"""
Robô Unimed Guarulhos - finaliza as guias SP/SADT em aberto no portal SGU-Card.

Portal: https://rda.unimedguarulhos.coop.br/cmagnet  (login manual: o robô abre o navegador e VOCÊ entra)

Função 1 - pela planilha (exportação da agenda do ClinicWeb):
  para cada número de guia da coluna Observação:
    Exames em aberto -> Número da guia (sem filtro de data) -> Filtrar -> abre a guia
    -> lê as datas de leitura/assinatura das sessões (campo da linha do procedimento)
    -> se o DIA da Data da Agenda estiver entre elas: Finalizar Parcial -> Confirmar
  Saída: cópia da planilha com as colunas "Status" e "Status Detalhado" (depois de Observação):
    SEM TR        linha sem número de guia
    QR Lido       a guia tem leitura no dia e foi finalizada
                  (paciente com falta e leitura: "QR Lido" + aviso no Status Detalhado, em amarelo)
    QR Não lido   sem leitura no dia, guia não encontrada, guia de outro paciente ou erro ao finalizar

Função 2 - por período:
  Exames em aberto -> Data de atendimento de/até (nenhum outro filtro) -> Filtrar
  -> junta as guias de TODAS as páginas -> abre uma por uma -> se tiver ALGUMA data de leitura/assinatura
  preenchida: Finalizar Parcial -> Confirmar; sem nenhuma data: passa para a próxima.
  Saída: "Finalização Unimed <período>.xlsx" na Área de Trabalho.

O portal usa um código de segurança (dynaHash) em cada link: o robô só navega clicando/abrindo os
links que o próprio portal mostra. Endereço montado à mão derruba a sessão.
"""
from __future__ import annotations

import asyncio
import datetime as dt
import difflib
import os
import re
import sys
import time
import unicodedata
from pathlib import Path
from urllib.parse import parse_qs, urljoin, urlparse

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill

import navegador  # noqa: F401  (acerta a pasta do navegador dentro do .exe)
import xlsx_seguro

URL_INICIAL = "https://rda.unimedguarulhos.coop.br/cmagnet/exames/emaberto/lista.do"
TEMPO_LOGIN_S = 15 * 60
MAX_ABAS = 5
PASTA_DADOS = Path.home() / "CentralAutomacoes" / "Unimed"

QR_LIDO, QR_NAO_LIDO, SEM_TR = "QR Lido", "QR Não lido", "SEM TR"
AMARELO = PatternFill("solid", fgColor="FFEB9C")
VERMELHO = PatternFill("solid", fgColor="F8CBAD")
VERDE = PatternFill("solid", fgColor="C6EFCE")
NAVY = PatternFill("solid", fgColor="002060")


# ----------------------------------------------------------------- utilidades

def _norm(txt):
    txt = "".join(c for c in unicodedata.normalize("NFD", str(txt or "")) if unicodedata.category(c) != "Mn")
    return " ".join(re.sub(r"[^A-Za-z0-9 ]+", " ", txt).upper().split())


def para_data(v):
    if isinstance(v, dt.datetime):
        return v.date()
    if isinstance(v, dt.date):
        return v
    m = re.search(r"(\d{1,2})/(\d{1,2})/(\d{4})", str(v or ""))
    if m:
        try:
            return dt.date(int(m.group(3)), int(m.group(2)), int(m.group(1)))
        except ValueError:
            return None
    return None


def guias_do_texto(v):
    """Observação -> números de guia (7 a 9 dígitos). 15346664.0 -> ['15346664']."""
    if v is None:
        return []
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    return list(dict.fromkeys(re.findall(r"(?<!\d)\d{7,9}(?!\d)", str(v).replace(".", ""))))


def eh_falta(status_agenda):
    n = _norm(status_agenda)
    return "FALT" in n or "NAO CHEGOU" in n


def mesmo_paciente(nome_planilha, nome_portal):
    if not nome_planilha or not nome_portal:
        return True
    return difflib.SequenceMatcher(None, _norm(nome_planilha), _norm(nome_portal)).ratio() >= 0.8


def _fmt(datas):
    return ", ".join(f"{d:%d/%m/%Y}" for d in sorted(datas)) or "nenhuma"


def area_de_trabalho():
    try:
        import gerador_agenda
        return gerador_agenda.area_de_trabalho()
    except Exception:  # noqa: BLE001
        casa = Path.home()
        for p in (casa / "Desktop", casa / "OneDrive" / "Desktop", casa / "Área de Trabalho"):
            if p.exists():
                return p
        return casa


# ----------------------------------------------------------------- planilha

COLUNAS = {
    "data": ("DATA DA AGENDA", "DIA", "DATA"),
    "paciente": ("PACIENTE",),
    "status_agenda": ("STATUS DA AGENDA COMBO", "STATUS DA AGENDA"),
    "obs": ("OBSERVACAO",),
}


def _mapear(ws):
    cab = {_norm(c.value): c.column for c in ws[1] if c.value is not None}
    cols = {}
    for chave, nomes in COLUNAS.items():
        for n in nomes:
            if n in cab:
                cols[chave] = cab[n]
                break
    faltam = [COLUNAS[k][0].title() for k in ("data", "obs") if k not in cols]
    if faltam:
        raise ValueError("Não achei a(s) coluna(s) " + ", ".join(faltam) + " na primeira linha da planilha.")
    cols["status"] = cab.get("STATUS")
    cols["detalhe"] = cab.get("STATUS DETALHADO")
    return cols


def ler_planilha(caminho):
    """Devolve (wb, ws, cols, linhas). Cada linha: {row, data, paciente, status_agenda, guias}."""
    wb = xlsx_seguro.carregar(caminho)
    ws = wb.active
    cols = _mapear(ws)
    linhas = []
    for r in range(2, ws.max_row + 1):
        val = lambda k: ws.cell(r, cols[k]).value if cols.get(k) else None  # noqa: E731
        if not any(ws.cell(r, c).value not in (None, "") for c in range(1, ws.max_column + 1)):
            continue
        linhas.append({
            "row": r, "data": para_data(val("data")), "paciente": " ".join(str(val("paciente") or "").split()),
            "status_agenda": str(val("status_agenda") or ""), "guias": guias_do_texto(val("obs")),
        })
    return wb, ws, cols, linhas


def status_da_linha(lin, res, simular=False):
    """lin: linha da planilha; res: resultado da guia (ou None). Devolve (Status, Status Detalhado, cor)."""
    if not lin["guias"]:
        return SEM_TR, "", None
    if res is None:
        return "", "guia não processada (teste limitado ou interrompido)", None
    if res.get("erro"):
        return QR_NAO_LIDO, res["erro"], VERMELHO
    if lin["data"] is None:
        return QR_NAO_LIDO, "Data da Agenda vazia ou inválida na planilha", VERMELHO
    leituras = res.get("leituras", set())
    if lin["data"] not in leituras:
        return QR_NAO_LIDO, f"sem leitura em {lin['data']:%d/%m/%Y} (leituras na guia: {_fmt(leituras)})", None
    if res.get("erro_finalizar"):
        return QR_NAO_LIDO, f"leitura OK, mas deu erro ao finalizar: {res['erro_finalizar']}", VERMELHO
    detalhes = []
    if simular:
        detalhes.append("SIMULAÇÃO: não foi finalizada")
    if eh_falta(lin["status_agenda"]):
        detalhes.append(f"ATENÇÃO: agenda com '{lin['status_agenda']}', mas o portal tem leitura no dia")
    return QR_LIDO, "; ".join(detalhes), (AMARELO if eh_falta(lin["status_agenda"]) else None)


def gravar_planilha(wb, ws, cols, linhas, resultados, destino, simular=False):
    """Põe Status / Status Detalhado logo depois de Observação (ou usa as que já existirem) e salva."""
    if not cols.get("status"):
        c = cols["obs"] + 1
        ws.insert_cols(c, 2)
        modelo = ws.cell(1, cols["obs"])
        for i, titulo in enumerate(("Status", "Status Detalhado")):
            cel = ws.cell(1, c + i, titulo)
            if modelo.has_style:
                cel.font, cel.fill = modelo.font.copy(), modelo.fill.copy()
                cel.alignment, cel.border = modelo.alignment.copy(), modelo.border.copy()
        ws.column_dimensions[openpyxl.utils.get_column_letter(c)].width = 14
        ws.column_dimensions[openpyxl.utils.get_column_letter(c + 1)].width = 60
        cols["status"], cols["detalhe"] = c, c + 1
    elif not cols.get("detalhe"):
        ws.insert_cols(cols["status"] + 1, 1)
        ws.cell(1, cols["status"] + 1, "Status Detalhado")
        cols["detalhe"] = cols["status"] + 1

    cont = {}
    for lin in linhas:
        res = None
        for g in lin["guias"]:
            if g in resultados:
                res = resultados[g]
                break
        st, det, cor = status_da_linha(lin, res, simular)
        if not st and not det:
            continue
        cs, cd = ws.cell(lin["row"], cols["status"], st), ws.cell(lin["row"], cols["detalhe"], det or None)
        if cor is not None:
            cs.fill = cd.fill = cor
        cont[st or "não processada"] = cont.get(st or "não processada", 0) + 1

    # aba de apoio com o que o portal mostrou de cada guia
    if "Guias Unimed" in wb.sheetnames:
        del wb["Guias Unimed"]
    wg = wb.create_sheet("Guias Unimed")
    cab = ["Guia", "Paciente (portal)", "Leituras na guia", "Datas na planilha", "Ação", "Mensagem do portal / erro"]
    wg.append(cab)
    for g, r in resultados.items():
        wg.append([g, r.get("beneficiario", ""), _fmt(r.get("leituras", set())), _fmt(r.get("datas_planilha", set())),
                   r.get("acao", ""), r.get("erro") or r.get("erro_finalizar") or r.get("mensagem", "")])
    _estilo_cabecalho(wg, [10, 40, 30, 22, 22, 80])
    wb.save(destino)
    return cont


def _estilo_cabecalho(ws, larguras):
    for cel in ws[1]:
        cel.font = Font(bold=True, color="FFFFFF")
        cel.fill = NAVY
        cel.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for i, w in enumerate(larguras, start=1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions


# ----------------------------------------------------------------- portal


class LoginNecessario(Exception):
    pass


class GuiaNaoEncontrada(Exception):
    pass


JS_LER_GUIA = """() => {
  const v = n => { const e = document.querySelector('[name="' + n + '"]'); return e ? (e.value || '').trim() : ''; };
  const series = [...document.querySelectorAll('input[id^="dt_serie_"], input[name^="dt_serie_"]')]
                   .map(e => (e.value || '').trim()).filter(Boolean);
  const procs = [];
  for (let i = 1; i <= 30; i++) {
    const c = v('CD_ITEM_' + i);
    if (!c) continue;
    procs.push({codigo: c, descricao: v('DS_ITEM_' + i), autorizada: v('QT_AUTORIZADA_' + i),
                realizada: v('QT_REALIZADA_' + i)});
  }
  const erro = document.querySelector('#error_label');
  return {guia: v('NR_GUIA'), beneficiario: v('NM_PESSOA'), carteira: v('NR_CARTEIRA'), series,
          procedimentos: procs, tem_parcial: !!document.querySelector('[name="Button_Parcial"]'),
          erro: erro ? erro.innerText.trim() : ''};
}"""

JS_FILTRAR = """(f) => {
  const form = document.querySelector('[name="Button_FIltro"]').form;
  for (const e of form.querySelectorAll('input[type="text"], input:not([type])')) e.value = '';
  for (const s of form.querySelectorAll('select')) s.selectedIndex = 0;
  const pon = (n, val) => { const e = form.querySelector('[name="' + n + '"]'); if (e) e.value = val; };
  pon('s_nr_guia', f.guia || ''); pon('s_dt_ini', f.ini || ''); pon('s_dt_fim', f.fim || '');
  return true;
}"""

JS_LINHAS = """() => {
  const out = [];
  for (const a of document.querySelectorAll('a[href*="sadt/execucao.do"]')) {
    const g = (a.innerText || '').trim();
    if (!/^\\d{6,}$/.test(g)) continue;
    const tds = [...a.closest('tr').querySelectorAll('td')].map(td => (td.innerText || '').trim());
    out.push({guia: g, href: a.getAttribute('href'), celulas: tds});
  }
  const m = (document.body.innerText.match(/(\\d+)\\s*exame\\(s\\) encontrado/) || [])[1];
  const prox = [...document.querySelectorAll('a')].find(a => (a.innerText || '').trim() === 'Próxima'
                 && a.getAttribute('href') && !a.getAttribute('href').startsWith('#'));
  return {linhas: out, total: m ? parseInt(m) : null, proxima: prox ? prox.getAttribute('href') : null};
}"""

JS_HREFS_LISTA = "els => els.map(a => a.getAttribute('href') || '')"


def _cd_guia(url):
    """cd_guia do endereço (é o código interno da guia; vem em execucao.do e sucesso_execucao.do)."""
    try:
        return (parse_qs(urlparse(url or "").query).get("cd_guia") or [""])[0]
    except Exception:
        return ""


def _primeira_linha(e):
    return (str(e).splitlines() or ["erro"])[0][:300] or e.__class__.__name__


class Aba:
    """Uma aba do navegador no SGU-Card. Cada aba cuida só das janelas que ela mesma abriu."""

    NOMES_CONFIRMAR = ("Confirmar", "Finalizar", "OK", "Sim")

    def __init__(self, page, nome, log, parar, simular):
        self.page, self.nome, self.log, self.parar, self.simular = page, nome, log, parar, simular
        self.dialogos, self.popups = [], []
        self.cd_guia_aberta = ""
        page.on("dialog", self._aceitar_dialogo)
        page.on("popup", self._nova_janela)

    def _nova_janela(self, p):
        self.popups.append(p)

    def _aceitar_dialogo(self, d):
        self.dialogos.append(" ".join((d.message or "").split()))
        asyncio.ensure_future(self._aceitar(d))

    @staticmethod
    async def _aceitar(d):
        try:
            await d.accept()
        except Exception:
            pass

    def parou(self):
        return self.parar is not None and self.parar.is_set()

    async def esperar(self, ms):
        await self.page.wait_for_timeout(ms)

    async def _visivel(self, loc):
        try:
            return await loc.count() > 0 and await loc.first.is_visible()
        except Exception:
            return False

    async def texto(self):
        try:
            return await self.page.inner_text("body", timeout=3000)
        except Exception:
            return ""

    # ---- login
    async def na_tela_de_login(self):
        return (await self._visivel(self.page.locator("input[type=password]"))
                or "Identifique-se" in (await self.texto()))

    async def logado(self):
        if await self.na_tela_de_login():
            return False
        return await self.page.locator('a[href*="emaberto/lista.do"]').count() > 0

    # ---- navegação (só por links do próprio portal)
    async def abrir_link(self, href):
        href = (href or "").strip()
        if href.lower().startswith("javascript"):
            raise RuntimeError(f"link do portal não é um endereço: {href[:80]}")
        await self.page.goto(urljoin(self.page.url, href), wait_until="domcontentloaded")
        if await self.na_tela_de_login():
            raise LoginNecessario()

    async def href_lista(self):
        """Endereço de 'Exames em aberto'. O item 'Exames' do menu é 'javascript: loadFirstMenu(...)':
        dele só se aproveita o endereço que está dentro."""
        try:
            hrefs = await self.page.eval_on_selector_all('a[href*="emaberto/lista.do"]', JS_HREFS_LISTA)
        except Exception:
            hrefs = []
        reais = [h for h in hrefs if not h.strip().lower().startswith("javascript")]
        if reais:
            return reais[0]
        for h in hrefs:
            m = re.search(r"['\"]([^'\"]*emaberto/lista\.do[^'\"]*)['\"]", h)
            if m:
                return m.group(1)
        return None

    async def ir_para_lista(self):
        href = await self.href_lista()
        if not href:
            if await self.na_tela_de_login():
                raise LoginNecessario()
            await self.page.go_back(wait_until="domcontentloaded")
            href = await self.href_lista()
            if not href:
                raise RuntimeError("não achei o menu 'Exames em aberto' no portal")
        await self.abrir_link(href)
        await self.page.locator('[name="Button_FIltro"]').wait_for(state="attached", timeout=30000)

    async def filtrar(self, guia="", ini="", fim=""):
        await self.ir_para_lista()
        await self.page.evaluate(JS_FILTRAR, {"guia": guia, "ini": ini, "fim": fim})
        async with self.page.expect_navigation(wait_until="domcontentloaded", timeout=60000):
            await self.page.locator('[name="Button_FIltro"]').click()
        if await self.na_tela_de_login():
            raise LoginNecessario()
        return await self.page.evaluate(JS_LINHAS)

    async def linhas_da_pagina(self):
        return await self.page.evaluate(JS_LINHAS)

    async def abrir_guia(self, guia, href=None):
        """Abre a guia. Com href (link da lista), tenta direto; senão (ou se falhar) pesquisa pelo número
        em Exames em aberto, sem filtro de data."""
        if href:
            try:
                await self.abrir_link(href)
                d = await self.ler_guia()
                if d["guia"] == guia:
                    return d
            except LoginNecessario:
                raise
            except Exception:
                pass
        res = await self.filtrar(guia=guia)
        alvo = [l for l in res["linhas"] if l["guia"] == guia]
        if not alvo:
            raise GuiaNaoEncontrada("guia não encontrada em Exames em aberto (pode já estar finalizada)")
        await self.abrir_link(alvo[0]["href"])
        d = await self.ler_guia()
        if d["guia"] and d["guia"] != guia:
            raise RuntimeError(f"abriu a guia {d['guia']} no lugar da {guia}")
        return d

    async def ler_guia(self):
        try:
            await self.page.locator('[name="NR_GUIA"]').wait_for(state="attached", timeout=30000)
        except Exception:
            if await self.na_tela_de_login():
                raise LoginNecessario()
            raise RuntimeError("a tela da guia não abriu")
        await self.esperar(1500)  # o portal termina de preencher as sessões por script
        d = await self.page.evaluate(JS_LER_GUIA)
        d["leituras"] = {x for x in (para_data(s) for s in d["series"]) if x}
        self.cd_guia_aberta = _cd_guia(self.page.url)
        return d

    # ---- finalizar
    # Fluxo do portal: Finalizar Parcial -> janela com "Confirmar" -> Confirmar -> página
    # sucesso_execucao.do. Depois disso a guia volta para "Exames em aberto" com um atendimento novo.
    def _sucesso(self):
        return "sucesso_execucao" in (self.page.url or "")

    def _minhas_paginas(self):
        return [self.page] + [p for p in self.popups if not p.is_closed()]

    async def _botao_confirmar(self):
        for nome in self.NOMES_CONFIRMAR:
            for pagina in self._minhas_paginas():
                for frame in pagina.frames:
                    try:
                        loc = frame.get_by_role("button", name=nome, exact=True)
                        for i in range(await loc.count()):
                            if await loc.nth(i).is_visible():
                                return nome, loc.nth(i)
                    except Exception:
                        continue
        return None, None

    async def _mensagens_visiveis(self):
        txts = []
        for pagina in self._minhas_paginas():
            for frame in pagina.frames:
                try:
                    for sel in (".ui-dialog", "#error_label", ".MensagemErro", ".mensagem", ".msg"):
                        loc = frame.locator(sel)
                        for i in range(min(await loc.count(), 5)):
                            if await loc.nth(i).is_visible():
                                t = " ".join((await loc.nth(i).inner_text(timeout=2000)).split())
                                if t and t not in txts:
                                    txts.append(t)
                except Exception:
                    continue
        return txts + self.dialogos

    async def _erro_na_tela(self):
        try:
            return (await self.page.evaluate(
                "() => ((document.querySelector('#error_label')||{}).innerText||'').trim()")) or ""
        except Exception:
            return ""

    async def _fechar_popups(self):
        for p in self.popups:
            try:
                if not p.is_closed():
                    await p.close()
            except Exception:
                pass
        self.popups.clear()

    async def finalizar_parcial(self):
        """Finalizar Parcial -> Confirmar na guia aberta. Só conta como finalizada se abrir a página de
        sucesso DESTA guia. Devolve a mensagem do portal; se não der certo -> RuntimeError."""
        if self.simular:
            return "simulação: nada foi clicado"
        self.dialogos.clear()
        cd = self.cd_guia_aberta
        await self.page.locator('[name="Button_Parcial"]').click()

        nome, botao = None, None
        for _ in range(80):  # até ~40 s pela janela do Confirmar
            if self.parou():
                raise RuntimeError("interrompido antes de clicar em Confirmar")
            if self._sucesso():  # confirmação nativa do navegador (já aceita) levou direto ao sucesso
                break
            nome, botao = await self._botao_confirmar()
            if botao is not None:
                break
            erro = await self._erro_na_tela()
            if erro:
                raise RuntimeError(f"o portal avisou: {erro}")
            await self.esperar(500)

        if not self._sucesso():
            if botao is None:
                msgs = await self._mensagens_visiveis()
                raise RuntimeError("não apareceu o botão Confirmar depois do Finalizar Parcial"
                                   + (f" (portal: {' | '.join(msgs)[:300]})" if msgs else ""))
            await botao.click()
            for _ in range(90):  # até ~45 s pela página de sucesso
                await self.esperar(500)
                if self._sucesso():
                    break
                erro = await self._erro_na_tela()
                if erro:
                    raise RuntimeError(f"o portal avisou: {erro}")

        if not self._sucesso():
            msgs = await self._mensagens_visiveis()
            raise RuntimeError(f"cliquei em {nome} mas a página de sucesso não abriu: confira a guia no portal"
                               + (f" (portal: {' | '.join(msgs)[:300]})" if msgs else ""))
        try:
            await self.page.wait_for_load_state("domcontentloaded", timeout=15000)
        except Exception:
            pass
        await self._fechar_popups()
        cd_ok = _cd_guia(self.page.url)
        if cd and cd_ok and cd != cd_ok:
            raise RuntimeError(f"ATENÇÃO: o portal mostrou sucesso de OUTRA guia (código {cd_ok}, esperado {cd}). "
                               "Confira as duas no portal")
        texto = " ".join((await self.texto()).split())
        m = re.search(r"[^.!]*(sucesso|finalizad|gravad|realizad)[^.!]*[.!]?", texto, re.I)
        return m.group(0).strip()[:200] if m else "finalizada (página de sucesso)"


# ----------------------------------------------------------------- navegador e abas

class _Navegador:
    def __init__(self, log, parar, simular, n_abas):
        self.log, self.parar, self.simular = log, parar, simular
        self.n_abas = max(1, min(int(n_abas or 1), MAX_ABAS))
        self._pw = self.ctx = None
        self.principal = None
        self.abas = []

    async def __aenter__(self):
        from playwright.async_api import async_playwright
        PASTA_DADOS.mkdir(parents=True, exist_ok=True)
        self._pw = await async_playwright().start()
        perfil = str(PASTA_DADOS / "navegador")
        opcoes = dict(headless=os.environ.get("UNIMED_HEADLESS") == "1", locale="pt-BR",
                      viewport={"width": 1366, "height": 860})
        try:  # Google Chrome instalado, se houver; senão o Chromium do Playwright
            self.ctx = await self._pw.chromium.launch_persistent_context(perfil, channel="chrome", **opcoes)
        except Exception:
            try:
                self.ctx = await self._pw.chromium.launch_persistent_context(perfil, **opcoes)
            except Exception as e:
                if not navegador.falta_navegador(e):
                    raise
                await asyncio.to_thread(navegador.instalar_chromium, self.log)
                self.ctx = await self._pw.chromium.launch_persistent_context(perfil, **opcoes)
        self.ctx.set_default_timeout(30000)
        page = self.ctx.pages[0] if self.ctx.pages else await self.ctx.new_page()
        self.principal = Aba(page, "[aba 1]", self.log, self.parar, self.simular)
        await page.goto(URL_INICIAL, wait_until="domcontentloaded")
        await esperar_login(self.principal, self.log, self.parar)
        self.abas = [self.principal]
        return self

    async def __aexit__(self, *exc):
        for fechar in (self.ctx.close, self._pw.stop):
            try:
                await fechar()
            except Exception:
                pass

    async def abrir_extras(self, quantas_guias):
        """Abre as outras abas a partir do link de 'Exames em aberto' da aba logada (mesma sessão)."""
        alvo = max(1, min(self.n_abas, quantas_guias))
        for n in range(2, alvo + 1):
            page = None
            try:
                href = await self.principal.href_lista()
                if not href:
                    raise RuntimeError("sem o link de Exames em aberto")
                page = await self.ctx.new_page()
                aba = Aba(page, f"[aba {n}]", self.log, self.parar, self.simular)
                await aba.abrir_link(urljoin(self.principal.page.url, href))
                await page.locator('[name="Button_FIltro"]').wait_for(state="attached", timeout=30000)
                self.abas.append(aba)
            except Exception as e:  # noqa: BLE001
                motivo = "abriu sem o login" if isinstance(e, LoginNecessario) else _primeira_linha(e)[:80]
                self.log(f"AVISO [aba {n}] não pôde ser usada ({motivo}); seguindo sem ela")
                if page is not None:
                    try:
                        await page.close()
                    except Exception:
                        pass
        if alvo > 1:
            self.log(f"{len(self.abas)} aba(s) trabalhando ao mesmo tempo")


async def esperar_login(aba, log, parar):
    if await aba.logado():
        return
    log("AVISO Faça o login no navegador que abriu (usuário e senha do SGU-Card). O robô continua sozinho depois.")
    for _ in range(TEMPO_LOGIN_S):
        if aba.parou():
            raise RuntimeError("Interrompido antes do login.")
        if await aba.logado():
            log("Login OK")
            await aba.esperar(1000)
            return
        await aba.esperar(1000)
    raise RuntimeError("O login não foi feito em 15 minutos.")


class _Repeticao:
    """Para o lote quando o MESMO erro de portal se repete em várias guias seguidas."""

    def __init__(self, limite=3):
        self.limite, self.ultimo, self.vezes = limite, None, 0

    def erro(self, msg):
        self.vezes = self.vezes + 1 if msg == self.ultimo else 1
        self.ultimo = msg
        return self.vezes >= self.limite

    def ok(self):
        self.ultimo, self.vezes = None, 0


async def _distribuir(nav, itens, tratar, log, progresso):
    """Põe os itens numa fila e cada aba vai pegando o próximo. tratar(aba, item) -> (chave, resultado, msg_erro).
    msg_erro (texto) conta para o 'mesmo erro seguidas'. Devolve {chave: resultado}."""
    fila = asyncio.Queue()
    for it in itens:
        fila.put_nowait(it)
    resultados, feitas = {}, [0]
    rep, travou = _Repeticao(), asyncio.Event()
    mostrar = len(nav.abas) > 1

    async def trabalhar(aba):
        while not (aba.parou() or travou.is_set()):
            try:
                it = fila.get_nowait()
            except asyncio.QueueEmpty:
                return
            for tentativa in (1, 2):
                try:
                    chave, res, erro = await tratar(aba, it, f"{aba.nome} " if mostrar else "")
                    resultados[chave] = res
                    if erro and rep.erro(erro):
                        log(f"ERRO o mesmo erro em {rep.vezes} guias seguidas: parei para não repetir em todas.")
                        travou.set()
                    elif not erro:
                        rep.ok()
                    break
                except LoginNecessario:
                    if aba is not nav.principal:
                        fila.put_nowait(it)  # devolve para outra aba
                        log(f"AVISO {aba.nome} perdeu o login; as outras abas continuam")
                        try:
                            await aba.page.close()
                        except Exception:
                            pass
                        return
                    if tentativa == 2:
                        raise RuntimeError("a sessão do portal caiu de novo")
                    log("AVISO a sessão do portal caiu: faça o login de novo no navegador")
                    await aba.page.goto(URL_INICIAL, wait_until="domcontentloaded")
                    await esperar_login(aba, log, aba.parar)
            feitas[0] += 1
            if progresso:
                progresso(feitas[0], len(itens))

    await asyncio.gather(*(trabalhar(a) for a in nav.abas))
    if not fila.empty() and not (nav.principal.parou() or travou.is_set()):
        await trabalhar(nav.principal)  # sobrou item de aba que caiu
    if nav.principal.parou():
        log("Interrompido pelo usuário.")
    return resultados


# ----------------------------------------------------------------- função 1: pela planilha

async def _planilha_async(guias, datas_por_guia, pac_por_guia, simular, n_abas, log, parar, progresso):
    async with _Navegador(log, parar, simular, n_abas) as nav:
        await nav.abrir_extras(len(guias))

        async def tratar(aba, g, rot):
            r = {"datas_planilha": datas_por_guia[g]}
            try:
                dados = await aba.abrir_guia(g)
            except GuiaNaoEncontrada as e:
                r.update(erro=str(e), acao="não encontrada")
                log(f"AVISO {rot}guia {g}: {e}")
                return g, r, None
            except LoginNecessario:
                raise
            except Exception as e:  # noqa: BLE001
                r.update(erro=_primeira_linha(e), acao="ERRO")
                log(f"ERRO {rot}guia {g}: {r['erro']}")
                return g, r, r["erro"]
            r.update(beneficiario=dados["beneficiario"], leituras=dados["leituras"])
            bate = datas_por_guia[g] & dados["leituras"]
            if not mesmo_paciente(pac_por_guia[g], dados["beneficiario"]):
                r["erro"] = f"a guia é de {dados['beneficiario']} (planilha: {pac_por_guia[g]}); não finalizada"
                r["acao"] = "não finalizada"
            elif not bate:
                r["acao"] = "sem leitura no dia"
            elif not dados["tem_parcial"]:
                r["erro_finalizar"] = "a guia não tem o botão Finalizar Parcial" + (
                    f" ({dados['erro']})" if dados["erro"] else "")
                r["acao"] = "não finalizada"
            else:
                try:
                    r["mensagem"] = await aba.finalizar_parcial()
                    r["acao"] = "SIMULADA" if simular else "FINALIZADA"
                except LoginNecessario:
                    r["erro_finalizar"] = "a sessão caiu durante a finalização: confira a guia no portal"
                    r["acao"] = "VERIFICAR"
                except Exception as e:  # noqa: BLE001
                    r["erro_finalizar"] = _primeira_linha(e)
                    r["acao"] = "ERRO"
            extra = r.get("erro") or r.get("erro_finalizar") or r.get("mensagem")
            tag = "OK" if r["acao"] in ("FINALIZADA", "SIMULADA") else ("ERRO" if r.get("erro") or r.get("erro_finalizar")
                                                                       else "AVISO")
            log(f"{tag} {rot}guia {g}: leituras {_fmt(dados['leituras'])} | planilha {_fmt(datas_por_guia[g])} | "
                f"{r['acao']}" + (f" - {extra}" if extra else ""))
            return g, r, r.get("erro_finalizar") if r["acao"] == "ERRO" else None

        return await _distribuir(nav, guias, tratar, log, progresso)


def executar_planilha(caminho, limite=0, simular=False, log=print, parar=None, progresso=None, n_abas=1):
    """Função 1. Devolve o caminho da planilha preenchida."""
    wb, ws, cols, linhas = ler_planilha(caminho)
    datas_por_guia, pac_por_guia = {}, {}
    for lin in linhas:
        for g in lin["guias"]:
            datas_por_guia.setdefault(g, set())
            if lin["data"]:
                datas_por_guia[g].add(lin["data"])
            pac_por_guia.setdefault(g, lin["paciente"])
    guias = list(datas_por_guia)
    sem = sum(1 for l in linhas if not l["guias"])
    if limite:
        guias = guias[:limite]
    log(f"{len(linhas)} linha(s) na planilha; {len(datas_por_guia)} guia(s); {sem} linha(s) sem guia"
        + (f"; testando só {len(guias)} guia(s)" if limite else "")
        + ("; MODO SIMULAÇÃO (não finaliza nada)" if simular else ""))
    if not guias:
        raise ValueError("Nenhum número de guia na coluna Observação.")

    resultados = asyncio.run(_planilha_async(guias, datas_por_guia, pac_por_guia, simular, n_abas, log, parar,
                                             progresso))
    resultados = {g: resultados[g] for g in guias if g in resultados}  # na ordem da planilha
    destino = Path(caminho).with_name(f"{Path(caminho).stem} - Unimed {dt.datetime.now():%Y-%m-%d %H%M}.xlsx")
    cont = gravar_planilha(wb, ws, cols, linhas, resultados, destino, simular)
    log("Pronto: " + ", ".join(f"{n} {s}" for s, n in cont.items()) + f". Planilha: {destino.name}")
    return destino


# ----------------------------------------------------------------- função 2: por período

CAB_PERIODO = ["Guia", "Data de atendimento", "Beneficiário", "Contratado executante", "Leituras na guia",
               "Situação", "Mensagem do portal / erro"]
CORES_PERIODO = {"FINALIZADA": VERDE, "SIMULADA": VERDE, "SEM ASSINATURA": None, "ERRO": VERMELHO,
                 "NÃO ENCONTRADA": AMARELO}


async def _coletar(aba, ini_txt, fim_txt, log):
    """Pesquisa o período e junta as guias de todas as páginas (antes de finalizar qualquer uma)."""
    res = await aba.filtrar(ini=ini_txt, fim=fim_txt)
    log(f"{res['total'] if res['total'] is not None else '?'} exame(s) em aberto de {ini_txt} a {fim_txt}")
    vistas, pagina = {}, 1
    while True:
        for l in res["linhas"]:
            vistas.setdefault(l["guia"], l)
        if not res["proxima"] or aba.parou() or pagina >= 500:
            break
        await aba.abrir_link(res["proxima"])
        pagina += 1
        novo = await aba.linhas_da_pagina()
        if [l["guia"] for l in novo["linhas"]] == [l["guia"] for l in res["linhas"]]:
            break  # a página não mudou
        res = novo
    return list(vistas.values()), pagina


async def _periodo_async(ini_txt, fim_txt, limite, simular, n_abas, log, parar, progresso):
    async with _Navegador(log, parar, simular, n_abas) as nav:
        for tentativa in (1, 2):
            try:
                itens, paginas = await _coletar(nav.principal, ini_txt, fim_txt, log)
                break
            except LoginNecessario:
                if tentativa == 2:
                    raise RuntimeError("a sessão do portal caiu de novo")
                log("AVISO a sessão do portal caiu: faça o login de novo no navegador")
                await nav.principal.page.goto(URL_INICIAL, wait_until="domcontentloaded")
                await esperar_login(nav.principal, log, parar)
        log(f"{len(itens)} guia(s) em {paginas} página(s)" + (f"; testando só {limite}" if limite else "")
            + ("; MODO SIMULAÇÃO (não finaliza nada)" if simular else ""))
        if limite:
            itens = itens[:limite]
        if not itens:
            return [], []
        await nav.abrir_extras(len(itens))

        async def tratar(aba, it, rot):
            g, cel = it["guia"], it["celulas"]
            linha = {"Guia": g,
                     "Data de atendimento": next((c for c in cel if re.match(r"\d{2}/\d{2}/\d{4}", c)), ""),
                     "Beneficiário": next((c for c in cel if re.match(r"[\d.\-]+ - ", c)), ""),
                     "Contratado executante": cel[-2] if len(cel) >= 2 else ""}
            try:
                dados = await aba.abrir_guia(g, it.get("href"))
            except GuiaNaoEncontrada as e:
                linha.update({"Situação": "NÃO ENCONTRADA", "Mensagem do portal / erro": str(e)})
                log(f"AVISO {rot}guia {g}: {e}")
                return g, linha, None
            except LoginNecessario:
                raise
            except Exception as e:  # noqa: BLE001
                linha.update({"Situação": "ERRO", "Mensagem do portal / erro": _primeira_linha(e)})
                log(f"ERRO {rot}guia {g}: {linha['Mensagem do portal / erro']}")
                return g, linha, linha["Mensagem do portal / erro"]
            linha["Leituras na guia"] = _fmt(dados["leituras"]) if dados["leituras"] else ""
            if not dados["series"]:
                linha.update({"Situação": "SEM ASSINATURA",
                              "Mensagem do portal / erro": "nenhuma data de leitura/assinatura: passou para a próxima"})
                log(f"AVISO {rot}guia {g}: sem data de assinatura, não finalizada")
                return g, linha, None
            if not dados["tem_parcial"]:
                msg = "a guia não tem o botão Finalizar Parcial" + (f" ({dados['erro']})" if dados["erro"] else "")
                linha.update({"Situação": "ERRO", "Mensagem do portal / erro": msg})
                log(f"ERRO {rot}guia {g}: {msg}")
                return g, linha, msg
            try:
                linha["Mensagem do portal / erro"] = await aba.finalizar_parcial()
                linha["Situação"] = "SIMULADA" if simular else "FINALIZADA"
                log(f"OK {rot}guia {g}: leituras {linha['Leituras na guia']} | {linha['Situação'].lower()} - "
                    f"{linha['Mensagem do portal / erro']}")
                return g, linha, None
            except LoginNecessario:
                linha.update({"Situação": "ERRO",
                              "Mensagem do portal / erro": "a sessão caiu durante a finalização: confira a guia"})
                return g, linha, None
            except Exception as e:  # noqa: BLE001
                linha.update({"Situação": "ERRO", "Mensagem do portal / erro": _primeira_linha(e)})
                log(f"ERRO {rot}guia {g}: {linha['Mensagem do portal / erro']}")
                return g, linha, linha["Mensagem do portal / erro"]

        res = await _distribuir(nav, itens, tratar, log, progresso)
        return itens, [res[it["guia"]] for it in itens if it["guia"] in res]


def executar_periodo(inicio, fim, limite=0, simular=False, log=print, parar=None, progresso=None, n_abas=1):
    """Função 2. Abre todas as guias em aberto do período e finaliza parcial as que têm data de assinatura."""
    if fim < inicio:
        raise ValueError("A data final é anterior à data inicial.")
    _, linhas_saida = asyncio.run(_periodo_async(f"{inicio:%d/%m/%Y}", f"{fim:%d/%m/%Y}", limite, simular, n_abas,
                                                 log, parar, progresso))
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Finalização"
    ws.append(CAB_PERIODO)
    for l in linhas_saida:
        ws.append([l.get(k, "") for k in CAB_PERIODO])
        cor = CORES_PERIODO.get(l.get("Situação"), AMARELO)
        if cor is not None:
            ws.cell(ws.max_row, CAB_PERIODO.index("Situação") + 1).fill = cor
    _estilo_cabecalho(ws, [12, 18, 48, 36, 26, 16, 80])
    destino = area_de_trabalho() / (f"Finalização Unimed {inicio:%d-%m} a {fim:%d-%m-%Y} "
                                    f"{dt.datetime.now():%Y-%m-%d %H%M}.xlsx")
    wb.save(destino)
    cont = {}
    for l in linhas_saida:
        cont[l["Situação"]] = cont.get(l["Situação"], 0) + 1
    log("Pronto: " + (", ".join(f"{n} {s}" for s, n in cont.items()) or "nenhuma guia") + f". Resultado: {destino.name}")
    return destino


if __name__ == "__main__":
    import argparse
    ap = argparse.ArgumentParser(description="Robô Unimed Guarulhos (SGU-Card)")
    sub = ap.add_subparsers(dest="modo", required=True)
    a1 = sub.add_parser("planilha")
    a1.add_argument("arquivo")
    a2 = sub.add_parser("periodo")
    a2.add_argument("inicio")
    a2.add_argument("fim")
    for a in (a1, a2):
        a.add_argument("--limite", type=int, default=0)
        a.add_argument("--simular", action="store_true")
        a.add_argument("--abas", type=int, default=3, help="abas do portal ao mesmo tempo (1 a 5)")
    a = ap.parse_args()
    try:
        if a.modo == "planilha":
            executar_planilha(a.arquivo, a.limite, a.simular, n_abas=a.abas)
        else:
            executar_periodo(para_data(a.inicio), para_data(a.fim), a.limite, a.simular, n_abas=a.abas)
    except (RuntimeError, ValueError) as e:
        sys.exit(str(e))
