"""
Robô Porto Saúde - conferência das guias da planilha (abas PORTO e PORTO CONV) com o portal.

Fluxo no portal (https://prestadores.portosaude.com.br/portal-prestador/):
  1. o robô abre o navegador e VOCÊ faz o login (CPF, senha e o reCAPTCHA do site);
  2. Autorização -> Pesquisar Autorização;
  3. para cada guia: Data Inicial + Data Final + Nº Guia Prestador -> Aplicar
     (o robô lê a lista que o próprio portal recebe: guia, senha, data, status);
  4. Expandir -> Imprimir: o robô pega o PDF da guia que o portal recebe e lê procedimento e quantidade.
  Nunca clica em outros botões da guia (ex.: Cancelar).

Regras da conferência:
  PORTO      : soma a quantidade autorizada de TODAS as guias do paciente e compara com o TOTAL MÊS
               (terapia não importa).
  PORTO CONV : o procedimento da guia tem que ser o do Grupo de Procedimento e a quantidade autorizada
               tem que ser a Quantidade.
  Nas duas   : avisa se a guia é de outro paciente (carteirinha/nome diferentes).

Saída: "Resultado Porto <data>.xlsx" na mesma pasta da planilha (a planilha de conferência não é alterada).
"""
from __future__ import annotations

import asyncio
import datetime as dt
import difflib
import json
import re
import sys
import unicodedata
from collections import defaultdict
from pathlib import Path

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill

import guia_porto_pdf
import navegador  # noqa: F401  (acerta a pasta do navegador dentro do .exe)

ABA = "PORTO"
ABA_CONV = "PORTO CONV"
ABAS = (ABA, ABA_CONV)

URL_INICIAL = "https://prestadores.portosaude.com.br/portal-prestador/"
URL_AUTORIZACAO = "https://prestadores.portosaude.com.br/portal-prestador/autorizacao"
MAX_DIAS = 89          # o portal recusa períodos maiores que 90 dias
TEMPO_LOGIN_S = 15 * 60
MAX_ABAS = 5

PASTA_DADOS = Path.home() / "CentralAutomacoes" / "Porto"


# ----------------------------------------------------------------- utilidades

def _norm(txt):
    txt = "".join(c for c in unicodedata.normalize("NFD", str(txt or "")) if unicodedata.category(c) != "Mn")
    return " ".join(re.sub(r"[^A-Za-z0-9 ]+", " ", txt).upper().split())


def _digitos(v):
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    return re.sub(r"\D", "", str(v or ""))


def _num(v):
    if v is None or v == "":
        return None
    if isinstance(v, (int, float)):
        return v
    try:
        return float(str(v).replace(",", "."))
    except ValueError:
        return None


def _unidade(v):
    m = re.search(r"\d+", str(v if v is not None else ""))
    return str(int(m.group())) if m else None


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


def separar_guias(valor):
    """'79.900.930 /  79.901.779' -> ['79900930', '79901779']; 79578715.0 -> ['79578715']."""
    if valor is None:
        return []
    if isinstance(valor, float) and valor.is_integer():
        valor = int(valor)
    txt = str(valor).replace(".", "")
    return list(dict.fromkeys(re.findall(r"\d{6,}", txt)))


def categoria(texto):
    """Terapia de um texto de procedimento: psicoped, psico, fono, to, fisio ou None."""
    n = f" {_norm(texto)} "
    if "PSICOPED" in n:
        return "psicoped"
    if "FONO" in n:
        return "fono"
    if "TERAPIA OCUPACIONAL" in n or " TO " in n:
        return "to"
    # fisioterapia: na guia da Porto pode vir como "Patologia neurológica com dependência de atividades
    # da vida diária" (e outras descrições de patologia)
    if "FISIO" in n or "PATOLOGIA" in n:
        return "fisio"
    if "PSICO" in n:
        return "psico"
    return None


NOME_CAT = {"psico": "Psicologia", "psicoped": "Psicopedagogia", "fono": "Fonoaudiologia",
            "to": "Terapia Ocupacional", "fisio": "Fisioterapia", None: "?"}


def _cabecalho(linhas, obrigatoria="PACIENTE"):
    for i, r in enumerate(linhas[:8]):
        nomes = [_norm(c) for c in (r or [])]
        if _norm(obrigatoria) in nomes:
            return i, nomes
    raise ValueError(f"Não achei o cabeçalho (coluna {obrigatoria}).")


# ----------------------------------------------------------------- planilha

def _ler_aba(caminho, aba):
    wb = openpyxl.load_workbook(caminho, data_only=True, read_only=True)
    if aba not in wb.sheetnames:
        wb.close()
        raise ValueError(f"Aba '{aba}' não encontrada na planilha.")
    linhas = [list(r) for r in wb[aba].iter_rows(values_only=True)]
    wb.close()
    return linhas


def ler_porto(caminho, unidade=None):
    """Aba PORTO (ABA): uma linha por paciente, com TOTAL MÊS e Nº GUIAS."""
    linhas = _ler_aba(caminho, ABA)
    i_cab, cab = _cabecalho(linhas)

    def col(*nomes):
        for n in nomes:
            if _norm(n) in cab:
                return cab.index(_norm(n))
        raise ValueError(f"Coluna '{nomes[0]}' não encontrada na aba {ABA}.")

    c_und, c_pac, c_cart = col("UND", "UND. AUT."), col("PACIENTE"), col("CARTEIRINHA")
    c_mes, c_guias = col("TOTAL MÊS", "TOTAL MES"), col("Nº GUIAS", "N GUIAS", "GUIAS")
    c_aut = cab.index(_norm("TOTAL AUT.")) if _norm("TOTAL AUT.") in cab else None
    c_meia = cab.index(_norm("1/2")) if _norm("1/2") in cab else None

    regs = []
    for n, row in enumerate(linhas[i_cab + 1:], start=i_cab + 2):
        row = row + [None] * (len(cab) - len(row))
        pac = " ".join(str(row[c_pac] or "").split())
        if not pac:
            continue
        und = _unidade(row[c_und])
        if unidade and und != _unidade(unidade):
            continue
        total = _num(row[c_mes])
        if total is None and c_aut is not None and _num(row[c_aut]) is not None:
            total = _num(row[c_aut]) * 4 + (_num(row[c_meia]) or 0 if c_meia is not None else 0)
        regs.append({
            "aba": ABA, "linha": n, "unidade": und, "paciente": pac,
            "carteirinha": _digitos(row[c_cart]),
            "guias": separar_guias(row[c_guias]),
            "guia_texto": str(row[c_guias] or "").strip(),
            "procedimento": "ABA (todas as terapias)",
            "qtd_planilha": int(total) if total is not None else None,
        })
    return regs


def ler_porto_conv(caminho, unidade=None):
    """Aba PORTO CONV: uma linha por especialidade, com Grupo de Procedimento, Quantidade e GUIAS."""
    linhas = _ler_aba(caminho, ABA_CONV)
    i_cab, cab = _cabecalho(linhas)

    def col(*nomes, obrigatoria=True):
        for n in nomes:
            if _norm(n) in cab:
                return cab.index(_norm(n))
        if obrigatoria:
            raise ValueError(f"Coluna '{nomes[0]}' não encontrada na aba {ABA_CONV}.")
        return None

    c_pac, c_proc, c_cart = col("PACIENTE"), col("GRUPO DE PROCEDIMENTO", "PROCEDIMENTO"), col("CARTEIRINHA")
    c_qt, c_guias = col("QUANTIDADE", "QNTD"), col("GUIAS", "GUIA")
    c_und = col("UND", "UND. AUT.", "UNIDADE", obrigatoria=False)
    if c_und is None:
        c_und = 0  # na planilha a 1ª coluna (sem título) é a unidade

    regs = []
    for n, row in enumerate(linhas[i_cab + 1:], start=i_cab + 2):
        row = row + [None] * (len(cab) - len(row))
        pac = " ".join(str(row[c_pac] or "").split())
        if not pac:
            continue
        und = _unidade(row[c_und])
        if unidade and und != _unidade(unidade):
            continue
        qt = _num(row[c_qt])
        proc = " ".join(str(row[c_proc] or "").split())
        regs.append({
            "aba": ABA_CONV, "linha": n, "unidade": und, "paciente": pac,
            "carteirinha": _digitos(row[c_cart]),
            "guias": separar_guias(row[c_guias]),
            "guia_texto": str(row[c_guias] or "").strip(),
            "procedimento": proc,
            "categoria": categoria(proc),
            "qtd_planilha": int(qt) if qt is not None else None,
        })
    return regs


def ler_registros(caminho, unidade=None, abas=ABAS):
    regs = []
    for aba in abas:
        regs += ler_porto(caminho, unidade) if aba == ABA else ler_porto_conv(caminho, unidade)
    return regs


# ----------------------------------------------------------------- conferência

def _mesmo_paciente(reg, dados):
    """(True/False/None, texto). None = não deu para conferir."""
    cart_portal = _digitos(dados.get("carteira")) or _digitos(dados.get("beneficiary_card"))
    nome_portal = dados.get("beneficiario") or dados.get("beneficiary_name") or ""
    cart_plan = reg["carteirinha"]
    if cart_plan and cart_portal:
        n = min(len(cart_plan), len(cart_portal), 14)
        if n >= 10 and cart_plan[:n] == cart_portal[:n]:
            return True, ""
    if nome_portal:
        sim = difflib.SequenceMatcher(None, _norm(reg["paciente"]), _norm(nome_portal)).ratio()
        if sim >= 0.85:
            if cart_plan and cart_portal:
                return True, f"carteirinha da planilha ({cart_plan}) diferente da guia ({cart_portal})"
            return True, ""
        return False, f"a guia é de {nome_portal}"
    return None, ""


def conferir(reg, consultas):
    """consultas: {guia: dados ou {"erro"}}. Devolve a linha do resultado."""
    linha = {
        "Aba": reg["aba"], "Linha": reg["linha"], "Unidade": reg["unidade"], "Paciente": reg["paciente"],
        "Carteirinha": reg["carteirinha"], "Procedimento planilha": reg["procedimento"],
        "Guia(s)": " / ".join(reg["guias"]) or reg["guia_texto"],
        "Qtd planilha": reg["qtd_planilha"],
    }
    if not reg["guias"]:
        linha.update({"Situação": "SEM GUIA",
                      "Divergências": f"sem número de guia na planilha ({reg['guia_texto'] or 'vazio'})"})
        return linha

    erros, avisos, sem_dados = [], [], []
    for g in reg["guias"]:
        if len(g) != 8:
            avisos.append(f"nº de guia {g} tem {len(g)} dígitos (confira a digitação)")

    status, senhas, datas, procs_txt = [], [], [], []
    total_portal = 0
    qt_conv = 0
    cats_guia = []
    for g in reg["guias"]:
        d = consultas.get(g) or {"erro": "não consultada"}
        if d.get("erro"):
            sem_dados.append(f"guia {g}: {d['erro']}")
            continue
        status.append(d.get("status") or "?")
        senhas.append(d.get("senha") or "")
        datas.append(d.get("data_autorizacao") or "")
        if d.get("status") and _norm(d["status"]) != "AUTORIZADO":
            avisos.append(f"guia {g} está {d['status']}")
        ok, msg = _mesmo_paciente(reg, d)
        if ok is False:
            erros.append(f"guia {g}: {msg}")
        elif msg:
            avisos.append(f"guia {g}: {msg}")
        procs = d.get("procedimentos") or []
        if not procs:
            sem_dados.append(f"guia {g}: não consegui ler os procedimentos do PDF")
        for p in procs:
            procs_txt.append(f"{g}: {p['codigo']} {p['descricao']} ({p['qt_autorizada']} aut.)")
            total_portal += p["qt_autorizada"]
            if reg["aba"] == ABA_CONV:
                cat = categoria(p["descricao"])
                cats_guia.append((cat, p))

    if reg["aba"] == ABA_CONV and cats_guia:
        esperado = reg.get("categoria")
        if not esperado:
            avisos.append(f"não reconheci a terapia do Grupo de Procedimento ('{reg['procedimento'] or 'vazio'}')")
            qt_conv = sum(p["qt_autorizada"] for _, p in cats_guia)
        else:
            iguais = [p for c, p in cats_guia if c == esperado]
            if not iguais and esperado == "psicoped":
                iguais = [p for c, p in cats_guia if c == "psico"]  # psicopedagogia é autorizada como psicologia
            outros = [p for c, p in cats_guia if p not in iguais]
            for p in outros:
                erros.append(f"procedimento da guia é {p['descricao']}, planilha é {reg['procedimento']}")
            if any("ABA" in _norm(p["descricao"]).split() for p in iguais):
                erros.append("a guia é de ABA, mas a linha é convencional")
            qt_conv = sum(p["qt_autorizada"] for p in iguais)
        total_portal = qt_conv

    linha["Status portal"] = " / ".join(dict.fromkeys(status))
    linha["Senha(s)"] = " / ".join(s for s in senhas if s)
    linha["Data aut. portal"] = " / ".join(d for d in datas if d)
    linha["Procedimentos portal"] = "; ".join(procs_txt)

    if sem_dados:
        linha["Qtd portal"] = total_portal if procs_txt else None
        linha["Situação"] = "ERRO"
        linha["Divergências"] = "; ".join(erros + avisos + sem_dados)
        return linha

    linha["Qtd portal"] = total_portal
    if reg["qtd_planilha"] is None:
        avisos.append("quantidade vazia na planilha")
    elif total_portal != reg["qtd_planilha"]:
        rot = "TOTAL MÊS" if reg["aba"] == ABA else "Quantidade"
        erros.append(f"{rot} {reg['qtd_planilha']} x portal {total_portal}")

    linha["Situação"] = "DIVERGENTE" if erros else ("AVISO" if avisos else "OK")
    linha["Divergências"] = "; ".join(list(dict.fromkeys(erros + avisos)))
    return linha


# ----------------------------------------------------------------- resultado

CAB = ["Aba", "Linha", "Unidade", "Paciente", "Carteirinha", "Procedimento planilha", "Guia(s)",
       "Status portal", "Senha(s)", "Data aut. portal", "Procedimentos portal",
       "Qtd planilha", "Qtd portal", "Situação", "Divergências"]
CAB_GUIAS = ["Guia", "Paciente (planilha)", "Beneficiário (portal)", "Carteira", "Senha", "Data autorização",
             "Validade senha", "Status", "Código", "Descrição", "Qt. solicitada", "Qt. autorizada", "Erro"]
CORES = {"OK": "C6EFCE", "AVISO": "FFEB9C", "SEM GUIA": "D9D9D9", "ERRO": "FFEB9C", "DIVERGENTE": "F8CBAD"}


def gravar(linhas, consultas, registros, destino):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Resultado"
    ws.append(CAB)
    for r in linhas:
        ws.append([r.get(k) for k in CAB])
        cel = ws.cell(ws.max_row, CAB.index("Situação") + 1)
        cel.fill = PatternFill("solid", fgColor=CORES.get(cel.value, "F8CBAD"))
        if r.get("Situação") == "DIVERGENTE" and r.get("Qtd portal") != r.get("Qtd planilha"):
            ws.cell(ws.max_row, CAB.index("Qtd portal") + 1).fill = PatternFill("solid", fgColor="F8CBAD")

    wg = wb.create_sheet("Guias")
    wg.append(CAB_GUIAS)
    pac_da_guia = defaultdict(list)
    for reg in registros:
        for g in reg["guias"]:
            pac_da_guia[g].append(reg["paciente"])
    for g, d in consultas.items():
        pac = " / ".join(dict.fromkeys(pac_da_guia.get(g, [])))
        if d.get("erro"):
            wg.append([g, pac] + [""] * (len(CAB_GUIAS) - 3) + [d["erro"]])
            continue
        base = [g, pac, d.get("beneficiario") or d.get("beneficiary_name"), d.get("carteira"), d.get("senha"),
                d.get("data_autorizacao"), d.get("validade_senha"), d.get("status")]
        for p in d.get("procedimentos") or [{}]:
            wg.append(base + [p.get("codigo"), p.get("descricao"), p.get("qt_solicitada"), p.get("qt_autorizada"), ""])

    for planilha, cab in ((ws, CAB), (wg, CAB_GUIAS)):
        for cel in planilha[1]:
            cel.font = Font(bold=True, color="FFFFFF")
            cel.fill = PatternFill("solid", fgColor="002060")
            cel.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
        planilha.row_dimensions[1].height = 30
        planilha.freeze_panes = "E2" if planilha is ws else "B2"
        planilha.auto_filter.ref = planilha.dimensions
        for i, k in enumerate(cab, start=1):
            letra = openpyxl.utils.get_column_letter(i)
            larg = max([len(str(c.value or "")) for c in planilha[letra][1:]] + [len(k), 8])
            planilha.column_dimensions[letra].width = min(larg + 2, 70 if k in ("Divergências", "Procedimentos portal") else 42)
    wb.save(destino)
    return destino


# ----------------------------------------------------------------- portal

class LoginNecessario(Exception):
    pass


class GuiaNaoEncontrada(Exception):
    pass


class Aba:
    """Uma aba do portal. Cada aba guarda só as respostas que ela mesma recebeu."""

    def __init__(self, page, nome, log, parar):
        self.page, self.nome, self.log, self.parar = page, nome, log, parar
        self._status = []      # respostas da pesquisa (/v1/status)
        self._pdfs = []        # respostas do Imprimir (print-legacy)
        page.on("response", self._ouvir)

    def _ouvir(self, resp):
        if "/v1/status" in resp.url and resp.request.method == "POST":
            self._status.append(resp)
        elif "print-legacy" in resp.url:
            self._pdfs.append(resp)

    async def _esperar(self, ms):
        await self.page.wait_for_timeout(ms)

    async def _visivel(self, loc):
        try:
            return await loc.count() > 0 and await loc.first.is_visible()
        except Exception:
            return False

    # ---- login
    async def na_tela_de_login(self):
        return await self._visivel(self.page.locator("#username-signIn-input"))

    async def logado(self):
        if await self.na_tela_de_login():
            return False
        return await self._visivel(self.page.get_by_text("Código Prestador"))

    # ---- tela de pesquisa
    async def fechar_avisos(self):
        """Fecha o aviso de boas-vindas e alertas de 'OK' que o portal mostrar."""
        ok = self.page.locator(".swal2-popup button.swal2-confirm")
        if await self._visivel(ok):
            try:
                await ok.first.click()
            except Exception:
                pass
        if await self._visivel(self.page.get_by_text("Bem-Vindo(a) ao Portal")):
            await self.page.keyboard.press("Escape")
            await self._esperar(400)

    def _campo_data(self, n):
        return self.page.get_by_placeholder("DD/MM/YYYY").nth(n)

    async def ir_para_pesquisa(self, recarregar=False):
        if not recarregar and await self._visivel(self._campo_data(0)):
            return
        await self.page.goto(URL_AUTORIZACAO, wait_until="domcontentloaded")
        for _ in range(60):
            await self.fechar_avisos()
            if await self.na_tela_de_login() or await self._visivel(self.page.get_by_text("Página não encontrada")):
                raise LoginNecessario()
            if await self._visivel(self._campo_data(0)):
                return
            await self._esperar(500)
        raise RuntimeError("não achei a tela 'Pesquisar Autorização' no portal")

    async def _digitar_data(self, n, data):
        alvo = data.strftime("%d/%m/%Y")
        campo = self._campo_data(n)
        if await campo.input_value() == alvo:
            return
        for _ in range(3):
            await campo.click()
            await campo.press("Control+a")
            await campo.press("Backspace")
            await campo.press_sequentially(alvo, delay=10)
            if await campo.input_value() == alvo:
                return
        raise RuntimeError(f"não consegui preencher a data {alvo} no portal")

    async def _digitar(self, rotulo, valor):
        campo = self.page.get_by_label(rotulo, exact=True)
        if await campo.input_value() == valor:
            return
        await campo.fill(valor)
        if await campo.input_value() != valor:  # campo não aceitou o preenchimento direto: digita
            await campo.click()
            await campo.fill("")
            await campo.press_sequentially(valor, delay=20)

    async def _alerta(self):
        pop = self.page.locator(".swal2-popup")
        if await self._visivel(pop):
            try:
                txt = " ".join((await pop.first.inner_text()).split())
            except Exception:
                txt = "alerta do portal"
            await self.fechar_avisos()
            return txt.strip("! ")
        return None

    async def pesquisar(self, guia, inicio, fim):
        """Pesquisa uma guia num período (até 89 dias). Devolve a lista que o portal recebeu."""
        await self.ir_para_pesquisa()
        await self.fechar_avisos()
        await self._digitar_data(0, inicio)
        await self._digitar_data(1, fim)
        await self._digitar("Nº Guia Prestador", guia)
        await self._digitar("Autorização", "")
        await self._digitar("Carteirinha", "")
        self._status.clear()
        await self.page.get_by_role("button", name="Aplicar", exact=True).click()
        for _ in range(300):  # até ~45 s
            await self._esperar(150)
            for resp in list(self._status):
                try:
                    corpo = json.loads(resp.request.post_data or "{}")
                except Exception:
                    corpo = {}
                if corpo.get("guiaPrestador") == guia:
                    if resp.status in (401, 403):
                        raise LoginNecessario()
                    if resp.status >= 400:
                        raise RuntimeError(f"o portal respondeu erro {resp.status} na pesquisa")
                    lista = await resp.json()
                    return lista if isinstance(lista, list) else (lista.get("data") or [])
            alerta = await self._alerta()
            if alerta:
                raise RuntimeError(f"o portal avisou: {alerta}")
            if await self.na_tela_de_login():
                raise LoginNecessario()
        raise RuntimeError("o portal não respondeu à pesquisa (tempo esgotado)")

    async def baixar_pdf(self, guia, ticket):
        """Expandir -> Imprimir na linha da guia; devolve os bytes do PDF que o portal recebeu."""
        linha = self.page.locator("main table tbody tr", has_text=guia).first
        await linha.wait_for(timeout=15000)
        imprimir = self.page.get_by_role("button", name="Imprimir", exact=True)
        if not await self._visivel(imprimir):
            await linha.get_by_role("button", name="Expandir").first.click()
            await imprimir.first.wait_for(state="visible", timeout=10000)
        self._pdfs.clear()
        await imprimir.first.click()
        try:
            for _ in range(270):  # até ~40 s
                await self._esperar(150)
                for resp in list(self._pdfs):
                    if not ticket or f"authorizationNumber={ticket}" in resp.url:
                        if resp.status in (401, 403):
                            raise LoginNecessario()
                        if resp.status >= 400:
                            raise RuntimeError(f"o portal respondeu erro {resp.status} no Imprimir")
                        return await resp.body()
                alerta = await self._alerta()
                if alerta:
                    raise RuntimeError(f"o portal avisou ao imprimir: {alerta}")
            raise RuntimeError("o PDF da guia não chegou (tempo esgotado)")
        finally:
            await self.fechar_pdf()

    async def fechar_pdf(self):
        try:
            fechar = self.page.get_by_label("Fechar", exact=True)
            if await self._visivel(fechar):
                await fechar.first.click()
            else:
                await self.page.keyboard.press("Escape")
            await self._esperar(150)
        except Exception:
            pass

    async def consultar(self, guia, periodos):
        """Procura a guia em cada período e lê o PDF. Devolve os dados da guia."""
        achadas = []
        for ini, fim in periodos:
            achadas = [e for e in await self.pesquisar(guia, ini, fim) if str(e.get("provider_guide")) == guia]
            if achadas:
                break
        if not achadas:
            txt = " e ".join(f"{a:%d/%m/%Y} a {b:%d/%m/%Y}" for a, b in periodos)
            raise GuiaNaoEncontrada(f"não encontrada no portal entre {txt}")
        e = achadas[0]
        dados = {
            "status": (e.get("status") or "").upper(),
            "senha": e.get("ticket") or "",
            "data_autorizacao": e.get("authorization_date") or "",
            "beneficiary_card": e.get("beneficiary_card") or "",
            "beneficiary_name": e.get("beneficiary_name") or "",
            "tipo": e.get("solicitation_type") or "",
            "procedimentos": [],
        }
        pdf = await self.baixar_pdf(guia, dados["senha"])
        lido = guia_porto_pdf.ler(pdf)
        for k in ("senha", "data_autorizacao", "validade_senha", "carteira", "beneficiario"):
            if lido.get(k):
                dados[k] = lido[k]
        dados["procedimentos"] = lido.get("procedimentos") or []
        return dados


def dividir_periodo(inicio, fim):
    """O portal aceita no máximo 90 dias: quebra em pedaços, do mais recente para o mais antigo."""
    if fim < inicio:
        raise ValueError("A data final é anterior à data inicial.")
    periodos, b = [], fim
    while b >= inicio:
        a = max(inicio, b - dt.timedelta(days=MAX_DIAS))
        periodos.append((a, b))
        b = a - dt.timedelta(days=1)
    return periodos


async def _esperar_login(aba, log, parar):
    if await aba.logado():
        return
    if not await aba.na_tela_de_login():
        try:
            await aba.page.goto(URL_INICIAL, wait_until="domcontentloaded")
            await aba._esperar(1500)
        except Exception:
            pass
    log("AVISO Faça o login no navegador que abriu (CPF e senha). O robô continua sozinho depois.")
    for _ in range(TEMPO_LOGIN_S):
        if parar is not None and parar.is_set():
            raise RuntimeError("Interrompido antes do login.")
        if await aba.logado():
            log("Login OK")
            await aba._esperar(1500)
            return
        await aba._esperar(1000)
    raise RuntimeError("O login não foi feito em 15 minutos.")


async def _abrir_aba_extra(principal, nome, log, parar):
    """Abre uma aba A PARTIR da aba logada (o Chrome copia o login para ela)."""
    ctx = principal.page.context
    page = None
    try:
        async with ctx.expect_page(timeout=15000) as nova:
            # janela separada (não aba de fundo): o Chrome não deixa ela "lenta" por estar escondida
            await principal.page.evaluate("u => window.open(u, '_blank', 'popup,width=1280,height=860')",
                                          URL_AUTORIZACAO)
        page = await nova.value
        aba = Aba(page, nome, log, parar)
        await page.wait_for_load_state("domcontentloaded")
        await aba.ir_para_pesquisa()
        return aba
    except Exception as e:  # noqa: BLE001
        motivo = "abriu sem o login" if isinstance(e, LoginNecessario) else (str(e).splitlines() or ["erro"])[0][:80]
        log(f"AVISO {nome} não pôde ser usada ({motivo}); seguindo sem ela")
        if page is not None:
            try:
                await page.close()
            except Exception:
                pass
        return None


async def _consultar_async(guias, periodos, n_abas, log, parar, progresso):
    from playwright.async_api import async_playwright

    resultado, feitas = {}, [0]
    fila = asyncio.Queue()
    for g in guias:
        fila.put_nowait(g)
    PASTA_DADOS.mkdir(parents=True, exist_ok=True)

    async with async_playwright() as pw:
        perfil = str(PASTA_DADOS / "navegador")
        opcoes = dict(headless=False, locale="pt-BR", viewport={"width": 1366, "height": 860})
        try:  # Google Chrome instalado, se houver; senão o Chromium do Playwright
            ctx = await pw.chromium.launch_persistent_context(perfil, channel="chrome", **opcoes)
        except Exception:
            try:
                ctx = await pw.chromium.launch_persistent_context(perfil, **opcoes)
            except Exception as e:
                if not navegador.falta_navegador(e):
                    raise
                await asyncio.to_thread(navegador.instalar_chromium, log)
                ctx = await pw.chromium.launch_persistent_context(perfil, **opcoes)
        try:
            page = ctx.pages[0] if ctx.pages else await ctx.new_page()
            principal = Aba(page, "[aba 1]", log, parar)
            await page.goto(URL_INICIAL, wait_until="domcontentloaded")
            await _esperar_login(principal, log, parar)

            abas = [principal]
            n_abas = max(1, min(int(n_abas or 1), MAX_ABAS, len(guias)))
            for n in range(2, n_abas + 1):
                aba = await _abrir_aba_extra(principal, f"[aba {n}]", log, parar)
                if aba:
                    abas.append(aba)
            if n_abas > 1:
                log(f"{len(abas)} aba(s) pesquisando ao mesmo tempo ({len(guias)} guia(s))")
            mostrar_aba = len(abas) > 1

            async def trabalhar(aba):
                while not (parar is not None and parar.is_set()):
                    try:
                        g = fila.get_nowait()
                    except asyncio.QueueEmpty:
                        return
                    rot = f"{aba.nome} " if mostrar_aba else ""
                    for tentativa in (1, 2, 3):
                        try:
                            dados = await aba.consultar(g, periodos)
                            resultado[g] = dados
                            procs = ", ".join(f"{p['descricao']} {p['qt_autorizada']}"
                                              for p in dados["procedimentos"]) or "-"
                            log(f"  {rot}guia {g}: {dados['status']} | senha {dados.get('senha')} | {procs}")
                            break
                        except LoginNecessario:
                            if aba is not principal:
                                fila.put_nowait(g)  # devolve a guia para outra aba
                                log(f"AVISO {aba.nome} perdeu o login; as outras abas continuam")
                                try:
                                    await aba.page.close()
                                except Exception:
                                    pass
                                return
                            log("AVISO a sessão do portal caiu: faça o login de novo no navegador")
                            await _esperar_login(aba, log, parar)
                        except GuiaNaoEncontrada as e:
                            resultado[g] = {"erro": str(e)}
                            log(f"AVISO {rot}guia {g}: {e}")
                            break
                        except Exception as e:  # noqa: BLE001
                            if tentativa == 3:
                                resultado[g] = {"erro": (str(e).splitlines() or ["erro"])[0][:200]}
                                log(f"ERRO {rot}guia {g}: {resultado[g]['erro']}")
                            else:
                                await aba.fechar_pdf()
                                try:
                                    await aba.ir_para_pesquisa(recarregar=True)
                                except LoginNecessario:
                                    pass
                                except Exception:
                                    pass
                    resultado.setdefault(g, {"erro": "não consegui consultar (a sessão caiu várias vezes)"})
                    feitas[0] += 1
                    if progresso:
                        progresso(feitas[0], len(guias))

            await asyncio.gather(*(trabalhar(a) for a in abas))
            if not fila.empty() and not (parar is not None and parar.is_set()):
                await trabalhar(principal)  # sobrou guia de aba que caiu
        finally:
            try:
                await ctx.close()
            except Exception:
                pass
    return resultado


def consultar_guias(guias, inicio, fim, log=print, parar=None, progresso=None, n_abas=1):
    """Abre o portal (login manual) e consulta as guias com até n_abas abas.
    Devolve {guia: dados ou {"erro"}}."""
    periodos = dividir_periodo(inicio, fim)
    if len(periodos) > 1:
        log(f"AVISO período maior que 90 dias: cada guia pode ser pesquisada até {len(periodos)} vezes")
    guias = list(dict.fromkeys(guias))
    if not guias:
        return {}
    resultado = asyncio.run(_consultar_async(guias, periodos, n_abas, log, parar, progresso))
    if parar is not None and parar.is_set():
        log("Interrompido pelo usuário.")
    for g in guias:
        resultado.setdefault(g, {"erro": "não consultada (interrompido)"})
    return resultado


def executar(caminho_planilha, inicio, fim, abas=ABAS, unidade=None, limite=0, log=print, parar=None,
             progresso=None, n_abas=1):
    """Lê a planilha, consulta as guias no portal, compara e grava o resultado. Devolve o caminho."""
    registros = ler_registros(caminho_planilha, unidade, abas)
    if limite:
        registros = registros[:limite]
    if not registros:
        raise ValueError("Nenhuma linha encontrada nas abas escolhidas.")
    guias = list(dict.fromkeys(g for r in registros for g in r["guias"]))
    sem = sum(1 for r in registros if not r["guias"])
    log(", ".join(f"{sum(1 for r in registros if r['aba'] == a)} linha(s) da {a}" for a in abas)
        + f"; {len(guias)} guia(s) para consultar" + (f"; {sem} linha(s) sem guia" if sem else ""))

    consultas = consultar_guias(guias, inicio, fim, log=log, parar=parar, progresso=progresso,
                                n_abas=n_abas) if guias else {}

    linhas = []
    for reg in registros:
        linha = conferir(reg, consultas)
        linhas.append(linha)
        quem = f"{reg['paciente']} ({reg['aba']}{' - ' + reg['procedimento'] if reg['aba'] == ABA_CONV else ''})"
        sit = linha["Situação"]
        tag = "OK" if sit == "OK" else ("ERRO" if sit in ("DIVERGENTE", "ERRO") else "AVISO")
        log(f"{tag} {quem}: planilha {reg['qtd_planilha']} x portal {linha.get('Qtd portal')}"
            + (f" - {linha['Divergências']}" if linha.get("Divergências") else ""))

    destino = Path(caminho_planilha).with_name(f"Resultado Porto {dt.datetime.now():%Y-%m-%d %H%M}.xlsx")
    gravar(linhas, consultas, registros, destino)
    cont = defaultdict(int)
    for l in linhas:
        cont[l["Situação"]] += 1
    log("Pronto: " + ", ".join(f"{n} {s}" for s, n in cont.items()) + f". Resultado em: {destino.name}")
    return destino


if __name__ == "__main__":
    import argparse
    ap = argparse.ArgumentParser(description="Confere as guias das abas PORTO e PORTO CONV no portal Porto")
    ap.add_argument("planilha")
    ap.add_argument("inicio", help="dd/mm/aaaa")
    ap.add_argument("fim", help="dd/mm/aaaa")
    ap.add_argument("--limite", type=int, default=0)
    ap.add_argument("--abas", type=int, default=3, help="abas do portal ao mesmo tempo (1 a 5)")
    a = ap.parse_args()
    try:
        executar(a.planilha, para_data(a.inicio), para_data(a.fim), limite=a.limite, n_abas=a.abas)
    except (RuntimeError, ValueError) as e:
        sys.exit(str(e))
