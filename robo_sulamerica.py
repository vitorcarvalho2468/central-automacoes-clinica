"""
Robô de consulta de guias - SulAmérica Saúde Online Referenciado

Fluxo:
  login -> menu Segurado -> Consulta -> "Número Guia" -> digita guia -> Pesquisar
  -> clica no paciente -> lê a página de detalhes da SP/SADT

Entrada : planilha de conferência (aba SAS), passada na linha de comando
Saída   : Resultado SulAmerica <data>.xlsx (planilha separada; a de conferência não é alterada)
          paginas/<guia>.html/.txt/.png (cópia de cada página, para conferência)

Uso:
  python robo_sulamerica.py "Conferência Setembro de 2026.xlsx"            # todas as guias
  python robo_sulamerica.py "Conferência Setembro de 2026.xlsx" --limite 3 # só as 3 primeiras (teste)

Credenciais: arquivo config.env (veja config.env.exemplo). Nunca coloque a senha no código.
"""

import argparse
import datetime as dt
import queue
import re
import sys
import threading
import unicodedata
from pathlib import Path

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
import navegador  # acerta a pasta do navegador dentro do .exe (antes de importar o playwright)
from playwright.sync_api import sync_playwright, TimeoutError as PWTimeout

ABA = "SAS"
ABA_CONV = "SAS CONV"
ABAS = (ABA, ABA_CONV)
LINHA_CABECALHO = 2

URL_LOGIN = "https://saude.sulamericaseguros.com.br/prestador/login/"
PASTA = Path(__file__).parent
if getattr(sys, "frozen", False):
    # dentro do CentralAutomacoes.exe a pasta do programa é temporária
    PASTA_PAGINAS = Path.home() / "CentralAutomacoes" / "SulAmerica" / "paginas"
else:
    PASTA_PAGINAS = PASTA / "paginas"

# Durante os testes: False = você vê o navegador trabalhando. Depois pode trocar para True.
HEADLESS = False
LENTIDAO_MS = 120  # pausa entre ações, deixa o robô com ritmo "humano" e evita sobrecarregar o site
MAX_ABAS = 5
REPESQUISAS = 4                     # no fim, as guias com erro/não encontradas são pesquisadas mais 4 vezes
ESPERA_MIN_SEM_REGISTRO_MS = 2800   # só aceita "nenhum registro" depois desse tempo da pesquisa


# ----------------------------------------------------------------- utilidades

def ler_config():
    arq = PASTA / "config.env"
    if not arq.exists():
        sys.exit("Crie o arquivo config.env (copie de config.env.exemplo) com código, usuário e senha.")
    cfg = {}
    for linha in arq.read_text(encoding="utf-8").splitlines():
        linha = linha.strip()
        if linha and not linha.startswith("#") and "=" in linha:
            k, v = linha.split("=", 1)
            cfg[k.strip()] = v.strip()
    for chave in ("CODIGO_REFERENCIADO", "USUARIO", "SENHA"):
        if not cfg.get(chave):
            sys.exit(f"Faltou {chave} no config.env")
    return cfg


def _num(v):
    """Converte o valor da célula em número (ou None)."""
    if v is None or v == "":
        return None
    if isinstance(v, (int, float)):
        return v
    try:
        return float(str(v).replace(",", "."))
    except ValueError:
        return None


def _data(v):
    if isinstance(v, dt.datetime):
        return v.date()
    if isinstance(v, dt.date):
        return v
    if isinstance(v, str):
        m = re.search(r"(\d{1,2})/(\d{1,2})/(\d{4})", v)  # aceita "05/10/2026 14:32"
        if m:
            try:
                return dt.date(int(m.group(3)), int(m.group(2)), int(m.group(1)))
            except ValueError:
                return None
    return None


def _unidade(v):
    """Normaliza a unidade: 1.0, "1", "UNIDADE 01" -> "1"."""
    if v is None:
        return None
    m = re.search(r"\d+", str(v))
    return str(int(float(m.group()))) if m else str(v).strip().upper()


def unidades_da_planilha(caminho):
    """Lista as unidades (coluna UND. AUT.) que têm guia na aba SAS."""
    return sorted({r["unidade"] for r in ler_planilha(caminho) if r["unidade"]},
                  key=lambda u: (not u.isdigit(), int(u) if u.isdigit() else 0, u))


def ler_planilha(caminho, unidade=None):
    """Lê a aba SAS e devolve uma lista de linhas com guia(s) e quantidades da planilha.
    Se 'unidade' for informada, devolve só as linhas daquela UND. AUT."""
    # data_only=True: lê o valor calculado das fórmulas (ex.: PSICO MÊS = PSICO * 4)
    wb = openpyxl.load_workbook(caminho, data_only=True, read_only=True)
    if ABA not in wb.sheetnames:
        raise ValueError(f"Aba '{ABA}' não encontrada na planilha.")
    ws = wb[ABA]
    linhas = list(ws.iter_rows(min_row=LINHA_CABECALHO, values_only=True))
    cab = [" ".join(str(c).split()).upper() if c is not None else "" for c in linhas[0]]

    def col(nome):
        if nome not in cab:
            raise ValueError(f"Coluna '{nome}' não encontrada na aba {ABA}.")
        return cab.index(nome)

    c_opc = {k: cab.index(k) for k in ("FATURAMENTO", "SENHA / VPP") if k in cab}  # colunas opcionais
    c = {k: col(k) for k in ("UND. AUT.", "PACIENTE", "CARTEIRINHA", "PSICO", "PSICO MÊS", "FONO", "FONO MÊS",
                             "TO", "TO MÊS", "TOTAL MÊS", "GUIA", "DATA AUT.")}

    registros = []
    for i, row in enumerate(linhas[1:], start=LINHA_CABECALHO + 1):
        row = list(row) + [None] * (len(cab) - len(row))
        guia_bruta = row[c["GUIA"]]
        if guia_bruta is None or str(guia_bruta).strip() in ("", "-"):
            continue
        if isinstance(guia_bruta, float):
            guia_bruta = str(int(guia_bruta))
        guias = re.findall(r"\d{6,}", str(guia_bruta))  # "240857667 / 242104435" -> duas guias
        if not guias:
            continue

        def mes(chave_mes, chave_sem):
            # usa o valor calculado; se a planilha não tiver o cálculo salvo, faz semanal x 4
            v = _num(row[c[chave_mes]])
            if v is None and _num(row[c[chave_sem]]) is not None:
                v = _num(row[c[chave_sem]]) * 4
            return int(v or 0)

        und = _unidade(row[c["UND. AUT."]])
        if unidade is not None and und != _unidade(unidade):
            continue

        psico, fono, to = mes("PSICO MÊS", "PSICO"), mes("FONO MÊS", "FONO"), mes("TO MÊS", "TO")
        total_mes = _num(row[c["TOTAL MÊS"]])
        total_mes = int(total_mes) if total_mes is not None else psico + fono + to

        registros.append({
            "aba": ABA,
            "especialidade": "ABA (psico, fono e TO)",
            "linha": i,
            "unidade": und,
            "paciente": (row[c["PACIENTE"]] or "").strip(),
            "carteirinha": str(row[c["CARTEIRINHA"]] or "").replace("​", "").strip(),
            "guias": guias,
            "guia_texto": str(guia_bruta),
            "faturamento": re.findall(r"\d+", re.sub(r"\.0\b", "", str(row[c_opc["FATURAMENTO"]] or "")))
            if "FATURAMENTO" in c_opc else [],
            "senha": re.sub(r"\.0$", "", str(row[c_opc["SENHA / VPP"]] or "").strip()) if "SENHA / VPP" in c_opc else "",
            "planilha": {
                "psico": psico,
                "fono": fono,
                "to": to,
                "total": total_mes,
                "data_aut": _data(row[c["DATA AUT."]]),
            },
        })
    wb.close()
    return registros


def _cab(c):
    txt = "".join(ch for ch in unicodedata.normalize("NFD", str(c or "")) if unicodedata.category(ch) != "Mn")
    return " ".join(txt.upper().split())


def ler_planilha_conv(caminho, unidade=None):
    """Lê a aba SAS CONV (convencional: uma linha por especialidade, com QNTD do mês)."""
    wb = openpyxl.load_workbook(caminho, data_only=True, read_only=True)
    if ABA_CONV not in wb.sheetnames:
        wb.close()
        raise ValueError(f"Aba '{ABA_CONV}' não encontrada na planilha.")
    linhas = list(wb[ABA_CONV].iter_rows(values_only=True))
    wb.close()
    lin_cab = next((i for i, r in enumerate(linhas[:8]) if r and "PACIENTE" in [_cab(x) for x in r]), None)
    if lin_cab is None:
        raise ValueError(f"Não achei o cabeçalho (coluna PACIENTE) na aba {ABA_CONV}.")
    cab = [_cab(x) for x in linhas[lin_cab]]

    def col(*nomes, obrigatoria=True):
        for n in nomes:
            if _cab(n) in cab:
                return cab.index(_cab(n))
        if obrigatoria:
            raise ValueError(f"Coluna '{nomes[0]}' não encontrada na aba {ABA_CONV}.")
        return None

    c_und = col("UND. AUTORIZADA", "UND. AUT.")
    c_pac, c_cart, c_qt = col("PACIENTE"), col("CARTEIRINHA"), col("QNTD")
    c_esp, c_guia = col("ESPECIALIDADE"), col("GUIA")
    c_senha = col("SENHA / VPP", obrigatoria=False)
    c_fat = col("FATURAMENTO", obrigatoria=False)
    c_data = col("DATA DE SOLICITAÇÃO", obrigatoria=False)

    registros = []
    for i, row in enumerate(linhas[lin_cab + 1:], start=lin_cab + 2):
        row = list(row) + [None] * (len(cab) - len(row))
        guias = re.findall(r"\d{6,}", re.sub(r"\.0\b", "", str(row[c_guia] or "")))
        if not guias or not row[c_pac]:
            continue
        und = _unidade(row[c_und])
        if unidade is not None and und != _unidade(unidade):
            continue
        esp = " ".join(str(row[c_esp] or "").split())
        ter = categoria(esp) if esp else "outros"
        qt = int(_num(row[c_qt]) or 0)
        plan = {"psico": 0, "fono": 0, "to": 0, "total": qt, "data_aut": None}
        if ter in plan:
            plan[ter] = qt
        registros.append({
            "aba": ABA_CONV,
            "especialidade": esp or "-",
            "terapia_unica": ter,
            "linha": i,
            "unidade": und,
            "paciente": " ".join(str(row[c_pac]).split()),
            "carteirinha": re.sub(r"[\s\u200b]", "", str(row[c_cart] or "")),
            "guias": guias,
            "guia_texto": str(row[c_guia]),
            "faturamento": re.findall(r"\d+", re.sub(r"\.0\b", "", str(row[c_fat] or ""))) if c_fat is not None else [],
            "senha": re.sub(r"\.0$", "", str(row[c_senha] or "").strip()) if c_senha is not None else "",
            "data_solicitacao": _data(row[c_data]) if c_data is not None else None,
            "tem_coluna_faturamento": c_fat is not None,
            "planilha": plan,
        })
    return registros


def agrupar_conv(registros):
    """
    No convencional, linhas do mesmo paciente com a MESMA GUIA são uma autorização só
    (ex.: Psicologia 5 + Psicopedagogia 5 = 10 de psicologia na guia e no faturamento).
    Junta essas linhas numa só, somando as quantidades.
    """
    grupos, saida = {}, []
    for r in registros:
        if r.get("aba") != ABA_CONV:
            saida.append(r)
            continue
        chave = tuple(sorted(r["guias"]))
        if chave in grupos:
            g = grupos[chave]
            g["_linhas"].append(r)
            continue
        g = dict(r, planilha=dict(r["planilha"]), _linhas=[r])
        grupos[chave] = g
        saida.append(g)
    for g in grupos.values():
        linhas = g.pop("_linhas")
        if len(linhas) == 1:
            continue
        plan = {"psico": 0, "fono": 0, "to": 0, "total": 0, "data_aut": None}
        for r in linhas:
            for k in ("psico", "fono", "to", "total"):
                plan[k] += r["planilha"][k]
        terapias = {r["terapia_unica"] for r in linhas}
        g.update({
            "planilha": plan,
            "linha": " / ".join(str(r["linha"]) for r in linhas),
            "especialidade": " + ".join(f"{r['especialidade']} ({r['planilha']['total']})" for r in linhas),
            "terapia_unica": terapias.pop() if len(terapias) == 1 else None,
            "faturamento": list(dict.fromkeys(f for r in linhas for f in r["faturamento"])),
            "senha": " / ".join(dict.fromkeys(r["senha"] for r in linhas if r["senha"])),
        })
    return saida


def ler_registros(caminho, unidade=None, abas=ABAS, ignorar_ausentes=True):
    """Linhas das abas escolhidas ('SAS' e/ou 'SAS CONV'), no mesmo formato.
    No convencional, linhas com a mesma GUIA já vêm somadas (agrupar_conv)."""
    regs = []
    for aba in abas:
        try:
            regs += ler_planilha(caminho, unidade) if aba == ABA else ler_planilha_conv(caminho, unidade)
        except ValueError as e:
            if not (ignorar_ausentes and "não encontrada na planilha" in str(e)):
                raise
    return agrupar_conv(regs)


def categoria(descricao):
    d = f" {descricao.upper()} "
    if "FONO" in d:
        return "fono"
    if re.search(r"\bTO\b|TERAPIA OCUPACIONAL", d):
        return "to"
    if "PSICO" in d:  # inclui psicopedagogia, que conta como psicologia
        return "psico"
    return "outros"


def comparar(registro, consultas):
    """Junta o que veio do portal (uma ou mais guias) e compara com a planilha."""
    portal = {"psico": 0, "fono": 0, "to": 0, "outros": 0}
    for dados in consultas:
        for p in dados.get("procedimentos", []):
            portal[categoria(p["descricao"])] += p["qt_autorizada"]
    portal["total"] = portal["psico"] + portal["fono"] + portal["to"] + portal["outros"]

    plan = registro["planilha"]
    if registro.get("terapia_unica"):  # convencional: a linha é de uma especialidade só
        t = registro["terapia_unica"]
        qt_portal = portal.get(t, 0)
        divergencias = [] if qt_portal == plan["total"] else [
            f"{registro.get('especialidade')}: QNTD {plan['total']} x portal {qt_portal}"]
        portal = {k: (portal.get(k, 0) if k == t else 0) for k in ("psico", "fono", "to", "outros")}
        portal["total"] = qt_portal
    else:
        divergencias = [
            f"{nome}: planilha {plan[k]} x portal {portal[k]}"
            for k, nome in (("psico", "Psico"), ("fono", "Fono"), ("to", "TO"), ("total", "Total"))
            if plan[k] != portal[k]
        ]
    datas_portal = [_data(d.get("data_autorizacao")) for d in consultas]
    datas_portal = [d for d in datas_portal if d]
    data_bate = None
    if plan["data_aut"] and datas_portal:
        data_bate = plan["data_aut"] in datas_portal
    return portal, divergencias, datas_portal, data_bate


def achar(page, montar_locator, timeout_ms=15000):
    """Procura um elemento na página e em todos os iframes (portais antigos usam muito iframe)."""
    fim = timeout_ms
    passo = 500
    while fim > 0:
        for frame in page.frames:
            try:
                loc = montar_locator(frame)
                if loc.count() > 0 and loc.first.is_visible():
                    return frame, loc.first
            except Exception:
                pass
        page.wait_for_timeout(passo)
        fim -= passo
    raise RuntimeError("Elemento não encontrado (veja o print de erro na pasta paginas/)")


# ----------------------------------------------------------------- etapas

def fazer_login(page, cfg):
    page.goto(URL_LOGIN, wait_until="domcontentloaded")

    # Fecha o aviso de cookies, se aparecer
    try:
        page.get_by_text("Continuar", exact=True).click(timeout=3000)
    except Exception:
        pass

    # Campos exatos do formulário de login (ids pegos no "Inspecionar" do site)
    frame, codigo = achar(page, lambda f: f.locator("input#code"))
    usuario = frame.locator("input#user")
    senha = frame.locator("input#senha")

    for loc, valor in ((codigo, cfg["CODIGO_REFERENCIADO"]), (usuario, cfg["USUARIO"]), (senha, cfg["SENHA"])):
        loc.click()
        loc.fill("")
        loc.type(valor, delay=50)  # digita tecla a tecla, como uma pessoa

    # O botão chama validarForm(event) no clique
    frame.locator("button#entrarLogin").click()

    # Login ok quando aparece "Dados do Prestador"
    achar(page, lambda f: f.get_by_text("Dados do Prestador"), timeout_ms=30000)
    print("Login OK")


def abrir_consulta(page):
    _, segurado = achar(page, lambda f: f.get_by_text("Segurado", exact=True))
    segurado.hover()
    _, consulta = achar(page, lambda f: f.get_by_text("Consulta", exact=True), timeout_ms=5000)
    consulta.click()
    achar(page, lambda f: f.get_by_text("Número Guia", exact=True), timeout_ms=30000)


class GuiaNaoEncontrada(Exception):
    """A pesquisa voltou sem resultado (guia de outra unidade/login, ou número errado)."""


def _esperar_resultado(page, guia, timeout_ms=30000):
    """Espera a linha da guia aparecer, ou a mensagem de que não há registros."""
    seletor = f"#tab-lista-consulta tbody tr[numero-guia='{guia}'] td:nth-child(3)"
    sem_registro = re.compile(r"nenhum registro|n[ãa]o (foi|foram) encontrad|0 registro\(s\) encontrado|"
                              r"nenhuma solicita[çc][ãa]o", re.I)
    # A mensagem "nenhum registro" pode ser a da pesquisa ANTERIOR, ainda na tela enquanto a nova carrega.
    # Por isso ela só vale depois de um tempo mínimo e se aparecer em 2 leituras seguidas.
    restante, decorrido, vezes_sem_registro = timeout_ms, 0, 0
    while restante > 0:
        viu_sem_registro = False
        for frame in page.frames:
            try:
                loc = frame.locator(seletor)
                if loc.count() > 0 and loc.first.is_visible():
                    return loc.first
                if sem_registro.search(frame.inner_text("body", timeout=1000)):
                    viu_sem_registro = True
            except Exception:
                pass
        vezes_sem_registro = vezes_sem_registro + 1 if viu_sem_registro else 0
        if viu_sem_registro and decorrido >= ESPERA_MIN_SEM_REGISTRO_MS and vezes_sem_registro >= 2:
            raise GuiaNaoEncontrada(f"guia {guia} não encontrada neste login")
        page.wait_for_timeout(700)
        restante -= 700
        decorrido += 700
    raise GuiaNaoEncontrada(f"guia {guia} não apareceu na pesquisa (tempo esgotado)")


def pesquisar_guia(page, guia):
    frame, radio = achar(page, lambda f: f.get_by_text("Número Guia", exact=True))
    radio.click()

    _, campo = achar(
        page,
        lambda f: f.locator("xpath=//*[contains(text(),'Digite o número da guia')]/following::input[1]"),
    )
    campo.fill(guia)

    _, botao = achar(
        page,
        lambda f: f.locator("input[value='Pesquisar'], button:has-text('Pesquisar'), a:has-text('Pesquisar')"),
    )
    botao.click()

    # Linha do resultado que contém o número da guia -> clica no nome do paciente
    # A linha inteira é clicável e tem o atributo numero-guia; clica na célula do nome do segurado
    link = _esperar_resultado(page, guia)

    # O detalhe pode abrir na mesma aba ou em aba nova
    try:
        with page.context.expect_page(timeout=4000) as nova:
            link.click()
        detalhe = nova.value
    except PWTimeout:
        detalhe = page

    achar(detalhe, lambda f: f.get_by_text("Dados do Beneficiário"), timeout_ms=30000)
    return detalhe


def extrair_dados(texto, guia):
    """Lê o texto da página de detalhes da SP/SADT e devolve os campos que interessam."""

    def pega(rotulo, ate=None):
        # valor depois de "Rótulo:" até o próximo rótulo da mesma linha (ou fim da linha)
        fim = rf"(?=\s+{re.escape(ate)}\s*:)" if ate else r"[ \t]*$"
        m = re.search(rf"(?<!\w){re.escape(rotulo)}\s*:[ \t]*(.*?){fim}", texto, flags=re.MULTILINE)
        valor = m.group(1).strip() if m else None
        return valor or None

    cbo = pega("Código e descrição CBO")
    cbo_codigo, cbo_descricao = (cbo.split(" - ", 1) + [None])[:2] if cbo else (None, None)

    # Tabela "Procedimentos e Exames Solicitados" (colunas separadas por tab)
    procedimentos = []
    trecho = texto.split("Procedimentos e Exames Solicitados", 1)
    if len(trecho) == 2:
        for linha in trecho[1].splitlines():
            cols = [c.strip() for c in linha.split("\t")]
            if len(cols) >= 6 and cols[0].isdigit() and cols[3].isdigit() and cols[4].isdigit():
                procedimentos.append({
                    "codigo": cols[0],
                    "seq": cols[1],
                    "descricao": cols[2],
                    "qt_solicitada": int(cols[3]),
                    "qt_autorizada": int(cols[4]),
                    "status": cols[5],
                    "observacao": cols[6] if len(cols) > 6 else "",
                    "tabela": cols[7] if len(cols) > 7 else "",
                })

    m_guia = re.search(r"Número da Guia\s*(\d+)", texto)
    return {
        "guia_pesquisada": guia,
        "numero_guia": m_guia.group(1) if m_guia else None,
        "status": pega("Status da solicitação"),
        "beneficiario": pega("Nome", "Data de Nascimento"),
        "carteira": pega("Carteira do Beneficiário"),
        "data_autorizacao": pega("Data de Autorização", "Data de Validade da Senha"),
        "data_validade": pega("Data de Validade da Senha"),
        "senha": pega("Senha", "Data de Autorização"),
        "conselho": pega("Conselho Profissional", "UF"),
        "uf": pega("UF", "Número no Conselho"),
        "numero_conselho": pega("Número no Conselho"),
        "cbo_codigo": cbo_codigo,
        "cbo_descricao": cbo_descricao,
        "procedimentos": procedimentos,
    }


# Salva HTML/texto/print de cada guia em paginas/ (útil nos testes; depois pode desligar)
SALVAR_PAGINAS = True


def extrair(detalhe, guia):
    textos = []
    for frame in detalhe.frames:
        try:
            textos.append(frame.inner_text("body"))
        except Exception:
            pass
    texto = "\n".join(textos)
    dados = extrair_dados(texto, guia)

    if SALVAR_PAGINAS:
        PASTA_PAGINAS.mkdir(exist_ok=True)
        (PASTA_PAGINAS / f"{guia}.html").write_text(detalhe.content(), encoding="utf-8")
        (PASTA_PAGINAS / f"{guia}.txt").write_text(texto, encoding="utf-8")
        detalhe.screenshot(path=str(PASTA_PAGINAS / f"{guia}.png"), full_page=True)
    return dados


# ----------------------------------------------------------------- resultado

VERMELHO = PatternFill("solid", fgColor="F8CBAD")
VERDE = PatternFill("solid", fgColor="C6EFCE")
AMARELO = PatternFill("solid", fgColor="FFEB9C")
CINZA = PatternFill("solid", fgColor="D9D9D9")


def _fmt_data(d):
    return d.strftime("%d/%m/%Y") if d else ""


def gravar_resultado(linhas_resultado, procedimentos, destino):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Resultado"
    cab = ["Aba", "Linha", "Unidade", "Paciente", "Especialidade", "Carteirinha", "Guia(s)", "Status portal",
           "Data aut. planilha", "Data aut. portal", "Data bate?",
           "Conselho", "UF", "Nº conselho", "CBO", "Descrição CBO",
           "Psico planilha", "Psico portal", "Fono planilha", "Fono portal",
           "TO planilha", "TO portal", "Total planilha", "Total portal",
           "Quantidades batem?", "Divergências", "Pedido médico", "Erro"]
    ws.append(cab)
    for r in linhas_resultado:
        ws.append([r.get(k) for k in cab])
        linha = ws.max_row
        q = ws.cell(linha, cab.index("Quantidades batem?") + 1)
        q.fill = VERDE if q.value == "SIM" else VERMELHO if q.value == "NÃO" else AMARELO
        d = ws.cell(linha, cab.index("Data bate?") + 1)
        if d.value == "NÃO":
            d.fill = VERMELHO
        if r.get("Erro"):
            ws.cell(linha, cab.index("Erro") + 1).fill = AMARELO
        ped = r.get("Pedido médico") or ""
        if ped and not ped.startswith("salvo"):
            ws.cell(linha, cab.index("Pedido médico") + 1).fill = AMARELO

    ws2 = wb.create_sheet("Procedimentos")
    cab2 = ["Linha", "Paciente", "Guia", "Código", "Seq", "Descrição", "Categoria",
            "Qt. solicitada", "Qt. autorizada", "Status", "Observação", "Tab.",
            "Data autorização", "Carteirinha", "Validade senha", "Senha portal"]
    ws2.append(cab2)
    for p in procedimentos:
        ws2.append(p)

    for planilha, cabecalho in ((ws, cab), (ws2, cab2)):
        for cel in planilha[1]:
            cel.font = Font(bold=True)
            cel.fill = CINZA
            cel.alignment = Alignment(wrap_text=True, vertical="center")
        planilha.freeze_panes = "C2"
        planilha.auto_filter.ref = planilha.dimensions
        for i, titulo in enumerate(cabecalho, start=1):
            maior = max([len(str(c.value or "")) for c in planilha[openpyxl.utils.get_column_letter(i)]] + [8])
            planilha.column_dimensions[openpyxl.utils.get_column_letter(i)].width = min(maior + 2, 45)
    wb.save(destino)


def montar_linha(registro, consultas, erros):
    portal, divergencias, datas_portal, data_bate = comparar(registro, consultas)
    plan = registro["planilha"]
    primeira = consultas[0] if consultas else {}

    def juntar(campo):
        vals = []
        for d in consultas:
            v = d.get(campo)
            if v and v not in vals:
                vals.append(v)
        return " / ".join(vals)

    if not consultas:
        batem = "ERRO"
    else:
        batem = "SIM" if not divergencias and not erros else "NÃO" if divergencias else "ERRO"
    return {
        "Aba": registro.get("aba", ABA),
        "Linha": registro["linha"],
        "Unidade": registro.get("unidade"),
        "Paciente": registro["paciente"],
        "Especialidade": registro.get("especialidade", ""),
        "Carteirinha": registro["carteirinha"],
        "Guia(s)": " / ".join(registro["guias"]),
        "Status portal": juntar("status"),
        "Data aut. planilha": _fmt_data(plan["data_aut"]),
        "Data aut. portal": " / ".join(_fmt_data(d) for d in datas_portal),
        "Data bate?": "" if data_bate is None else ("SIM" if data_bate else "NÃO"),
        "Conselho": juntar("conselho"),
        "UF": juntar("uf"),
        "Nº conselho": juntar("numero_conselho"),
        "CBO": juntar("cbo_codigo"),
        "Descrição CBO": juntar("cbo_descricao"),
        "Psico planilha": plan["psico"], "Psico portal": portal["psico"] if consultas else None,
        "Fono planilha": plan["fono"], "Fono portal": portal["fono"] if consultas else None,
        "TO planilha": plan["to"], "TO portal": portal["to"] if consultas else None,
        "Total planilha": plan["total"], "Total portal": portal["total"] if consultas else None,
        "Quantidades batem?": batem,
        "Divergências": "; ".join(divergencias) if consultas else "",
        "Erro": "; ".join(erros),
    }


# ----------------------------------------------------------------- pedido médico (anexos da guia)

SEM_PEDIDO = "sem pedido anexado"

_JS_LINKS = """() => Array.from(document.querySelectorAll('a')).map((a, i) => ({
    i, href: a.href || '', raw: a.getAttribute('href') || '', texto: (a.innerText || '').trim().slice(0, 80),
    classe: a.className || '', visivel: !!(a.offsetParent || a.getClientRects().length)
}))"""


def _tipo_arquivo(dados, content_type=""):
    if dados[:4] == b"%PDF":
        return "pdf"
    if dados[:3] == b"\xff\xd8\xff":
        return "jpg"
    if dados[:8] == b"\x89PNG\r\n\x1a\n":
        return "png"
    if dados[:4] in (b"II*\x00", b"MM\x00*"):
        return "tif"
    if dados[:6] in (b"GIF87a", b"GIF89a"):
        return "gif"
    if dados[:4] == b"RIFF" and dados[8:12] == b"WEBP":
        return "webp"
    if dados[:2] == b"BM":
        return "bmp"
    ct = (content_type or "").lower()
    for chave, ext in (("pdf", "pdf"), ("jpeg", "jpg"), ("jpg", "jpg"), ("png", "png"), ("html", "html")):
        if chave in ct:
            return ext
    return "bin"


def _eh_toggle(link):
    return "bt_recuperar_anexos" in link["classe"] or "anexo" in link["texto"].lower() and (
        "exibir" in link["texto"].lower() or "ocultar" in link["texto"].lower())


def _baixar_link(pagina, frame, indice, link):
    """Baixa o arquivo de um link de anexo. Devolve (bytes, extensão) ou None."""
    contexto = pagina.context
    href = link["href"]
    if href.startswith("http") and not link["raw"].startswith("#") and "javascript" not in href.lower():
        try:
            resp = contexto.request.get(href, timeout=60000)
            if resp.ok:
                dados = resp.body()
                tipo = _tipo_arquivo(dados, resp.headers.get("content-type", ""))
                if dados and tipo != "html":
                    return dados, tipo
        except Exception:
            pass
    alvo = frame.locator("a").nth(indice)
    try:  # link que dispara download
        with pagina.expect_download(timeout=8000) as dl:
            alvo.click()
        caminho = dl.value.path()
        dados = Path(caminho).read_bytes()
        return dados, _tipo_arquivo(dados, dl.value.suggested_filename)
    except Exception:
        pass
    try:  # link que abre o anexo numa aba/janela nova
        with contexto.expect_page(timeout=8000) as nova:
            alvo.click()
        aba = nova.value
        try:
            aba.wait_for_load_state("load", timeout=20000)
        except Exception:
            pass
        try:
            if aba.url.startswith("http"):
                resp = contexto.request.get(aba.url, timeout=60000)
                dados = resp.body()
                tipo = _tipo_arquivo(dados, resp.headers.get("content-type", ""))
                if dados and tipo != "html":
                    return dados, tipo
            return aba.screenshot(full_page=True), "png"  # não deu para pegar o arquivo: tira print
        finally:
            aba.close()
    except Exception:
        return None


def baixar_anexos(pagina):
    """
    Na página de detalhe da guia: clica em "Exibir anexos" e baixa todos os anexos (PDFs e fotos).
    Devolve None se não houver o botão (sem pedido anexado), senão a lista [(bytes, extensão)].
    """
    antes = {}
    for frame in pagina.frames:
        try:
            antes[frame] = frame.evaluate(_JS_LINKS)
        except Exception:
            antes[frame] = []
    toggles = [(f, l) for f, ls in antes.items() for l in ls if _eh_toggle(l) and l["visivel"]]
    if not toggles:
        return None
    for frame, l in toggles:
        try:
            frame.locator("a").nth(l["i"]).click()
        except Exception:
            pass

    def novos():
        achados = []
        for frame in pagina.frames:
            try:
                depois = frame.evaluate(_JS_LINKS)
            except Exception:
                continue
            visiveis_antes = {(l["href"], l["texto"], l["classe"]) for l in antes.get(frame, []) if l["visivel"]}
            for l in depois:
                chave = (l["href"], l["texto"], l["classe"])
                if l["visivel"] and chave not in visiveis_antes and not _eh_toggle(l):
                    achados.append((frame, l))
        return achados

    candidatos = []
    for _ in range(20):  # a lista pode carregar por ajax: espera até ~10 s
        pagina.wait_for_timeout(500)
        candidatos = novos()
        if candidatos:
            pagina.wait_for_timeout(800)  # deixa terminar de carregar a lista toda
            candidatos = novos()
            break

    arquivos, vistos = [], set()
    for frame, l in candidatos:
        chave = l["href"] if l["href"] and not l["raw"].startswith("#") else (l["texto"], l["i"])
        if chave in vistos:
            continue
        vistos.add(chave)
        r = _baixar_link(pagina, frame, l["i"], l)
        if r:
            arquivos.append(r)
    return arquivos


def _nome_arquivo(txt):
    txt = re.sub(r'[\\/:*?"<>|]+', "", str(txt)).strip().rstrip(". -")
    return re.sub(r"\s+", " ", txt)[:90] or "paciente"


def salvar_pedido(pasta, paciente, arquivos):
    """Junta todos os anexos num único PDF '<PACIENTE>.pdf' (só isso na pasta). Devolve o texto para o resultado."""
    from io import BytesIO

    nome = _nome_arquivo(paciente)
    pasta = Path(pasta)
    pasta.mkdir(parents=True, exist_ok=True)
    try:
        from pypdf import PdfWriter
        from PIL import Image

        escritor = PdfWriter()
        for dados, ext in arquivos:
            if ext == "pdf":
                escritor.append(BytesIO(dados))
            else:
                img = Image.open(BytesIO(dados))
                quadros = []
                for k in range(getattr(img, "n_frames", 1)):  # TIFF pode ter várias páginas
                    img.seek(k)
                    quadros.append(img.convert("RGB"))
                buf = BytesIO()
                quadros[0].save(buf, "PDF", save_all=True, append_images=quadros[1:], resolution=150)
                buf.seek(0)
                escritor.append(buf)
        paginas = len(escritor.pages)
        with open(pasta / f"{nome}.pdf", "wb") as f:
            escritor.write(f)
        return f"salvo: {len(arquivos)} arquivo(s), {paginas} página(s)"
    except Exception as e:
        # não deu para juntar: salva o primeiro anexo como veio, ainda só com o nome do paciente
        dados, ext = arquivos[0]
        (pasta / f"{nome}.{ext}").write_bytes(dados)
        return f"salvo só o 1º de {len(arquivos)} anexo(s) (não deu para juntar em PDF: {e})"


# ----------------------------------------------------------------- principal

def consultar(page, cfg, guia, pegar_anexos=False):
    """Consulta uma guia; se a sessão tiver caído, faz login de novo e tenta mais uma vez."""
    for tentativa in (1, 2):
        try:
            abrir_consulta(page)
            detalhe = pesquisar_guia(page, guia)
            dados = extrair(detalhe, guia)
            if pegar_anexos:
                try:
                    dados["anexos"] = baixar_anexos(detalhe)
                except Exception as e:
                    dados["anexos"] = []
                    dados["erro_anexos"] = str(e)
            if detalhe is not page:
                detalhe.close()
            return dados
        except GuiaNaoEncontrada:
            raise  # não adianta refazer login: a guia não existe neste login
        except Exception as e:
            PASTA_PAGINAS.mkdir(exist_ok=True)
            try:
                page.screenshot(path=str(PASTA_PAGINAS / f"ERRO_{guia}.png"), full_page=True)
            except Exception:
                pass
            if tentativa == 2:
                raise
            print(f"  falhou ({e}); refazendo login e tentando de novo...")
            fazer_login(page, cfg)


class GuiaTrocada(Exception):
    """O portal mostrou outra guia (sessão compartilhada entre abas se confundiu)."""


def _consultar_conferindo(page, cfg, guia, pegar_anexos):
    """Consulta e confere se a página lida é mesmo da guia pesquisada (até 3 tentativas)."""
    for tentativa in range(3):
        dados = consultar(page, cfg, guia, pegar_anexos=pegar_anexos)
        lida = (dados.get("numero_guia") or "").strip()
        if not lida or lida == guia:
            return dados
        print(f"  a página veio com a guia {lida} em vez de {guia}; tentando de novo...")
    raise GuiaTrocada(f"o portal mostrou outra guia ({lida}) no lugar da {guia}; use menos abas")


def _processar_fila(page, cfg, fila, resultado, trava, contador, total, pegar_anexos, parar, progresso, log,
                    rotulo):
    while not (parar is not None and parar.is_set()):
        try:
            guia = fila.get_nowait()
        except queue.Empty:
            return
        log(f"{rotulo} guia {guia}")
        try:
            dados = _consultar_conferindo(page, cfg, guia, pegar_anexos)
            log(f"  {rotulo} {guia}: {dados.get('status')} | autorização {dados.get('data_autorizacao')}"
                f" | validade {dados.get('data_validade')}")
        except Exception as e:
            dados = {"erro": str(e)}
            log(f"  ERRO {rotulo} {guia}: {e}")
        with trava:
            resultado[guia] = dados
            contador[0] += 1
            feitas = contador[0]
        if progresso:
            progresso(feitas, total)


def _aba_extra(n, cfg, estado, url_inicial, headless, fila, resultado, trava, contador, total, pegar_anexos,
               parar, progresso, log):
    """Uma aba a mais, em outro navegador, aproveitando a sessão já logada."""
    rotulo = f"[aba {n}]"
    try:
        with sync_playwright() as p:
            nav = _lancar(p, headless, log)
            ctx = nav.new_context(locale="pt-BR", viewport={"width": 1366, "height": 900}, storage_state=estado)
            page = ctx.new_page()
            try:
                page.goto(url_inicial, wait_until="domcontentloaded")
                achar(page, lambda f: f.get_by_text("Dados do Prestador"), timeout_ms=20000)
            except Exception:
                log(f"{rotulo} sessão não aproveitada; fazendo login nesta aba")
                fazer_login(page, cfg)
            _processar_fila(page, cfg, fila, resultado, trava, contador, total, pegar_anexos, parar, progresso, log,
                            rotulo)
            nav.close()
    except Exception as e:
        log(f"AVISO {rotulo} parou ({e}); as outras abas continuam")


_trava_instalacao = threading.Lock()


def _lancar(p, headless, log):
    """Abre o Chromium; se ele não estiver instalado (ex.: primeira vez no .exe), baixa e tenta de novo."""
    try:
        return p.chromium.launch(headless=headless, slow_mo=LENTIDAO_MS)
    except Exception as e:
        if not navegador.falta_navegador(e):
            raise
        with _trava_instalacao:  # várias abas ao mesmo tempo: baixa uma vez só
            try:
                return p.chromium.launch(headless=headless, slow_mo=LENTIDAO_MS)
            except Exception:
                navegador.instalar_chromium(log)
        return p.chromium.launch(headless=headless, slow_mo=LENTIDAO_MS)


def consultar_guias(cfg, guias, headless=None, n_abas=1, pegar_anexos=False, log=print, parar=None,
                    progresso=None):
    """
    Faz login e consulta a lista de guias, usando até n_abas abas ao mesmo tempo.
    Devolve {guia: dados}; as que falharam vêm como {guia: {"erro": "..."}}.
    """
    global print
    print = log
    if headless is None:
        headless = HEADLESS
    guias = list(dict.fromkeys(g for g in guias if g))
    resultado = {}
    if not guias:
        return resultado
    n_abas = max(1, min(int(n_abas or 1), MAX_ABAS, len(guias)))
    fila = queue.Queue()
    for g in guias:
        fila.put(g)
    trava, contador = threading.Lock(), [0]

    with sync_playwright() as p:
        navegador = _lancar(p, headless, log)
        contexto = navegador.new_context(locale="pt-BR", viewport={"width": 1366, "height": 900})
        page = contexto.new_page()
        try:
            fazer_login(page, cfg)
        except Exception as e:
            PASTA_PAGINAS.mkdir(exist_ok=True)
            page.screenshot(path=str(PASTA_PAGINAS / "ERRO_login.png"), full_page=True)
            navegador.close()
            raise RuntimeError(f"Falha no login. Confira código, usuário e senha. ({e})")

        extras = []
        if n_abas > 1:
            log(f"Abrindo {n_abas} abas no portal ({len(guias)} guia(s))")
            estado, url_inicial = contexto.storage_state(), page.url
            for n in range(2, n_abas + 1):
                t = threading.Thread(target=_aba_extra, daemon=True, args=(
                    n, cfg, estado, url_inicial, headless, fila, resultado, trava, contador, len(guias),
                    pegar_anexos, parar, progresso, log))
                t.start()
                extras.append(t)
        _processar_fila(page, cfg, fila, resultado, trava, contador, len(guias), pegar_anexos, parar, progresso,
                        log, "[aba 1]" if n_abas > 1 else "")
        for t in extras:
            t.join()
        _repesquisar(page, cfg, guias, resultado, trava, pegar_anexos, parar, progresso, log)
        navegador.close()

    if parar is not None and parar.is_set():
        log("Interrompido pelo usuário.")
    for g in guias:
        resultado.setdefault(g, {"erro": "não consultada (interrompido)"})
    return resultado


def _repesquisar(page, cfg, guias, resultado, trava, pegar_anexos, parar, progresso, log):
    """No fim da conferência, pesquisa de novo (até REPESQUISAS vezes) as guias que deram erro ou não foram
    achadas. A guia que for achada numa repesquisa sai da lista; as outras continuam com o erro."""
    def pendentes():
        return [g for g in guias if (resultado.get(g) or {}).get("erro")]

    iniciais = pendentes()
    for rodada in range(1, REPESQUISAS + 1):
        lista = pendentes()
        if not lista or (parar is not None and parar.is_set()):
            break
        log(f"AVISO Repesquisa {rodada} de {REPESQUISAS}: {len(lista)} guia(s) não encontrada(s) ou com erro: "
            + ", ".join(lista))
        page.wait_for_timeout(2000)
        fila = queue.Queue()
        for g in lista:
            fila.put(g)
        _processar_fila(page, cfg, fila, resultado, trava, [0], len(lista), pegar_anexos, parar, progresso, log,
                        f"[repesquisa {rodada}]")
    if not iniciais:
        return
    achadas = [g for g in iniciais if not (resultado.get(g) or {}).get("erro")]
    faltam = pendentes()
    if achadas:
        log(f"OK Repesquisa: {len(achadas)} guia(s) achada(s) depois: " + ", ".join(achadas))
    for g in faltam:
        if not (parar is not None and parar.is_set()):
            resultado[g]["erro"] = f"{resultado[g]['erro']} (pesquisada {REPESQUISAS + 1} vezes)"
    if faltam:
        log(f"AVISO {len(faltam)} guia(s) continuam sem resultado depois de {REPESQUISAS + 1} tentativas: "
            + ", ".join(faltam))


def consultar_lista(cfg, guias, headless=None, log=print, parar=None, progresso=None, n_abas=1):
    """Consulta uma lista de guias avulsas (usado pelas agendas e pelo faturamento)."""
    return consultar_guias(cfg, guias, headless=headless, n_abas=n_abas, log=log, parar=parar, progresso=progresso)


def executar(cfg, caminho_planilha, limite=0, headless=None, log=print, parar=None, progresso=None,
             unidade=None, pasta_pedidos=None, abas=(ABA,), n_abas=1, pasta_saida=None):
    """
    Conferência de solicitação: consulta no portal as guias da(s) aba(s) escolhida(s) e compara.
      abas          : ("SAS",), ("SAS CONV",) ou as duas
      n_abas        : quantas abas do portal usar ao mesmo tempo (1 a 5)
      pasta_pedidos : se informada, baixa os anexos (pedido médico) e salva "<PACIENTE>.pdf" nessa pasta
      pasta_saida   : onde gravar o resultado (padrão: a pasta da planilha)
    Devolve o caminho do arquivo de resultado.
    """
    global print
    print = log
    registros = ler_registros(caminho_planilha, unidade, abas, ignorar_ausentes=len(abas) > 1)
    if not registros:
        raise ValueError(f"Nenhuma guia encontrada para a unidade {unidade} em {', '.join(abas)}.")
    if limite:
        registros = registros[:limite]
    guias = list(dict.fromkeys(g for r in registros for g in r["guias"]))
    nome_und = f"unidade {unidade}" if unidade else "todas as unidades"
    por_aba = {a: sum(1 for r in registros if r["aba"] == a) for a in abas}
    log(f"{nome_und}: " + ", ".join(f"{n} linha(s) da {a}" for a, n in por_aba.items())
        + f"; {len(guias)} guia(s) para consultar")

    sufixo = f" Unidade {unidade}" if unidade else ""
    nome = f"Resultado SulAmerica{sufixo} {dt.datetime.now():%Y-%m-%d %H%M}.xlsx"
    destino = (Path(pasta_saida) / nome) if pasta_saida else Path(caminho_planilha).with_name(nome)

    dados_guias = consultar_guias(cfg, guias, headless=headless, n_abas=n_abas, pegar_anexos=bool(pasta_pedidos),
                                  log=log, parar=parar, progresso=progresso)

    linhas_resultado, procedimentos = [], []
    for reg in registros:
        consultas, erros = [], []
        for guia in reg["guias"]:
            dados = dados_guias.get(guia) or {"erro": "não consultada"}
            if dados.get("erro"):
                erros.append(f"guia {guia}: {dados['erro']}")
                continue
            consultas.append(dados)
            for pr in dados["procedimentos"]:
                procedimentos.append([
                    reg["linha"], reg["paciente"], guia, pr["codigo"], pr["seq"], pr["descricao"],
                    categoria(pr["descricao"]), pr["qt_solicitada"], pr["qt_autorizada"],
                    pr["status"], pr["observacao"], pr["tabela"],
                    dados.get("data_autorizacao") or "", reg["carteirinha"],
                    dados.get("data_validade") or "", dados.get("senha") or "",
                ])

        linha = montar_linha(reg, consultas, erros)
        if pasta_pedidos:
            if not consultas:
                linha["Pedido médico"] = "não verificado (erro na consulta)"
            else:
                anexos = [a for d in consultas for a in (d.get("anexos") or [])]
                erros_anexo = [d["erro_anexos"] for d in consultas if d.get("erro_anexos")]
                if anexos:
                    linha["Pedido médico"] = salvar_pedido(pasta_pedidos, reg["paciente"], anexos)
                elif erros_anexo:
                    linha["Pedido médico"] = "erro ao baixar: " + "; ".join(erros_anexo)
                elif all(d.get("anexos") is None for d in consultas):
                    linha["Pedido médico"] = SEM_PEDIDO
                else:
                    linha["Pedido médico"] = "botão 'Exibir anexos' sem arquivos"
            tag = "AVISO" if not linha["Pedido médico"].startswith("salvo") else "  pedido"
            log(f"{tag} {reg['paciente']}: pedido médico - {linha['Pedido médico']}")
        quem = f"{reg['paciente']} ({reg['aba']}{' - ' + reg['especialidade'] if reg['aba'] == ABA_CONV else ''})"
        log(f"{quem} -> quantidades batem: {linha['Quantidades batem?']} {linha['Divergências']}")
        linhas_resultado.append(linha)

    gravar_resultado(linhas_resultado, procedimentos, destino)
    if pasta_pedidos:
        com = sum(1 for r in linhas_resultado if str(r.get("Pedido médico", "")).startswith("salvo"))
        log(f"Pedidos médicos: {com} de {len(linhas_resultado)} linha(s) salvos em {pasta_pedidos}")
    batem = sum(1 for r in linhas_resultado if r["Quantidades batem?"] == "SIM")
    log(f"Pronto: {batem} de {len(linhas_resultado)} linha(s) batem. Resultado em: {destino}")
    return destino


def main():
    parser = argparse.ArgumentParser(description="Consulta guias da aba SAS no portal SulAmérica")
    parser.add_argument("planilha", help="caminho da planilha de conferência (.xlsx)")
    parser.add_argument("--limite", type=int, default=0, help="consultar só as N primeiras linhas (teste)")
    parser.add_argument("--unidade", default=None, help="consultar só as guias dessa UND. AUT.")
    args = parser.parse_args()
    try:
        executar(ler_config(), args.planilha, args.limite, unidade=args.unidade)
    except (RuntimeError, ValueError) as e:
        sys.exit(str(e))


if __name__ == "__main__":
    main()
