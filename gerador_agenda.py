"""
Gerador de agendas SulAmérica - um PDF por paciente.

Entradas:
  - exportação da agenda (.xlsx, aba "Agendamentos": Unidade, Dia, Hora, Paciente, Terapia, Convênio/Plano,
    Carteirinha, Autorização, Status, Observação ...)
  - um ou mais arquivos "Resultado SulAmerica ....xlsx" gerados pelo robô do portal
    (data de autorização e quantidade autorizada de cada terapia, por guia)
  - período (início e fim) e feriados

Regras:
  - só Psicologia (psicopedagogia conta como psicologia), Fonoaudiologia e Terapia Ocupacional
  - sessões de reposição ficam de fora (convênio "Reposição" ou "reposição" na observação)
  - qualquer status entra (atendido, faltou, convênio suspenso...); convênio suspenso gera aviso
  - a carga é semanal: nada muda de semana
  - sessão antes da autorização, na mesma semana dela, vai para os dias seguintes dessa semana;
    sessão de semana anterior à autorização é perdida
  - sábado não é faturado: sessão de sábado vai para seg-sex da mesma semana
  - sessão em feriado é remanejada para outro dia da mesma semana, sem passar de 5 terapias no dia
  - a quantidade de cada terapia não passa do autorizado (corta as últimas); excesso e falta vão para o log
  - terapia sem autorização no portal: entra tudo o que veio na exportação, com aviso
"""

import datetime as dt
import re
import unicodedata
from collections import defaultdict
from pathlib import Path

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment

MAX_POR_DIA = 5
VALOR_SESSAO = 90.0  # R$ por sessão, igual para psico, fono e TO


def reais(v):
    """1080 -> 'R$ 1.080,00'"""
    return "R$ " + f"{v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
TERAPIAS = ("psico", "fono", "to")
TITULO = {"psico": "PSICO", "fono": "FONO", "to": "TO"}
DIA_SEMANA = ["seg", "ter", "qua", "qui", "sex", "sáb", "dom"]


# ----------------------------------------------------------------- utilidades

def _sem_acento(txt):
    return "".join(c for c in unicodedata.normalize("NFD", str(txt)) if unicodedata.category(c) != "Mn").upper()


def _limpa(v):
    return re.sub(r"[\s​‌‍﻿]", "", str(v or ""))


def categoria_terapia(nome):
    n = _sem_acento(nome)
    if "PSICOPED" in n:  # psicopedagogia conta como psicologia
        return "psico"
    if "FONO" in n:
        return "fono"
    if "TERAPIA OCUPACIONAL" in n or re.search(r"\bTO\b", n):
        return "to"
    if "PSICOLOGIA" in n or re.search(r"\bPSICO\b", n):
        return "psico"
    return None


def para_data(v, ano_padrao=None):
    if v is None or v == "":
        return None
    if isinstance(v, dt.datetime):
        return v.date()
    if isinstance(v, dt.date):
        return v
    txt = str(v).strip()
    for fmt in ("%d/%m/%Y", "%d/%m/%y", "%Y-%m-%d"):
        try:
            return dt.datetime.strptime(txt, fmt).date()
        except ValueError:
            pass
    m = re.fullmatch(r"(\d{1,2})/(\d{1,2})", txt)
    if m and ano_padrao:
        return dt.date(ano_padrao, int(m.group(2)), int(m.group(1)))
    # data com horário ou texto junto, ex.: "05/10/2026 14:32" -> só a data
    m = re.search(r"(\d{1,2})/(\d{1,2})/(\d{4})", txt)
    if m:
        try:
            return dt.date(int(m.group(3)), int(m.group(2)), int(m.group(1)))
        except ValueError:
            return None
    return None


def ler_feriados(texto, ano_padrao):
    """'12/10/2026, 02/11' -> {date(2026,10,12), date(2026,11,2)}. Aceita vírgula, ponto e vírgula ou espaço."""
    feriados, invalidos = set(), []
    for parte in re.split(r"[,;\s]+", texto or ""):
        if not parte:
            continue
        d = para_data(parte, ano_padrao)
        if d:
            feriados.add(d)
        else:
            invalidos.append(parte)
    if invalidos:
        raise ValueError("Feriado(s) em formato inválido: " + ", ".join(invalidos) + " (use dd/mm/aaaa)")
    return feriados


def _cabecalho(ws, obrigatorias, linha_max=5):
    """Acha a linha de cabeçalho e devolve {nome_normalizado: índice}."""
    for i, row in enumerate(ws.iter_rows(min_row=1, max_row=linha_max, values_only=True), start=1):
        nomes = {_sem_acento(" ".join(str(c).split())): j for j, c in enumerate(row) if c is not None}
        if all(_sem_acento(o) in nomes for o in obrigatorias):
            return i, nomes
    raise ValueError(f"Não achei as colunas {', '.join(obrigatorias)} na aba '{ws.title}'.")


# ----------------------------------------------------------------- leitura

def ler_autorizacoes(arquivos_resultado):
    """
    Lê os resultados do robô do portal.
    Devolve {guia: {"data": date|None, "qt": {terapia: n}, "carteirinha": str, "paciente": str}}
    """
    guias, vistos = {}, set()
    for arq in arquivos_resultado:
        wb = openpyxl.load_workbook(arq, data_only=True, read_only=True)
        if "Procedimentos" not in wb.sheetnames:
            raise ValueError(f"'{Path(arq).name}' não parece um resultado do robô (falta a aba Procedimentos).")

        # datas por guia vindas da aba Resultado (para arquivos antigos, sem a coluna na aba Procedimentos)
        datas_resultado, cart_resultado = {}, {}
        if "Resultado" in wb.sheetnames:
            ws = wb["Resultado"]
            lin, c = _cabecalho(ws, ["Guia(s)", "Data aut. portal"])
            for row in ws.iter_rows(min_row=lin + 1, values_only=True):
                g = re.findall(r"\d{6,}", str(row[c["GUIA(S)"]] or ""))
                ds = [para_data(x) for x in str(row[c["DATA AUT. PORTAL"]] or "").split("/ ") if x.strip()]
                ds = [d for d in ds if d] or [para_data(row[c["DATA AUT. PORTAL"]])]
                for k, guia in enumerate(g):
                    datas_resultado[guia] = ds[k] if k < len(ds) else (ds[0] if ds else None)
                    if "CARTEIRINHA" in c:
                        cart_resultado[guia] = _limpa(row[c["CARTEIRINHA"]])

        ws = wb["Procedimentos"]
        lin, c = _cabecalho(ws, ["Guia", "Descrição", "Qt. autorizada"])
        for row in ws.iter_rows(min_row=lin + 1, values_only=True):
            guia = _limpa(row[c["GUIA"]])
            if not guia:
                continue
            g = guias.setdefault(guia, {"data": None, "validade": None, "qt": defaultdict(int),
                                        "carteirinha": "", "paciente": ""})
            chave = (guia, row[c["CODIGO"]] if "CODIGO" in c else None, row[c["SEQ"]] if "SEQ" in c else None,
                     row[c["DESCRICAO"]])
            cat = categoria_terapia(row[c["DESCRICAO"]] or "")
            if cat and chave not in vistos:  # a mesma guia pode aparecer em mais de um arquivo/linha
                vistos.add(chave)
                g["qt"][cat] += int(row[c["QT. AUTORIZADA"]] or 0)
            if "VALIDADE SENHA" in c and row[c["VALIDADE SENHA"]]:
                g["validade"] = para_data(row[c["VALIDADE SENHA"]])
            d = para_data(row[c["DATA AUTORIZACAO"]]) if "DATA AUTORIZACAO" in c else None
            g["data"] = d or g["data"] or datas_resultado.get(guia)
            if "CARTEIRINHA" in c and row[c["CARTEIRINHA"]]:
                g["carteirinha"] = _limpa(row[c["CARTEIRINHA"]])
            g["carteirinha"] = g["carteirinha"] or cart_resultado.get(guia, "")
            if "PACIENTE" in c and row[c["PACIENTE"]]:
                g["paciente"] = str(row[c["PACIENTE"]]).strip()
        wb.close()
    return guias


def autorizacoes_do_portal(consultas):
    """Converte o que o robô leu no portal ({guia: dados}) para o formato de ler_autorizacoes."""
    guias = {}
    for guia, d in consultas.items():
        if not d or d.get("erro"):
            continue
        qt = defaultdict(int)
        for pr in d.get("procedimentos", []):
            cat = categoria_terapia(pr.get("descricao", ""))
            if cat:
                qt[cat] += int(pr.get("qt_autorizada") or 0)
        guias[guia] = {
            "data": para_data(d.get("data_autorizacao")),
            "validade": para_data(d.get("data_validade")),
            "qt": qt,
            "carteirinha": _limpa(d.get("carteirinha")),
            "paciente": d.get("beneficiario") or "",
            "status": d.get("status") or "",
        }
    return guias


def unidade_num(txt):
    """'04 · Tucuruvi' -> '4'; 'UNIDADE 01' -> '1'; 3.0 -> '3'."""
    m = re.search(r"\d+", str(txt or ""))
    return str(int(m.group())) if m else None


def unidades_da_agenda(arquivo):
    cont = defaultdict(set)
    for s in ler_agenda(arquivo):
        u = unidade_num(s["unidade"])
        if u and s["terapia"]:
            cont[u].add(s["carteirinha"] or s["paciente"])
    return {u: len(p) for u, p in cont.items()}


def ler_agenda(arquivo):
    wb = openpyxl.load_workbook(arquivo, data_only=True, read_only=True)
    ws = wb["Agendamentos"] if "Agendamentos" in wb.sheetnames else wb.worksheets[0]
    lin, c = _cabecalho(ws, ["Dia", "Paciente", "Terapia", "Carteirinha"])
    pega = lambda row, nome: row[c[nome]] if nome in c and c[nome] < len(row) else None

    sessoes = []
    for n, row in enumerate(ws.iter_rows(min_row=lin + 1, values_only=True), start=lin + 1):
        if not row or not any(row):
            continue
        sessoes.append({
            "linha": n,
            "data": para_data(pega(row, "DIA")),
            "hora": str(pega(row, "HORA") or ""),
            "paciente": " ".join(str(pega(row, "PACIENTE") or "").split()),
            "carteirinha": _limpa(pega(row, "CARTEIRINHA")),
            "terapia_nome": str(pega(row, "TERAPIA") or ""),
            "terapia": categoria_terapia(pega(row, "TERAPIA") or ""),
            "guia": _limpa(pega(row, "AUTORIZACAO")),
            "unidade": str(pega(row, "UNIDADE") or ""),
            "convenio": str(pega(row, "CONVENIO/PLANO") or ""),
            "obs": str(pega(row, "OBSERVACAO") or ""),
            "status": str(pega(row, "STATUS") or ""),
        })
    wb.close()
    return sessoes


def eh_reposicao(s):
    return "REPOSI" in _sem_acento(s["convenio"]) or "REPOSI" in _sem_acento(s["obs"])


# ----------------------------------------------------------------- regras

def dias_uteis(inicio, fim, feriados):
    d, dias = inicio, []
    while d <= fim:
        if d.weekday() < 5 and d not in feriados:
            dias.append(d)
        d += dt.timedelta(days=1)
    return dias


def montar_agenda_paciente(sessoes, autorizacao, inicio, fim, feriados):
    """
    sessoes     : sessões do paciente (já só psico/fono/to, sem reposição, dentro do período)
    autorizacao : {"data": {terapia: date}, "qt": {terapia: n}}
                  terapia sem quantidade autorizada = sem autorização: não filtra data nem corta
    Devolve (agenda {date: {terapia: n}}, ocorrencias {terapia: {...}})
    """
    agenda = defaultdict(lambda: defaultdict(int))
    autorizado = {t: (autorizacao["qt"].get(t) or None) for t in TERAPIAS}
    # sem autorização: vale desde o início do período
    data_min = {t: (autorizacao["data"].get(t) if autorizado[t] else inicio) for t in TERAPIAS}
    oc = {t: {"exportadas": 0, "antes_autorizacao": 0, "remanejadas": 0, "sem_vaga_feriado": 0,
              "remanejadas_aut": 0, "sem_vaga_aut": 0, "remanejadas_sab": 0, "sem_vaga_sab": 0,
              "cortadas": 0, "final": 0, "autorizado": autorizado[t]} for t in TERAPIAS}

    def total_dia(d):
        return sum(agenda[d].values())

    def segunda(d):
        return d - dt.timedelta(days=d.weekday())

    movimentos = []

    def remanejar(s, motivo):
        """Põe a sessão em outro dia útil da MESMA semana (a carga é semanal: nunca muda de semana)."""
        t = s["terapia"]
        seg = segunda(s["data"])
        candidatos = []
        for k in range(5):
            d = seg + dt.timedelta(days=k)
            if d == s["data"] or d in feriados or not (inicio <= d <= fim) or d < data_min[t]:
                continue
            if total_dia(d) >= MAX_POR_DIA:
                continue
            # prefere dia sem essa terapia, depois o dia mais vazio, depois o mais próximo
            candidatos.append((agenda[d][t] > 0, total_dia(d), abs((d - s["data"]).days), d))
        if candidatos:
            d = min(candidatos)[3]
            agenda[d][t] += 1
            movimentos.append(f"{TITULO[t]} {s['data']:%d/%m} -> {d:%d/%m} ({motivo})")
            return True
        movimentos.append(f"{TITULO[t]} {s['data']:%d/%m} sem dia livre na semana ({motivo})")
        return False

    em_feriado, antes_na_semana, fim_de_semana = [], [], []
    for s in sorted(sessoes, key=lambda s: (s["data"], s["hora"])):
        t = s["terapia"]
        oc[t]["exportadas"] += 1
        data_aut = data_min[t]
        if data_aut is None:
            oc[t]["antes_autorizacao"] += 1
            continue
        if s["data"] < data_aut:
            if segunda(s["data"]) == segunda(data_aut):
                antes_na_semana.append(s)  # mesma semana da autorização: vai para os dias seguintes
            else:
                oc[t]["antes_autorizacao"] += 1  # semana anterior à autorização: cobrança perdida
                movimentos.append(f"{TITULO[t]} {s['data']:%d/%m} perdida (semana anterior à autorização "
                                  f"{data_aut:%d/%m})")
            continue
        if s["data"].weekday() >= 5:  # sábado/domingo não é faturado: vai para seg-sex da mesma semana
            fim_de_semana.append(s)
            continue
        if s["data"] in feriados:
            em_feriado.append(s)
            continue
        agenda[s["data"]][t] += 1

    # Remaneja dentro da própria semana: primeiro as anteriores à autorização, depois os feriados
    for s in antes_na_semana:
        t = s["terapia"]
        oc[t]["remanejadas_aut" if remanejar(s, f"antes da autorização {data_min[t]:%d/%m}")
              else "sem_vaga_aut"] += 1
    for s in em_feriado:
        t = s["terapia"]
        oc[t]["remanejadas" if remanejar(s, "feriado") else "sem_vaga_feriado"] += 1
    for s in fim_de_semana:
        t = s["terapia"]
        oc[t]["remanejadas_sab" if remanejar(s, "sábado") else "sem_vaga_sab"] += 1

    # Limita ao autorizado, cortando as últimas
    for t in TERAPIAS:
        total = sum(agenda[d][t] for d in agenda)
        excesso = total - oc[t]["autorizado"] if oc[t]["autorizado"] is not None else 0
        for d in sorted(agenda, reverse=True):
            while excesso > 0 and agenda[d][t] > 0:
                agenda[d][t] -= 1
                excesso -= 1
                oc[t]["cortadas"] += 1
        oc[t]["final"] = sum(agenda[d][t] for d in agenda)

    oc["_movimentos"] = movimentos
    limpa = {d: {t: n for t, n in v.items() if n} for d, v in agenda.items()}
    return {d: v for d, v in limpa.items() if v}, oc


# ----------------------------------------------------------------- PDF

def _ficha(ficha, largura, altura, compacta=False):
    """Monta os elementos (nome, informações e tabela) da agenda de um paciente para caber em largura x altura."""
    from reportlab.lib import colors
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.lib.units import mm
    from reportlab.platypus import Paragraph, Spacer, Table, TableStyle

    paciente, info, dias, agenda, autorizado = (ficha[k] for k in ("paciente", "info", "dias", "agenda", "autorizado"))
    est_nome = ParagraphStyle("nome", fontName="Helvetica-Bold", fontSize=12 if compacta else 15,
                              leading=15 if compacta else 19)
    est_info = ParagraphStyle("info", fontName="Helvetica", fontSize=7.5 if compacta else 9,
                              leading=9.5 if compacta else 12, textColor=colors.HexColor("#333333"))

    meses = {d.month for d in dias}
    linhas = [["", "PSICO", "FONO", "TO"],
              ["Autorizado"] + [("-" if autorizado.get(t) is None else str(autorizado[t])) for t in TERAPIAS]]
    for d in dias:
        rot = f"{d.day}" if len(meses) == 1 else f"{d:%d/%m}"
        rot += f"  {DIA_SEMANA[d.weekday()]}"
        linhas.append([rot] + [" ".join(["|"] * agenda.get(d, {}).get(t, 0)) for t in TERAPIAS])
    qtd = {t: sum(agenda.get(d, {}).get(t, 0) for d in dias) for t in TERAPIAS}
    linhas.append(["Total"] + [str(qtd[t]) for t in TERAPIAS])
    linhas.append([f"Valor ({reais(VALOR_SESSAO)})"] + [reais(qtd[t] * VALOR_SESSAO) for t in TERAPIAS])
    linhas.append(["Valor total", reais(sum(qtd.values()) * VALOR_SESSAO), "", ""])

    cabecalho = (14 if compacta else 24) * mm
    altura_linha = max(11, min(30, (altura - cabecalho) / len(linhas)))
    col_dia = largura * 0.22
    col = (largura - col_dia) / 3
    fonte, fonte_marca = (8, 12) if compacta else (10, 16)
    tabela = Table(linhas, colWidths=[col_dia, col, col, col], rowHeights=altura_linha, repeatRows=2)
    tabela.setStyle(TableStyle([
        ("GRID", (0, 0), (-1, -1), 0.8 if compacta else 1, colors.black),
        ("FONT", (0, 0), (-1, -1), "Helvetica", fonte),
        ("FONT", (0, 0), (-1, 1), "Helvetica-Bold", fonte),
        ("FONT", (0, -3), (-1, -1), "Helvetica-Bold", fonte),
        ("FONT", (1, 2), (-1, -4), "Helvetica-Bold", fonte_marca),
        ("SPAN", (1, -1), (-1, -1)),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING", (0, 0), (-1, -1), 1),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 1),
        ("BACKGROUND", (0, 0), (-1, 1), colors.HexColor("#E8E8E8")),
        ("BACKGROUND", (0, -3), (-1, -2), colors.HexColor("#E8E8E8")),
        ("BACKGROUND", (0, -1), (-1, -1), colors.HexColor("#002060")),
        ("TEXTCOLOR", (0, -1), (-1, -1), colors.white),
    ]))
    return [Paragraph(paciente, est_nome), Paragraph(info, est_info), Spacer(1, 2 * mm if compacta else 4 * mm),
            tabela]


def gerar_pdf(destino, ficha):
    """PDF individual: uma agenda numa folha A4 em pé."""
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.units import mm
    from reportlab.platypus import SimpleDocTemplate

    margem = 12 * mm
    doc = SimpleDocTemplate(str(destino), pagesize=A4, leftMargin=15 * mm, rightMargin=15 * mm,
                            topMargin=margem, bottomMargin=margem, title=f"Agenda - {ficha['paciente']}")
    doc.build(_ficha(ficha, A4[0] - 30 * mm, A4[1] - 2 * margem - 6 * mm))


def gerar_pdf_duas_por_folha(destino, fichas, titulo="Agendas"):
    """Todas as agendas num PDF só, duas por folha (A4 deitado, uma em cada metade, com linha de corte)."""
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.lib.units import mm
    from reportlab.platypus import BaseDocTemplate, Frame, FrameBreak, PageTemplate

    pagina = landscape(A4)
    margem, meio = 8 * mm, 8 * mm
    largura = (pagina[0] - 2 * margem - 2 * meio) / 2
    altura = pagina[1] - 2 * margem

    def linha_de_corte(canvas, doc):
        canvas.saveState()
        canvas.setDash(4, 4)
        canvas.setLineWidth(0.5)
        canvas.setStrokeGray(0.6)
        canvas.line(pagina[0] / 2, 5 * mm, pagina[0] / 2, pagina[1] - 5 * mm)
        canvas.restoreState()

    esquerda = Frame(margem, margem, largura, altura, id="esq", leftPadding=0, rightPadding=0,
                     topPadding=0, bottomPadding=0)
    direita = Frame(pagina[0] / 2 + meio, margem, largura, altura, id="dir", leftPadding=0, rightPadding=0,
                    topPadding=0, bottomPadding=0)
    doc = BaseDocTemplate(str(destino), pagesize=pagina, leftMargin=margem, rightMargin=margem,
                          topMargin=margem, bottomMargin=margem, title=titulo)
    doc.addPageTemplates([PageTemplate(id="duas", frames=[esquerda, direita], onPage=linha_de_corte)])

    elementos = []
    for i, ficha in enumerate(fichas):
        if i:
            elementos.append(FrameBreak())  # a próxima agenda começa na outra metade (ou na próxima folha)
        elementos.extend(_ficha(ficha, largura, altura - 4 * mm, compacta=True))
    doc.build(elementos)


def area_de_trabalho():
    """Pasta da Área de Trabalho do usuário (inclusive quando ela está dentro do OneDrive)."""
    try:
        import ctypes
        from ctypes import wintypes

        class GUID(ctypes.Structure):
            _fields_ = [("Data1", wintypes.DWORD), ("Data2", wintypes.WORD), ("Data3", wintypes.WORD),
                        ("Data4", ctypes.c_ubyte * 8)]

        # FOLDERID_Desktop {B4BFCC3A-DB2C-424C-B029-7FE99A87C641}
        guid = GUID(0xB4BFCC3A, 0xDB2C, 0x424C, (ctypes.c_ubyte * 8)(0xB0, 0x29, 0x7F, 0xE9, 0x9A, 0x87, 0xC6, 0x41))
        caminho = ctypes.c_wchar_p()
        if ctypes.windll.shell32.SHGetKnownFolderPath(ctypes.byref(guid), 0, None, ctypes.byref(caminho)) == 0:
            pasta = Path(caminho.value)
            ctypes.windll.ole32.CoTaskMemFree(caminho)
            if pasta.exists():
                return pasta
    except Exception:
        pass
    casa = Path.home()
    for opcao in (casa / "Desktop", casa / "OneDrive" / "Desktop", casa / "OneDrive" / "Área de Trabalho",
                  casa / "Área de Trabalho"):
        if opcao.exists():
            return opcao
    return casa


def _nome_arquivo(txt):
    txt = re.sub(r'[\\/:*?"<>|]+', "", txt).strip()
    return re.sub(r"\s+", " ", txt)[:90] or "paciente"


# ----------------------------------------------------------------- principal

def gerar(arquivo_agenda, arquivos_resultado, inicio, fim, feriados, pasta_saida=None, log=print,
          unidade=None, consultar_portal=None, pacientes_unidade=None):
    """
    Gera os PDFs e o resumo. Devolve (pasta, lista_do_resumo).
      arquivos_resultado : resultados já salvos do robô (opcional; evita consultar de novo)
      unidade            : nome da unidade (pasta de saída; e filtro pela coluna Unidade da agenda,
                           só quando pacientes_unidade não for informado)
      pacientes_unidade  : pacientes da unidade vindos da aba SAS (UND. AUT.):
                           [{"carteirinha", "paciente", "guias", "planilha": {psico, fono, to, total}}].
                           Todos ganham agenda, atendidos
                           em qualquer unidade.
      consultar_portal   : função(lista_de_guias) -> {guia: dados}; consulta no site as guias que faltam
    """
    if fim < inicio:
        raise ValueError("A data de fim é anterior à data de início.")
    autorizacoes = ler_autorizacoes(arquivos_resultado or [])
    if arquivos_resultado:
        log(f"{len(autorizacoes)} guia(s) lidas dos resultados já salvos")
    sessoes = ler_agenda(arquivo_agenda)
    log(f"{len(sessoes)} linha(s) na exportação da agenda")
    meta = {}  # carteirinha -> dados do paciente vindos da aba SAS
    if pacientes_unidade is not None:
        for p in pacientes_unidade:
            cart = _limpa(p["carteirinha"])
            m = meta.setdefault(cart, {"paciente": " ".join(str(p["paciente"]).split()), "guias": set(),
                                       "planilha": None})
            m["guias"].update(p.get("guias") or [])
            if p.get("planilha"):  # paciente em mais de uma linha da SAS: soma
                if m["planilha"] is None:
                    m["planilha"] = {k: 0 for k in ("psico", "fono", "to", "total")}
                for k in m["planilha"]:
                    m["planilha"][k] += int(p["planilha"].get(k) or 0)
        if not meta:
            raise ValueError(f"Nenhum paciente com UND. AUT. {unidade} na aba SAS da planilha.")
        sessoes = [s for s in sessoes if s["carteirinha"] in meta]
        log(f"{len(meta)} paciente(s) com UND. AUT. {unidade} na aba SAS; "
            f"{len(sessoes)} linha(s) deles na agenda exportada")
    elif unidade:
        sessoes = [s for s in sessoes if unidade_num(s["unidade"]) == unidade_num(unidade)]
        log(f"{len(sessoes)} linha(s) da unidade {unidade}")
        if not sessoes:
            raise ValueError(f"Nenhuma sessão da unidade {unidade} na agenda exportada.")

    ignoradas = defaultdict(int)
    por_paciente = defaultdict(list)
    for s in sessoes:
        if s["data"] is None:
            ignoradas["sem data"] += 1
        elif not (inicio <= s["data"] <= fim):
            ignoradas["fora do período"] += 1
        elif s["terapia"] is None:
            ignoradas["outras terapias"] += 1
        elif eh_reposicao(s):
            ignoradas["reposição"] += 1
        else:
            por_paciente[s["carteirinha"] or s["paciente"]].append(s)
    for motivo, n in ignoradas.items():
        log(f"  {n} sessão(ões) ignorada(s): {motivo}")
    for cart in meta:  # paciente da SAS sem nenhuma sessão válida na agenda também ganha a folha
        por_paciente.setdefault(cart, [])

    if consultar_portal:
        if meta:
            # guias da coluna GUIA da aba SAS (a coluna Autorização da agenda não é usada)
            faltam = {g for m in meta.values() for g in m["guias"]}
        else:
            faltam = {s["guia"] for l in por_paciente.values() for s in l if s["guia"]}
        faltam = sorted(g for g in faltam if g not in autorizacoes)
        log(f"{len(faltam)} guia(s) para consultar no portal")
        consultas = consultar_portal(faltam)
        novas = autorizacoes_do_portal(consultas)
        autorizacoes.update(novas)
        nao_achadas = [g for g in faltam if g not in novas]
        if nao_achadas:
            log(f"AVISO {len(nao_achadas)} guia(s) sem resposta do portal: {', '.join(nao_achadas)}")

    guias_por_cart = defaultdict(set)
    for guia, g in autorizacoes.items():
        if g["carteirinha"]:
            guias_por_cart[g["carteirinha"]].add(guia)

    dias = dias_uteis(inicio, fim, feriados)
    nome_pasta = f"Agendas SulAmerica{' Unidade ' + str(unidade) if unidade else ''} {inicio:%d-%m} a {fim:%d-%m-%Y}"
    pasta = Path(pasta_saida) if pasta_saida else area_de_trabalho() / nome_pasta
    pasta.mkdir(parents=True, exist_ok=True)

    resumo, fichas = [], []

    def nome_de(chave, lista):
        return meta[chave]["paciente"] if chave in meta else lista[0]["paciente"]

    for chave, lista in sorted(por_paciente.items(), key=lambda kv: _sem_acento(nome_de(*kv))):
        paciente = nome_de(chave, lista)
        cart = chave if chave in meta else lista[0]["carteirinha"]
        und_atend = ", ".join(sorted({s["unidade"] for s in lista})) or "-"

        if chave in meta:
            # paciente da aba SAS: vale só a(s) guia(s) da coluna GUIA da planilha
            guias = {g for g in meta[chave]["guias"] if g in autorizacoes}
        else:
            guias = {s["guia"] for s in lista if s["guia"] in autorizacoes}
            if not guias:  # sem guia na exportação: tenta pela carteirinha
                guias = set(guias_por_cart.get(cart, set()))
        ignoradas_guia = []
        for g in sorted(guias):
            a = autorizacoes[g]
            if a.get("validade") and a["validade"] < inicio:
                ignoradas_guia.append(f"guia {g} vencida em {a['validade']:%d/%m/%Y}")
            elif a.get("data") and a["data"] > fim:
                ignoradas_guia.append(f"guia {g} autorizada só em {a['data']:%d/%m/%Y}")
        guias -= {x.split()[1] for x in ignoradas_guia}
        linha = {"Paciente": paciente, "Carteirinha": cart, "Unidade": und_atend,
                 "Guias": " / ".join(sorted(guias)), "Data autorização": "", "Situação": "", "Detalhes": ""}

        qt, datas = defaultdict(int), {}
        for g in guias:
            a = autorizacoes[g]
            for t, n in a["qt"].items():
                if n:
                    qt[t] += n
                    if a["data"] and (t not in datas or a["data"] < datas[t]):
                        datas[t] = a["data"]
        sem_data = [TITULO[t] for t in TERAPIAS if qt.get(t) and t not in datas]
        aut = {"qt": qt, "data": {t: datas.get(t) for t in TERAPIAS}}

        agenda, oc = montar_agenda_paciente(lista, aut, inicio, fim, feriados)

        problemas, avisos = [], []
        if not lista:
            avisos.append("nenhuma sessão de psico/fono/TO desse paciente na agenda exportada no período")
        if ignoradas_guia:
            avisos.append("ignorada(s) por estar fora do período: " + ", ".join(ignoradas_guia))
        if not guias:
            if chave in meta and not meta[chave]["guias"]:
                motivo = "sem número na coluna GUIA da aba SAS"
            elif chave in meta and any(g in autorizacoes for g in meta[chave]["guias"]):
                motivo = "guia da aba SAS fora do período"
            elif chave in meta:
                motivo = "guia da aba SAS não encontrada no portal (" + ", ".join(sorted(meta[chave]["guias"])) + ")"
            else:
                motivo = "nenhuma guia no resultado do portal"
            avisos.append(motivo + ": agenda gerada com tudo o que veio na exportação")
        sem_aut = []
        for t in TERAPIAS:
            o = oc[t]
            if o["autorizado"] is None:
                if o["exportadas"] and guias:
                    avisos.append(f"{TITULO[t]}: sem autorização no portal, {o['final']} sessão(ões) mantida(s)")
                if o["exportadas"]:
                    sem_aut.append(t)
                continue
            if o["antes_autorizacao"]:
                problemas.append(f"{TITULO[t]}: {o['antes_autorizacao']} sessão(ões) em semana anterior à "
                                 f"autorização (cobrança perdida)")
            if o["remanejadas_aut"]:
                avisos.append(f"{TITULO[t]}: {o['remanejadas_aut']} sessão(ões) antes da autorização "
                              f"remanejada(s) na mesma semana")
            if o["sem_vaga_aut"]:
                problemas.append(f"{TITULO[t]}: {o['sem_vaga_aut']} sessão(ões) antes da autorização sem dia livre "
                                 f"na semana")
            if o["cortadas"]:
                problemas.append(f"{TITULO[t]}: EXCESSO - {o['cortadas']} cortada(s) (autorizado {o['autorizado']})")
            if o["final"] < o["autorizado"]:
                problemas.append(f"{TITULO[t]}: FALTA - agendado {o['final']} de {o['autorizado']} autorizado(s)")
        for t in TERAPIAS:
            if oc[t]["sem_vaga_sab"]:
                problemas.append(f"{TITULO[t]}: {oc[t]['sem_vaga_sab']} sessão(ões) de sábado sem dia livre na semana")
            if oc[t]["sem_vaga_feriado"]:
                problemas.append(f"{TITULO[t]}: {oc[t]['sem_vaga_feriado']} sessão(ões) de feriado sem dia livre na semana")
        if sem_data:
            problemas.append("sem data de autorização no resultado para " + ", ".join(sem_data))
        suspensas = sorted({s["data"] for s in lista if "SUSPENS" in _sem_acento(s["status"])})
        if suspensas:
            avisos.append("convênio suspenso em " + ", ".join(f"{d:%d/%m}" for d in suspensas)
                          + " (sessões mantidas na agenda)")

        plan = meta.get(chave, {}).get("planilha")
        if plan and guias:
            dif = []
            for t in TERAPIAS:
                portal_t = oc[t]["autorizado"] or 0
                if plan[t] != portal_t:
                    dif.append(f"{TITULO[t]} planilha {plan[t]} x portal {portal_t}")
            total_portal = sum(oc[t]["autorizado"] or 0 for t in TERAPIAS)
            if plan["total"] != total_portal:
                dif.append(f"TOTAL MÊS planilha {plan['total']} x portal {total_portal}")
            if dif:
                avisos.append("planilha SAS diferente do portal: " + ", ".join(dif))

        excesso = any(oc[t]["cortadas"] for t in TERAPIAS)
        falta = any(oc[t]["autorizado"] is not None and oc[t]["final"] < oc[t]["autorizado"] for t in TERAPIAS)
        if excesso and falta:
            linha["Situação"] = "EXCESSO E FALTA"
        elif excesso:
            linha["Situação"] = "EXCESSO"
        elif falta:
            linha["Situação"] = "FALTA"
        elif problemas:
            linha["Situação"] = "VERIFICAR"
        elif not guias or sem_aut:
            linha["Situação"] = "SEM AUTORIZAÇÃO"
        elif avisos:
            linha["Situação"] = "AVISO"
        else:
            linha["Situação"] = "OK"
        data_aut = min((d for d in aut["data"].values() if d), default=None)
        linha["Data autorização"] = f"{data_aut:%d/%m/%Y}" if data_aut else ""
        linha["Detalhes"] = "; ".join(problemas + avisos)
        for t in TERAPIAS:
            linha[f"{TITULO[t]} planilha"] = plan[t] if plan else ""
            linha[f"{TITULO[t]} autorizado"] = "-" if oc[t]["autorizado"] is None else oc[t]["autorizado"]
            linha[f"{TITULO[t]} agendado"] = oc[t]["final"]
        linha["TOTAL MÊS planilha"] = plan["total"] if plan else ""
        linha["Total autorizado"] = sum(oc[t]["autorizado"] or 0 for t in TERAPIAS) if guias else "-"
        linha["Total agendado"] = sum(oc[t]["final"] for t in TERAPIAS)
        linha["Valor total (R$)"] = linha["Total agendado"] * VALOR_SESSAO
        linha["Remanejadas (feriado)"] = sum(oc[t]["remanejadas"] for t in TERAPIAS)
        linha["Autorização por terapia"] = ", ".join(
            f"{TITULO[t]} {aut['data'][t]:%d/%m}" for t in TERAPIAS if oc[t]["autorizado"] and aut["data"].get(t))
        linha["Remanejamentos"] = "; ".join(oc["_movimentos"])
        linha["Remanejadas (sábado)"] = sum(oc[t]["remanejadas_sab"] for t in TERAPIAS)
        linha["Remanejadas (autorização)"] = sum(oc[t]["remanejadas_aut"] for t in TERAPIAS)
        linha["Perdidas (semana antes da autorização)"] = sum(oc[t]["antes_autorizacao"] for t in TERAPIAS)

        dias_pdf = sorted(set(dias) | set(agenda))
        info = (f"Carteirinha: {cart} &nbsp;|&nbsp; Unidade: {und_atend} &nbsp;|&nbsp; "
                f"Período: {inicio:%d/%m/%Y} a {fim:%d/%m/%Y}<br/>"
                f"Guia(s): {linha['Guias']} &nbsp;|&nbsp; Autorização: {linha['Data autorização']}")
        ficha = {"paciente": paciente, "info": info, "dias": dias_pdf, "agenda": agenda,
                 "autorizado": {t: oc[t]["autorizado"] for t in TERAPIAS}}
        fichas.append(ficha)
        pasta_ind = pasta / "Individuais"
        pasta_ind.mkdir(exist_ok=True)
        arquivo_pdf = pasta_ind / f"{_nome_arquivo(paciente)}.pdf"
        gerar_pdf(arquivo_pdf, ficha)
        linha["PDF"] = f"Individuais/{arquivo_pdf.name}"

        if problemas:
            log(f"ERRO {paciente}: {'; '.join(problemas)}")
        if avisos:
            log(f"AVISO {paciente}: {'; '.join(avisos)}")
        if not problemas and not avisos:
            log(f"OK {paciente}")
        resumo.append(linha)

    if fichas:
        nome = f"Agendas{' Unidade ' + str(unidade) if unidade else ''} - 2 por folha.pdf"
        gerar_pdf_duas_por_folha(pasta / nome, fichas, titulo=nome[:-4])
        log(f"PDF para impressão: {nome} ({(len(fichas) + 1) // 2} folha(s))")
    gravar_resumo(resumo, pasta / "Resumo das agendas.xlsx")
    ok = sum(1 for r in resumo if r["Situação"] == "OK")
    log(f"Pronto: {len(resumo)} agenda(s) gerada(s), {ok} OK e {len(resumo) - ok} com divergência ou aviso. "
        f"Pasta: {pasta}")
    return pasta, resumo


def gravar_resumo(resumo, destino):
    cab = ["Paciente", "Carteirinha", "Unidade", "Guias", "Data autorização",
           "PSICO planilha", "PSICO autorizado", "PSICO agendado",
           "FONO planilha", "FONO autorizado", "FONO agendado",
           "TO planilha", "TO autorizado", "TO agendado",
           "TOTAL MÊS planilha", "Total autorizado", "Total agendado", "Valor total (R$)", "Remanejadas (feriado)",
           "Remanejadas (sábado)", "Remanejadas (autorização)", "Perdidas (semana antes da autorização)",
           "Autorização por terapia", "Remanejamentos", "Situação", "Detalhes", "PDF"]
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Resumo"
    ws.append(cab)
    cores = {"OK": "C6EFCE", "VERIFICAR": "FFEB9C", "AVISO": "FFEB9C", "SEM AUTORIZAÇÃO": "FFEB9C"}
    for r in resumo:
        ws.append([r.get(k, "") for k in cab])
        cel = ws.cell(ws.max_row, cab.index("Situação") + 1)
        cel.fill = PatternFill("solid", fgColor=cores.get(cel.value, "F8CBAD"))
        ws.cell(ws.max_row, cab.index("Valor total (R$)") + 1).number_format = '"R$" #,##0.00'
    if resumo:  # linha final com a soma da unidade
        n = ws.max_row
        col = openpyxl.utils.get_column_letter(cab.index("Valor total (R$)") + 1)
        col_qtd = openpyxl.utils.get_column_letter(cab.index("Total agendado") + 1)
        ws.append([])
        ws.cell(n + 2, 1, "TOTAL DA UNIDADE").font = Font(bold=True)
        ws.cell(n + 2, cab.index("Total agendado") + 1, f"=SUM({col_qtd}2:{col_qtd}{n})").font = Font(bold=True)
        tot = ws.cell(n + 2, cab.index("Valor total (R$)") + 1, f"=SUM({col}2:{col}{n})")
        tot.font = Font(bold=True)
        tot.number_format = '"R$" #,##0.00'
    ref_filtro = f"A1:{openpyxl.utils.get_column_letter(len(cab))}{len(resumo) + 1}"
    for cel in ws[1]:
        cel.font = Font(bold=True)
        cel.fill = PatternFill("solid", fgColor="D9D9D9")
        cel.alignment = Alignment(wrap_text=True, vertical="center")
    ws.freeze_panes = "B2"
    ws.auto_filter.ref = ref_filtro if resumo else ws.dimensions
    for i, _ in enumerate(cab, start=1):
        letra = openpyxl.utils.get_column_letter(i)
        ws.column_dimensions[letra].width = min(max(len(str(c.value or "")) for c in ws[letra]) + 2, 60)
    wb.save(destino)
