"""
Conferência de faturamento SulAmérica: relatório do ClinicWeb x planilha de conferência x portal.
Vale para a aba SAS (ABA) e para a aba SAS CONV (convencional).

Ligação: coluna FATURAMENTO da aba  <->  coluna "n° guia" do relatório do ClinicWeb.

Batimentos de cada linha:
  - nome faturado = PACIENTE da planilha
  - senha faturada = SENHA / VPP da planilha = senha da GUIA no portal
  - quantidades: SAS -> PSICO MÊS, FONO MÊS, TO MÊS e TOTAL MÊS | SAS CONV -> QNTD da especialidade
  - códigos: estão na tabela da aba, são da terapia certa e foram autorizados na GUIA do portal
  - valor faturado = quantidade x valor da sessão (ABA R$ 90,00 | convencional R$ 57,88)
"""

import datetime as dt
import re
import unicodedata
from collections import defaultdict
from pathlib import Path

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill

import xlsx_seguro

VALOR_SESSAO = 90.0
TERAPIAS = ("psico", "fono", "to")
TITULO = {"psico": "PSICO", "fono": "FONO", "to": "TO"}

# código de faturamento -> (terapia, tipo)
CODIGOS = {
    "66600510": ("psico", "Psicologia Inicial"),
    "66600480": ("psico", "Psicologia sessão"),
    "66600529": ("fono", "Fonoaudiologia Inicial"),
    "66600499": ("fono", "Fonoaudiologia sessão"),
    "66600537": ("to", "Terapia Ocupacional Inicial"),
    "66600502": ("to", "Terapia Ocupacional Sessão"),
}

# convencional (aba SAS CONV); psicopedagogia é faturada como psicologia
CODIGOS_CONV = {
    "50000470": ("psico", "Psicologia (convencional)"),
    "50000616": ("fono", "Fonoaudiologia (convencional)"),
}

REGRAS = {
    "SAS": {"codigos": CODIGOS, "valor": 90.0, "rotulo": "MÊS"},
    "SAS CONV": {"codigos": CODIGOS_CONV, "valor": 57.88, "rotulo": "QNTD"},
}


# ----------------------------------------------------------------- utilidades

def _norm(txt):
    txt = "".join(c for c in unicodedata.normalize("NFD", str(txt or "")) if unicodedata.category(c) != "Mn")
    return " ".join(re.sub(r"[^A-Za-z0-9 ]+", " ", txt).upper().split())


def _num(v):
    """9803.0 -> '9803'; '5058046196' -> '5058046196'; '-' -> ''"""
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    t = re.sub(r"[\s​]", "", str(v))
    return "" if t in ("-", "None") else t


def terapia_do_nome(nome):
    n = _norm(nome)
    if "FONO" in n:
        return "fono"
    if "TERAPIA OCUPACIONAL" in n or re.search(r"\bTO\b", n):
        return "to"
    if "PSICO" in n:  # inclui psicopedagogia
        return "psico"
    return None


def _valor(v):
    """90.0 / '1.190,10' / 'R$ 90,00' -> float"""
    if v is None or v == "":
        return 0.0
    if isinstance(v, (int, float)):
        return float(v)
    t = re.sub(r"[^\d,.-]", "", str(v))
    if "," in t:
        t = t.replace(".", "").replace(",", ".")
    try:
        return float(t)
    except ValueError:
        return 0.0


def reais(v):
    return "R$ " + f"{v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


# ----------------------------------------------------------------- leitura do ClinicWeb

def ler_relatorio(arquivo):
    """Devolve {nº guia: [linhas]} do Relatório de Conferência de Faturamento do ClinicWeb."""
    wb = xlsx_seguro.carregar(arquivo)
    ws = wb.active
    cab, linha_cab = None, 1
    for i, row in enumerate(ws.iter_rows(min_row=1, max_row=10, values_only=True), start=1):
        nomes = [_norm(c) for c in row]
        if "N GUIA" in nomes and "PROCEDIMENTOS" in nomes:
            cab, linha_cab = nomes, i
            break
    if cab is None:
        raise ValueError("Não achei as colunas 'n° guia' e 'Procedimentos' no relatório do ClinicWeb.")
    col = {k: cab.index(k) for k in ("UNIDADE", "N GUIA", "SENHA", "VALIDADE SENHA", "DATA", "PACIENTE",
                                      "CONVENIO PLANO", "PROCEDIMENTOS", "TOTAL R", "STATUS") if k in cab}
    pega = lambda r, k: r[col[k]] if k in col and col[k] < len(r) else None  # noqa: E731

    guias = defaultdict(list)
    for row in ws.iter_rows(min_row=linha_cab + 1, values_only=True):
        n = _num(pega(row, "N GUIA"))
        if not n:
            continue
        procs = [(int(q), cod, nome.strip())
                 for q, cod, nome in re.findall(r"(\d+)\s*x\s*\[(\d+)\]\s*([^,]+)", str(pega(row, "PROCEDIMENTOS") or ""))]
        guias[n].append({
            "n_guia": n,
            "unidade": str(pega(row, "UNIDADE") or ""),
            "senha": _num(pega(row, "SENHA")),
            "validade": str(pega(row, "VALIDADE SENHA") or ""),
            "data": str(pega(row, "DATA") or ""),
            "paciente": " ".join(str(pega(row, "PACIENTE") or "").split()),
            "convenio": str(pega(row, "CONVENIO PLANO") or ""),
            "procedimentos": procs,
            "texto_procedimentos": str(pega(row, "PROCEDIMENTOS") or ""),
            "total": _valor(pega(row, "TOTAL R")),
            "status": str(pega(row, "STATUS") or ""),
        })
    return guias


def ler_resultados_salvos(arquivos):
    """Dos 'Resultado SulAmerica ....xlsx': {guia: {"codigos": set, "senha": str}} (o que der para aproveitar)."""
    out = {}
    for arq in arquivos or []:
        try:
            wb = openpyxl.load_workbook(arq, data_only=True, read_only=True)
        except Exception:
            continue
        if "Procedimentos" not in wb.sheetnames:
            continue
        rows = list(wb["Procedimentos"].iter_rows(values_only=True))
        if not rows:
            continue
        cab = [_norm(c) for c in rows[0]]
        if "SENHA PORTAL" not in cab or "CODIGO" not in cab or "GUIA" not in cab:
            continue  # resultado antigo, sem senha do portal: consulta de novo
        for r in rows[1:]:
            g = _num(r[cab.index("GUIA")])
            if not g:
                continue
            d = out.setdefault(g, {"codigos": set(), "senha": ""})
            d["codigos"].add(_num(r[cab.index("CODIGO")]))
            d["senha"] = d["senha"] or _num(r[cab.index("SENHA PORTAL")])
        wb.close()
    return out


def portal_para_dict(consultas):
    """Saída do robo.consultar_lista -> {guia: {"codigos", "senha"}} ou {"erro"}."""
    out = {}
    for guia, d in (consultas or {}).items():
        if not d or d.get("erro"):
            out[guia] = {"erro": (d or {}).get("erro") or "sem resposta do portal"}
        else:
            out[guia] = {"codigos": {_num(p.get("codigo")) for p in d.get("procedimentos", [])},
                         "senha": _num(d.get("senha"))}
    return out


# ----------------------------------------------------------------- conferência

def conferir_paciente(reg, relatorio, portal):
    """reg: linha da aba SAS ou SAS CONV (robo.ler_registros). Devolve a linha do resultado."""
    plan = reg["planilha"]
    aba = reg.get("aba", "SAS")
    regra = REGRAS.get(aba, REGRAS["SAS"])
    valor_sessao, rot = regra["valor"], regra["rotulo"]
    linha = {
        "Aba": aba, "Linha": reg["linha"], "Unidade": reg.get("unidade"), "Paciente": reg["paciente"],
        "Especialidade": reg.get("especialidade", ""),
        "Guia (portal)": " / ".join(reg["guias"]), "Senha SAS": reg.get("senha", ""),
        "Nº faturamento": " / ".join(reg.get("faturamento") or []),
    }
    erros, avisos = [], []

    if not reg.get("faturamento"):
        motivo = (f"a aba {aba} não tem a coluna FATURAMENTO" if reg.get("tem_coluna_faturamento") is False
                  else f"coluna FATURAMENTO vazia na aba {aba}")
        linha.update({"Situação": "SEM FATURAMENTO", "Divergências": motivo})
        return linha

    faturadas = []
    for n in reg["faturamento"]:
        if n in relatorio:
            faturadas.extend(relatorio[n])
        else:
            erros.append(f"nº guia {n} não encontrado no relatório do ClinicWeb")
    if not faturadas:
        linha.update({"Situação": "DIVERGENTE", "Divergências": "; ".join(erros)})
        return linha

    # nome
    nomes = sorted({f["paciente"] for f in faturadas})
    linha["Paciente faturado"] = " / ".join(nomes)
    if any(_norm(n) != _norm(reg["paciente"]) for n in nomes):
        erros.append(f"nome faturado diferente: {', '.join(nomes)}")

    # senha: faturada x SAS x portal
    senhas_fat = sorted({f["senha"] for f in faturadas if f["senha"]})
    linha["Senha faturada"] = " / ".join(senhas_fat) or "-"
    if not senhas_fat:
        erros.append("guia faturada sem senha")
    elif reg.get("senha") and any(s != reg["senha"] for s in senhas_fat):
        erros.append(f"senha faturada {', '.join(senhas_fat)} diferente da SENHA / VPP da {aba} ({reg['senha']})")

    senhas_portal, codigos_portal, sem_portal = set(), set(), []
    for g in reg["guias"]:
        p = portal.get(g)
        if not p or p.get("erro"):
            sem_portal.append(f"guia {g}: {p.get('erro') if p else 'não consultada'}")
            continue
        if p.get("senha"):
            senhas_portal.add(p["senha"])
        codigos_portal |= {c for c in p.get("codigos", set()) if c}
    linha["Senha portal"] = " / ".join(sorted(senhas_portal)) or "-"
    if sem_portal:
        avisos.append("sem dados do portal (" + "; ".join(sem_portal) + "): senha e códigos não conferidos com o portal")
    elif senhas_portal and senhas_fat and any(s not in senhas_portal for s in senhas_fat):
        erros.append(f"senha faturada {', '.join(senhas_fat)} não é a senha da guia "
                     f"{' / '.join(reg['guias'])} no portal ({', '.join(sorted(senhas_portal))})")

    # quantidades e códigos
    qtd = defaultdict(int)
    codigos_fat = []
    for f in faturadas:
        for q, cod, nome in f["procedimentos"]:
            codigos_fat.append(f"{cod} {nome}")
            info = regra["codigos"].get(cod)
            ter_nome = terapia_do_nome(nome)
            if info is None:
                erros.append(f"código {cod} ({nome}) não está na tabela de códigos da {aba}")
                ter = ter_nome
            else:
                ter = info[0]
                if ter_nome and ter_nome != ter:
                    erros.append(f"código {cod} é de {info[1]}, mas foi faturado como {nome}")
            if codigos_portal and cod not in codigos_portal:
                erros.append(f"código {cod} ({nome}) não foi autorizado na guia do portal "
                             f"(autorizados: {', '.join(sorted(codigos_portal))})")
            if ter:
                qtd[ter] += q
            else:
                qtd["outros"] += q
    linha["Códigos faturados"] = "; ".join(dict.fromkeys(codigos_fat))
    linha["Códigos autorizados (portal)"] = ", ".join(sorted(codigos_portal)) or "-"

    total_fat = sum(qtd.values())
    for t in TERAPIAS:
        linha[f"{TITULO[t]} planilha"] = plan[t]
        linha[f"{TITULO[t]} faturado"] = qtd[t]
        if qtd[t] != plan[t]:
            if reg.get("terapia_unica") and t != reg["terapia_unica"]:
                erros.append(f"{TITULO[t]}: faturado {qtd[t]}, mas a linha é de {reg.get('especialidade')}")
            else:
                erros.append(f"{TITULO[t]}: faturado {qtd[t]} x {TITULO[t]} {rot} {plan[t]}"
                             if aba == "SAS" else f"{TITULO[t]}: faturado {qtd[t]} x QNTD {plan[t]}")
    linha["Total planilha"] = plan["total"]
    linha["Total faturado"] = total_fat
    if total_fat != plan["total"]:
        erros.append(f"TOTAL: faturado {total_fat} x {'TOTAL MÊS' if aba == 'SAS' else 'QNTD'} {plan['total']}")
    if qtd.get("outros"):
        erros.append(f"{qtd['outros']} sessão(ões) de procedimento que não é psico/fono/TO")

    # valor
    valor_fat = sum(f["total"] for f in faturadas)
    valor_esperado = round(plan["total"] * valor_sessao, 2)
    linha["Valor sessão"] = valor_sessao
    linha["Valor esperado"] = valor_esperado
    linha["Valor faturado"] = valor_fat
    if abs(valor_fat - total_fat * valor_sessao) > 0.05:
        erros.append(f"valor faturado {reais(valor_fat)} não é quantidade x {reais(valor_sessao)} "
                     f"({reais(total_fat * valor_sessao)})")
    if abs(valor_fat - valor_esperado) > 0.05:
        erros.append(f"valor faturado {reais(valor_fat)} x esperado {reais(valor_esperado)}")

    linha["Status ClinicWeb"] = " / ".join(sorted({f["status"] for f in faturadas}))
    linha["Situação"] = "DIVERGENTE" if erros else ("AVISO" if avisos else "OK")
    linha["Divergências"] = "; ".join(list(dict.fromkeys(erros)) + avisos)
    return linha


def conferir(arquivo_relatorio, registros, consultar_portal=None, resultados_salvos=None, log=print):
    """
    registros         : linhas da aba SAS da unidade (robo.ler_planilha(planilha, unidade))
    consultar_portal  : função(lista_de_guias) -> {guia: dados do robô}; None = não consulta o portal
    Devolve a lista de linhas do resultado.
    """
    relatorio = ler_relatorio(arquivo_relatorio)
    log(f"{sum(len(v) for v in relatorio.values())} guia(s) no relatório do ClinicWeb")
    com_fat = [r for r in registros if r.get("faturamento")]
    for aba in dict.fromkeys(r.get("aba", "SAS") for r in registros):
        n = sum(1 for r in registros if r.get("aba", "SAS") == aba)
        nf = sum(1 for r in com_fat if r.get("aba", "SAS") == aba)
        log(f"{aba}: {n} linha(s); {nf} com nº de FATURAMENTO preenchido")

    portal = ler_resultados_salvos(resultados_salvos)
    if consultar_portal:
        faltam = sorted({g for r in com_fat for g in r["guias"] if g not in portal})
        if faltam:
            log(f"{len(faltam)} guia(s) para pesquisar no portal")
            portal.update(portal_para_dict(consultar_portal(faltam)))

    resultado = []
    for reg in sorted(registros, key=lambda r: _norm(r["paciente"])):
        linha = conferir_paciente(reg, relatorio, portal)
        resultado.append(linha)
        if linha["Situação"] == "DIVERGENTE":
            log(f"ERRO {reg['paciente']}: {linha['Divergências']}")
        elif linha["Situação"] in ("AVISO", "SEM FATURAMENTO"):
            log(f"AVISO {reg['paciente']}: {linha['Divergências']}")
        else:
            log(f"OK {reg['paciente']}")
    ok = sum(1 for r in resultado if r["Situação"] == "OK")
    div = sum(1 for r in resultado if r["Situação"] == "DIVERGENTE")
    sem = sum(1 for r in resultado if r["Situação"] == "SEM FATURAMENTO")
    log(f"Pronto: {ok} OK, {div} divergente(s), {sem} sem faturamento, "
        f"{len(resultado) - ok - div - sem} com aviso.")
    return resultado


CABECALHO = ["Aba", "Linha", "Unidade", "Paciente", "Especialidade", "Paciente faturado", "Guia (portal)",
             "Nº faturamento",
             "Senha SAS", "Senha faturada", "Senha portal",
             "PSICO planilha", "PSICO faturado", "FONO planilha", "FONO faturado", "TO planilha", "TO faturado",
             "Total planilha", "Total faturado", "Valor sessão", "Valor esperado", "Valor faturado",
             "Códigos faturados", "Códigos autorizados (portal)", "Status ClinicWeb", "Situação", "Divergências"]


def gravar(resultado, destino):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Faturamento"
    ws.append(CABECALHO)
    cores = {"OK": "C6EFCE", "AVISO": "FFEB9C", "SEM FATURAMENTO": "D9D9D9", "DIVERGENTE": "F8CBAD"}
    vermelho = PatternFill("solid", fgColor="F8CBAD")
    for r in resultado:
        ws.append([r.get(k, "") for k in CABECALHO])
        n = ws.max_row
        sit = ws.cell(n, CABECALHO.index("Situação") + 1)
        sit.fill = PatternFill("solid", fgColor=cores.get(sit.value, "F8CBAD"))
        for k in ("Valor sessão", "Valor esperado", "Valor faturado"):
            ws.cell(n, CABECALHO.index(k) + 1).number_format = '"R$" #,##0.00'
        if r.get("Situação") in ("DIVERGENTE", "AVISO"):  # pinta os pares que não batem
            for a, b in (("PSICO planilha", "PSICO faturado"), ("FONO planilha", "FONO faturado"),
                         ("TO planilha", "TO faturado"), ("Total planilha", "Total faturado"),
                         ("Valor esperado", "Valor faturado")):
                if r.get(b, "") != "" and r.get(a) != r.get(b):
                    ws.cell(n, CABECALHO.index(b) + 1).fill = vermelho
    if resultado:
        n = ws.max_row
        ws.append([])
        ws.cell(n + 2, 1, "TOTAL").font = Font(bold=True)
        for k in ("Total planilha", "Total faturado", "Valor esperado", "Valor faturado"):
            letra = openpyxl.utils.get_column_letter(CABECALHO.index(k) + 1)
            c = ws.cell(n + 2, CABECALHO.index(k) + 1, f"=SUM({letra}2:{letra}{n})")
            c.font = Font(bold=True)
            if k.startswith("Valor"):
                c.number_format = '"R$" #,##0.00'
    for cel in ws[1]:
        cel.font = Font(bold=True, color="FFFFFF")
        cel.fill = PatternFill("solid", fgColor="002060")
        cel.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
    ws.row_dimensions[1].height = 32
    ws.freeze_panes = "D2"
    ws.auto_filter.ref = f"A1:{openpyxl.utils.get_column_letter(len(CABECALHO))}{len(resultado) + 1}"
    for i, k in enumerate(CABECALHO, start=1):
        letra = openpyxl.utils.get_column_letter(i)
        larg = max([len(str(c.value or "")) for c in ws[letra][1:]] + [10])
        ws.column_dimensions[letra].width = min(max(larg + 2, 11), 70 if k == "Divergências" else 40)
    wb.save(destino)
    return destino
