"""
União de conferências: junta várias planilhas de conferência (uma por convênio/dia) em uma
única "Conferência geral", no mesmo padrão de 12 colunas.

Para cada arquivo:
  - escolhe a aba certa: se existir "RESULT"/"Resultado" (versão final, ex.: Intermédica), usa ela;
    senão usa "Planilha1" (as outras abas, como Planilha2 / Guias, são ignoradas);
  - lê o cabeçalho e mapeia as colunas por NOME (tolerando variações: "Satus"->"Status",
    "Convênio/Plano"->"Convenio/Plano", "Observação"->"Observacao", etc.);
  - empilha as linhas na geral.

Saída: uma planilha só, com o cabeçalho padrão e todas as linhas juntas.
"""
from __future__ import annotations

import datetime as dt
import re
import unicodedata
from pathlib import Path

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill

import xlsx_seguro

# colunas da geral, na ordem (nomes exatamente como no modelo "Conferência geral")
COLUNAS_GERAL = ["Empresa", "Profissional", "Data da Agenda", "Hora", "Paciente",
                 "Grupo de Procedimento", "Convenio/Plano", "Status da Agenda (combo)",
                 "Observacao", "Status", "Status Detalhado", "Carteirinha"]

# cabeçalho de origem (normalizado) -> coluna da geral
ALIASES = {
    "EMPRESA": "Empresa",
    "PROFISSIONAL": "Profissional",
    "DATA DA AGENDA": "Data da Agenda", "DATA": "Data da Agenda", "DIA": "Data da Agenda",
    "HORA": "Hora",
    "PACIENTE": "Paciente",
    "GRUPO DE PROCEDIMENTO": "Grupo de Procedimento", "PROCEDIMENTO": "Grupo de Procedimento",
    "CONVENIO PLANO": "Convenio/Plano", "CONVENIO": "Convenio/Plano", "PLANO": "Convenio/Plano",
    "STATUS DA AGENDA COMBO": "Status da Agenda (combo)", "STATUS DA AGENDA": "Status da Agenda (combo)",
    "OBSERVACAO": "Observacao", "OBS": "Observacao",
    "SATUS": "Status", "STATUS": "Status",
    "STATUS DETALHADO": "Status Detalhado", "SATUS DETALHADO": "Status Detalhado",
    "CARTEIRINHA": "Carteirinha", "CARTEIRA": "Carteirinha",
}

NAVY = PatternFill("solid", fgColor="002060")
ABAS_IGNORAR = {"PLANILHA2", "GUIAS UNIMED", "GUIAS AMIL", "APOIO"}


def _norm(txt):
    txt = "".join(c for c in unicodedata.normalize("NFD", str(txt or "")) if unicodedata.category(c) != "Mn")
    return " ".join(re.sub(r"[^A-Za-z0-9 ]+", " ", txt).upper().split())


def padronizar_empresa(v):
    """Padroniza a coluna Empresa:
       '01 · Gopoúva (Sede)'                   -> 'UNIDADE 01'
       'HC · Home Care (atendimento em casa)'  -> 'HC'   (só o prefixo antes do separador)
       'THERAPY MINDS'                         -> 'THERAPY MINDS'  (sem separador, mantém)
    """
    t = str(v or "").strip()
    m = re.match(r"^\s*(\d{1,2})(?=\D|$)", t)  # começa com 1-2 dígitos
    if m:
        return f"UNIDADE {int(m.group(1)):02d}"
    m2 = re.match(r"^\s*([^·\-–—]+?)\s*[·\-–—]\s*\S", t)  # tem separador e não é número -> usa o prefixo
    if m2:
        return m2.group(1).strip()
    return t


def _escolher_aba(wb):
    """RESULT/Resultado tem prioridade (versão final); senão Planilha1; senão a 1ª que não for de apoio."""
    por_nome = {_norm(ws.title): ws for ws in wb.worksheets}
    for chave in ("RESULTADO", "RESULT"):
        if chave in por_nome:
            return por_nome[chave]
    if "PLANILHA1" in por_nome:
        return por_nome["PLANILHA1"]
    for ws in wb.worksheets:
        if _norm(ws.title) not in ABAS_IGNORAR:
            return ws
    return wb.active


def _mapa_colunas(ws):
    """{coluna da geral: índice (0-based) na origem}, pelo cabeçalho da linha 1."""
    mapa = {}
    for i, cel in enumerate(ws[1]):
        destino = ALIASES.get(_norm(cel.value))
        if destino and destino not in mapa:
            mapa[destino] = i
    return mapa


def ler_conferencia(caminho):
    """Devolve (nome_aba, [linhas]) onde cada linha é uma lista nas 12 colunas da geral."""
    wb = xlsx_seguro.carregar(caminho)
    ws = _escolher_aba(wb)
    mapa = _mapa_colunas(ws)
    if "Observacao" not in mapa and "Carteirinha" not in mapa:
        raise ValueError(f"'{Path(caminho).name}' (aba {ws.title}): não reconheci as colunas da conferência.")
    linhas = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not any(v not in (None, "") for v in row):
            continue
        linha = []
        for col in COLUNAS_GERAL:
            idx = mapa.get(col)
            val = row[idx] if (idx is not None and idx < len(row)) else None
            if col == "Empresa":
                val = padronizar_empresa(val)
            linha.append(val)
        linhas.append(linha)
    return ws.title, linhas


def combinar(arquivos, destino, log=print):
    """Junta os arquivos numa geral. Devolve o caminho de saída."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Planilha1"
    ws.append(COLUNAS_GERAL)
    total = 0
    for caminho in arquivos:
        nome = Path(caminho).name
        try:
            aba, linhas = ler_conferencia(caminho)
        except Exception as e:  # noqa: BLE001
            log(f"AVISO pulei '{nome}': {e}")
            continue
        for linha in linhas:
            ws.append(linha)
        total += len(linhas)
        log(f"OK {nome} (aba {aba}): {len(linhas)} linha(s)")

    for cel in ws[1]:
        cel.font = Font(bold=True, color="FFFFFF")
        cel.fill = NAVY
        cel.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.freeze_panes = "A2"
    larguras = [22, 34, 14, 8, 30, 24, 28, 20, 18, 14, 22, 20]
    for i, w in enumerate(larguras, start=1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w
    if ws.max_row >= 1:
        ws.auto_filter.ref = f"A1:{openpyxl.utils.get_column_letter(len(COLUNAS_GERAL))}{ws.max_row}"
    wb.save(destino)
    log(f"Pronto: {total} linha(s) de {len(arquivos)} arquivo(s). Geral: {Path(destino).name}")
    return destino


if __name__ == "__main__":
    import argparse
    ap = argparse.ArgumentParser(description="Une conferências (por convênio/dia) numa Conferência geral")
    ap.add_argument("arquivos", nargs="+", help="planilhas de conferência (.xlsx)")
    ap.add_argument("-o", "--saida", default="", help="arquivo de saída (.xlsx)")
    a = ap.parse_args()
    saida = a.saida or f"Conferência geral {dt.datetime.now():%Y-%m-%d %H%M}.xlsx"
    combinar(a.arquivos, saida)
