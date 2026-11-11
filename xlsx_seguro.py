"""
Leitura tolerante de planilhas .xlsx.

Alguns sistemas (ex.: ClinicWeb) exportam planilhas com a formatação de cores levemente
fora do padrão. O Excel abre normalmente, mas o openpyxl falha com
"expected <class 'openpyxl.styles.fills.Fill'>". Aqui a gente tenta abrir do jeito normal e,
só se falhar, conserta a formatação EM MEMÓRIA (o arquivo original não é alterado e os dados
não mudam).
"""
from __future__ import annotations

import io
import re
import zipfile

import openpyxl

_FILL_PADRAO = '<fill><patternFill patternType="none"/></fill>'


def _reescrever_styles(caminho, transformar) -> io.BytesIO:
    buf = io.BytesIO()
    with zipfile.ZipFile(caminho) as zin, zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            dados = zin.read(item.filename)
            if item.filename == "xl/styles.xml":
                dados = transformar(dados.decode("utf-8")).encode("utf-8")
            zout.writestr(item, dados)
    buf.seek(0)
    return buf


def _consertar_fills_vazios(xml: str) -> str:
    """Nível 1: <fill/> ou <fill></fill> vazios viram 'sem preenchimento'."""
    def sub(m):
        p = m.group(1) or ""
        return f'<{p}fill><{p}patternFill patternType="none"/></{p}fill>'
    xml = re.sub(r"<((?:\w+:)?)fill\s*/>", sub, xml)
    xml = re.sub(r"<((?:\w+:)?)fill>\s*</\1fill>", sub, xml)
    return xml


def _zerar_fills(xml: str) -> str:
    """Nível 2 (último recurso): descarta as cores de fundo; os dados ficam intactos."""
    xml = re.sub(r"<((?:\w+:)?)fills\b[^>]*>.*?</\1fills>",
                 lambda m: f'<{m.group(1)}fills count="2">{_FILL_PADRAO}'
                           f'<{m.group(1)}fill><{m.group(1)}patternFill patternType="gray125"/>'
                           f'</{m.group(1)}fill></{m.group(1)}fills>', xml, flags=re.S)
    return re.sub(r'\bfillId="\d+"', 'fillId="0"', xml)


def carregar(caminho):
    """openpyxl.load_workbook com conserto automático de formatação defeituosa."""
    try:
        return openpyxl.load_workbook(caminho)
    except Exception as erro_original:  # noqa: BLE001
        for conserto in (_consertar_fills_vazios, _zerar_fills):
            try:
                return openpyxl.load_workbook(_reescrever_styles(caminho, conserto))
            except Exception:  # noqa: BLE001
                continue
        raise erro_original
