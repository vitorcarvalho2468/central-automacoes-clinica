"""
Leitura do PDF da guia SP/SADT da Porto Saúde (o mesmo que sai no botão "Imprimir" do portal).

O PDF tem texto (não é imagem), então os campos são achados pela posição: o valor fica logo
abaixo do rótulo numerado do campo (ex.: "5-Senha" -> "1126624141").

Uso:
    import guia_porto_pdf
    dados = guia_porto_pdf.ler(bytes_ou_caminho)
    python guia_porto_pdf.py guia.pdf        # mostra os dados na tela
"""
from __future__ import annotations

import io
import re
import sys
from pathlib import Path

import pdfplumber

# rótulo -> chave. O rótulo é procurado pelo número do campo no começo do texto.
CAMPOS = {
    "1": "registro_ans",
    "4": "data_autorizacao",
    "5": "senha",
    "6": "validade_senha",
    "7": "guia_operadora",
    "8": "carteira",
    "9": "validade_carteira",
    "10": "beneficiario",
    "89": "nome_social",
    "13": "codigo_solicitante",
    "14": "contratado_solicitante",
    "15": "profissional_solicitante",
    "16": "conselho",
    "17": "numero_conselho",
    "18": "uf",
    "21": "carater",
    "22": "data_solicitacao",
    "29": "codigo_executante",
    "30": "contratado_executante",
    "31": "cnes",
    "32": "tipo_atendimento",
    "91": "regime_atendimento",
}

# "5-Senha", "8 \u00ad Número" ou "8   Número" (o tracinho invisível às vezes vira espaço)
_ROTULO = re.compile(r"^(\d{1,3})(?:\s*[-\u2013\u00ad]\s*|\s+)(?=[^\W\d_])")
_HIFENS = str.maketrans({"\u00ad": "-", "\u2013": "-"})


def _palavras(pagina):
    return pagina.extract_words(keep_blank_chars=True, x_tolerance=2, y_tolerance=2)


def _abrir(origem):
    if isinstance(origem, (bytes, bytearray)):
        return pdfplumber.open(io.BytesIO(origem))
    return pdfplumber.open(str(origem))


def ler(origem) -> dict:
    """origem: bytes do PDF ou caminho. Devolve dict com os campos e a lista 'procedimentos'."""
    with _abrir(origem) as pdf:
        pagina = pdf.pages[0]
        palavras = _palavras(pagina)
        largura = pagina.width

    for w in palavras:
        w["texto"] = " ".join(w["text"].translate(_HIFENS).split())

    rotulos = []
    for w in palavras:
        m = _ROTULO.match(w["texto"])
        if m and w["texto"].lower() not in ("1 - __/__/____",):
            rotulos.append((m.group(1), w))

    def mesma_linha(a, b):
        return abs(a["top"] - b["top"]) < 3

    dados = {}
    for num, rot in rotulos:
        chave = CAMPOS.get(num)
        if not chave or chave in dados:
            continue
        # limite direito: o próximo rótulo da mesma linha
        direita = min([w["x0"] for _, w in rotulos if mesma_linha(w, rot) and w["x0"] > rot["x0"] + 2]
                      + [largura])
        valores = [w for w in palavras
                   if rot["top"] + 3 < w["top"] < rot["top"] + 14
                   and rot["x0"] - 3 <= w["x0"] < direita - 2
                   and not _ROTULO.match(w["texto"]) and w["texto"]]
        dados[chave] = " ".join(w["texto"] for w in sorted(valores, key=lambda w: w["x0"])) or ""

    # nº da guia no prestador: número grande no canto superior direito
    topo = [w for w in palavras if w["top"] < 30 and re.fullmatch(r"\d{6,}", w["texto"])]
    dados["guia_prestador"] = topo[0]["texto"] if topo else (dados.get("guia_operadora") or "")

    dados["procedimentos"] = _procedimentos(palavras)
    return dados


def _procedimentos(palavras):
    """Tabela 'Procedimentos ou Itens Assistenciais Solicitados' (campos 24 a 28)."""
    def acha(prefixo):
        return next((w for w in palavras if w["texto"].startswith(prefixo)), None)

    c24, c25, c26 = acha("24-Tabela"), acha("25 - Código do Procedimento"), acha("26 - Descrição")
    c27, c28 = acha("27-Qtde"), acha("28-Qtde")
    fim = acha("Dados do Contrato Executante")
    if not all((c24, c25, c26, c27, c28)):
        return []
    limite = fim["top"] if fim else c24["top"] + 80

    linhas = {}
    for w in palavras:
        if c24["top"] + 4 < w["top"] < limite - 2 and w["texto"]:
            linhas.setdefault(round(w["top"]), []).append(w)

    procs = []
    for _, ws in sorted(linhas.items()):
        col = {"seq": [], "tabela": [], "codigo": [], "descricao": [], "qt_solicitada": [], "qt_autorizada": []}
        for w in ws:
            x = w["x0"]
            if x < c24["x0"] - 4:
                col["seq"].append(w)
            elif x < c25["x0"] - 4:
                col["tabela"].append(w)
            elif x < c26["x0"] - 4:
                col["codigo"].append(w)
            elif x < c27["x0"] - 4:
                col["descricao"].append(w)
            elif x < c28["x0"] - 4:
                col["qt_solicitada"].append(w)
            else:
                col["qt_autorizada"].append(w)
        p = {k: " ".join(w["texto"] for w in sorted(v, key=lambda w: w["x0"])) for k, v in col.items()}
        # código só com números ou com pontos/traços (ex.: 25130358, 25.12.0620)
        if (not re.fullmatch(r"[\d.\-/ ]+", p["codigo"]) or len(re.sub(r"\D", "", p["codigo"])) < 5
                or (p["seq"] and not p["seq"].isdigit())):
            continue
        for k in ("qt_solicitada", "qt_autorizada"):
            m = re.search(r"\d+", p[k])
            p[k] = int(m.group()) if m else 0
        procs.append(p)
    return procs


if __name__ == "__main__":
    import json
    for arq in sys.argv[1:]:
        print(json.dumps(ler(Path(arq)), ensure_ascii=False, indent=2))
