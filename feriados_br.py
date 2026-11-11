"""
Calendário de feriados (offline) para o programa descobrir sozinho qual dia da agenda é feriado.

Inclui só feriados que a clínica certamente não abre: nacionais, Sexta-feira Santa e
9 de julho (estadual de SP). Pontos facultativos (Carnaval, Corpus Christi) e feriados
municipais NÃO entram automaticamente: digite a data no campo "Feriado" ou cadastre
em "Feriados fixos da clínica" (ex.: 08/12).
"""
from __future__ import annotations

from datetime import date, timedelta


def pascoa(ano: int) -> date:
    """Domingo de Páscoa (algoritmo gregoriano)."""
    a, b, c = ano % 19, ano // 100, ano % 100
    d, e = b // 4, b % 4
    f = (b + 8) // 25
    g = (b - f + 1) // 3
    h = (19 * a + b - d - g + 15) % 30
    i, k = c // 4, c % 4
    l = (32 + 2 * e + 2 * i - h - k) % 7
    m = (a + 11 * h + 22 * l) // 451
    mes = (h + l - 7 * m + 114) // 31
    dia = (h + l - 7 * m + 114) % 31 + 1
    return date(ano, mes, dia)


_FIXOS = {
    (1, 1): "Confraternização Universal",
    (4, 21): "Tiradentes",
    (5, 1): "Dia do Trabalho",
    (7, 9): "Revolução Constitucionalista (SP)",
    (9, 7): "Independência do Brasil",
    (10, 12): "Nossa Senhora Aparecida",
    (11, 2): "Finados",
    (11, 15): "Proclamação da República",
    (11, 20): "Consciência Negra",
    (12, 25): "Natal",
}


def feriados_do_ano(ano: int) -> dict[date, str]:
    out = {date(ano, m, d): nome for (m, d), nome in _FIXOS.items()}
    out[pascoa(ano) - timedelta(days=2)] = "Sexta-feira Santa"
    return out


def detectar(datas, extras: list[tuple[int, int]] | None = None) -> dict[date, str]:
    """Das datas da agenda, devolve as que são feriado ({data: nome}).
    `extras` = feriados fixos da clínica, como (dia, mes)."""
    datas = sorted(set(datas))
    cache: dict[int, dict[date, str]] = {}
    achados: dict[date, str] = {}
    for d in datas:
        if d.weekday() >= 5:  # sábado/domingo não entram no cálculo
            continue
        tabela = cache.setdefault(d.year, feriados_do_ano(d.year))
        if d in tabela:
            achados[d] = tabela[d]
        elif extras and (d.day, d.month) in set(extras):
            achados[d] = "Feriado cadastrado"
    return achados
