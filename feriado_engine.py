"""
Motor da automação de feriado.

Entrada : planilha .xlsx com a agenda da semana (o feriado é detectado sozinho,
          ou informado à mão).
Saída   : novo .xlsx com as abas
          - Planilha de feriado (resumo do feriado + coluna Redistribuição)
          - Apoio               (todos os pacientes e suas distribuições)
          - uma aba por dia útil restante (Terça, Quarta...) com quem vem e o que faz
"""
from __future__ import annotations

import re
import unicodedata
from collections import defaultdict
from dataclasses import dataclass, field
from datetime import date, datetime, time
from pathlib import Path

import openpyxl

import xlsx_seguro
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.workbook.properties import CalcProperties

import feriados_br

# ----------------------------------------------------------------- constantes
BASES = THERAPIES = [  # (nome base, sigla da coluna)
    ("Psicologia", "PSICO"),
    ("Fonoaudiologia", "FONO"),
    ("Terapia Ocupacional", "TO"),
    ("Psicopedagogia", "PSICOPE"),
    ("Fisioterapia", "FISIO"),
    ("Musicoterapia", "MUSICO"),
]
ORDER = {name: i for i, (name, _) in enumerate(THERAPIES)}
ALIASES = {  # nomes aceitos para cada coluna da agenda (sem acento / maiúscula)
    "dia": ["data da agenda", "dia", "data"],
    "hora": ["hora", "horario"],
    "pac": ["paciente"],
    "ter": ["grupo de procedimento", "terapia", "procedimento"],
    "emp": ["empresa", "unidade"],
    "conv": ["convenio/plano", "convenio"],
    "cart": ["carteirinha"],
}
DIAS = ["segunda", "terça", "quarta", "quinta", "sexta", "sábado", "domingo"]

NAVY = PatternFill("solid", fgColor="002060")
YELLOW = PatternFill("solid", fgColor="FFFF00")
RED_LIGHT = PatternFill("solid", fgColor="FCE4D6")
H_FONT = Font(name="Calibri", size=12, bold=True, color="FFFFFF")
CENTER = Alignment(horizontal="center", vertical="center")
LEFT = Alignment(horizontal="left", vertical="center")


@dataclass
class Options:
    conector: str = "de"      # "de" -> "de terça"  |  "na" -> "na terça"
    teto_confortavel: int = 4  # tenta manter cada dia com no máximo isso
    teto_maximo: int = 5       # nunca passa disso (exceto se só houver 1 outro dia)
    modalidade: bool = True    # True: "Psicologia ABA" | False: só "Psicologia"


@dataclass
class Appt:
    dia: date
    hora: time | None
    paciente: str
    terapia: str
    und: str
    convenio: str
    carteirinha: str
    chave: str = ""  # identifica o paciente (nome, ou nome+carteirinha se houver homônimos)


@dataclass
class PatientResult:
    paciente: str
    chave: str
    und: str
    convenio: str
    carteirinha: str
    hora_feriado: time | None
    feriado_terapias: list[str]
    dias_feriado: list[date] = field(default_factory=list)  # feriados em que ele tinha terapia
    por_dia: dict[date, list[str]] = field(default_factory=dict)  # redistribuição
    so_feriado: bool = False
    avisos: list[str] = field(default_factory=list)
    texto: str = ""


@dataclass
class Summary:
    arquivo_saida: Path
    pacientes: int
    so_feriado: int
    com_aviso: int
    avisos: list[str]
    notas: list[str] = field(default_factory=list)
    feriados: list[tuple[date, str]] = field(default_factory=list)  # (data, nome)


# ---------------------------------------------------------------- utilidades
def _norm(s: str) -> str:
    s = unicodedata.normalize("NFD", str(s or ""))
    return "".join(c for c in s if unicodedata.category(c) != "Mn").lower().strip()


def parse_date(v) -> date | None:
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    if isinstance(v, str):
        v = v.strip()
        for fmt in ("%d/%m/%Y", "%d/%m/%y", "%Y-%m-%d"):
            try:
                return datetime.strptime(v, fmt).date()
            except ValueError:
                pass
    return None


def parse_time(v) -> time | None:
    if isinstance(v, datetime):
        return v.time().replace(second=0, microsecond=0)
    if isinstance(v, time):
        return v.replace(second=0, microsecond=0)
    if isinstance(v, str):
        m = re.match(r"^\s*(\d{1,2}):(\d{2})", v)
        if m:
            return time(int(m.group(1)), int(m.group(2)))
    return None


def parse_und(empresa: str) -> str:
    empresa = str(empresa or "").strip()
    m = re.match(r"^(\d+)", empresa)
    if m:
        return str(int(m.group(1)))
    words = re.findall(r"[A-Za-zÀ-ÿ]+", empresa)
    return "".join(w[0] for w in words).upper() or "?"


def dia_nome(d: date) -> str:
    return DIAS[d.weekday()]


def com_conector(d: date, conector: str) -> str:
    nome = dia_nome(d)
    if conector == "na":
        return f"{'no' if d.weekday() >= 5 else 'na'} {nome}"
    return f"de {nome}"


def base_of(terapia: str) -> str | None:
    """'Psicologia ABA' -> 'Psicologia' (None se não for uma das 6 bases)."""
    n = _norm(terapia)
    for name, _ in BASES:
        if n.startswith(_norm(name)):
            return name
    return None


def label(terapia: str, modalidade: bool) -> str:
    b = base_of(terapia)
    return terapia if (modalidade or b is None) else b


def _sort_terapias(items: list[str]) -> list[str]:
    return sorted(items, key=lambda t: (ORDER.get(base_of(t), 99), t))


def _contar(items: list[str]) -> str:
    cont: dict[str, int] = defaultdict(int)
    for t in items:
        cont[t] += 1
    return " + ".join(f"{cont[t]} {t}" for t in _sort_terapias(list(cont)))


# ------------------------------------------------------------------- leitura
def _map_header(ws):
    header = {_norm(c.value): i for i, c in enumerate(ws[1]) if c.value}
    cols = {}
    for key, names in ALIASES.items():
        for n in names:
            if n in header:
                cols[key] = header[n]
                break
    return cols


def find_sheet(wb: openpyxl.Workbook):
    """Acha a aba da agenda: 'Agenda', 'Agendamentos' ou qualquer uma com as colunas certas."""
    ordem = [n for n in ("Agenda", "Agendamentos") if n in wb.sheetnames] + \
            [n for n in wb.sheetnames if n not in ("Agenda", "Agendamentos")]
    for n in ordem:
        cols = _map_header(wb[n])
        if all(k in cols for k in ("dia", "hora", "pac", "ter", "emp")):
            return wb[n], cols
    raise ValueError("Não achei a aba da agenda. Ela precisa ter as colunas: "
                     "Dia/Data da Agenda, Hora, Paciente, Terapia/Grupo de Procedimento, "
                     "Unidade/Empresa.")


def read_agenda(wb: openpyxl.Workbook) -> tuple[list[Appt], str, list[str]]:
    ws, cols = find_sheet(wb)
    get = lambda row, k: row[cols[k]] if k in cols else None  # noqa: E731

    out: list[Appt] = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        pac, dia, ter = get(row, "pac"), parse_date(get(row, "dia")), get(row, "ter")
        if not pac or not dia or not ter:
            continue
        cart = get(row, "cart")
        out.append(Appt(
            dia=dia, hora=parse_time(get(row, "hora")), paciente=str(pac).strip(),
            terapia=str(ter).strip(), und=parse_und(get(row, "emp")),
            convenio=str(get(row, "conv") or ""),
            carteirinha=str(cart).strip() if cart not in (None, "") else "",
        ))

    # homônimos: mesmo nome com carteirinhas diferentes = pessoas diferentes
    carts: dict[str, set[str]] = defaultdict(set)
    for a in out:
        if a.carteirinha:
            carts[a.paciente].add(a.carteirinha)
    homonimos = {n for n, c in carts.items() if len(c) > 1}
    for a in out:
        a.chave = f"{a.paciente}|{a.carteirinha}" if a.paciente in homonimos else a.paciente
    notas = ([f"Nome repetido com carteirinhas diferentes (tratados como pessoas distintas): "
              f"{', '.join(sorted(homonimos))}"] if homonimos else [])
    return out, ws.title, notas


# ------------------------------------------------------------ distribuição
def distribute(feriado_terapias: list[str], cargas: dict[date, int], opt: Options) -> tuple[dict[date, list[str]], list[str]]:
    """Devolve (terapia por dia, avisos). `cargas` = terapias já existentes por dia."""
    dias = sorted(d for d, n in cargas.items() if n > 0)  # cargas já exclui os feriados
    avisos: list[str] = []
    res: dict[date, list[str]] = defaultdict(list)

    if len(dias) == 1:  # só mais um dia: pode passar do teto
        d = dias[0]
        res[d] = list(feriado_terapias)
        total = cargas[d] + len(feriado_terapias)
        if total > opt.teto_maximo:
            avisos.append(f"{total} terapias {com_conector(d, 'de')} (único outro dia)")
        return dict(res), avisos

    carga = dict(cargas)
    for t in feriado_terapias:
        cand = [d for d in dias if carga[d] < opt.teto_confortavel]
        if not cand:
            cand = [d for d in dias if carga[d] < opt.teto_maximo]
        if not cand:
            cand = [min(dias, key=lambda d: (carga[d], d))]
            avisos.append(f"estourou {opt.teto_maximo} terapias {com_conector(cand[0], 'de')}")
        d = cand[0]
        res[d].append(t)
        carga[d] += 1
    return dict(res), avisos


def _nomes_dias(ds: list[date]) -> str:
    return " e ".join(dia_nome(d) for d in sorted(ds))


def build_text(r: PatientResult, opt: Options) -> str:
    if r.so_feriado:  # "Só passa de segunda" / "Só passa de segunda e terça"
        ds = sorted(r.dias_feriado)
        return f"Só passa {com_conector(ds[0], opt.conector)}" + "".join(
            f" e {dia_nome(d)}" for d in ds[1:])
    partes = [f"{_contar(r.por_dia[d])} {com_conector(d, opt.conector)}"
              for d in sorted(r.por_dia)]
    return " e ".join(partes)


def process(appts: list[Appt], feriados: list[date], opt: Options) -> list[PatientResult]:
    fer = set(feriados)
    por_pac: dict[str, list[Appt]] = defaultdict(list)
    for a in appts:
        por_pac[a.chave].append(a)

    results: list[PatientResult] = []
    for chave, lst in por_pac.items():
        do_feriado = sorted((a for a in lst if a.dia in fer),
                            key=lambda a: (a.hora or time(23, 59), ORDER.get(base_of(a.terapia), 99)))
        if not do_feriado:
            continue
        cargas: dict[date, int] = defaultdict(int)
        for a in lst:
            if a.dia not in fer:
                cargas[a.dia] += 1
        ter = [label(a.terapia, opt.modalidade) for a in do_feriado]
        r = PatientResult(
            paciente=do_feriado[0].paciente, chave=chave, und=do_feriado[0].und,
            convenio=do_feriado[0].convenio, carteirinha=do_feriado[0].carteirinha,
            hora_feriado=do_feriado[0].hora, feriado_terapias=ter,
            dias_feriado=sorted({a.dia for a in do_feriado}),
        )
        if not cargas:
            r.so_feriado = True
        else:
            r.por_dia, r.avisos = distribute(ter, cargas, opt)
        r.texto = build_text(r, opt)
        results.append(r)
    return results  # mantém a ordem em que os pacientes aparecem na Agenda


# --------------------------------------------------------------------- saída
def _hdr(ws, row, headers, start_col=1):
    for i, h in enumerate(headers):
        c = ws.cell(row, start_col + i, h)
        c.font, c.fill, c.alignment = H_FONT, NAVY, CENTER


def _fit(ws, widths: dict[str, float]):
    for k, w in widths.items():
        ws.column_dimensions[k].width = w


def write_output(wb, fonte: str, results, appts, feriados: list[date], opt: Options):
    for name in list(wb.sheetnames):  # recria tudo, mantém só a aba da agenda
        if name != fonte:
            del wb[name]

    feriados = sorted(feriados)
    nomes_fer = _nomes_dias(feriados)
    datas_fer = " e ".join(f"{d:%d/%m/%Y}" for d in feriados)
    outros = sorted({a.dia for a in appts if a.dia not in set(feriados) and a.dia.weekday() < 5})

    # cargas e unidades por paciente/dia
    cargas = defaultdict(lambda: defaultdict(int))
    unds = defaultdict(lambda: defaultdict(lambda: defaultdict(int)))
    horas = defaultdict(lambda: defaultdict(list))
    for a in appts:
        cargas[a.chave][a.dia] += 1
        unds[a.chave][a.dia][a.und] += 1
        if a.hora:
            horas[a.chave][a.dia].append(a.hora)

    def und_do_dia(r, d):
        u = unds[r.chave][d]
        return max(u, key=u.get) if u else r.und

    # -- Planilha de feriado ----------------------------------------------------
    ws = wb.create_sheet("Planilha de feriado")
    c_ter = 6                              # F: primeira coluna de terapia
    c_outros_t = c_ter + len(BASES)        # coluna OUTROS
    c_total = c_outros_t + 1
    c_dias = c_total + 2                   # deixa uma coluna de respiro
    c_tot2 = c_dias + len(outros)
    c_red = c_tot2 + 2
    L = get_column_letter

    ws.merge_cells(start_row=1, start_column=c_ter, end_row=1, end_column=c_total)
    t = ws.cell(1, c_ter, f"{nomes_fer.upper()} - FERIADO {datas_fer}")
    t.font, t.fill, t.alignment = Font(name="Calibri", size=13, bold=True), YELLOW, CENTER
    _hdr(ws, 2, ["UND", "Hora", "Paciente", "Convenio/Plano", "Carteirinha"]
         + [sg for _, sg in BASES] + ["OUTROS", "TOTAL"])
    for i, d in enumerate(outros):
        c = ws.cell(1, c_dias + i, dia_nome(d).upper())
        c.font, c.fill, c.alignment = Font(name="Calibri", size=13, bold=True), YELLOW, CENTER
    _hdr(ws, 2, [dia_nome(d)[:3].upper() for d in outros], c_dias)
    _hdr(ws, 2, ["TOTAL"], c_tot2)
    _hdr(ws, 2, ["Redistribuição"], c_red)

    for i, r in enumerate(results):
        row = 3 + i
        for j, v in enumerate([r.und, r.hora_feriado, r.paciente, r.convenio, r.carteirinha]):
            ws.cell(row, 1 + j, v)
        for j, (nome, _) in enumerate(BASES):
            ws.cell(row, c_ter + j, sum(1 for x in r.feriado_terapias if base_of(x) == nome))
        ws.cell(row, c_outros_t, sum(1 for x in r.feriado_terapias if base_of(x) is None))
        ws.cell(row, c_total, f"=SUM({L(c_ter)}{row}:{L(c_outros_t)}{row})")
        for j, d in enumerate(outros):
            ws.cell(row, c_dias + j, cargas[r.chave].get(d, 0))
        ws.cell(row, c_tot2, f"=SUM({L(c_dias)}{row}:{L(c_tot2 - 1)}{row})" if outros else 0)
        ws.cell(row, c_red, r.texto)
        for c in range(1, c_red + 1):
            ws.cell(row, c).alignment = CENTER if c != c_red else LEFT
        ws.cell(row, 2).number_format = "hh:mm"
        if r.avisos:
            ws.cell(row, c_red).fill = RED_LIGHT
    ws.auto_filter.ref = f"A2:{L(c_red)}{2 + len(results)}"
    ws.freeze_panes = "D3"
    _fit(ws, {"A": 6, "B": 8, "C": 42, "D": 26, "E": 22, L(c_outros_t): 9,
              L(c_total + 1): 1.5, L(c_tot2 + 1): 1.5, L(c_red): 80})

    # -- Apoio ------------------------------------------------------------------
    ap = wb.create_sheet("Apoio")
    n_end = 5 + len(outros)
    _hdr(ap, 1, ["UND", "Paciente", "Carteirinha", "Redistribuição"]
         + [dia_nome(d).capitalize() for d in outros] + ["Atenção"])
    for i, r in enumerate(results):
        row = 2 + i
        ap.cell(row, 1, r.und)
        ap.cell(row, 2, r.paciente)
        ap.cell(row, 3, r.carteirinha)
        ap.cell(row, 4, r.texto)
        atencao = []
        for j, d in enumerate(outros):
            if d in r.por_dia:
                ap.cell(row, 5 + j, _contar(r.por_dia[d]))
                total = cargas[r.chave].get(d, 0) + len(r.por_dia[d])
                if total > opt.teto_maximo:
                    atencao.append(f"{total} terapias {com_conector(d, 'de')}")
                u = und_do_dia(r, d)
                if u != r.und:
                    atencao.append(f"unidade {u} {com_conector(d, 'de')} (feriado era {r.und})")
        ap.cell(row, n_end, "; ".join(atencao))
        for c in range(1, n_end + 1):
            ap.cell(row, c).alignment = LEFT if c in (2, 4) else CENTER
            if atencao:
                ap.cell(row, c).fill = RED_LIGHT
    ap.auto_filter.ref = f"A1:{L(n_end)}{1 + len(results)}"
    ap.freeze_panes = "C2"
    _fit(ap, {"A": 6, "B": 42, "C": 22, "D": 80, L(n_end): 50})
    for j in range(len(outros)):
        ap.column_dimensions[L(5 + j)].width = 40

    # -- abas por dia -----------------------------------------------------------
    for d in outros:
        tab = wb.create_sheet(dia_nome(d).capitalize())
        _hdr(tab, 1, ["UND", "Paciente", "Carteirinha", "Hora", "Agenda", "CONF"])
        tab["E1"].fill = tab["F1"].fill = YELLOW
        tab["E1"].font = tab["F1"].font = Font(name="Calibri", size=12, bold=True)
        linhas = []
        for r in results:
            if d in r.por_dia:
                hs = horas[r.chave].get(d)
                linhas.append((min(hs) if hs else None, r))
        linhas.sort(key=lambda x: (x[0] or time(23, 59), x[1].paciente))
        for i, (h, r) in enumerate(linhas):
            row = 2 + i
            tab.cell(row, 1, und_do_dia(r, d))
            tab.cell(row, 2, r.paciente)
            tab.cell(row, 3, r.carteirinha)
            tab.cell(row, 4, h).number_format = "hh:mm"
            tab.cell(row, 5, f"{_contar(r.por_dia[d])} (da {_nomes_dias(r.dias_feriado)})")
            for c in range(1, 7):
                tab.cell(row, c).alignment = LEFT if c in (2, 5) else CENTER
        tab.auto_filter.ref = f"A1:F{max(2, 1 + len(linhas))}"
        tab.freeze_panes = "A2"
        _fit(tab, {"A": 6, "B": 42, "C": 22, "D": 8, "E": 70, "F": 8})

    wb.move_sheet("Planilha de feriado", offset=-wb.index(wb["Planilha de feriado"]))
    wb.active = 0


# ---------------------------------------------------------------- API pública
def run(input_path: str | Path, feriado: date | list[date] | None = None,
        opt: Options | None = None, output_path: str | Path | None = None,
        extras: list[tuple[int, int]] | None = None) -> Summary:
    """feriado=None -> detecta sozinho pelas datas da agenda (calendário em feriados_br.py).
    `extras` = feriados fixos da clínica, [(dia, mes), ...]."""
    opt = opt or Options()
    input_path = Path(input_path)
    wb = xlsx_seguro.carregar(input_path)  # mantém a aba da agenda e suas fórmulas
    appts, fonte, notas = read_agenda(wb)
    if not appts:
        raise ValueError("Não encontrei nenhum atendimento na agenda.")
    datas = sorted({a.dia for a in appts})
    periodo = f"{datas[0]:%d/%m/%Y} a {datas[-1]:%d/%m/%Y}"

    if feriado is None:
        achados = feriados_br.detectar(datas, extras)
        if not achados:
            raise ValueError(
                f"Não encontrei nenhum feriado entre {periodo.replace(" a ", " e ")}. Se houver feriado municipal "
                "ou ponto facultativo, digite a data no campo Feriado.")
        feriados = sorted(achados)
        nomes = [(d, achados[d]) for d in feriados]
    else:
        feriados = sorted(set([feriado] if isinstance(feriado, date) else feriado))
        faltam = [d for d in feriados if d not in set(datas)]
        if faltam:
            raise ValueError(
                f"Não há atendimentos em {', '.join(f'{d:%d/%m/%Y}' for d in faltam)} na agenda. "
                f"Datas encontradas: {', '.join(f'{d:%d/%m}' for d in datas)}")
        nomes = [(d, feriados_br.detectar([d], extras).get(d, "informado por você")) for d in feriados]

    semanas = {d.isocalendar()[:2] for d in feriados}
    if len(semanas) > 1:
        raise ValueError("Há feriados em semanas diferentes. Envie uma semana de cada vez.")
    semana = next(iter(semanas))
    appts = [a for a in appts if a.dia.isocalendar()[:2] == semana]  # redistribui só dentro da semana

    results = process(appts, feriados, opt)

    # linhas repetidas nos dias de feriado (mesmo paciente/hora/terapia)
    rep: dict[tuple, int] = defaultdict(int)
    for a in appts:
        if a.dia in set(feriados):
            rep[(a.paciente, a.dia, a.hora, a.terapia)] += 1
    for (pac, _d, h, ter), n in rep.items():
        if n > 1:
            quando = f"{h:%H:%M} " if h else ""
            notas.append(f"Linha repetida no feriado ({n}x): {pac}, {quando}{ter}")

    write_output(wb, fonte, results, appts, feriados, opt)
    out = Path(output_path) if output_path else input_path.with_name(
        f"{input_path.stem}_feriado_{'_'.join(f'{d:%d-%m-%Y}' for d in feriados)}.xlsx")
    if wb.calculation is None:  # alguns exports não trazem a configuração de cálculo
        wb.calculation = CalcProperties()
    wb.calculation.fullCalcOnLoad = True
    wb.save(out)
    avisos = [f"{r.paciente}: {'; '.join(r.avisos)}" for r in results if r.avisos]
    return Summary(out, len(results), sum(r.so_feriado for r in results), len(avisos),
                   avisos, notas, nomes)
