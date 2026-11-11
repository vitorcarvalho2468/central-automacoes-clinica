"""
Extrator de TR – janela da automação (lógica do monitor_pasta_tr.py + interface).

Regra (igual à original): na coluna "Observação" da planilha, se houver um código TR123
no texto, a célula passa a ter só o código (em maiúsculas). Sem TR, o texto é mantido.

Fluxo: [Adicionar planilhas] -> copia para a pasta de ENTRADA -> processa -> grava em SAÍDA
e move o original para PROCESSADOS. O monitoramento da pasta também pode ficar ligado.
"""
from __future__ import annotations

import json
import os
import queue
import re
import shutil
import threading
import time
import unicodedata
from pathlib import Path
from tkinter import filedialog, messagebox

import customtkinter as ctk
import openpyxl

import xlsx_seguro

NAVY, NAVY_HOVER, BG = "#002060", "#0A3A8C", "#F3F5FA"

BASE_PADRAO = Path.home() / "CentralAutomacoes" / "TR"
CONFIG_PATH = Path.home() / "CentralAutomacoes" / "config_tr.json"
TR_PATTERN = re.compile(r"(?i)\bTR\d+\b")


# ============================================================ lógica (sem Tk)
def _norm(s) -> str:
    s = unicodedata.normalize("NFD", str(s or ""))
    return "".join(c for c in s if unicodedata.category(c) != "Mn").lower().strip()


def encontrar_coluna_observacao(ws) -> int:
    for cell in ws[1]:
        if cell.value and _norm(cell.value) == "observacao":
            return cell.column
    raise ValueError("Coluna 'Observação' não encontrada na primeira linha da planilha.")


def processar_planilha(caminho_entrada, caminho_saida):
    wb = xlsx_seguro.carregar(caminho_entrada)
    ws = wb.active
    col_obs = encontrar_coluna_observacao(ws)

    total = com_tr = sem_tr = 0
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        cell = row[col_obs - 1]
        match = TR_PATTERN.search(str(cell.value)) if cell.value is not None else None
        total += 1
        if match:
            cell.value = match.group(0).upper()
            com_tr += 1
        else:  # mantém o que já estava escrito
            sem_tr += 1
    wb.save(caminho_saida)
    return total, com_tr, sem_tr


def arquivo_estavel(caminho, espera=1.0) -> bool:
    """O arquivo parou de ser copiado/salvo? (tamanho não muda)"""
    try:
        t1 = os.path.getsize(caminho)
        time.sleep(espera)
        return t1 == os.path.getsize(caminho)
    except OSError:
        return False


def nome_livre(pasta: Path, nome: str) -> Path:
    """Evita sobrescrever: 'a.xlsx' -> 'a (2).xlsx' se já existir."""
    destino = pasta / nome
    n = 2
    while destino.exists():
        destino = pasta / f"{Path(nome).stem} ({n}){Path(nome).suffix}"
        n += 1
    return destino


def verificar_pasta(entrada: Path, saida: Path, processados: Path, log,
                    checar_estavel: bool = True, falhos: set | None = None) -> int:
    """Processa todas as .xlsx da pasta de entrada. Devolve quantas foram processadas."""
    for p in (entrada, saida, processados):
        p.mkdir(parents=True, exist_ok=True)
    feitos = 0
    for nome in sorted(os.listdir(entrada)):
        if not nome.lower().endswith(".xlsx") or nome.startswith("~$"):
            continue
        origem = entrada / nome
        try:
            chave = (nome, origem.stat().st_mtime)
        except OSError:
            continue
        if falhos is not None and chave in falhos:  # já deu erro; não repete a mensagem
            continue
        if checar_estavel and not arquivo_estavel(origem):
            log(f"[AGUARDANDO] '{nome}' ainda está sendo copiado/salvo.")
            continue
        try:
            destino = nome_livre(saida, nome)
            total, com_tr, sem_tr = processar_planilha(origem, destino)
            shutil.move(str(origem), str(nome_livre(processados, nome)))
            log(f"[OK] '{nome}' -> {total} linhas ({com_tr} com TR, {sem_tr} sem TR). "
                f"Salvo como '{destino.name}'.")
            feitos += 1
        except Exception as e:  # noqa: BLE001
            log(f"[ERRO] Falha ao processar '{nome}': {e} (o arquivo continua na pasta de entrada)")
            if falhos is not None:
                falhos.add(chave)
    return feitos


def carregar_config() -> dict:
    cfg = {"entrada": str(BASE_PADRAO / "entrada"), "saida": str(BASE_PADRAO / "saida"),
           "processados": str(BASE_PADRAO / "processados"), "intervalo": 10}
    try:
        cfg.update(json.loads(CONFIG_PATH.read_text(encoding="utf-8")))
    except (OSError, ValueError):
        pass
    return cfg


def salvar_config(cfg: dict):
    try:
        CONFIG_PATH.parent.mkdir(parents=True, exist_ok=True)
        CONFIG_PATH.write_text(json.dumps(cfg, ensure_ascii=False, indent=2), encoding="utf-8")
    except OSError:
        pass


def abrir_pasta(path: Path):
    import sys
    try:
        path.mkdir(parents=True, exist_ok=True)
        if sys.platform.startswith("win"):
            os.startfile(path)  # type: ignore[attr-defined]
        elif sys.platform == "darwin":
            os.system(f'open "{path}"')
        else:
            os.system(f'xdg-open "{path}"')
    except Exception:  # noqa: BLE001
        pass


# ================================================================== janela
class TRWindow(ctk.CTkToplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.title("Extrator de TR")
        self.geometry("740x760")
        self.minsize(700, 700)
        self.configure(fg_color=BG)
        self.cfg = carregar_config()
        self.fila: queue.Queue = queue.Queue()
        self.lock = threading.Lock()
        self.parar = threading.Event()
        self.monitor: threading.Thread | None = None
        self.falhos: set = set()

        ctk.CTkLabel(self, text="Extrator de TR", text_color=NAVY,
                     font=ctk.CTkFont(size=24, weight="bold")).pack(pady=(22, 2))
        ctk.CTkLabel(self, text="Deixa só o código TR na coluna Observação das planilhas.",
                     text_color="#51607A").pack(pady=(0, 14))

        # --- card 1: adicionar
        c1 = ctk.CTkFrame(self, fg_color="white", corner_radius=16)
        c1.pack(fill="x", padx=28)
        ctk.CTkLabel(c1, text="1. Planilhas", text_color=NAVY,
                     font=ctk.CTkFont(weight="bold")).pack(anchor="w", padx=20, pady=(16, 2))
        ctk.CTkLabel(c1, text="Elas são copiadas para a pasta de entrada e processadas na hora.",
                     text_color="#7A869A").pack(anchor="w", padx=20)
        ctk.CTkButton(c1, text="Adicionar planilhas (.xlsx)", height=44, corner_radius=12,
                      fg_color=NAVY, hover_color=NAVY_HOVER,
                      font=ctk.CTkFont(size=15, weight="bold"),
                      command=self.escolher).pack(fill="x", padx=20, pady=(10, 18))

        # --- card 2: pastas
        c2 = ctk.CTkFrame(self, fg_color="white", corner_radius=16)
        c2.pack(fill="x", padx=28, pady=(12, 0))
        c2.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(c2, text="2. Pastas", text_color=NAVY,
                     font=ctk.CTkFont(weight="bold")).grid(row=0, column=0, columnspan=3,
                                                          sticky="w", padx=20, pady=(16, 6))
        self.lbl_pastas = {}
        for i, (chave, nome) in enumerate([("entrada", "Entrada"), ("saida", "Saída"),
                                           ("processados", "Processados")], start=1):
            ctk.CTkLabel(c2, text=nome, width=90, anchor="w").grid(row=i, column=0, padx=(20, 4), pady=3)
            lbl = ctk.CTkLabel(c2, text=self._curto(self.cfg[chave]), anchor="w", text_color="#51607A")
            lbl.grid(row=i, column=1, sticky="ew", pady=3)
            self.lbl_pastas[chave] = lbl
            ctk.CTkButton(c2, text="Alterar", width=80, height=28, fg_color="transparent",
                          text_color=NAVY, border_width=1, border_color=NAVY, hover_color="#E3E8F4",
                          command=lambda k=chave: self.trocar_pasta(k)).grid(row=i, column=2, padx=(6, 20), pady=3)
        ctk.CTkLabel(c2, text="").grid(row=4, column=0, pady=2)

        # --- card 3: monitoramento
        c3 = ctk.CTkFrame(self, fg_color="white", corner_radius=16)
        c3.pack(fill="x", padx=28, pady=(12, 0))
        linha = ctk.CTkFrame(c3, fg_color="transparent")
        linha.pack(fill="x", padx=20, pady=14)
        self.sw = ctk.CTkSwitch(linha, text="Monitorar a pasta de entrada", progress_color=NAVY,
                                command=self.alternar_monitor)
        self.sw.pack(side="left")
        ctk.CTkLabel(linha, text="a cada (s):").pack(side="left", padx=(24, 6))
        self.ent_int = ctk.CTkEntry(linha, width=56)
        self.ent_int.insert(0, str(self.cfg["intervalo"]))
        self.ent_int.pack(side="left")

        # --- botão inferior + log
        bot = ctk.CTkFrame(self, fg_color="transparent")
        bot.pack(side="bottom", pady=(0, 16))
        for txt, cmd in [("Abrir pasta de saída", lambda: abrir_pasta(Path(self.cfg["saida"]))),
                         ("Abrir pasta de entrada", lambda: abrir_pasta(Path(self.cfg["entrada"])))]:
            ctk.CTkButton(bot, text=txt, fg_color="transparent", text_color=NAVY, hover_color="#E3E8F4",
                          border_width=1, border_color=NAVY, command=cmd).pack(side="left", padx=6)
        self.log = ctk.CTkTextbox(self, fg_color="white", corner_radius=12, text_color="#243B53")
        self.log.pack(fill="both", expand=True, padx=28, pady=12)
        self.log.configure(state="disabled")

        self.protocol("WM_DELETE_WINDOW", self.fechar)
        self.after(150, self._poll)
        self.after(200, self.focus)

    # ---- util
    @staticmethod
    def _curto(p: str, n: int = 52) -> str:
        return p if len(p) <= n else "…" + p[-(n - 1):]

    def escrever(self, msg: str):
        self.log.configure(state="normal")
        self.log.insert("end", f"{time.strftime('%H:%M:%S')}  {msg}\n")
        self.log.see("end")
        self.log.configure(state="disabled")

    def _poll(self):  # único lugar que toca na janela a partir do resultado das threads
        try:
            while True:
                self.escrever(self.fila.get_nowait())
        except queue.Empty:
            pass
        if self.winfo_exists():
            self.after(150, self._poll)

    def _pastas(self):
        return Path(self.cfg["entrada"]), Path(self.cfg["saida"]), Path(self.cfg["processados"])

    # ---- ações
    def trocar_pasta(self, chave):
        p = filedialog.askdirectory(title=f"Pasta de {chave}", initialdir=self.cfg[chave], parent=self)
        if p:
            self.cfg[chave] = p
            self.lbl_pastas[chave].configure(text=self._curto(p))
            salvar_config(self.cfg)

    def escolher(self):
        arqs = filedialog.askopenfilenames(title="Planilhas", parent=self,
                                           filetypes=[("Planilha Excel", "*.xlsx")])
        if arqs:
            self.adicionar([Path(a) for a in arqs])

    def adicionar(self, caminhos: list[Path]):
        entrada, saida, proc = self._pastas()
        entrada.mkdir(parents=True, exist_ok=True)
        copiados = 0
        for c in caminhos:
            if c.suffix.lower() != ".xlsx":
                self.escrever(f"[IGNORADO] '{c.name}' não é .xlsx")
                continue
            try:
                shutil.copy2(c, nome_livre(entrada, c.name))
                copiados += 1
            except OSError as e:
                self.escrever(f"[ERRO] Não consegui copiar '{c.name}': {e}")
        if copiados:
            self.escrever(f"{copiados} planilha(s) enviada(s) para a pasta de entrada.")
            threading.Thread(target=self._processar, daemon=True).start()

    def _processar(self):  # roda em thread; só usa a fila para falar com a janela
        with self.lock:
            e, s, p = self._pastas()
            verificar_pasta(e, s, p, self.fila.put, checar_estavel=False, falhos=self.falhos)

    def alternar_monitor(self):
        if self.sw.get():
            try:
                self.cfg["intervalo"] = max(2, int(self.ent_int.get()))
            except ValueError:
                self.cfg["intervalo"] = 10
            salvar_config(self.cfg)
            self.parar.clear()
            self.monitor = threading.Thread(target=self._loop, args=(self.cfg["intervalo"],), daemon=True)
            self.monitor.start()
            self.escrever(f"Monitorando '{self.cfg['entrada']}' a cada {self.cfg['intervalo']}s.")
        else:
            self.parar.set()
            self.escrever("Monitoramento desligado.")

    def _loop(self, intervalo):
        while not self.parar.is_set():
            with self.lock:
                e, s, p = self._pastas()
                verificar_pasta(e, s, p, self.fila.put, falhos=self.falhos)
            self.parar.wait(intervalo)

    def fechar(self):
        self.parar.set()
        self.destroy()


def abrir(parent):
    TRWindow(parent)
