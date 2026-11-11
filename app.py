"""Central de Automações – janela principal com um botão para cada automação."""
import json
import os
import queue
import re
import sys
import threading
from datetime import datetime
from pathlib import Path
from tkinter import filedialog, messagebox

import customtkinter as ctk

import automacao_anterior
import navegador  # noqa: F401  (pasta do navegador dos robôs dentro do .exe)
import feriado_engine as engine

ctk.set_appearance_mode("light")
ctk.set_default_color_theme("blue")

CONFIG_FERIADO = Path.home() / "CentralAutomacoes" / "config_feriado.json"

NAVY = "#002060"
NAVY_HOVER = "#0A3A8C"
YELLOW = "#FFD84D"
BG = "#F3F5FA"


BASE_DADOS = Path.home() / "CentralAutomacoes"


def abrir_pasta(path: Path):
    try:
        if sys.platform.startswith("win"):
            os.startfile(path)  # type: ignore[attr-defined]
        elif sys.platform == "darwin":
            os.system(f'open "{path}"')
        else:
            os.system(f'xdg-open "{path}"')
    except Exception:
        pass


# ---------------------------------------------------------------- limpar prints de erro
# Os robôs salvam HTML/texto/print (.png) de cada guia em subpastas "paginas"
# (ex.: CentralAutomacoes\SulAmerica\paginas). Isso ajuda a conferir erros, mas pesa.
# Aqui a gente só mexe nessas pastas "paginas": logins, configurações e planilhas ficam intactos.

def _pastas_de_prints():
    if not BASE_DADOS.exists():
        return []
    return [p for p in BASE_DADOS.glob("*/paginas") if p.is_dir()]


def _tam_legivel(n):
    for u in ("B", "KB", "MB", "GB"):
        if n < 1024 or u == "GB":
            return f"{n:.0f} {u}" if u == "B" else f"{n:.1f} {u}"
        n /= 1024


def medir_prints():
    """Devolve (qtde_arquivos, bytes_totais) das pastas 'paginas'."""
    arquivos, total = 0, 0
    for pasta in _pastas_de_prints():
        for f in pasta.rglob("*"):
            if f.is_file():
                try:
                    total += f.stat().st_size
                    arquivos += 1
                except OSError:
                    pass
    return arquivos, total


def limpar_prints():
    """Apaga o conteúdo das pastas 'paginas'. Devolve (removidos, bytes_liberados)."""
    removidos, liberado = 0, 0
    for pasta in _pastas_de_prints():
        for f in list(pasta.rglob("*")):
            if f.is_file():
                try:
                    liberado += f.stat().st_size
                    f.unlink()
                    removidos += 1
                except OSError:
                    pass
        for d in sorted((p for p in pasta.rglob("*") if p.is_dir()), reverse=True):
            try:
                d.rmdir()
            except OSError:
                pass
    return removidos, liberado


# ---------------------------------------------------------------- feriado
class FeriadoWindow(ctk.CTkToplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.title("Automação de Feriado")
        self.geometry("720x780")
        self.configure(fg_color=BG)
        self.minsize(680, 740)
        self.arquivo: Path | None = None
        self.ultima_saida: Path | None = None
        self.extras: list = []
        self.fila: queue.Queue = queue.Queue()

        ctk.CTkLabel(self, text="Redistribuição de Feriado", text_color=NAVY,
                     font=ctk.CTkFont(size=24, weight="bold")).pack(pady=(22, 2))
        ctk.CTkLabel(self, text="Envie a agenda da semana, informe o feriado e receba tudo pronto.",
                     text_color="#51607A").pack(pady=(0, 16))

        card = ctk.CTkFrame(self, fg_color="white", corner_radius=16)
        card.pack(fill="x", padx=28)
        card.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(card, text="1. Agenda da semana (.xlsx)", text_color=NAVY,
                     font=ctk.CTkFont(weight="bold")).grid(row=0, column=0, columnspan=2,
                                                          sticky="w", padx=20, pady=(18, 4))
        self.lbl_arquivo = ctk.CTkLabel(card, text="Nenhum arquivo selecionado",
                                        text_color="#7A869A", anchor="w")
        self.lbl_arquivo.grid(row=1, column=0, columnspan=2, sticky="ew", padx=20)
        ctk.CTkButton(card, text="Escolher arquivo", fg_color=NAVY, hover_color=NAVY_HOVER,
                      command=self.escolher).grid(row=2, column=0, sticky="w", padx=20, pady=(6, 14))

        ctk.CTkLabel(card, text="2. Feriado (opcional)", text_color=NAVY,
                     font=ctk.CTkFont(weight="bold")).grid(row=3, column=0, columnspan=2,
                                                          sticky="w", padx=20, pady=(4, 4))
        self.ent_data = ctk.CTkEntry(card, placeholder_text="Vazio = detectar sozinho", width=240)
        self.ent_data.grid(row=4, column=0, sticky="w", padx=20, pady=(0, 14))
        ctk.CTkLabel(card, text="Deixe vazio: o programa acha o feriado pela agenda.\nOu digite dd/mm/aaaa (vários: separe por vírgula).",
                     text_color="#7A869A", justify="left").grid(row=4, column=1, sticky="w")

        ctk.CTkLabel(card, text="3. Opções", text_color=NAVY,
                     font=ctk.CTkFont(weight="bold")).grid(row=5, column=0, columnspan=2,
                                                          sticky="w", padx=20, pady=(4, 4))
        linha = ctk.CTkFrame(card, fg_color="transparent")
        linha.grid(row=6, column=0, columnspan=2, sticky="w", padx=20, pady=(0, 10))
        ctk.CTkLabel(linha, text="Escrever como:").pack(side="left")
        self.cb_conector = ctk.CTkOptionMenu(linha, values=["de terça", "na terça"], width=110,
                                             fg_color=NAVY, button_color=NAVY_HOVER)
        self.cb_conector.pack(side="left", padx=(8, 22))
        ctk.CTkLabel(linha, text="Máx. por dia:").pack(side="left")
        self.cb_max = ctk.CTkOptionMenu(linha, values=["4", "5", "6"], width=70,
                                        fg_color=NAVY, button_color=NAVY_HOVER)
        self.cb_max.set("5")
        self.cb_max.pack(side="left", padx=8)
        self.sw_modal = ctk.CTkCheckBox(card, text="Mostrar modalidade (ABA / Convencional / PediaSuit)",
                                        fg_color=NAVY, hover_color=NAVY_HOVER)
        self.sw_modal.select()
        self.sw_modal.grid(row=7, column=0, columnspan=2, sticky="w", padx=20, pady=(0, 10))
        ex = ctk.CTkFrame(card, fg_color="transparent")
        ex.grid(row=8, column=0, columnspan=2, sticky="w", padx=20, pady=(0, 18))
        ctk.CTkLabel(ex, text="Feriados fixos da clínica (todo ano):").pack(side="left")
        self.ent_extras = ctk.CTkEntry(ex, width=170, placeholder_text="ex.: 08/12, 15/12")
        self.ent_extras.pack(side="left", padx=8)
        try:
            self.ent_extras.insert(0, json.loads(CONFIG_FERIADO.read_text(encoding="utf-8")).get("extras", ""))
        except (OSError, ValueError):
            pass

        self.btn_gerar = ctk.CTkButton(self, text="Gerar planilha", height=46, corner_radius=12,
                                       fg_color=NAVY, hover_color=NAVY_HOVER,
                                       font=ctk.CTkFont(size=16, weight="bold"),
                                       command=self.gerar)
        self.btn_gerar.pack(fill="x", padx=28, pady=16)

        self.btn_pasta = ctk.CTkButton(self, text="Abrir pasta do resultado", fg_color="transparent",
                                       text_color=NAVY, hover_color="#E3E8F4", border_width=1,
                                       border_color=NAVY, state="disabled",
                                       command=lambda: abrir_pasta(self.ultima_saida.parent))
        self.btn_pasta.pack(side="bottom", pady=(0, 16))
        self.log = ctk.CTkTextbox(self, height=170, fg_color="white", corner_radius=12,
                                  text_color="#243B53")
        self.log.pack(fill="both", expand=True, padx=28, pady=(0, 10))
        self.log.configure(state="disabled")

        self.after(150, self.focus)

    # -- helpers
    def escrever(self, msg: str):
        self.log.configure(state="normal")
        self.log.insert("end", msg + "\n")
        self.log.see("end")
        self.log.configure(state="disabled")

    def escolher(self):
        f = filedialog.askopenfilename(title="Agenda da semana",
                                       filetypes=[("Planilha Excel", "*.xlsx *.xlsm")])
        if f:
            self.arquivo = Path(f)
            self.lbl_arquivo.configure(text=self.arquivo.name, text_color=NAVY)

    def _ler_data(self):
        """Vazio -> None (detecção automática). Senão, lista de datas."""
        txt = self.ent_data.get().lower().replace("feriado:", "").strip()
        if not txt:
            return None
        datas = [engine.parse_date(t) for t in re.findall(r"\d{1,2}/\d{1,2}/\d{2,4}", txt)]
        if not datas or any(d is None for d in datas):
            raise ValueError("Data inválida. Use dd/mm/aaaa, ex.: 02/11/2026 "
                             "(ou deixe vazio para detectar sozinho).")
        return datas

    def _ler_extras(self):
        txt = self.ent_extras.get()
        try:
            CONFIG_FERIADO.parent.mkdir(parents=True, exist_ok=True)
            CONFIG_FERIADO.write_text(json.dumps({"extras": txt}, ensure_ascii=False), encoding="utf-8")
        except OSError:
            pass
        return [(int(d), int(m)) for d, m in re.findall(r"(\d{1,2})/(\d{1,2})", txt)]

    def gerar(self):
        try:
            if not self.arquivo:
                raise ValueError("Escolha o arquivo da agenda primeiro.")
            feriado = self._ler_data()
            self.extras = self._ler_extras()
        except ValueError as e:
            messagebox.showwarning("Falta algo", str(e), parent=self)
            return
        opt = engine.Options(
            conector="na" if self.cb_conector.get().startswith("na") else "de",
            teto_maximo=int(self.cb_max.get()),
            teto_confortavel=min(4, int(self.cb_max.get())),
            modalidade=bool(self.sw_modal.get()),
        )
        self.btn_gerar.configure(state="disabled", text="Gerando...")
        self.escrever(f"▶ {self.arquivo.name} | feriado: "
                      + ("detectando pela agenda..." if feriado is None
                         else ", ".join(f"{d:%d/%m/%Y}" for d in feriado)))
        threading.Thread(target=self._trabalho, args=(feriado, opt), daemon=True).start()
        self.after(100, self._poll)

    def _trabalho(self, feriado, opt):
        # roda fora da janela; só põe o resultado na fila (nada de Tk aqui)
        try:
            self.fila.put(("ok", engine.run(self.arquivo, feriado, opt, extras=self.extras)))
        except PermissionError:
            self.fila.put(("erro", "Não consegui salvar. Feche a planilha no Excel e tente de novo."))
        except Exception as e:  # noqa: BLE001
            self.fila.put(("erro", str(e)))

    def _poll(self):
        try:
            tipo, dado = self.fila.get_nowait()
        except queue.Empty:
            self.after(100, self._poll)
            return
        (self._ok if tipo == "ok" else self._erro)(dado)

    def _ok(self, s):
        self.btn_gerar.configure(state="normal", text="Gerar planilha")
        self.ultima_saida = s.arquivo_saida
        for d, nome in s.feriados:
            self.escrever(f"📅 Feriado: {d:%d/%m/%Y} ({engine.dia_nome(d)}) – {nome}")
        self.escrever(f"✔ Pronto: {s.arquivo_saida.name}")
        self.escrever(f"   {s.pacientes} pacientes | {s.so_feriado} só passam no dia do feriado")
        if s.avisos:
            self.escrever(f"⚠ {s.com_aviso} paciente(s) com dia acima do limite (marcados em rosa):")
            for a in s.avisos:
                self.escrever(f"   • {a}")
        for n in s.notas:
            self.escrever(f"ℹ {n}")
        self.btn_pasta.configure(state="normal")

    def _erro(self, msg):
        self.btn_gerar.configure(state="normal", text="Gerar planilha")
        self.escrever(f"✖ {msg}")
        messagebox.showerror("Erro", msg, parent=self)


# ---------------------------------------------------------------- launcher
def abrir_feriado(parent):
    FeriadoWindow(parent)


def abrir_tr(parent):
    automacao_anterior.abrir(parent)


def _abrir_robo(parent, modulo, titulo):
    # importa só na hora: se faltar alguma biblioteca do robô, o resto da Central continua funcionando
    try:
        mod = __import__(modulo)
    except ImportError as e:
        messagebox.showerror(
            titulo,
            f"Falta instalar uma biblioteca do robô ({e.name or e}).\n\n"
            "Dê dois cliques em \"Instalar dependencias.bat\" e abra a Central de novo.",
            parent=parent)
        return
    mod.abrir(parent)


def abrir_sulamerica(parent):
    _abrir_robo(parent, "interface_robo", "Robô SulAmérica")


def abrir_porto(parent):
    _abrir_robo(parent, "interface_porto", "Robô Porto")


def abrir_unimed(parent):
    _abrir_robo(parent, "interface_unimed", "Robô Unimed")


def abrir_uniao(parent):
    _abrir_robo(parent, "interface_uniao", "União de Conferências")


AUTOMACOES = [
    ("Extrator de TR", "Deixa só o código TR na coluna Observação das planilhas", abrir_tr),
    ("Redistribuição de feriado", "Distribui a agenda do feriado pelos outros dias", abrir_feriado),
    ("Robô SulAmérica", "Confere guias no portal, gera agendas e confere o faturamento", abrir_sulamerica),
    ("Robô Porto", "Confere as guias das abas PORTO e PORTO CONV no portal", abrir_porto),
    ("Robô Unimed", "Confere a leitura e finaliza as guias em aberto no SGU-Card", abrir_unimed),
    ("União de conferências", "Junta as conferências da semana (por convênio/dia) numa geral", abrir_uniao),
]
COLUNAS = 2 if len(AUTOMACOES) in (2, 4) else min(3, len(AUTOMACOES))
LINHAS = -(-len(AUTOMACOES) // COLUNAS)


class Launcher(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Central de Automações")
        self.geometry(f"{140 + 250 * COLUNAS}x{230 + 250 * LINHAS}")
        self.configure(fg_color=BG)
        self.resizable(False, False)

        ctk.CTkLabel(self, text="Central de Automações", text_color=NAVY,
                     font=ctk.CTkFont(size=28, weight="bold")).pack(pady=(34, 4))
        ctk.CTkLabel(self, text="Escolha o que você quer fazer hoje", text_color="#51607A",
                     font=ctk.CTkFont(size=14)).pack(pady=(0, 26))

        box = ctk.CTkFrame(self, fg_color="transparent")
        box.pack(fill="both", expand=True, padx=34)
        box.grid_columnconfigure(tuple(range(COLUNAS)), weight=1, uniform="a")
        box.grid_rowconfigure(tuple(range(LINHAS)), weight=1, uniform="l")
        for i, (titulo, sub, fn) in enumerate(AUTOMACOES):
            card = ctk.CTkFrame(box, fg_color="white", corner_radius=18)
            card.grid(row=i // COLUNAS, column=i % COLUNAS, sticky="nsew", padx=10,
                      pady=(0, 34 if i // COLUNAS == LINHAS - 1 else 20))
            ctk.CTkLabel(card, text=titulo, text_color=NAVY, wraplength=220,
                         font=ctk.CTkFont(size=18, weight="bold")).pack(pady=(26, 6), padx=14)
            ctk.CTkLabel(card, text=sub, text_color="#7A869A", wraplength=210).pack(padx=14)
            ctk.CTkButton(card, text="Abrir", height=42, corner_radius=10,
                          fg_color=NAVY, hover_color=NAVY_HOVER,
                          font=ctk.CTkFont(size=15, weight="bold"),
                          command=lambda f=fn: f(self)).pack(side="bottom", pady=20, padx=26, fill="x")

        rodape = ctk.CTkFrame(self, fg_color="transparent")
        rodape.pack(side="bottom", pady=(0, 18))
        ctk.CTkButton(rodape, text="🧹  Limpar prints de erro", height=38, corner_radius=10,
                      fg_color="transparent", text_color=NAVY, hover_color="#E3E8F4",
                      border_width=1, border_color=NAVY, font=ctk.CTkFont(size=14, weight="bold"),
                      command=self._limpar_prints).pack(side="left", padx=6)
        ctk.CTkButton(rodape, text="Abrir pasta", height=38, corner_radius=10,
                      fg_color="transparent", text_color=NAVY, hover_color="#E3E8F4",
                      border_width=1, border_color=NAVY, font=ctk.CTkFont(size=14),
                      command=self._abrir_pasta_prints).pack(side="left", padx=6)

    def _abrir_pasta_prints(self):
        pastas = _pastas_de_prints()
        abrir_pasta(pastas[0] if pastas else BASE_DADOS)

    def _limpar_prints(self):
        arquivos, total = medir_prints()
        if arquivos == 0:
            messagebox.showinfo("Limpar prints de erro",
                                "Não há prints/páginas salvas para limpar.", parent=self)
            return
        if not messagebox.askyesno(
                "Limpar prints de erro",
                f"Apagar {arquivos} arquivo(s) salvos ({_tam_legivel(total)})?\n\n"
                "Isso remove só os prints e HTML de conferência dos robôs (pastas \"paginas\").\n"
                "NÃO mexe nos logins, nas configurações nem nas planilhas de resultado.\n\n"
                "Essa ação não tem como desfazer.", parent=self):
            return
        removidos, liberado = limpar_prints()
        messagebox.showinfo("Pronto",
                            f"{removidos} arquivo(s) apagado(s). {_tam_legivel(liberado)} liberados.",
                            parent=self)


if __name__ == "__main__":
    Launcher().mainloop()
