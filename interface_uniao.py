"""
União de Conferências - junta várias planilhas de conferência (por convênio/dia) numa geral.

Abre pela Central de Automações (botão "União de conferências") ou sozinho: python interface_uniao.py
"""
from __future__ import annotations

import datetime as dt
import os
import queue
import sys
import threading
import time
from pathlib import Path
from tkinter import filedialog, messagebox

import customtkinter as ctk

import uniao

ctk.set_appearance_mode("light")
ctk.set_default_color_theme("blue")

NAVY, NAVY_HOVER, BG = "#002060", "#0A3A8C", "#F3F5FA"
CINZA, CINZA_ESCURO, TEXTO = "#7A869A", "#51607A", "#243B53"
CONTORNO_HOVER = "#E3E8F4"
VERMELHO, LARANJA, VERDE = "#C00000", "#B26A00", "#2E7D32"


def abrir_arquivo(caminho):
    try:
        if sys.platform.startswith("win"):
            os.startfile(caminho)  # type: ignore[attr-defined]
        elif sys.platform == "darwin":
            os.system(f'open "{caminho}"')
        else:
            os.system(f'xdg-open "{caminho}"')
    except Exception:
        pass


class _JanelaUniao:
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.title("União de Conferências")
        altura = max(620, min(860, self.winfo_screenheight() - 90))
        self.geometry(f"820x{altura}+60+20")
        self.minsize(720, 560)
        self.configure(fg_color=BG)
        self.fila = queue.Queue()
        self.arquivos = []
        self.resultado = None
        self.thread = None
        self._montar_tela()
        self.protocol("WM_DELETE_WINDOW", self._fechar)
        self.after(150, self._ler_fila)

    def _fonte(self, tamanho=13, negrito=False):
        return ctk.CTkFont(size=tamanho, weight="bold" if negrito else "normal")

    def _botao_contorno(self, pai, texto, comando, **kw):
        return ctk.CTkButton(pai, text=texto, command=comando, fg_color="transparent", text_color=NAVY,
                             hover_color=CONTORNO_HOVER, border_width=1, border_color=NAVY,
                             text_color_disabled="#A9B3C4", **kw)

    def _botao_cheio(self, pai, texto, comando, **kw):
        return ctk.CTkButton(pai, text=texto, command=comando, fg_color=NAVY, hover_color=NAVY_HOVER, **kw)

    def _montar_tela(self):
        corpo = ctk.CTkScrollableFrame(self, fg_color=BG)
        corpo.pack(fill="both", expand=True)
        ctk.CTkLabel(corpo, text="União de Conferências", text_color=NAVY, font=self._fonte(26, True)).pack(pady=(20, 2))
        ctk.CTkLabel(corpo, text="Junta todas as conferências da semana (por convênio/dia) numa única "
                                 "Conferência geral, no padrão de 12 colunas.",
                     text_color=CINZA_ESCURO, font=self._fonte(14), wraplength=720).pack(pady=(0, 16))

        card = ctk.CTkFrame(corpo, fg_color="white", corner_radius=16)
        card.pack(fill="x", padx=28)
        ctk.CTkLabel(card, text="Adicione as planilhas de conferência. Para a Intermédica, o robô usa a aba "
                                "\"RESULT/Resultado\" sozinho; das outras usa a \"Planilha1\". Cabeçalhos diferentes "
                                "(Satus/Status, Convênio/Convenio...) são ajustados automaticamente.",
                     text_color=CINZA_ESCURO, justify="left", wraplength=720).pack(anchor="w", padx=16, pady=(14, 8))

        linha = ctk.CTkFrame(card, fg_color="transparent")
        linha.pack(fill="x", padx=16)
        self._botao_cheio(linha, "Adicionar planilhas", self._adicionar, width=170).pack(side="left")
        self._botao_contorno(linha, "Limpar lista", self._limpar, width=120).pack(side="left", padx=8)
        self.lbl_qtd = ctk.CTkLabel(linha, text="Nenhum arquivo", text_color=CINZA)
        self.lbl_qtd.pack(side="left", padx=8)

        self.lista = ctk.CTkTextbox(card, height=150, fg_color="#F7F9FC", corner_radius=10, text_color=TEXTO,
                                    font=ctk.CTkFont(family="Consolas", size=12))
        self.lista.pack(fill="x", padx=16, pady=(10, 6))
        self.lista.configure(state="disabled")

        per = ctk.CTkFrame(card, fg_color="transparent")
        per.pack(anchor="w", padx=16, pady=(2, 6))
        ctk.CTkLabel(per, text="Nome da geral (opcional):", text_color=TEXTO).pack(side="left")
        self.ent_nome = ctk.CTkEntry(per, width=320, placeholder_text="Conferência geral <data e hora>")
        self.ent_nome.pack(side="left", padx=8)

        self.bt_juntar = self._botao_cheio(card, "Juntar numa geral", self._iniciar,
                                           height=46, corner_radius=12, font=self._fonte(16, True))
        self.bt_juntar.pack(fill="x", padx=16, pady=(8, 16))

        self.log = ctk.CTkTextbox(corpo, height=200, fg_color="white", corner_radius=12, text_color=TEXTO,
                                  font=ctk.CTkFont(family="Consolas", size=12))
        self.log.pack(fill="x", padx=28, pady=(12, 0))
        self.log.tag_config("erro", foreground=VERMELHO)
        self.log.tag_config("aviso", foreground=LARANJA)
        self.log.tag_config("ok", foreground=VERDE)
        self.log.configure(state="disabled")

        rodape = ctk.CTkFrame(corpo, fg_color="transparent")
        rodape.pack(pady=(12, 20))
        self.bt_resultado = self._botao_contorno(rodape, "Abrir resultado", self._abrir_resultado, state="disabled")
        self.bt_resultado.pack(side="left", padx=5)
        self.bt_pasta = self._botao_contorno(rodape, "Abrir pasta", self._abrir_pasta, state="disabled", width=100)
        self.bt_pasta.pack(side="left", padx=5)

    # ---------------------------------------------------------------- lista
    def _adicionar(self):
        inicial = str(Path(self.arquivos[-1]).parent) if self.arquivos else str(Path.home() / "Downloads")
        novos = filedialog.askopenfilenames(parent=self, title="Planilhas de conferência", initialdir=inicial,
                                            filetypes=[("Planilha Excel", "*.xlsx *.xlsm")])
        for n in novos:
            if n not in self.arquivos:
                self.arquivos.append(n)
        self._mostrar_lista()

    def _limpar(self):
        self.arquivos = []
        self._mostrar_lista()

    def _mostrar_lista(self):
        self.lista.configure(state="normal")
        self.lista.delete("1.0", "end")
        for i, a in enumerate(self.arquivos, start=1):
            self.lista.insert("end", f"{i:>2}. {Path(a).name}\n")
        self.lista.configure(state="disabled")
        self.lbl_qtd.configure(text=(f"{len(self.arquivos)} arquivo(s)" if self.arquivos else "Nenhum arquivo"),
                               text_color=(NAVY if self.arquivos else CINZA))

    # ---------------------------------------------------------------- execução
    def _iniciar(self):
        if len(self.arquivos) < 2:
            messagebox.showwarning("União", "Adicione pelo menos 2 planilhas para juntar.", parent=self)
            return
        nome = self.ent_nome.get().strip() or f"Conferência geral {dt.datetime.now():%Y-%m-%d %H%M}"
        if not nome.lower().endswith(".xlsx"):
            nome += ".xlsx"
        destino = str(Path(self.arquivos[0]).parent / nome)
        arquivos = list(self.arquivos)
        self.bt_juntar.configure(state="disabled", text="Juntando...")
        self.resultado = None
        self._escrever(f"▶ Juntando {len(arquivos)} planilha(s)...", "ok")

        def trabalho():
            try:
                saida = uniao.combinar(arquivos, destino, log=lambda m: self.fila.put(("log", str(m))))
                self.fila.put(("fim", saida))
            except PermissionError:
                self.fila.put(("erro", "Não consegui salvar. Feche a planilha aberta e tente de novo."))
            except Exception as e:  # noqa: BLE001
                self.fila.put(("erro", str(e).splitlines()[0] if str(e) else repr(e)))

        self.thread = threading.Thread(target=trabalho, daemon=True)
        self.thread.start()

    def _fim(self, destino):
        self.resultado = destino
        self.bt_juntar.configure(state="normal", text="Juntar numa geral")
        self.bt_resultado.configure(state="normal")
        self.bt_pasta.configure(state="normal")
        if messagebox.askyesno("Concluído", f"Pronto.\n\n{Path(destino).name}\n\nAbrir agora?", parent=self):
            abrir_arquivo(destino)

    def _abrir_resultado(self):
        if self.resultado and Path(self.resultado).exists():
            abrir_arquivo(self.resultado)

    def _abrir_pasta(self):
        if self.resultado:
            abrir_arquivo(Path(self.resultado).parent)

    def _fechar(self):
        if self.thread and self.thread.is_alive():
            if not messagebox.askyesno("Rodando", "Ainda está juntando. Fechar mesmo assim?", parent=self):
                return
        self.destroy()

    # ---------------------------------------------------------------- log
    def _escrever(self, texto, tag=None):
        if tag is None:
            t = texto.lstrip()
            if t.startswith(("ERRO", "✖")):
                tag = "erro"
            elif t.startswith("AVISO"):
                tag = "aviso"
            elif t.startswith(("OK", "Pronto")):
                tag = "ok"
        self.log.configure(state="normal")
        self.log.insert("end", f"{time.strftime('%H:%M:%S')}  {texto}\n", tag)
        self.log.see("end")
        self.log.configure(state="disabled")

    def _ler_fila(self):
        try:
            while True:
                tipo, valor = self.fila.get_nowait()
                if tipo == "log":
                    self._escrever(valor)
                elif tipo == "fim":
                    self._fim(valor)
                elif tipo == "erro":
                    self.bt_juntar.configure(state="normal", text="Juntar numa geral")
                    self._escrever(f"✖ {valor}", "erro")
                    messagebox.showerror("Erro", valor, parent=self)
        except queue.Empty:
            pass
        try:
            if self.winfo_exists():
                self.after(150, self._ler_fila)
        except Exception:
            pass


class App(_JanelaUniao, ctk.CTk):
    """Aberto sozinho."""


class UniaoWindow(_JanelaUniao, ctk.CTkToplevel):
    """Aberto como janela da Central."""

    def __init__(self, parent):
        super().__init__(parent)
        self.after(250, self._trazer_para_frente)

    def _trazer_para_frente(self):
        self.lift()
        self.focus_force()


_janela_aberta = None


def abrir(parent):
    global _janela_aberta
    if _janela_aberta is not None and _janela_aberta.winfo_exists():
        _janela_aberta.deiconify()
        _janela_aberta._trazer_para_frente()
        return _janela_aberta
    _janela_aberta = UniaoWindow(parent)
    return _janela_aberta


if __name__ == "__main__":
    App().mainloop()
