"""
Robô Porto Saúde - janela da conferência das abas PORTO e PORTO CONV com o portal.

Abre pela Central de Automações (botão "Robô Porto") ou sozinho: python interface_porto.py
"""
from __future__ import annotations

import datetime as dt
import json
import os
import queue
import sys
import threading
import time
from pathlib import Path
from tkinter import filedialog, messagebox

import customtkinter as ctk

import robo_porto as robo

ctk.set_appearance_mode("light")
ctk.set_default_color_theme("blue")

ARQ_PREFS = robo.PASTA_DADOS / "preferencias.json"  # últimos arquivos e opções (não há senha: o login é manual)

NAVY = "#002060"
NAVY_HOVER = "#0A3A8C"
BG = "#F3F5FA"
CINZA = "#7A869A"
CINZA_ESCURO = "#51607A"
TEXTO = "#243B53"
CONTORNO_HOVER = "#E3E8F4"
NAO_SELECIONADO = "#9AA8C2"
NAO_SELECIONADO_HOVER = "#7F8FAE"
VERMELHO = "#C00000"
LARANJA = "#B26A00"
VERDE = "#2E7D32"
TODAS = "Todas"


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


def _curto(p, n=70):
    p = str(p)
    return p if len(p) <= n else "…" + p[-(n - 1):]


class _JanelaPorto:
    """Janela do robô. Vira App (sozinho) ou PortoWindow (aberta pela Central)."""

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.title("Robô Porto Saúde")
        altura = max(640, min(900, self.winfo_screenheight() - 90))
        self.geometry(f"860x{altura}+60+20")
        self.minsize(760, 600)
        self.configure(fg_color=BG)

        self.fila = queue.Queue()
        self.parar = threading.Event()
        self.thread = None
        self.planilha = ""
        self.resultado = None
        self.prefs = {}

        self._montar_tela()
        self._carregar_prefs()
        self.protocol("WM_DELETE_WINDOW", self._fechar)
        self.after(150, self._ler_fila)

    # ================================================================= tela
    def _fonte(self, tamanho=13, negrito=False):
        return ctk.CTkFont(size=tamanho, weight="bold" if negrito else "normal")

    def _cartao(self, pai, **pack):
        card = ctk.CTkFrame(pai, fg_color="white", corner_radius=16)
        card.pack(fill="x", padx=28, **pack)
        return card

    def _titulo(self, card, texto):
        ctk.CTkLabel(card, text=texto, text_color=NAVY, font=self._fonte(14, True)).pack(
            anchor="w", padx=20, pady=(16, 6))

    def _botao_contorno(self, pai, texto, comando, **kw):
        return ctk.CTkButton(pai, text=texto, command=comando, fg_color="transparent", text_color=NAVY,
                             hover_color=CONTORNO_HOVER, border_width=1, border_color=NAVY,
                             text_color_disabled="#A9B3C4", **kw)

    def _botao_cheio(self, pai, texto, comando, **kw):
        return ctk.CTkButton(pai, text=texto, command=comando, fg_color=NAVY, hover_color=NAVY_HOVER, **kw)

    def _montar_tela(self):
        corpo = ctk.CTkScrollableFrame(self, fg_color=BG)
        corpo.pack(fill="both", expand=True)

        ctk.CTkLabel(corpo, text="Robô Porto Saúde", text_color=NAVY, font=self._fonte(26, True)).pack(pady=(20, 2))
        ctk.CTkLabel(corpo, text="Confere as guias das abas PORTO e PORTO CONV no portal do prestador.",
                     text_color=CINZA_ESCURO, font=self._fonte(14)).pack(pady=(0, 16))

        # ---- 1. planilha
        c1 = self._cartao(corpo)
        self._titulo(c1, "1. Planilha de conferência (abas PORTO e PORTO CONV)")
        linha = ctk.CTkFrame(c1, fg_color="transparent")
        linha.pack(fill="x", padx=20, pady=(0, 16))
        self.lbl_planilha = ctk.CTkLabel(linha, text="Nenhum arquivo selecionado", text_color=CINZA, anchor="w")
        self.lbl_planilha.pack(side="left", fill="x", expand=True)
        self._botao_cheio(linha, "Escolher arquivo", self._escolher_planilha, width=150).pack(side="right")

        # ---- 2. o que conferir
        c2 = self._cartao(corpo, pady=(12, 0))
        self._titulo(c2, "2. O que conferir")
        abas = ctk.CTkFrame(c2, fg_color="#EEF2FA", corner_radius=10)
        abas.pack(anchor="w", padx=20, pady=(0, 10))
        self.var_aba = ctk.BooleanVar(value=True)
        self.var_conv = ctk.BooleanVar(value=True)
        for texto, var in (("ABA (aba PORTO) - soma das guias x TOTAL MÊS", self.var_aba),
                           ("Convencional (aba PORTO CONV) - procedimento e quantidade", self.var_conv)):
            ctk.CTkCheckBox(abas, text=texto, variable=var, fg_color=NAVY, hover_color=NAVY_HOVER,
                            text_color=TEXTO, command=self._mostrar_contagem).pack(anchor="w", padx=12, pady=6)
        lu = ctk.CTkFrame(c2, fg_color="transparent")
        lu.pack(anchor="w", padx=20, pady=(0, 8))
        ctk.CTkLabel(lu, text="Unidade", text_color=TEXTO).pack(side="left")
        self.seg_unidade = ctk.CTkSegmentedButton(
            lu, values=[TODAS], command=lambda v: self._mostrar_contagem(),
            selected_color=NAVY, selected_hover_color=NAVY_HOVER, unselected_color=NAO_SELECIONADO,
            unselected_hover_color=NAO_SELECIONADO_HOVER, fg_color=NAO_SELECIONADO, text_color="white",
            font=self._fonte(13, True))
        self.seg_unidade.set(TODAS)
        self.seg_unidade.pack(side="left", padx=12)
        self.lbl_contagem = ctk.CTkLabel(lu, text="", text_color=CINZA)
        self.lbl_contagem.pack(side="left")
        lt = ctk.CTkFrame(c2, fg_color="transparent")
        lt.pack(anchor="w", padx=20, pady=(0, 8))
        ctk.CTkLabel(lt, text="Testar só as primeiras", text_color=TEXTO).pack(side="left")
        self.ent_limite = ctk.CTkEntry(lt, width=60, justify="center")
        self.ent_limite.insert(0, "0")
        self.ent_limite.pack(side="left", padx=8)
        ctk.CTkLabel(lt, text="linhas (0 = todas)", text_color=CINZA).pack(side="left")
        la = ctk.CTkFrame(c2, fg_color="transparent")
        la.pack(anchor="w", padx=20, pady=(0, 16))
        ctk.CTkLabel(la, text="Abas no portal ao mesmo tempo", text_color=TEXTO).pack(side="left")
        self.seg_abas = ctk.CTkSegmentedButton(
            la, values=[str(n) for n in range(1, robo.MAX_ABAS + 1)], selected_color=NAVY,
            selected_hover_color=NAVY_HOVER, unselected_color=NAO_SELECIONADO,
            unselected_hover_color=NAO_SELECIONADO_HOVER, fg_color=NAO_SELECIONADO, text_color="white",
            font=self._fonte(13, True), width=200)
        self.seg_abas.set("3")
        self.seg_abas.pack(side="left", padx=12)
        ctk.CTkLabel(la, text="mais abas = mais rápido (se o portal reclamar, use 1)",
                     text_color=CINZA).pack(side="left")

        # ---- 3. período
        c3 = self._cartao(corpo, pady=(12, 0))
        self._titulo(c3, "3. Período da pesquisa no portal (data da solicitação)")
        per = ctk.CTkFrame(c3, fg_color="transparent")
        per.pack(anchor="w", padx=20, pady=(0, 6))
        ctk.CTkLabel(per, text="de", text_color=TEXTO).pack(side="left")
        self.ent_inicio = ctk.CTkEntry(per, width=120, placeholder_text="dd/mm/aaaa", justify="center")
        self.ent_inicio.pack(side="left", padx=8)
        ctk.CTkLabel(per, text="até", text_color=TEXTO).pack(side="left")
        self.ent_fim = ctk.CTkEntry(per, width=120, placeholder_text="dd/mm/aaaa", justify="center")
        self.ent_fim.pack(side="left", padx=8)
        self._botao_contorno(per, "Últimos 90 dias", self._ultimos_90, width=130, height=28).pack(side="left", padx=8)
        ctk.CTkLabel(c3, text="O portal pesquisa no máximo 90 dias por vez. Se o período for maior, o robô "
                              "pesquisa em partes (mais demorado).",
                     text_color=CINZA, wraplength=740, justify="left").pack(anchor="w", padx=20, pady=(0, 16))

        # ---- 4. como funciona + botão
        c4 = self._cartao(corpo, pady=(12, 0))
        self._titulo(c4, "4. Conferir")
        ctk.CTkLabel(c4, text="O robô abre o navegador no portal da Porto. Faça o LOGIN você mesmo (CPF e senha) "
                              "e ele continua sozinho: pesquisa cada guia, abre o PDF dela e compara com a "
                              "planilha. O resultado sai numa planilha nova, na mesma pasta da planilha de "
                              "conferência. Não mexa no navegador enquanto ele trabalha.",
                     text_color=CINZA_ESCURO, wraplength=740, justify="left").pack(anchor="w", padx=20)
        self.bt_conferir = self._botao_cheio(c4, "Abrir o portal e conferir", self._iniciar, height=46,
                                             corner_radius=12, font=self._fonte(16, True))
        self.bt_conferir.pack(fill="x", padx=20, pady=(14, 18))

        # ---- andamento e log
        andamento = ctk.CTkFrame(corpo, fg_color="transparent")
        andamento.pack(fill="x", padx=28, pady=(14, 0))
        self.lbl_status = ctk.CTkLabel(andamento, text="Pronto para começar.", text_color=CINZA_ESCURO, anchor="w")
        self.lbl_status.pack(fill="x")
        self.barra = ctk.CTkProgressBar(andamento, progress_color=NAVY, height=10)
        self.barra.set(0)
        self.barra.pack(fill="x", pady=(4, 0))
        self.log = ctk.CTkTextbox(corpo, height=190, fg_color="white", corner_radius=12, text_color=TEXTO,
                                  font=ctk.CTkFont(family="Consolas", size=12))
        self.log.pack(fill="x", padx=28, pady=(10, 0))
        self.log.tag_config("erro", foreground=VERMELHO)
        self.log.tag_config("aviso", foreground=LARANJA)
        self.log.tag_config("ok", foreground=VERDE)
        self.log.configure(state="disabled")

        rodape = ctk.CTkFrame(corpo, fg_color="transparent")
        rodape.pack(pady=(12, 20))
        self.bt_parar = self._botao_contorno(rodape, "■  Parar", self._parar, state="disabled", width=90)
        self.bt_parar.pack(side="left", padx=5)
        self.bt_resultado = self._botao_contorno(rodape, "Abrir resultado", self._abrir_resultado, state="disabled")
        self.bt_resultado.pack(side="left", padx=5)
        self.bt_pasta = self._botao_contorno(rodape, "Abrir pasta", self._abrir_pasta, state="disabled", width=100)
        self.bt_pasta.pack(side="left", padx=5)

    # ================================================================= campos e preferências
    @staticmethod
    def _set(entry, texto):
        entry.delete(0, "end")
        if texto:
            entry.insert(0, texto)

    def _ultimos_90(self):
        hoje = dt.date.today()
        self._set(self.ent_inicio, f"{hoje - dt.timedelta(days=robo.MAX_DIAS):%d/%m/%Y}")
        self._set(self.ent_fim, f"{hoje:%d/%m/%Y}")

    def _carregar_prefs(self):
        try:
            self.prefs = json.loads(ARQ_PREFS.read_text(encoding="utf-8"))
        except Exception:
            self.prefs = {}
        if self.prefs.get("planilha") and Path(self.prefs["planilha"]).exists():
            self._usar_planilha(self.prefs["planilha"])
        self.var_aba.set(self.prefs.get("conferir_aba", True))
        self.var_conv.set(self.prefs.get("conferir_conv", True))
        self._set(self.ent_inicio, self.prefs.get("inicio", ""))
        self._set(self.ent_fim, self.prefs.get("fim", ""))
        if not self.ent_inicio.get():
            self._ultimos_90()
        if str(self.prefs.get("abas_portal", "")) in self.seg_abas.cget("values"):
            self.seg_abas.set(str(self.prefs["abas_portal"]))
        und = self.prefs.get("unidade", TODAS)
        if und in self.seg_unidade.cget("values"):
            self.seg_unidade.set(und)
        self._mostrar_contagem()

    def _salvar_prefs(self):
        self.prefs.update({"planilha": self.planilha, "conferir_aba": bool(self.var_aba.get()),
                           "conferir_conv": bool(self.var_conv.get()), "unidade": self.seg_unidade.get(),
                           "abas_portal": self.seg_abas.get(),
                           "inicio": self.ent_inicio.get().strip(), "fim": self.ent_fim.get().strip()})
        try:
            ARQ_PREFS.parent.mkdir(parents=True, exist_ok=True)
            ARQ_PREFS.write_text(json.dumps(self.prefs, ensure_ascii=False, indent=2), encoding="utf-8")
        except Exception:
            pass

    def _escolher_planilha(self):
        inicial = str(Path(self.planilha).parent) if self.planilha else str(Path.home() / "Downloads")
        caminho = filedialog.askopenfilename(parent=self, title="Planilha de conferência", initialdir=inicial,
                                             filetypes=[("Planilha Excel", "*.xlsx *.xlsm")])
        if caminho:
            self._usar_planilha(caminho)

    def _usar_planilha(self, caminho):
        self.planilha = caminho
        self.lbl_planilha.configure(text=_curto(Path(caminho).name), text_color=NAVY)
        self._registros = {}
        for aba in robo.ABAS:
            try:
                self._registros[aba] = robo.ler_registros(caminho, None, (aba,))
            except Exception as e:
                self._registros[aba] = []
                self._escrever(f"AVISO {e}")
        unidades = sorted({r["unidade"] for rs in self._registros.values() for r in rs if r["unidade"]}, key=int)
        atual = self.seg_unidade.get()
        self.seg_unidade.configure(values=[TODAS] + unidades)
        self.seg_unidade.set(atual if atual in [TODAS] + unidades else TODAS)
        self._mostrar_contagem()

    def _abas(self):
        return tuple(a for a, v in ((robo.ABA, self.var_aba), (robo.ABA_CONV, self.var_conv)) if v.get())

    def _unidade(self):
        u = self.seg_unidade.get()
        return None if u == TODAS else u

    def _mostrar_contagem(self):
        regs = getattr(self, "_registros", None)
        if not regs:
            self.lbl_contagem.configure(text="escolha a planilha")
            return
        und = self._unidade()
        partes = []
        for aba in self._abas():
            rs = [r for r in regs.get(aba, []) if und is None or r["unidade"] == und]
            guias = {g for r in rs for g in r["guias"]}
            partes.append(f"{len(rs)} linha(s) / {len(guias)} guia(s) na {aba}")
        self.lbl_contagem.configure(text=" | ".join(partes) or "marque o que conferir")

    # ================================================================= execução
    def _iniciar(self):
        if not self.planilha or not Path(self.planilha).exists():
            messagebox.showwarning("Planilha", "Escolha a planilha de conferência.", parent=self)
            return
        abas = self._abas()
        if not abas:
            messagebox.showwarning("Conferir", "Marque ABA (PORTO), Convencional (PORTO CONV) ou os dois.",
                                   parent=self)
            return
        inicio, fim = robo.para_data(self.ent_inicio.get()), robo.para_data(self.ent_fim.get())
        if not inicio or not fim:
            messagebox.showwarning("Período", "Preencha o período no formato dd/mm/aaaa.", parent=self)
            return
        try:
            partes = robo.dividir_periodo(inicio, fim)
        except ValueError as e:
            messagebox.showwarning("Período", str(e), parent=self)
            return
        if len(partes) > 1 and not messagebox.askyesno(
                "Período longo", f"O período tem {(fim - inicio).days + 1} dias e o portal pesquisa no máximo 90 "
                f"por vez. As guias fora dos últimos 90 dias vão precisar de até {len(partes)} pesquisas cada "
                "(mais demorado).\n\nContinuar?", parent=self):
            return
        try:
            limite = max(0, int(self.ent_limite.get() or 0))
        except ValueError:
            limite = 0
        und = self._unidade()
        n_abas = int(self.seg_abas.get() or 1)
        quais = " e ".join("PORTO" if a == robo.ABA else "PORTO CONV" for a in abas)
        if not messagebox.askyesno(
                "Confirmar", f"Conferir {quais} ({'todas as unidades' if und is None else 'unidade ' + und}) "
                f"pesquisando de {inicio:%d/%m/%Y} a {fim:%d/%m/%Y}?\n\nO navegador vai abrir: faça o login "
                "nele e o robô continua sozinho.", parent=self):
            return
        self._salvar_prefs()

        self.parar.clear()
        self.resultado = None
        self._set_ocupado(True)
        self._status("Abrindo o portal... faça o login no navegador.")
        self._escrever(f"▶ Conferência Porto - {quais} - {inicio:%d/%m/%Y} a {fim:%d/%m/%Y}"
                       + (f" - unidade {und}" if und else "") + f" - {n_abas} aba(s)", "ok")
        planilha = self.planilha

        def trabalho():
            try:
                destino = robo.executar(
                    planilha, inicio, fim, abas=abas, unidade=und, limite=limite, parar=self.parar, n_abas=n_abas,
                    log=lambda m: self.fila.put(("log", str(m))),
                    progresso=lambda f, t: self.fila.put(("progresso", (f, t))))
                self.fila.put(("fim", destino))
            except PermissionError:
                self.fila.put(("erro", "Não consegui salvar o resultado. Feche a planilha aberta e tente de novo."))
            except Exception as e:  # noqa: BLE001
                self.fila.put(("erro", str(e).splitlines()[0] if str(e) else repr(e)))

        self.thread = threading.Thread(target=trabalho, daemon=True)
        self.thread.start()

    def _set_ocupado(self, ocupado):
        self.bt_conferir.configure(state="disabled" if ocupado else "normal")
        self.bt_parar.configure(state="normal" if ocupado else "disabled")
        if ocupado:
            self.barra.set(0)

    def _parar(self):
        self.parar.set()
        self.bt_parar.configure(state="disabled")
        self._status("Parando depois da guia atual...")

    def _fim(self, destino):
        self.resultado = destino
        self._set_ocupado(False)
        self.barra.set(1)
        self.bt_resultado.configure(state="normal")
        self.bt_pasta.configure(state="normal")
        self._status(f"Conferência concluída: {Path(destino).name}")
        if messagebox.askyesno("Concluído", f"Conferência finalizada.\n\n{Path(destino).name}\n\nAbrir agora?",
                               parent=self):
            abrir_arquivo(destino)

    def _abrir_resultado(self):
        if self.resultado and Path(self.resultado).exists():
            abrir_arquivo(self.resultado)

    def _abrir_pasta(self):
        if self.resultado:
            abrir_arquivo(Path(self.resultado).parent)

    def _fechar(self):
        if self.thread and self.thread.is_alive():
            if not messagebox.askyesno("Robô rodando", "O robô ainda está rodando. Fechar mesmo assim?", parent=self):
                return
            self.parar.set()
        self._salvar_prefs()
        self.destroy()

    # ================================================================= log e fila
    def _status(self, texto):
        self.lbl_status.configure(text=texto)

    def _escrever(self, texto, tag=None):
        if tag is None:
            t = texto.lstrip()
            if t.startswith(("ERRO", "✖")):
                tag = "erro"
            elif t.startswith("AVISO"):
                tag = "aviso"
            elif t.startswith(("OK", "Pronto", "Login OK")):
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
                    if "Login OK" in valor:
                        self._status("Login feito. Pesquisando as guias...")
                    elif "Faça o login" in valor:
                        self._status("Aguardando você fazer o login no navegador...")
                elif tipo == "progresso":
                    feitas, total = valor
                    self.barra.set(feitas / total if total else 0)
                    self._status(f"{feitas} de {total} guia(s) pesquisada(s)")
                elif tipo == "fim":
                    self._fim(valor)
                elif tipo == "erro":
                    self._set_ocupado(False)
                    self._escrever(f"✖ {valor}", "erro")
                    self._status("Parou com erro (veja abaixo).")
                    messagebox.showerror("Erro", valor, parent=self)
        except queue.Empty:
            pass
        try:
            if self.winfo_exists():
                self.after(150, self._ler_fila)
        except Exception:
            pass


class App(_JanelaPorto, ctk.CTk):
    """Robô aberto sozinho."""


class PortoWindow(_JanelaPorto, ctk.CTkToplevel):
    """Robô aberto como janela da Central de Automações."""

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
    _janela_aberta = PortoWindow(parent)
    return _janela_aberta


if __name__ == "__main__":
    App().mainloop()
