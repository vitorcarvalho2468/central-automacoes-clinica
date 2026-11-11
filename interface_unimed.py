"""
Robô Unimed Guarulhos - janela das duas funções (pela planilha e por período).

Abre pela Central de Automações (botão "Robô Unimed") ou sozinho: python interface_unimed.py
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

import robo_unimed as robo

ctk.set_appearance_mode("light")
ctk.set_default_color_theme("blue")

ARQ_PREFS = robo.PASTA_DADOS / "preferencias.json"  # últimos arquivos e opções (sem senha: o login é manual)

NAVY, NAVY_HOVER, BG = "#002060", "#0A3A8C", "#F3F5FA"
CINZA, CINZA_ESCURO, TEXTO = "#7A869A", "#51607A", "#243B53"
CONTORNO_HOVER = "#E3E8F4"
NAO_SELECIONADO, NAO_SELECIONADO_HOVER = "#9AA8C2", "#7F8FAE"
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


def _curto(p, n=70):
    p = str(p)
    return p if len(p) <= n else "…" + p[-(n - 1):]


class _JanelaUnimed:
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.title("Robô Unimed Guarulhos")
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

    def _botao_contorno(self, pai, texto, comando, **kw):
        return ctk.CTkButton(pai, text=texto, command=comando, fg_color="transparent", text_color=NAVY,
                             hover_color=CONTORNO_HOVER, border_width=1, border_color=NAVY,
                             text_color_disabled="#A9B3C4", **kw)

    def _botao_cheio(self, pai, texto, comando, **kw):
        return ctk.CTkButton(pai, text=texto, command=comando, fg_color=NAVY, hover_color=NAVY_HOVER, **kw)

    def _montar_tela(self):
        corpo = ctk.CTkScrollableFrame(self, fg_color=BG)
        corpo.pack(fill="both", expand=True)
        ctk.CTkLabel(corpo, text="Robô Unimed Guarulhos", text_color=NAVY, font=self._fonte(26, True)).pack(pady=(20, 2))
        ctk.CTkLabel(corpo, text="Finaliza as guias em aberto no SGU-Card (Finalizar Parcial → Confirmar).",
                     text_color=CINZA_ESCURO, font=self._fonte(14)).pack(pady=(0, 16))

        self.abas = ctk.CTkTabview(corpo, fg_color="white", corner_radius=16, height=330,
                                   segmented_button_fg_color=NAO_SELECIONADO,
                                   segmented_button_selected_color=NAVY,
                                   segmented_button_selected_hover_color=NAVY_HOVER,
                                   segmented_button_unselected_color=NAO_SELECIONADO,
                                   segmented_button_unselected_hover_color=NAO_SELECIONADO_HOVER,
                                   text_color="white")
        self.abas._segmented_button.configure(font=self._fonte(13, True))
        self.abas.pack(fill="x", padx=28)
        a1 = self.abas.add("Pela planilha")
        a2 = self.abas.add("Por período")

        # ---- aba 1
        ctk.CTkLabel(a1, text="Pega o nº da guia da coluna Observação, abre a guia, confere se tem leitura no DIA da "
                              "Data da Agenda e, se tiver, finaliza. Devolve a planilha com Status (QR Lido / QR Não "
                              "lido / SEM TR) e Status Detalhado. Falta com leitura: QR Lido + aviso em amarelo.",
                     text_color=CINZA_ESCURO, wraplength=740, justify="left").pack(anchor="w", padx=12, pady=(8, 10))
        linha = ctk.CTkFrame(a1, fg_color="transparent")
        linha.pack(fill="x", padx=12)
        self.lbl_planilha = ctk.CTkLabel(linha, text="Nenhum arquivo selecionado", text_color=CINZA, anchor="w")
        self.lbl_planilha.pack(side="left", fill="x", expand=True)
        self._botao_cheio(linha, "Escolher planilha", self._escolher_planilha, width=150).pack(side="right")
        self.bt_planilha = self._botao_cheio(a1, "Abrir o portal e finalizar pela planilha", self._iniciar_planilha,
                                             height=46, corner_radius=12, font=self._fonte(16, True))
        self.bt_planilha.pack(fill="x", padx=12, pady=(18, 8), side="bottom")

        # ---- aba 2
        ctk.CTkLabel(a2, text="Pesquisa Exames em aberto só pela Data de atendimento (de/até), junta as guias de "
                              "todas as páginas e abre uma por uma: se tiver ALGUMA data de assinatura preenchida, faz "
                              "Finalizar Parcial → Confirmar (qualquer data); sem assinatura, passa para a próxima. "
                              "O resultado vai para a Área de Trabalho.",
                     text_color=CINZA_ESCURO, wraplength=740, justify="left").pack(anchor="w", padx=12, pady=(8, 10))
        per = ctk.CTkFrame(a2, fg_color="transparent")
        per.pack(anchor="w", padx=12)
        ctk.CTkLabel(per, text="Data de atendimento de", text_color=TEXTO).pack(side="left")
        self.ent_inicio = ctk.CTkEntry(per, width=120, placeholder_text="dd/mm/aaaa", justify="center")
        self.ent_inicio.pack(side="left", padx=8)
        ctk.CTkLabel(per, text="até", text_color=TEXTO).pack(side="left")
        self.ent_fim = ctk.CTkEntry(per, width=120, placeholder_text="dd/mm/aaaa", justify="center")
        self.ent_fim.pack(side="left", padx=8)
        self._botao_contorno(per, "Ontem", self._ontem, width=80, height=28).pack(side="left", padx=8)
        self.bt_periodo = self._botao_cheio(a2, "Abrir o portal e finalizar o período", self._iniciar_periodo,
                                            height=46, corner_radius=12, font=self._fonte(16, True))
        self.bt_periodo.pack(fill="x", padx=12, pady=(18, 8), side="bottom")

        # ---- opções comuns
        op = ctk.CTkFrame(corpo, fg_color="white", corner_radius=16)
        op.pack(fill="x", padx=28, pady=(12, 0))
        lt = ctk.CTkFrame(op, fg_color="transparent")
        lt.pack(anchor="w", padx=20, pady=(14, 6))
        ctk.CTkLabel(lt, text="Testar só as primeiras", text_color=TEXTO).pack(side="left")
        self.ent_limite = ctk.CTkEntry(lt, width=60, justify="center")
        self.ent_limite.insert(0, "0")
        self.ent_limite.pack(side="left", padx=8)
        ctk.CTkLabel(lt, text="guias (0 = todas)", text_color=CINZA).pack(side="left")
        la = ctk.CTkFrame(op, fg_color="transparent")
        la.pack(anchor="w", padx=20, pady=(0, 6))
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
        self.sw_simular = ctk.CTkSwitch(op, text="Simulação: confere tudo, mas NÃO clica em Finalizar Parcial",
                                        progress_color=NAVY, text_color=TEXTO)
        self.sw_simular.pack(anchor="w", padx=20, pady=(0, 6))
        ctk.CTkLabel(op, text="O navegador abre no SGU-Card: faça o LOGIN você mesmo e o robô continua sozinho. "
                              "Não mexa no navegador enquanto ele trabalha. Finalizar não tem volta: na primeira vez, "
                              "teste com 1 guia e confira no portal.",
                     text_color=CINZA, wraplength=740, justify="left").pack(anchor="w", padx=20, pady=(0, 14))

        # ---- andamento e log
        andamento = ctk.CTkFrame(corpo, fg_color="transparent")
        andamento.pack(fill="x", padx=28, pady=(14, 0))
        self.lbl_status = ctk.CTkLabel(andamento, text="Pronto para começar.", text_color=CINZA_ESCURO, anchor="w")
        self.lbl_status.pack(fill="x")
        self.barra = ctk.CTkProgressBar(andamento, progress_color=NAVY, height=10)
        self.barra.set(0)
        self.barra.pack(fill="x", pady=(4, 0))
        self.log = ctk.CTkTextbox(corpo, height=200, fg_color="white", corner_radius=12, text_color=TEXTO,
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

    def _ontem(self):
        d = dt.date.today() - dt.timedelta(days=1)
        self._set(self.ent_inicio, f"{d:%d/%m/%Y}")
        self._set(self.ent_fim, f"{d:%d/%m/%Y}")

    def _carregar_prefs(self):
        try:
            self.prefs = json.loads(ARQ_PREFS.read_text(encoding="utf-8"))
        except Exception:
            self.prefs = {}
        if self.prefs.get("planilha") and Path(self.prefs["planilha"]).exists():
            self._usar_planilha(self.prefs["planilha"])
        self._set(self.ent_inicio, self.prefs.get("inicio", ""))
        self._set(self.ent_fim, self.prefs.get("fim", ""))
        if str(self.prefs.get("abas_portal", "")) in self.seg_abas.cget("values"):
            self.seg_abas.set(str(self.prefs["abas_portal"]))

    def _salvar_prefs(self):
        self.prefs.update({"planilha": self.planilha, "inicio": self.ent_inicio.get().strip(),
                           "fim": self.ent_fim.get().strip(), "abas_portal": self.seg_abas.get()})
        try:
            ARQ_PREFS.parent.mkdir(parents=True, exist_ok=True)
            ARQ_PREFS.write_text(json.dumps(self.prefs, ensure_ascii=False, indent=2), encoding="utf-8")
        except Exception:
            pass

    def _escolher_planilha(self):
        inicial = str(Path(self.planilha).parent) if self.planilha else str(Path.home() / "Downloads")
        caminho = filedialog.askopenfilename(parent=self, title="Planilha da agenda (Unimed)", initialdir=inicial,
                                             filetypes=[("Planilha Excel", "*.xlsx *.xlsm")])
        if caminho:
            self._usar_planilha(caminho)

    def _usar_planilha(self, caminho):
        self.planilha = caminho
        try:
            _, _, _, linhas = robo.ler_planilha(caminho)
            guias = {g for l in linhas for g in l["guias"]}
            info = f"  ({len(linhas)} linhas, {len(guias)} guias)"
        except Exception as e:  # noqa: BLE001
            info = ""
            self._escrever(f"AVISO {e}")
        self.lbl_planilha.configure(text=_curto(Path(caminho).name, 60) + info, text_color=NAVY)

    def _n_abas(self):
        try:
            return int(self.seg_abas.get() or 1)
        except ValueError:
            return 1

    def _limite(self):
        try:
            return max(0, int(self.ent_limite.get() or 0))
        except ValueError:
            return 0

    # ================================================================= execução
    def _iniciar_planilha(self):
        if not self.planilha or not Path(self.planilha).exists():
            messagebox.showwarning("Planilha", "Escolha a planilha da agenda.", parent=self)
            return
        limite, simular, n_abas = self._limite(), bool(self.sw_simular.get()), self._n_abas()
        quanto = f"as primeiras {limite} guia(s)" if limite else "todas as guias"
        acao = "CONFERIR (simulação, sem finalizar)" if simular else "conferir e FINALIZAR"
        if not messagebox.askyesno("Confirmar", f"{acao.capitalize()} {quanto} da planilha\n{Path(self.planilha).name}?"
                                   "\n\nO navegador vai abrir: faça o login nele e o robô continua sozinho.", parent=self):
            return
        planilha = self.planilha
        self._rodar(f"▶ Pela planilha - {Path(planilha).name} - {n_abas} aba(s)" + (" - SIMULAÇÃO" if simular else ""),
                    lambda log, prog: robo.executar_planilha(planilha, limite, simular, log=log, parar=self.parar,
                                                             progresso=prog, n_abas=n_abas))

    def _iniciar_periodo(self):
        inicio, fim = robo.para_data(self.ent_inicio.get()), robo.para_data(self.ent_fim.get())
        if not inicio or not fim:
            messagebox.showwarning("Período", "Preencha as datas no formato dd/mm/aaaa.", parent=self)
            return
        if fim < inicio:
            messagebox.showwarning("Período", "A data final é anterior à inicial.", parent=self)
            return
        limite, simular, n_abas = self._limite(), bool(self.sw_simular.get()), self._n_abas()
        quanto = f"as primeiras {limite} guia(s)" if limite else "TODAS as guias"
        if not messagebox.askyesno(
                "Confirmar", f"Abrir {quanto} em aberto com data de atendimento de {inicio:%d/%m/%Y} a {fim:%d/%m/%Y} e "
                f"{'simular a finalização das' if simular else 'fazer Finalizar Parcial → Confirmar nas'} que tiverem "
                "data de assinatura (qualquer data)?\n\nO navegador vai abrir: faça o login nele.", parent=self):
            return
        self._rodar(f"▶ Por período - {inicio:%d/%m/%Y} a {fim:%d/%m/%Y} - {n_abas} aba(s)" + (" - SIMULAÇÃO" if simular else ""),
                    lambda log, prog: robo.executar_periodo(inicio, fim, limite, simular, log=log, parar=self.parar,
                                                            progresso=prog, n_abas=n_abas))

    def _rodar(self, titulo, funcao):
        self._salvar_prefs()
        self.parar.clear()
        self.resultado = None
        self._set_ocupado(True)
        self._status("Abrindo o portal... faça o login no navegador.")
        self._escrever(titulo, "ok")

        def trabalho():
            try:
                destino = funcao(lambda m: self.fila.put(("log", str(m))),
                                 lambda f, t: self.fila.put(("progresso", (f, t))))
                self.fila.put(("fim", destino))
            except PermissionError:
                self.fila.put(("erro", "Não consegui salvar o resultado. Feche a planilha aberta e tente de novo."))
            except Exception as e:  # noqa: BLE001
                self.fila.put(("erro", str(e).splitlines()[0] if str(e) else repr(e)))

        self.thread = threading.Thread(target=trabalho, daemon=True)
        self.thread.start()

    def _set_ocupado(self, ocupado):
        for b in (self.bt_planilha, self.bt_periodo):
            b.configure(state="disabled" if ocupado else "normal")
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
        self._status(f"Concluído: {Path(destino).name}")
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
                        self._status("Login feito. Trabalhando nas guias...")
                    elif "Faça o login" in valor:
                        self._status("Aguardando você fazer o login no navegador...")
                elif tipo == "progresso":
                    feitas, total = valor
                    self.barra.set(feitas / total if total else 0)
                    self._status(f"{feitas} de {total} guia(s)")
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


class App(_JanelaUnimed, ctk.CTk):
    """Robô aberto sozinho."""


class UnimedWindow(_JanelaUnimed, ctk.CTkToplevel):
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
    _janela_aberta = UnimedWindow(parent)
    return _janela_aberta


if __name__ == "__main__":
    App().mainloop()
