"""
Robô SulAmérica - conferência de guias no portal e geração das agendas.

Abra com dois cliques no "Abrir Robo SulAmerica.bat" (ou: python interface_robo.py).
Precisa estar na mesma pasta que robo_sulamerica.py e gerador_agenda.py.
"""

import json
import os
import queue
import re
import sys
import threading
import time
from pathlib import Path
from tkinter import filedialog, messagebox

import customtkinter as ctk

import conferencia_faturamento as fat
import gerador_agenda as agenda_gen
import robo_sulamerica as robo

ctk.set_appearance_mode("light")
ctk.set_default_color_theme("blue")

PASTA = Path(__file__).parent
if getattr(sys, "frozen", False):
    # dentro do CentralAutomacoes.exe a pasta do programa é temporária: guarda as preferências no usuário
    PASTA_DADOS = Path.home() / "CentralAutomacoes" / "SulAmerica"
    PASTA_DADOS.mkdir(parents=True, exist_ok=True)
else:
    PASTA_DADOS = PASTA
ARQ_PREFS = PASTA_DADOS / "preferencias.json"  # código e usuário de cada unidade e últimos arquivos (nunca a senha)
UNIDADES_PADRAO = ["1", "2", "3", "4"]

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


def _ordem(u):
    return (not u.isdigit(), int(u) if u.isdigit() else 0, u)


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


def _curto(p, n=60):
    p = str(p)
    return p if len(p) <= n else "…" + p[-(n - 1):]


class _JanelaRobo:
    """Janela do robô. Vira App (janela principal, pelo .bat) ou RoboWindow (aberta pela Central)."""

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.title("Robô SulAmérica")
        altura = max(640, min(1060, self.winfo_screenheight() - 90))
        self.geometry(f"900x{altura}+40+10")
        self.minsize(760, 600)
        self.configure(fg_color=BG)

        self.fila = queue.Queue()
        self.parar = threading.Event()
        self.thread = None
        self.resultado = None          # arquivo de resultado da conferência
        self.pasta_agendas = None      # pasta das agendas
        self.planilha = ""
        self.agenda = ""
        self.relatorio = ""            # relatório de faturamento do ClinicWeb
        self.pasta_pedidos = None      # pasta dos pedidos médicos
        self.resultado_fat = None      # planilha da conferência de faturamento
        self.resultados_salvos = []
        self._cont_planilha = {}
        self.prefs = {"unidades": {}}
        self._unidade_anterior = None

        self._montar_tela()
        self._carregar_prefs()
        self.protocol("WM_DELETE_WINDOW", self._fechar)
        self.after(150, self._ler_fila)

    def _ler_fila_se_aberta(self):
        try:
            if self.winfo_exists():
                self.after(150, self._ler_fila)
        except Exception:
            pass

    # ================================================================= tela
    def _fonte(self, tamanho=13, negrito=False):
        return ctk.CTkFont(size=tamanho, weight="bold" if negrito else "normal")

    def _cartao(self, pai, **pack):
        card = ctk.CTkFrame(pai, fg_color="white", corner_radius=16)
        card.pack(fill="x", padx=28, **pack)
        return card

    def _titulo_cartao(self, card, texto, linha=0, colunas=4):
        ctk.CTkLabel(card, text=texto, text_color=NAVY, font=self._fonte(14, True)).grid(
            row=linha, column=0, columnspan=colunas, sticky="w", padx=20, pady=(16, 6))

    def _botao_contorno(self, pai, texto, comando, **kw):
        return ctk.CTkButton(pai, text=texto, command=comando, fg_color="transparent", text_color=NAVY,
                             hover_color=CONTORNO_HOVER, border_width=1, border_color=NAVY,
                             text_color_disabled="#A9B3C4", **kw)

    def _botao_cheio(self, pai, texto, comando, **kw):
        return ctk.CTkButton(pai, text=texto, command=comando, fg_color=NAVY, hover_color=NAVY_HOVER, **kw)

    def _montar_tela(self):
        corpo = ctk.CTkScrollableFrame(self, fg_color=BG)
        corpo.pack(fill="both", expand=True)

        ctk.CTkLabel(corpo, text="Robô SulAmérica", text_color=NAVY,
                     font=self._fonte(26, True)).pack(pady=(20, 2))
        ctk.CTkLabel(corpo, text="Confere as guias no portal e monta as agendas do mês.",
                     text_color=CINZA_ESCURO, font=self._fonte(14)).pack(pady=(0, 16))

        # ---- 1. unidade e login
        c1 = self._cartao(corpo)
        for i in range(3):
            c1.grid_columnconfigure(i, weight=1, uniform="login")
        self._titulo_cartao(c1, "1. Unidade e login do portal", colunas=3)
        linha_und = ctk.CTkFrame(c1, fg_color="transparent")
        linha_und.grid(row=1, column=0, columnspan=3, sticky="w", padx=20, pady=(0, 4))
        ctk.CTkLabel(linha_und, text="Unidade (UND. AUT.)", text_color=TEXTO).pack(side="left")
        self.seg_unidade = ctk.CTkSegmentedButton(
            linha_und, values=UNIDADES_PADRAO, command=lambda v: self._trocar_unidade(),
            selected_color=NAVY, selected_hover_color=NAVY_HOVER, unselected_color=NAO_SELECIONADO,
            unselected_hover_color=NAO_SELECIONADO_HOVER, fg_color=NAO_SELECIONADO, text_color="white",
            font=self._fonte(13, True), width=180)
        self.seg_unidade.pack(side="left", padx=12)
        self.lbl_qtd_unidade = ctk.CTkLabel(linha_und, text="", text_color=CINZA)
        self.lbl_qtd_unidade.pack(side="left")

        for col, rotulo in enumerate(("Código referenciado", "Usuário", "Senha")):
            ctk.CTkLabel(c1, text=rotulo, text_color=TEXTO).grid(
                row=2, column=col, sticky="w", padx=(20 if col == 0 else 8, 8), pady=(8, 0))
        self.ent_codigo = ctk.CTkEntry(c1, placeholder_text="ex.: 100000021392", height=36)
        self.ent_usuario = ctk.CTkEntry(c1, placeholder_text="usuário", height=36)
        self.ent_senha = ctk.CTkEntry(c1, placeholder_text="senha", show="•", height=36)
        for col, ent in enumerate((self.ent_codigo, self.ent_usuario, self.ent_senha)):
            ent.grid(row=3, column=col, sticky="ew", padx=(20 if col == 0 else 8, 20 if col == 2 else 8))
        opc = ctk.CTkFrame(c1, fg_color="transparent")
        opc.grid(row=4, column=0, columnspan=3, sticky="ew", padx=20, pady=(12, 10))
        self.chk_lembrar = ctk.CTkCheckBox(opc, text="Lembrar código e usuário desta unidade (a senha nunca é salva)",
                                           fg_color=NAVY, hover_color=NAVY_HOVER, text_color=TEXTO)
        self.chk_lembrar.select()
        self.chk_lembrar.pack(side="left")
        self.sw_navegador = ctk.CTkSwitch(opc, text="Mostrar o navegador", progress_color=NAVY,
                                          text_color=TEXTO)
        self.sw_navegador.select()
        self.sw_navegador.pack(side="right")
        linha_abas = ctk.CTkFrame(c1, fg_color="transparent")
        linha_abas.grid(row=5, column=0, columnspan=3, sticky="ew", padx=20, pady=(0, 16))
        ctk.CTkLabel(linha_abas, text="Abas no portal ao mesmo tempo", text_color=TEXTO).pack(side="left")
        self.seg_abas = ctk.CTkSegmentedButton(
            linha_abas, values=["1", "2", "3", "4", "5"], selected_color=NAVY, selected_hover_color=NAVY_HOVER,
            unselected_color=NAO_SELECIONADO, unselected_hover_color=NAO_SELECIONADO_HOVER, fg_color=NAO_SELECIONADO,
            text_color="white", font=self._fonte(13, True), width=200)
        self.seg_abas.set("3")
        self.seg_abas.pack(side="left", padx=12)
        ctk.CTkLabel(linha_abas, text="mais abas = mais rápido (se o portal reclamar, use 1)",
                     text_color=CINZA).pack(side="left")

        # ---- 2. planilha de conferência
        c2 = self._cartao(corpo, pady=(12, 0))
        c2.grid_columnconfigure(0, weight=1)
        self._titulo_cartao(c2, "2. Planilha de conferência (abas SAS e SAS CONV)", colunas=2)
        self.lbl_planilha = ctk.CTkLabel(c2, text="Nenhum arquivo selecionado", text_color=CINZA, anchor="w")
        self.lbl_planilha.grid(row=1, column=0, sticky="ew", padx=20, pady=(0, 16))
        self._botao_cheio(c2, "Escolher arquivo", self._escolher_planilha, width=150).grid(
            row=1, column=1, sticky="e", padx=20, pady=(0, 16))

        # ---- 3. o que fazer (abas)
        self.abas = ctk.CTkTabview(corpo, fg_color="white", corner_radius=16, height=330,
                                   segmented_button_fg_color=NAO_SELECIONADO,
                                   segmented_button_selected_color=NAVY,
                                   segmented_button_selected_hover_color=NAVY_HOVER,
                                   segmented_button_unselected_color=NAO_SELECIONADO,
                                   segmented_button_unselected_hover_color=NAO_SELECIONADO_HOVER,
                                   text_color="white")
        self.abas._segmented_button.configure(font=self._fonte(13, True))
        self.abas.pack(fill="x", padx=28, pady=(12, 0))
        aba1 = self.abas.add("Conferir guias no portal")
        aba2 = self.abas.add("Gerar agendas")
        aba3 = self.abas.add("Conferir faturamento")
        self.abas.set("Gerar agendas")
        self._montar_aba_conferencia(aba1)
        self._montar_aba_agendas(aba2)
        self._montar_aba_faturamento(aba3)

        # ---- andamento e log
        andamento = ctk.CTkFrame(corpo, fg_color="transparent")
        andamento.pack(fill="x", padx=28, pady=(14, 0))
        self.lbl_status = ctk.CTkLabel(andamento, text="Pronto para começar.", text_color=CINZA_ESCURO,
                                       anchor="w")
        self.lbl_status.pack(fill="x")
        self.barra = ctk.CTkProgressBar(andamento, progress_color=NAVY, height=10)
        self.barra.set(0)
        self.barra.pack(fill="x", pady=(4, 0))

        self.log = ctk.CTkTextbox(corpo, height=170, fg_color="white", corner_radius=12, text_color=TEXTO,
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
        self.bt_imprimir = self._botao_contorno(rodape, "Abrir PDF para imprimir", self._abrir_pdf_impressao,
                                                state="disabled")
        self.bt_imprimir.pack(side="left", padx=5)
        self.bt_resumo = self._botao_contorno(rodape, "Abrir resumo", self._abrir_resumo, state="disabled")
        self.bt_resumo.pack(side="left", padx=5)
        self.bt_resultado = self._botao_contorno(rodape, "Abrir resultado",
                                                 self._abrir_resultado, state="disabled")
        self.bt_resultado.pack(side="left", padx=5)
        self.bt_pasta = self._botao_contorno(rodape, "Abrir pasta", self._abrir_pasta, state="disabled",
                                             width=100)
        self.bt_pasta.pack(side="left", padx=5)

    def _seletor_abas(self, pai):
        """Caixas 'ABA (SAS)' e 'Convencional (SAS CONV)'; as duas abas da janela usam as mesmas variáveis."""
        if not hasattr(self, "var_sas"):
            self.var_sas = ctk.BooleanVar(value=True)
            self.var_conv = ctk.BooleanVar(value=True)
        quadro = ctk.CTkFrame(pai, fg_color="#EEF2FA", corner_radius=10)
        ctk.CTkLabel(quadro, text="Conferir:", text_color=NAVY, font=self._fonte(13, True)).pack(
            side="left", padx=(12, 10), pady=8)
        for texto, var in (("ABA (aba SAS)", self.var_sas), ("Convencional (aba SAS CONV)", self.var_conv)):
            ctk.CTkCheckBox(quadro, text=texto, variable=var, fg_color=NAVY, hover_color=NAVY_HOVER,
                            text_color=TEXTO, command=self._mostrar_qtd_unidade).pack(side="left", padx=(0, 18))
        return quadro

    def _abas_escolhidas(self):
        abas = [a for a, v in ((robo.ABA, self.var_sas), (robo.ABA_CONV, self.var_conv)) if v.get()]
        if not abas:
            messagebox.showwarning("Conferir", "Marque ABA (SAS), Convencional (SAS CONV) ou os dois.", parent=self)
        return abas

    def _montar_aba_conferencia(self, aba):
        ctk.CTkLabel(aba, text="Entra no portal com o login da unidade, pesquisa a GUIA de cada linha e compara: "
                               "ABA com PSICO MÊS, FONO MÊS, TO MÊS e TOTAL MÊS; convencional com a QNTD da "
                               "especialidade.",
                     text_color=CINZA_ESCURO, wraplength=760, justify="left").pack(anchor="w", padx=12, pady=(8, 8))
        self._seletor_abas(aba).pack(anchor="w", padx=12, pady=(0, 12))
        linha = ctk.CTkFrame(aba, fg_color="transparent")
        linha.pack(anchor="w", padx=12)
        ctk.CTkLabel(linha, text="Testar só as primeiras", text_color=TEXTO).pack(side="left")
        self.ent_limite = ctk.CTkEntry(linha, width=60, justify="center")
        self.ent_limite.insert(0, "0")
        self.ent_limite.pack(side="left", padx=8)
        ctk.CTkLabel(linha, text="linhas (0 = todas)", text_color=CINZA).pack(side="left")
        self.sw_pedidos = ctk.CTkSwitch(aba, text="Salvar o pedido médico de cada paciente (anexos da guia)",
                                        progress_color=NAVY, text_color=TEXTO)
        self.sw_pedidos.select()
        self.sw_pedidos.pack(anchor="w", padx=12, pady=(14, 0))
        ctk.CTkLabel(aba, text="O resultado vai para a Área de Trabalho. Os pedidos vão para a pasta "
                               "\"Pedidos Médicos SulAmérica Unidade X\" na Área de Trabalho, um PDF por paciente com só o nome dele "
                               "(sem anexo: \"sem pedido anexado\" no resultado).",
                     text_color=CINZA, wraplength=760, justify="left").pack(anchor="w", padx=12, pady=(10, 0))
        self.bt_conferir = self._botao_cheio(aba, "Conferir no portal", self._iniciar_conferencia, height=46,
                                             corner_radius=12, font=self._fonte(16, True))
        self.bt_conferir.pack(fill="x", padx=12, pady=(20, 8), side="bottom")

    def _montar_aba_faturamento(self, aba):
        aba.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(aba, text="Liga a coluna FATURAMENTO ao nº guia do ClinicWeb e confere nome, senha (planilha e "
                               "portal), códigos, quantidades do mês e valor (ABA R$ 90,00 | convencional R$ 57,88).",
                     text_color=CINZA_ESCURO, wraplength=760, justify="left").grid(
            row=0, column=0, columnspan=3, sticky="w", padx=12, pady=(8, 8))
        self._seletor_abas(aba).grid(row=5, column=0, columnspan=3, sticky="w", padx=12, pady=(12, 0))
        ctk.CTkLabel(aba, text="Relatório ClinicWeb", text_color=NAVY, font=self._fonte(13, True)).grid(
            row=1, column=0, sticky="w", padx=12)
        self.lbl_relatorio = ctk.CTkLabel(aba, text="Nenhum arquivo selecionado", text_color=CINZA, anchor="w")
        self.lbl_relatorio.grid(row=1, column=1, sticky="ew", padx=8)
        self._botao_cheio(aba, "Escolher arquivo", self._escolher_relatorio, width=150).grid(
            row=1, column=2, sticky="e", padx=12)
        self.sw_portal_fat = ctk.CTkSwitch(aba, text="Conferir senha e códigos no portal (usa o login da unidade)",
                                           progress_color=NAVY, text_color=TEXTO)
        self.sw_portal_fat.select()
        self.sw_portal_fat.grid(row=2, column=0, columnspan=3, sticky="w", padx=12, pady=(14, 0))
        ctk.CTkLabel(aba, text="O resultado vai para a Área de Trabalho. Guias já conferidas na aba 1 (com a versão "
                               "nova do robô) não são pesquisadas de novo.", text_color=CINZA, wraplength=760,
                     justify="left").grid(row=3, column=0, columnspan=3, sticky="w", padx=12, pady=(12, 0))
        self.bt_faturamento = self._botao_cheio(aba, "Conferir faturamento", self._iniciar_faturamento, height=46,
                                                corner_radius=12, font=self._fonte(16, True))
        self.bt_faturamento.grid(row=6, column=0, columnspan=3, sticky="ew", padx=12, pady=(18, 8))

    def _montar_aba_agendas(self, aba):
        aba.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(aba, text="Agenda exportada", text_color=NAVY, font=self._fonte(13, True)).grid(
            row=0, column=0, sticky="w", padx=12, pady=(8, 0))
        self.lbl_agenda = ctk.CTkLabel(aba, text="Nenhum arquivo selecionado", text_color=CINZA, anchor="w")
        self.lbl_agenda.grid(row=0, column=1, sticky="ew", padx=8, pady=(8, 0))
        self._botao_cheio(aba, "Escolher arquivo", self._escolher_agenda, width=150).grid(
            row=0, column=2, sticky="e", padx=12, pady=(8, 0))

        ctk.CTkLabel(aba, text="Período", text_color=NAVY, font=self._fonte(13, True)).grid(
            row=1, column=0, sticky="w", padx=12, pady=(14, 0))
        per = ctk.CTkFrame(aba, fg_color="transparent")
        per.grid(row=1, column=1, columnspan=2, sticky="w", padx=8, pady=(14, 0))
        ctk.CTkLabel(per, text="de", text_color=TEXTO).pack(side="left")
        self.ent_inicio = ctk.CTkEntry(per, width=120, placeholder_text="dd/mm/aaaa", justify="center")
        self.ent_inicio.pack(side="left", padx=8)
        ctk.CTkLabel(per, text="até", text_color=TEXTO).pack(side="left")
        self.ent_fim = ctk.CTkEntry(per, width=120, placeholder_text="dd/mm/aaaa", justify="center")
        self.ent_fim.pack(side="left", padx=8)
        ctk.CTkLabel(per, text="(preenchido pelo nome do arquivo)", text_color=CINZA).pack(side="left", padx=4)

        ctk.CTkLabel(aba, text="Feriados", text_color=NAVY, font=self._fonte(13, True)).grid(
            row=2, column=0, sticky="w", padx=12, pady=(14, 0))
        self.ent_feriados = ctk.CTkEntry(aba, placeholder_text="ex.: 12/10/2026, 02/11/2026 (separe por vírgula)")
        self.ent_feriados.grid(row=2, column=1, columnspan=2, sticky="ew", padx=(8, 12), pady=(14, 0))

        ctk.CTkLabel(aba, text="Já conferidas", text_color=NAVY, font=self._fonte(13, True)).grid(
            row=3, column=0, sticky="w", padx=12, pady=(14, 0))
        self.lbl_salvos = ctk.CTkLabel(aba, text="nenhum resultado (opcional)", text_color=CINZA, anchor="w")
        self.lbl_salvos.grid(row=3, column=1, sticky="ew", padx=8, pady=(14, 0))
        bts = ctk.CTkFrame(aba, fg_color="transparent")
        bts.grid(row=3, column=2, sticky="e", padx=12, pady=(14, 0))
        self._botao_contorno(bts, "Adicionar", self._adicionar_resultados, width=80, height=28).pack(side="left")
        self._botao_contorno(bts, "Limpar", self._limpar_resultados, width=64, height=28).pack(side="left", padx=(6, 0))

        ctk.CTkLabel(aba, text="As agendas vão para uma pasta na Área de Trabalho, com um PDF de 2 por folha "
                               "para imprimir.", text_color=CINZA).grid(
            row=4, column=0, columnspan=3, sticky="w", padx=12, pady=(14, 0))
        self.bt_gerar = self._botao_cheio(aba, "Gerar agendas", self._iniciar_agendas, height=46, corner_radius=12,
                                          font=self._fonte(16, True))
        self.bt_gerar.grid(row=5, column=0, columnspan=3, sticky="ew", padx=12, pady=(14, 8))

    # ================================================================= campos
    @staticmethod
    def _set(entry, texto):
        entry.delete(0, "end")
        if texto:
            entry.insert(0, texto)

    def _mostrar_arquivo(self, rotulo, caminho):
        if caminho:
            rotulo.configure(text=_curto(Path(caminho).name, 70), text_color=NAVY)
        else:
            rotulo.configure(text="Nenhum arquivo selecionado", text_color=CINZA)

    def _mostrar_salvos(self):
        if not self.resultados_salvos:
            self.lbl_salvos.configure(text="nenhum resultado (opcional)", text_color=CINZA)
        else:
            nomes = ", ".join(Path(c).name for c in self.resultados_salvos)
            self.lbl_salvos.configure(text=_curto(f"{len(self.resultados_salvos)}: {nomes}", 70),
                                      text_color=NAVY)

    # ================================================================= preferências e unidades
    def _carregar_prefs(self):
        try:
            prefs = json.loads(ARQ_PREFS.read_text(encoding="utf-8"))
        except Exception:
            prefs = {}
        if "unidades" not in prefs:  # formato antigo (um login só) -> unidade 1
            antigo = {k: prefs.get(k, "") for k in ("codigo", "usuario") if prefs.get(k)}
            prefs = {"unidades": {"1": antigo} if antigo else {}, "planilha": prefs.get("planilha", "")}
        self.prefs = prefs
        if prefs.get("planilha") and Path(prefs["planilha"]).exists():
            self.planilha = prefs["planilha"]
        if prefs.get("relatorio") and Path(prefs["relatorio"]).exists():
            self.relatorio = prefs["relatorio"]
        if prefs.get("agenda") and Path(prefs["agenda"]).exists():
            self.agenda = prefs["agenda"]
            self._periodo_pelo_nome(self.agenda)
        self._set(self.ent_feriados, prefs.get("feriados", ""))
        if str(prefs.get("abas_portal", "")) in ("1", "2", "3", "4", "5"):
            self.seg_abas.set(str(prefs["abas_portal"]))
        self.var_sas.set(prefs.get("conferir_sas", True))
        self.var_conv.set(prefs.get("conferir_conv", True))
        self._mostrar_arquivo(self.lbl_planilha, self.planilha)
        self._mostrar_arquivo(self.lbl_agenda, self.agenda)
        self._mostrar_arquivo(self.lbl_relatorio, self.relatorio)
        self._contar_planilha()
        unidades = self._atualizar_unidades()
        self.seg_unidade.set(prefs.get("unidade") if prefs.get("unidade") in unidades else unidades[0])
        self._trocar_unidade(guardar_atual=False)

    def _guardar_login_da_unidade(self, und):
        if und and self.chk_lembrar.get():
            self.prefs.setdefault("unidades", {})[und] = {
                "codigo": self.ent_codigo.get().strip(),
                "usuario": self.ent_usuario.get().strip(),
            }

    def _trocar_unidade(self, guardar_atual=True):
        und = self.seg_unidade.get()
        if guardar_atual and self._unidade_anterior and self._unidade_anterior != und:
            self._guardar_login_da_unidade(self._unidade_anterior)
        login = self.prefs.get("unidades", {}).get(und, {})
        self._set(self.ent_codigo, login.get("codigo", ""))
        self._set(self.ent_usuario, login.get("usuario", ""))
        self._set(self.ent_senha, "")  # cada unidade tem a sua senha
        self._unidade_anterior = und
        self._mostrar_qtd_unidade()
        (self.ent_senha if self.ent_codigo.get() else self.ent_codigo).focus_set()

    def _contar_planilha(self):
        self._cont_planilha = {}   # unidade -> {aba: nº de linhas}
        if self.planilha and Path(self.planilha).exists():
            try:
                for r in robo.ler_registros(self.planilha):
                    if r["unidade"]:
                        d = self._cont_planilha.setdefault(r["unidade"], {})
                        d[r["aba"]] = d.get(r["aba"], 0) + 1
            except Exception as e:
                messagebox.showwarning("Planilha", f"Não consegui ler a planilha de conferência: {e}", parent=self)

    def _atualizar_unidades(self):
        # linhas com UND. AUT. em branco ou inválida não viram uma "unidade" na lista
        unidades = sorted((u for u in self._cont_planilha if str(u).isdigit()), key=_ordem) or UNIDADES_PADRAO
        atual = self.seg_unidade.get()
        self.seg_unidade.configure(values=unidades)
        if atual not in unidades:
            self.seg_unidade.set(unidades[0])
            self._trocar_unidade(guardar_atual=False)
        else:
            self.seg_unidade.set(atual)
        self._mostrar_qtd_unidade()
        return unidades

    def _mostrar_qtd_unidade(self):
        d = self._cont_planilha.get(self.seg_unidade.get())
        if not d:
            texto = "escolha a planilha de conferência" if not self._cont_planilha else "nenhuma linha nesta unidade"
        else:
            texto = " | ".join(f"{d.get(a, 0)} na {a}" for a in (robo.ABA, robo.ABA_CONV))
        sem = sum(n for u, dd in self._cont_planilha.items() if not str(u).isdigit() for n in dd.values())
        if sem:
            texto += f"   ⚠ {sem} linha(s) sem unidade na planilha"
        self.lbl_qtd_unidade.configure(text=texto)

    def _salvar_prefs(self):
        und = self.seg_unidade.get()
        self._guardar_login_da_unidade(und)
        if not self.chk_lembrar.get():
            self.prefs.get("unidades", {}).pop(und, None)
        self.prefs.update({"unidade": und, "planilha": self.planilha, "agenda": self.agenda,
                           "relatorio": self.relatorio, "abas_portal": self.seg_abas.get(),
                           "conferir_sas": bool(self.var_sas.get()), "conferir_conv": bool(self.var_conv.get()),
                           "feriados": self.ent_feriados.get().strip()})
        try:
            ARQ_PREFS.write_text(json.dumps(self.prefs, ensure_ascii=False, indent=2), encoding="utf-8")
        except Exception:
            pass

    # ================================================================= escolher arquivos
    def _pasta_inicial(self, caminho):
        return str(Path(caminho).parent) if caminho else str(Path.home() / "Downloads")

    def _escolher_planilha(self):
        caminho = filedialog.askopenfilename(parent=self, title="Planilha de conferência",
                                             initialdir=self._pasta_inicial(self.planilha),
                                             filetypes=[("Planilha Excel", "*.xlsx *.xlsm")])
        if caminho:
            self.planilha = caminho
            self._mostrar_arquivo(self.lbl_planilha, caminho)
            self._contar_planilha()
            self._atualizar_unidades()

    def _escolher_relatorio(self):
        caminho = filedialog.askopenfilename(parent=self, title="Relatório de Conferência de Faturamento (ClinicWeb)",
                                             initialdir=self._pasta_inicial(self.relatorio or self.planilha),
                                             filetypes=[("Planilha Excel", "*.xlsx")])
        if caminho:
            self.relatorio = caminho
            self._mostrar_arquivo(self.lbl_relatorio, caminho)

    def _escolher_agenda(self):
        caminho = filedialog.askopenfilename(parent=self, title="Agenda exportada",
                                             initialdir=self._pasta_inicial(self.agenda),
                                             filetypes=[("Planilha Excel", "*.xlsx")])
        if caminho:
            self.agenda = caminho
            self._mostrar_arquivo(self.lbl_agenda, caminho)
            self._periodo_pelo_nome(caminho)

    def _periodo_pelo_nome(self, caminho):
        # "agendamentos-2026-10-05-a-2026-10-31.xlsx" -> preenche início e fim
        m = re.search(r"(\d{4})-(\d{2})-(\d{2})-a-(\d{4})-(\d{2})-(\d{2})", Path(caminho).name)
        if m:
            a1, m1, d1, a2, m2, d2 = m.groups()
            self._set(self.ent_inicio, f"{d1}/{m1}/{a1}")
            self._set(self.ent_fim, f"{d2}/{m2}/{a2}")

    def _adicionar_resultados(self, caminhos=None):
        if caminhos is None:
            caminhos = filedialog.askopenfilenames(parent=self, title="Resultados da conferência",
                                                   initialdir=str(agenda_gen.area_de_trabalho()),
                                                   filetypes=[("Resultado do robô", "Resultado SulAmerica*.xlsx"),
                                                              ("Planilha Excel", "*.xlsx")])
        for c in caminhos:
            if c and str(c) not in self.resultados_salvos:
                self.resultados_salvos.append(str(c))
        self._mostrar_salvos()

    def _limpar_resultados(self):
        self.resultados_salvos = []
        self._mostrar_salvos()

    # ================================================================= validação
    def _validar_login(self, pergunta):
        codigo = self.ent_codigo.get().strip()
        usuario = self.ent_usuario.get().strip()
        senha = self.ent_senha.get()
        faltando = [n for n, v in (("Código referenciado", codigo), ("Usuário", usuario), ("Senha", senha)) if not v]
        if faltando:
            messagebox.showwarning("Faltam dados", "Preencha no login: " + ", ".join(faltando), parent=self)
            return None
        avisos = []
        if len(codigo) > 12:
            avisos.append("o código referenciado tem mais de 12 caracteres")
        if len(usuario) > 8:
            avisos.append("o usuário tem mais de 8 caracteres")
        if len(senha) > 8:
            avisos.append("a senha tem mais de 8 caracteres")
        if avisos and not messagebox.askyesno(
                "Confira os dados", "O site aceita no máximo 12 dígitos no código e 8 caracteres no usuário e na "
                "senha, mas " + "; ".join(avisos) + ".\n\nContinuar mesmo assim?", parent=self):
            return None
        und = self.seg_unidade.get()
        if not messagebox.askyesno("Confirmar", pergunta.format(und=und, usuario=usuario, codigo=codigo),
                                   parent=self):
            return None
        return {"CODIGO_REFERENCIADO": codigo, "USUARIO": usuario, "SENHA": senha}

    def _planilha_ok(self):
        if not self.planilha or not Path(self.planilha).exists():
            messagebox.showwarning("Planilha", "Escolha a planilha de conferência (aba SAS).", parent=self)
            return False
        return True

    # ================================================================= execução
    def _set_ocupado(self, ocupado):
        estado = "disabled" if ocupado else "normal"
        self.bt_conferir.configure(state=estado)
        self.bt_gerar.configure(state=estado)
        self.bt_faturamento.configure(state=estado)
        self.seg_unidade.configure(state=estado)
        self.bt_parar.configure(state="normal" if ocupado else "disabled")
        if ocupado:
            self.barra.set(0)

    def _n_abas(self):
        try:
            return int(self.seg_abas.get())
        except (TypeError, ValueError):
            return 1

    def _iniciar_conferencia(self):
        if not self._planilha_ok():
            return
        abas = self._abas_escolhidas()
        if not abas:
            return
        quais = " e ".join("ABA" if a == robo.ABA else "convencional" for a in abas)
        cfg = self._validar_login("Conferir as guias (" + quais + ") da UNIDADE {und} com o login {usuario} "
                                  "(código {codigo})?")
        if not cfg:
            return
        self._salvar_prefs()
        try:
            limite = max(0, int(self.ent_limite.get() or 0))
        except ValueError:
            limite = 0
        unidade, planilha = self.seg_unidade.get(), self.planilha
        headless = not self.sw_navegador.get()
        pasta_pedidos = (agenda_gen.area_de_trabalho() / f"Pedidos Médicos SulAmérica Unidade {unidade}"
                         if self.sw_pedidos.get() else None)
        self.pasta_pedidos = pasta_pedidos

        self.parar.clear()
        self.resultado = None
        self._set_ocupado(True)
        self._status(f"Unidade {unidade}: conferindo guias no portal...")
        n_abas = self._n_abas()
        self._escrever(f"▶ Conferência - unidade {unidade} - {', '.join(abas)} - {n_abas} aba(s) no portal", "ok")

        def trabalho():
            try:
                destino = robo.executar(
                    self._cfg_copia(cfg), planilha, limite=limite, headless=headless, unidade=unidade,
                    pasta_pedidos=pasta_pedidos, abas=tuple(abas), n_abas=n_abas, parar=self.parar,
                    pasta_saida=agenda_gen.area_de_trabalho(),
                    log=lambda m: self.fila.put(("log", str(m))),
                    progresso=lambda f, t: self.fila.put(("progresso", (f, t, "guia(s) pesquisada(s)"))))
                self.fila.put(("fim_conferencia", destino))
            except Exception as e:
                self.fila.put(("erro", str(e)))

        self.thread = threading.Thread(target=trabalho, daemon=True)
        self.thread.start()

    @staticmethod
    def _cfg_copia(cfg):
        return dict(cfg)

    def _iniciar_agendas(self):
        if not self._planilha_ok():
            return
        if not self.agenda or not Path(self.agenda).exists():
            messagebox.showwarning("Agenda", "Escolha o arquivo da agenda exportada.", parent=self)
            return
        inicio = agenda_gen.para_data(self.ent_inicio.get().strip())
        fim = agenda_gen.para_data(self.ent_fim.get().strip())
        if not inicio or not fim:
            messagebox.showwarning("Período", "Preencha o período no formato dd/mm/aaaa.", parent=self)
            return
        try:
            feriados = agenda_gen.ler_feriados(self.ent_feriados.get(), inicio.year)
        except ValueError as e:
            messagebox.showwarning("Feriados", str(e), parent=self)
            return
        fora = [d for d in feriados if not (inicio <= d <= fim)]
        if fora and not messagebox.askyesno(
                "Feriados", "Estes feriados estão fora do período e não vão mudar nada: "
                + ", ".join(f"{d:%d/%m/%Y}" for d in sorted(fora)) + "\n\nContinuar?", parent=self):
            return
        cfg = self._validar_login("Entrar no portal com o login da UNIDADE {und} ({usuario}), pesquisar as guias "
                                  "dos pacientes com UND. AUT. {und} e gerar as agendas deles?")
        if not cfg:
            return
        self._salvar_prefs()
        unidade, planilha, agenda = self.seg_unidade.get(), self.planilha, self.agenda
        resultados = list(self.resultados_salvos)
        headless = not self.sw_navegador.get()
        n_abas = self._n_abas()

        self.parar.clear()
        self._set_ocupado(True)
        self._status(f"Unidade {unidade}: lendo a planilha e a agenda...")
        self._escrever(f"▶ Agendas - unidade {unidade} - {inicio:%d/%m/%Y} a {fim:%d/%m/%Y}", "ok")
        if feriados:
            self._escrever("Feriados: " + ", ".join(f"{d:%d/%m}" for d in sorted(feriados)))
        log = lambda m: self.fila.put(("log", str(m)))  # noqa: E731

        def consultar_portal(guias):
            self.fila.put(("status", f"Unidade {unidade}: pesquisando {len(guias)} guia(s) no portal..."))
            consultas = robo.consultar_lista(
                cfg, guias, headless=headless, log=log, parar=self.parar, n_abas=n_abas,
                progresso=lambda f, t: self.fila.put(("progresso", (f, t, "guia(s) pesquisada(s)"))))
            if self.parar.is_set():
                raise RuntimeError("Interrompido. Nenhuma agenda foi gerada.")
            self.fila.put(("status", f"Unidade {unidade}: gerando os PDFs..."))
            return consultas

        def trabalho():
            try:
                pacientes = [{"carteirinha": r["carteirinha"], "paciente": r["paciente"],
                              "guias": r["guias"], "planilha": r["planilha"]}
                             for r in robo.ler_planilha(planilha, unidade)]
                pasta, resumo = agenda_gen.gerar(agenda, resultados, inicio, fim, feriados, log=log,
                                                 unidade=unidade, consultar_portal=consultar_portal,
                                                 pacientes_unidade=pacientes)
                self.fila.put(("fim_agendas", (pasta, resumo)))
            except PermissionError:
                self.fila.put(("erro", "Não consegui salvar. Feche o PDF ou o resumo que estiver aberto e "
                                       "tente de novo."))
            except Exception as e:
                self.fila.put(("erro", str(e)))

        self.thread = threading.Thread(target=trabalho, daemon=True)
        self.thread.start()

    def _iniciar_faturamento(self):
        if not self._planilha_ok():
            return
        abas = self._abas_escolhidas()
        if not abas:
            return
        if not self.relatorio or not Path(self.relatorio).exists():
            messagebox.showwarning("Relatório", "Escolha o Relatório de Conferência de Faturamento do ClinicWeb.",
                                   parent=self)
            return
        usar_portal = bool(self.sw_portal_fat.get())
        cfg = None
        if usar_portal:
            cfg = self._validar_login("Conferir o faturamento da UNIDADE {und}, pesquisando as guias no portal com o "
                                      "login {usuario} (código {codigo})?")
            if not cfg:
                return
        elif not messagebox.askyesno("Confirmar", f"Conferir o faturamento da UNIDADE {self.seg_unidade.get()} "
                                     "sem consultar o portal (senha e códigos do portal não serão conferidos)?",
                                     parent=self):
            return
        self._salvar_prefs()
        unidade, planilha, relatorio = self.seg_unidade.get(), self.planilha, self.relatorio
        resultados = list(self.resultados_salvos)
        headless = not self.sw_navegador.get()
        n_abas = self._n_abas()

        self.parar.clear()
        self.resultado_fat = None
        self._set_ocupado(True)
        self._status(f"Unidade {unidade}: lendo o relatório do ClinicWeb...")
        self._escrever(f"▶ Faturamento - unidade {unidade} - {', '.join(abas)} - {Path(relatorio).name}", "ok")
        log = lambda m: self.fila.put(("log", str(m)))  # noqa: E731

        def consultar_portal(guias):
            self.fila.put(("status", f"Unidade {unidade}: pesquisando {len(guias)} guia(s) no portal..."))
            consultas = robo.consultar_lista(
                cfg, guias, headless=headless, log=log, parar=self.parar, n_abas=n_abas,
                progresso=lambda f, t: self.fila.put(("progresso", (f, t, "guia(s) pesquisada(s)"))))
            if self.parar.is_set():
                raise RuntimeError("Interrompido. A conferência de faturamento não foi gravada.")
            return consultas

        def trabalho():
            try:
                registros = robo.ler_registros(planilha, unidade, tuple(abas), ignorar_ausentes=len(abas) > 1)
                linhas = fat.conferir(relatorio, registros, consultar_portal=consultar_portal if usar_portal else None,
                                      resultados_salvos=resultados, log=log)
                destino = agenda_gen.area_de_trabalho() / (
                    f"Conferencia Faturamento SulAmerica Unidade {unidade} {time.strftime('%Y-%m-%d %H%M')}.xlsx")
                fat.gravar(linhas, destino)
                self.fila.put(("fim_faturamento", (destino, linhas)))
            except PermissionError:
                self.fila.put(("erro", "Não consegui salvar. Feche a planilha que estiver aberta e tente de novo."))
            except Exception as e:
                self.fila.put(("erro", str(e)))

        self.thread = threading.Thread(target=trabalho, daemon=True)
        self.thread.start()

    def _fim_faturamento(self, destino, linhas):
        self.resultado = destino
        self._set_ocupado(False)
        self.barra.set(1)
        self.bt_resultado.configure(state="normal")
        self.bt_pasta.configure(state="normal")
        cont = {}
        for r in linhas:
            cont[r["Situação"]] = cont.get(r["Situação"], 0) + 1
        resumo = (f"{cont.get('OK', 0)} OK | {cont.get('DIVERGENTE', 0)} divergente(s) | "
                  f"{cont.get('SEM FATURAMENTO', 0)} sem faturamento | {cont.get('AVISO', 0)} com aviso")
        self._status(f"Faturamento conferido: {resumo}")
        div = [r for r in linhas if r["Situação"] == "DIVERGENTE"]
        texto = f"{resumo}\n\n"
        if div:
            texto += "\n".join(f"• {r['Paciente']}" for r in div[:15])
            if len(div) > 15:
                texto += f"\n... e mais {len(div) - 15}"
            texto += "\n\n"
        texto += f"Planilha salva na Área de Trabalho:\n{Path(destino).name}\n\nAbrir agora?"
        if messagebox.askyesno("Conferência de faturamento", texto, parent=self):
            abrir_arquivo(destino)

    def _parar(self):
        self.parar.set()
        self.bt_parar.configure(state="disabled")
        self._status("Parando...")

    def _fim_conferencia(self, destino):
        self.resultado = destino
        self._adicionar_resultados([str(destino)])  # já serve para as agendas
        self._set_ocupado(False)
        self.barra.set(1)
        self.bt_resultado.configure(state="normal")
        self.bt_pasta.configure(state="normal")
        self._status(f"Conferência concluída: {Path(destino).name}")
        extra = ""
        if self.pasta_pedidos and Path(self.pasta_pedidos).exists():
            extra = f"\n\nPedidos médicos salvos em:\n{self.pasta_pedidos}"
            abrir_arquivo(self.pasta_pedidos)
        if messagebox.askyesno("Concluído", f"Conferência finalizada.{extra}\n\nResultado salvo na Área de "
                               f"Trabalho:\n{Path(destino).name}\n\nAbrir a planilha de resultado?",
                               parent=self):
            abrir_arquivo(destino)

    def _fim_agendas(self, pasta, resumo):
        self.pasta_agendas = pasta
        self._set_ocupado(False)
        self.barra.set(1)
        for b in (self.bt_imprimir, self.bt_resumo, self.bt_pasta):
            b.configure(state="normal")
        problemas = [r for r in resumo if r["Situação"] != "OK"]
        gerados = sum(1 for r in resumo if r.get("PDF"))
        self._status(f"{gerados} agenda(s) gerada(s) | {len(problemas)} com divergência ou aviso | "
                     f"pasta na Área de Trabalho")
        if problemas:
            linhas = [f"• {r['Paciente']}: {r['Situação']}" for r in problemas[:18]]
            if len(problemas) > 18:
                linhas.append(f"... e mais {len(problemas) - 18}")
            messagebox.showwarning(
                "Divergências nas agendas",
                f"{gerados} agenda(s) gerada(s). {len(problemas)} paciente(s) com divergência ou aviso:\n\n"
                + "\n".join(linhas) + "\n\nOs detalhes estão no log e no 'Resumo das agendas.xlsx'.", parent=self)
        else:
            messagebox.showinfo("Agendas", f"{gerados} agenda(s) gerada(s), todas batendo com o autorizado.",
                                parent=self)

    # ================================================================= botões do rodapé
    def _abrir_resultado(self):
        if self.resultado and Path(self.resultado).exists():
            abrir_arquivo(self.resultado)
        elif self.resultado:
            messagebox.showwarning("Resultado", f"Não achei o arquivo:\n{self.resultado}\n\nEle pode ter sido "
                                   "movido ou apagado.", parent=self)

    def _abrir_pasta(self):
        if self.resultado and Path(self.resultado).exists() and (
                not self.pasta_agendas or Path(self.resultado).stat().st_mtime > Path(self.pasta_agendas).stat().st_mtime):
            abrir_arquivo(Path(self.resultado).parent)
        elif self.pasta_agendas and Path(self.pasta_agendas).exists():
            abrir_arquivo(self.pasta_agendas)
        elif self.resultado:
            abrir_arquivo(Path(self.resultado).parent)

    def _abrir_pdf_impressao(self):
        if self.pasta_agendas:
            for arq in Path(self.pasta_agendas).glob("*2 por folha.pdf"):
                abrir_arquivo(arq)
                return

    def _abrir_resumo(self):
        if self.pasta_agendas:
            arq = Path(self.pasta_agendas) / "Resumo das agendas.xlsx"
            if arq.exists():
                abrir_arquivo(arq)

    def _fechar(self):
        if self.thread and self.thread.is_alive():
            if not messagebox.askyesno("Robô rodando", "O robô ainda está rodando. Fechar mesmo assim?",
                                       parent=self):
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
            if t.startswith(("ERRO", "✖")) or "falhou" in t or "quantidades batem: NÃO" in t:
                tag = "erro"
            elif t.startswith("AVISO"):
                tag = "aviso"
            elif t.startswith(("OK", "Pronto", "Login OK", "PDF para impressão")):
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
                elif tipo == "status":
                    self._status(valor)
                elif tipo == "progresso":
                    feitas, total, rotulo = valor
                    self.barra.set(feitas / total if total else 0)
                    self._status(f"{feitas} de {total} {rotulo}")
                elif tipo == "fim_conferencia":
                    self._fim_conferencia(valor)
                elif tipo == "fim_agendas":
                    self._fim_agendas(*valor)
                elif tipo == "fim_faturamento":
                    self._fim_faturamento(*valor)
                elif tipo == "erro":
                    self._set_ocupado(False)
                    self._escrever(f"✖ {valor}", "erro")
                    self._status("Parou com erro (veja abaixo).")
                    messagebox.showerror("Erro", valor, parent=self)
        except queue.Empty:
            pass
        self._ler_fila_se_aberta()


class App(_JanelaRobo, ctk.CTk):
    """Robô aberto sozinho (Abrir Robo SulAmerica.bat)."""


class RoboWindow(_JanelaRobo, ctk.CTkToplevel):
    """Robô aberto como janela da Central de Automações."""

    def __init__(self, parent):
        super().__init__(parent)
        self.after(250, self._trazer_para_frente)

    def _trazer_para_frente(self):
        self.lift()
        self.focus_force()


_janela_aberta = None


def abrir(parent):
    """Abre o robô pela Central; se já estiver aberto, só traz a janela para a frente."""
    global _janela_aberta
    if _janela_aberta is not None and _janela_aberta.winfo_exists():
        _janela_aberta.deiconify()
        _janela_aberta._trazer_para_frente()
        return _janela_aberta
    _janela_aberta = RoboWindow(parent)
    return _janela_aberta


if __name__ == "__main__":
    App().mainloop()
