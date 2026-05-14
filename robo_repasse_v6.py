import tkinter as tk
from tkinter import ttk, filedialog, scrolledtext, messagebox
import customtkinter as ctk
import requests
import getpass
from datetime import datetime
import pandas as pd
import os
import sys
import threading
import time
from pathlib import Path

# ====================== SELENIUM ======================
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import (
    TimeoutException, StaleElementReferenceException, ElementClickInterceptedException,
    ElementNotInteractableException, WebDriverException, NoSuchElementException
)

# ====================== CONFIGURAÇÕES GLOBAIS ======================
URL_VALIDACAO = "https://raw.githubusercontent.com/diogodiasyt-blip/validacaofoco/refs/heads/main/chave"
URL_CORAL_LOGIN = "https://coral.aluguefoco.com.br/login"
URL_CORAL_CONTRATOS = "https://coral.aluguefoco.com.br/contratos"
XPATH_ABA_CONTRATOS = "/html/body/foco-app/div[1]/foco-rent-agreement-home/div/ngb-tabset/ul/li[3]/a"
XPATH_CAMPO_BUSCA_CONTRATOS = "/html/body/foco-app/div[1]/foco-rent-agreement-home/div/div/div[2]/input"

COLUNAS = {
    "status": 3,
    "contrato": 4,
    "valor_repasse": 10
}

TIMEOUT_PADRAO = 45
TIMEOUT_ACAO = 20
TIMEOUT_PAGAMENTOS = 60
PAUSA_CURTA = 1.0
PAUSA_MEDIA = 2.0
MAX_TENTATIVAS_POR_CONTRATO = 2

MAIN_BG = "#f6f4f1"
CARD_BG = "#ffffff"
CARD_BORDER = "#eadfdb"
PRIMARY_TEXT = "#d81919"
MUTED_TEXT = "#5c5c5c"
BUTTON_BG = "#ef1a14"
BUTTON_ACTIVE_BG = "#c91410"
SUCCESS_TEXT = "#187a2f"
SOFT_RED = "#fff1ef"


def localizar_logo():
    candidatos = []
    if getattr(sys, "_MEIPASS", None):
        candidatos.append(os.path.join(sys._MEIPASS, "assets", "logo.png"))
    base_atual = os.path.dirname(os.path.abspath(__file__))
    candidatos.append(os.path.join(os.path.dirname(base_atual), "assets", "logo.png"))
    candidatos.append(os.path.join(os.getcwd(), "DESENVOLVIMENTO", "assets", "logo.png"))
    candidatos.append(os.path.join(os.getcwd(), "assets", "logo.png"))
    for caminho in candidatos:
        if os.path.exists(caminho):
            return caminho
    return None


class AppRepasse:
    def __init__(self, root):
        ctk.set_appearance_mode("light")
        ctk.set_default_color_theme("blue")
        self.root = root
        self.root.title("🤖 Robô de Repasse de Recebimentos v6 - Criado por Diogo Medeiros")
        self.root.geometry("1120x820")
        self.root.minsize(980, 720)
        self.root.resizable(True, True)
        self.root.configure(bg=MAIN_BG)

        self.driver = None
        self.wait = None
        self.campo_busca = None
        self.processando = False

        self.caminho_planilha = tk.StringVar()
        self.usuario = tk.StringVar()
        self.senha = tk.StringVar()
        self.aba_selecionada = tk.StringVar()
        self.data_repasse_var = tk.StringVar()
        self.modo_invisivel = tk.BooleanVar(value=False)

        self.progresso = tk.DoubleVar(value=0)
        self.df = None
        self.total_aptos = 0
        self.lancados = 0
        self.ignorados = 0
        self.erros = 0
        self.atual = 0

        self.contratos_lancados = []
        self.contratos_com_erro = []
        self.contratos_ignorados = []
        self.caminho_relatorio_final = None
        self.logo_image = None

        self.indice_inicio = 0
        self.caminho_relatorio_parcial = None
        self.tentativas_do_contrato_atual = 0

        self.configurar_estilo()
        self.criar_interface()
        self.validar_abertura()

    def configurar_estilo(self):
        style = ttk.Style()
        try:
            style.theme_use("clam")
        except Exception:
            pass
        style.configure("Modern.TCombobox", fieldbackground="#ffffff", background="#ffffff", foreground="#303030", padding=8)

    def carregar_logo(self, reducao=2):
        caminho_logo = localizar_logo()
        if not caminho_logo:
            return None
        try:
            logo = tk.PhotoImage(file=caminho_logo)
            if reducao > 1:
                logo = logo.subsample(reducao, reducao)
            self.logo_image = logo
            return logo
        except Exception:
            return None

    def criar_secao(self, parent, titulo):
        frame = ctk.CTkFrame(
            parent,
            fg_color=CARD_BG,
            corner_radius=22,
            border_width=1,
            border_color=CARD_BORDER,
        )
        frame.pack(fill="x", padx=8, pady=(0, 14))
        ctk.CTkLabel(
            frame,
            text=titulo,
            text_color=PRIMARY_TEXT,
            font=("Segoe UI", 18, "bold"),
        ).pack(anchor="w", padx=18, pady=(16, 12))
        return frame

    # ====================== INTERFACE GR?FICA ======================
    def criar_interface(self):
        container = ctk.CTkFrame(self.root, fg_color=MAIN_BG, corner_radius=0)
        container.pack(fill="both", expand=True, padx=12, pady=12)

        scroll = ctk.CTkScrollableFrame(container, fg_color=MAIN_BG, corner_radius=0)
        scroll.pack(fill="both", expand=True)

        hero = ctk.CTkFrame(scroll, fg_color=CARD_BG, corner_radius=26, border_width=1, border_color=CARD_BORDER)
        hero.pack(fill="x", padx=8, pady=(8, 14))

        hero_inner = ctk.CTkFrame(hero, fg_color="transparent")
        hero_inner.pack(fill="x", padx=24, pady=24)

        logo = self.carregar_logo(reducao=2)
        if logo:
            ctk.CTkLabel(hero_inner, text="", image=logo).pack(side="left", padx=(0, 18))

        header_texto = ctk.CTkFrame(hero_inner, fg_color="transparent")
        header_texto.pack(side="left", fill="x", expand=True)
        ctk.CTkLabel(header_texto, text="Repasse FOCO", text_color=PRIMARY_TEXT, font=("Segoe UI", 30, "bold")).pack(anchor="w")
        ctk.CTkLabel(
            header_texto,
            text="Processamento de repasses com validacao, progresso em tempo real e relatorio final.",
            text_color=MUTED_TEXT,
            font=("Segoe UI", 14),
        ).pack(anchor="w", pady=(6, 0))
        ctk.CTkLabel(header_texto, text="OPERACAO DE REPASSE", text_color="#a65f56", font=("Segoe UI", 12, "bold")).pack(anchor="w", pady=(10, 0))

        acesso = self.criar_secao(scroll, "Acesso ao Coral")
        acesso_grid = ctk.CTkFrame(acesso, fg_color="transparent")
        acesso_grid.pack(fill="x", padx=18, pady=(0, 18))
        acesso_grid.grid_columnconfigure((0, 1), weight=1)
        ctk.CTkLabel(acesso_grid, text="Usuario", font=("Segoe UI", 13, "bold"), text_color="#303030").grid(row=0, column=0, sticky="w", padx=(0, 10), pady=(0, 6))
        ctk.CTkLabel(acesso_grid, text="Senha", font=("Segoe UI", 13, "bold"), text_color="#303030").grid(row=0, column=1, sticky="w", padx=(10, 0), pady=(0, 6))
        self.entry_usuario = ctk.CTkEntry(acesso_grid, textvariable=self.usuario, height=42, corner_radius=12)
        self.entry_usuario.grid(row=1, column=0, sticky="ew", padx=(0, 10))
        self.entry_senha = ctk.CTkEntry(acesso_grid, textvariable=self.senha, show="*", height=42, corner_radius=12)
        self.entry_senha.grid(row=1, column=1, sticky="ew", padx=(10, 0))

        planilha = self.criar_secao(scroll, "Planilha de Repasses")
        planilha_form = ctk.CTkFrame(planilha, fg_color="transparent")
        planilha_form.pack(fill="x", padx=18, pady=(0, 18))
        planilha_form.grid_columnconfigure(0, weight=1)
        self.entry_planilha = ctk.CTkEntry(planilha_form, textvariable=self.caminho_planilha, height=44, corner_radius=12)
        self.entry_planilha.grid(row=0, column=0, sticky="ew", padx=(0, 10), pady=(0, 10))
        ctk.CTkButton(
            planilha_form,
            text="Selecionar",
            command=self.selecionar_planilha,
            height=44,
            width=150,
            corner_radius=14,
            fg_color="#ffffff",
            text_color=PRIMARY_TEXT,
            hover_color=SOFT_RED,
            border_width=1,
            border_color="#f0d7d2",
            font=("Segoe UI", 14, "bold"),
        ).grid(row=0, column=1, sticky="e", pady=(0, 10))
        ctk.CTkLabel(planilha_form, text="Aba da planilha", font=("Segoe UI", 13, "bold"), text_color="#303030").grid(row=1, column=0, sticky="w", pady=(0, 6))
        self.combo_abas = ttk.Combobox(planilha_form, textvariable=self.aba_selecionada, state="readonly", style="Modern.TCombobox")
        self.combo_abas.grid(row=2, column=0, columnspan=2, sticky="ew")

        indicadores = self.criar_secao(scroll, "Indicadores")
        indicadores_grid = ctk.CTkFrame(indicadores, fg_color="transparent")
        indicadores_grid.pack(fill="x", padx=18, pady=(0, 18))
        indicadores_grid.grid_columnconfigure((0, 1, 2, 3), weight=1)
        self.label_contador = ctk.CTkLabel(indicadores_grid, text="Aptos: 0", font=("Segoe UI", 13, "bold"), text_color="#303030")
        self.label_contador.grid(row=0, column=0, sticky="w", padx=(0, 12))
        self.label_lancados = ctk.CTkLabel(indicadores_grid, text="Lancados: 0", font=("Segoe UI", 13, "bold"), text_color=SUCCESS_TEXT)
        self.label_lancados.grid(row=0, column=1, sticky="w", padx=12)
        self.label_ignorados = ctk.CTkLabel(indicadores_grid, text="Ignorados: 0", font=("Segoe UI", 13, "bold"), text_color="#9a6c2f")
        self.label_ignorados.grid(row=0, column=2, sticky="w", padx=12)
        self.label_erros = ctk.CTkLabel(indicadores_grid, text="Erros: 0", font=("Segoe UI", 13, "bold"), text_color="#bf2121")
        self.label_erros.grid(row=0, column=3, sticky="w", padx=(12, 0))

        configuracoes = self.criar_secao(scroll, "Configuracoes")
        config_grid = ctk.CTkFrame(configuracoes, fg_color="transparent")
        config_grid.pack(fill="x", padx=18, pady=(0, 18))
        config_grid.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(config_grid, text="Data do repasse", font=("Segoe UI", 13, "bold"), text_color="#303030").grid(row=0, column=0, sticky="w", pady=(0, 6))
        self.entry_data = ctk.CTkEntry(config_grid, textvariable=self.data_repasse_var, height=42, corner_radius=12, width=180)
        self.entry_data.grid(row=1, column=0, sticky="w", padx=(0, 10))
        self.entry_data.insert(0, datetime.now().strftime("%d/%m/%Y"))
        ctk.CTkButton(
            config_grid,
            text="Hoje",
            command=self.colocar_data_hoje,
            height=42,
            width=120,
            corner_radius=14,
            fg_color="#ffffff",
            text_color=PRIMARY_TEXT,
            hover_color=SOFT_RED,
            border_width=1,
            border_color="#f0d7d2",
            font=("Segoe UI", 14, "bold"),
        ).grid(row=1, column=1, sticky="w")
        self.check_modo = ctk.CTkCheckBox(
            config_grid,
            text="Executar em modo invisivel",
            variable=self.modo_invisivel,
            font=("Segoe UI", 13),
            text_color="#303030",
            checkbox_width=22,
            checkbox_height=22,
            corner_radius=8,
        )
        self.check_modo.grid(row=2, column=0, columnspan=2, sticky="w", pady=(14, 0))

        progresso = self.criar_secao(scroll, "Progresso")
        self.progressbar = ctk.CTkProgressBar(progresso, height=16, corner_radius=999, progress_color=BUTTON_BG, fg_color="#f2dfdb")
        self.progressbar.pack(fill="x", padx=18, pady=(0, 10))
        self.progressbar.set(0)
        self.label_status = ctk.CTkLabel(progresso, text="0/0 - Aguardando inicio...", text_color=MUTED_TEXT, font=("Segoe UI", 13))
        self.label_status.pack(pady=(0, 18))

        acoes = ctk.CTkFrame(scroll, fg_color="transparent")
        acoes.pack(fill="x", padx=8, pady=(0, 8))
        self.btn_iniciar = ctk.CTkButton(
            acoes,
            text="Iniciar Processamento",
            command=self.iniciar_processamento_wrapper,
            height=48,
            width=220,
            corner_radius=14,
            fg_color=BUTTON_BG,
            hover_color=BUTTON_ACTIVE_BG,
            font=("Segoe UI", 16, "bold"),
        )
        self.btn_iniciar.pack(side="left", padx=(0, 10))
        self.btn_limpar = ctk.CTkButton(
            acoes,
            text="Limpar Log",
            command=self.limpar_log,
            height=48,
            width=160,
            corner_radius=14,
            fg_color="#ffffff",
            text_color=PRIMARY_TEXT,
            hover_color=SOFT_RED,
            border_width=1,
            border_color="#f0d7d2",
            font=("Segoe UI", 16, "bold"),
        )
        self.btn_limpar.pack(side="left")

        logs = self.criar_secao(scroll, "Log do Processamento")
        self.log_text = ctk.CTkTextbox(
            logs,
            height=320,
            corner_radius=16,
            fg_color="#fffaf9",
            border_width=1,
            border_color=CARD_BORDER,
            text_color="#2d2d2d",
            font=("Consolas", 12),
        )
        self.log_text.pack(fill="both", expand=True, padx=18, pady=(0, 18))
        self.log_text.configure(state="disabled")

        ctk.CTkLabel(scroll, text="Desenvolvido por Diogo Medeiros ? 2026", text_color="#b85b52", font=("Segoe UI", 11)).pack(anchor="w", padx=12, pady=(0, 12))


    # ====================== VALIDAÇÃO E CONTROLE ======================
    def validar_abertura(self):
        self.escrever_log("Inicializando robô...")
        registrar_abertura()
        if verificar_chave():
            self.escrever_log("✅ Robô pronto para execução.")
        else:
            self.escrever_log("❌ Robô BLOQUEADO pela validação remota.")
            messagebox.showerror("Acesso Bloqueado", "Este robô não está autorizado para execução.")
            self.btn_iniciar.configure(state="disabled")

    def validar_data_repasse(self):
        data_txt = self.data_repasse_var.get().strip()
        try:
            datetime.strptime(data_txt, "%d/%m/%Y")
            return True, ""
        except ValueError:
            return False, "A data do repasse deve estar no formato DD/MM/AAAA."

    def iniciar_processamento_wrapper(self):
        if self.processando:
            return
        if not self.caminho_planilha.get() or not self.aba_selecionada.get():
            messagebox.showwarning("Atenção", "Selecione planilha e aba.")
            return
        if not self.usuario.get() or not self.senha.get():
            messagebox.showwarning("Atenção", "Informe usuário e senha.")
            return
        data_ok, msg_data = self.validar_data_repasse()
        if not data_ok:
            messagebox.showwarning("Atenção", msg_data)
            return

        self.processando = True
        self.btn_iniciar.configure(state="disabled")
        thread = threading.Thread(target=self.processamento_com_reset, daemon=True)
        thread.start()

    # ====================== LÓGICA DE RESET PRINCIPAL ======================
    def processamento_com_reset(self):
        try:
            if self.df is None:
                self.escrever_log("Carregando planilha e filtrando contratos aptos...")
                self.df = pd.read_excel(self.caminho_planilha.get(), sheet_name=self.aba_selecionada.get(), dtype=str)
                self.df = self.filtrar_contratos_aptos(self.df)
                self.total_aptos = len(self.df)
                self.atualizar_contador()
                if self.total_aptos == 0:
                    self.root.after(0, lambda: messagebox.showwarning("Aviso", "Nenhum contrato aberto encontrado."))
                    return
                self.caminho_relatorio_parcial = self.criar_arquivo_relatorio_parcial()

            while self.indice_inicio < len(self.df):
                self.tentativas_do_contrato_atual = 0
                contrato_atual = None
                deve_avancar_indice = True

                while self.tentativas_do_contrato_atual <= MAX_TENTATIVAS_POR_CONTRATO:
                    try:
                        row = self.df.iloc[self.indice_inicio]
                        contrato_atual = str(row.iloc[COLUNAS["contrato"]]).strip()
                        valor_raw = str(row.iloc[COLUNAS["valor_repasse"]]).strip()

                        self.atual = self.indice_inicio + 1
                        self.atualizar_progresso_processamento()

                        if not contrato_atual or contrato_atual.lower() == "nan":
                            self.ignorados += 1
                            self.contratos_ignorados.append(f"{contrato_atual} | Contrato vazio")
                            self.atualizar_labels_resultados()
                            self.escrever_log(f"⚠️ [{self.atual}/{self.total_aptos}] Linha ignorada: contrato vazio.")
                            self.salvar_relatorio_parcial()
                            break

                        valor_ok, valor_str = self.validar_e_normalizar_valor(valor_raw)
                        if not valor_ok:
                            self.ignorados += 1
                            self.contratos_ignorados.append(f"{contrato_atual} | Valor inválido: {valor_raw}")
                            self.atualizar_labels_resultados()
                            self.escrever_log(f"⚠️ [{self.atual}/{self.total_aptos}] Ignorado - valor inválido.")
                            self.salvar_relatorio_parcial()
                            break

                        self.escrever_log(f"🔍 Processando: {contrato_atual} | R$ {valor_str}" + (f" (Tentativa {self.tentativas_do_contrato_atual + 1})" if self.tentativas_do_contrato_atual > 0 else ""))

                        if self.tentativas_do_contrato_atual > 0:
                            self.escrever_log("🔄 Erro anterior. Retornando para a URL base de contratos...")
                            if self.driver:
                                self.preparar_tela_contratos_para_proximo_loop()
                            else:
                                if not self.abrir_coral():
                                    raise Exception("Falha ao abrir o Coral.")

                        if not self.driver or not self.driver.current_url or "coral" not in self.driver.current_url:
                            if not self.abrir_coral():
                                raise Exception("Falha ao abrir o Coral.")

                        self.buscar_contrato_seguro(contrato_atual)
                        status_real = self.obter_status_real()
                        if "AGUARDANDO DEVOLUÇÃO" not in status_real:
                            self.ignorados += 1
                            self.contratos_ignorados.append(f"{contrato_atual} | Status: {status_real}")
                            self.atualizar_labels_resultados()
                            self.escrever_log(f"⏭️ Ignorado - Status no Coral: {status_real}")
                            self.salvar_relatorio_parcial()
                            break

                        sucesso = self.lancar_repasse(contrato_atual, valor_str)
                        if sucesso:
                            self.lancados += 1
                            self.contratos_lancados.append(contrato_atual)
                            self.atualizar_labels_resultados()
                            self.escrever_log(f"✅ [{self.atual}/{self.total_aptos}] Contrato {contrato_atual} lançado com sucesso.")
                            self.salvar_relatorio_parcial()
                            break
                        else:
                            raise Exception(f"Falha no lançamento do contrato {contrato_atual}.")

                    except Exception as e:
                        self.tentativas_do_contrato_atual += 1
                        self.escrever_log(f"❌ Erro na tentativa {self.tentativas_do_contrato_atual} para o contrato {contrato_atual}: {e}")

                        if self.tentativas_do_contrato_atual > MAX_TENTATIVAS_POR_CONTRATO:
                            self.erros += 1
                            self.contratos_com_erro.append(f"{contrato_atual} | {e}")
                            self.atualizar_labels_resultados()
                            self.escrever_log(f"🛑 Segunda tentativa falhou. Contrato {contrato_atual} marcado como ERRO.")
                            self.salvar_relatorio_parcial()
                            self.indice_inicio += 1
                            deve_avancar_indice = False
                            break

                if deve_avancar_indice:
                    self.indice_inicio += 1

            self.escrever_log("\n🎉 Processamento finalizado com sucesso para todos os contratos aptos!")

        finally:
            self.escrever_log("\nExecutando ações finais...")
            try:
                if self.caminho_relatorio_parcial:
                    nome_final = self.caminho_relatorio_parcial.replace("Relatorio_Parcial", "Relatorio_Final")
                    os.rename(self.caminho_relatorio_parcial, nome_final)
                    self.caminho_relatorio_final = nome_final
                    self.escrever_log(f"Relatório parcial renomeado para final: {self.caminho_relatorio_final}")
            except Exception as e:
                self.escrever_log(f"⚠️ Erro ao gerar o relatório final: {e}")

            try:
                if self.driver:
                    self.driver.quit()
                    self.driver = None
            except:
                pass

            self.processando = False
            self.root.after(0, lambda: self.btn_iniciar.configure(state="normal"))
            self.mostrar_resumo_final()

    # ====================== LANÇAMENTO CORRIGIDO ======================
    def lancar_repasse(self, contrato, valor_str):
        try:
            # Menu Ações
            self.clicar_seguro(
                By.XPATH,
                '/html/body/foco-app/div[1]/foco-rent-agreement-home/div/ngb-tabset/div/div/div/div/foco-rent-agreement-list/div/div/div[3]/table/tbody/tr/td[8]/div/div/button',
                "Menu Ações"
            )
            time.sleep(PAUSA_CURTA)

            # Editar
            self.clicar_seguro(
                By.XPATH,
                '/html/body/foco-app/div[1]/foco-rent-agreement-home/div/ngb-tabset/div/div/div/div/foco-rent-agreement-list/div/div/div[3]/table/tbody/tr/td[8]/div/div/div/button[1]',
                "Editar contrato"
            )
            time.sleep(PAUSA_MEDIA)

            # Verificação de pop-up de CPF
            xpath_popup_cadastro = '/html/body/ngb-modal-window/div/div/foco-confirm-modal'
            ja_esta_em_pagamentos = False

            if self.elemento_existe(By.XPATH, xpath_popup_cadastro, timeout=8):
                self.escrever_log("⚠️ Pop-up de cadastro pendente detectado → Caminho secundário")
                self.clicar_seguro(By.XPATH, f'{xpath_popup_cadastro}/div[3]/button[2]', "Botão 'Sim'")
                time.sleep(PAUSA_MEDIA)

                for i in range(1, 5):
                    btn_xpath = '/html/body/foco-app/div[1]/foco-rent-agreement-edit/div/div[3]/div/div/div[2]/button'
                    if i == 2:
                        btn_xpath += '[3]'
                    elif i >= 3:
                        btn_xpath += '[2]'
                    self.clicar_seguro(By.XPATH, btn_xpath, f"Avançar {i}/4")
                    time.sleep(PAUSA_MEDIA)

                ja_esta_em_pagamentos = True
                self.escrever_log("✅ Caminho secundário concluído (já na aba Pagamentos)")

            else:
                self.escrever_log("✅ Sem pop-up. Caminho normal.")

            # Aba Pagamentos (só se necessário)
            if not ja_esta_em_pagamentos:
                self.escrever_log(f"⏳ Aguardando a aba Pagamentos ficar realmente clicável (até {TIMEOUT_PAGAMENTOS}s)...")
                self.clicar_seguro(
                    By.XPATH,
                    '/html/body/foco-app/div[1]/foco-rent-agreement-edit/div/div[1]/div/div/div[2]/div[11]/button',
                    "Aba Pagamentos",
                    timeout=TIMEOUT_PAGAMENTOS
                )
                time.sleep(PAUSA_MEDIA)

            # Preenchimento
            self.clicar_seguro(
                By.XPATH,
                '/html/body/foco-app/div[1]/foco-rent-agreement-edit/div/div[2]/div[6]/foco-rent-agreement-payment/div/div[2]/div/div[2]/div[1]/div[1]/button[4]',
                "Boleto"
            )
            time.sleep(PAUSA_CURTA)

            self.digitar_seguro(
                By.XPATH,
                '/html/body/foco-app/div[1]/foco-rent-agreement-edit/div/div[2]/div[6]/foco-rent-agreement-payment/div/div[2]/div/div[2]/div[6]/div/div[1]/foco-form-date/div/foco-date/div/input[1]',
                self.data_repasse_var.get(),
                "Data"
            )
            time.sleep(PAUSA_CURTA)

            self.digitar_seguro(
                By.XPATH,
                '/html/body/foco-app/div[1]/foco-rent-agreement-edit/div/div[2]/div[6]/foco-rent-agreement-payment/div/div[2]/div/div[2]/div[6]/div/div[2]/foco-form-input/div/div[1]/input',
                valor_str,
                "Valor"
            )
            time.sleep(PAUSA_MEDIA)

            # CONCLUSÃO
            self.clicar_seguro(
                By.XPATH,
                '/html/body/foco-app/div[1]/foco-rent-agreement-edit/div/div[2]/div[6]/foco-rent-agreement-payment/div/div[2]/div/div[2]/div[15]/button',
                "Efetuar pagamento"
            )
            time.sleep(PAUSA_MEDIA)

            self.clicar_seguro(
                By.XPATH,
                '/html/body/foco-app/div[1]/foco-rent-agreement-edit/div/div[3]/div/div/div[2]/button[2]',
                "Concluir"
            )
            time.sleep(PAUSA_MEDIA)

            self.clicar_seguro(
                By.XPATH,
                '/html/body/ngb-modal-window/div/div/foco-confirm-modal/div[3]/button[2]',
                "Atualizar contrato"
            )
            time.sleep(PAUSA_MEDIA)

            # FECHAR MODAL FINAL (ÚNICO MOMENTO QUE CONTA COMO LANÇADO)
            fechar_xpath = '/html/body/ngb-modal-window/div/div/foco-reservation-created/div[3]/button'
            if self.elemento_existe(By.XPATH, fechar_xpath, timeout=10):
                self.clicar_seguro(By.XPATH, fechar_xpath, "Fechar")
                time.sleep(PAUSA_MEDIA)
                self.escrever_log(f"✅ Repasse concluído e modal fechado com sucesso → {contrato}")
                self.preparar_tela_contratos_para_proximo_loop()
                return True
            else:
                self.escrever_log("⚠️ Modal de sucesso não apareceu. Tentando fluxo de descarte...")
                self.descartar_evento_e_ir_para_contratos()
                return False

        except Exception as e:
            self.escrever_log(f"❌ Falha crítica no lançamento de {contrato}: {e}")
            return False

    # ====================== PREPARAÇÃO DA TELA (CORRIGIDO) ======================
    def preparar_tela_contratos_para_proximo_loop(self):
        self.escrever_log("🔄 Preparando tela de contratos para o próximo item...")
        try:
            self.driver.get(URL_CORAL_CONTRATOS)
            time.sleep(PAUSA_CURTA)
            self.clicar_seguro(By.XPATH, XPATH_ABA_CONTRATOS, "Aba de Contratos")
            time.sleep(PAUSA_MEDIA)
            self.atualizar_campo_busca()
            self.escrever_log("✅ Tela de contratos pronta para o próximo contrato.")
            return True
        except Exception as e:
            self.escrever_log(f"⚠️ Falha ao preparar tela: {e}")
            return False

    # ====================== OUTROS MÉTODOS ======================
    def filtrar_contratos_aptos(self, df):
        serie_status = df.iloc[:, COLUNAS["status"]].fillna("").str.upper()
        return df[serie_status.str.contains("ABERTO|ATIVO", na=False)].copy()

    def clicar_seguro(self, by, locator, descricao="", timeout=None):
        timeout_final = timeout if timeout is not None else TIMEOUT_ACAO
        try:
            self.escrever_log(f"🖱️ Clicando em '{descricao}' (timeout de {timeout_final}s)...")
            elemento = WebDriverWait(self.driver, timeout_final).until(
                EC.element_to_be_clickable((by, locator))
            )
            self.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", elemento)
            time.sleep(0.2)
            elemento.click()
            self.escrever_log(f"✅ Clique em '{descricao}' bem-sucedido.")
            return True
        except Exception as e:
            self.escrever_log(f"❌ FALHA ao clicar em '{descricao}': {e}")
            raise

    def digitar_seguro(self, by, locator, texto, descricao=""):
        try:
            self.escrever_log(f"⌨️ Preenchendo '{descricao}' com '{texto}' (timeout de {TIMEOUT_ACAO}s)...")
            campo = WebDriverWait(self.driver, TIMEOUT_ACAO).until(EC.presence_of_element_located((by, locator)))
            self.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", campo)
            time.sleep(0.2)
            campo.clear()
            campo.send_keys(texto)
            self.escrever_log(f"✅ Campo '{descricao}' preenchido com sucesso.")
            return True
        except Exception as e:
            self.escrever_log(f"❌ FALHA ao digitar em '{descricao}': {e}")
            raise

    def elemento_existe(self, by, locator, timeout=8):
        try:
            WebDriverWait(self.driver, timeout).until(EC.presence_of_element_located((by, locator)))
            return True
        except:
            return False

    def buscar_contrato_seguro(self, contrato):
        self.atualizar_campo_busca()
        self.campo_busca.clear()
        self.campo_busca.send_keys(contrato)
        self.campo_busca.send_keys("\n")
        time.sleep(1.5)
        WebDriverWait(self.driver, 15).until(EC.presence_of_element_located((By.XPATH, f'//td[2][contains(text(), "{contrato}")]')))

    def obter_status_real(self):
        xpath = '/html/body/foco-app/div[1]/foco-rent-agreement-home/div/ngb-tabset/div/div/div/div/foco-rent-agreement-list/div/div/div[3]/table/tbody/tr/td[7]'
        elem = WebDriverWait(self.driver, 10).until(EC.presence_of_element_located((By.XPATH, xpath)))
        return elem.text.strip().upper()

    def abrir_coral(self, relogin=False):
        try:
            if self.driver is None:
                options = Options()
                if self.modo_invisivel.get():
                    options.add_argument("--headless=new")
                    options.add_argument("--window-size=1920,1080")
                    options.add_argument("--disable-gpu")
                    options.add_argument("--no-sandbox")
                    options.add_argument("--disable-dev-shm-usage")
                else:
                    options.add_argument("--start-maximized")
                self.driver = webdriver.Chrome(options=options)

            self.driver.get(URL_CORAL_LOGIN)
            time.sleep(PAUSA_MEDIA)

            self.digitar_seguro(By.XPATH, '/html/body/foco-app/div[1]/foco-login/div/div/div/div/div[2]/form/div[1]/input', self.usuario.get(), "Usuário login")
            self.digitar_seguro(By.XPATH, '/html/body/foco-app/div[1]/foco-login/div/div/div/div/div[2]/form/div[2]/input', self.senha.get(), "Senha login")
            self.clicar_seguro(By.XPATH, '/html/body/foco-app/div[1]/foco-login/div/div/div/div/div[2]/form/button', "Botão Entrar")

            self.preparar_tela_contratos_para_proximo_loop()
            self.escrever_log("✅ Coral pronto para lançamentos.")
            return True
        except Exception as e:
            self.escrever_log(f"❌ Erro ao abrir o Coral: {e}")
            return False

    def atualizar_campo_busca(self):
        self.campo_busca = WebDriverWait(self.driver, 15).until(
            EC.presence_of_element_located((By.XPATH, XPATH_CAMPO_BUSCA_CONTRATOS))
        )
        return self.campo_busca

    def descartar_evento_e_ir_para_contratos(self):
        self.escrever_log("⚠️ Executando descarte e retorno para contratos...")
        self.clicar_seguro(By.XPATH, '/html/body/foco-app/div[1]/div/ul/li[5]/a/i', "Menu Contratos para descarte")
        time.sleep(PAUSA_CURTA)
        self.clicar_seguro(By.XPATH, '/html/body/ngb-modal-window/div/div/foco-discard-event-modal/div[3]/button[1]', "Confirmar descarte")
        time.sleep(PAUSA_MEDIA)
        self.preparar_tela_contratos_para_proximo_loop()

    def criar_arquivo_relatorio_parcial(self):
        try:
            base_dir = Path(self.caminho_planilha.get()).parent if self.caminho_planilha.get() else Path.cwd()
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            nome_arquivo = f"Relatorio_Parcial_Repasse_{timestamp}.xlsx"
            caminho_arquivo = base_dir / nome_arquivo
            pd.DataFrame(columns=["CONTRATOS_LANCADOS", "CONTRATOS_COM_ERRO", "CONTRATOS_IGNORADOS"]).to_excel(caminho_arquivo, index=False)
            self.escrever_log(f"📄 Arquivo de relatório parcial criado: {caminho_arquivo}")
            return str(caminho_arquivo)
        except Exception as e:
            self.escrever_log(f"❌ Erro ao criar relatório parcial: {e}")
            return None

    def salvar_relatorio_parcial(self):
        if not self.caminho_relatorio_parcial:
            return
        try:
            max_len = max(len(self.contratos_lancados), len(self.contratos_com_erro), len(self.contratos_ignorados), 1)
            df = pd.DataFrame({
                "CONTRATOS_LANCADOS": self.contratos_lancados + [""] * (max_len - len(self.contratos_lancados)),
                "CONTRATOS_COM_ERRO": self.contratos_com_erro + [""] * (max_len - len(self.contratos_com_erro)),
                "CONTRATOS_IGNORADOS": self.contratos_ignorados + [""] * (max_len - len(self.contratos_ignorados))
            })
            df.to_excel(self.caminho_relatorio_parcial, index=False)
        except:
            pass

    def validar_e_normalizar_valor(self, valor_str):
        valor = valor_str.strip().replace("R$", "").replace(" ", "")
        if not valor or valor.lower() == "nan":
            return False, None
        if "," in valor:
            valor = valor.replace(",", ".")
        try:
            numero = float(valor)
            return True, f"{numero:.2f}"
        except:
            return False, None

    def escrever_log(self, mensagem):
        horario = datetime.now().strftime("%H:%M:%S")
        self.root.after(0, self._escrever_log_safe, f"[{horario}] {mensagem}\n")

    def _escrever_log_safe(self, texto):
        self.log_text.configure(state="normal")
        self.log_text.insert("end", texto)
        self.log_text.see("end")
        self.log_text.configure(state="disabled")

    def limpar_log(self):
        self.log_text.configure(state="normal")
        self.log_text.delete("1.0", "end")
        self.log_text.configure(state="disabled")

    def atualizar_contador(self):
        self.root.after(0, lambda: self.label_contador.configure(text=f"Aptos: {self.total_aptos}"))

    def atualizar_labels_resultados(self):
        self.root.after(0, self._atualizar_labels_resultados_safe)

    def _atualizar_labels_resultados_safe(self):
        self.label_lancados.configure(text=f"Lancados: {self.lancados}")
        self.label_ignorados.configure(text=f"Ignorados: {self.ignorados}")
        self.label_erros.configure(text=f"Erros: {self.erros}")

    def atualizar_progresso_processamento(self):
        self.root.after(0, self._atualizar_progresso_safe)

    def _atualizar_progresso_safe(self):
        if self.total_aptos > 0:
            percent = (self.atual / self.total_aptos) * 100
            self.progresso.set(percent)
            self.progressbar.set(percent / 100)
            self.label_status.configure(text=f"{self.atual}/{self.total_aptos} - Em andamento")

    def colocar_data_hoje(self):
        self.data_repasse_var.set(datetime.now().strftime("%d/%m/%Y"))

    def selecionar_planilha(self):
        arquivo = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx *.xls")])
        if arquivo:
            self.caminho_planilha.set(arquivo)
            self.escrever_log(f"Planilha selecionada: {os.path.basename(arquivo)}")
            self.carregar_abas_e_contar()

    def carregar_abas_e_contar(self):
        try:
            xl = pd.ExcelFile(self.caminho_planilha.get())
            self.combo_abas["values"] = xl.sheet_names
            if xl.sheet_names:
                self.combo_abas.current(0)
                self.aba_selecionada.set(xl.sheet_names[0])
            self.contar_contratos_aptos()
        except Exception as e:
            messagebox.showerror("Erro", str(e))

    def contar_contratos_aptos(self):
        try:
            aba = self.aba_selecionada.get() if self.aba_selecionada.get() else 0
            df_temp = pd.read_excel(self.caminho_planilha.get(), sheet_name=aba, dtype=str)
            df_temp = self.filtrar_contratos_aptos(df_temp)
            self.total_aptos = len(df_temp)
            self.atualizar_contador()
        except Exception as e:
            self.escrever_log(f"⚠️ Não foi possível contar contratos aptos: {e}")

    def mostrar_resumo_final(self):
        def montar_lista(lista, limite=20):
            if not lista:
                return "Nenhum"
            if len(lista) <= limite:
                return "\n".join(f"- {item}" for item in lista)
            exibidos = "\n".join(f"- {item}" for item in lista[:limite])
            restantes = len(lista) - limite
            return f"{exibidos}\n- ... e mais {restantes} item(ns)"

        caminho_relatorio_txt = self.caminho_relatorio_final if self.caminho_relatorio_final else "Não foi possível gerar"
        resumo = (
            f"Processamento concluído.\n\n"
            f"Total aptos: {self.total_aptos}\n"
            f"Lançados: {self.lancados}\n"
            f"Ignorados: {self.ignorados}\n"
            f"Erros: {self.erros}\n\n"
            f"RELATÓRIO EXCEL:\n"
            f"{caminho_relatorio_txt}\n\n"
            f"CONTRATOS LANÇADOS:\n"
            f"{montar_lista(self.contratos_lancados)}\n\n"
            f"CONTRATOS COM ERRO:\n"
            f"{montar_lista(self.contratos_com_erro)}"
        )
        self.root.after(0, lambda: messagebox.showinfo("Resumo final do processamento", resumo))


# ====================== VALIDAÇÃO REMOTA + PING ======================
def registrar_abertura():
    try:
        url = "https://docs.google.com/forms/d/e/1FAIpQLScmxNbTO-vXw0LEOKIyEhSpIl9aTbw8x5hnEI5VY2eVMRh5gQ/formResponse"
        data = {"entry.846583903": getpass.getuser(), "entry.1509395143": datetime.now().strftime("%d/%m/%Y %H:%M:%S")}
        requests.post(url, data=data, timeout=5)
    except:
        pass

def verificar_chave():
    try:
        r = requests.get(URL_VALIDACAO, timeout=10)
        return r.text.strip().upper() == "ATIVO"
    except:
        return True


if __name__ == "__main__":
    root = tk.Tk()
    app = AppRepasse(root)
    root.mainloop()
