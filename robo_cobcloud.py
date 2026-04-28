import json
import csv
import os
import queue
import socket
import subprocess
import sys
import threading
import time
import unicodedata
import urllib.request
from dataclasses import dataclass
from datetime import datetime
from tkinter import messagebox

import customtkinter as ctk
import tkinter as tk
from PIL import Image
from selenium import webdriver
from selenium.common.exceptions import TimeoutException, WebDriverException
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select, WebDriverWait
from webdriver_manager.chrome import ChromeDriverManager


if getattr(sys, "frozen", False):
    os.environ["WDM_LOG_LEVEL"] = "0"


MAIN_BG = "#f6f4f1"
CARD_BG = "#ffffff"
CARD_BORDER = "#eadfdb"
PRIMARY_TEXT = "#d81919"
MUTED_TEXT = "#5c5c5c"
BUTTON_BG = "#ef1a14"
BUTTON_ACTIVE_BG = "#c91410"
SOFT_RED = "#fff1ef"
LINK_BLUE = "#2f64d6"
INTERVALO_ENTRE_ACOES = 0.45
ESPERA_APOS_CLIQUE = 0.75
ESPERA_APOS_SELECAO = 0.45


LOGIN_URL = "https://foco.cobcloud.com.br/login"
BASE_URL = "https://foco.cobcloud.com.br/app/backof/operacional/negativacao?page=0&limit=10"
USUARIO_PADRAO = "admin"
SENHA_PADRAO = "Foco@2025"
APP_DATA_DIR = os.path.join(
    os.environ.get("LOCALAPPDATA") or os.path.expanduser("~"),
    "SistemaFOCO",
    "CobCloud",
)
os.makedirs(APP_DATA_DIR, exist_ok=True)
CHROME_PROFILE_DIR = os.path.join(APP_DATA_DIR, "chrome_profile_cobcloud")
CHECKPOINT_PATH = os.path.join(APP_DATA_DIR, "checkpoint_cobcloud.json")
DEBUG_PORT = 9222
ABA_ENDERECO_XPATH = "/html/body/div[1]/div/div/div[2]/main/div/div/div[1]/div[2]/div/div[2]/div/div[1]/ul/li[4]/a"
LINHAS_ENDERECOS_XPATH = "//div[not(@hidden) and contains(@class,'tab-pane')]//table[.//th[normalize-space()='Endereço']]//tbody/tr"
TABELA_ENDERECOS_XPATH = "//div[not(@hidden) and contains(@class,'tab-pane')]//table[.//th[normalize-space()='Endereço']]"
CONFIRMAR_TORNAR_PRINCIPAL_XPATH = "/html/body/div[5]/div[3]/div/div[3]/button[1]/span[1]"
MENSAGEM_SUCESSO_ENDERECO_XPATH = "/html/body/div[1]/div/div/div[2]/main/div/div/div[1]/div[2]/div/div[2]/div/div[1]/div[1]/div[4]/div[3]"
CONFIRMAR_INCLUSAO_XPATH = "/html/body/div[6]/div[3]/div/div[3]/button[1]/span[1]"
CANCELAR_INCLUSAO_XPATH = "/html/body/div[4]/div[3]/div/div[3]/button[2]/span[1]"
ERRO_DATA_DOCUMENTO_DIRETO_XPATH = "/html/body/div[1]/div/div/div[2]/main/div/div/div[6]/div/div[1]"
ERRO_DATA_DOCUMENTO_XPATH = "//div[contains(@class,'MuiSnackbarContent-message')]"


@dataclass(frozen=True)
class XPaths:
    usuario: str = "/html/body/div[1]/div/div/div[2]/div/div[1]/div/div[3]/input"
    senha: str = "/html/body/div[1]/div/div/div[2]/div/div[2]/div/div[3]/input"
    humano: str = "/html/body//div/div/div[1]/div/label/input"
    acessar: str = "/html/body/div[1]/div/div/div[2]/div/button/span[1]/span"
    validacao_login: str = "/html/body/div[1]/div/div/div[2]/main/div/div/div/div/div[2]/div/div[1]/h3"
    backoffice: str = "/html/body/div[1]/div/div/div[2]/div/header/div/ul/li[1]/a[4]/button/span[1]/svg"
    operacional: str = "/html/body/div[1]/div/div/div[1]/div/div/div[2]/div[1]/nav/div/div[2]/div[3]/div[1]/span[1]"
    negativacao: str = "/html/body/div[1]/div/div/div[1]/div/div/div[2]/div[1]/nav/div/div[2]/div[3]/div[2]/div/div/div/div[4]/a/span"
    tipo: str = "/html/body/div[1]/div/div/div[2]/main/div/div/div[2]/div[1]/div[3]/div/div/select"
    status: str = "/html/body/div[1]/div/div/div[2]/main/div/div/div[2]/div[2]/div[3]/div/div/select"
    pesquisar: str = "/html/body/div[1]/div/div/div[2]/main/div/div/div[2]/div[2]/div[4]/button/span[1]/span"
    localizar_processo: str = "/html/body/div[1]/div/div/div[2]/main/div/div/div[2]/div[1]/div[2]/div/div/input"
    tabela_linhas: str = "//tbody[contains(@class,'MuiTableBody-root')]/tr"
    proxima_pagina: str = "/html/body/div[1]/div/div/div[2]/main/div/div/div[2]/div[3]/div/div/div[3]/button[2]"
    paginacao_caption: str = "//div[contains(@class,'MuiTablePagination-root')]//p[contains(@class,'MuiTablePagination-caption')][last()]"


XP = XPaths()


def localizar_logo():
    candidatos = []
    env_logo = os.environ.get("FOCO_LOGO_PNG", "").strip()
    env_assets = os.environ.get("FOCO_ASSETS_DIR", "").strip()
    if env_logo:
        candidatos.append(env_logo)
    if env_assets:
        candidatos.append(os.path.join(env_assets, "logo.png"))
    if getattr(sys, "_MEIPASS", None):
        candidatos.append(os.path.join(sys._MEIPASS, "assets", "logo.png"))

    base_atual = os.path.dirname(os.path.abspath(__file__))
    candidatos.extend(
        [
            os.path.join(base_atual, "assets", "logo.png"),
            os.path.join(os.getcwd(), "assets", "logo.png"),
            os.path.join(
                os.path.expanduser("~"),
                "OneDrive - Foco Aluguel de Carros",
                "Área de Trabalho",
                "projeto gaby",
                "QUINZENAL",
                "DESENVOLVIMENTO",
                "assets",
                "logo.png",
            ),
        ]
    )
    for caminho in candidatos:
        if os.path.exists(caminho):
            return caminho
    return None


def encontrar_chrome():
    candidatos = [
        os.path.join(os.environ.get("PROGRAMFILES", ""), "Google", "Chrome", "Application", "chrome.exe"),
        os.path.join(os.environ.get("PROGRAMFILES(X86)", ""), "Google", "Chrome", "Application", "chrome.exe"),
        os.path.join(os.environ.get("LOCALAPPDATA", ""), "Google", "Chrome", "Application", "chrome.exe"),
    ]
    for caminho in candidatos:
        if caminho and os.path.exists(caminho):
            return caminho
    return "chrome.exe"


def porta_esta_aberta(porta):
    try:
        with socket.create_connection(("127.0.0.1", porta), timeout=0.5):
            return True
    except OSError:
        return False


def obter_abas_chrome():
    try:
        with urllib.request.urlopen(f"http://127.0.0.1:{DEBUG_PORT}/json", timeout=1) as resposta:
            return json.loads(resposta.read().decode("utf-8", errors="ignore"))
    except Exception:
        return []


def normalizar_texto(texto):
    texto = unicodedata.normalize("NFKD", texto or "")
    texto = "".join(char for char in texto if not unicodedata.combining(char))
    return texto.casefold().strip()


class ControleExecucao:
    def __init__(self):
        self.pausado = threading.Event()
        self.parar = threading.Event()

    def aguardar_liberacao(self):
        while self.pausado.is_set() and not self.parar.is_set():
            time.sleep(0.25)
        if self.parar.is_set():
            raise RuntimeError("Execucao interrompida pelo usuario.")


class CobCloudBot:
    def __init__(self, usuario, senha, headless, controle, log_callback, status_callback, contador_callback):
        self.usuario = usuario
        self.senha = senha
        self.headless = headless
        self.controle = controle
        self.log = log_callback
        self.status = status_callback
        self.contador = contador_callback
        self.driver = None
        self.realizados = 0
        self.avaliados = 0
        self.chaves_ignoradas = set()
        self.processos_tratados = set()
        self.ultima_acao = 0
        self.relatorio_path = os.path.join(
            APP_DATA_DIR,
            f"Relatorio_CobCloud_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
        )
        self.carregar_checkpoint()

    def carregar_checkpoint(self):
        if not os.path.exists(CHECKPOINT_PATH):
            return
        try:
            with open(CHECKPOINT_PATH, "r", encoding="utf-8") as arquivo:
                dados = json.load(arquivo)
            self.processos_tratados = set(dados.get("processos_tratados", []))
            self.chaves_ignoradas = set(dados.get("chaves_ignoradas", []))
            self.log(
                f"Checkpoint carregado: {len(self.processos_tratados)} processo(s) tratado(s), "
                f"{len(self.chaves_ignoradas)} linha(s) ignorada(s)."
            )
        except Exception as exc:
            self.log(f"Nao foi possivel carregar checkpoint: {exc}")

    def salvar_checkpoint(self):
        try:
            dados = {
                "atualizado_em": datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
                "processos_tratados": sorted(p for p in self.processos_tratados if p),
                "chaves_ignoradas": sorted(c for c in self.chaves_ignoradas if c),
            }
            tmp_path = CHECKPOINT_PATH + ".tmp"
            with open(tmp_path, "w", encoding="utf-8") as arquivo:
                json.dump(dados, arquivo, ensure_ascii=False, indent=2)
            os.replace(tmp_path, CHECKPOINT_PATH)
        except Exception as exc:
            self.log(f"Nao foi possivel salvar checkpoint: {exc}")

    def marcar_processo_tratado(self, processo):
        if processo:
            self.processos_tratados.add(processo)
            self.salvar_checkpoint()

    def marcar_linha_ignorada(self, chave):
        if chave:
            self.chaves_ignoradas.add(chave)
            self.salvar_checkpoint()

    def iniciar_driver(self):
        self.log("Abrindo Chrome normal para login manual, ainda sem conectar o Selenium...")
        os.makedirs(CHROME_PROFILE_DIR, exist_ok=True)
        chrome = encontrar_chrome()

        if not porta_esta_aberta(DEBUG_PORT):
            subprocess.Popen(
                [
                    chrome,
                    f"--remote-debugging-port={DEBUG_PORT}",
                    f"--user-data-dir={CHROME_PROFILE_DIR}",
                    "--profile-directory=Default",
                    "--start-maximized",
                    "--lang=pt-BR",
                    LOGIN_URL,
                ],
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
            )
            time.sleep(3)
        else:
            self.log("Chrome de login manual ja esta aberto.")
            subprocess.Popen(
                [chrome, LOGIN_URL],
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
            )

        self.log("Aguardando porta do Chrome ficar disponivel...")
        fim = time.time() + 20
        while time.time() < fim and not porta_esta_aberta(DEBUG_PORT):
            time.sleep(0.5)

        if not porta_esta_aberta(DEBUG_PORT):
            raise RuntimeError("Nao consegui abrir o Chrome para login manual.")

    def conectar_selenium_ao_chrome(self):
        self.log("Conectando Selenium ao Chrome ja logado...")
        options = webdriver.ChromeOptions()
        options.debugger_address = f"127.0.0.1:{DEBUG_PORT}"
        service = Service(ChromeDriverManager().install())
        self.driver = webdriver.Chrome(service=service, options=options)
        self.selecionar_aba_cobcloud()

    def selecionar_aba_cobcloud(self):
        for handle in self.driver.window_handles:
            self.driver.switch_to.window(handle)
            url = self.driver.current_url
            if "foco.cobcloud.com.br/app" in url:
                self.log(f"Aba logada selecionada: {url}")
                return
        for handle in self.driver.window_handles:
            self.driver.switch_to.window(handle)
            url = self.driver.current_url
            if "foco.cobcloud.com.br" in url:
                self.log(f"Aba CobCloud selecionada: {url}")
                return

    def wait(self, timeout=30):
        return WebDriverWait(self.driver, timeout)

    def ritmo_seguro(self, segundos=INTERVALO_ENTRE_ACOES):
        self.controle.aguardar_liberacao()
        decorrido = time.time() - self.ultima_acao
        if decorrido < segundos:
            time.sleep(segundos - decorrido)
        self.ultima_acao = time.time()

    def aguardar_pagina_pronta(self, timeout=30):
        self.controle.aguardar_liberacao()
        try:
            self.wait(timeout).until(
                lambda driver: driver.execute_script("return document.readyState") == "complete"
            )
        except Exception:
            pass

    def aguardar_carregamentos(self, timeout=20):
        self.aguardar_pagina_pronta(timeout=timeout)
        seletores_loading = [
            "//*[contains(@class,'MuiCircularProgress-root')]",
            "//*[contains(@class,'MuiLinearProgress-root')]",
            "//*[contains(@class,'MuiBackdrop-root') and not(contains(@style,'visibility: hidden'))]",
            "//*[contains(@class,'MuiSkeleton-root')]",
            "//*[@role='progressbar']",
        ]
        fim = time.time() + timeout
        while time.time() < fim:
            self.controle.aguardar_liberacao()
            carregando = False
            for seletor in seletores_loading:
                elementos = self.driver.find_elements(By.XPATH, seletor)
                if any(elemento.is_displayed() for elemento in elementos):
                    carregando = True
                    break
            if not carregando:
                return
            time.sleep(0.4)

    def esperar_presente(self, xpath, timeout=30):
        self.controle.aguardar_liberacao()
        return self.wait(timeout).until(EC.presence_of_element_located((By.XPATH, xpath)))

    def esperar_clicavel(self, xpath, timeout=30):
        self.controle.aguardar_liberacao()
        elemento = self.wait(timeout).until(EC.element_to_be_clickable((By.XPATH, xpath)))
        self.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", elemento)
        return elemento

    def clicar(self, xpath, descricao, timeout=30):
        self.log(f"Aguardando {descricao}...")
        self.ritmo_seguro()
        self.aguardar_carregamentos(timeout=3)
        elemento = self.esperar_clicavel(xpath, timeout)
        elemento.click()
        self.log(f"Clique confirmado: {descricao}")
        time.sleep(ESPERA_APOS_CLIQUE)
        self.aguardar_carregamentos(timeout=8)
        return elemento

    def preencher(self, xpath, valor, descricao, timeout=30):
        self.log(f"Aguardando campo {descricao}...")
        self.ritmo_seguro(segundos=0.25)
        self.aguardar_carregamentos(timeout=3)
        campo = self.esperar_clicavel(xpath, timeout)
        campo.clear()
        campo.send_keys(valor)
        self.log(f"Campo preenchido: {descricao}")
        time.sleep(0.5)

    def selecionar_por_valor(self, xpath, valor, descricao, timeout=30):
        self.log(f"Aguardando lista {descricao}...")
        self.ritmo_seguro()
        self.aguardar_carregamentos(timeout=3)
        elemento = self.esperar_clicavel(xpath, timeout)
        Select(elemento).select_by_value(valor)
        self.log(f"Selecionado em {descricao}: {valor}")
        time.sleep(ESPERA_APOS_SELECAO)
        self.aguardar_carregamentos(timeout=6)

    def preencher_input_por_xpath(self, xpath, valor, descricao, timeout=30):
        self.log(f"Aguardando campo {descricao}...")
        self.ritmo_seguro(segundos=0.25)
        self.aguardar_carregamentos(timeout=3)
        campo = self.esperar_clicavel(xpath, timeout)
        campo.clear()
        campo.send_keys(valor)
        self.log(f"Campo preenchido: {descricao} = {valor}")
        time.sleep(0.5)

    def login_cobcloud(self):
        self.status("Abrindo login")

        self.status("Aguardando login manual")
        self.log("Navegador aberto na tela de login.")
        self.log("Digite usuario e senha manualmente, resolva a verificacao de humano e clique em Acessar.")
        self.log("O Selenium so vai conectar depois que o navegador sair da tela de login.")
        self.aguardar_login_manual_por_url(timeout=300)
        self.conectar_selenium_ao_chrome()
        self.log("O robo vai validar a sessao agora.")
        self.validar_login(timeout=20)
        self.abrir_tela_base_direta()

    def aguardar_login_manual_por_url(self, timeout=300):
        fim = time.time() + timeout
        while time.time() < fim:
            self.controle.aguardar_liberacao()
            abas = obter_abas_chrome()
            for aba in abas:
                url = aba.get("url", "")
                if "foco.cobcloud.com.br/app" in url:
                    self.log(f"Login manual detectado: {url}")
                    return
            time.sleep(1)
        raise RuntimeError("Tempo esgotado aguardando login manual no CobCloud.")

    def validar_login(self, timeout=45):
        self.status("Validando login")
        url_atual = self.driver.current_url
        if "foco.cobcloud.com.br/app" in url_atual:
            self.log(f"Login validado pela URL: {url_atual}")
            return

        try:
            titulo = self.esperar_presente(XP.validacao_login, timeout=timeout)
            texto = titulo.text.strip()
            esperado = "Recebimentos (Últimos 30 dias)"
            if texto == esperado:
                self.log("Login validado pelo dashboard.")
                return
            raise RuntimeError(f"Esperado '{esperado}', encontrado '{texto}'.")
        except Exception as exc:
            url_atual = self.driver.current_url
            if "foco.cobcloud.com.br/app" in url_atual:
                self.log(f"Titulo do dashboard nao apareceu, mas a sessao esta logada: {url_atual}")
                return
            raise RuntimeError(f"Login nao validado. URL atual: {url_atual}. Detalhe: {exc}")

    def navegar_tela_base(self):
        self.status("Indo para tela base")
        if self.abrir_tela_base_direta():
            return
        try:
            self.clicar(XP.backoffice, "BackOffice", timeout=20)
            self.clicar(XP.operacional, "menu Operacional", timeout=20)
            self.clicar(XP.negativacao, "submenu Negativacao", timeout=20)
        except Exception as exc:
            self.log(f"Navegacao por menu falhou ({exc}). Abrindo link direto da tela base.")
            self.driver.get(BASE_URL)
        self.validar_tela_base()

    def abrir_tela_base_direta(self):
        self.status("Abrindo tela base")
        self.log("Abrindo link direto da Negativacao...")
        self.driver.get(BASE_URL)
        try:
            self.validar_tela_base()
            return True
        except Exception as exc:
            self.log(f"Link direto da Negativacao ainda nao validou: {exc}")
            return False

    def validar_tela_base(self):
        self.aguardar_carregamentos(timeout=25)
        self.esperar_clicavel(XP.tipo, timeout=35)
        self.esperar_clicavel(XP.status, timeout=35)
        self.log("Tela base de negativacao pronta.")

    def aplicar_filtros(self):
        self.status("Aplicando filtros")
        self.validar_tela_base()
        self.selecionar_por_valor(XP.tipo, "0", "tipo de movimento")
        self.selecionar_por_valor(XP.status, "erro", "status")
        self.clicar(XP.pesquisar, "Pesquisar")
        self.aguardar_carregamentos(timeout=10)
        self.esperar_presente(XP.tabela_linhas, timeout=45)
        time.sleep(0.5)
        self.log("Pesquisa concluida.")

    def pesquisar_processo_negativacao(self, processo):
        self.status(f"Pesquisando negativacao {processo}")
        self.driver.get(BASE_URL)
        self.validar_tela_base()
        self.preencher_input_por_xpath(XP.localizar_processo, processo, "localizar processo")
        self.selecionar_por_valor(XP.tipo, "0", "tipo de movimento")
        self.selecionar_por_valor(XP.status, "erro", "status")
        self.clicar(XP.pesquisar, "Pesquisar")
        self.aguardar_carregamentos(timeout=10)
        self.esperar_presente(XP.tabela_linhas, timeout=45)
        time.sleep(0.5)
        linhas = self.obter_linhas_visiveis()
        self.log(f"Pesquisa do processo {processo}: {len(linhas)} linha(s) encontrada(s).")
        for posicao, linha in enumerate(linhas, 1):
            dados = self.extrair_dados_linha(linha)
            self.log(
                f"Resultado {posicao}: {dados['processo']} | {dados['cliente']} | "
                f"{dados['titulo']} | {dados['parcela']} | {dados['erro'] or 'sem erro'}"
            )
        return linhas

    def obter_linhas_visiveis(self):
        self.controle.aguardar_liberacao()
        return self.driver.find_elements(By.XPATH, XP.tabela_linhas)

    def extrair_dados_linha(self, linha):
        colunas = linha.find_elements(By.XPATH, "./td")
        textos = [c.text.strip() for c in colunas]
        processo = textos[2] if len(textos) > 2 else ""
        carteira = textos[3] if len(textos) > 3 else ""
        cliente = textos[4] if len(textos) > 4 else ""
        documento = textos[5] if len(textos) > 5 else ""
        titulo = textos[7] if len(textos) > 7 else ""
        parcela = textos[8] if len(textos) > 8 else ""
        valor = textos[11] if len(textos) > 11 else ""
        status = textos[12] if len(textos) > 12 else ""
        erro = textos[-1] if textos else ""
        texto_linha = " | ".join(textos)
        return {
            "processo": processo,
            "carteira": carteira,
            "cliente": cliente,
            "documento": documento,
            "titulo": titulo,
            "parcela": parcela,
            "valor": valor,
            "status": status,
            "erro": erro,
            "texto_linha": texto_linha,
        }

    def deve_processar_linha(self, dados):
        return "informar o endereco" in normalizar_texto(dados["texto_linha"])

    def chave_linha(self, dados):
        return "|".join(
            [
                dados.get("processo", ""),
                dados.get("documento", ""),
                dados.get("titulo", ""),
                dados.get("parcela", ""),
            ]
        )

    def registrar_relatorio(self, dados, acao, observacao=""):
        novo_arquivo = not os.path.exists(self.relatorio_path)
        campos = [
            "data_hora",
            "acao",
            "processo",
            "cliente",
            "documento",
            "titulo",
            "parcela",
            "valor",
            "status",
            "erro",
            "observacao",
        ]
        with open(self.relatorio_path, "a", newline="", encoding="utf-8-sig") as arquivo:
            writer = csv.DictWriter(arquivo, fieldnames=campos, delimiter=";")
            if novo_arquivo:
                writer.writeheader()
            writer.writerow(
                {
                    "data_hora": datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
                    "acao": acao,
                    "processo": dados.get("processo", ""),
                    "cliente": dados.get("cliente", ""),
                    "documento": dados.get("documento", ""),
                    "titulo": dados.get("titulo", ""),
                    "parcela": dados.get("parcela", ""),
                    "valor": dados.get("valor", ""),
                    "status": dados.get("status", ""),
                    "erro": dados.get("erro", ""),
                    "observacao": observacao,
                }
            )

    def abrir_processo_linha(self, linha):
        self.ritmo_seguro()
        self.aguardar_carregamentos(timeout=3)
        link = linha.find_element(By.XPATH, "./td[3]//a")
        href = link.get_attribute("href")
        texto = link.text.strip()
        self.log(f"Abrindo processo {texto}: {href}")
        self.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", link)
        self.wait(15).until(lambda _driver: link.is_enabled() and link.is_displayed())
        link.click()
        time.sleep(ESPERA_APOS_CLIQUE)
        self.aguardar_carregamentos(timeout=8)
        self.log(f"Tela do processo aberta: {self.driver.current_url}")

    def valor_input(self, campo_id):
        self.esperar_presente(f"//*[@id='{campo_id}']", timeout=25)
        valor = self.driver.execute_script(
            "const el = document.getElementById(arguments[0]); return el ? (el.value || '') : '';",
            campo_id,
        )
        return str(valor or "").strip()

    def validar_endereco_cadastral(self):
        campos = {
            "Endereco": "tx_endereco",
            "Numero": "tx_end_numero",
            "Bairro": "tx_end_bairro",
            "CEP": "tx_end_cep",
            "Cidade": "tx_end_cidade",
            "UF": "tx_end_uf",
        }
        valores = {nome: self.valor_input(campo_id) for nome, campo_id in campos.items()}
        faltantes = [nome for nome, valor in valores.items() if not valor.strip()]
        if valores["UF"] and len(valores["UF"].strip()) != 2:
            faltantes.append("UF")
        self.log(
            "Endereco atual: "
            + " | ".join(f"{nome}: {valor or 'VAZIO'}" for nome, valor in valores.items())
        )
        return valores, sorted(set(faltantes))

    def clicar_aba_enderecos(self):
        self.clicar(ABA_ENDERECO_XPATH, "aba Endereco", timeout=25)
        self.esperar_presente(TABELA_ENDERECOS_XPATH, timeout=25)

    def dados_endereco_linha(self, linha):
        colunas = linha.find_elements(By.XPATH, "./td")
        textos = [coluna.text.strip() for coluna in colunas]
        return {
            "cadastro": textos[1] if len(textos) > 1 else "",
            "endereco": textos[2] if len(textos) > 2 else "",
            "bairro": textos[3] if len(textos) > 3 else "",
            "cidade": textos[4] if len(textos) > 4 else "",
            "uf": textos[5] if len(textos) > 5 else "",
            "cep": textos[6] if len(textos) > 6 else "",
            "status": textos[7] if len(textos) > 7 else "",
        }

    def endereco_linha_completo(self, endereco):
        obrigatorios = ["endereco", "bairro", "cidade", "uf", "cep"]
        if any(not endereco[campo].strip() for campo in obrigatorios):
            return False
        return len(endereco["uf"].strip()) == 2

    def confirmar_dialogo_opcional(self):
        botoes_confirmacao = [
            CONFIRMAR_TORNAR_PRINCIPAL_XPATH,
            "//button[.//span[contains(normalize-space(),'Confirmar')]]",
            "//button[.//span[contains(normalize-space(),'Sim')]]",
            "//button[.//span[contains(normalize-space(),'OK')]]",
        ]
        for xpath in botoes_confirmacao:
            try:
                botao = self.esperar_clicavel(xpath, timeout=3)
                botao.click()
                self.log("Confirmacao do dialogo realizada.")
                time.sleep(ESPERA_APOS_CLIQUE)
                self.aguardar_carregamentos(timeout=15)
                return True
            except Exception:
                continue
        return False

    def aguardar_sucesso_endereco(self):
        self.log("Aguardando mensagem: Alteracao realizada com sucesso...")
        mensagem = self.esperar_presente(MENSAGEM_SUCESSO_ENDERECO_XPATH, timeout=35)
        texto = mensagem.text.strip()
        if "Alteração realizada com sucesso" not in texto and "Alteracao realizada com sucesso" not in texto:
            raise RuntimeError(f"Mensagem de sucesso inesperada: {texto}")
        self.log(f"Confirmacao recebida: {texto}")

    def selecionar_primeiro_endereco_completo(self):
        self.clicar_aba_enderecos()
        linhas = self.driver.find_elements(By.XPATH, LINHAS_ENDERECOS_XPATH)
        if not linhas:
            self.log("Nenhum endereco localizado na aba Enderecos.")
            return None

        for posicao, linha in enumerate(linhas, 1):
            endereco = self.dados_endereco_linha(linha)
            self.log(
                f"Endereco {posicao}: {endereco['endereco'] or 'VAZIO'} | "
                f"{endereco['bairro'] or 'VAZIO'} | {endereco['cidade'] or 'VAZIO'} | "
                f"{endereco['uf'] or 'VAZIO'} | {endereco['cep'] or 'VAZIO'}"
            )
            if not self.endereco_linha_completo(endereco):
                continue

            botao_principal = linha.find_element(By.XPATH, "./td[1]//button")
            self.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", botao_principal)
            self.wait(15).until(lambda _driver: botao_principal.is_enabled() and botao_principal.is_displayed())
            self.ritmo_seguro()
            botao_principal.click()
            self.log(f"Endereco completo selecionado: {endereco['endereco']} - {endereco['cidade']}/{endereco['uf']}")
            time.sleep(ESPERA_APOS_CLIQUE)
            self.confirmar_dialogo_opcional()
            self.aguardar_sucesso_endereco()
            self.aguardar_carregamentos(timeout=25)
            return endereco

        self.log("Nenhum endereco completo encontrado para tornar principal.")
        return None

    def tratar_endereco_do_processo(self, dados):
        _valores, faltantes = self.validar_endereco_cadastral()
        if not faltantes:
            self.log(f"Processo {dados['processo']} ja possui endereco cadastral completo.")
            self.registrar_relatorio(dados, "ENDERECO_OK", "Endereco cadastral completo")
            return "ok"

        self.log(f"Endereco cadastral incompleto. Campos faltando: {', '.join(faltantes)}")
        endereco = self.selecionar_primeiro_endereco_completo()
        if endereco:
            observacao = (
                f"Endereco principal escolhido: {endereco['endereco']} | {endereco['bairro']} | "
                f"{endereco['cidade']}/{endereco['uf']} | {endereco['cep']}"
            )
            self.registrar_relatorio(dados, "ENDERECO_ATUALIZADO", observacao)
            return "atualizado"

        self.registrar_relatorio(dados, "SEM_ENDERECO_COMPLETO", "Nao encontrou endereco completo na aba Enderecos")
        return "sem_endereco"

    def clicar_acao_inclusao(self, linha):
        self.ritmo_seguro()
        self.aguardar_carregamentos(timeout=3)
        botao_acoes = linha.find_element(By.XPATH, "./td[2]//button")
        self.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", botao_acoes)
        self.wait(15).until(lambda _driver: botao_acoes.is_enabled() and botao_acoes.is_displayed())
        self.driver.execute_script("arguments[0].click();", botao_acoes)
        self.log("Menu Acoes aberto.")
        time.sleep(0.2)
        if self.clicar_opcao_inclusao_rapida(timeout=4):
            self.log("Acao Inclusao selecionada.")
            time.sleep(0.25)
            return
        raise RuntimeError("Nao encontrei a opcao Inclusao no menu Acoes.")

    def clicar_opcao_inclusao_rapida(self, timeout=4):
        fim = time.time() + timeout
        script = """
            const norm = (s) => (s || '').normalize('NFD').replace(/[\\u0300-\\u036f]/g, '').toLowerCase();
            const menus = Array.from(document.querySelectorAll('.MuiMenu-paper, .MuiPopover-paper, [role="menu"]'));
            for (const menu of menus) {
                const menuRect = menu.getBoundingClientRect();
                const menuStyle = window.getComputedStyle(menu);
                if (menuRect.width === 0 || menuRect.height === 0 || menuStyle.visibility === 'hidden' || menuStyle.display === 'none') continue;
                const items = Array.from(menu.querySelectorAll('li[role="menuitem"], [role="menuitem"]'));
                for (const item of items) {
                    if ((item.getAttribute('aria-disabled') || '').toLowerCase() === 'true') continue;
                    const text = norm(item.innerText || item.textContent || '');
                    if (text.trim() !== 'inclusao' && !text.includes('inclusao')) continue;
                    const rect = item.getBoundingClientRect();
                    const style = window.getComputedStyle(item);
                    if (rect.width === 0 || rect.height === 0 || style.visibility === 'hidden' || style.display === 'none') continue;
                    item.click();
                    return true;
                }
            }
            const nodes = Array.from(document.querySelectorAll('li[role="menuitem"], [role="menuitem"]'));
            for (const node of nodes) {
                const text = norm(node.innerText || node.textContent || '');
                if (!text.includes('inclusao')) continue;
                if ((node.getAttribute('aria-disabled') || '').toLowerCase() === 'true') continue;
                const rect = node.getBoundingClientRect();
                const style = window.getComputedStyle(node);
                if (rect.width === 0 || rect.height === 0 || style.visibility === 'hidden' || style.display === 'none') continue;
                node.click();
                return true;
            }
            return false;
        """
        while time.time() < fim:
            self.controle.aguardar_liberacao()
            try:
                if self.driver.execute_script(script):
                    return True
            except Exception:
                pass
            time.sleep(0.15)
        return False

    def confirmar_inclusao(self):
        self.clicar_confirmar_inclusao()
        time.sleep(1.0)
        self.cancelar_modal_inclusao()
        return "confirmada_cancelada"

    def clicar_confirmar_inclusao(self):
        seletores_confirmar = [
            CONFIRMAR_INCLUSAO_XPATH,
            "//div[contains(@class,'MuiDialog-root')]//button[.//span[contains(normalize-space(),'Confirmar')]]",
            "//div[contains(@class,'MuiDialog-root')]//button[contains(normalize-space(),'Confirmar')]",
            "//button[.//span[contains(normalize-space(),'Confirmar')]]",
            "//button[contains(normalize-space(),'Confirmar')]",
            "(//button[.//span[contains(normalize-space(),'Confirmar')]])[last()]",
        ]
        ultimo_erro = None
        for xpath in seletores_confirmar:
            try:
                return self.clicar_rapido_por_xpath(xpath, "confirmar inclusao", timeout=4)
            except Exception as exc:
                ultimo_erro = exc
        raise RuntimeError(f"Nao consegui clicar em confirmar inclusao: {ultimo_erro}")

    def clicar_rapido_por_xpath(self, xpath, descricao, timeout=8):
        self.log(f"Aguardando {descricao}...")
        fim = time.time() + timeout
        ultimo_erro = None
        while time.time() < fim:
            self.controle.aguardar_liberacao()
            try:
                elementos = self.driver.find_elements(By.XPATH, xpath)
                for elemento in elementos:
                    if elemento.is_displayed():
                        clicavel = self.driver.execute_script(
                            "return arguments[0].closest('button, [role=\"button\"], li, [role=\"menuitem\"]') || arguments[0];",
                            elemento,
                        )
                        self.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", clicavel)
                        self.driver.execute_script("arguments[0].click();", clicavel)
                        self.log(f"Clique confirmado: {descricao}")
                        return clicavel
            except Exception as exc:
                ultimo_erro = exc
            time.sleep(0.15)
        raise RuntimeError(f"Nao consegui clicar em {descricao}: {ultimo_erro}")

    def erro_data_documento_apareceu(self, timeout=8):
        fim = time.time() + timeout
        while time.time() < fim:
            mensagens = []
            mensagens.extend(self.driver.find_elements(By.XPATH, ERRO_DATA_DOCUMENTO_DIRETO_XPATH))
            mensagens.extend(self.driver.find_elements(By.XPATH, ERRO_DATA_DOCUMENTO_XPATH))
            for mensagem in mensagens:
                try:
                    if not mensagem.is_displayed():
                        continue
                    texto = mensagem.text.strip()
                    if "data do documento e obrigatoria" in normalizar_texto(texto):
                        self.log(f"Erro esperado na inclusao: {texto}")
                        return True
                except Exception:
                    continue
            time.sleep(0.25)
        return False

    def cancelar_modal_inclusao(self):
        self.log("Cancelando modal de inclusao apos erro esperado...")
        seletores_cancelar = [
            CANCELAR_INCLUSAO_XPATH,
            "//div[contains(@class,'MuiDialog-root')]//button[.//span[contains(normalize-space(),'Cancelar')]]",
            "//button[.//span[contains(normalize-space(),'Cancelar')]]",
            "(//button[.//span[contains(normalize-space(),'Cancelar')]])[last()]",
        ]
        ultimo_erro = None
        for xpath in seletores_cancelar:
            try:
                botao = self.clicar_rapido_por_xpath(xpath, "cancelar inclusao", timeout=2)
                self.log("Modal de inclusao cancelado.")
                time.sleep(0.3)
                return
            except Exception as exc:
                ultimo_erro = exc
        raise RuntimeError(f"Nao consegui clicar em Cancelar no modal de inclusao: {ultimo_erro}")

    def incluir_negativacoes_do_processo(self, dados):
        processo = dados["processo"]
        self.pesquisar_processo_negativacao(processo)
        inclusoes = 0
        tentativas = 0
        linhas_processadas = set()

        while tentativas < 30:
            tentativas += 1
            linhas = self.obter_linhas_visiveis()
            alvo = None
            alvo_dados = None
            for linha in linhas:
                linha_dados = self.extrair_dados_linha(linha)
                chave = self.chave_linha(linha_dados)
                if linha_dados["processo"] != processo or chave in linhas_processadas:
                    continue
                alvo = linha
                alvo_dados = linha_dados
                break

            if alvo is None:
                break

            self.log(
                f"Incluindo negativacao: {alvo_dados['processo']} | "
                f"{alvo_dados['titulo']} | {alvo_dados['parcela']} | {alvo_dados['valor']}"
            )
            self.clicar_acao_inclusao(alvo)
            resultado_inclusao = self.confirmar_inclusao()
            inclusoes += 1
            linhas_processadas.add(self.chave_linha(alvo_dados))
            if resultado_inclusao in {"erro_data_documento", "confirmada_cancelada"}:
                self.registrar_relatorio(
                    alvo_dados,
                    "INCLUSAO_CONTABILIZADA_CANCELADA",
                    "Confirmou inclusao e clicou em cancelar apos 1 segundo",
                )
                self.log("Inclusao contabilizada e modal cancelado. Verificando outras linhas do mesmo processo.")
            else:
                self.registrar_relatorio(alvo_dados, "INCLUSAO_SOLICITADA", "Acao Inclusao confirmada")

            time.sleep(0.35)

        if inclusoes == 0:
            self.registrar_relatorio(dados, "SEM_LINHAS_PARA_INCLUSAO", "Nenhuma linha encontrada para inclusao")
            self.log(f"Nenhuma linha para inclusao encontrada para {processo}.")
        else:
            self.log(f"Inclusoes solicitadas para {processo}: {inclusoes}")
        return inclusoes

    def processar_linha(self, indice, linha):
        dados = self.extrair_dados_linha(linha)
        processo = dados["processo"]
        cliente = dados["cliente"]
        documento = dados["documento"]
        erro = dados["erro"]
        chave = self.chave_linha(dados)
        if chave in self.chaves_ignoradas or processo in self.processos_tratados:
            self.log(f"[{indice}] Ignorando ja tratado: {processo or 'sem processo'}")
            return False

        self.avaliados += 1

        if not self.deve_processar_linha(dados):
            self.log(f"[{indice}] Pulando {processo or 'sem processo'} | Erro diferente: {erro or 'sem descricao'}")
            self.registrar_relatorio(dados, "PULADO", "Erro diferente de Informar o endereco")
            self.marcar_linha_ignorada(chave)
            return False

        self.log(f"[{indice}] SELECIONADO para endereco: Processo {processo} | {cliente} | {documento}")
        self.registrar_relatorio(dados, "SELECIONADO", "Informar o endereco")
        self.status(f"Processando {processo or indice}")

        try:
            self.abrir_processo_linha(linha)
            time.sleep(0.7)
            resultado = self.tratar_endereco_do_processo(dados)
            self.realizados += 1
            self.contador(self.realizados)
            self.marcar_processo_tratado(processo)
            self.log(f"Tratamento de endereco finalizado para {processo}: {resultado}")
            if resultado in {"ok", "atualizado"}:
                inclusoes = self.incluir_negativacoes_do_processo(dados)
                self.log(f"Tratamento de inclusao finalizado para {processo}: {inclusoes} linha(s)")
            self.recuperar_tela_base()
            return True
        except Exception as exc:
            self.log(f"Falha ao abrir o processo {processo}: {exc}")
            self.marcar_linha_ignorada(chave)
            self.recuperar_tela_base()
            return True

    def recuperar_tela_base(self):
        self.status("Recuperando tela base")
        self.driver.get(BASE_URL)
        self.validar_tela_base()
        self.aplicar_filtros()

    def ir_proxima_pagina(self):
        try:
            caption_antes = self.texto_paginacao()
            botao = self.esperar_presente(XP.proxima_pagina, timeout=5)
            if self.botao_desabilitado(botao):
                self.log("Botao de proxima pagina esta desabilitado.")
                return False
            self.ritmo_seguro()
            self.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", botao)
            self.driver.execute_script("arguments[0].click();", botao)
            time.sleep(ESPERA_APOS_CLIQUE)
            self.aguardar_carregamentos(timeout=8)
            self.esperar_presente(XP.tabela_linhas, timeout=20)
            mudou = self.aguardar_paginacao_mudar(caption_antes, timeout=8)
            if not mudou:
                self.log("Clique em proxima pagina nao alterou a paginacao; considerando fim.")
                return False
            self.log(f"Pagina alterada: {caption_antes or 'sem contador'} -> {self.texto_paginacao()}")
            time.sleep(0.25)
            return True
        except Exception:
            return False

    def botao_desabilitado(self, elemento):
        classes = elemento.get_attribute("class") or ""
        return (
            elemento.get_attribute("disabled") is not None
            or elemento.get_attribute("aria-disabled") == "true"
            or "Mui-disabled" in classes
        )

    def texto_paginacao(self):
        try:
            elementos = self.driver.find_elements(By.XPATH, XP.paginacao_caption)
            textos = [e.text.strip() for e in elementos if e.is_displayed() and e.text.strip()]
            return textos[-1] if textos else ""
        except Exception:
            return ""

    def aguardar_paginacao_mudar(self, caption_antes, timeout=8):
        fim = time.time() + timeout
        while time.time() < fim:
            atual = self.texto_paginacao()
            if atual and atual != caption_antes:
                return True
            time.sleep(0.25)
        return False

    def executar(self, limite=None):
        inicio = time.time()
        try:
            self.iniciar_driver()
            self.login_cobcloud()
            self.aplicar_filtros()
            self.log(f"Relatorio em tempo real: {self.relatorio_path}")

            pagina = 1
            while not self.controle.parar.is_set():
                linhas = self.obter_linhas_visiveis()
                if not linhas:
                    self.log("Nenhuma linha encontrada na pagina atual.")
                    break

                self.log(f"Pagina {pagina}: {len(linhas)} linhas visiveis.")
                reiniciou_base = False
                for indice, linha in enumerate(linhas, 1):
                    self.controle.aguardar_liberacao()
                    if limite and self.realizados >= limite:
                        self.log(f"Limite de teste atingido: {limite}.")
                        return
                    if self.processar_linha(indice, linha):
                        reiniciou_base = True
                        break

                if reiniciou_base:
                    continue

                if not self.ir_proxima_pagina():
                    self.log("Nao ha proxima pagina disponivel.")
                    break
                pagina += 1
        finally:
            tempo = time.time() - inicio
            self.status("Finalizado")
            self.log(
                f"Finalizado. Avaliados: {self.avaliados} | Selecionados: {self.realizados} | "
                f"Tempo: {formatar_tempo(tempo)}"
            )
            if self.driver:
                try:
                    self.driver.quit()
                except WebDriverException:
                    pass


def formatar_tempo(segundos):
    segundos = int(segundos)
    horas, resto = divmod(segundos, 3600)
    minutos, seg = divmod(resto, 60)
    if horas:
        return f"{horas:02d}:{minutos:02d}:{seg:02d}"
    return f"{minutos:02d}:{seg:02d}"


class RoboCobCloudApp:
    def __init__(self):
        ctk.set_appearance_mode("light")
        ctk.set_default_color_theme("blue")

        self.root = ctk.CTk()
        self.root.title("Robo CobCloud - Registro de Endereco e Inclusao de Cobranca")
        self.root.geometry("980x760")
        self.root.minsize(900, 680)
        self.root.configure(fg_color=MAIN_BG)

        self.usuario_var = tk.StringVar(value=USUARIO_PADRAO)
        self.senha_var = tk.StringVar(value=SENHA_PADRAO)
        self.headless_var = tk.BooleanVar(value=False)
        self.limite_var = tk.StringVar(value="")
        self.status_var = tk.StringVar(value="Aguardando inicio")
        self.contador_var = tk.StringVar(value="0")
        self.tempo_var = tk.StringVar(value="00:00")

        self.log_queue = queue.Queue()
        self.controle = ControleExecucao()
        self.thread = None
        self.inicio_execucao = None
        self.logo_image = None

        self.criar_widgets()
        self.root.after(200, self.processar_filas)
        self.root.after(1000, self.atualizar_tempo)

    def carregar_logo(self, reducao=2):
        caminho_logo = localizar_logo()
        if not caminho_logo:
            return None
        try:
            imagem = Image.open(caminho_logo)
            largura, altura = imagem.size
            largura = max(1, largura // reducao)
            altura = max(1, altura // reducao)
            self.logo_image = ctk.CTkImage(light_image=imagem, dark_image=imagem, size=(largura, altura))
            return self.logo_image
        except Exception:
            return None

    def criar_secao(self, parent, titulo):
        frame = ctk.CTkFrame(parent, fg_color=CARD_BG, corner_radius=20, border_width=1, border_color=CARD_BORDER)
        frame.pack(fill="x", padx=8, pady=8)
        ctk.CTkLabel(frame, text=titulo, text_color=PRIMARY_TEXT, font=("Segoe UI", 18, "bold")).pack(
            anchor="w", padx=18, pady=(16, 12)
        )
        return frame

    def criar_widgets(self):
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

        texto = ctk.CTkFrame(hero_inner, fg_color="transparent")
        texto.pack(side="left", fill="x", expand=True)
        ctk.CTkLabel(texto, text="CobCloud FOCO", text_color=PRIMARY_TEXT, font=("Segoe UI", 30, "bold")).pack(anchor="w")
        ctk.CTkLabel(
            texto,
            text="Registro de endereco e inclusao de cobranca com esperas seguras em cada etapa.",
            text_color=MUTED_TEXT,
            font=("Segoe UI", 14),
        ).pack(anchor="w", pady=(6, 0))
        ctk.CTkLabel(texto, text="BACKOFFICE - NEGATIVACAO", text_color="#a65f56", font=("Segoe UI", 12, "bold")).pack(
            anchor="w", pady=(10, 0)
        )

        acesso = self.criar_secao(scroll, "Acesso CobCloud")
        grid = ctk.CTkFrame(acesso, fg_color="transparent")
        grid.pack(fill="x", padx=18, pady=(0, 18))
        grid.grid_columnconfigure((0, 1), weight=1)

        ctk.CTkLabel(grid, text="Usuario", font=("Segoe UI", 13, "bold"), text_color="#303030").grid(
            row=0, column=0, sticky="w", padx=(0, 10), pady=(0, 6)
        )
        ctk.CTkLabel(grid, text="Senha", font=("Segoe UI", 13, "bold"), text_color="#303030").grid(
            row=0, column=1, sticky="w", padx=(10, 0), pady=(0, 6)
        )
        ctk.CTkEntry(grid, textvariable=self.usuario_var, height=42, corner_radius=12).grid(
            row=1, column=0, sticky="ew", padx=(0, 10)
        )
        ctk.CTkEntry(grid, textvariable=self.senha_var, show="*", height=42, corner_radius=12).grid(
            row=1, column=1, sticky="ew", padx=(10, 0)
        )

        config = self.criar_secao(scroll, "Configuracoes")
        config_grid = ctk.CTkFrame(config, fg_color="transparent")
        config_grid.pack(fill="x", padx=18, pady=(0, 18))
        config_grid.grid_columnconfigure(1, weight=1)

        ctk.CTkCheckBox(
            config_grid,
            text="Executar em modo invisivel",
            variable=self.headless_var,
            font=("Segoe UI", 13),
            text_color="#303030",
            checkbox_width=22,
            checkbox_height=22,
            corner_radius=8,
        ).grid(row=0, column=0, sticky="w", padx=(0, 18))
        ctk.CTkLabel(config_grid, text="Limite de teste", font=("Segoe UI", 13, "bold"), text_color="#303030").grid(
            row=0, column=1, sticky="e", padx=(0, 8)
        )
        ctk.CTkEntry(config_grid, textvariable=self.limite_var, width=90, height=36, corner_radius=10).grid(
            row=0, column=2, sticky="e"
        )

        progresso = self.criar_secao(scroll, "Progresso da Execucao")
        cards = ctk.CTkFrame(progresso, fg_color="transparent")
        cards.pack(fill="x", padx=18, pady=(0, 18))
        cards.grid_columnconfigure((0, 1, 2), weight=1)
        self.criar_indicador(cards, "Realizados", self.contador_var, 0)
        self.criar_indicador(cards, "Tempo", self.tempo_var, 1)
        self.criar_indicador(cards, "Status", self.status_var, 2)

        acoes = ctk.CTkFrame(scroll, fg_color="transparent")
        acoes.pack(fill="x", padx=8, pady=(0, 8))
        self.btn_iniciar = self.criar_botao_primario(acoes, "Iniciar", self.iniciar_robo)
        self.btn_iniciar.pack(side="left", padx=(0, 10))
        self.btn_pausar = self.criar_botao_secundario(acoes, "Pausar", self.alternar_pausa)
        self.btn_pausar.pack(side="left", padx=(0, 10))
        self.btn_parar = self.criar_botao_secundario(acoes, "Parar", self.parar_robo)
        self.btn_parar.pack(side="left")

        logs = self.criar_secao(scroll, "Logs em Tempo Real")
        self.log_text = ctk.CTkTextbox(
            logs,
            height=290,
            corner_radius=16,
            fg_color="#fffaf9",
            border_width=1,
            border_color=CARD_BORDER,
            text_color="#2d2d2d",
            font=("Consolas", 12),
        )
        self.log_text.pack(fill="both", expand=True, padx=18, pady=(0, 18))
        self.log_text.configure(state="disabled")

        ctk.CTkLabel(scroll, text="Desenvolvido por Diogo Medeiros © 2026", text_color="#b85b52", font=("Segoe UI", 11)).pack(
            anchor="w", padx=12, pady=(0, 12)
        )

    def criar_indicador(self, parent, titulo, variavel, coluna):
        frame = ctk.CTkFrame(parent, fg_color="#fffaf9", corner_radius=16, border_width=1, border_color=CARD_BORDER)
        frame.grid(row=0, column=coluna, sticky="ew", padx=6)
        ctk.CTkLabel(frame, text=titulo, text_color=MUTED_TEXT, font=("Segoe UI", 12, "bold")).pack(pady=(14, 2))
        ctk.CTkLabel(frame, textvariable=variavel, text_color=PRIMARY_TEXT, font=("Segoe UI", 24, "bold")).pack(pady=(0, 14))

    def criar_botao_primario(self, parent, texto, comando):
        return ctk.CTkButton(
            parent,
            text=texto,
            command=comando,
            height=48,
            width=160,
            corner_radius=14,
            fg_color=BUTTON_BG,
            hover_color=BUTTON_ACTIVE_BG,
            font=("Segoe UI", 16, "bold"),
        )

    def criar_botao_secundario(self, parent, texto, comando):
        return ctk.CTkButton(
            parent,
            text=texto,
            command=comando,
            height=48,
            width=140,
            corner_radius=14,
            fg_color="#ffffff",
            text_color=PRIMARY_TEXT,
            hover_color=SOFT_RED,
            border_width=1,
            border_color="#f0d7d2",
            font=("Segoe UI", 16, "bold"),
        )

    def log(self, mensagem):
        horario = datetime.now().strftime("%H:%M:%S")
        self.log_queue.put(f"[{horario}] {mensagem}")

    def set_status(self, texto):
        self.log_queue.put(("status", texto))

    def set_contador(self, valor):
        self.log_queue.put(("contador", valor))

    def processar_filas(self):
        while not self.log_queue.empty():
            item = self.log_queue.get()
            if isinstance(item, tuple) and item[0] == "status":
                self.status_var.set(item[1])
                continue
            if isinstance(item, tuple) and item[0] == "contador":
                self.contador_var.set(str(item[1]))
                continue
            self.log_text.configure(state="normal")
            self.log_text.insert("end", item + "\n")
            self.log_text.see("end")
            self.log_text.configure(state="disabled")
        self.root.after(200, self.processar_filas)

    def atualizar_tempo(self):
        if self.inicio_execucao and self.thread and self.thread.is_alive():
            self.tempo_var.set(formatar_tempo(time.time() - self.inicio_execucao))
        self.root.after(1000, self.atualizar_tempo)

    def iniciar_robo(self):
        if self.thread and self.thread.is_alive():
            messagebox.showwarning("Atencao", "O robo ja esta em execucao.")
            return

        usuario = self.usuario_var.get().strip()
        senha = self.senha_var.get().strip()
        if not usuario or not senha:
            messagebox.showwarning("Atencao", "Informe usuario e senha.")
            return

        limite = None
        if self.limite_var.get().strip():
            try:
                limite = int(self.limite_var.get().strip())
            except ValueError:
                messagebox.showwarning("Atencao", "O limite de teste precisa ser numerico.")
                return

        self.controle = ControleExecucao()
        self.inicio_execucao = time.time()
        self.contador_var.set("0")
        self.tempo_var.set("00:00")
        self.status_var.set("Iniciando")
        self.btn_iniciar.configure(state="disabled")
        self.btn_pausar.configure(text="Pausar")
        self.log_text.configure(state="normal")
        self.log_text.delete("1.0", "end")
        self.log_text.configure(state="disabled")

        self.thread = threading.Thread(target=self.executar_thread, args=(usuario, senha, limite), daemon=True)
        self.thread.start()

    def executar_thread(self, usuario, senha, limite):
        try:
            bot = CobCloudBot(
                usuario=usuario,
                senha=senha,
                headless=self.headless_var.get(),
                controle=self.controle,
                log_callback=self.log,
                status_callback=self.set_status,
                contador_callback=self.set_contador,
            )
            bot.executar(limite=limite)
        except Exception as exc:
            self.log(f"ERRO: {exc}")
            self.set_status("Erro")
        finally:
            self.btn_iniciar.configure(state="normal")

    def alternar_pausa(self):
        if self.controle.pausado.is_set():
            self.controle.pausado.clear()
            self.btn_pausar.configure(text="Pausar")
            self.status_var.set("Retomando")
            self.log("Execucao retomada.")
        else:
            self.controle.pausado.set()
            self.btn_pausar.configure(text="Retomar")
            self.status_var.set("Pausado")
            self.log("Pausa solicitada. O robo vai parar no proximo ponto seguro.")

    def parar_robo(self):
        self.controle.parar.set()
        self.status_var.set("Parando")
        self.log("Parada solicitada. Aguardando ponto seguro para encerrar.")


if __name__ == "__main__":
    app = RoboCobCloudApp()
    app.root.mainloop()
