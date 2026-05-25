import os
import queue
import re
import socket
import subprocess
import sys
import threading
import time
import urllib.request
from datetime import date, datetime
from decimal import Decimal, InvalidOperation
from tkinter import messagebox

import customtkinter as ctk
import tkinter as tk
from PIL import Image
from selenium import webdriver
from selenium.common.exceptions import StaleElementReferenceException, WebDriverException
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
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

LOGIN_URL = "https://foco.cobcloud.com.br/login"
BASE_URL = "https://foco.cobcloud.com.br/app/cob/acordo?page=0&limit=100"
RECEBIMENTOS_LIBERADOS_URL = "https://foco.cobcloud.com.br/app/fin/recebimentos/liberados-baixa?page=0&limit=10"
BOLETOS_BANCARIOS_URL = "https://foco.cobcloud.com.br/app/fin/boletos?page=0&limit=10"
BANCO_DO_BRASIL_ID = "3ea0ee9c-c8e0-473b-9a22-b0544eb27d47"
APP_DATA_DIR = os.path.join(
    os.environ.get("LOCALAPPDATA") or os.path.expanduser("~"),
    "SistemaFOCO",
    "CobCloudBaixaPagamentos",
)
os.makedirs(APP_DATA_DIR, exist_ok=True)

CHROME_PROFILE_DIR = os.path.join(APP_DATA_DIR, "chrome_profile_cobcloud_baixa")
CHECKPOINT_PATH = os.path.join(APP_DATA_DIR, "checkpoint_cobcloud_baixa_pagamentos.json")
DEBUG_PORT = 9223
LIMITE_RECEBIMENTOS_POR_EXECUCAO = 200


def primeiro_dia_mes_vigente(data_base=None):
    data_base = data_base or date.today()
    return date(data_base.year, data_base.month, 1).strftime("%d/%m/%Y")


def primeiro_dia_mes_vigente_iso(data_base=None):
    data_base = data_base or date.today()
    return date(data_base.year, data_base.month, 1).strftime("%Y-%m-%d")


def valor_primeiro_dia_mes_para_tipo_input(tipo_input, data_base=None):
    if (tipo_input or "").strip().lower() == "date":
        return primeiro_dia_mes_vigente_iso(data_base)
    return primeiro_dia_mes_vigente(data_base)


def converter_moeda_brasileira(valor):
    normalizado = valor.replace(".", "").replace(",", ".")
    try:
        return Decimal(normalizado)
    except InvalidOperation:
        return Decimal("0")


def texto_indica_recebimentos_liberados(texto):
    if not texto:
        return False
    valores = re.findall(r"R\$\s*([0-9.]+,\d{2})", texto)
    return any(converter_moeda_brasileira(valor) > 0 for valor in valores)


def montar_teclas_data_texto(data_texto):
    return [Keys.CONTROL, "a"] + [Keys.BACKSPACE] * 12 + list(data_texto)


def montar_teclas_data_recebimento(data_recebimento):
    return montar_teclas_data_texto(data_recebimento)


def eh_bolinha_recebimento_pendente(cor_css):
    if not cor_css:
        return False
    numeros = []
    atual = ""
    for caractere in cor_css:
        if caractere.isdigit():
            atual += caractere
            continue
        if atual:
            numeros.append(int(atual))
            atual = ""
    if atual:
        numeros.append(int(atual))
    if len(numeros) < 3:
        return False
    r, g, b = numeros[:3]
    return abs(r - 51) <= 3 and abs(g - 153) <= 3 and abs(b - 255) <= 3


def elemento_visivel_seguro(elemento):
    try:
        return elemento.is_displayed()
    except StaleElementReferenceException:
        return False


def elemento_habilitado_seguro(elemento):
    try:
        return elemento.is_enabled()
    except StaleElementReferenceException:
        return False


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
            os.path.join(os.path.dirname(base_atual), "assets", "logo.png"),
            os.path.join(os.getcwd(), "assets", "logo.png"),
            os.path.join(
                os.path.expanduser("~"),
                "OneDrive - Foco Aluguel de Carros",
                "Area de Trabalho",
                "SISTEMA FOCO",
                "DESENVOLVIMENTO",
                "assets",
                "logo.png",
            ),
            os.path.join(
                os.path.expanduser("~"),
                "OneDrive - Foco Aluguel de Carros",
                "Área de Trabalho",
                "SISTEMA FOCO",
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


def abrir_chrome_debug():
    if porta_esta_aberta(DEBUG_PORT):
        return

    os.makedirs(CHROME_PROFILE_DIR, exist_ok=True)
    comando = [
        encontrar_chrome(),
        f"--remote-debugging-port={DEBUG_PORT}",
        f"--user-data-dir={CHROME_PROFILE_DIR}",
        "--start-maximized",
        LOGIN_URL,
    ]
    subprocess.Popen(comando, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)


def obter_abas_chrome():
    try:
        with urllib.request.urlopen(f"http://127.0.0.1:{DEBUG_PORT}/json", timeout=1) as resposta:
            import json

            return json.loads(resposta.read().decode("utf-8", errors="ignore"))
    except Exception:
        return []


def formatar_tempo(segundos):
    segundos = int(segundos)
    horas, resto = divmod(segundos, 3600)
    minutos, seg = divmod(resto, 60)
    if horas:
        return f"{horas:02d}:{minutos:02d}:{seg:02d}"
    return f"{minutos:02d}:{seg:02d}"


class ControleExecucao:
    def __init__(self):
        self.pausado = threading.Event()
        self.parar = threading.Event()

    def aguardar_liberacao(self):
        while self.pausado.is_set() and not self.parar.is_set():
            time.sleep(0.25)
        if self.parar.is_set():
            raise RuntimeError("Execucao interrompida pelo usuario.")


class CobCloudBaixaPagamentosBot:
    def __init__(
        self,
        controle,
        log_callback,
        status_callback,
        recebimentos_callback,
        boleto_callback,
        pix_callback,
    ):
        self.controle = controle
        self.log = log_callback
        self.status = status_callback
        self.recebimentos_callback = recebimentos_callback
        self.boleto_callback = boleto_callback
        self.pix_callback = pix_callback
        self.driver = None
        self.recebimentos_lancados = 0
        self.baixados_boleto = 0
        self.baixados_pix = 0

    def conectar_selenium_ao_chrome(self):
        options = webdriver.ChromeOptions()
        options.add_experimental_option("debuggerAddress", f"127.0.0.1:{DEBUG_PORT}")
        service = Service(ChromeDriverManager().install())
        self.driver = webdriver.Chrome(service=service, options=options)
        self.driver.maximize_window()

    def wait(self, timeout=30):
        return WebDriverWait(self.driver, timeout)

    def aguardar_pagina_pronta(self, timeout=30):
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
                try:
                    elementos = self.driver.find_elements(By.XPATH, seletor)
                    if any(elemento_visivel_seguro(elemento) for elemento in elementos):
                        carregando = True
                        break
                except StaleElementReferenceException:
                    carregando = True
                    break
            if not carregando:
                return
            time.sleep(0.4)

    def esperar_clicavel(self, xpath, timeout=30):
        self.controle.aguardar_liberacao()
        elemento = self.wait(timeout).until(EC.element_to_be_clickable((By.XPATH, xpath)))
        self.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", elemento)
        return elemento

    def clicar(self, xpath, descricao, timeout=30):
        self.status(f"Clicando: {descricao}")
        self.aguardar_carregamentos(timeout=3)
        elemento = self.esperar_clicavel(xpath, timeout)
        elemento.click()
        time.sleep(0.75)
        self.aguardar_carregamentos(timeout=8)
        return elemento

    def primeiro_clicavel(self, xpaths, timeout=20):
        fim = time.time() + timeout
        ultimo_erro = None
        while time.time() < fim:
            self.controle.aguardar_liberacao()
            for xpath in xpaths:
                try:
                    elementos = self.driver.find_elements(By.XPATH, xpath)
                    for elemento in elementos:
                        if elemento_visivel_seguro(elemento) and elemento_habilitado_seguro(elemento):
                            self.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", elemento)
                            return elemento
                except StaleElementReferenceException as exc:
                    ultimo_erro = exc
                except Exception as exc:
                    ultimo_erro = exc
            time.sleep(0.3)
        raise RuntimeError(f"Elemento clicavel nao encontrado. Ultimo erro: {ultimo_erro}")

    def clicar_elemento(self, elemento, descricao):
        self.status(f"Clicando: {descricao}")
        self.controle.aguardar_liberacao()
        try:
            self.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", elemento)
            elemento.click()
        except StaleElementReferenceException:
            raise
        except Exception:
            self.driver.execute_script("arguments[0].click();", elemento)
        time.sleep(0.75)
        self.aguardar_carregamentos(timeout=8)

    def preencher_input(self, elemento, valor, descricao):
        self.status(f"Preenchendo: {descricao}")
        self.controle.aguardar_liberacao()
        try:
            self.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", elemento)
            elemento.click()
            elemento.send_keys(Keys.CONTROL, "a")
            elemento.send_keys(Keys.BACKSPACE)
            elemento.send_keys(valor)
            self.driver.execute_script(
                "arguments[0].dispatchEvent(new Event('input', {bubbles:true}));"
                "arguments[0].dispatchEvent(new Event('change', {bubbles:true}));",
                elemento,
            )
        except StaleElementReferenceException:
            raise
        time.sleep(0.3)

    def digitar_data_texto_sem_colar(self, elemento, data_texto, descricao):
        self.status(f"Preenchendo: {descricao}")
        self.controle.aguardar_liberacao()
        try:
            self.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", elemento)
            elemento.click()
            teclas = montar_teclas_data_texto(data_texto)
            elemento.send_keys(teclas[0], teclas[1])
            for tecla in teclas[2:]:
                elemento.send_keys(tecla)
                time.sleep(0.03)
            self.driver.execute_script(
                "arguments[0].dispatchEvent(new Event('input', {bubbles:true}));"
                "arguments[0].dispatchEvent(new Event('change', {bubbles:true}));",
                elemento,
            )
            valor_atual = (elemento.get_attribute("value") or "").strip()
            if valor_atual != data_texto:
                self.driver.execute_script(
                    "arguments[0].value = '';"
                    "arguments[0].dispatchEvent(new Event('input', {bubbles:true}));",
                    elemento,
                )
                elemento.click()
                for caractere in data_texto:
                    elemento.send_keys(caractere)
                    time.sleep(0.03)
                self.driver.execute_script(
                    "arguments[0].dispatchEvent(new Event('input', {bubbles:true}));"
                    "arguments[0].dispatchEvent(new Event('change', {bubbles:true}));",
                    elemento,
                )
                valor_atual = (elemento.get_attribute("value") or "").strip()
            if valor_atual != data_texto:
                raise RuntimeError(
                    f"{descricao.capitalize()} nao ficou correta. Esperado {data_texto}, encontrado {valor_atual or 'vazio'}."
                )
        except StaleElementReferenceException:
            raise
        time.sleep(0.3)

    def digitar_data_recebimento_sem_colar(self, elemento, data_recebimento):
        self.digitar_data_texto_sem_colar(elemento, data_recebimento, "data de recebimento")

    def preencher_input_data_html(self, elemento, valor_iso, descricao):
        self.status(f"Preenchendo: {descricao}")
        self.controle.aguardar_liberacao()
        self.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", elemento)
        self.driver.execute_script(
            "arguments[0].value = arguments[1];"
            "arguments[0].dispatchEvent(new Event('input', {bubbles:true}));"
            "arguments[0].dispatchEvent(new Event('change', {bubbles:true}));",
            elemento,
            valor_iso,
        )
        time.sleep(0.3)

    def preencher_data_mes_vigente(self, elemento, descricao):
        tipo_input = ""
        try:
            tipo_input = elemento.get_attribute("type") or ""
        except StaleElementReferenceException:
            raise
        valor = valor_primeiro_dia_mes_para_tipo_input(tipo_input)
        if tipo_input.strip().lower() == "date":
            self.preencher_input_data_html(elemento, valor, descricao)
        else:
            self.digitar_data_texto_sem_colar(elemento, valor, descricao)
        return valor

    def validar_tela_acordos_programados(self):
        self.status("Validando tela")
        xpath_titulo = "/html/body/div[1]/div/div/div[2]/main/div/div/div[1]/div[1]/h3"

        def titulo_valido(driver):
            try:
                titulo = driver.find_element(By.XPATH, xpath_titulo)
                if not elemento_visivel_seguro(titulo):
                    return False
                return titulo.text.strip() == "Acordos Programados"
            except StaleElementReferenceException:
                return False

        self.wait(30).until(titulo_valido)
        self.log("Tela validada: Acordos Programados.")

    def pesquisar_primeiro_dia_mes_vigente(self):
        campo_data_xpath = "/html/body/div[1]/div/div/div[2]/main/div/div/div[2]/div[1]/div[4]/div/div/input"
        botao_pesquisar_xpath = "/html/body/div[1]/div/div/div[2]/main/div/div/div[2]/div[3]/div[2]/button/span[1]"
        campo_data = self.esperar_clicavel(campo_data_xpath, timeout=20)
        data_inicial = self.preencher_data_mes_vigente(campo_data, "data inicial")
        self.log(f"Data inicial informada: {data_inicial}.")
        self.clicar(botao_pesquisar_xpath, "Pesquisar", timeout=20)
        self.aguardar_resultados_acordos()

    def aguardar_resultados_acordos(self, timeout=40):
        self.status("Aguardando resultados")
        self.aguardar_carregamentos(timeout=timeout)
        self.wait(timeout).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, ".table-responsive-material .MuiDataGrid-root"))
        )

        def grade_carregada(driver):
            try:
                if driver.find_elements(By.CSS_SELECTOR, ".MuiDataGrid-row"):
                    return True
                paginacao = driver.find_elements(By.CSS_SELECTOR, ".MuiTablePagination-displayedRows")
                return bool(paginacao and "0-0" in paginacao[0].text)
            except StaleElementReferenceException:
                return False

        self.wait(timeout).until(grade_carregada)
        total = len(self.driver.find_elements(By.CSS_SELECTOR, ".MuiDataGrid-row"))
        self.log(f"Resultados carregados na grade: {total} linhas visiveis.")

    def selecionar_100_registros_por_pagina(self):
        self.status("Ajustando paginação")
        try:
            seletor = self.wait(15).until(
                EC.element_to_be_clickable(
                    (
                        By.XPATH,
                        "//p[normalize-space()='Registros por Página']/following-sibling::div"
                        "//*[@role='combobox']",
                    )
                )
            )
            if seletor.text.strip() == "100":
                self.log("Grade ja configurada para 100 registros por pagina.")
                return
            self.clicar_elemento(seletor, "Registros por pagina")
            opcao_100 = self.wait(10).until(
                EC.element_to_be_clickable((By.XPATH, "//li[@role='option' and @data-value='100']"))
            )
            self.clicar_elemento(opcao_100, "100 registros por pagina")
            self.aguardar_resultados_acordos()
            self.log("Grade configurada para 100 registros por pagina.")
        except Exception as exc:
            raise RuntimeError(f"Nao foi possivel selecionar 100 registros por pagina: {exc}") from exc

    def obter_cor_bolinha_linha(self, linha):
        try:
            bolinhas = linha.find_elements(
                By.XPATH,
                ".//*[@data-field='actions']//button[not(@value)][1]//span[contains(@class,'MuiIconButton-label')]/div",
            )
            if not bolinhas:
                return ""
            return bolinhas[0].value_of_css_property("background-color")
        except StaleElementReferenceException:
            return ""

    def linha_tem_recebimento_pendente(self, linha):
        return eh_bolinha_recebimento_pendente(self.obter_cor_bolinha_linha(linha))

    def obter_dados_linha_acordo(self, linha):
        processo = ""
        forma_pagto = ""
        try:
            processo = linha.find_element(By.XPATH, ".//*[@data-field='processo']//a").text.strip()
        except Exception:
            pass
        try:
            forma_pagto = linha.find_element(By.XPATH, ".//*[@data-field='forma_pagto']").text.strip()
        except Exception:
            pass
        vencimento = linha.find_element(By.XPATH, ".//*[@data-field='data_vencto_br']").text.strip()
        if not vencimento:
            raise RuntimeError(f"Linha {processo or 'sem processo'} esta sem data de vencimento.")
        return processo or "sem processo", vencimento, forma_pagto

    def encontrar_primeira_linha_pendente(self):
        linhas = self.driver.find_elements(By.CSS_SELECTOR, ".MuiDataGrid-row")
        for linha in linhas:
            self.controle.aguardar_liberacao()
            try:
                if self.linha_tem_recebimento_pendente(linha):
                    return linha
            except StaleElementReferenceException:
                continue
        return None

    def abrir_menu_lancamento_recebimento(self, linha):
        try:
            botoes_seta = linha.find_elements(By.XPATH, ".//*[@data-field='actions']//button[@value]")
        except StaleElementReferenceException:
            raise
        if not botoes_seta:
            raise RuntimeError("Setinha de acoes nao encontrada na linha pendente.")
        self.clicar_elemento(botoes_seta[0], "setinha do contrato")
        opcao = self.wait(10).until(
            EC.element_to_be_clickable((By.XPATH, "//li[@role='menuitem' and contains(normalize-space(.),'Recebimento')]"))
        )
        self.clicar_elemento(opcao, "Lancar Recebimento")

    def lancar_recebimento_com_data(self, data_recebimento):
        campo_data_xpath = "/html/body/div[5]/div[3]/div/form/div[2]/div[1]/div[1]/div/div/input"
        gravar_xpath = "/html/body/div[5]/div[3]/div/form/div[3]/button[1]/span[1]"
        campo_data = self.primeiro_clicavel(
            [
                campo_data_xpath,
                "(//div[contains(@class,'MuiDialog-root')]//form//input)[1]",
            ],
            timeout=20,
        )
        self.digitar_data_recebimento_sem_colar(campo_data, data_recebimento)
        botao_gravar = self.primeiro_clicavel(
            [
                gravar_xpath,
                "//div[contains(@class,'MuiDialog-root')]//form//button[contains(normalize-space(.),'Gravar')]",
            ],
            timeout=20,
        )
        self.clicar_elemento(botao_gravar, "Gravar recebimento")
        try:
            WebDriverWait(self.driver, 12).until_not(
                EC.presence_of_element_located((By.XPATH, campo_data_xpath))
            )
        except Exception:
            self.aguardar_carregamentos(timeout=8)

    def processar_recebimentos_pendentes(self):
        self.status("Mapeando recebimentos")
        processados = 0
        while processados < LIMITE_RECEBIMENTOS_POR_EXECUCAO:
            self.controle.aguardar_liberacao()
            self.aguardar_resultados_acordos(timeout=20)
            linha = self.encontrar_primeira_linha_pendente()
            if not linha:
                self.log("Nenhuma bolinha azul pendente encontrada na pagina.")
                return processados

            try:
                processo, vencimento, forma_pagto = self.obter_dados_linha_acordo(linha)
                self.log(f"Recebimento pendente encontrado: {processo} | vencimento {vencimento}.")
                self.abrir_menu_lancamento_recebimento(linha)
                self.lancar_recebimento_com_data(vencimento)
            except StaleElementReferenceException:
                self.log("Grade atualizada durante a leitura da linha; reconsultando resultados.")
                time.sleep(0.5)
                continue

            self.recebimentos_lancados += 1
            self.atualizar_contadores()
            processados += 1
            self.log(f"Recebimento lancado: {processo} | data {vencimento}.")
            if "BOLETO" in forma_pagto.upper():
                self.log(f"Forma de pagamento identificada para proxima etapa: boleto ({processo}).")
            if "PIX" in forma_pagto.upper():
                self.log(f"Forma de pagamento identificada para proxima etapa: pix ({processo}).")
            time.sleep(1)
        raise RuntimeError(
            f"Limite de seguranca atingido: {LIMITE_RECEBIMENTOS_POR_EXECUCAO} recebimentos na mesma execucao."
        )

    def pesquisar_recebimentos_liberados_mes_vigente(self):
        campo_data = self.primeiro_clicavel(
            [
                "//input[@type='date']",
                "/html/body/div[1]/div/div/div[2]/main/div/div/div[2]//input[@type='date']",
            ],
            timeout=25,
        )
        data_inicial = self.preencher_data_mes_vigente(campo_data, "data inicial financeiro")
        self.log(f"Data inicial do financeiro informada: {data_inicial}.")
        self.clicar(
            "/html/body/div[1]/div/div/div[2]/main/div/div/div[2]/div[3]/div[5]/button/span[1]/span",
            "Pesquisar recebimentos liberados",
            timeout=20,
        )

    def obter_texto_tabela_recebimentos_liberados(self, timeout=30):
        self.status("Aguardando recebimentos liberados")
        self.aguardar_carregamentos(timeout=timeout)
        self.wait(timeout).until(EC.presence_of_element_located((By.CSS_SELECTOR, "tbody.MuiTableBody-root")))
        time.sleep(0.5)
        textos = []
        for corpo in self.driver.find_elements(By.CSS_SELECTOR, "tbody.MuiTableBody-root"):
            try:
                texto = corpo.text.strip()
                if texto:
                    textos.append(texto)
            except StaleElementReferenceException:
                return self.obter_texto_tabela_recebimentos_liberados(timeout=10)
        return "\n".join(textos)

    def tem_recebimentos_liberados_para_baixa(self):
        texto_tabela = self.obter_texto_tabela_recebimentos_liberados()
        if texto_indica_recebimentos_liberados(texto_tabela):
            self.log("Recebimentos liberados encontrados para baixa em lote.")
            return True
        self.log("Nenhum recebimento liberado para baixa em lote encontrado.")
        return False

    def confirmar_baixa_em_lote_recebimentos_liberados(self):
        self.status("Baixa em lote")
        self.clicar(
            "/html/body/div[1]/div/div/div[2]/main/div/div/div[1]/div[3]/button/span[2]",
            "Baixar em lote",
            timeout=20,
        )
        selecionar_todos = self.primeiro_clicavel(
            [
                "/html/body/div[4]/div[3]/div/div[2]/table/thead/tr/th[1]/span/span[1]/input",
                "(//div[contains(@class,'MuiDialog-root')]//table//thead//input[@type='checkbox'])[1]",
            ],
            timeout=20,
        )
        self.clicar_elemento(selecionar_todos, "selecionar todos")
        confirmar = self.primeiro_clicavel(
            [
                "/html/body/div[4]/div[3]/div/div[3]/button[1]/span[1]",
                "//div[contains(@class,'MuiDialog-root')]//button[contains(normalize-space(.),'Confirmar')]",
            ],
            timeout=20,
        )
        self.clicar_elemento(confirmar, "Confirmar baixa em lote")
        self.log("Baixa em lote confirmada em Recebimentos Liberados.")

    def processar_recebimentos_liberados_baixa(self):
        self.status("Recebimentos liberados")
        self.driver.get(RECEBIMENTOS_LIBERADOS_URL)
        self.aguardar_carregamentos(timeout=20)
        self.log("Tela Financeiro - Recebimentos Liberados aberta.")
        self.pesquisar_recebimentos_liberados_mes_vigente()
        if self.tem_recebimentos_liberados_para_baixa():
            self.confirmar_baixa_em_lote_recebimentos_liberados()

    def selecionar_banco_do_brasil_boletos(self):
        self.status("Selecionando banco")
        seletor_banco = self.primeiro_clicavel(
            [
                f"//select[option[@value='{BANCO_DO_BRASIL_ID}']]",
                "//select[option[normalize-space()='BANCO DO BRASIL']]",
            ],
            timeout=20,
        )
        self.driver.execute_script(
            "arguments[0].value = arguments[1];"
            "arguments[0].dispatchEvent(new Event('input', {bubbles:true}));"
            "arguments[0].dispatchEvent(new Event('change', {bubbles:true}));",
            seletor_banco,
            BANCO_DO_BRASIL_ID,
        )
        time.sleep(0.3)
        self.log("Banco selecionado: BANCO DO BRASIL.")

    def processar_boletos_bancarios_retorno(self):
        self.status("Boletos bancarios")
        self.driver.get(BOLETOS_BANCARIOS_URL)
        self.aguardar_carregamentos(timeout=20)
        self.log("Tela Financeiro - Boletos Bancarios aberta.")
        campo_data = self.esperar_clicavel(
            "/html/body/div[1]/div/div/div[2]/main/div/div/div[2]/div[1]/div[4]/div/div/input",
            timeout=20,
        )
        data_inicial = self.preencher_data_mes_vigente(campo_data, "data inicial boletos")
        self.log(f"Data inicial de boletos informada: {data_inicial}.")
        self.selecionar_banco_do_brasil_boletos()
        self.clicar(
            "/html/body/div[1]/div/div/div[2]/main/div/div/div[2]/div[3]/div[5]/button/span[1]/span",
            "Pesquisar boletos",
            timeout=20,
        )
        time.sleep(2)
        self.clicar(
            "/html/body/div[1]/div/div/div[2]/main/div/div/div[2]/div[3]/div[4]/button/span[1]/span",
            "Retorno boletos",
            timeout=20,
        )
        self.log("Retorno de boletos acionado para BANCO DO BRASIL.")

    def aguardar_login_manual(self, timeout=300):
        self.status("Aguardando login manual")
        self.log("Chrome aberto no CobCloud. Faça login manualmente e aguarde o robô detectar a sessão.")
        fim = time.time() + timeout
        while time.time() < fim:
            self.controle.aguardar_liberacao()
            abas = obter_abas_chrome()
            for aba in abas:
                url = aba.get("url", "")
                if "foco.cobcloud.com.br/app" in url:
                    self.log("Login CobCloud detectado.")
                    return
            time.sleep(1)
        raise RuntimeError("Tempo esgotado aguardando login manual no CobCloud.")

    def executar(self):
        inicio = time.time()
        try:
            self.atualizar_contadores()

            self.status("Abrindo CobCloud")
            abrir_chrome_debug()
            self.aguardar_login_manual()
            self.conectar_selenium_ao_chrome()
            self.aguardar_carregamentos(timeout=20)

            self.status("Acordos programados")
            self.driver.get(BASE_URL)
            self.aguardar_carregamentos(timeout=20)
            self.log("Base de login e cliques seguros carregada.")
            self.validar_tela_acordos_programados()
            self.pesquisar_primeiro_dia_mes_vigente()
            self.selecionar_100_registros_por_pagina()
            total = self.processar_recebimentos_pendentes()
            self.log(f"Processamento de recebimentos concluido. Lancamentos realizados: {total}.")
            self.processar_recebimentos_liberados_baixa()
            self.processar_boletos_bancarios_retorno()
        finally:
            self.status("Finalizado")
            self.log(f"Tempo de execucao: {formatar_tempo(time.time() - inicio)}")
            if self.driver:
                try:
                    self.driver.quit()
                except WebDriverException:
                    pass

    def atualizar_contadores(self):
        self.recebimentos_callback(self.recebimentos_lancados)
        self.boleto_callback(self.baixados_boleto)
        self.pix_callback(self.baixados_pix)


class RoboCobCloudBaixaPagamentosApp:
    def __init__(self):
        ctk.set_appearance_mode("light")
        ctk.set_default_color_theme("blue")

        self.root = ctk.CTk()
        self.root.title("Robo CobCloud - Baixa de Pagamentos")
        self.root.geometry("1040x760")
        self.root.minsize(920, 680)
        self.root.configure(fg_color=MAIN_BG)

        self.status_var = tk.StringVar(value="Aguardando inicio")
        self.recebimentos_var = tk.StringVar(value="0")
        self.boleto_var = tk.StringVar(value="0")
        self.pix_var = tk.StringVar(value="0")
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
        ctk.CTkLabel(texto, text="Baixa de Pagamentos", text_color=PRIMARY_TEXT, font=("Segoe UI", 30, "bold")).pack(anchor="w")
        ctk.CTkLabel(
            texto,
            text="Automacao CobCloud para lancar e baixar recebimentos direto pela plataforma.",
            text_color=MUTED_TEXT,
            font=("Segoe UI", 14),
        ).pack(anchor="w", pady=(6, 0))
        ctk.CTkLabel(texto, text="COBCLOUD - BACKOFFICE", text_color="#a65f56", font=("Segoe UI", 12, "bold")).pack(
            anchor="w", pady=(10, 0)
        )

        progresso = self.criar_secao(scroll, "Progresso da Execucao")
        cards = ctk.CTkFrame(progresso, fg_color="transparent")
        cards.pack(fill="x", padx=18, pady=(0, 18))
        cards.grid_columnconfigure((0, 1, 2, 3), weight=1)
        self.criar_indicador(cards, "Lancados", self.recebimentos_var, 0)
        self.criar_indicador(cards, "Boleto", self.boleto_var, 1)
        self.criar_indicador(cards, "Pix", self.pix_var, 2)
        self.criar_indicador(cards, "Tempo", self.tempo_var, 3)

        status_card = self.criar_secao(scroll, "Status")
        status_linha = ctk.CTkFrame(status_card, fg_color="#fffaf9", corner_radius=16, border_width=1, border_color=CARD_BORDER)
        status_linha.pack(fill="x", padx=18, pady=(0, 18))
        ctk.CTkLabel(status_linha, textvariable=self.status_var, text_color=PRIMARY_TEXT, font=("Segoe UI", 18, "bold")).pack(
            anchor="w", padx=18, pady=16
        )

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
            height=250,
            corner_radius=16,
            fg_color="#fffaf9",
            border_width=1,
            border_color=CARD_BORDER,
            text_color="#2d2d2d",
            font=("Consolas", 12),
        )
        self.log_text.pack(fill="both", expand=True, padx=18, pady=(0, 18))
        self.log_text.configure(state="disabled")

        ctk.CTkLabel(
            scroll,
            text="Desenvolvido por Diogo Medeiros © 2026",
            text_color="#b85b52",
            font=("Segoe UI", 11),
        ).pack(anchor="w", padx=12, pady=(0, 12))

    def criar_indicador(self, parent, titulo, variavel, coluna):
        frame = ctk.CTkFrame(parent, fg_color="#fffaf9", corner_radius=16, border_width=1, border_color=CARD_BORDER)
        frame.grid(row=0, column=coluna, sticky="ew", padx=6)
        ctk.CTkLabel(frame, text=titulo, text_color=MUTED_TEXT, font=("Segoe UI", 12, "bold")).pack(pady=(14, 2))
        ctk.CTkLabel(frame, textvariable=variavel, text_color=PRIMARY_TEXT, font=("Segoe UI", 22, "bold")).pack(pady=(0, 14))

    def criar_botao_primario(self, parent, texto, comando):
        return ctk.CTkButton(
            parent,
            text=texto,
            command=comando,
            height=48,
            width=170,
            corner_radius=14,
            fg_color=BUTTON_BG,
            hover_color=BUTTON_ACTIVE_BG,
            font=("Segoe UI", 16, "bold"),
        )

    def criar_botao_secundario(self, parent, texto, comando, largura=140):
        return ctk.CTkButton(
            parent,
            text=texto,
            command=comando,
            height=48,
            width=largura,
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

    def set_recebimentos(self, valor):
        self.log_queue.put(("recebimentos", valor))

    def set_boleto(self, valor):
        self.log_queue.put(("boleto", valor))

    def set_pix(self, valor):
        self.log_queue.put(("pix", valor))

    def processar_filas(self):
        while not self.log_queue.empty():
            item = self.log_queue.get()
            if isinstance(item, tuple) and item[0] == "status":
                self.status_var.set(item[1])
                continue
            if isinstance(item, tuple) and item[0] == "contador":
                self.recebimentos_var.set(str(item[1]))
                continue
            if isinstance(item, tuple) and item[0] == "recebimentos":
                self.recebimentos_var.set(str(item[1]))
                continue
            if isinstance(item, tuple) and item[0] == "boleto":
                self.boleto_var.set(str(item[1]))
                continue
            if isinstance(item, tuple) and item[0] == "pix":
                self.pix_var.set(str(item[1]))
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

        self.controle = ControleExecucao()
        self.inicio_execucao = time.time()
        self.recebimentos_var.set("0")
        self.boleto_var.set("0")
        self.pix_var.set("0")
        self.tempo_var.set("00:00")
        self.status_var.set("Iniciando")
        self.btn_iniciar.configure(state="disabled")
        self.btn_pausar.configure(text="Pausar")
        self.log_text.configure(state="normal")
        self.log_text.delete("1.0", "end")
        self.log_text.configure(state="disabled")

        self.thread = threading.Thread(target=self.executar_thread, daemon=True)
        self.thread.start()

    def executar_thread(self):
        try:
            bot = CobCloudBaixaPagamentosBot(
                controle=self.controle,
                log_callback=self.log,
                status_callback=self.set_status,
                recebimentos_callback=self.set_recebimentos,
                boleto_callback=self.set_boleto,
                pix_callback=self.set_pix,
            )
            bot.executar()
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
    app = RoboCobCloudBaixaPagamentosApp()
    app.root.mainloop()
