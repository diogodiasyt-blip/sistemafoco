import os
import queue
import sys
import threading
import time
from datetime import datetime
import tkinter as tk
from tkinter import filedialog, messagebox

import customtkinter as ctk
import multiprocessing
import pandas as pd
import winsound
from PIL import Image
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
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
LINK_BLUE = "#2f64d6"


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


def get_desktop_path():
    try:
        import winreg

        key = winreg.OpenKey(
            winreg.HKEY_CURRENT_USER,
            r"Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders",
        )
        desktop = winreg.QueryValueEx(key, "Desktop")[0]
        winreg.CloseKey(key)
        return desktop
    except Exception:
        return os.path.join(os.path.expanduser("~"), "Desktop")


def iniciar_driver(pasta_download, headless, log_callback):
    status = "INVISIVEL" if headless else "VISIVEL"
    log_callback(f"Iniciando Chrome ({status}) - Pasta: {pasta_download}")

    options = webdriver.ChromeOptions()
    if headless:
        options.add_argument("--headless=new")
        options.add_argument("--window-size=1920,1080")
        options.add_argument("--disable-gpu")

    prefs = {
        "download.default_directory": pasta_download,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "plugins.always_open_pdf_externally": True,
    }
    options.add_experimental_option("prefs", prefs)
    options.add_argument("--disable-blink-features=AutomationControlled")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")

    service = Service(ChromeDriverManager().install())
    return webdriver.Chrome(service=service, options=options)


def login(driver, usuario, senha, log_callback):
    log_callback("Fazendo login...")
    driver.get("https://coral.aluguefoco.com.br/login")
    wait = WebDriverWait(driver, 20)

    seletores_usuario = [
        (By.XPATH, '//input[@placeholder="Usuário"]'),
        (By.XPATH, "/html/body/foco-app/div[1]/foco-login/div/div/div/div/div[2]/form/div[1]/input"),
        (By.XPATH, '//input[@placeholder="Usuario"]'),
        (By.XPATH, '//input[@placeholder="Usuário"]'),
        (By.XPATH, '//input[@type="text"]'),
    ]
    seletores_senha = [
        (By.XPATH, "/html/body/foco-app/div[1]/foco-login/div/div/div/div/div[2]/form/div[2]/input"),
        (By.XPATH, '//input[@placeholder="Senha"]'),
        (By.XPATH, '//input[@type="password"]'),
    ]

    usuario_field = None
    for by, seletor in seletores_usuario:
        try:
            usuario_field = wait.until(EC.presence_of_element_located((by, seletor)))
            break
        except Exception:
            continue
    if usuario_field is None:
        raise RuntimeError("Nao foi possivel localizar o campo de usuario na tela de login.")

    senha_field = None
    for by, seletor in seletores_senha:
        try:
            senha_field = wait.until(EC.presence_of_element_located((by, seletor)))
            break
        except Exception:
            continue
    if senha_field is None:
        raise RuntimeError("Nao foi possivel localizar o campo de senha na tela de login.")

    usuario_field.clear()
    usuario_field.send_keys(usuario)
    senha_field.clear()
    senha_field.send_keys(senha)
    senha_field.send_keys(Keys.ENTER)
    time.sleep(5)
    log_callback("Login realizado")


def esperar_download(pasta_download, contrato, log_callback, timeout=60):
    inicio = time.time()
    log_callback(f"Aguardando download do contrato {contrato}...")

    while time.time() - inicio < timeout:
        arquivos = [f for f in os.listdir(pasta_download) if f.endswith(".pdf")]
        if arquivos:
            ultimo = max(arquivos, key=lambda x: os.path.getctime(os.path.join(pasta_download, x)))
            log_callback(f"{contrato} -> PDF baixado -> {ultimo}")
            return True
        time.sleep(1.0)

    log_callback(f"{contrato} -> Nenhum PDF detectado")
    return False


def buscar(driver, numero, log_callback):
    wait = WebDriverWait(driver, 15)
    try:
        driver.find_element(By.XPATH, "/html/body/foco-app/div[1]/div/ul/li[5]/a/i").click()
        time.sleep(1)
        driver.find_element(By.XPATH, '//*[@id="tab-ra-list"]').click()
        time.sleep(1.5)
        campo = driver.find_element(
            By.XPATH,
            "/html/body/foco-app/div[1]/foco-rent-agreement-home/div/div/div[2]/input",
        )
        campo.clear()
        campo.send_keys(numero)
        campo.send_keys(Keys.ENTER)
        wait.until(
            EC.presence_of_element_located(
                (By.XPATH, f"//*[@id='tableCRUD']/tbody/tr/td[2][contains(text(), '{numero}')]")
            )
        )
        log_callback(f"{numero} encontrado")
        return True
    except Exception:
        log_callback(f"{numero} nao encontrado")
        return False


def baixar(driver, log_callback):
    wait = WebDriverWait(driver, 20)
    log_callback("Iniciando download...")
    wait.until(EC.element_to_be_clickable((By.XPATH, "//i[contains(@class,'ellipsis')]"))).click()
    time.sleep(1)
    wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(.,'Reenviar')]"))).click()
    time.sleep(1)
    wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(.,'Baixar PDF')]"))).click()
    time.sleep(1)
    botao_idioma = None
    for seletor in [
        "//button[contains(.,'Português')]",
        "//button[contains(.,'Portugues')]",
    ]:
        try:
            botao_idioma = wait.until(EC.element_to_be_clickable((By.XPATH, seletor)))
            break
        except Exception:
            continue
    if botao_idioma is None:
        raise RuntimeError("Nao foi possivel localizar o botao de idioma para baixar o PDF.")
    botao_idioma.click()
    time.sleep(3)

    if len(driver.window_handles) > 1:
        driver.switch_to.window(driver.window_handles[-1])
        time.sleep(5)
        driver.close()
        driver.switch_to.window(driver.window_handles[0])

    wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(.,'Fechar')]"))).click()
    log_callback("Download enviado")


def executar_robo(usuario, senha, planilha_path, pasta_download, headless, log_callback, queue_result, progress_callback):
    start_time = time.time()
    success = []
    failed = []
    processed = set()
    driver = None

    log_callback("Validando pasta de salvamento...")
    try:
        os.makedirs(pasta_download, exist_ok=True)
        log_callback(f"Pasta pronta: {pasta_download}")
    except Exception as e:
        log_callback(f"Erro ao criar pasta: {str(e)}")
        return

    try:
        driver = iniciar_driver(pasta_download, headless, log_callback)
        login(driver, usuario, senha, log_callback)

        df = pd.read_excel(planilha_path)
        contratos = df["contrato"].dropna().astype(str).tolist()
        total = len(contratos)
        log_callback(f"Iniciando processamento de {total} contratos...")

        for i, contrato in enumerate(contratos, 1):
            if contrato in processed:
                continue
            processed.add(contrato)

            log_callback(f"\n[{i}/{total}] Processando: {contrato}")
            try:
                if not buscar(driver, contrato, log_callback):
                    failed.append(contrato)
                    continue

                baixar(driver, log_callback)

                if esperar_download(pasta_download, contrato, log_callback, timeout=65):
                    success.append(contrato)
                else:
                    failed.append(contrato)

                progress_callback((i / total) * 100)

                driver.get("https://coral.aluguefoco.com.br/dashboard")
                time.sleep(2.0)

            except Exception as e:
                log_callback(f"Erro em {contrato}: {str(e)}")
                failed.append(contrato)

    except Exception as e:
        log_callback(f"ERRO FATAL: {str(e)}")
    finally:
        if driver:
            try:
                driver.quit()
            except Exception:
                pass

    elapsed = time.time() - start_time
    minutes = int(elapsed // 60)
    seconds = int(elapsed % 60)

    queue_result.put({"success": success, "failed": failed, "time": f"{minutes}min {seconds}s"})
    log_callback(f"\nFINALIZADO em {minutes}min {seconds}s -> Sucessos: {len(success)} | Erros: {len(failed)}")

    try:
        winsound.Beep(1000, 800)
        time.sleep(0.3)
        winsound.Beep(1200, 600)
    except Exception:
        pass


class RoboContratosApp:
    def __init__(self):
        ctk.set_appearance_mode("light")
        ctk.set_default_color_theme("blue")

        self.root = ctk.CTk()
        self.root.title("Robo de Contratos Coral - Desenvolvido por Diogo Medeiros © 2026")
        self.root.geometry("1120x820")
        self.root.minsize(980, 720)
        self.root.configure(fg_color=MAIN_BG)

        desktop = get_desktop_path()
        pasta_padrao = os.path.join(desktop, "Contratos_Foco")

        self.usuario_var = tk.StringVar(value="")
        self.senha_var = tk.StringVar(value="")
        self.planilha_var = tk.StringVar(value="")
        self.pasta_var = tk.StringVar(value=pasta_padrao)
        self.headless_var = tk.BooleanVar(value=True)

        self.log_queue = queue.Queue()
        self.result_queue = queue.Queue()
        self.logo_image = None
        self.logo_label = None

        self.create_widgets()

    def carregar_logo(self, reducao=2):
        caminho_logo = localizar_logo()
        if not caminho_logo:
            return None
        try:
            imagem = Image.open(caminho_logo)
            largura, altura = imagem.size
            if reducao > 1:
                largura = max(1, largura // reducao)
                altura = max(1, altura // reducao)
            logo = ctk.CTkImage(light_image=imagem, dark_image=imagem, size=(largura, altura))
            self.logo_image = logo
            return logo
        except Exception:
            return None

    def criar_secao(self, parent, titulo):
        frame = ctk.CTkFrame(
            parent,
            fg_color=CARD_BG,
            corner_radius=20,
            border_width=1,
            border_color=CARD_BORDER,
        )
        frame.pack(fill="x", padx=8, pady=8)
        ctk.CTkLabel(
            frame,
            text=titulo,
            text_color=PRIMARY_TEXT,
            font=("Segoe UI", 18, "bold"),
        ).pack(anchor="w", padx=18, pady=(16, 12))
        return frame

    def create_widgets(self):
        container = ctk.CTkFrame(self.root, fg_color=MAIN_BG, corner_radius=0)
        container.pack(fill="both", expand=True, padx=12, pady=12)

        scroll = ctk.CTkScrollableFrame(container, fg_color=MAIN_BG, corner_radius=0)
        scroll.pack(fill="both", expand=True)

        hero = ctk.CTkFrame(
            scroll,
            fg_color=CARD_BG,
            corner_radius=26,
            border_width=1,
            border_color=CARD_BORDER,
        )
        hero.pack(fill="x", padx=8, pady=(8, 14))

        hero_inner = ctk.CTkFrame(hero, fg_color="transparent")
        hero_inner.pack(fill="x", padx=24, pady=24)

        logo = self.carregar_logo(reducao=2)
        if logo:
            self.logo_label = ctk.CTkLabel(hero_inner, text="", image=logo)
            self.logo_label.pack(side="left", padx=(0, 18))

        texto = ctk.CTkFrame(hero_inner, fg_color="transparent")
        texto.pack(side="left", fill="x", expand=True)

        ctk.CTkLabel(
            texto,
            text="Contratos FOCO",
            text_color=PRIMARY_TEXT,
            font=("Segoe UI", 30, "bold"),
        ).pack(anchor="w")
        ctk.CTkLabel(
            texto,
            text="Busca, reenvio e download de contratos com progresso em tempo real.",
            text_color=MUTED_TEXT,
            font=("Segoe UI", 14),
        ).pack(anchor="w", pady=(6, 0))
        ctk.CTkLabel(
            texto,
            text="GESTAO DE CONTRATOS",
            text_color="#a65f56",
            font=("Segoe UI", 12, "bold"),
        ).pack(anchor="w", pady=(10, 0))

        acesso = self.criar_secao(scroll, "Acesso")
        acesso_grid = ctk.CTkFrame(acesso, fg_color="transparent")
        acesso_grid.pack(fill="x", padx=18, pady=(0, 18))
        acesso_grid.grid_columnconfigure((0, 1), weight=1)

        ctk.CTkLabel(acesso_grid, text="Usuario", font=("Segoe UI", 13, "bold"), text_color="#303030").grid(row=0, column=0, sticky="w", padx=(0, 10), pady=(0, 6))
        ctk.CTkLabel(acesso_grid, text="Senha", font=("Segoe UI", 13, "bold"), text_color="#303030").grid(row=0, column=1, sticky="w", padx=(10, 0), pady=(0, 6))

        self.entry_usuario = ctk.CTkEntry(acesso_grid, textvariable=self.usuario_var, height=42, corner_radius=12)
        self.entry_usuario.grid(row=1, column=0, sticky="ew", padx=(0, 10))
        self.entry_senha = ctk.CTkEntry(acesso_grid, textvariable=self.senha_var, show="*", height=42, corner_radius=12)
        self.entry_senha.grid(row=1, column=1, sticky="ew", padx=(10, 0))

        config = self.criar_secao(scroll, "Configuracoes")
        self.check_headless = ctk.CTkCheckBox(
            config,
            text="Executar em modo invisivel (sem abrir janela do Chrome)",
            variable=self.headless_var,
            font=("Segoe UI", 13),
            text_color="#303030",
            checkbox_width=22,
            checkbox_height=22,
            corner_radius=8,
        )
        self.check_headless.pack(anchor="w", padx=18, pady=(0, 18))

        arquivos = self.criar_secao(scroll, "Planilha e Pasta")
        botoes = ctk.CTkFrame(arquivos, fg_color="transparent")
        botoes.pack(fill="x", padx=18, pady=(0, 12))
        botoes.grid_columnconfigure((0, 1), weight=1)

        ctk.CTkButton(
            botoes,
            text="Selecionar Planilha Excel",
            command=self.selecionar_planilha,
            height=44,
            corner_radius=14,
            fg_color=BUTTON_BG,
            hover_color=BUTTON_ACTIVE_BG,
            font=("Segoe UI", 14, "bold"),
        ).grid(row=0, column=0, sticky="ew", padx=(0, 8))

        ctk.CTkButton(
            botoes,
            text="Mudar Pasta de Salvamento",
            command=self.selecionar_pasta,
            height=44,
            corner_radius=14,
            fg_color="#ffffff",
            text_color=PRIMARY_TEXT,
            hover_color=SOFT_RED,
            border_width=1,
            border_color="#f0d7d2",
            font=("Segoe UI", 14, "bold"),
        ).grid(row=0, column=1, sticky="ew", padx=(8, 0))

        self.label_planilha = ctk.CTkLabel(
            arquivos,
            text="Nenhuma planilha selecionada",
            text_color=LINK_BLUE,
            font=("Segoe UI", 12),
            justify="left",
            anchor="w",
        )
        self.label_planilha.pack(fill="x", padx=18, pady=(0, 6))

        self.label_pasta = ctk.CTkLabel(
            arquivos,
            text=self.pasta_var.get(),
            text_color=LINK_BLUE,
            font=("Segoe UI", 12),
            justify="left",
            anchor="w",
        )
        self.label_pasta.pack(fill="x", padx=18, pady=(0, 18))

        progresso = self.criar_secao(scroll, "Progresso da Execucao")
        self.progress_bar = ctk.CTkProgressBar(
            progresso,
            height=16,
            corner_radius=999,
            progress_color=BUTTON_BG,
            fg_color="#f2dfdb",
        )
        self.progress_bar.pack(fill="x", padx=18, pady=(0, 10))
        self.progress_bar.set(0)

        self.label_progress = ctk.CTkLabel(
            progresso,
            text="0% - Aguardando inicio...",
            text_color=MUTED_TEXT,
            font=("Segoe UI", 13),
        )
        self.label_progress.pack(pady=(0, 18))

        acoes = ctk.CTkFrame(scroll, fg_color="transparent")
        acoes.pack(fill="x", padx=8, pady=(0, 8))

        self.btn_iniciar = ctk.CTkButton(
            acoes,
            text="Iniciar Robo",
            command=self.iniciar_robo,
            height=48,
            width=180,
            corner_radius=14,
            fg_color=BUTTON_BG,
            hover_color=BUTTON_ACTIVE_BG,
            font=("Segoe UI", 16, "bold"),
        )
        self.btn_iniciar.pack(side="left", padx=(0, 10))

        self.btn_salvar_log = ctk.CTkButton(
            acoes,
            text="Salvar Log",
            command=self.salvar_log,
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
        self.btn_salvar_log.pack(side="left")

        logs = self.criar_secao(scroll, "Logs em Tempo Real")
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

        ctk.CTkLabel(
            scroll,
            text="Desenvolvido por Diogo Medeiros © 2026",
            text_color="#b85b52",
            font=("Segoe UI", 11),
        ).pack(anchor="w", padx=12, pady=(0, 12))

        self.root.after(100, self.process_queue)

    def log(self, msg):
        self.log_text.configure(state="normal")
        self.log_text.insert("end", msg + "\n")
        self.log_text.see("end")
        self.log_text.configure(state="disabled")

    def process_queue(self):
        while not self.log_queue.empty():
            msg = self.log_queue.get()
            self.log(msg)
        if not self.result_queue.empty():
            self.mostrar_resumo(self.result_queue.get())
        self.root.after(100, self.process_queue)

    def atualizar_progresso(self, valor):
        progresso = max(0.0, min(1.0, valor / 100))
        self.progress_bar.set(progresso)
        self.label_progress.configure(text=f"{int(valor)}% concluido")

    def selecionar_planilha(self):
        arq = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx *.xls")])
        if arq:
            self.planilha_var.set(arq)
            self.label_planilha.configure(text=arq)

    def selecionar_pasta(self):
        pasta = filedialog.askdirectory()
        if pasta:
            self.pasta_var.set(pasta)
            self.label_pasta.configure(text=pasta)

    def iniciar_robo(self):
        if not self.planilha_var.get():
            messagebox.showwarning("Atencao", "Selecione a planilha Excel!")
            return

        self.btn_iniciar.configure(state="disabled")
        self.progress_bar.set(0)
        self.label_progress.configure(text="0% - Iniciando...")
        self.log_text.configure(state="normal")
        self.log_text.delete("1.0", "end")
        self.log_text.configure(state="disabled")

        thread = threading.Thread(
            target=executar_robo,
            args=(
                self.usuario_var.get(),
                self.senha_var.get(),
                self.planilha_var.get(),
                self.pasta_var.get(),
                self.headless_var.get(),
                self.log,
                self.result_queue,
                self.atualizar_progresso,
            ),
            daemon=True,
        )
        thread.start()

    def mostrar_resumo(self, result):
        s = result["success"]
        f = result["failed"]
        tempo = result.get("time", "N/A")

        messagebox.showinfo(
            "Execucao Finalizada",
            f"Sucessos: {len(s)}\nErros: {len(f)}\nTempo total: {tempo}\n\n"
            + "Sucessos:\n"
            + "\n".join(s)
            + "\n\nErros:\n"
            + "\n".join(f),
        )

        self.btn_iniciar.configure(state="normal")
        self.progress_bar.set(1)
        self.label_progress.configure(text="100% - Finalizado")

    def salvar_log(self):
        try:
            log_content = self.log_text.get("1.0", "end")
            filename = f"Log_Robo_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
            with open(filename, "w", encoding="utf-8") as arquivo:
                arquivo.write(log_content)
            messagebox.showinfo("Salvo", f"Log salvo com sucesso!\nArquivo: {filename}")
        except Exception as e:
            messagebox.showerror("Erro", f"Nao foi possivel salvar o log:\n{str(e)}")


if __name__ == "__main__":
    multiprocessing.freeze_support()
    app = RoboContratosApp()
    app.root.mainloop()
