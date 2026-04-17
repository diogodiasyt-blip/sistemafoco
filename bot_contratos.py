import sys
import os
import tkinter as tk
from tkinter import ttk, filedialog, scrolledtext, messagebox
import threading
import queue
import time
import pandas as pd
import multiprocessing
from datetime import datetime
import winsound  # Para notificação sonora

# ====================== CONFIGURAÇÕES PARA .EXE ======================
if getattr(sys, 'frozen', False):
    os.environ['WDM_LOG_LEVEL'] = '0'

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager

print("🚀 ROBÔ INICIADO")

MAIN_BG = "#f6f4f1"
CARD_BG = "#ffffff"
PRIMARY_TEXT = "#d81919"
MUTED_TEXT = "#5c5c5c"
BUTTON_BG = "#ef1a14"
BUTTON_ACTIVE_BG = "#c91410"


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
        key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders")
        desktop = winreg.QueryValueEx(key, "Desktop")[0]
        winreg.CloseKey(key)
        return desktop
    except:
        return os.path.join(os.path.expanduser("~"), "Desktop")


# ====================== FUNÇÕES DO ROBÔ ======================
def iniciar_driver(pasta_download, headless, log_callback):
    status = "INVISIBLE" if headless else "VISÍVEL"
    log_callback(f"🚀 Iniciando Chrome ({status}) - Pasta: {pasta_download}")
    
    options = webdriver.ChromeOptions()
    
    if headless:
        options.add_argument("--headless=new")
        options.add_argument("--window-size=1920,1080")
        options.add_argument("--disable-gpu")

    prefs = {
        "download.default_directory": pasta_download,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "plugins.always_open_pdf_externally": True
    }
    options.add_experimental_option("prefs", prefs)
    options.add_argument("--disable-blink-features=AutomationControlled")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")

    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=options)
    return driver


def login(driver, usuario, senha, log_callback):
    log_callback("🔑 Fazendo login...")
    driver.get("https://coral.aluguefoco.com.br/login")
    time.sleep(3)
    driver.find_element(By.XPATH, '//input[@placeholder="Usuário"]').send_keys(usuario)
    senha_field = driver.find_element(By.XPATH, '//input[@placeholder="Senha"]')
    senha_field.send_keys(senha)
    senha_field.send_keys(Keys.ENTER)
    time.sleep(5)
    log_callback("✅ Login realizado")


def esperar_download(pasta_download, contrato, log_callback, timeout=60):
    inicio = time.time()
    log_callback(f"⏳ Aguardando download do contrato {contrato}...")

    while time.time() - inicio < timeout:
        arquivos = [f for f in os.listdir(pasta_download) if f.endswith(".pdf")]
        if arquivos:
            ultimo = max(arquivos, key=lambda x: os.path.getctime(os.path.join(pasta_download, x)))
            log_callback(f"✅ {contrato} → PDF baixado → {ultimo}")
            return True
        time.sleep(1.0)

    log_callback(f"❌ {contrato} → Nenhum PDF detectado")
    return False


def buscar(driver, numero, log_callback):
    wait = WebDriverWait(driver, 15)
    try:
        driver.find_element(By.XPATH, "/html/body/foco-app/div[1]/div/ul/li[5]/a/i").click()
        time.sleep(1)
        driver.find_element(By.XPATH, '//*[@id="tab-ra-list"]').click()
        time.sleep(1.5)
        campo = driver.find_element(By.XPATH, '/html/body/foco-app/div[1]/foco-rent-agreement-home/div/div/div[2]/input')
        campo.clear()
        campo.send_keys(numero)
        campo.send_keys(Keys.ENTER)
        wait.until(EC.presence_of_element_located((By.XPATH, f"//*[@id='tableCRUD']/tbody/tr/td[2][contains(text(), '{numero}')]")))
        log_callback(f"✔ {numero} encontrado")
        return True
    except:
        log_callback(f"❌ {numero} não encontrado")
        return False


def baixar(driver, log_callback):
    wait = WebDriverWait(driver, 20)
    log_callback("📥 Iniciando download...")
    wait.until(EC.element_to_be_clickable((By.XPATH, "//i[contains(@class,'ellipsis')]"))).click()
    time.sleep(1)
    wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(.,'Reenviar')]"))).click()
    time.sleep(1)
    wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(.,'Baixar PDF')]"))).click()
    time.sleep(1)
    wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(.,'Português')]"))).click()
    time.sleep(3)

    if len(driver.window_handles) > 1:
        driver.switch_to.window(driver.window_handles[-1])
        time.sleep(5)
        driver.close()
        driver.switch_to.window(driver.window_handles[0])

    wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(.,'Fechar')]"))).click()
    log_callback("✅ Download enviado")


# ====================== EXECUÇÃO ======================
def executar_robo(usuario, senha, planilha_path, pasta_download, headless, log_callback, queue_result, progress_callback):
    start_time = time.time()
    success = []
    failed = []
    processed = set()

    log_callback("📂 Validando pasta de salvamento...")
    try:
        os.makedirs(pasta_download, exist_ok=True)
        log_callback(f"✅ Pasta pronta: {pasta_download}")
    except Exception as e:
        log_callback(f"❌ Erro ao criar pasta: {str(e)}")
        return

    try:
        driver = iniciar_driver(pasta_download, headless, log_callback)
        login(driver, usuario, senha, log_callback)

        df = pd.read_excel(planilha_path)
        contratos = df["contrato"].dropna().astype(str).tolist()
        total = len(contratos)

        log_callback(f"📋 Iniciando processamento de {total} contratos...")

        for i, c in enumerate(contratos, 1):
            if c in processed: continue
            processed.add(c)

            log_callback(f"\n[{i}/{total}] Processando: {c}")

            try:
                if not buscar(driver, c, log_callback):
                    failed.append(c)
                    continue

                baixar(driver, log_callback)

                if esperar_download(pasta_download, c, log_callback, timeout=65):
                    success.append(c)
                else:
                    failed.append(c)

                # Atualiza barra de progresso
                progress = (i / total) * 100
                progress_callback(progress)

                driver.get("https://coral.aluguefoco.com.br/dashboard")
                time.sleep(2.0)

            except Exception as e:
                log_callback(f"❌ Erro em {c}: {str(e)}")
                failed.append(c)

        driver.quit()

    except Exception as e:
        log_callback(f"❌ ERRO FATAL: {str(e)}")

    # Tempo total
    elapsed = time.time() - start_time
    minutes = int(elapsed // 60)
    seconds = int(elapsed % 60)

    queue_result.put({"success": success, "failed": failed, "time": f"{minutes}min {seconds}s"})
    log_callback(f"\n🏁 FINALIZADO em {minutes}min {seconds}s → Sucessos: {len(success)} | Erros: {len(failed)}")

    # Notificação sonora
    try:
        winsound.Beep(1000, 800)   # Beep médio
        time.sleep(0.3)
        winsound.Beep(1200, 600)
    except:
        pass


# ====================== INTERFACE ======================
class RoboContratosApp:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("🤖 Robô de Contratos Coral - Desenvolvido por Diogo © 2026")
        self.root.geometry("1000x760")
        self.root.configure(bg=MAIN_BG)

        desktop = get_desktop_path()
        pasta_padrao = os.path.join(desktop, "Contratos_Foco")

        self.usuario_var = tk.StringVar(value="")
        self.senha_var = tk.StringVar(value="")
        self.planilha_var = tk.StringVar()
        self.pasta_var = tk.StringVar(value=pasta_padrao)
        self.headless_var = tk.BooleanVar(value=True)

        self.log_queue = queue.Queue()
        self.result_queue = queue.Queue()
        self.progress_var = tk.DoubleVar(value=0)
        self.logo_image = None

        self.configurar_estilo()
        self.create_widgets()

    def configurar_estilo(self):
        style = ttk.Style()
        try:
            style.theme_use("clam")
        except Exception:
            pass
        style.configure("App.TFrame", background=MAIN_BG)
        style.configure("TLabelframe", background=CARD_BG, borderwidth=1, relief="solid")
        style.configure("TLabelframe.Label", background=CARD_BG, foreground=PRIMARY_TEXT, font=("Segoe UI", 10, "bold"))
        style.configure("TLabel", background=CARD_BG, foreground="#303030", font=("Segoe UI", 10))
        style.configure("Primary.TButton", background=BUTTON_BG, foreground="#ffffff", padding=(14, 8), font=("Segoe UI", 10, "bold"), borderwidth=0)
        style.map("Primary.TButton", background=[("active", BUTTON_ACTIVE_BG), ("pressed", BUTTON_ACTIVE_BG)])
        style.configure("Secondary.TButton", background="#ffffff", foreground=PRIMARY_TEXT, padding=(12, 8), font=("Segoe UI", 10, "bold"), borderwidth=1)
        style.map("Secondary.TButton", background=[("active", "#fff3f2")], foreground=[("active", PRIMARY_TEXT)])
        style.configure("Accent.Horizontal.TProgressbar", troughcolor="#f3e6e3", background=BUTTON_BG, bordercolor="#f3e6e3", lightcolor=BUTTON_BG, darkcolor=BUTTON_BG)

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

    def create_widgets(self):
        hero = tk.Frame(self.root, bg=CARD_BG, highlightthickness=1, highlightbackground="#eadfdb")
        hero.pack(fill="x", padx=12, pady=(12, 8))
        hero_inner = tk.Frame(hero, bg=CARD_BG)
        hero_inner.pack(fill="x", padx=20, pady=18)
        logo = self.carregar_logo(reducao=2)
        if logo:
            tk.Label(hero_inner, image=logo, bg=CARD_BG).pack(side="left", padx=(0, 16))
        header_texto = tk.Frame(hero_inner, bg=CARD_BG)
        header_texto.pack(side="left", fill="x", expand=True)
        tk.Label(header_texto, text="Contratos FOCO", bg=CARD_BG, fg=PRIMARY_TEXT, font=("Segoe UI", 21, "bold")).pack(anchor="w")
        tk.Label(
            header_texto,
            text="Busca, reenvio e download de contratos com progresso em tempo real.",
            bg=CARD_BG,
            fg=MUTED_TEXT,
            font=("Segoe UI", 10)
        ).pack(anchor="w", pady=(4, 0))
        tk.Label(header_texto, text="GESTAO DE CONTRATOS", bg=CARD_BG, fg="#a65f56", font=("Segoe UI", 9, "bold")).pack(anchor="w", pady=(8, 0))

        # Login
        login_frame = ttk.LabelFrame(self.root, text="🔑 Login do Sistema", padding=10)
        login_frame.pack(fill="x", padx=10, pady=5)
        ttk.Label(login_frame, text="Usuário:").grid(row=0, column=0, sticky="w", pady=2)
        ttk.Entry(login_frame, textvariable=self.usuario_var, width=30).grid(row=0, column=1, pady=2)
        ttk.Label(login_frame, text="Senha:").grid(row=1, column=0, sticky="w", pady=2)
        ttk.Entry(login_frame, textvariable=self.senha_var, show="*", width=30).grid(row=1, column=1, pady=2)

        # Configurações
        config_frame = ttk.LabelFrame(self.root, text="⚙️ Configurações", padding=10)
        config_frame.pack(fill="x", padx=10, pady=5)
        ttk.Checkbutton(config_frame, text="Executar em modo invisível (sem abrir janela do Chrome)", 
                       variable=self.headless_var).pack(anchor="w", pady=5)

        # Planilha e Pasta
        file_frame = ttk.LabelFrame(self.root, text="📁 Planilha e Pasta", padding=10)
        file_frame.pack(fill="x", padx=10, pady=5)
        ttk.Button(file_frame, text="Selecionar Planilha Excel", command=self.selecionar_planilha, style="Secondary.TButton").pack(pady=5)
        ttk.Label(file_frame, textvariable=self.planilha_var, wraplength=900, foreground="blue").pack(anchor="w")
        ttk.Button(file_frame, text="Mudar Pasta de Salvamento", command=self.selecionar_pasta, style="Secondary.TButton").pack(pady=5)
        ttk.Label(file_frame, textvariable=self.pasta_var, wraplength=900, foreground="blue").pack(anchor="w")

        # Progresso
        progress_frame = ttk.LabelFrame(self.root, text="Progresso da Execução", padding=10)
        progress_frame.pack(fill="x", padx=10, pady=5)
        self.progress_bar = ttk.Progressbar(progress_frame, variable=self.progress_var, maximum=100, style="Accent.Horizontal.TProgressbar")
        self.progress_bar.pack(fill="x", pady=5)
        self.label_progress = ttk.Label(progress_frame, text="0% - Aguardando início...")
        self.label_progress.pack()

        # Botão Iniciar
        self.btn_iniciar = ttk.Button(self.root, text="🚀 INICIAR ROBÔ", command=self.iniciar_robo, style="Primary.TButton")
        self.btn_iniciar.pack(pady=15)

        # Logs
        log_frame = ttk.LabelFrame(self.root, text="📜 Logs em Tempo Real", padding=10)
        log_frame.pack(fill="both", expand=True, padx=10, pady=5)
        self.log_text = scrolledtext.ScrolledText(log_frame, height=18, state='disabled', font=("Consolas", 10))
        self.log_text.pack(fill="both", expand=True)
        self.log_text.configure(bg="#fffaf9", fg="#303030", insertbackground=PRIMARY_TEXT, relief="flat", bd=0, highlightthickness=1, highlightbackground="#eadfdb")

        ttk.Button(self.root, text="💾 Salvar Log", command=self.salvar_log, style="Secondary.TButton").pack(pady=5)

        ttk.Label(self.root, text="Desenvolvido por Diogo © 2026", foreground="gray").pack(pady=5)

        self.root.after(100, self.process_queue)

    def log(self, msg):
        self.log_text.configure(state='normal')
        self.log_text.insert(tk.END, msg + "\n")
        self.log_text.see(tk.END)
        self.log_text.configure(state='disabled')

    def process_queue(self):
        while not self.log_queue.empty():
            msg = self.log_queue.get()
            self.log(msg)
        if not self.result_queue.empty():
            self.mostrar_resumo(self.result_queue.get())
        self.root.after(100, self.process_queue)

    def atualizar_progresso(self, valor):
        self.progress_var.set(valor)
        self.label_progress.config(text=f"{int(valor)}% concluído")

    def selecionar_planilha(self):
        arq = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx *.xls")])
        if arq:
            self.planilha_var.set(arq)

    def selecionar_pasta(self):
        pasta = filedialog.askdirectory()
        if pasta:
            self.pasta_var.set(pasta)

    def iniciar_robo(self):
        if not self.planilha_var.get():
            messagebox.showwarning("Atenção", "Selecione a planilha Excel!")
            return

        self.btn_iniciar.configure(state="disabled")
        self.progress_var.set(0)
        self.label_progress.config(text="0% - Iniciando...")
        self.log_text.configure(state='normal')
        self.log_text.delete(1.0, tk.END)
        self.log_text.configure(state='disabled')

        thread = threading.Thread(
            target=executar_robo,
            args=(self.usuario_var.get(), self.senha_var.get(), self.planilha_var.get(),
                  self.pasta_var.get(), self.headless_var.get(), self.log, 
                  self.result_queue, self.atualizar_progresso),
            daemon=True
        )
        thread.start()

    def mostrar_resumo(self, result):
        s = result["success"]
        f = result["failed"]
        tempo = result.get("time", "N/A")

        messagebox.showinfo("🏁 Execução Finalizada", 
            f"✅ Sucessos: {len(s)}\n❌ Erros: {len(f)}\n⏱️ Tempo total: {tempo}\n\n" +
            "Sucessos:\n" + "\n".join(s) + "\n\nErros:\n" + "\n".join(f))
        
        self.btn_iniciar.configure(state="normal")
        self.progress_var.set(100)
        self.label_progress.config(text="100% - Finalizado")

    def salvar_log(self):
        try:
            log_content = self.log_text.get("1.0", tk.END)
            filename = f"Log_Robo_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
            with open(filename, "w", encoding="utf-8") as f:
                f.write(log_content)
            messagebox.showinfo("Salvo", f"Log salvo com sucesso!\nArquivo: {filename}")
        except Exception as e:
            messagebox.showerror("Erro", f"Não foi possível salvar o log:\n{str(e)}")


if __name__ == "__main__":
    multiprocessing.freeze_support()
    app = RoboContratosApp()
    app.root.mainloop()
