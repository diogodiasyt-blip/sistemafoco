from __future__ import annotations
import getpass
import os
import queue
import threading
import time
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from tkinter import filedialog, messagebox

import customtkinter as ctk
import pandas as pd
import requests
import tkinter as tk
import unicodedata
from PIL import Image
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait
from webdriver_manager.chrome import ChromeDriverManager


APP_TITLE = "Robo de Cobranca de Cartoes"
APP_GEOMETRY = "1120x760"

URL_VALIDACAO = "https://raw.githubusercontent.com/diogodiasyt-blip/validacaofoco/refs/heads/main/chave"
URL_PING_ABERTURA = "https://docs.google.com/forms/d/e/1FAIpQLScmxNbTO-vXw0LEOKIyEhSpIl9aTbw8x5hnEI5VY2eVMRh5gQ/formResponse"
URL_CORAL = "https://coral.aluguefoco.com.br/login"
URL_CONTRATOS_PENDENTES = "https://coral.aluguefoco.com.br/contratos-pendentes"

XPATH_LOGIN = "/html/body/foco-app/div[1]/foco-login/div/div/div/div/div[2]/form/div[1]/input"
XPATH_SENHA = "/html/body/foco-app/div[1]/foco-login/div/div/div/div/div[2]/form/div[2]/input"
XPATH_ENTRAR = "/html/body/foco-app/div[1]/foco-login/div/div/div/div/div[2]/form/button"
XPATH_MENU_CADASTROS = "/html/body/foco-app/div[1]/div/ul/li[8]/a/i"
XPATH_CONTRATOS_PENDENTES = "/html/body/foco-app/div[1]/foco-crud-list/div/div/div/div/div[4]/div[2]/a[4]"
XPATH_CAMPO_BUSCA = "/html/body/foco-app/div[1]/foco-adyen-profit-payments/div/div/div/div/div/div/article[1]/div/div/input"
XPATH_TITULO_CONTRATOS_PENDENTES = "//*[contains(translate(normalize-space(.), 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'), 'CONTRATOS PENDENTES DE PAGAMENTO')]"
XPATH_BOTAO_EFETUAR_PAGAMENTO = "/html/body/foco-app/div[1]/foco-adyen-profit-payments/div/div/div/div/div/div/article[2]/table/tbody/tr/td[4]/button[1]"
XPATH_CARTOES_RADIO = "//ngb-modal-window//section[contains(@class,'container-card')]//input[@formcontrolname='cardSelected' and @type='radio']"
XPATH_VALOR_MODAL = "/html/body/ngb-modal-window/div/div/foco-adyen-profit-payments-modal/div[2]/div[1]/form/div[1]/section[1]/input[2]"
XPATH_PARCELAMENTO = "//ngb-modal-window//select[@formcontrolname='profitInstallments']"
XPATH_SALVAR_MODAL = "/html/body/ngb-modal-window/div/div/foco-adyen-profit-payments-modal/div[2]/div[2]/button[2]"

MAIN_BG = "#f6f4f1"
CARD_BG = "#ffffff"
CARD_BORDER = "#eadfdb"
PRIMARY_TEXT = "#d81919"
MUTED_TEXT = "#5c5c5c"
BUTTON_BG = "#ef1a14"
BUTTON_ACTIVE_BG = "#c91410"
SOFT_RED = "#fff1ef"
SUCCESS_GREEN = "#0f8a4b"
WARNING_ORANGE = "#b96a10"


@dataclass
class ValidationResult:
    sheet_name: str
    total_rows: int
    ready_rows: int
    ignored_rows: int
    missing_columns: list[str]
    columns: list[str]
    contract_column: str | None = None
    value_column: str | None = None
    ready_records: list[dict[str, object]] | None = None


def normalize_text(value: object) -> str:
    text = unicodedata.normalize("NFKD", str(value or "")).encode("ASCII", "ignore").decode("ASCII")
    return " ".join(text.strip().upper().split())


def parse_money_value(value: object) -> float:
    if value is None or pd.isna(value):
        return 0.0
    if isinstance(value, str):
        cleaned = value.replace("R$", "").replace(" ", "").strip()
        if "," in cleaned:
            cleaned = cleaned.replace(".", "").replace(",", ".")
        numeric_value = pd.to_numeric(cleaned, errors="coerce")
        return float(0 if pd.isna(numeric_value) else numeric_value)
    numeric_value = pd.to_numeric(value, errors="coerce")
    return float(0 if pd.isna(numeric_value) else numeric_value)


def resolve_logo_candidates() -> list[Path]:
    candidates: list[Path] = []
    env_logo = os.environ.get("FOCO_LOGO_PNG", "").strip()
    env_assets = os.environ.get("FOCO_ASSETS_DIR", "").strip()
    if env_logo:
        candidates.append(Path(env_logo))
    if env_assets:
        candidates.append(Path(env_assets) / "logo.png")

    base_dir = Path(__file__).resolve().parent
    candidates.append(base_dir.parent / "assets" / "logo.png")
    candidates.append(Path.cwd() / "DESENVOLVIMENTO" / "assets" / "logo.png")
    candidates.append(Path.cwd() / "assets" / "logo.png")
    return candidates


class RoboCobrancaCartoesApp(ctk.CTk):
    def __init__(self) -> None:
        super().__init__()
        ctk.set_appearance_mode("light")
        ctk.set_default_color_theme("blue")

        self.title(APP_TITLE)
        self.geometry(APP_GEOMETRY)
        self.minsize(1040, 700)
        self.configure(fg_color=MAIN_BG)

        self.file_path_var = ctk.StringVar()
        self.status_var = ctk.StringVar(value="Aguardando planilha")

        self.total_var = ctk.StringVar(value="0")
        self.ready_var = ctk.StringVar(value="0")
        self.ignored_var = ctk.StringVar(value="0")
        self.error_var = ctk.StringVar(value="0")

        self.username_var = ctk.StringVar()
        self.password_var = ctk.StringVar()
        self.visible_mode_var = ctk.BooleanVar(value=True)

        self.validation_result: ValidationResult | None = None
        self.processing_thread: threading.Thread | None = None
        self.driver = None
        self.report_path: Path | None = None
        self.report_rows: list[dict[str, object]] = []
        self.stop_requested = False
        self.log_queue: queue.Queue[str] = queue.Queue()
        self.logo_image = self._load_logo()

        self._build_layout()
        self._update_action_buttons()
        self.after(150, self._drain_log_queue)

    def _load_logo(self):
        for candidate in resolve_logo_candidates():
            try:
                if candidate.exists():
                    image = Image.open(candidate)
                    return ctk.CTkImage(light_image=image, dark_image=image, size=(92, 56))
            except Exception:
                continue
        return None

    def _build_layout(self) -> None:
        container = ctk.CTkScrollableFrame(self, fg_color="transparent", corner_radius=0)
        container.pack(fill="both", expand=True, padx=22, pady=22)
        container.grid_columnconfigure(0, weight=1)

        self._build_header(container)
        self._build_file_section(container)
        self._build_access_section(container)
        self._build_indicator_section(container)
        self._build_execution_section(container)

        footer = ctk.CTkLabel(
            container,
            text="Criado por Diogo Medeiros - 2026",
            text_color="#ef574f",
            font=("Segoe UI", 12),
        )
        footer.grid(row=6, column=0, sticky="w", padx=6, pady=(6, 0))

    def _build_header(self, parent) -> None:
        header = ctk.CTkFrame(parent, fg_color=CARD_BG, corner_radius=26, border_width=1, border_color=CARD_BORDER)
        header.grid(row=0, column=0, sticky="ew", pady=(0, 18))
        header.grid_columnconfigure(1, weight=1)

        brand = ctk.CTkFrame(header, fg_color="transparent")
        brand.grid(row=0, column=0, padx=(22, 14), pady=22, sticky="nw")
        if self.logo_image is not None:
            ctk.CTkLabel(brand, text="", image=self.logo_image).pack(anchor="w")
        else:
            ctk.CTkLabel(brand, text="foco,", text_color=PRIMARY_TEXT, font=("Segoe UI", 24, "bold")).pack(anchor="w")
            ctk.CTkLabel(brand, text="aluguel de carros", text_color="#c7463f", font=("Segoe UI", 10, "bold")).pack(anchor="w")

        texts = ctk.CTkFrame(header, fg_color="transparent")
        texts.grid(row=0, column=1, sticky="ew", padx=(0, 22), pady=22)
        ctk.CTkLabel(texts, text="Cobranca de Cartoes", text_color=PRIMARY_TEXT, font=("Segoe UI", 30, "bold")).pack(anchor="w")
        ctk.CTkLabel(
            texts,
            text="Validacao de contratos aptos para cobranca em cartao.",
            text_color="#4b5563",
            font=("Segoe UI", 15),
        ).pack(anchor="w", pady=(8, 0))
    def _create_section(self, parent, row: int, title: str):
        section = ctk.CTkFrame(parent, fg_color=CARD_BG, corner_radius=22, border_width=1, border_color=CARD_BORDER)
        section.grid(row=row, column=0, sticky="ew", pady=(0, 16))
        section.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(section, text=title, text_color=PRIMARY_TEXT, font=("Segoe UI", 18, "bold")).pack(
            anchor="w", padx=18, pady=(16, 12)
        )
        return section

    def _build_file_section(self, parent) -> None:
        section = self._create_section(parent, 1, "Planilha de cobranca")
        content = ctk.CTkFrame(section, fg_color="transparent")
        content.pack(fill="x", padx=18, pady=(0, 18))
        content.grid_columnconfigure(0, weight=1)

        self._form_label(content, "Arquivo Excel").grid(row=0, column=0, sticky="w", pady=(0, 6))
        self._entry(content, self.file_path_var).grid(row=1, column=0, sticky="ew", padx=(0, 10))
        self._secondary_button(content, "Selecionar", self.select_excel_file, width=150).grid(row=1, column=1, sticky="e")
        self._secondary_button(content, "Validar planilha", self.validate_workbook, width=170).grid(
            row=1, column=2, sticky="e", padx=(10, 0)
        )

        hint = (
            "Apos anexar o modelo, o robo vai validar contratos aptos para cobranca usando "
            "as colunas 'Nº Contrato' e 'Saldo'."
        )
        ctk.CTkLabel(content, text=hint, text_color=MUTED_TEXT, font=("Segoe UI", 13), wraplength=900, justify="left").grid(
            row=2, column=0, columnspan=3, sticky="w", pady=(12, 0)
        )

    def _build_access_section(self, parent) -> None:
        section = self._create_section(parent, 2, "Acesso ao Coral")
        content = ctk.CTkFrame(section, fg_color="transparent")
        content.pack(fill="x", padx=18, pady=(0, 18))
        content.grid_columnconfigure((0, 1), weight=1)

        self._form_label(content, "Usuario").grid(row=0, column=0, sticky="w", padx=(0, 10), pady=(0, 6))
        self._form_label(content, "Senha").grid(row=0, column=1, sticky="w", padx=(10, 0), pady=(0, 6))
        self._entry(content, self.username_var).grid(row=1, column=0, sticky="ew", padx=(0, 10))
        self._entry(content, self.password_var, show="*").grid(row=1, column=1, sticky="ew", padx=(10, 0))
        ctk.CTkCheckBox(
            content,
            text="Executar em modo visivel",
            variable=self.visible_mode_var,
            text_color="#1f2937",
            fg_color=BUTTON_BG,
            hover_color=BUTTON_ACTIVE_BG,
            border_color=CARD_BORDER,
        ).grid(row=2, column=0, columnspan=2, sticky="w", pady=(12, 0))

    def _build_indicator_section(self, parent) -> None:
        section = self._create_section(parent, 3, "Indicadores da planilha")
        grid = ctk.CTkFrame(section, fg_color="transparent")
        grid.pack(fill="x", padx=18, pady=(0, 18))
        grid.grid_columnconfigure((0, 1, 2, 3), weight=1)

        self._indicator_card(grid, 0, "Total", self.total_var)
        self._indicator_card(grid, 1, "Aptos", self.ready_var, SUCCESS_GREEN)
        self._indicator_card(grid, 2, "Ignorados", self.ignored_var, WARNING_ORANGE)
        self._indicator_card(grid, 3, "Pendencias", self.error_var, PRIMARY_TEXT)

    def _build_execution_section(self, parent) -> None:
        section = self._create_section(parent, 4, "Execucao")
        content = ctk.CTkFrame(section, fg_color="transparent")
        content.pack(fill="both", expand=True, padx=18, pady=(0, 18))
        content.grid_columnconfigure(0, weight=1)

        actions = ctk.CTkFrame(content, fg_color="transparent")
        actions.grid(row=0, column=0, sticky="ew", pady=(0, 12))
        actions.grid_columnconfigure((0, 1, 2), weight=1)

        self.start_button = self._primary_button(actions, "Iniciar robo", self.start_processing)
        self.start_button.grid(row=0, column=0, sticky="ew", padx=(0, 8))
        self.pause_button = self._secondary_button(actions, "Pausar", self.pause_processing)
        self.pause_button.grid(row=0, column=1, sticky="ew", padx=8)
        self.stop_button = self._secondary_button(actions, "Parar", self.stop_processing)
        self.stop_button.grid(row=0, column=2, sticky="ew", padx=(8, 0))

        self.progress = ctk.CTkProgressBar(content, height=14, progress_color=BUTTON_BG, fg_color="#eadfdb")
        self.progress.grid(row=1, column=0, sticky="ew", pady=(0, 10))
        self.progress.set(0)

        ctk.CTkLabel(content, textvariable=self.status_var, text_color="#374151", font=("Segoe UI", 14, "bold")).grid(
            row=2, column=0, sticky="w", pady=(0, 8)
        )

        self.log_box = ctk.CTkTextbox(
            content,
            height=190,
            fg_color="#fffaf8",
            border_width=1,
            border_color=CARD_BORDER,
            corner_radius=16,
            text_color="#1f2937",
            font=("Consolas", 12),
        )
        self.log_box.grid(row=3, column=0, sticky="nsew")
        self.log("Interface carregada. Selecione a planilha modelo para iniciar a validacao.")

    def _form_label(self, parent, text: str):
        return ctk.CTkLabel(parent, text=text, text_color="#111827", font=("Segoe UI", 13, "bold"))

    def _entry(self, parent, variable, show: str | None = None):
        return ctk.CTkEntry(
            parent,
            textvariable=variable,
            show=show,
            height=42,
            corner_radius=12,
            border_color=CARD_BORDER,
            fg_color="#fffdfb",
            text_color="#111827",
        )

    def _primary_button(self, parent, text: str, command, width: int | None = None):
        return ctk.CTkButton(
            parent,
            text=text,
            command=command,
            width=width or 180,
            height=42,
            corner_radius=12,
            fg_color=BUTTON_BG,
            hover_color=BUTTON_ACTIVE_BG,
            text_color="white",
            font=("Segoe UI", 14, "bold"),
        )

    def _secondary_button(self, parent, text: str, command, width: int | None = None):
        return ctk.CTkButton(
            parent,
            text=text,
            command=command,
            width=width or 150,
            height=42,
            corner_radius=12,
            fg_color="#fffdfb",
            hover_color=SOFT_RED,
            border_width=1,
            border_color=CARD_BORDER,
            text_color=PRIMARY_TEXT,
            font=("Segoe UI", 14, "bold"),
        )

    def _indicator_card(self, parent, column: int, title: str, variable, color: str = PRIMARY_TEXT):
        card = ctk.CTkFrame(parent, fg_color="#fffaf8", corner_radius=16, border_width=1, border_color=CARD_BORDER)
        card.grid(row=0, column=column, sticky="ew", padx=6)
        ctk.CTkLabel(card, text=title, text_color="#4b5563", font=("Segoe UI", 12, "bold")).pack(anchor="w", padx=14, pady=(12, 0))
        ctk.CTkLabel(card, textvariable=variable, text_color=color, font=("Segoe UI", 26, "bold")).pack(anchor="w", padx=14, pady=(4, 12))

    def log(self, message: str) -> None:
        if threading.current_thread() is not threading.main_thread():
            self.log_queue.put(message)
            return
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log_box.insert("end", f"[{timestamp}] {message}\n")
        self.log_box.see("end")
        self.update_idletasks()

    def _drain_log_queue(self) -> None:
        while not self.log_queue.empty():
            self.log(self.log_queue.get())
        self.after(150, self._drain_log_queue)

    def select_excel_file(self) -> None:
        file_path = filedialog.askopenfilename(
            title="Selecionar planilha de cobranca de cartoes",
            filetypes=[("Excel", "*.xlsx *.xlsm *.xls")],
        )
        if not file_path:
            return
        self.file_path_var.set(file_path)
        self.validation_result = None
        self.progress.set(0)
        self.status_var.set("Planilha selecionada. Clique em Validar planilha.")
        self.log(f"Planilha selecionada: {file_path}")
        self._update_action_buttons()

    def validate_workbook(self) -> None:
        path_text = self.file_path_var.get().strip()
        if not path_text:
            messagebox.showwarning("Planilha", "Selecione a planilha antes de validar.")
            return

        workbook_path = Path(path_text)
        if not workbook_path.exists():
            messagebox.showerror("Arquivo nao encontrado", f"A planilha nao foi encontrada:\n{workbook_path}")
            return

        self.log("Lendo planilha Excel...")
        self.progress.set(0.15)
        try:
            excel_file = pd.ExcelFile(workbook_path)
            sheet_name = excel_file.sheet_names[0]
            dataframe = pd.read_excel(workbook_path, sheet_name=sheet_name)
            dataframe = dataframe.dropna(how="all")

            columns = list(dataframe.columns)
            normalized_columns = {normalize_text(column): column for column in columns}
            contract_col = self._find_column(
                normalized_columns,
                ["Nº CONTRATO", "NO CONTRATO", "N CONTRATO", "NUMERO CONTRATO"],
            )
            value_col = self._find_column(normalized_columns, ["SALDO"])

            missing_columns = []
            if not contract_col:
                missing_columns.append("Nº Contrato")
            if not value_col:
                missing_columns.append("Saldo")

            total_rows = len(dataframe.index)
            if contract_col:
                valid_contract = dataframe[contract_col].fillna("").astype(str).str.strip().ne("")
            else:
                valid_contract = pd.Series([False] * total_rows, index=dataframe.index)

            if value_col:
                values = dataframe[value_col].apply(parse_money_value)
                positive_value = values.gt(0)
            else:
                positive_value = pd.Series([True] * total_rows, index=dataframe.index)

            ready_mask = valid_contract & positive_value
            ready_rows = int(ready_mask.sum())
            ignored_rows = int(total_rows - ready_rows)
            ready_records = []
            if contract_col and value_col:
                for _, row in dataframe.loc[ready_mask].iterrows():
                    ready_records.append(
                        {
                            "contrato": str(row.get(contract_col, "")).strip(),
                            "saldo": parse_money_value(row.get(value_col, 0)),
                        }
                    )

            self.validation_result = ValidationResult(
                sheet_name=sheet_name,
                total_rows=total_rows,
                ready_rows=ready_rows,
                ignored_rows=ignored_rows,
                missing_columns=missing_columns,
                columns=[str(column) for column in columns],
                contract_column=str(contract_col) if contract_col else None,
                value_column=str(value_col) if value_col else None,
                ready_records=ready_records,
            )

            self.total_var.set(str(total_rows))
            self.ready_var.set(str(ready_rows if not missing_columns else 0))
            self.ignored_var.set(str(ignored_rows))
            self.error_var.set(str(len(missing_columns)))
            self.progress.set(1)

            if missing_columns:
                self.status_var.set("Validacao com pendencias de estrutura.")
                self.log(f"Colunas ausentes: {', '.join(missing_columns)}")
                messagebox.showwarning(
                    "Validacao com pendencias",
                    "A planilha foi lida, mas faltam colunas obrigatorias:\n\n- " + "\n- ".join(missing_columns),
                )
            else:
                self.status_var.set(f"Planilha validada. {ready_rows} contrato(s) apto(s) para cobranca.")
                self.log(f"Aba validada: {sheet_name}")
                self.log(f"Colunas utilizadas: {contract_col} | {value_col}")
                self.log(f"Total: {total_rows} | Aptos: {ready_rows} | Ignorados: {ignored_rows}")
        except Exception as exc:
            self.validation_result = None
            self.progress.set(0)
            self.total_var.set("0")
            self.ready_var.set("0")
            self.ignored_var.set("0")
            self.error_var.set("1")
            self.status_var.set("Falha na validacao da planilha.")
            self.log(f"Erro de validacao: {exc}")
            messagebox.showerror("Falha na validacao", str(exc))
        finally:
            self._update_action_buttons()

    @staticmethod
    def _find_column(normalized_columns: dict[str, object], aliases: list[str]):
        for alias in aliases:
            alias_norm = normalize_text(alias)
            if alias_norm in normalized_columns:
                return normalized_columns[alias_norm]
        return None

    def registrar_abertura(self) -> None:
        try:
            usuario_coral = self.username_var.get().strip()
            usuario_windows = os.environ.get("USERNAME", "").strip() or getpass.getuser().strip()
            usuario_registro = usuario_coral or usuario_windows or "usuario_desconhecido"
            data = {
                "entry.1320712185": "Robo Cobranca Cartoes",
                "entry.1823299431": usuario_registro,
                "entry.1825360926": datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
            }
            requests.post(URL_PING_ABERTURA, data=data, timeout=5)
            self.log(f"Abertura liberada. Usuario registrado: {usuario_registro}.")
        except Exception as exc:
            self.log(f"Falha ao registrar abertura: {exc}")

    def verificar_chave(self) -> bool:
        try:
            response = requests.get(URL_VALIDACAO, timeout=10)
            status = response.text.strip().upper()
            self.log(f"STATUS DO ROBO: {status or 'INDEFINIDO'}")
            return status == "ATIVO"
        except Exception as exc:
            self.log(f"Falha ao validar chave remota: {exc}. Mantendo execucao liberada.")
            return True

    def _create_driver(self):
        options = webdriver.ChromeOptions()
        if not self.visible_mode_var.get():
            options.add_argument("--headless=new")
        options.add_argument("--start-maximized")
        options.add_argument("--disable-gpu")
        options.add_argument("--no-sandbox")
        options.add_argument("--disable-dev-shm-usage")
        options.add_argument("--log-level=3")
        options.add_experimental_option("excludeSwitches", ["enable-logging"])

        self.log("Criando sessao do Chrome...")
        return webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)

    def _wait_clickable(self, xpath: str, description: str, timeout: int = 30):
        self.log(f"Aguardando {description} ficar disponivel...")
        return WebDriverWait(self.driver, timeout).until(EC.element_to_be_clickable((By.XPATH, xpath)))

    def _wait_visible(self, xpath: str, description: str, timeout: int = 30):
        self.log(f"Aguardando {description} aparecer...")
        return WebDriverWait(self.driver, timeout).until(EC.visibility_of_element_located((By.XPATH, xpath)))

    def _safe_click(self, xpath: str, description: str, timeout: int = 30) -> None:
        last_error = None
        for attempt in range(1, 4):
            try:
                element = self._wait_clickable(xpath, description, timeout=timeout)
                self.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", element)
                time.sleep(0.2)
                try:
                    element.click()
                except Exception:
                    self.driver.execute_script("arguments[0].click();", element)
                self.log(f"Clique OK em {description}.")
                return
            except Exception as exc:
                last_error = exc
                self.log(f"Clique falhou em {description} ({attempt}/3): {exc}")
                time.sleep(1)
        raise RuntimeError(f"Nao foi possivel clicar em {description} apos 3 tentativas: {last_error}")

    def _safe_type(self, xpath: str, value: str, description: str, timeout: int = 30, press_enter: bool = False) -> None:
        last_error = None
        for attempt in range(1, 4):
            try:
                element = self._wait_clickable(xpath, description, timeout=timeout)
                self.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", element)
                time.sleep(0.2)
                element.click()
                element.send_keys(Keys.CONTROL, "a")
                element.send_keys(Keys.DELETE)
                element.send_keys(str(value))
                if press_enter:
                    element.send_keys(Keys.ENTER)
                safe_value = "********" if "senha" in normalize_text(description).lower() else value
                self.log(f"Texto preenchido em {description}: {safe_value}")
                return
            except Exception as exc:
                last_error = exc
                self.log(f"Preenchimento falhou em {description} ({attempt}/3): {exc}")
                time.sleep(1)
        raise RuntimeError(f"Nao foi possivel preencher {description} apos 3 tentativas: {last_error}")

    def _clear_input_field(self, element) -> None:
        element.click()
        element.send_keys(Keys.CONTROL, "a")
        element.send_keys(Keys.DELETE)
        element.clear()
        self.driver.execute_script(
            """
            arguments[0].value = '';
            arguments[0].dispatchEvent(new Event('input', { bubbles: true }));
            arguments[0].dispatchEvent(new Event('change', { bubbles: true }));
            """,
            element,
        )
        time.sleep(0.2)

    def _read_input_value(self, xpath: str, description: str, timeout: int = 10) -> str:
        last_error = None
        for attempt in range(1, 4):
            try:
                element = self._wait_visible(xpath, description, timeout=timeout)
                return (element.get_attribute("value") or "").strip()
            except Exception as exc:
                last_error = exc
                self.log(f"Leitura falhou em {description} ({attempt}/3): {exc}")
                time.sleep(1)
        raise RuntimeError(f"Nao foi possivel ler {description} apos 3 tentativas: {last_error}")

    @staticmethod
    def _money_values_match(expected: object, actual: object) -> bool:
        return abs(parse_money_value(expected) - parse_money_value(actual)) < 0.01

    def _fill_and_validate_money(self, xpath: str, value: object, description: str, timeout: int = 30) -> None:
        expected_text = self._format_money_for_coral(value)
        last_read = ""

        for attempt in range(1, 4):
            try:
                element = self._wait_clickable(xpath, description, timeout=timeout)
                self.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", element)
                time.sleep(0.2)
                self._clear_input_field(element)
                element.send_keys(expected_text)
                time.sleep(0.4)
                last_read = self._read_input_value(xpath, description, timeout=10)
            except Exception as exc:
                last_read = f"erro: {exc}"
                self.log(f"Erro ao preencher valor em {description} ({attempt}/3): {exc}")
                time.sleep(1)
                continue

            if self._money_values_match(expected_text, last_read):
                self.log(f"Valor validado em {description}: esperado {expected_text} | lido {last_read}")
                return

            self.log(
                f"Validacao do valor falhou ({attempt}/3). "
                f"Esperado: {expected_text} | Lido: {last_read or '<vazio>'}"
            )

        raise RuntimeError(
            f"Valor divergente em {description}. Esperado {expected_text}, lido {last_read or '<vazio>'}."
        )

    def _safe_select_value(self, xpath: str, value: str, description: str, timeout: int = 30) -> None:
        last_error = None
        for attempt in range(1, 4):
            try:
                element = self._wait_clickable(xpath, description, timeout=timeout)
                self.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", element)
                time.sleep(0.2)
                Select(element).select_by_value(str(value))
                self.log(f"Selecao OK em {description}: {value}")
                return
            except Exception as exc:
                last_error = exc
                self.log(f"Selecao falhou em {description} ({attempt}/3): {exc}")
                time.sleep(1)
        raise RuntimeError(f"Nao foi possivel selecionar {description} apos 3 tentativas: {last_error}")

    @staticmethod
    def _format_money_for_coral(value: object) -> str:
        return f"{parse_money_value(value):.2f}".replace(".", ",")

    def _go_to_pending_payments_safe(self) -> None:
        self.status_var.set("Retornando ao porto seguro...")
        self.log("Redirecionando para Contratos pendentes de pagamento...")
        self.driver.get(URL_CONTRATOS_PENDENTES)
        self._wait_visible(XPATH_TITULO_CONTRATOS_PENDENTES, "titulo Contratos pendentes de pagamento", timeout=30)
        self._wait_visible(XPATH_CAMPO_BUSCA, "campo de busca de contratos", timeout=30)
        self.log("Porto seguro validado: Contratos pendentes de pagamento.")

    def _find_pending_contract_row(self, contract_number: str, timeout: int = 30):
        contract_xpath = (
            "//foco-adyen-profit-payments//table//tbody//tr"
            f"[td[normalize-space()='{contract_number}']]"
        )
        self.log(f"Aguardando resultado da busca para contrato {contract_number}...")
        return WebDriverWait(self.driver, timeout).until(EC.visibility_of_element_located((By.XPATH, contract_xpath)))

    def _count_cards(self, timeout: int = 30) -> int:
        self.log("Aguardando cartoes disponiveis no modal...")
        WebDriverWait(self.driver, timeout).until(
            lambda driver: len(driver.find_elements(By.XPATH, XPATH_CARTOES_RADIO)) > 0
        )
        cards = self.driver.find_elements(By.XPATH, XPATH_CARTOES_RADIO)
        self.log(f"Cartoes disponiveis no contrato: {len(cards)}")
        return len(cards)

    def _select_card_by_index(self, index: int, timeout: int = 30) -> None:
        self.log(f"Selecionando cartao {index + 1}...")
        WebDriverWait(self.driver, timeout).until(
            lambda driver: len(driver.find_elements(By.XPATH, XPATH_CARTOES_RADIO)) > index
        )
        card = self.driver.find_elements(By.XPATH, XPATH_CARTOES_RADIO)[index]
        self.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", card)
        time.sleep(0.2)
        try:
            card.click()
        except Exception:
            self.driver.execute_script("arguments[0].click();", card)
        self.log(f"Cartao {index + 1} selecionado.")

    def _prepare_report(self) -> None:
        workbook_path = Path(self.file_path_var.get().strip())
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        self.report_path = workbook_path.with_name(f"Relatorio_Cobranca_Cartoes_{timestamp}.xlsx")
        self.report_rows = []
        self._save_report()
        self.log(f"Relatorio criado: {self.report_path}")

    def _append_report_row(
        self,
        contract_number: str,
        total_cards: int,
        attempted_cards: int,
        attempted_balance: object,
        status: str,
        detail: str = "",
    ) -> None:
        self.report_rows.append(
            {
                "contrato": contract_number,
                "quantidade_cartoes": total_cards,
                "cartoes_tentados": attempted_cards,
                "saldo_tentado": attempted_balance,
                "status": status,
                "detalhe": detail,
                "data_hora": datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
            }
        )
        self._save_report()

    def _save_report(self) -> None:
        if self.report_path is None:
            return
        pd.DataFrame(
            self.report_rows,
            columns=[
                "contrato",
                "quantidade_cartoes",
                "cartoes_tentados",
                "saldo_tentado",
                "status",
                "detalhe",
                "data_hora",
            ],
        ).to_excel(self.report_path, index=False)

    def _open_contract_payment_modal(self, contract_number: str) -> None:
        self._safe_type(XPATH_CAMPO_BUSCA, contract_number, "campo de busca de contratos", timeout=30, press_enter=True)
        self._find_pending_contract_row(contract_number, timeout=45)
        self._safe_click(XPATH_BOTAO_EFETUAR_PAGAMENTO, "botao Efetuar pagamento", timeout=30)

    def _attempt_charge_current_modal(self, contract_number: str, balance: object) -> tuple[str, int, int, str]:
        try:
            total_cards = self._count_cards(timeout=45)
        except Exception as exc:
            return "NAO_COBRADO", 0, 0, f"Nenhum cartao disponivel/localizado: {exc}"
        balance_text = self._format_money_for_coral(balance)
        detail = "Aguardando mapeamento do modal de retorno da cobrança."

        for card_index in range(total_cards):
            if self.stop_requested:
                return "INTERROMPIDO", total_cards, card_index, "Execucao interrompida pelo usuario."
            try:
                self._select_card_by_index(card_index, timeout=30)
                self._fill_and_validate_money(XPATH_VALOR_MODAL, balance_text, "campo Valor", timeout=30)
                self._safe_select_value(XPATH_PARCELAMENTO, "1", "parcelamento 1x", timeout=30)
                current_value = self._read_input_value(XPATH_VALOR_MODAL, "campo Valor", timeout=10)
                if not self._money_values_match(balance_text, current_value):
                    raise RuntimeError(
                        f"Valor mudou antes de salvar. Esperado {balance_text}, lido {current_value or '<vazio>'}."
                    )
                self._safe_click(XPATH_SALVAR_MODAL, "botao Salvar pagamento", timeout=30)
                return "AGUARDANDO_CONFIRMACAO", total_cards, card_index + 1, detail
            except Exception as exc:
                self.log(f"Falha ao tentar cartao {card_index + 1}: {exc}")

        return "NAO_COBRADO", total_cards, total_cards, "Todas as tentativas de cartao falharam antes da confirmacao."

    def _process_contract_payment(self, contract_number: str, balance: object) -> None:
        self.status_var.set(f"Processando contrato {contract_number}...")
        last_error = ""

        for attempt in range(1, 4):
            if self.stop_requested:
                self._append_report_row(
                    contract_number,
                    0,
                    0,
                    balance,
                    "INTERROMPIDO",
                    "Execucao interrompida pelo usuario antes da tentativa.",
                )
                return

            self.log(f"Tentativa {attempt}/3 para contrato {contract_number}.")
            try:
                self._go_to_pending_payments_safe()
                self._open_contract_payment_modal(contract_number)
                status, total_cards, attempted_cards, detail = self._attempt_charge_current_modal(contract_number, balance)
                self._append_report_row(contract_number, total_cards, attempted_cards, balance, status, detail)
                self.log(
                    f"Relatorio atualizado: contrato={contract_number} | cartoes={total_cards} | "
                    f"tentativas={attempted_cards} | saldo={self._format_money_for_coral(balance)} | status={status}"
                )
                return
            except Exception as exc:
                last_error = str(exc)
                self.log(f"Tentativa {attempt}/3 falhou no contrato {contract_number}: {exc}")
                try:
                    self._go_to_pending_payments_safe()
                except Exception as safe_exc:
                    self.log(f"Falha ao retornar ao porto seguro apos erro: {safe_exc}")
                time.sleep(2)

        self._append_report_row(
            contract_number,
            0,
            0,
            balance,
            "ERRO",
            f"Contrato falhou apos 3 tentativas. Ultimo erro: {last_error}",
        )
        self.log(f"Contrato {contract_number} registrado como ERRO apos 3 tentativas.")

    def _login_and_open_pending_payments(self) -> None:
        usuario = self.username_var.get().strip()
        senha = self.password_var.get().strip()
        if not usuario or not senha:
            raise RuntimeError("Informe usuario e senha do Coral antes de iniciar.")

        self.driver = self._create_driver()
        self.status_var.set("Acessando Coral...")
        self.log("Acessando tela de login do Coral...")
        self.driver.get(URL_CORAL)

        self._safe_type(XPATH_LOGIN, usuario, "campo de login")
        self._safe_type(XPATH_SENHA, senha, "campo de senha")
        self._safe_click(XPATH_ENTRAR, "botao Entrar")

        self.status_var.set("Navegando para contratos pendentes...")
        time.sleep(3)
        self._go_to_pending_payments_safe()

    def start_processing(self) -> None:
        if self.validation_result is None:
            messagebox.showwarning("Validacao", "Valide a planilha antes de iniciar.")
            return
        if self.validation_result.missing_columns:
            messagebox.showwarning("Validacao", "Corrija as colunas obrigatorias antes de iniciar.")
            return
        if self.validation_result.ready_rows <= 0:
            messagebox.showwarning("Validacao", "Nao ha contratos aptos para cobranca.")
            return
        if self.processing_thread is not None and self.processing_thread.is_alive():
            messagebox.showinfo("Execucao", "O robo ja esta em execucao.")
            return

        self.stop_requested = False
        self.start_button.configure(state="disabled")
        self.stop_button.configure(state="normal")
        self.pause_button.configure(state="disabled")
        self.progress.set(0.05)
        self.status_var.set("Validando liberacao remota...")
        self.processing_thread = threading.Thread(target=self._run_processing_stub, daemon=True)
        self.processing_thread.start()

    def _run_processing_stub(self) -> None:
        try:
            self.registrar_abertura()
            if not self.verificar_chave():
                self.status_var.set("Robo bloqueado remotamente.")
                self.log("Execucao bloqueada pela validacao remota.")
                return

            self.log("Robo liberado para execucao.")
            self._login_and_open_pending_payments()
            self._prepare_report()
            records = self.validation_result.ready_records or []
            total_records = len(records)
            if not records:
                raise RuntimeError("Nenhum contrato apto foi carregado da planilha.")

            for index, record in enumerate(records, start=1):
                if self.stop_requested:
                    self.log("Execucao interrompida antes de concluir todos os contratos.")
                    break

                contract_number = str(record.get("contrato", "")).strip()
                balance = record.get("saldo", 0)
                self.log(f"[{index}/{total_records}] Contrato {contract_number} | Saldo {self._format_money_for_coral(balance)}")

                try:
                    self._process_contract_payment(contract_number, balance)
                except Exception as exc:
                    self.log(f"Erro no contrato {contract_number}: {exc}")
                    self._append_report_row(contract_number, 0, 0, balance, "ERRO", str(exc))
                    try:
                        self._go_to_pending_payments_safe()
                    except Exception as safe_exc:
                        self.log(f"Nao foi possivel retornar ao porto seguro: {safe_exc}")

                self.progress.set(min(1, index / total_records))

            self.status_var.set("Execucao finalizada. Confira o relatorio.")
            messagebox.showinfo(
                "Execucao finalizada",
                f"Processamento finalizado.\n\nRelatorio:\n{self.report_path}",
            )
        except Exception as exc:
            self.status_var.set("Erro na preparacao")
            self.log(f"ERRO: {exc}")
            messagebox.showerror("Erro", str(exc))
        finally:
            self._update_action_buttons()

    def pause_processing(self) -> None:
        messagebox.showinfo("Pausa", "Controle de pausa sera habilitado junto com a automacao operacional.")

    def stop_processing(self) -> None:
        self.stop_requested = True
        self.status_var.set("Execucao interrompida.")
        self.log("Execucao interrompida pelo usuario.")
        self._update_action_buttons()

    def _update_action_buttons(self) -> None:
        running = self.processing_thread is not None and self.processing_thread.is_alive()
        valid = self.validation_result is not None and not self.validation_result.missing_columns and self.validation_result.ready_rows > 0
        self.start_button.configure(state="normal" if valid and not running else "disabled")
        self.pause_button.configure(state="disabled")
        self.stop_button.configure(state="normal" if running else "disabled")


if __name__ == "__main__":
    app = RoboCobrancaCartoesApp()
    app.mainloop()
