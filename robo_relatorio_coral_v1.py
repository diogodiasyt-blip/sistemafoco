from __future__ import annotations

import getpass
import os
import queue
import re
import tempfile
import threading
import time
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from tkinter import filedialog, messagebox

import customtkinter as ctk
import pandas as pd
import requests
from PIL import Image
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from webdriver_manager.chrome import ChromeDriverManager


APP_TITLE = "Robo de Relatorio Coral"
APP_GEOMETRY = "1120x760"

URL_VALIDACAO = "https://raw.githubusercontent.com/diogodiasyt-blip/validacaofoco/refs/heads/main/chave"
URL_PING_ABERTURA = "https://docs.google.com/forms/d/e/1FAIpQLScmxNbTO-vXw0LEOKIyEhSpIl9aTbw8x5hnEI5VY2eVMRh5gQ/formResponse"
URL_CORAL = "https://coral.aluguefoco.com.br/login"
URL_DASHBOARD_POS_LOGIN = "https://coral.aluguefoco.com.br/precificacao/dashboard"
URL_RELATORIOS = "https://coral.aluguefoco.com.br/relatorios"

XPATH_LOGIN = "/html/body/foco-app/div[1]/foco-login/div/div/div/div/div[2]/form/div[1]/input"
XPATH_SENHA = "/html/body/foco-app/div[1]/foco-login/div/div/div/div/div[2]/form/div[2]/input"
XPATH_ENTRAR = "/html/body/foco-app/div[1]/foco-login/div/div/div/div/div[2]/form/button"
XPATH_GERAR_RELATORIO = "/html/body/foco-app/div[1]/foco-analytics-home/div/div/div/div/div/div/div[2]/div[1]/div[3]/button/span"
REPORT_NAME = "Relatório Eficiência Brokers Financeiro"

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
class ConversionResult:
    source_csv: Path
    output_xlsx: Path
    rows: int
    columns: int


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


def get_desktop_dir() -> Path:
    """Detecta automaticamente a Área de Trabalho real do usuário (mesmo com OneDrive personalizado)"""
    user_profile = Path(os.environ.get("USERPROFILE", str(Path.home())))
    
    candidates: list[Path] = []

    # 1. Tentativa via variáveis de ambiente (mais confiável)
    for env_var in ("OneDriveCommercial", "OneDrive", "OneDriveConsumer"):
        onedrive_path = os.environ.get(env_var)
        if onedrive_path:
            onedrive = Path(onedrive_path)
            candidates.extend([
                onedrive / "Desktop",
                onedrive / "Área de Trabalho",
            ])

    # 2. Busca inteligente por pastas OneDrive no perfil do usuário
    try:
        for item in user_profile.iterdir():
            if item.is_dir() and item.name.startswith("OneDrive"):
                candidates.extend([
                    item / "Desktop",
                    item / "Área de Trabalho",
                ])
    except Exception:
        pass  # se der erro na listagem, ignora

    # 3. Pastas locais tradicionais (fallback)
    candidates.extend([
        user_profile / "Desktop",
        user_profile / "Área de Trabalho",
        Path.home() / "Desktop",
        Path.home() / "Área de Trabalho",
    ])

    # Remove duplicatas preservando ordem
    seen = set()
    unique_candidates = []
    for path in candidates:
        resolved = path.resolve()
        if resolved not in seen:
            seen.add(resolved)
            unique_candidates.append(path)

    # Testa qual existe e retorna a primeira válida
    for candidate in unique_candidates:
        if candidate.exists():
            return candidate

    # Último fallback
    return user_profile / "Desktop"


def parse_ptbr_date(value: str) -> datetime:
    return datetime.strptime(value.strip(), "%d/%m/%Y")


def format_output_name(start_date: str, end_date: str) -> str:
    start_safe = parse_ptbr_date(start_date).strftime("%Y%m%d")
    end_safe = parse_ptbr_date(end_date).strftime("%Y%m%d")
    timestamp = datetime.now().strftime("%H%M%S")
    return f"Relatorio_Coral_{start_safe}_a_{end_safe}_{timestamp}.xlsx"


def read_coral_csv(csv_path: Path) -> pd.DataFrame:
    errors: list[str] = []
    for encoding in ("utf-8-sig", "utf-8", "latin1"):
        try:
            dataframe = pd.read_csv(csv_path, sep=",", decimal=".", encoding=encoding)
            return normalize_numeric_columns(dataframe)
        except Exception as exc:
            errors.append(f"{encoding}: {exc}")
    raise RuntimeError("Nao foi possivel ler o CSV exportado pelo Coral. " + " | ".join(errors))


def normalize_numeric_columns(dataframe: pd.DataFrame) -> pd.DataFrame:
    numeric_pattern = re.compile(r"^-?\d+(\.\d+)?$")
    result = dataframe.copy()
    for column in result.columns:
        if not pd.api.types.is_object_dtype(result[column]):
            continue
        values = result[column].dropna().astype(str).str.strip()
        if values.empty:
            continue
        sample = values.head(50)
        numeric_like = sample.map(lambda item: bool(numeric_pattern.fullmatch(item))).mean()
        if numeric_like >= 0.9:
            result[column] = pd.to_numeric(result[column], errors="coerce")
    return result


def convert_coral_csv_to_xlsx(csv_path: Path, output_dir: Path, start_date: str, end_date: str) -> ConversionResult:
    if not csv_path.exists():
        raise FileNotFoundError(f"CSV nao encontrado: {csv_path}")

    dataframe = read_coral_csv(csv_path)
    output_dir.mkdir(parents=True, exist_ok=True)
    output_path = output_dir / format_output_name(start_date, end_date)
    dataframe.to_excel(output_path, index=False)
    return ConversionResult(
        source_csv=csv_path,
        output_xlsx=output_path,
        rows=len(dataframe),
        columns=len(dataframe.columns),
    )


class RoboRelatorioCoralApp(ctk.CTk):
    def __init__(self) -> None:
        super().__init__()
        ctk.set_appearance_mode("light")
        ctk.set_default_color_theme("blue")

        self.title(APP_TITLE)
        self.geometry(APP_GEOMETRY)
        self.minsize(1040, 700)
        self.configure(fg_color=MAIN_BG)

        self.username_var = ctk.StringVar()
        self.password_var = ctk.StringVar()
        today = datetime.now().strftime("%d/%m/%Y")
        self.start_date_var = ctk.StringVar(value=today)
        self.end_date_var = ctk.StringVar(value=today)
        self.output_dir_var = ctk.StringVar(value=str(get_desktop_dir()))
        self.status_var = ctk.StringVar(value="Aguardando inicio")
        self.visible_mode_var = ctk.BooleanVar(value=False)
        self.keep_csv_var = ctk.BooleanVar(value=False)

        self.log_queue: queue.Queue[str] = queue.Queue()
        self.processing_thread: threading.Thread | None = None
        self.stop_requested = False
        self.driver = None
        self.download_dir = Path(tempfile.gettempdir()) / "SistemaFOCO" / "downloads_relatorio_coral"
        self.logo_image = None

        self._build_layout()
        self._poll_log_queue()
        self._update_action_buttons()

    def _build_layout(self) -> None:
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)

        scroll = ctk.CTkScrollableFrame(self, fg_color=MAIN_BG, corner_radius=0)
        scroll.grid(row=0, column=0, sticky="nsew")
        scroll.grid_columnconfigure(0, weight=1)

        container = ctk.CTkFrame(scroll, fg_color="transparent")
        container.grid(row=0, column=0, sticky="nsew", padx=24, pady=20)
        container.grid_columnconfigure(0, weight=1)

        self._build_header(container)
        self._build_access_section(container)
        self._build_period_section(container)
        self._build_execution_section(container)

    def _build_header(self, parent) -> None:
        card = self._card(parent)
        card.grid(row=0, column=0, sticky="ew", pady=(0, 16))
        card.grid_columnconfigure(1, weight=1)

        logo_loaded = False
        for candidate in resolve_logo_candidates():
            try:
                if candidate.exists():
                    image = Image.open(candidate)
                    self.logo_image = ctk.CTkImage(light_image=image, dark_image=image, size=(112, 54))
                    ctk.CTkLabel(card, image=self.logo_image, text="").grid(row=0, column=0, rowspan=3, padx=(18, 22), pady=18)
                    logo_loaded = True
                    break
            except Exception:
                continue

        if not logo_loaded:
            ctk.CTkLabel(card, text="foco", text_color=PRIMARY_TEXT, font=("Segoe UI", 32, "bold")).grid(
                row=0, column=0, rowspan=3, padx=(18, 22), pady=18
            )

        ctk.CTkLabel(
            card,
            text="Relatorio Coral",
            text_color=PRIMARY_TEXT,
            font=("Segoe UI", 28, "bold"),
            anchor="w",
        ).grid(row=0, column=1, sticky="ew", padx=(0, 20), pady=(22, 4))
        ctk.CTkLabel(
            card,
            text="Emite relatorio no Coral, baixa CSV e entrega Excel corrigido para PT-BR.",
            text_color=MUTED_TEXT,
            font=("Segoe UI", 14),
            anchor="w",
        ).grid(row=1, column=1, sticky="ew", padx=(0, 20))
        ctk.CTkLabel(
            card,
            text="OPERACAO DE RELATORIOS",
            text_color="#b65748",
            font=("Segoe UI", 12, "bold"),
            anchor="w",
        ).grid(row=2, column=1, sticky="ew", padx=(0, 20), pady=(10, 22))

    def _build_access_section(self, parent) -> None:
        card = self._section(parent, "Acesso ao Coral", 1)
        card.grid_columnconfigure((0, 1), weight=1)

        self._label(card, "Usuario").grid(row=1, column=0, sticky="w", padx=18, pady=(8, 4))
        self._entry(card, self.username_var).grid(row=2, column=0, sticky="ew", padx=(18, 10), pady=(0, 14))

        self._label(card, "Senha").grid(row=1, column=1, sticky="w", padx=18, pady=(8, 4))
        self._entry(card, self.password_var, show="*").grid(row=2, column=1, sticky="ew", padx=(10, 18), pady=(0, 14))

    def _build_period_section(self, parent) -> None:
        card = self._section(parent, "Periodo do relatorio", 2)
        card.grid_columnconfigure((0, 1), weight=1)

        start_frame = ctk.CTkFrame(card, fg_color="transparent")
        start_frame.grid(row=1, column=0, sticky="ew", padx=(18, 10), pady=(8, 14))
        start_frame.grid_columnconfigure(0, weight=1)
        self._label(start_frame, "Data inicial").grid(row=0, column=0, sticky="w", pady=(0, 4))
        self._entry(start_frame, self.start_date_var).grid(row=1, column=0, sticky="ew", padx=(0, 8))
        self._secondary_button(start_frame, "Calendario", lambda: self._open_calendar(self.start_date_var), width=130).grid(
            row=1, column=1, sticky="e"
        )

        end_frame = ctk.CTkFrame(card, fg_color="transparent")
        end_frame.grid(row=1, column=1, sticky="ew", padx=(10, 18), pady=(8, 14))
        end_frame.grid_columnconfigure(0, weight=1)
        self._label(end_frame, "Data final").grid(row=0, column=0, sticky="w", pady=(0, 4))
        self._entry(end_frame, self.end_date_var).grid(row=1, column=0, sticky="ew", padx=(0, 8))
        self._secondary_button(end_frame, "Calendario", lambda: self._open_calendar(self.end_date_var), width=130).grid(
            row=1, column=1, sticky="e"
        )

    def _build_execution_section(self, parent) -> None:
        card = self._section(parent, "Execucao", 3)
        card.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(card, textvariable=self.status_var, text_color=MUTED_TEXT, font=("Segoe UI", 13)).grid(
            row=1, column=0, sticky="w", padx=18, pady=(8, 8)
        )
        self.progress_bar = ctk.CTkProgressBar(
            card,
            height=16,
            progress_color=BUTTON_BG,
            fg_color=SOFT_RED,
            corner_radius=12,
        )
        self.progress_bar.grid(row=2, column=0, sticky="ew", padx=18, pady=(0, 14))
        self.progress_bar.set(0)

        actions = ctk.CTkFrame(card, fg_color="transparent")
        actions.grid(row=3, column=0, sticky="ew", padx=18, pady=(4, 18))
        actions.grid_columnconfigure((0, 1, 2), weight=1)

        self.start_button = self._primary_button(actions, "Iniciar robo", self.start_processing)
        self.start_button.grid(row=0, column=0, sticky="ew", padx=(0, 8))
        self.manual_button = self._secondary_button(actions, "Converter CSV manual", self.convert_csv_manual)
        self.manual_button.grid(row=0, column=1, sticky="ew", padx=8)
        self.stop_button = self._secondary_button(actions, "Parar", self.stop_processing)
        self.stop_button.grid(row=0, column=2, sticky="ew", padx=(8, 0))

    def _card(self, parent):
        return ctk.CTkFrame(parent, fg_color=CARD_BG, border_color=CARD_BORDER, border_width=1, corner_radius=24)

    def _section(self, parent, title: str, row: int):
        card = self._card(parent)
        card.grid(row=row, column=0, sticky="ew", pady=(0, 14))
        ctk.CTkLabel(card, text=title, text_color=PRIMARY_TEXT, font=("Segoe UI", 16, "bold")).grid(
            row=0, column=0, columnspan=4, sticky="w", padx=18, pady=(16, 0)
        )
        return card

    def _label(self, parent, text: str):
        return ctk.CTkLabel(parent, text=text, text_color="#242424", font=("Segoe UI", 12, "bold"))

    def _entry(self, parent, variable, show: str | None = None):
        return ctk.CTkEntry(
            parent,
            textvariable=variable,
            show=show,
            height=42,
            fg_color="#ffffff",
            border_color=CARD_BORDER,
            border_width=1,
            corner_radius=12,
            text_color="#202020",
            font=("Segoe UI", 13),
        )

    def _primary_button(self, parent, text: str, command, width: int | None = None):
        return ctk.CTkButton(
            parent,
            text=text,
            command=command,
            width=width or 180,
            height=42,
            fg_color=BUTTON_BG,
            hover_color=BUTTON_ACTIVE_BG,
            text_color="#ffffff",
            corner_radius=14,
            font=("Segoe UI", 13, "bold"),
        )

    def _secondary_button(self, parent, text: str, command, width: int | None = None):
        return ctk.CTkButton(
            parent,
            text=text,
            command=command,
            width=width or 180,
            height=42,
            fg_color="#ffffff",
            hover_color=SOFT_RED,
            border_color=CARD_BORDER,
            border_width=1,
            text_color=PRIMARY_TEXT,
            corner_radius=14,
            font=("Segoe UI", 13, "bold"),
        )

    def log(self, message: str) -> None:
        self.log_queue.put(f"[{datetime.now().strftime('%H:%M:%S')}] {message}")

    def _poll_log_queue(self) -> None:
        latest_message = None
        try:
            while True:
                latest_message = self.log_queue.get_nowait()
        except queue.Empty:
            pass
        if latest_message:
            clean_message = re.sub(r"^\[\d{2}:\d{2}:\d{2}\]\s*", "", latest_message)
            self.status_var.set(clean_message)
        self.after(200, self._poll_log_queue)

    def choose_output_dir(self) -> None:
        folder = filedialog.askdirectory(title="Selecione a pasta de salvamento")
        if folder:
            self.output_dir_var.set(folder)

    def _open_calendar(self, target_var: ctk.StringVar) -> None:
        try:
            selected = parse_ptbr_date(target_var.get())
        except Exception:
            selected = datetime.now()

        popup = ctk.CTkToplevel(self)
        popup.title("Selecionar data")
        popup.geometry("320x330")
        popup.resizable(False, False)
        popup.configure(fg_color=MAIN_BG)
        popup.transient(self)
        popup.grab_set()

        current = {"year": selected.year, "month": selected.month}

        header = ctk.CTkFrame(popup, fg_color="transparent")
        header.pack(fill="x", padx=16, pady=(16, 8))
        title_var = ctk.StringVar()

        def render_calendar() -> None:
            for widget in days.winfo_children():
                widget.destroy()
            title_var.set(datetime(current["year"], current["month"], 1).strftime("%B/%Y").upper())
            week_days = ["S", "T", "Q", "Q", "S", "S", "D"]
            for col, label in enumerate(week_days):
                ctk.CTkLabel(days, text=label, text_color=MUTED_TEXT, font=("Segoe UI", 11, "bold")).grid(row=0, column=col, padx=2, pady=2)
            first = datetime(current["year"], current["month"], 1)
            start_col = first.weekday()
            if current["month"] == 12:
                next_month = datetime(current["year"] + 1, 1, 1)
            else:
                next_month = datetime(current["year"], current["month"] + 1, 1)
            total_days = (next_month - first).days
            row = 1
            col = start_col
            for day in range(1, total_days + 1):
                chosen = datetime(current["year"], current["month"], day)
                button = ctk.CTkButton(
                    days,
                    text=str(day),
                    width=36,
                    height=32,
                    corner_radius=10,
                    fg_color=BUTTON_BG if chosen.date() == selected.date() else "#ffffff",
                    hover_color=SOFT_RED,
                    text_color="#ffffff" if chosen.date() == selected.date() else "#202020",
                    border_color=CARD_BORDER,
                    border_width=1,
                    command=lambda date_value=chosen: choose_date(date_value),
                )
                button.grid(row=row, column=col, padx=2, pady=2)
                col += 1
                if col > 6:
                    col = 0
                    row += 1

        def previous_month() -> None:
            if current["month"] == 1:
                current["month"] = 12
                current["year"] -= 1
            else:
                current["month"] -= 1
            render_calendar()

        def next_month() -> None:
            if current["month"] == 12:
                current["month"] = 1
                current["year"] += 1
            else:
                current["month"] += 1
            render_calendar()

        def choose_date(date_value: datetime) -> None:
            target_var.set(date_value.strftime("%d/%m/%Y"))
            popup.destroy()

        self._secondary_button(header, "<", previous_month, width=44).pack(side="left")
        ctk.CTkLabel(header, textvariable=title_var, text_color=PRIMARY_TEXT, font=("Segoe UI", 14, "bold")).pack(side="left", expand=True)
        self._secondary_button(header, ">", next_month, width=44).pack(side="right")

        days = ctk.CTkFrame(popup, fg_color="transparent")
        days.pack(padx=16, pady=8)
        render_calendar()

    def _validate_inputs(self) -> bool:
        if not self.username_var.get().strip():
            messagebox.showwarning("Validacao", "Informe o usuario do Coral.")
            return False
        if not self.password_var.get().strip():
            messagebox.showwarning("Validacao", "Informe a senha do Coral.")
            return False
        try:
            start = parse_ptbr_date(self.start_date_var.get())
            end = parse_ptbr_date(self.end_date_var.get())
        except Exception:
            messagebox.showwarning("Validacao", "Informe as datas no formato dd/mm/aaaa.")
            return False
        if end < start:
            messagebox.showwarning("Validacao", "A data final nao pode ser menor que a data inicial.")
            return False
        output_dir = Path(self.output_dir_var.get().strip())
        if not output_dir.exists():
            messagebox.showwarning("Validacao", "A pasta de salvamento nao existe.")
            return False
        return True

    def registrar_abertura(self) -> None:
        try:
            usuario_coral = self.username_var.get().strip()
            usuario_windows = os.environ.get("USERNAME", "").strip() or getpass.getuser().strip()
            usuario_registro = usuario_coral or usuario_windows or "usuario_desconhecido"
            data = {
                "entry.1320712185": "Robo Relatorio Coral",
                "entry.1823299431": usuario_registro,
                "entry.1825360926": datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
            }
            requests.post(URL_PING_ABERTURA, data=data, timeout=5)
            self.log(f"Abertura registrada. Usuario: {usuario_registro}.")
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

    def start_processing(self) -> None:
        if not self._validate_inputs():
            return
        if self.processing_thread is not None and self.processing_thread.is_alive():
            messagebox.showinfo("Execucao", "O robo ja esta em execucao.")
            return
        self.stop_requested = False
        self.status_var.set("Iniciando processamento...")
        self.progress_bar.set(0.03)
        self._update_action_buttons()
        self.processing_thread = threading.Thread(target=self._run_processing, daemon=True)
        self.processing_thread.start()

    def stop_processing(self) -> None:
        self.stop_requested = True
        self.status_var.set("Parada solicitada...")
        self.log("Parada solicitada pelo usuario.")

    def convert_csv_manual(self) -> None:
        csv_file = filedialog.askopenfilename(
            title="Selecione o CSV exportado pelo Coral",
            filetypes=[("CSV", "*.csv"), ("Todos os arquivos", "*.*")],
        )
        if not csv_file:
            return
        try:
            self._validate_dates_only()
            self.status_var.set("Convertendo CSV manual...")
            self.progress_bar.set(0.45)
            result = convert_coral_csv_to_xlsx(
                Path(csv_file),
                Path(self.output_dir_var.get().strip()),
                self.start_date_var.get().strip(),
                self.end_date_var.get().strip(),
            )
            self.progress_bar.set(1)
            self.status_var.set("CSV convertido com sucesso.")
            self.log(f"CSV convertido: {result.source_csv}")
            self.log(f"Excel gerado: {result.output_xlsx}")
            messagebox.showinfo("Conversao concluida", f"Arquivo gerado:\n{result.output_xlsx}")
        except Exception as exc:
            self.progress_bar.set(0)
            self.status_var.set("Falha ao converter CSV.")
            self.log(f"Erro na conversao manual: {exc}")
            messagebox.showerror("Erro na conversao", str(exc))

    def _validate_dates_only(self) -> None:
        start = parse_ptbr_date(self.start_date_var.get())
        end = parse_ptbr_date(self.end_date_var.get())
        if end < start:
            raise ValueError("A data final nao pode ser menor que a data inicial.")

    def _run_processing(self) -> None:
        try:
            self.progress_bar.set(0.08)
            self.registrar_abertura()
            self.progress_bar.set(0.14)
            if not self.verificar_chave():
                self.status_var.set("Robo bloqueado remotamente.")
                self.log("Execucao bloqueada pela validacao remota.")
                self.progress_bar.set(0)
                return

            self.progress_bar.set(0.22)
            self.download_dir.mkdir(parents=True, exist_ok=True)
            self._clean_download_dir()
            self.driver = self._create_driver()
            self.progress_bar.set(0.35)
            self._login_coral()
            self.progress_bar.set(0.52)
            csv_path = self._navigate_and_emit_report()
            self.progress_bar.set(0.82)
            result = convert_coral_csv_to_xlsx(
                csv_path,
                Path(self.output_dir_var.get().strip()),
                self.start_date_var.get().strip(),
                self.end_date_var.get().strip(),
            )
            self.progress_bar.set(1)
            self.status_var.set("Relatorio convertido com sucesso.")
            self.log(f"Arquivo final gerado: {result.output_xlsx}")
            if not self.keep_csv_var.get():
                try:
                    result.source_csv.unlink(missing_ok=True)
                    self.log("CSV original removido apos conversao.")
                except Exception as exc:
                    self.log(f"Nao foi possivel remover CSV original: {exc}")
            messagebox.showinfo("Concluido", f"Relatorio gerado com sucesso:\n{result.output_xlsx}")
        except Exception as exc:
            self.progress_bar.set(0)
            self.status_var.set("Falha na execucao.")
            self.log(f"ERRO: {exc}")
            messagebox.showerror("Erro na execucao", str(exc))
        finally:
            self._close_driver()
            self._update_action_buttons()

    def _create_driver(self):
        options = webdriver.ChromeOptions()
        options.add_argument("--headless=new")
        options.add_argument("--start-maximized")
        options.add_argument("--disable-gpu")
        options.add_argument("--no-sandbox")
        options.add_argument("--disable-dev-shm-usage")
        options.add_argument("--log-level=3")
        prefs = {
            "download.default_directory": str(self.download_dir),
            "download.prompt_for_download": False,
            "download.directory_upgrade": True,
            "safebrowsing.enabled": True,
        }
        options.add_experimental_option("prefs", prefs)
        options.add_experimental_option("excludeSwitches", ["enable-logging"])
        self.log("Criando sessao do Chrome...")
        return webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)

    def _wait_clickable(self, xpath: str, description: str, timeout: int = 30):
        self.log(f"Aguardando {description} ficar disponivel...")
        return WebDriverWait(self.driver, timeout).until(EC.element_to_be_clickable((By.XPATH, xpath)))

    def _wait_visible(self, xpath: str, description: str, timeout: int = 30):
        self.log(f"Aguardando {description} aparecer...")
        return WebDriverWait(self.driver, timeout).until(EC.visibility_of_element_located((By.XPATH, xpath)))

    def _wait_present(self, xpath: str, description: str, timeout: int = 30):
        self.log(f"Aguardando {description} existir...")
        return WebDriverWait(self.driver, timeout).until(EC.presence_of_element_located((By.XPATH, xpath)))

    def _wait_login_completed(self, timeout: int = 90) -> None:
        self.log("Aguardando confirmacao de login pelo dashboard do Coral...")
        try:
            WebDriverWait(self.driver, timeout).until(EC.url_to_be(URL_DASHBOARD_POS_LOGIN))
        except Exception as exc:
            current_url = self.driver.current_url
            raise RuntimeError(
                "Login nao foi confirmado dentro do tempo limite. "
                f"URL atual: {current_url}"
            ) from exc
        self.log("Login confirmado no dashboard do Coral.")

    def _safe_click(self, xpath: str, description: str, timeout: int = 30) -> None:
        last_error = None
        for attempt in range(1, 4):
            try:
                element = self._wait_clickable(xpath, description, timeout)
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
        raise RuntimeError(f"Nao foi possivel clicar em {description}: {last_error}")

    def _safe_click_existing(self, element, description: str) -> None:
        self.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", element)
        time.sleep(0.2)
        try:
            element.click()
        except Exception:
            self.driver.execute_script("arguments[0].click();", element)
        self.log(f"Clique OK em {description}.")

    def _safe_type(self, xpath: str, text: str, description: str, timeout: int = 30, secret: bool = False) -> None:
        last_error = None
        for attempt in range(1, 4):
            try:
                element = self._wait_clickable(xpath, description, timeout)
                self.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", element)
                element.click()
                element.send_keys(Keys.CONTROL, "a")
                element.send_keys(Keys.DELETE)
                element.send_keys(text)
                self.log(f"Texto digitado em {description}: {'*' * len(text) if secret else text}")
                return
            except Exception as exc:
                last_error = exc
                self.log(f"Digitacao falhou em {description} ({attempt}/3): {exc}")
                time.sleep(1)
        raise RuntimeError(f"Nao foi possivel preencher {description}: {last_error}")

    def _login_coral(self) -> None:
        self.log("Abrindo Coral...")
        self.driver.get(URL_CORAL)
        self._safe_type(XPATH_LOGIN, self.username_var.get().strip(), "campo usuario")
        self._safe_type(XPATH_SENHA, self.password_var.get().strip(), "campo senha", secret=True)
        self._safe_click(XPATH_ENTRAR, "botao Entrar")
        self.log("Login enviado.")
        self._wait_login_completed()

    def _navigate_and_emit_report(self) -> Path:
        self.log("Acessando pagina de relatorios...")
        self.driver.get(URL_RELATORIOS)
        self._wait_present("//foco-app", "aplicacao Coral", timeout=45)
        self._wait_visible(
            "//*[contains(normalize-space(.), 'Selecione uma categoria') or contains(normalize-space(.), 'REPORT_GROUP')]",
            "tela de relatorios",
            timeout=45,
        )

        self._select_report_category("Financeiro")
        self._select_report_name(REPORT_NAME)
        self._select_coral_period(self.start_date_var.get().strip(), self.end_date_var.get().strip())
        self._click_export_report()
        return self._wait_for_csv_download()

    def _select_report_category(self, category: str) -> None:
        self.log(f"Selecionando categoria do relatorio: {category}")
        dropdown_xpath = (
            "(//foco-dropdown[.//button//*[contains(normalize-space(.), 'Selecione uma categoria') "
            "or contains(normalize-space(.), 'REPORT_GROUP')]])[1]//button[contains(@class, 'dropdown-toggle')]"
        )
        self._safe_click(dropdown_xpath, "lista de categorias", timeout=45)
        option_xpath = (
            f"//div[contains(@class, 'dropdown-menu') and contains(@class, 'show')]"
            f"//button[contains(@class, 'dropdown-item')][.//span[normalize-space()='{category}']]"
        )
        self._safe_click(option_xpath, f"categoria {category}", timeout=30)

    def _select_report_name(self, report_name: str) -> None:
        self.log(f"Selecionando relatorio: {report_name}")
        dropdown_xpath = (
            "(//foco-dropdown[.//button//*[contains(normalize-space(.), 'Selecione um relat')]])[1]"
            "//button[contains(@class, 'dropdown-toggle')]"
        )
        self._safe_click(dropdown_xpath, "lista de relatorios", timeout=45)

        search_xpath = "//div[contains(@class, 'dropdown-menu') and contains(@class, 'show')]//input[@placeholder='Buscar']"
        try:
            self._safe_type(search_xpath, report_name, "busca do relatorio", timeout=8)
            time.sleep(1)
        except Exception as exc:
            self.log(f"Busca do relatorio nao ficou disponivel, tentando lista direta: {exc}")

        option_xpath = (
            f"//div[contains(@class, 'dropdown-menu') and contains(@class, 'show')]"
            f"//button[contains(@class, 'dropdown-item')][.//span[normalize-space()='{report_name}']]"
        )
        self._safe_click(option_xpath, report_name, timeout=45)

    def _select_coral_period(self, start_date: str, end_date: str) -> None:
        start = parse_ptbr_date(start_date)
        end = parse_ptbr_date(end_date)
        self.log(f"Selecionando periodo no Coral: {start_date} ate {end_date}")
        self._safe_click("//input[@id='dateRange' or @name='dp']", "campo de periodo", timeout=45)
        self._click_datepicker_day(start, "data inicial")
        self._click_datepicker_day(end, "data final")
        time.sleep(1)

    def _click_datepicker_day(self, target_date: datetime, description: str) -> None:
        target_label = f"{target_date.day}/{target_date.month}/{target_date.year}"
        target_xpath = f"//ngb-datepicker//div[@role='gridcell' and @aria-label='{target_label}' and not(contains(@class, 'hidden'))]"

        for attempt in range(1, 25):
            elements = self.driver.find_elements(By.XPATH, target_xpath)
            visible_elements = [element for element in elements if element.is_displayed()]
            if visible_elements:
                self._safe_click_existing(visible_elements[0], f"{description} {target_label}")
                return
            self._move_datepicker_towards(target_date)
            time.sleep(0.4)
            self.log(f"Procurando {description} {target_label} no calendario ({attempt}/24)...")

        raise RuntimeError(f"Nao foi possivel selecionar {description}: {target_label}")

    def _move_datepicker_towards(self, target_date: datetime) -> None:
        visible_months = self.driver.find_elements(By.XPATH, "//ngb-datepicker//div[contains(@class, 'ngb-dp-month-name')]")
        month_dates: list[datetime] = []
        month_names = {
            "janeiro": 1,
            "fevereiro": 2,
            "marco": 3,
            "março": 3,
            "abril": 4,
            "maio": 5,
            "junho": 6,
            "julho": 7,
            "agosto": 8,
            "setembro": 9,
            "outubro": 10,
            "novembro": 11,
            "dezembro": 12,
        }
        for element in visible_months:
            text = element.text.strip().lower()
            parts = text.split()
            if len(parts) >= 2 and parts[0] in month_names and parts[-1].isdigit():
                month_dates.append(datetime(int(parts[-1]), month_names[parts[0]], 1))

        if not month_dates:
            raise RuntimeError("Nao foi possivel identificar o mes atual do calendario.")

        first_visible = min(month_dates)
        last_visible = max(month_dates)
        target_month = datetime(target_date.year, target_date.month, 1)
        if target_month < first_visible:
            self._safe_click("//ngb-datepicker//button[@title='Previous month' or @aria-label='Previous month']", "mes anterior", timeout=10)
        elif target_month > last_visible:
            self._safe_click("//ngb-datepicker//button[@title='Next month' or @aria-label='Next month']", "proximo mes", timeout=10)
        else:
            raise RuntimeError("Mes visivel, mas o dia nao ficou disponivel para clique.")

    def _click_export_report(self) -> None:
        self.log("Tentando iniciar exportacao do relatorio...")
        try:
            self._safe_click(XPATH_GERAR_RELATORIO, "botao Gerar relatorio", timeout=30)
            return
        except Exception as exc:
            self.log(f"XPath principal de Gerar relatorio falhou, tentando fallbacks: {exc}")

        export_candidates = [
            "//button[.//*[contains(normalize-space(.), 'CSV')] or contains(normalize-space(.), 'CSV')]",
            "//button[.//*[contains(normalize-space(.), 'Gerar relatório')] or contains(normalize-space(.), 'Gerar relatório')]",
            "//button[.//*[contains(normalize-space(.), 'Gerar relatorio')] or contains(normalize-space(.), 'Gerar relatorio')]",
            "//button[contains(normalize-space(.), 'Exportar')]",
            "//button[contains(normalize-space(.), 'Gerar')]",
            "//button[contains(normalize-space(.), 'Baixar')]",
            "//button[contains(normalize-space(.), 'Download')]",
            "//a[contains(normalize-space(.), 'CSV') or contains(normalize-space(.), 'Exportar') or contains(normalize-space(.), 'Baixar')]",
        ]
        last_error = None
        for xpath in export_candidates:
            try:
                self._safe_click(xpath, "botao de exportacao/CSV", timeout=8)
                return
            except Exception as exc:
                last_error = exc
        raise RuntimeError(
            "Relatorio e periodo foram selecionados, mas o botao de gerar/exportar CSV ainda nao foi localizado. "
            f"Ultima falha: {last_error}"
        )

    def _wait_for_csv_download(self, timeout: int = 120) -> Path:
        self.log("Aguardando download do CSV...")
        deadline = time.time() + timeout
        last_size = -1
        stable_since = None
        latest: Path | None = None
        while time.time() < deadline:
            csv_files = sorted(self.download_dir.glob("*.csv"), key=lambda path: path.stat().st_mtime, reverse=True)
            partial_files = list(self.download_dir.glob("*.crdownload"))
            if csv_files:
                latest = csv_files[0]
                size = latest.stat().st_size
                if size == last_size and not partial_files:
                    if stable_since is None:
                        stable_since = time.time()
                    if time.time() - stable_since >= 2:
                        self.log(f"CSV baixado: {latest}")
                        return latest
                else:
                    stable_since = None
                    last_size = size
            time.sleep(1)
        raise TimeoutError("O CSV nao foi baixado dentro do tempo limite.")

    def _clean_download_dir(self) -> None:
        for pattern in ("*.csv", "*.crdownload"):
            for file_path in self.download_dir.glob(pattern):
                try:
                    file_path.unlink()
                except Exception:
                    pass

    def _close_driver(self) -> None:
        if self.driver is None:
            return
        try:
            self.driver.quit()
        except Exception:
            pass
        self.driver = None

    def _update_action_buttons(self) -> None:
        running = self.processing_thread is not None and self.processing_thread.is_alive()
        self.start_button.configure(state="disabled" if running else "normal")
        self.manual_button.configure(state="disabled" if running else "normal")
        self.stop_button.configure(state="normal" if running else "disabled")


if __name__ == "__main__":
    app = RoboRelatorioCoralApp()
    app.mainloop()
