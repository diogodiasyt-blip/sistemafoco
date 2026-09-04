from __future__ import annotations

import getpass
import os
import queue
import re
import tempfile
import threading
import time
from dataclasses import dataclass
from datetime import datetime, timedelta, timezone
from zoneinfo import ZoneInfo, ZoneInfoNotFoundError
from pathlib import Path
from tkinter import filedialog, messagebox

import customtkinter as ctk
import pandas as pd
import requests
from PIL import Image


APP_TITLE = "Robo de Relatorio Coral"
APP_GEOMETRY = "660x470"

URL_VALIDACAO = "https://raw.githubusercontent.com/diogodiasyt-blip/validacaofoco/refs/heads/main/chave"
URL_PING_ABERTURA = "https://docs.google.com/forms/d/e/1FAIpQLScmxNbTO-vXw0LEOKIyEhSpIl9aTbw8x5hnEI5VY2eVMRh5gQ/formResponse"
URL_CORAL = "https://coral.aluguefoco.com.br/login"
URL_DASHBOARD_POS_LOGIN = "https://coral.aluguefoco.com.br/precificacao/dashboard"
URL_RELATORIOS = "https://coral.aluguefoco.com.br/relatorios"
CORAL_USERNAME = os.getenv("FOCO_CORAL_USERNAME", "codp")
CORAL_PASSWORD = os.getenv("FOCO_CORAL_PASSWORD", "Foco@2026")

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
SOFT_RED = "#fff1f0"
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
            dataframe = pd.read_csv(csv_path, sep=",", decimal=".", encoding=encoding, dtype=str)
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


CORAL_API_BASE_URL = "https://servicescoral.aluguefoco.com.br"
CORAL_REPORT_KEY = "finance-brokers-efficiency"


class ReportCancelled(Exception):
    pass


def report_api_params(start_date: str, end_date: str) -> dict[str, str]:
    start, end = parse_ptbr_date(start_date), parse_ptbr_date(end_date)
    if end < start:
        raise ValueError("A data final nao pode ser menor que a data inicial.")
    if (start.year, start.month) != (end.year, end.month):
        raise ValueError("Selecione datas dentro do mesmo mes, como no portal de faturamento.")
    try:
        zone = ZoneInfo("America/Sao_Paulo")
    except ZoneInfoNotFoundError:
        if start.year < 2020:
            raise ValueError("Fuso historico indisponivel para o periodo selecionado.") from None
        zone = timezone(timedelta(hours=-3))
    def iso(value):
        return value.replace(tzinfo=zone).astimezone(timezone.utc).isoformat(timespec="milliseconds").replace("+00:00", "Z")
    return {"periodStart": start.strftime("%Y-%m-%d"), "periodEnd": end.strftime("%Y-%m-%d"),
            "start": iso(start), "end": iso(end)}


def download_coral_report_api(start_date, end_date, download_dir, *, login=None,
                              password=None, log=lambda message: None,
                              should_stop=lambda: False, session_factory=requests.Session):
    params = report_api_params(start_date, end_date)
    def check_stop():
        if should_stop():
            raise ReportCancelled("Execucao interrompida pelo usuario.")
    def request(session, method, path, **kwargs):
        check_stop()
        try:
            response = session.request(method, CORAL_API_BASE_URL + path,
                                       allow_redirects=False, **kwargs)
        except requests.RequestException:
            raise RuntimeError("Coral API: falha de conexao ou tempo limite; tente novamente.") from None
        check_stop()
        if not 200 <= response.status_code < 300:
            raise RuntimeError(f"Coral API: HTTP {response.status_code} na etapa {path.split('/')[2]}.")
        return response
    def json_data(response):
        try:
            return response.json()["data"]
        except (ValueError, KeyError, TypeError):
            raise RuntimeError("Coral API: resposta JSON invalida ou sem data.") from None

    with session_factory() as session:
        log("Coral API: autenticando.")
        auth = json_data(request(session, "POST", "/api/auth/login",
                                 json={"login": login if login is not None else CORAL_USERNAME,
                                       "password": password if password is not None else CORAL_PASSWORD},
                                 headers={"Accept": "application/json"}, timeout=(10, 30)))
        token = auth.get("token") if isinstance(auth, dict) else None
        if not isinstance(token, str) or not token.strip():
            raise RuntimeError("Coral API: autenticacao nao retornou token.")
        headers = {"Accept": "application/json", "Authorization": f"Bearer {token}"}
        log("Coral API: gerando relatorio Eficiência Brokers Financeiro.")
        data = json_data(request(session, "GET",
            f"/api/analytics/report/generate/{CORAL_REPORT_KEY}",
            params=params, headers=headers, timeout=(10, 60)))
        filename = data[0] if isinstance(data, list) and data else None
        if not isinstance(filename, str) or not re.fullmatch(r"[a-zA-Z0-9._-]+\.csv", filename, re.IGNORECASE):
            raise RuntimeError("Coral API: nome do CSV ausente ou invalido.")
        log("Coral API: baixando CSV.")
        response = request(session, "GET",
            f"/api/analytics/bitstream/{CORAL_REPORT_KEY}/{filename}",
            headers={**headers, "Accept": "text/csv, application/octet-stream"}, timeout=(10, 60))
        content = response.content
        sample = content[:2048].decode("utf-8-sig", errors="replace").lstrip()
        content_type = response.headers.get("Content-Type", "").lower()
        if (not content or not any(separator in sample for separator in ",;\t")
                or sample.startswith(("<", "{", "[")) or "json" in content_type or "html" in content_type):
            raise RuntimeError("Coral API: resposta de download vazia ou nao e um CSV valido.")
        check_stop()
        # Cada execucao tem sua propria pasta: nao reutiliza nem apaga CSV de outra instancia.
        run_dir = Path(tempfile.mkdtemp(prefix="coral_report_", dir=download_dir))
        csv_path = run_dir / filename
        csv_path.write_bytes(content)
        log("Coral API: download confirmado.")
        return csv_path


class RoboRelatorioCoralApp(ctk.CTk):
    def __init__(self) -> None:
        super().__init__()
        ctk.set_appearance_mode("light")
        ctk.set_default_color_theme("blue")

        self.title(APP_TITLE)
        self.geometry(APP_GEOMETRY)
        self.minsize(600, 390)
        self.configure(fg_color=MAIN_BG)

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

        container = ctk.CTkFrame(self, fg_color="transparent")
        container.grid(row=0, column=0, sticky="nsew", padx=16, pady=14)
        container.grid_columnconfigure(0, weight=1)

        self._build_header(container)
        self._build_period_section(container)
        self._build_execution_section(container)

    def _build_header(self, parent) -> None:
        card = self._card(parent)
        card.grid(row=0, column=0, sticky="ew", pady=(0, 10))
        card.grid_columnconfigure(1, weight=1)

        logo_loaded = False
        for candidate in resolve_logo_candidates():
            try:
                if candidate.exists():
                    image = Image.open(candidate)
                    self.logo_image = ctk.CTkImage(light_image=image, dark_image=image, size=(82, 40))
                    ctk.CTkLabel(card, image=self.logo_image, text="").grid(row=0, column=0, rowspan=3, padx=(14, 16), pady=12)
                    logo_loaded = True
                    break
            except Exception:
                continue

        if not logo_loaded:
            ctk.CTkLabel(card, text="foco", text_color=PRIMARY_TEXT, font=("Segoe UI", 26, "bold")).grid(
                row=0, column=0, rowspan=3, padx=(14, 16), pady=12
            )

        ctk.CTkLabel(
            card,
            text="Relatorio Coral",
            text_color=PRIMARY_TEXT,
            font=("Segoe UI", 22, "bold"),
            anchor="w",
        ).grid(row=0, column=1, sticky="ew", padx=(0, 16), pady=(14, 2))
        ctk.CTkLabel(
            card,
            text="Selecione o periodo e gere o Excel corrigido.",
            text_color=MUTED_TEXT,
            font=("Segoe UI", 13),
            anchor="w",
        ).grid(row=1, column=1, sticky="ew", padx=(0, 20))
        ctk.CTkLabel(
            card,
            text="OPERACAO DE RELATORIOS",
            text_color="#b65748",
            font=("Segoe UI", 11, "bold"),
            anchor="w",
        ).grid(row=2, column=1, sticky="ew", padx=(0, 16), pady=(6, 14))

    def _build_period_section(self, parent) -> None:
        card = self._section(parent, "Periodo do relatorio", 1)
        card.grid_columnconfigure((0, 1), weight=1)

        start_frame = ctk.CTkFrame(card, fg_color="transparent")
        start_frame.grid(row=1, column=0, sticky="ew", padx=(14, 8), pady=(6, 12))
        start_frame.grid_columnconfigure(0, weight=1)
        self._label(start_frame, "Data inicial").grid(row=0, column=0, sticky="w", pady=(0, 4))
        self._entry(start_frame, self.start_date_var).grid(row=1, column=0, sticky="ew", padx=(0, 8))
        self._secondary_button(start_frame, "Calendario", lambda: self._open_calendar(self.start_date_var), width=112).grid(
            row=1, column=1, sticky="e"
        )

        end_frame = ctk.CTkFrame(card, fg_color="transparent")
        end_frame.grid(row=1, column=1, sticky="ew", padx=(8, 14), pady=(6, 12))
        end_frame.grid_columnconfigure(0, weight=1)
        self._label(end_frame, "Data final").grid(row=0, column=0, sticky="w", pady=(0, 4))
        self._entry(end_frame, self.end_date_var).grid(row=1, column=0, sticky="ew", padx=(0, 8))
        self._secondary_button(end_frame, "Calendario", lambda: self._open_calendar(self.end_date_var), width=112).grid(
            row=1, column=1, sticky="e"
        )

    def _build_execution_section(self, parent) -> None:
        card = self._section(parent, "Execucao", 2)
        card.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(card, textvariable=self.status_var, text_color=MUTED_TEXT, font=("Segoe UI", 13)).grid(
            row=1, column=0, sticky="w", padx=14, pady=(6, 6)
        )
        self.progress_bar = ctk.CTkProgressBar(
            card,
            height=16,
            progress_color=BUTTON_BG,
            fg_color=SOFT_RED,
            corner_radius=12,
        )
        self.progress_bar.grid(row=2, column=0, sticky="ew", padx=14, pady=(0, 12))
        self.progress_bar.set(0)

        actions = ctk.CTkFrame(card, fg_color="transparent")
        actions.grid(row=3, column=0, sticky="ew", padx=14, pady=(2, 14))
        actions.grid_columnconfigure((0, 1, 2), weight=1)

        self.start_button = self._primary_button(actions, "Iniciar robo", self.start_processing)
        self.start_button.grid(row=0, column=0, sticky="ew", padx=(0, 8))
        self.manual_button = self._secondary_button(actions, "Converter CSV manual", self.convert_csv_manual)
        self.manual_button.grid(row=0, column=1, sticky="ew", padx=8)
        self.stop_button = self._secondary_button(actions, "Parar", self.stop_processing)
        self.stop_button.grid(row=0, column=2, sticky="ew", padx=(8, 0))

    def _card(self, parent):
        return ctk.CTkFrame(parent, fg_color=CARD_BG, border_color=CARD_BORDER, border_width=1, corner_radius=18)

    def _section(self, parent, title: str, row: int):
        card = self._card(parent)
        card.grid(row=row, column=0, sticky="ew", pady=(0, 10))
        ctk.CTkLabel(card, text=title, text_color=PRIMARY_TEXT, font=("Segoe UI", 16, "bold")).grid(
            row=0, column=0, columnspan=4, sticky="w", padx=14, pady=(12, 0)
        )
        return card

    def _label(self, parent, text: str):
        return ctk.CTkLabel(parent, text=text, text_color="#242424", font=("Segoe UI", 12, "bold"))

    def _entry(self, parent, variable, show: str | None = None):
        return ctk.CTkEntry(
            parent,
            textvariable=variable,
            show=show,
            height=38,
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
            width=width or 150,
            height=38,
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
            width=width or 150,
            height=38,
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
        if not CORAL_USERNAME or not CORAL_PASSWORD:
            messagebox.showwarning(
                "Credenciais Coral",
                "Defina FOCO_CORAL_USERNAME e FOCO_CORAL_PASSWORD nas variaveis de ambiente antes de executar.",
            )
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
        try:
            report_api_params(self.start_date_var.get(), self.end_date_var.get())
        except ValueError as exc:
            messagebox.showwarning("Validacao", str(exc))
            return False
        output_dir = Path(self.output_dir_var.get().strip())
        if not output_dir.exists():
            messagebox.showwarning("Validacao", "A pasta de salvamento nao existe.")
            return False
        return True

    def registrar_abertura(self) -> None:
        try:
            usuario_windows = os.environ.get("USERNAME", "").strip() or getpass.getuser().strip()
            usuario_registro = CORAL_USERNAME or usuario_windows or "usuario_desconhecido"
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
            csv_path = download_coral_report_api(
                self.start_date_var.get().strip(), self.end_date_var.get().strip(),
                self.download_dir, log=self.log, should_stop=lambda: self.stop_requested,
            )
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
        except ReportCancelled:
            self.progress_bar.set(0)
            self.status_var.set("Execucao interrompida pelo usuario.")
        except Exception as exc:
            self.progress_bar.set(0)
            self.status_var.set("Falha na execucao.")
            self.log(f"ERRO: {exc}")
            messagebox.showerror("Erro na execucao", str(exc))
        finally:
            self._update_action_buttons()

    def _update_action_buttons(self) -> None:
        running = self.processing_thread is not None and self.processing_thread.is_alive()
        self.start_button.configure(state="disabled" if running else "normal")
        self.manual_button.configure(state="disabled" if running else "normal")
        self.stop_button.configure(state="normal" if running else "disabled")


if __name__ == "__main__":
    app = RoboRelatorioCoralApp()
    app.mainloop()
