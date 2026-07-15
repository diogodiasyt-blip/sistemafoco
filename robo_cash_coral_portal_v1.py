from __future__ import annotations

import csv
import json
import os
import queue
import re
import tempfile
import threading
import time
from dataclasses import dataclass
from datetime import datetime, timedelta
from pathlib import Path
from tkinter import messagebox

import customtkinter as ctk
import requests
from PIL import Image
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from webdriver_manager.chrome import ChromeDriverManager


APP_TITLE = "Atualizar Cash Coral"
APP_GEOMETRY = "760x640"
URL_CORAL_LOGIN = "https://coral.aluguefoco.com.br/login"
URL_CORAL_RELATORIOS = "https://coral.aluguefoco.com.br/relatorios"
REPORT_CATEGORY = "Financeiro"
REPORT_NAME = "Relatorio de Cash"
SUPABASE_URL = "https://pervrszvwgfuzebrkuup.supabase.co"
SUPABASE_ANON_KEY = (
    "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9."
    "eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InBlcnZyc3p2d2dmdXplYnJrdXVwIiwicm9sZSI6ImFub24iLCJpYXQiOjE3ODM5OTcwMTgsImV4cCI6MjA5OTU3MzAxOH0."
    "0wQQtwfw7GA_cbzm4vBQmwIp8G93-FK7md221i_XLKg"
)

XPATH_LOGIN = "/html/body/foco-app/div[1]/foco-login/div/div/div/div/div[2]/form/div[1]/input"
XPATH_PASSWORD = "/html/body/foco-app/div[1]/foco-login/div/div/div/div/div[2]/form/div[2]/input"
XPATH_ENTER = "/html/body/foco-app/div[1]/foco-login/div/div/div/div/div[2]/form/button"
XPATH_GENERATE_REPORT = "/html/body/foco-app/div[1]/foco-analytics-home/div/div/div/div/div/div/div[2]/div[1]/div[3]/button/span"

MAIN_BG = "#f6f4f1"
CARD_BG = "#ffffff"
CARD_BORDER = "#eadfdb"
PRIMARY_TEXT = "#d81919"
MUTED_TEXT = "#5c5c5c"
BUTTON_BG = "#ef1a14"
BUTTON_ACTIVE_BG = "#c91410"
SOFT_RED = "#fff1f0"

STORE_IDS = {
    "AJU10": "aju10",
    "BPS10": "bps10",
    "CGR10": "cgr10",
    "CNF10": "cnf10",
    "CWB10": "cwb10",
    "FLN10": "fln10",
    "FOR10": "for10",
    "GIG10": "gig10",
    "GYN10": "gyn10",
    "JPA10": "jpa10",
    "MCZ10": "mcz10",
    "NAT10": "nat10",
    "NVT10": "nvt10",
    "POA10": "poa10",
    "QNS10": "qns10",
    "REC20": "rec20",
    "RAO10": "rao10",
    "SAO10": "sao10",
    "SAO11": "sao11",
    "SAO12": "sao12",
    "SDU10": "sdu10",
    "SSA10": "ssa10",
    "VCP10": "vcp10",
    "VIX10": "vix10",
}


def resolve_logo_candidates() -> list[Path]:
    candidates: list[Path] = []
    env_logo = os.environ.get("FOCO_LOGO_PNG", "").strip()
    env_assets = os.environ.get("FOCO_ASSETS_DIR", "").strip()
    if env_logo:
        candidates.append(Path(env_logo))
    if env_assets:
        candidates.append(Path(env_assets) / "logo.png")
    base_dir = Path(__file__).resolve().parent
    candidates.extend([
        base_dir.parent / "assets" / "logo.png",
        Path.cwd() / "assets" / "logo.png",
        Path.cwd() / "DESENVOLVIMENTO" / "assets" / "logo.png",
    ])
    return candidates


@dataclass
class CashEntry:
    id: str
    store_id: str
    date: str
    period: str
    contract: str
    amount: float
    origin: str
    source_key: str


def config_path() -> Path:
    root = Path(os.environ.get("APPDATA", tempfile.gettempdir())) / "SistemaFOCO"
    root.mkdir(parents=True, exist_ok=True)
    return root / "cash_coral_portal_config.json"


def load_config() -> dict:
    path = config_path()
    if not path.exists():
        return {}
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except Exception:
        return {}


def save_config(config: dict) -> None:
    config_path().write_text(json.dumps(config, ensure_ascii=False, indent=2), encoding="utf-8")


def parse_ptbr_date(value: str) -> datetime:
    return datetime.strptime(value.strip(), "%d/%m/%Y")


def normalize_header(value: str) -> str:
    replacements = {
        "ç": "c",
        "ã": "a",
        "á": "a",
        "à": "a",
        "â": "a",
        "é": "e",
        "ê": "e",
        "í": "i",
        "ó": "o",
        "ô": "o",
        "õ": "o",
        "ú": "u",
    }
    clean = value.strip().lower()
    for source, target in replacements.items():
        clean = clean.replace(source, target)
    return re.sub(r"[^a-z0-9]+", "_", clean).strip("_")


def parse_money(value: str) -> float:
    text = str(value or "").strip().replace(".", "").replace(",", ".")
    return float(text)


def parse_csv_date(value: str) -> str:
    return parse_ptbr_date(value).strftime("%Y-%m-%d")


def make_source_key(store_code: str, contract: str, date: str, amount: float) -> str:
    cents = round(amount * 100)
    return f"{store_code.lower()}-{contract.lower()}-{date}-{cents}"


def first_filled(record: dict[str, str], keys: list[str]) -> str:
    for key in keys:
        value = str(record.get(key, "") or "").strip()
        if value:
            return value
    return ""


def payment_store_code(record: dict[str, str]) -> str:
    # Regra operacional: o cash entra na loja que recebeu o dinheiro, nao na loja de origem do contrato.
    payment_store = first_filled(
        record,
        ["loja_pagamento", "loja_de_pagamento", "loja_pgto", "loja_pagto", "filial_pagamento"],
    )
    return (payment_store or record.get("loja", "")).strip().upper()


def read_cash_csv(csv_path: Path) -> list[CashEntry]:
    with csv_path.open("r", encoding="utf-8-sig", newline="") as handle:
        reader = csv.reader(handle)
        rows = list(reader)
    if not rows:
        return []
    headers = [normalize_header(item) for item in rows[0]]
    entries: list[CashEntry] = []
    for row in rows[1:]:
        if not row or all(not item.strip() for item in row):
            continue
        record = {headers[index]: row[index].strip() if index < len(row) else "" for index in range(len(headers))}
        if record.get("tipo", "").lower() != "money":
            continue
        store_code = payment_store_code(record)
        store_id = STORE_IDS.get(store_code)
        contract = (record.get("contrato") or "").upper()
        if not store_id or not contract:
            continue
        date = parse_csv_date(record.get("data_de_criacao", ""))
        amount = parse_money(record.get("total_cash", "0"))
        if amount <= 0:
            continue
        source_key = make_source_key(store_code, contract, date, amount)
        entries.append(
            CashEntry(
                id=f"coral-{source_key}",
                store_id=store_id,
                date=date,
                period=date[:7],
                contract=contract,
                amount=amount,
                origin=record.get("origem", ""),
                source_key=source_key,
            )
        )
    unique = {entry.source_key: entry for entry in entries}
    return list(unique.values())


class PortalCaixaClient:
    def __init__(self, admin_user: str, admin_password: str) -> None:
        self.admin_user = admin_user.strip().lower()
        self.admin_password = admin_password
        self.access_token = ""

    def login(self) -> None:
        response = requests.post(
            f"{SUPABASE_URL}/auth/v1/token?grant_type=password",
            headers={"apikey": SUPABASE_ANON_KEY, "content-type": "application/json"},
            json={"email": f"{self.admin_user}@portal-caixa.local", "password": self.admin_password},
            timeout=30,
        )
        if response.status_code >= 400:
            raise RuntimeError(f"Falha no login do Portal Caixa: {response.text}")
        self.access_token = response.json()["access_token"]

    def headers(self, prefer: str | None = None) -> dict:
        headers = {
            "apikey": SUPABASE_ANON_KEY,
            "authorization": f"Bearer {self.access_token}",
            "content-type": "application/json",
        }
        if prefer:
            headers["prefer"] = prefer
        return headers

    def ignored_keys(self, keys: list[str]) -> set[str]:
        if not keys:
            return set()
        quoted = ",".join(f'"{key}"' for key in keys)
        response = requests.get(
            f"{SUPABASE_URL}/rest/v1/cash_entry_ignored?select=source_key&source_key=in.({quoted})",
            headers=self.headers(),
            timeout=30,
        )
        if response.status_code >= 400:
            raise RuntimeError(f"Falha ao consultar cash ignorado: {response.text}")
        return {row["source_key"] for row in response.json()}

    def apply_entries(self, entries: list[CashEntry]) -> int:
        ignored = self.ignored_keys([entry.source_key for entry in entries])
        created_at = datetime.utcnow().isoformat(timespec="milliseconds") + "Z"
        rows = [
            {
                "id": entry.id,
                "store_id": entry.store_id,
                "kind": "receita",
                "entry_date": entry.date,
                "period": entry.period,
                "category": "Cash Coral",
                "description": entry.contract,
                "amount": entry.amount,
                "bank": None,
                "authorizer": None,
                "receipt_path": None,
                "invoice_path": entry.origin or None,
                "source_key": entry.source_key,
                "status": "baixado",
                "created_by": None,
                "created_at": created_at,
            }
            for entry in entries
            if entry.source_key not in ignored
        ]
        if not rows:
            return 0
        response = requests.post(
            f"{SUPABASE_URL}/rest/v1/cash_entries?on_conflict=id",
            headers=self.headers("resolution=merge-duplicates"),
            json=rows,
            timeout=60,
        )
        if response.status_code >= 400:
            raise RuntimeError(f"Falha ao aplicar cash no Portal Caixa: {response.text}")
        return len(rows)


class CoralCashDownloader:
    def __init__(self, coral_user: str, coral_password: str, visible: bool, log) -> None:
        self.coral_user = coral_user
        self.coral_password = coral_password
        self.visible = visible
        self.log = log
        self.download_dir = Path(tempfile.gettempdir()) / "SistemaFOCO" / "downloads_cash_coral"
        self.driver = None

    def download(self, start: datetime, end: datetime) -> Path:
        self.download_dir.mkdir(parents=True, exist_ok=True)
        self._clean_downloads()
        self.driver = self._create_driver()
        try:
            self._open_reports()
            self._select_category(REPORT_CATEGORY)
            self._select_report(REPORT_NAME)
            self._select_period(start, end)
            self._click_generate()
            return self._wait_csv()
        finally:
            self._close()

    def _create_driver(self):
        options = webdriver.ChromeOptions()
        if not self.visible:
            options.add_argument("--headless=new")
        options.add_argument("--start-maximized")
        options.add_argument("--disable-gpu")
        options.add_argument("--no-sandbox")
        options.add_argument("--disable-dev-shm-usage")
        options.add_argument("--log-level=3")
        options.add_experimental_option(
            "prefs",
            {
                "download.default_directory": str(self.download_dir),
                "download.prompt_for_download": False,
                "download.directory_upgrade": True,
                "safebrowsing.enabled": True,
            },
        )
        options.add_experimental_option("excludeSwitches", ["enable-logging"])
        self.log("Criando sessao do Chrome...")
        return webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)

    def _open_reports(self) -> None:
        self.driver.get(URL_CORAL_RELATORIOS)
        time.sleep(2)
        if "/login" in self.driver.current_url:
            self.driver.get(URL_CORAL_LOGIN)
            self._type(XPATH_LOGIN, self.coral_user, "usuario Coral")
            self._type(XPATH_PASSWORD, self.coral_password, "senha Coral")
            self._click(XPATH_ENTER, "botao Entrar")
            WebDriverWait(self.driver, 90).until(lambda driver: "/login" not in driver.current_url)
            self.driver.get(URL_CORAL_RELATORIOS)
        self._visible(
            "//*[contains(normalize-space(.), 'Selecione uma categoria') or contains(normalize-space(.), 'REPORT_GROUP')]",
            "tela de relatorios",
            60,
        )

    def _select_category(self, category: str) -> None:
        self.log(f"Selecionando categoria: {category}")
        self._click(
            "(//foco-dropdown[.//button//*[contains(normalize-space(.), 'Selecione uma categoria') or contains(normalize-space(.), 'REPORT_GROUP')]])[1]//button[contains(@class, 'dropdown-toggle')]",
            "lista de categorias",
            45,
        )
        self._click(
            f"//div[contains(@class, 'dropdown-menu') and contains(@class, 'show')]//button[contains(@class, 'dropdown-item')][.//span[normalize-space()='{category}']]",
            f"categoria {category}",
            30,
        )

    def _select_report(self, report: str) -> None:
        self.log(f"Selecionando relatorio: {report}")
        self._click(
            "(//foco-dropdown[.//button//*[contains(normalize-space(.), 'Selecione um relat')]])[1]//button[contains(@class, 'dropdown-toggle')]",
            "lista de relatorios",
            45,
        )
        menu_xpath = "//div[contains(@class, 'dropdown-menu') and contains(@class, 'show')]"
        WebDriverWait(self.driver, 20).until(EC.visibility_of_element_located((By.XPATH, menu_xpath)))
        try:
            search = WebDriverWait(self.driver, 5).until(
                EC.element_to_be_clickable((By.XPATH, f"{menu_xpath}//input"))
            )
            search.click()
            search.send_keys(Keys.CONTROL, "a")
            search.send_keys(Keys.DELETE)
            search.send_keys("Cash")
            self.log("Filtro aplicado na lista de relatorios: Cash")
            time.sleep(0.8)
            clicked = self._click_visible_report_option()
            if clicked:
                self.log(f"Clique OK: relatorio {clicked}")
                return
        except Exception as exc:
            self.log(f"Filtro de relatorio indisponivel, tentando scroll: {exc}")

        options = [
            f"//div[contains(@class, 'dropdown-menu') and contains(@class, 'show')]//button[contains(@class, 'dropdown-item')][.//span[normalize-space()='{report}']]",
            "//div[contains(@class, 'dropdown-menu') and contains(@class, 'show')]//button[contains(@class, 'dropdown-item')][.//span[normalize-space()='Relatório de Cash']]",
            "//div[contains(@class, 'dropdown-menu') and contains(@class, 'show')]//button[contains(@class, 'dropdown-item')][.//span[contains(normalize-space(), 'Cash')]]",
        ]
        last_error = None
        for xpath in options:
            try:
                self._click(xpath, report, 20)
                return
            except Exception as exc:
                last_error = exc
        clicked = self.driver.execute_script(
            """
            const wanted = 'cash';
            const menu = document.querySelector('div.dropdown-menu.show');
            if (!menu) return null;
            const scrollables = [menu, ...menu.querySelectorAll('.ps, .ps-content, perfect-scrollbar')];
            for (let step = 0; step < 30; step += 1) {
              const buttons = Array.from(menu.querySelectorAll('button.dropdown-item'));
              const button = buttons.find((item) => (item.innerText || '').trim().toLowerCase().includes(wanted));
              if (button) {
                button.scrollIntoView({block: 'center'});
                button.click();
                return (button.innerText || '').trim();
              }
              for (const item of scrollables) {
                item.scrollTop = (item.scrollTop || 0) + 240;
                item.dispatchEvent(new WheelEvent('wheel', {deltaY: 240, bubbles: true}));
              }
            }
            return null;
            """
        )
        if clicked:
            self.log(f"Clique OK: relatorio {clicked}")
            return
        raise RuntimeError(f"Nao foi possivel selecionar Relatorio de Cash: {last_error}")

    def _click_visible_report_option(self) -> str | None:
        return self.driver.execute_script(
            """
            const menu = document.querySelector('div.dropdown-menu.show');
            if (!menu) return null;
            const normalize = (text) => (text || '')
              .normalize('NFD')
              .replace(/[\\u0300-\\u036f]/g, '')
              .trim()
              .toLowerCase();
            const spans = Array.from(menu.querySelectorAll('span'));
            const span = spans.find((item) => normalize(item.textContent) === 'relatorio de cash')
              || spans.find((item) => normalize(item.textContent).includes('cash'));
            if (!span) return null;
            const button = span.closest('button');
            if (!button) return null;
            button.scrollIntoView({block: 'center', inline: 'center'});
            for (const type of ['mouseover', 'mousemove', 'mousedown', 'mouseup', 'click']) {
              button.dispatchEvent(new MouseEvent(type, {
                bubbles: true,
                cancelable: true,
                view: window,
                buttons: type === 'mousedown' ? 1 : 0
              }));
            }
            return (span.textContent || button.innerText || '').trim();
            """
        )

    def _select_period(self, start: datetime, end: datetime) -> None:
        self.log(f"Selecionando periodo: {start:%d/%m/%Y} ate {end:%d/%m/%Y}")
        self._click("//input[@id='dateRange' or @name='dp']", "campo de periodo", 45)
        self._click_day(start, "data inicial")
        self._click_day(end, "data final")
        time.sleep(1)

    def _click_day(self, target: datetime, description: str) -> None:
        label = f"{target.day}/{target.month}/{target.year}"
        xpath = f"//ngb-datepicker//div[@role='gridcell' and @aria-label='{label}' and not(contains(@class, 'hidden'))]"
        for _ in range(24):
            visible = [element for element in self.driver.find_elements(By.XPATH, xpath) if element.is_displayed()]
            if visible:
                self._click_element(visible[0], f"{description} {label}")
                return
            self._move_calendar(target)
            time.sleep(0.35)
        raise RuntimeError(f"Nao foi possivel selecionar {description}: {label}")

    def _move_calendar(self, target: datetime) -> None:
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
        visible_months = []
        for element in self.driver.find_elements(By.XPATH, "//ngb-datepicker//div[contains(@class, 'ngb-dp-month-name')]"):
            parts = element.text.strip().lower().split()
            if len(parts) >= 2 and parts[0] in month_names and parts[-1].isdigit():
                visible_months.append(datetime(int(parts[-1]), month_names[parts[0]], 1))
        if not visible_months:
            raise RuntimeError("Nao foi possivel identificar o mes do calendario.")
        target_month = datetime(target.year, target.month, 1)
        if target_month < min(visible_months):
            self._click("//ngb-datepicker//button[@title='Previous month' or @aria-label='Previous month']", "mes anterior", 10)
        elif target_month > max(visible_months):
            self._click("//ngb-datepicker//button[@title='Next month' or @aria-label='Next month']", "proximo mes", 10)
        else:
            raise RuntimeError("Mes visivel, mas o dia nao ficou disponivel.")

    def _click_generate(self) -> None:
        self.log("Gerando relatorio...")
        candidates = [
            XPATH_GENERATE_REPORT,
            "//button[.//*[contains(normalize-space(.), 'Gerar relatorio')] or contains(normalize-space(.), 'Gerar relatorio')]",
            "//button[.//*[contains(normalize-space(.), 'Gerar relatório')] or contains(normalize-space(.), 'Gerar relatório')]",
            "//button[contains(normalize-space(.), 'Gerar')]",
        ]
        last_error = None
        for xpath in candidates:
            try:
                self._click(xpath, "botao Gerar relatorio", 10)
                return
            except Exception as exc:
                last_error = exc
        raise RuntimeError(f"Nao foi possivel clicar em Gerar relatorio: {last_error}")

    def _wait_csv(self, timeout: int = 180) -> Path:
        self.log("Aguardando download do CSV...")
        deadline = time.time() + timeout
        last_size = -1
        stable_since = None
        while time.time() < deadline:
            csv_files = sorted(self.download_dir.glob("*.csv"), key=lambda item: item.stat().st_mtime, reverse=True)
            partials = list(self.download_dir.glob("*.crdownload"))
            if csv_files:
                latest = csv_files[0]
                size = latest.stat().st_size
                if size == last_size and not partials:
                    stable_since = stable_since or time.time()
                    if time.time() - stable_since >= 2:
                        self.log(f"CSV baixado: {latest}")
                        return latest
                else:
                    stable_since = None
                    last_size = size
            time.sleep(1)
        raise TimeoutError("CSV de cash nao foi baixado dentro do tempo limite.")

    def _click(self, xpath: str, description: str, timeout: int = 30) -> None:
        last_error = None
        for _ in range(3):
            try:
                element = WebDriverWait(self.driver, timeout).until(EC.element_to_be_clickable((By.XPATH, xpath)))
                self._click_element(element, description)
                return
            except Exception as exc:
                last_error = exc
                time.sleep(1)
        raise RuntimeError(f"Nao foi possivel clicar em {description}: {last_error}")

    def _click_element(self, element, description: str) -> None:
        self.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", element)
        time.sleep(0.2)
        try:
            element.click()
        except Exception:
            self.driver.execute_script("arguments[0].click();", element)
        self.log(f"Clique OK: {description}")

    def _type(self, xpath: str, value: str, description: str) -> None:
        element = WebDriverWait(self.driver, 30).until(EC.element_to_be_clickable((By.XPATH, xpath)))
        element.click()
        element.send_keys(Keys.CONTROL, "a")
        element.send_keys(Keys.DELETE)
        element.send_keys(value)
        self.log(f"Preenchido: {description}")

    def _visible(self, xpath: str, description: str, timeout: int) -> None:
        WebDriverWait(self.driver, timeout).until(EC.visibility_of_element_located((By.XPATH, xpath)))
        self.log(f"Visivel: {description}")

    def _clean_downloads(self) -> None:
        for pattern in ("*.csv", "*.crdownload"):
            for file_path in self.download_dir.glob(pattern):
                try:
                    file_path.unlink()
                except Exception:
                    pass

    def _close(self) -> None:
        if self.driver is not None:
            self.driver.quit()


class RoboCashCoralPortalApp(ctk.CTk):
    def __init__(self) -> None:
        super().__init__()
        ctk.set_appearance_mode("light")
        ctk.set_default_color_theme("blue")

        config = load_config()
        yesterday = datetime.now() - timedelta(days=1)
        self.coral_user_var = ctk.StringVar(value=config.get("coral_user", "ddm"))
        self.coral_password_var = ctk.StringVar(value=config.get("coral_password", ""))
        self.portal_admin_var = ctk.StringVar(value=config.get("portal_admin", "admin"))
        self.portal_password_var = ctk.StringVar(value=config.get("portal_password", ""))
        self.start_date_var = ctk.StringVar(value=yesterday.strftime("%d/%m/%Y"))
        self.end_date_var = ctk.StringVar(value=yesterday.strftime("%d/%m/%Y"))
        self.visible_var = ctk.BooleanVar(value=bool(config.get("visible", True)))
        self.status_var = ctk.StringVar(value="Aguardando periodo")
        self.log_queue: queue.Queue[str] = queue.Queue()
        self.processing_thread: threading.Thread | None = None
        self.logo_image = None

        self.title(APP_TITLE)
        self.geometry(APP_GEOMETRY)
        self.minsize(700, 560)
        self.configure(fg_color=MAIN_BG)
        self._build_layout()
        self._poll_logs()

    def _build_layout(self) -> None:
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)
        container = ctk.CTkScrollableFrame(self, fg_color="transparent", scrollbar_button_color=BUTTON_BG, scrollbar_button_hover_color=BUTTON_ACTIVE_BG)
        container.grid(row=0, column=0, sticky="nsew", padx=16, pady=14)
        container.grid_columnconfigure(0, weight=1)

        self._header(container)
        self._credentials(container)
        self._period(container)
        self._execution(container)
        self._logs(container)

    def _header(self, parent) -> None:
        card = self._card(parent)
        card.grid(row=0, column=0, sticky="ew", pady=(0, 10))
        card.grid_columnconfigure(1, weight=1)
        logo_loaded = False
        for candidate in resolve_logo_candidates():
            try:
                if candidate.exists():
                    image = Image.open(candidate)
                    self.logo_image = ctk.CTkImage(light_image=image, dark_image=image, size=(96, 46))
                    ctk.CTkLabel(card, image=self.logo_image, text="").grid(row=0, column=0, rowspan=2, padx=(18, 16), pady=16)
                    logo_loaded = True
                    break
            except Exception:
                continue
        if not logo_loaded:
            ctk.CTkLabel(card, text="foco", text_color=PRIMARY_TEXT, font=("Segoe UI", 28, "bold")).grid(row=0, column=0, rowspan=2, padx=(18, 16), pady=16)
        ctk.CTkLabel(card, text="Atualizar Cash Coral", text_color=PRIMARY_TEXT, font=("Segoe UI", 24, "bold"), anchor="w").grid(
            row=0, column=1, sticky="ew", padx=(0, 18), pady=(16, 2)
        )
        ctk.CTkLabel(
            card,
            text="Baixa o Relatorio de Cash do Coral e grava as receitas no Portal Caixa.",
            text_color=MUTED_TEXT,
            font=("Segoe UI", 13),
            anchor="w",
        ).grid(row=1, column=1, sticky="ew", padx=(0, 18), pady=(0, 16))

    def _credentials(self, parent) -> None:
        card = self._section(parent, "Credenciais", 1)
        card.grid_columnconfigure((0, 1), weight=1)
        self._field(card, "Usuario Coral", self.coral_user_var).grid(row=1, column=0, sticky="ew", padx=(14, 8), pady=(8, 10))
        self._field(card, "Senha Coral", self.coral_password_var, show="*").grid(row=1, column=1, sticky="ew", padx=(8, 14), pady=(8, 10))
        self._field(card, "Usuario admin portal", self.portal_admin_var).grid(row=2, column=0, sticky="ew", padx=(14, 8), pady=(0, 14))
        self._field(card, "Senha admin portal", self.portal_password_var, show="*").grid(row=2, column=1, sticky="ew", padx=(8, 14), pady=(0, 14))

    def _period(self, parent) -> None:
        card = self._section(parent, "Periodo", 2)
        card.grid_columnconfigure((0, 1), weight=1)
        self._field(card, "Data inicial", self.start_date_var).grid(row=1, column=0, sticky="ew", padx=(14, 8), pady=(8, 10))
        self._field(card, "Data final", self.end_date_var).grid(row=1, column=1, sticky="ew", padx=(8, 14), pady=(8, 10))
        ctk.CTkCheckBox(
            card,
            text="Abrir Chrome visivel",
            variable=self.visible_var,
            text_color="#242424",
            fg_color=BUTTON_BG,
            hover_color=BUTTON_ACTIVE_BG,
            border_color=CARD_BORDER,
        ).grid(row=2, column=0, columnspan=2, sticky="w", padx=14, pady=(0, 14))

    def _execution(self, parent) -> None:
        card = self._section(parent, "Execucao", 3)
        card.grid_columnconfigure((0, 1, 2), weight=1)
        ctk.CTkLabel(card, textvariable=self.status_var, text_color=MUTED_TEXT, font=("Segoe UI", 13)).grid(
            row=1, column=0, columnspan=3, sticky="w", padx=14, pady=(8, 8)
        )
        self.progress = ctk.CTkProgressBar(card, height=14, progress_color=BUTTON_BG, fg_color=SOFT_RED, corner_radius=12)
        self.progress.grid(row=2, column=0, columnspan=3, sticky="ew", padx=14, pady=(0, 12))
        self.progress.set(0)
        self._primary_button(card, "Atualizar periodo", self.start_update).grid(row=3, column=0, sticky="ew", padx=(14, 8), pady=(0, 14))
        self._secondary_button(card, "Ontem", self.set_yesterday).grid(row=3, column=1, sticky="ew", padx=8, pady=(0, 14))
        self._secondary_button(card, "Fim de semana", self.set_weekend).grid(row=3, column=2, sticky="ew", padx=(8, 14), pady=(0, 14))

    def _logs(self, parent) -> None:
        card = self._section(parent, "Log", 4)
        card.grid_columnconfigure(0, weight=1)
        card.grid_rowconfigure(1, weight=1)
        self.log_box = ctk.CTkTextbox(card, fg_color="#fffafa", border_color=CARD_BORDER, border_width=1, text_color="#242424", font=("Consolas", 12))
        self.log_box.grid(row=1, column=0, sticky="nsew", padx=14, pady=(8, 14))

    def start_update(self) -> None:
        if self.processing_thread and self.processing_thread.is_alive():
            messagebox.showinfo("Atualizacao", "O robo ja esta executando.")
            return
        try:
            start = parse_ptbr_date(self.start_date_var.get())
            end = parse_ptbr_date(self.end_date_var.get())
            if end < start:
                raise ValueError("Data final menor que a inicial.")
            if not self.coral_user_var.get().strip() or not self.coral_password_var.get().strip():
                raise ValueError("Informe usuario e senha do Coral.")
            if not self.portal_admin_var.get().strip() or not self.portal_password_var.get().strip():
                raise ValueError("Informe usuario e senha admin do Portal Caixa.")
        except Exception as exc:
            messagebox.showwarning("Validacao", str(exc))
            return

        save_config(
            {
                "coral_user": self.coral_user_var.get().strip(),
                "coral_password": self.coral_password_var.get(),
                "portal_admin": self.portal_admin_var.get().strip(),
                "portal_password": self.portal_password_var.get(),
                "visible": self.visible_var.get(),
            }
        )
        self.progress.set(0.06)
        self.log_box.delete("1.0", "end")
        self.status_var.set("Iniciando atualizacao...")
        self.processing_thread = threading.Thread(target=self._run_update, args=(start, end), daemon=True)
        self.processing_thread.start()

    def _run_update(self, start: datetime, end: datetime) -> None:
        try:
            self._log(f"Periodo: {start:%d/%m/%Y} ate {end:%d/%m/%Y}")
            self.progress.set(0.18)
            downloader = CoralCashDownloader(
                self.coral_user_var.get().strip(),
                self.coral_password_var.get(),
                self.visible_var.get(),
                self._log,
            )
            csv_path = downloader.download(start, end)
            self.progress.set(0.62)
            entries = read_cash_csv(csv_path)
            self._log(f"Contratos validos no CSV: {len(entries)}")
            client = PortalCaixaClient(self.portal_admin_var.get().strip(), self.portal_password_var.get())
            client.login()
            applied = client.apply_entries(entries)
            self.progress.set(1)
            self.status_var.set("Atualizacao concluida.")
            self._log(f"Cash aplicado no Portal Caixa: {applied}")
            self.after(0, lambda: messagebox.showinfo("Concluido", f"Cash atualizado. Lancamentos aplicados: {applied}"))
        except Exception as exc:
            self.progress.set(0)
            self.status_var.set("Falha na atualizacao.")
            self._log(f"ERRO: {exc}")
            self.after(0, lambda: messagebox.showerror("Erro", str(exc)))

    def set_yesterday(self) -> None:
        yesterday = datetime.now() - timedelta(days=1)
        self.start_date_var.set(yesterday.strftime("%d/%m/%Y"))
        self.end_date_var.set(yesterday.strftime("%d/%m/%Y"))

    def set_weekend(self) -> None:
        today = datetime.now()
        self.start_date_var.set((today - timedelta(days=3)).strftime("%d/%m/%Y"))
        self.end_date_var.set((today - timedelta(days=1)).strftime("%d/%m/%Y"))

    def _log(self, message: str) -> None:
        self.log_queue.put(f"[{datetime.now():%H:%M:%S}] {message}")

    def _poll_logs(self) -> None:
        try:
            while True:
                message = self.log_queue.get_nowait()
                self.log_box.insert("end", f"{message}\n")
                self.log_box.see("end")
        except queue.Empty:
            pass
        self.after(200, self._poll_logs)

    def _card(self, parent):
        return ctk.CTkFrame(parent, fg_color=CARD_BG, border_color=CARD_BORDER, border_width=1, corner_radius=14)

    def _section(self, parent, title: str, row: int):
        card = self._card(parent)
        card.grid(row=row, column=0, sticky="nsew" if row == 4 else "ew", pady=(0, 10))
        ctk.CTkLabel(card, text=title, text_color=PRIMARY_TEXT, font=("Segoe UI", 16, "bold")).grid(
            row=0, column=0, columnspan=4, sticky="w", padx=14, pady=(12, 0)
        )
        return card

    def _field(self, parent, label: str, variable, show: str | None = None):
        frame = ctk.CTkFrame(parent, fg_color="transparent")
        frame.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(frame, text=label, text_color="#242424", font=("Segoe UI", 12, "bold")).grid(row=0, column=0, sticky="w", pady=(0, 5))
        ctk.CTkEntry(frame, textvariable=variable, show=show, height=40, fg_color="#ffffff", border_color=CARD_BORDER, border_width=1, corner_radius=10).grid(
            row=1, column=0, sticky="ew"
        )
        return frame

    def _primary_button(self, parent, text: str, command):
        return ctk.CTkButton(parent, text=text, command=command, height=40, fg_color=BUTTON_BG, hover_color=BUTTON_ACTIVE_BG, corner_radius=12, font=("Segoe UI", 13, "bold"))

    def _secondary_button(self, parent, text: str, command):
        return ctk.CTkButton(
            parent,
            text=text,
            command=command,
            height=40,
            fg_color="#ffffff",
            hover_color=SOFT_RED,
            border_color=CARD_BORDER,
            border_width=1,
            text_color=PRIMARY_TEXT,
            corner_radius=12,
            font=("Segoe UI", 13, "bold"),
        )


def main() -> None:
    app = RoboCashCoralPortalApp()
    app.mainloop()


if __name__ == "__main__":
    main()
