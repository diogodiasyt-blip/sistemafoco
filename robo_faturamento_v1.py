from __future__ import annotations

# Unified single-file version generated from the working modules.
# The original modules are preserved in the workspace as backup.

# ===== config.py =====
import os
import shutil
from pathlib import Path


APP_TITLE = "Robô de Faturamento de Avarias"
APP_GEOMETRY = "1180x760"
SAVE_DATE_FORMAT = "%d/%m/%Y %H:%M:%S"

DEFAULT_PROTHEUS_URL = "https://focoaluguel162907.protheus.cloudtotvs.com.br:1453/webapp/"
EXCEL_REQUIRED_COLUMNS = [
    "CLIENTE",
    "MODALIDADE",
    "CONTRATO",
    "CPF",
    "LOJA",
    "VALOR",
    "MOTIVO",
    "HISTORICO",
    "FATURA",
    "FATURAMENTO_STATUS",
]

EXCEL_OUTPUT_COLUMNS = [
    "FATURA",
    "STATUS_PROCESSAMENTO",
    "DETALHE_PROCESSAMENTO",
    "DATA_PROCESSAMENTO",
]

EXCEL_COLUMN_ALIASES = {
    "CLIENTE": {"CLIENTE", "NOME CLIENTE", "CLIENTE NOME", "NOM_CLIENTE"},
    "MODALIDADE": {"MODALIDADE", "SEGREGACAO", "SEGREGAÇÃO", "MODALIDADE FATURAMENTO"},
    "CONTRATO": {"CONTRATO", "N CONTRATO", "NR CONTRATO", "NUMERO CONTRATO", "Nº CONTRATO"},
    "CPF": {"CPF", "CPF/CNPJ", "DOCUMENTO"},
    "LOJA": {"LOJA", "CENTRO DE CUSTO", "C DE CUSTO", "C. DE CUSTO", "CCUSTO"},
    "VALOR": {"VALOR", "VALOR AVARIA", "VLR", "VLR AVARIA"},
    "MOTIVO": {"MOTIVO", "MOTIVO AVARIA"},
    "HISTORICO": {"HISTORICO", "HISTÓRICO", "OBSERVACAO", "OBS"},
    "FATURA": {"FATURA", "NUMERO FATURA", "N FATURA", "Nº FATURA", "TOTVS"},
    "FATURAMENTO_STATUS": {"FATURAMENTO", "STATUS FATURAMENTO", "APROVACAO FATURAMENTO", "STATUS"},
}

FIELD_MAP = {
    "prefixo": {
        "protheus_name": "E1_PREFIXO",
        "tag": "wa-text-input",
        "kind": "text",
        "label": "Prefixo",
    },
    "numero_titulo": {
        "protheus_name": "E1_NUM",
        "tag": "wa-text-input",
        "kind": "read_only_capture",
        "label": "No. Titulo",
    },
    "tipo_pg": {
        "protheus_name": "E1_XTIPOPG",
        "tag": "wa-combobox",
        "kind": "combo",
        "label": "Tipo Pg",
    },
    "tipo": {
        "protheus_name": "E1_TIPO",
        "tag": "wa-text-input",
        "kind": "lookup",
        "label": "Tipo",
    },
    "natureza": {
        "protheus_name": "E1_NATUREZ",
        "tag": "wa-text-input",
        "kind": "lookup",
        "label": "Natureza",
    },
    "cliente": {
        "protheus_name": "E1_CLIENTE",
        "tag": "wa-text-input",
        "kind": "lookup",
        "label": "Cliente",
    },
    "contrato": {
        "protheus_name": "E1_XCONTRA",
        "tag": "wa-text-input",
        "kind": "text",
        "label": "Nº Contrato",
    },
    "vencimento": {
        "protheus_name": "E1_VENCTO",
        "tag": "wa-text-input",
        "kind": "date",
        "label": "Vencimento",
    },
    "valor_titulo": {
        "protheus_name": "E1_VALOR",
        "tag": "wa-text-input",
        "kind": "number",
        "label": "Vlr. Titulo",
    },
    "centro_custo": {
        "protheus_name": "E1_CCUSTO",
        "tag": "wa-text-input",
        "kind": "lookup",
        "label": "C. de Custo",
    },
    "negocio": {
        "protheus_name": "E1_NEGOCIO",
        "tag": "wa-combobox",
        "kind": "combo",
        "label": "Negocio",
    },
    "segregacao": {
        "protheus_name": "E1_XSEGREG",
        "tag": "wa-text-input",
        "kind": "lookup",
        "label": "Segregação",
    },
    "motivo": {
        "protheus_name": "E1_MOTIVO2",
        "tag": "wa-combobox",
        "kind": "combo",
        "label": "Motivo",
    },
    "historico": {
        "protheus_name": "E1_HIST",
        "tag": "wa-text-input",
        "kind": "text",
        "label": "Historico",
    },
}

LOOKUP_SELECTORS = {
    "dialog_title": 'wa-text-view[caption*="Consulta Padrão - Cliente"]',
    "dialog_root": "wa-dialog.dict-msdialog",
    "search_input": 'wa-text-input[placeholder="Pesquisar"]',
    "search_input_fallback": "wa-text-input.dict-tget",
    "search_button": "wa-button.dict-tbutton",
    "search_button_xpath": "/html/body/wa-dialog/wa-panel/wa-panel[2]/wa-tab-view/wa-tab-page/wa-dialog[3]/wa-panel[2]/wa-panel[2]/wa-panel[4]/wa-panel[3]/wa-button[2]",
    "grid": "wa-tgrid",
    "first_result_name": 'td[id="2"] label',
    "ok_button": 'wa-button[name="M->E1_CLIENTE"][caption="OK"]',
    "cancel_button": 'wa-button[caption="Cancelar"]',
}

PROTHEUS_NAVIGATION = {
    "webagent_dialog_selector": 'wa-dialog h1',
    "webagent_checkbox_selector": '#do-not-show-again',
    "start_parameters_dialog_selector": "wa-dialog.startParameters",
    "start_parameters_program_selector": "#selectStartProg",
    "start_parameters_environment_selector": "#selectEnv",
    "start_parameters_ok_xpath": "//wa-dialog[contains(@class,'startParameters')]//button[.//span[normalize-space()='Ok' or normalize-space()='OK' or normalize-space()='ok'] or normalize-space()='Ok' or normalize-space()='OK' or normalize-space()='ok']",
    "initial_ok_xpath": "/html/body/wa-dialog[2]//footer/wa-button[2]//button",
    "initial_ok_fallback_xpath": "//button[.//span[normalize-space()='Ok' or normalize-space()='OK' or normalize-space()='ok']]",
    "login_username_xpath": "/html/body/ld-root/ng-component/pro-login/po-page-login/po-page-background/div/div/div[2]/div/form/div/div[1]/div[1]/po-login/po-field-container/div/div[2]/input",
    "login_password_xpath": "/html/body/ld-root/ng-component/pro-login/po-page-login/po-page-background/div/div/div[2]/div/form/div/div[2]/div[1]/po-password/po-field-container/div/div[2]/input",
    "login_submit_xpath": "/html/body/ld-root/ng-component/pro-login/po-page-login/po-page-background/div/div/div[2]/div/form/div/po-button/button",
    "environment_input_xpath": "/html/body/ld-root/ng-component/pro-session-settings/pro-page-background/div/div/div[1]/div/form/div/pro-system-module-lookup/div/po-lookup/po-field-container/div/div[2]/div/input",
    "environment_confirm_xpath": "/html/body/ld-root/ng-component/pro-session-settings/pro-page-background/div/div/div[1]/div/form/div/div[4]/po-button[2]/button",
    "menu_updates_selector": 'span.caption[title*="Atualizacoes"], span.caption[title*="Atualizações"]',
    "menu_contas_receber_selector": 'wa-menu-item[caption*="Contas a Receber"]',
    "menu_funcoes_receber_selector": 'wa-menu-item[caption*="Funcoes Contas a Receber"], wa-menu-item[caption*="Funções Contas a Receber"]',
    "menu_funcoes_confirm_xpath": "/html/body/wa-dialog[2]/wa-panel[3]/wa-panel/wa-panel/wa-button[1]",
    "browse_button_selector": 'wa-button[caption*="Ctas a Receber"]',
    "browse_incluir_selector": 'wa-text-view[caption*="Incluir"]',
    "browse_incluir_confirm_xpath": "/html/body/wa-dialog/wa-panel/wa-panel[2]/wa-tab-view/wa-tab-page/wa-dialog[2]/wa-panel[1]/wa-button[1]",
    "faturamento_title_selector": 'wa-text-view[caption="Contas a Receber"]',
    "faturamento_title_caption": "Contas a Receber",
    "environment_expected_value": "6",
}


# ===== excel_handler.py =====
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Tuple

import pandas as pd



@dataclass
class WorkbookContext:
    source_path: Path
    dataframe: pd.DataFrame
    normalized_columns: Dict[str, str]


def _normalize_header(header: str) -> str:
    return (
        str(header)
        .strip()
        .upper()
        .replace("Á", "A")
        .replace("À", "A")
        .replace("Ã", "A")
        .replace("Â", "A")
        .replace("É", "E")
        .replace("Ê", "E")
        .replace("Í", "I")
        .replace("Ó", "O")
        .replace("Ô", "O")
        .replace("Õ", "O")
        .replace("Ú", "U")
        .replace("Ç", "C")
    )


def load_workbook(path: str | Path) -> WorkbookContext:
    source_path = Path(path)
    dataframe = pd.read_excel(source_path)
    normalized_columns = {_normalize_header(column): column for column in dataframe.columns}

    rename_map: Dict[str, str] = {}
    for target, aliases in EXCEL_COLUMN_ALIASES.items():
        for alias in aliases:
            alias_normalized = _normalize_header(alias)
            if alias_normalized in normalized_columns:
                rename_map[normalized_columns[alias_normalized]] = target
                break

    dataframe = dataframe.rename(columns=rename_map)

    for output_column in EXCEL_OUTPUT_COLUMNS:
        if output_column not in dataframe.columns:
            dataframe[output_column] = ""

    return WorkbookContext(
        source_path=source_path,
        dataframe=dataframe,
        normalized_columns=rename_map,
    )


def validate_workbook(context: WorkbookContext) -> List[str]:
    missing = [column for column in EXCEL_REQUIRED_COLUMNS if column not in context.dataframe.columns]
    return missing


def _normalize_cpf_for_lookup(raw_value: str) -> str | None:
    text = str(raw_value).strip()
    if not text:
        return None
    if any(char.isalpha() for char in text):
        return None

    digits = "".join(char for char in text if char.isdigit())
    if not digits or len(digits) > 11:
        return None
    return digits.zfill(11)


def prepare_rows(context: WorkbookContext) -> Tuple[pd.DataFrame, pd.DataFrame]:
    dataframe = context.dataframe.copy()

    for column in [
        "CLIENTE",
        "MODALIDADE",
        "CONTRATO",
        "CPF",
        "LOJA",
        "MOTIVO",
        "HISTORICO",
        "FATURA",
        "FATURAMENTO_STATUS",
    ]:
        if column in dataframe.columns:
            dataframe[column] = dataframe[column].fillna("").astype(str).str.strip()

    dataframe["CPF_NORMALIZADO"] = dataframe["CPF"].apply(_normalize_cpf_for_lookup)
    dataframe["CPF"] = dataframe["CPF_NORMALIZADO"].fillna(dataframe["CPF"])
    dataframe["VALOR"] = pd.to_numeric(dataframe.get("VALOR"), errors="coerce")
    faturamento_status = dataframe["FATURAMENTO_STATUS"].str.upper()

    ignored_mask = (
        dataframe["FATURA"].astype(str).str.strip().ne("")
        | dataframe["CLIENTE"].eq("")
        | dataframe["MODALIDADE"].eq("")
        | dataframe["CONTRATO"].eq("")
        | dataframe["LOJA"].eq("")
        | dataframe["CPF_NORMALIZADO"].isna()
        | dataframe["MOTIVO"].eq("")
        | dataframe["HISTORICO"].eq("")
        | faturamento_status.ne("APROVADO")
        | dataframe["VALOR"].isna()
        | dataframe["VALOR"].le(0)
    )

    apt = dataframe.loc[~ignored_mask].copy()
    ignored = dataframe.loc[ignored_mask].copy()
    return apt, ignored


def mark_row(
    context: WorkbookContext,
    row_index: int,
    status: str,
    detail: str,
    invoice: str = "",
) -> None:
    context.dataframe.at[row_index, "STATUS_PROCESSAMENTO"] = status
    context.dataframe.at[row_index, "DETALHE_PROCESSAMENTO"] = detail
    context.dataframe.at[row_index, "DATA_PROCESSAMENTO"] = datetime.now().strftime(SAVE_DATE_FORMAT)
    if invoice:
        context.dataframe.at[row_index, "FATURA"] = invoice


def save_workbook(context: WorkbookContext, output_path: str | Path) -> Path:
    output = Path(output_path)
    output.parent.mkdir(parents=True, exist_ok=True)
    context.dataframe.to_excel(output, index=False)
    return output


# ===== report_handler.py =====
from datetime import datetime
from pathlib import Path

import pandas as pd


REPORT_COLUMNS = ["CONTRATO", "VALOR", "NUMERO_TITULO"]


def build_daily_report_path(base_dir: str | Path) -> Path:
    base_dir = Path(base_dir)
    date_label = datetime.now().strftime("%Y-%m-%d")
    return base_dir / f"relatorio_faturamento_{date_label}.xlsx"


def append_report_entry(base_dir: str | Path, contract: str, value, invoice_number: str) -> Path:
    report_path = build_daily_report_path(base_dir)

    if report_path.exists():
        dataframe = pd.read_excel(report_path)
    else:
        dataframe = pd.DataFrame(columns=REPORT_COLUMNS)

    if "CONTRATO" not in dataframe.columns:
        dataframe = pd.DataFrame(columns=REPORT_COLUMNS)

    contract = str(contract).strip()
    invoice_number = str(invoice_number).strip()

    try:
        numeric_value = float(value)
    except Exception:
        text_value = str(value).strip().replace(".", "").replace(",", ".")
        numeric_value = float(text_value)

    new_row = {
        "CONTRATO": contract,
        "VALOR": numeric_value,
        "NUMERO_TITULO": invoice_number,
    }

    existing_mask = dataframe["CONTRATO"].fillna("").astype(str).str.strip().eq(contract)
    if existing_mask.any():
        dataframe.loc[existing_mask, REPORT_COLUMNS] = [new_row["CONTRATO"], new_row["VALOR"], new_row["NUMERO_TITULO"]]
    else:
        dataframe = pd.concat([dataframe, pd.DataFrame([new_row])], ignore_index=True)

    report_path.parent.mkdir(parents=True, exist_ok=True)
    dataframe.to_excel(report_path, index=False)
    return report_path


# ===== protheus_bot.py =====
from datetime import datetime
import re
from typing import Callable


try:
    from playwright.sync_api import Browser, BrowserContext, Frame, Page, Playwright, sync_playwright
except Exception:  # pragma: no cover
    Browser = BrowserContext = Frame = Page = Playwright = object
    sync_playwright = None


LogFn = Callable[[str], None]


SPECIAL_NATURE_MOTIVES = {
    "AVARIA/SINISTRO",
    "APREENSAO AUTORIDADES",
    "APREENSÃO AUTORIDADES",
    "ASSISTENCIA 24 H",
    "ASSISTÊNCIA 24 H",
}

MOTIVE_CODE_GROUPS = {
    "1": {"AVARIA/SINISTRO"},
    "2": {"DIARIAS", "DIARIAS E ITENS EXTRAS", "ITENS EXTRAS", "EXTRACONTRATUAL"},
    "3": {"BUSCA/ RECUPERACAO", "BUSCA/ RECUPERAÇÃO", "APROPRIACAO", "APROPRIAÇÃO", "APREENSAO AUTORIDADES", "APREENSÃO AUTORIDADES"},
    "4": {"ASSISTENCIA 24 H", "ASSISTÊNCIA 24 H"},
}


class ProtheusBot:
    def __init__(
        self,
        username: str,
        password: str,
        log: LogFn,
        base_url: str = DEFAULT_PROTHEUS_URL,
        headless: bool = False,
    ) -> None:
        self.username = username
        self.password = password
        self.log = log
        self.base_url = base_url
        self.headless = headless

        self.playwright: Playwright | None = None
        self.browser: Browser | None = None
        self.context: BrowserContext | None = None
        self.page: Page | None = None

    def _resolve_browser_executable(self) -> str | None:
        env_candidates = [
            os.environ.get("FOCO_BROWSER_EXECUTABLE", "").strip(),
            os.environ.get("PLAYWRIGHT_CHROME_EXECUTABLE", "").strip(),
        ]
        for candidate in env_candidates:
            if candidate and os.path.exists(candidate):
                return candidate

        common_paths = [
            os.path.join(os.environ.get("PROGRAMFILES", r"C:\Program Files"), "Google", "Chrome", "Application", "chrome.exe"),
            os.path.join(os.environ.get("PROGRAMFILES(X86)", r"C:\Program Files (x86)"), "Google", "Chrome", "Application", "chrome.exe"),
            os.path.join(os.environ.get("LOCALAPPDATA", ""), "Google", "Chrome", "Application", "chrome.exe"),
            os.path.join(os.environ.get("PROGRAMFILES", r"C:\Program Files"), "Microsoft", "Edge", "Application", "msedge.exe"),
            os.path.join(os.environ.get("PROGRAMFILES(X86)", r"C:\Program Files (x86)"), "Microsoft", "Edge", "Application", "msedge.exe"),
            os.path.join(os.environ.get("LOCALAPPDATA", ""), "Microsoft", "Edge", "Application", "msedge.exe"),
        ]
        for candidate in common_paths:
            if candidate and os.path.exists(candidate):
                return candidate

        for command in ("chrome.exe", "chrome", "msedge.exe", "msedge"):
            resolved = shutil.which(command)
            if resolved:
                return resolved

        return None

    def start(self) -> None:
        if self.page:
            return
        if sync_playwright is None:
            raise RuntimeError("Playwright nao esta disponivel neste ambiente.")

        self.log("Inicializando automacao do Protheus.")
        self.playwright = sync_playwright().start()
        launch_options = {
            "headless": self.headless,
            "args": ["--start-maximized"],
        }
        launch_errors: list[str] = []

        browser_executable = self._resolve_browser_executable()
        if browser_executable:
            try:
                self.log(f"Usando navegador local: {browser_executable}")
                self.browser = self.playwright.chromium.launch(
                    executable_path=browser_executable,
                    **launch_options,
                )
            except Exception as exc:
                launch_errors.append(f"executavel local: {exc}")
                self.log(f"Falha ao iniciar navegador local. Tentando fallback Playwright: {exc}")

        if self.browser is None:
            try:
                self.log("Tentando navegador padrao do Playwright.")
                self.browser = self.playwright.chromium.launch(**launch_options)
            except Exception as exc:
                launch_errors.append(f"playwright chromium: {exc}")

        if self.browser is None:
            if self.playwright:
                try:
                    self.playwright.stop()
                except Exception:
                    pass
                self.playwright = None
            details = " | ".join(launch_errors) if launch_errors else "nenhum navegador disponivel"
            raise RuntimeError(f"Nao foi possivel iniciar o navegador do faturamento: {details}")

        self.context = self.browser.new_context(viewport=None)
        self.page = self.context.new_page()

    def stop(self) -> None:
        if self.context:
            self.context.close()
        if self.browser:
            self.browser.close()
        if self.playwright:
            self.playwright.stop()

        self.playwright = None
        self.browser = None
        self.context = None
        self.page = None

    def login(self) -> None:
        page = self._require_page()
        self.log("Abrindo Protheus.")
        page.goto(self.base_url, wait_until="load")
        page.wait_for_timeout(700)

        self._dismiss_webagent_if_present()
        self._close_popup_ok_if_present()
        self._perform_login()
        self._confirm_second_enter_if_present()
        self._wait_main_menu_ready()
        self.log("Login concluido.")

    def navigate_to_faturamento(self) -> None:
        page = self._require_page()
        self.log("Abrindo tela de faturamento.")

        self._navigate_menu(page)
        self._click_ctas_receber_button(page)
        self._click_incluir_menu(page)
        self._finalize_incluir(page)
        self.wait_until_faturamento_screen()
        self.log("Tela de faturamento carregada com sucesso.")

    def wait_until_faturamento_screen(self) -> None:
        page = self._require_page()
        dialog = page.locator('wa-dialog.dict-msdialog[title="Contas a Receber"][opened]').last
        dialog.wait_for(state="visible", timeout=30000)

        self._wait_any_visible(
            [
                dialog.locator('wa-text-view[caption="Contas a Receber"]'),
                dialog.locator('wa-tab-button#BUTTON-COMP6003[active]'),
                dialog.locator('wa-tab-button#BUTTON-COMP6003'),
            ],
            timeout=30000,
        )
        self._wait_any_visible(
            [
                dialog.locator('wa-button[caption="Salvar"]'),
                dialog.locator('wa-button#COMP6156'),
            ],
            timeout=30000,
        )
        self._wait_any_visible(
            [
                dialog.locator('wa-button[caption="Cancelar"]'),
                dialog.locator('wa-button#COMP6157'),
            ],
            timeout=30000,
        )

        dialog.locator(self.build_field_selector("prefixo")).last.wait_for(state="visible", timeout=30000)
        dialog.locator(self.build_field_selector("numero_titulo")).last.wait_for(state="visible", timeout=30000)
        page.wait_for_timeout(400)

    def process_row(self, row: dict) -> dict:
        contract = str(row.get("CONTRATO", "")).strip()
        cpf = self._normalize_cpf_for_lookup(str(row.get("CPF", "")).strip())
        self.log(f"Iniciando contrato {contract}.")
        self.wait_until_faturamento_screen()

        if not cpf:
            return {
                "status": "IGNORADO_CPF_INVALIDO",
                "detail": "CPF invalido para pesquisa do cliente.",
                "invoice": "",
            }

        expected_client_name = str(row.get("CLIENTE", "")).strip()
        if not self.lookup_cliente_by_cpf(cpf, expected_client_name):
            return {
                "status": "IGNORADO_CLIENTE_NAO_ENCONTRADO",
                "detail": f"Nenhum cliente valido encontrado para o CPF {cpf}.",
                "invoice": "",
            }

        invoice_number = self.fill_billing_fields(row)

        return {
            "status": "SALVO",
            "detail": "Cliente localizado, campos preenchidos e faturamento salvo.",
            "invoice": invoice_number,
        }

    def lookup_cliente_by_cpf(self, cpf: str, expected_client_name: str) -> bool:
        page = self._require_page()
        self.log(f"Pesquisando cliente pelo CPF {cpf}.")

        self._click_first_visible_locator(
            [
                page.get_by_title("Codigo do Cliente        ").get_by_role("button"),
                page.locator(self.build_lookup_button_selector("cliente")),
            ],
            timeout=45000,
            pause_ms=500,
        )

        dialog = page.locator(LOOKUP_SELECTORS["dialog_root"]).filter(has=page.locator('wa-text-view[caption*="Cliente"]')).last
        dialog.wait_for(state="visible", timeout=45000)

        search_input = self._first_visible_locator(
            [
                dialog.get_by_role("textbox", name="Pesquisar"),
                dialog.get_by_role("textbox", name="CNPJ/CPF"),
                dialog.locator(LOOKUP_SELECTORS["search_input"]).last,
                dialog.locator(LOOKUP_SELECTORS["search_input_fallback"]).last,
            ]
        )
        if search_input is None:
            raise RuntimeError("Campo de pesquisa do cliente nao foi localizado.")

        self._type_lookup_search_value(search_input, cpf)

        self._click_first_visible_locator(
            [
                dialog.locator("#COMP7534 > button"),
                dialog.locator("#COMP7534"),
            ],
            timeout=45000,
            pause_ms=500,
        )

        dialog.locator(LOOKUP_SELECTORS["grid"]).wait_for(state="visible", timeout=45000)
        page.wait_for_timeout(1800)

        first_result = self._first_visible_locator(
            [
                dialog.locator(LOOKUP_SELECTORS["first_result_name"]).first,
                dialog.locator('td[id="2"] label').first,
                dialog.locator("tbody tr td:nth-child(2) label").first,
            ]
        )
        if first_result is None:
            try:
                cancel_button = self._first_visible_locator(
                    [
                        dialog.get_by_role("button", name="Cancelar"),
                        dialog.locator(LOOKUP_SELECTORS["cancel_button"]),
                    ]
                )
                if cancel_button is not None:
                    cancel_button.click(force=True)
            except Exception:
                pass
            return False

        found_name = self._read_locator_text(first_result)
        self.log(f"Primeiro nome retornado na busca: {found_name or '(vazio)'}")

        expected_name_normalized = self._normalize_text(expected_client_name)
        found_name_normalized = self._normalize_text(found_name)
        if not self._names_match(expected_name_normalized, found_name_normalized):
            self.log(
                f"Cliente retornado nao confere. Esperado: {expected_client_name} | Encontrado: {found_name or '(vazio)'}"
            )
            try:
                cancel_button = self._first_visible_locator(
                    [
                        dialog.get_by_role("button", name="Cancelar"),
                        dialog.locator(LOOKUP_SELECTORS["cancel_button"]),
                    ]
                )
                if cancel_button is not None:
                    cancel_button.click(force=True)
                    page.wait_for_timeout(700)
            except Exception:
                pass
            return False

        first_result.scroll_into_view_if_needed(timeout=3000)
        first_result.click(force=True)
        page.wait_for_timeout(700)

        ok_button = self._first_visible_locator(
            [
                dialog.get_by_role("button", name="OK"),
                dialog.get_by_text("OK", exact=True),
                dialog.locator('wa-button[caption="OK"]'),
                dialog.locator(LOOKUP_SELECTORS["ok_button"]),
            ]
        )
        if ok_button is None:
            raise RuntimeError("Botao OK da consulta de cliente nao foi localizado.")

        ok_button.click(force=True)
        page.wait_for_timeout(800)
        return True

    def _type_lookup_search_value(self, locator, value: str) -> None:
        if locator is None:
            raise RuntimeError("Campo de pesquisa da consulta nao foi localizado.")

        page = self._require_page()
        locator.wait_for(state="visible", timeout=45000)
        locator.scroll_into_view_if_needed(timeout=3000)

        # O campo CNPJ/CPF da consulta do TOTVS responde melhor a eventos reais
        # de teclado do que a preenchimento direto por DOM.
        locator.click(force=True)
        page.wait_for_timeout(250)
        for _ in range(16):
            page.keyboard.press("Backspace")
        page.wait_for_timeout(300)
        page.keyboard.type(str(value), delay=120)
        page.wait_for_timeout(350)
        page.keyboard.press("Tab")
        page.wait_for_timeout(500)

    def fill_billing_fields(self, row: dict) -> str:
        page = self._require_page()
        dialog = self._faturamento_dialog()
        payload = self._build_billing_payload(row)

        prefixo_input = self._first_visible_locator(
            [dialog.get_by_title("Prefixo do titulo        ").get_by_role("textbox")]
        )
        numero_titulo_input = self._first_visible_locator(
            [dialog.get_by_title("Numero do Titulo         ").get_by_role("textbox")]
        )
        tipo_pg_combo = self._first_visible_locator(
            [dialog.get_by_title("Tipo de pagamento        ").get_by_role("combobox")]
        )
        tipo_input = self._first_visible_locator(
            [dialog.get_by_title("Tipo do titulo           ").get_by_role("textbox")]
        )
        natureza_input = self._first_visible_locator([dialog.locator("#COMP6023 > input")])
        contrato_input = self._first_visible_locator([dialog.locator("#COMP6028 > input")])
        vencimento_input = self._first_visible_locator([dialog.locator("#COMP6030 > input")])
        valor_input = self._first_visible_locator([dialog.locator("#COMP6032 > input")])
        loja_input = self._first_visible_locator([dialog.locator("#COMP6036 > input")])
        negocio_combo = self._first_visible_locator(
            [dialog.get_by_title("Negocio                  ").get_by_role("combobox")]
        )
        segregacao_input = self._first_visible_locator([dialog.locator("#COMP6038 > input")])
        motivo_combo = self._first_visible_locator(
            [dialog.get_by_title("Motivo                   ").get_by_role("combobox")]
        )
        historico_input = self._first_visible_locator([dialog.locator("#COMP6040 > input")])

        self._click_and_fill_field(prefixo_input, payload["prefixo"])
        self._settle_after_field()

        invoice_number = self._read_from_locator(numero_titulo_input)

        self._select_combo_option(tipo_pg_combo, payload["tipo_pg"])
        self._settle_after_field()

        self._click_and_fill_field(tipo_input, payload["tipo"])
        self._settle_after_field()
        self._click_and_fill_field(natureza_input, payload["natureza"])
        self._settle_after_field()
        self._click_and_fill_field(contrato_input, payload["contrato"])
        self._settle_after_field()
        self._click_and_fill_field(vencimento_input, payload["vencimento"])
        self._settle_after_field()
        self._click_and_fill_field(valor_input, payload["valor_titulo"])
        self._settle_after_field()
        self._click_and_fill_field(loja_input, payload["centro_custo"])
        self._settle_after_field()

        self._select_combo_option(negocio_combo, payload["negocio"])
        self._settle_after_field()

        self._click_and_fill_field(segregacao_input, payload["segregacao"])
        self._settle_after_field()

        self._select_combo_option(motivo_combo, payload["motivo"])
        self._settle_after_field()

        self._click_and_fill_field(historico_input, payload["historico"])
        self._settle_after_field()

        if not invoice_number:
            page.wait_for_timeout(600)
            invoice_number = self._read_from_locator(numero_titulo_input)

        self._complete_save_flow(dialog)
        if invoice_number:
            self.log(f"Numero do titulo capturado: {invoice_number}")
        return invoice_number

    def _dismiss_webagent_if_present(self) -> None:
        page = self._require_page()
        try:
            dialog = page.locator("wa-dialog").filter(has=page.locator("h1"))
            if dialog.count() == 0 or not dialog.first.locator("h1").first.is_visible():
                return

            title = (dialog.first.locator("h1").first.inner_text() or "").strip()
            if "TOTVS WebAgent" not in title:
                return

            checkbox = page.locator("#do-not-show-again")
            if checkbox.count() > 0:
                checkbox.click(force=True)

            page.evaluate(
                """
                () => {
                    const dialogs = Array.from(document.querySelectorAll('wa-dialog'));
                    for (const dialog of dialogs) {
                        const title = dialog.querySelector('h1');
                        if (title && title.textContent.includes('TOTVS WebAgent')) {
                            dialog.remove();
                        }
                    }
                }
                """
            )
        except Exception:
            pass

    def _close_popup_ok_if_present(self) -> None:
        page = self._require_page()
        try:
            page.wait_for_selector("text=OK", timeout=10000)
            page.click("text=OK", force=True)
        except Exception:
            pass

    def _perform_login(self) -> None:
        page = self._require_page()
        login_frame = self._find_login_frame(page)

        username_input = login_frame.locator('input[name="login"]')
        username_input.wait_for(state="visible", timeout=30000)
        username_input.click(force=True)
        username_input.fill("")
        username_input.type(self.username, delay=120)

        page.wait_for_timeout(400)

        password_input = login_frame.locator('input[name="password"]')
        password_input.wait_for(state="visible", timeout=30000)
        password_input.click(force=True)
        password_input.fill("")
        password_input.type(self.password, delay=120)

        page.wait_for_timeout(1200)
        self._click_first_visible_locator(
            [
                login_frame.locator("button:has-text('Entrar')"),
                login_frame.get_by_role("button", name="Entrar"),
            ],
            timeout=45000,
            pause_ms=500,
        )

    def _confirm_second_enter_if_present(self) -> None:
        page = self._require_page()
        for _ in range(30):
            page.wait_for_timeout(800)
            for frame in page.frames:
                try:
                    self._click_first_visible_locator(
                        [
                            frame.locator("button:has-text('Entrar')"),
                            frame.get_by_role("button", name="Entrar"),
                        ],
                        timeout=3000,
                        pause_ms=400,
                    )
                    if self._wait_main_menu_ready_if_possible():
                        return
                except Exception:
                    pass
            page.wait_for_timeout(800)

    def _wait_main_menu_ready(self) -> None:
        page = self._require_page()
        candidates = [
            'wa-menu-item:has-text("Atualizacoes")',
            'wa-menu-item:has-text("Atualizações")',
            'text="Atualizacoes"',
            'text="Atualizações"',
            'wa-menu-item:has-text("Contas a Receber")',
            'text="Contas a Receber"',
        ]
        for selector in candidates:
            try:
                page.locator(selector).first.wait_for(state="visible", timeout=15000)
                page.wait_for_timeout(250)
                return
            except Exception:
                continue
        raise RuntimeError("O menu principal do Protheus nao carregou apos o login.")

    def _wait_main_menu_ready_if_possible(self) -> bool:
        try:
            self._wait_main_menu_ready()
            return True
        except Exception:
            return False

    def _navigate_menu(self, page: Page) -> None:
        possible_steps = [
            [
                page.get_by_text("Atualizações (17)", exact=True),
                page.get_by_text("Contas a Receber (27)", exact=True),
                page.get_by_title("Funções Contas a Receber"),
            ],
            [
                page.get_by_text("Atualizacoes (17)", exact=True),
                page.get_by_text("Contas a Receber (27)", exact=True),
                page.get_by_title("Funcoes Contas a Receber"),
            ],
        ]

        for steps in possible_steps:
            try:
                for level, locator in enumerate(steps, 1):
                    self._click_first_visible_locator([locator], timeout=45000, pause_ms=500)
                    page.wait_for_timeout(1200 + level * 350)
                return
            except Exception:
                page.wait_for_timeout(1200)

        raise RuntimeError("Nenhum caminho de navegacao por texto funcionou no menu do Protheus.")

    def _click_ctas_receber_button(self, page: Page) -> None:
        candidates = [
            page.get_by_role("button", name="Ctas a Receber"),
            page.locator("#COMP4599"),
            page.locator('wa-button[caption*="Ctas a Receber"]'),
        ]
        self._click_first_visible_locator(candidates, timeout=45000, pause_ms=500)
        page.wait_for_timeout(1200)

    def _click_confirmar_menu(self, page: Page) -> None:
        browse_button = page.locator("#COMP4599")
        if self._is_any_locator_visible([browse_button]):
            return

        candidates = [
            page.locator("#COMP6057"),
            page.get_by_text("Confirmar", exact=True),
            page.locator('wa-button:has-text("Confirmar")'),
            page.locator('button:has-text("Confirmar")'),
            page.locator("xpath=/html/body/wa-dialog[2]/wa-panel[3]/wa-panel/wa-panel/wa-button[1]"),
        ]

        end_time = page.evaluate("Date.now()") + 30000
        while page.evaluate("Date.now()") < end_time:
            if self._is_any_locator_visible([browse_button]):
                return

            for locator in candidates:
                try:
                    count = locator.count()
                    if count == 0:
                        continue
                    for index in range(count):
                        candidate = locator.nth(index)
                        if not candidate.is_visible():
                            continue
                        candidate.click(force=True)
                        page.wait_for_timeout(450)
                        break
                except Exception:
                    continue

            page.wait_for_timeout(250)

        if not self._is_any_locator_visible([browse_button]):
            raise RuntimeError("Nao foi possivel concluir a etapa do botao Confirmar.")

    def _click_incluir_menu(self, page: Page) -> None:
        candidates = [
            page.get_by_text("Incluir", exact=True),
            page.get_by_text("Incluir", exact=True),
            page.locator('wa-menu-popup-item[caption="Incluir"]'),
            page.locator('wa-menu-popup-item:has-text("Incluir")'),
            page.locator('wa-text-view[caption*="Incluir"]'),
            page.locator("wa-menu-popup-item#COMP4602"),
        ]
        self._click_first_visible_locator(candidates, timeout=45000, pause_ms=500)
        page.wait_for_timeout(1200)

    def _finalize_incluir(self, page: Page) -> None:
        candidates = [
            page.get_by_role("button", name="OK"),
            page.locator("#COMP6057"),
            page.get_by_text("OK", exact=True),
            page.locator('wa-button:has-text("OK")'),
            page.locator('button:has-text("OK")'),
        ]

        for locator in candidates:
            try:
                if locator.count() > 0 and locator.first.is_visible():
                    locator.first.click(force=True)
                    page.wait_for_timeout(250)
                    break
            except Exception:
                continue

        self.wait_until_faturamento_screen()

    def _find_login_frame(self, page: Page) -> Frame:
        for _ in range(30):
            for frame in page.frames:
                try:
                    if frame.locator('input[name="login"]').count() > 0:
                        return frame
                except Exception:
                    pass
            page.wait_for_timeout(400)
        raise RuntimeError("Iframe de login nao encontrado.")

    def _fill_lookup_input(self, locator, value: str) -> None:
        page = self._require_page()
        self._fill_host_input(locator, value)

    def _wait_for_visible(self, locator, timeout: int = 30000) -> None:
        locator.wait_for(state="visible", timeout=timeout)
        self._require_page().wait_for_timeout(250)

    def _build_text_locators(self, page: Page, text: str):
        variants = [text]
        if "Atualizações" in text:
            variants.append(text.replace("Atualizações", "Atualizacoes"))
        if "Funções" in text:
            variants.append(text.replace("Funções", "Funcoes"))
        if "Módulos" in text:
            variants.append(text.replace("Módulos", "Modulos"))

        selectors = []
        for variant in variants:
            text_without_counter = re.sub(r"\s*\(\d+\)\s*$", "", variant).strip()
            selectors.extend(
                [
                    page.locator(f'span.caption[title="{variant}"]'),
                    page.locator(f'span.caption[title*="{text_without_counter}"]'),
                    page.locator("span.caption").filter(has_text=variant),
                    page.locator("span.caption").filter(has_text=text_without_counter),
                    page.locator(f'wa-menu-item[caption="{variant}"]'),
                    page.locator(f'wa-menu-popup-item[caption="{variant}"]'),
                    page.locator("wa-menu-item").filter(has_text=variant),
                    page.locator("wa-menu-popup-item").filter(has_text=variant),
                    page.get_by_text(variant, exact=True),
                    page.get_by_text(text_without_counter, exact=True),
                    page.locator(f'text="{variant}"'),
                    page.locator(f'text="{text_without_counter}"'),
                ]
            )
        return selectors

    def _click_first_visible_locator(self, locators, timeout: int = 30000, pause_ms: int = 200) -> None:
        page = self._require_page()
        end_time = page.evaluate("Date.now()") + timeout

        while page.evaluate("Date.now()") < end_time:
            for locator in locators:
                try:
                    count = locator.count()
                    if count == 0:
                        continue
                    for index in range(count):
                        candidate = locator.nth(index)
                        if not candidate.is_visible():
                            continue
                        candidate.scroll_into_view_if_needed(timeout=3000)
                        candidate.click(force=True)
                        page.wait_for_timeout(pause_ms)
                        return
                except Exception:
                    continue
            page.wait_for_timeout(pause_ms)

        raise RuntimeError("Nao foi possivel clicar no elemento esperado do Protheus.")

    def _wait_any_visible(self, locators, timeout: int = 30000) -> None:
        page = self._require_page()
        end_time = page.evaluate("Date.now()") + timeout

        while page.evaluate("Date.now()") < end_time:
            for locator in locators:
                try:
                    count = locator.count()
                    if count == 0:
                        continue
                    for index in range(count):
                        if locator.nth(index).is_visible():
                            return
                except Exception:
                    continue
            page.wait_for_timeout(250)

        raise RuntimeError("Nao foi possivel validar a tela de faturamento.")

    def _is_any_locator_visible(self, locators) -> bool:
        for locator in locators:
            try:
                count = locator.count()
                if count == 0:
                    continue
                for index in range(count):
                    if locator.nth(index).is_visible():
                        return True
            except Exception:
                continue
        return False

    def _first_visible_locator(self, locators):
        for locator in locators:
            try:
                count = locator.count()
                if count == 0:
                    continue
                for index in range(count):
                    candidate = locator.nth(index)
                    if candidate.is_visible():
                        return candidate
            except Exception:
                continue
        return None

    def _wait_for_first_visible_locator(self, locators, timeout: int = 30000, pause_ms: int = 250):
        page = self._require_page()
        end_time = page.evaluate("Date.now()") + timeout

        while page.evaluate("Date.now()") < end_time:
            candidate = self._first_visible_locator(locators)
            if candidate is not None:
                return candidate
            page.wait_for_timeout(pause_ms)

        return None

    @staticmethod
    def _names_match(expected_name: str, found_name: str) -> bool:
        if not expected_name or not found_name:
            return False
        return expected_name == found_name or expected_name in found_name or found_name in expected_name

    @staticmethod
    def _normalize_document(text: str) -> str:
        return "".join(char for char in text if char.isdigit())

    @staticmethod
    def _normalize_text(value: str) -> str:
        replacements = str.maketrans(
            {
                "Á": "A",
                "À": "A",
                "Ã": "A",
                "Â": "A",
                "Ä": "A",
                "É": "E",
                "Ê": "E",
                "Ë": "E",
                "Í": "I",
                "Î": "I",
                "Ï": "I",
                "Ó": "O",
                "Ô": "O",
                "Õ": "O",
                "Ö": "O",
                "Ú": "U",
                "Û": "U",
                "Ü": "U",
                "Ç": "C",
            }
        )
        return " ".join(str(value).strip().upper().translate(replacements).split())

    @classmethod
    def _normalize_cpf_for_lookup(cls, text: str) -> str:
        raw_text = str(text).strip()
        if not raw_text:
            return ""
        if any(char.isalpha() for char in raw_text):
            return ""
        digits = cls._normalize_document(raw_text)
        if not digits or len(digits) > 11:
            return ""
        return digits.zfill(11)

    def _require_page(self) -> Page:
        if self.page is None:
            raise RuntimeError("Pagina do Protheus ainda nao foi inicializada.")
        return self.page

    @staticmethod
    def build_field_selector(field_key: str) -> str:
        field = FIELD_MAP[field_key]
        return f'{field["tag"]}[name="M->{field["protheus_name"]}"]'

    @staticmethod
    def build_lookup_button_selector(field_key: str) -> str:
        return f'{ProtheusBot.build_field_selector(field_key)} button.button-image'

    def _faturamento_dialog(self):
        dialog = self._require_page().locator('wa-dialog.dict-msdialog[title="Contas a Receber"][opened]').last
        dialog.wait_for(state="visible", timeout=30000)
        return dialog

    def _build_billing_payload(self, row: dict) -> dict:
        motivo = str(row.get("MOTIVO", "")).strip()
        modalidade = str(row.get("MODALIDADE", "")).strip()
        modalidade_norm = self._normalize_text(modalidade)
        motivo_norm = self._normalize_text(motivo)

        return {
            "prefixo": "NF",
            "tipo_pg": "0",
            "tipo": "NF",
            "natureza": "SP0000012" if motivo_norm in {self._normalize_text(item) for item in SPECIAL_NATURE_MOTIVES} else "SP0000006",
            "contrato": str(row.get("CONTRATO", "")).strip(),
            "vencimento": datetime.now().strftime("%d/%m/%Y"),
            "valor_titulo": self._format_currency_value(row.get("VALOR")),
            "centro_custo": str(row.get("LOJA", "")).strip(),
            "negocio": "1" if modalidade_norm == "RAC" else "2",
            "segregacao": modalidade,
            "motivo": self._resolve_motivo_codigo(motivo),
            "historico": str(row.get("HISTORICO", "")).strip(),
        }

    def _resolve_motivo_codigo(self, motivo: str) -> str:
        motivo_norm = self._normalize_text(motivo)
        if motivo_norm == "AVARIA/SINISTRO":
            return "0"
        if motivo_norm in {"DIARIAS", "DIARIAS E ITENS EXTRAS", "ITENS EXTRAS"}:
            return "1"
        if motivo_norm in {"BUSCA/ RECUPERACAO", "BUSCA/ RECUPERAÇÃO", "APROPRIACAO", "APROPRIAÇÃO", "APREENSAO AUTORIDADES", "APREENSÃO AUTORIDADES"}:
            return "2"
        if motivo_norm in {"ASSISTENCIA 24 H", "ASSISTÊNCIA 24 H"}:
            return "3"
        if motivo_norm == "EXTRACONTRATUAL":
            return "4"
        raise RuntimeError(f"Motivo sem mapeamento para o Protheus: {motivo}")

    @staticmethod
    def _format_currency_value(value) -> str:
        if value is None or value == "":
            return ""
        try:
            numeric_value = float(value)
        except Exception:
            text_value = str(value).strip().replace(".", "").replace(",", ".")
            numeric_value = float(text_value)
        return f"{numeric_value:.2f}"

    def _fill_text_field(self, dialog, field_key: str, value: str) -> None:
        locator = dialog.locator(self.build_field_selector(field_key)).last
        locator.wait_for(state="visible", timeout=30000)
        self._fill_host_input(locator, value)
        self._require_page().wait_for_timeout(250)

    def _fill_combo_field(self, dialog, field_key: str, value: str) -> None:
        locator = dialog.locator(self.build_field_selector(field_key)).last
        locator.wait_for(state="visible", timeout=30000)
        locator.click(force=True)
        self._require_page().keyboard.press("Control+A")
        self._require_page().keyboard.press("Backspace")
        self._require_page().keyboard.type(str(value), delay=60)
        self._require_page().keyboard.press("Enter")
        self._require_page().wait_for_timeout(300)

    def _fill_host_input(self, locator, value: str) -> None:
        if locator is None:
            raise RuntimeError("Campo de entrada nao foi localizado.")
        page = self._require_page()
        locator.wait_for(state="visible", timeout=45000)
        locator.scroll_into_view_if_needed(timeout=3000)
        locator.click(force=True)

        filled = False
        try:
            filled = bool(
                locator.evaluate(
                    """
                    (host, value) => {
                        const root = host.shadowRoot || host;
                        const input =
                            root.querySelector('input') ||
                            root.querySelector('textarea') ||
                            host.querySelector('input') ||
                            host.querySelector('textarea');
                        if (!input) return false;
                        input.focus();
                        input.value = '';
                        input.dispatchEvent(new Event('input', { bubbles: true, composed: true }));
                        input.value = String(value);
                        input.dispatchEvent(new Event('input', { bubbles: true, composed: true }));
                        input.dispatchEvent(new Event('change', { bubbles: true, composed: true }));
                        return true;
                    }
                    """,
                    value,
                )
            )
        except Exception:
            filled = False

        if not filled:
            locator.click(force=True)
            page.keyboard.press("Control+A")
            page.keyboard.press("Backspace")
            page.keyboard.type(str(value), delay=110)

        page.wait_for_timeout(700)
        self._dismiss_help_popup_if_present()

    def _click_and_fill_field(self, locator, value: str) -> None:
        if locator is None:
            raise RuntimeError("Campo de entrada nao foi localizado.")

        page = self._require_page()
        locator.wait_for(state="visible", timeout=45000)
        locator.scroll_into_view_if_needed(timeout=3000)
        locator.click(force=True)
        page.wait_for_timeout(350)

        filled = False
        try:
            page.keyboard.press("Control+A")
            page.wait_for_timeout(150)
            page.keyboard.type(str(value), delay=120)
            filled = True
        except Exception:
            filled = False

        if not filled:
            try:
                page.keyboard.press("Control+A")
                page.wait_for_timeout(100)
                page.keyboard.type(str(value), delay=120)
                filled = True
            except Exception:
                filled = False

        if not filled:
            self._fill_host_input(locator, value)
            return

        page.wait_for_timeout(800)
        self._dismiss_help_popup_if_present()

    def _select_combo_option(self, locator, value: str) -> None:
        if locator is None:
            raise RuntimeError("Combobox nao localizado.")
        locator.wait_for(state="visible", timeout=45000)
        locator.scroll_into_view_if_needed(timeout=3000)
        locator.select_option(str(value))
        self._require_page().wait_for_timeout(800)
        self._dismiss_help_popup_if_present()

    def _settle_after_field(self) -> None:
        page = self._require_page()
        page.wait_for_timeout(500)
        self._dismiss_help_popup_if_present()
        page.wait_for_timeout(250)

    def _dismiss_help_popup_if_present(self) -> None:
        page = self._require_page()
        try:
            help_dialog = page.locator("wa-dialog.dict-msdialog").filter(
                has=page.get_by_text("Problema:", exact=False)
            ).last
            if help_dialog.count() == 0 or not help_dialog.is_visible():
                return

            close_button = self._first_visible_locator(
                [
                    help_dialog.get_by_role("button", name="Fechar"),
                    help_dialog.get_by_text("Fechar", exact=True),
                ]
            )
            if close_button is not None:
                close_button.click(force=True)
                page.wait_for_timeout(700)
        except Exception:
            pass

    def _complete_save_flow(self, dialog) -> None:
        page = self._require_page()

        first_save_candidates = [
            dialog.get_by_role("button", name="Salvar"),
            dialog.get_by_text("Salvar", exact=True),
        ]

        first_save = self._first_visible_locator(first_save_candidates)
        if first_save is None:
            raise RuntimeError("Primeiro botao Salvar nao foi localizado na tela de faturamento.")

        self._click_first_visible_locator([first_save], timeout=45000, pause_ms=700)
        page.wait_for_timeout(1200)

        second_save_candidates = [
            page.get_by_role("button", name="Salvar"),
            page.get_by_text("Salvar", exact=True),
        ]
        second_save = self._wait_for_first_visible_locator(second_save_candidates, timeout=20000, pause_ms=700)
        if second_save is None:
            raise RuntimeError("Segundo botao Salvar da confirmacao contabil nao apareceu.")
        self._click_first_visible_locator([second_save], timeout=20000, pause_ms=700)
        page.wait_for_timeout(1500)

        close_candidates = [
            page.get_by_role("button", name="Fechar"),
            page.get_by_text("Fechar", exact=True),
        ]

        for _ in range(2):
            close_button = self._wait_for_first_visible_locator(close_candidates, timeout=7000, pause_ms=700)
            if close_button is None:
                break
            try:
                self._click_first_visible_locator([close_button], timeout=7000, pause_ms=700)
                page.wait_for_timeout(1500)
            except Exception:
                break

        self.wait_until_faturamento_screen()

    def _read_from_locator(self, locator) -> str:
        if locator is None:
            return ""
        try:
            if locator.evaluate("el => el.tagName && el.tagName.toLowerCase() === 'input'"):
                return str(locator.input_value()).strip()
        except Exception:
            pass
        try:
            return str(
                locator.evaluate(
                    """
                    (host) => {
                        const root = host.shadowRoot || host;
                        const input =
                            root.querySelector('input') ||
                            root.querySelector('textarea') ||
                            host.querySelector('input') ||
                            host.querySelector('textarea');
                        return input ? input.value || '' : '';
                    }
                    """
                )
            ).strip()
        except Exception:
            return ""

    def _read_locator_text(self, locator) -> str:
        if locator is None:
            return ""
        try:
            return str(locator.inner_text() or "").strip()
        except Exception:
            pass
        try:
            return str(locator.text_content() or "").strip()
        except Exception:
            return ""

    def _read_field_value(self, dialog, field_key: str) -> str:
        locator = dialog.locator(self.build_field_selector(field_key)).last
        try:
            value = locator.evaluate(
                """
                (host) => {
                    const root = host.shadowRoot || host;
                    const input =
                        root.querySelector('input') ||
                        root.querySelector('textarea') ||
                        host.querySelector('input') ||
                        host.querySelector('textarea');
                    return input ? input.value || '' : '';
                }
                """
            )
            return str(value).strip()
        except Exception:
            return ""


# ===== main.py =====
import getpass
import queue
import threading
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox

import customtkinter as ctk
import requests
try:
    from PIL import Image
except Exception:
    Image = None

URL_VALIDACAO = "https://raw.githubusercontent.com/diogodiasyt-blip/validacaofoco/refs/heads/main/chave"
URL_PING_ABERTURA = "https://docs.google.com/forms/d/e/1FAIpQLScmxNbTO-vXw0LEOKIyEhSpIl9aTbw8x5hnEI5VY2eVMRh5gQ/formResponse"



class BillingApp(ctk.CTk):
    MAIN_BG = "#f6f4f1"
    CARD_BG = "#ffffff"
    CARD_BORDER = "#eadfdb"
    PRIMARY_TEXT = "#d81919"
    MUTED_TEXT = "#5c5c5c"
    BUTTON_BG = "#ef1a14"
    BUTTON_ACTIVE_BG = "#c91410"
    SUCCESS_TEXT = "#187a2f"
    SOFT_RED = "#fff1ef"

    def __init__(self) -> None:
        ctk.set_appearance_mode("light")
        ctk.set_default_color_theme("blue")
        super().__init__()

        self.title(APP_TITLE)
        self.geometry(APP_GEOMETRY)
        self.minsize(980, 720)
        self.configure(fg_color=self.MAIN_BG)

        self.workbook_context: WorkbookContext | None = None
        self.processing_thread: threading.Thread | None = None
        self.bot_instance: ProtheusBot | None = None
        self.ui_queue: queue.Queue = queue.Queue()
        self.logo_image = None
        self.logo_label = None
        self.pause_requested = threading.Event()
        self.stop_requested = threading.Event()
        self.is_processing = False
        self.is_paused = False
        self.session_ready = False

        self.username_var = ctk.StringVar()
        self.password_var = ctk.StringVar()
        self.file_path_var = ctk.StringVar()
        self.headless_var = ctk.BooleanVar(value=False)

        self.total_var = ctk.StringVar(value="0")
        self.ready_var = ctk.StringVar(value="0")
        self.success_var = ctk.StringVar(value="0")
        self.ignored_var = ctk.StringVar(value="0")
        self.error_var = ctk.StringVar(value="0")

        self.validate_button: ctk.CTkButton | None = None
        self.start_button: ctk.CTkButton | None = None
        self.pause_button: ctk.CTkButton | None = None
        self.stop_button: ctk.CTkButton | None = None
        self.counter_value_labels: dict[str, ctk.CTkLabel] = {}

        self._build_layout()
        self._update_action_buttons()
        self.after(100, self._process_ui_queue)

    def _build_layout(self) -> None:
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)

        container = ctk.CTkFrame(self, fg_color=self.MAIN_BG, corner_radius=0)
        container.grid(row=0, column=0, sticky="nsew", padx=12, pady=12)
        container.grid_columnconfigure(0, weight=1)
        container.grid_rowconfigure(0, weight=1)

        scroll = ctk.CTkScrollableFrame(container, fg_color=self.MAIN_BG, corner_radius=0)
        scroll.grid(row=0, column=0, sticky="nsew")
        scroll.grid_columnconfigure(0, weight=1)

        self._build_hero(scroll)
        self._build_access_section(scroll)
        self._build_plan_section(scroll)
        self._build_indicator_section(scroll)
        self._build_execution_section(scroll)

    def _build_hero(self, parent: ctk.CTkScrollableFrame) -> None:
        hero = ctk.CTkFrame(
            parent,
            fg_color=self.CARD_BG,
            corner_radius=26,
            border_width=1,
            border_color=self.CARD_BORDER,
        )
        hero.grid(row=0, column=0, sticky="ew", padx=8, pady=(8, 14))

        hero_inner = ctk.CTkFrame(hero, fg_color="transparent")
        hero_inner.pack(fill="x", padx=24, pady=24)

        logo = self._load_logo()
        if logo:
            try:
                self.logo_label = ctk.CTkLabel(hero_inner, text="", image=logo)
                self.logo_label.pack(side="left", padx=(0, 18))
            except Exception:
                self.logo_image = None
                self._build_logo_text(hero_inner)
        else:
            self._build_logo_text(hero_inner)

        header_text = ctk.CTkFrame(hero_inner, fg_color="transparent")
        header_text.pack(side="left", fill="x", expand=True)

        ctk.CTkLabel(
            header_text,
            text="Faturamento de Avarias",
            text_color=self.PRIMARY_TEXT,
            font=("Segoe UI", 30, "bold"),
        ).pack(anchor="w")
        ctk.CTkLabel(
            header_text,
            text="Automacao do Protheus com validacao de cliente, progresso em tempo real e rastreabilidade na planilha.",
            text_color=self.MUTED_TEXT,
            font=("Segoe UI", 14),
        ).pack(anchor="w", pady=(6, 0))
        ctk.CTkLabel(
            header_text,
            text="OPERACAO DE FATURAMENTO",
            text_color="#a65f56",
            font=("Segoe UI", 12, "bold"),
        ).pack(anchor="w", pady=(10, 0))

    def _load_logo(self):
        logo_candidates = []

        env_logo = os.environ.get("FOCO_LOGO_PNG", "").strip()
        if env_logo:
            logo_candidates.append(Path(env_logo))

        env_assets = os.environ.get("FOCO_ASSETS_DIR", "").strip()
        if env_assets:
            logo_candidates.append(Path(env_assets) / "logo.png")

        logo_candidates.append(Path(__file__).with_name("logo.png"))
        logo_candidates.append(Path(__file__).resolve().parent.parent / "assets" / "logo.png")

        for logo_path in logo_candidates:
            try:
                if not logo_path.exists() or Image is None:
                    continue
                image = Image.open(logo_path)
                self.logo_image = ctk.CTkImage(light_image=image, dark_image=image, size=(86, 52))
                return self.logo_image
            except Exception:
                continue
        return None

    def _build_logo_text(self, parent: ctk.CTkFrame) -> None:
        brand = ctk.CTkFrame(parent, fg_color="transparent")
        brand.pack(side="left", padx=(0, 18))
        ctk.CTkLabel(
            brand,
            text="foco,",
            text_color=self.PRIMARY_TEXT,
            font=("Segoe UI", 24, "bold"),
        ).pack(anchor="w")
        ctk.CTkLabel(
            brand,
            text="aluguel de carros",
            text_color="#c7463f",
            font=("Segoe UI", 10, "bold"),
        ).pack(anchor="w", pady=(0, 0))

    def run(self) -> None:
        self.mainloop()

    def _build_access_section(self, parent: ctk.CTkScrollableFrame) -> None:
        access = self._create_section(parent, 1, "Acesso ao Protheus")
        grid = ctk.CTkFrame(access, fg_color="transparent")
        grid.pack(fill="x", padx=18, pady=(0, 18))
        grid.grid_columnconfigure((0, 1), weight=1)

        self._form_label(grid, "Login").grid(row=0, column=0, sticky="w", padx=(0, 10), pady=(0, 6))
        self._form_label(grid, "Senha").grid(row=0, column=1, sticky="w", padx=(10, 0), pady=(0, 6))
        self._entry(grid, self.username_var).grid(row=1, column=0, sticky="ew", padx=(0, 10))
        self._entry(grid, self.password_var, show="*").grid(row=1, column=1, sticky="ew", padx=(10, 0))

    def _build_plan_section(self, parent: ctk.CTkScrollableFrame) -> None:
        plan = self._create_section(parent, 2, "Planilha")
        grid = ctk.CTkFrame(plan, fg_color="transparent")
        grid.pack(fill="x", padx=18, pady=(0, 18))
        grid.grid_columnconfigure(0, weight=1)

        self._form_label(grid, "Planilha").grid(row=0, column=0, sticky="w", pady=(0, 6))
        self._entry(grid, self.file_path_var).grid(row=1, column=0, sticky="ew", padx=(0, 10), pady=(0, 10))
        self._secondary_button(grid, "Selecionar", self.select_excel_file, width=150).grid(row=1, column=1, sticky="e", pady=(0, 10))

        ctk.CTkCheckBox(
            grid,
            text="Modo invisivel",
            variable=self.headless_var,
            text_color="#303030",
            fg_color=self.BUTTON_BG,
            hover_color=self.BUTTON_ACTIVE_BG,
            border_color="#d9c9c3",
        ).grid(row=2, column=0, sticky="w")

    def _build_indicator_section(self, parent: ctk.CTkScrollableFrame) -> None:
        indicators = self._create_section(parent, 3, "Indicadores")
        counters = ctk.CTkFrame(indicators, fg_color="transparent")
        counters.pack(fill="x", padx=18, pady=(0, 12))
        for index in range(5):
            counters.grid_columnconfigure(index, weight=1)

        self._counter_card(counters, 0, "Total", "total", self.total_var)
        self._counter_card(counters, 1, "Aptos", "ready", self.ready_var)
        self._counter_card(counters, 2, "Faturados", "success", self.success_var)
        self._counter_card(counters, 3, "Ignorados", "ignored", self.ignored_var)
        self._counter_card(counters, 4, "Erros", "error", self.error_var)
        self._refresh_counter_cards()

    def _build_execution_section(self, parent: ctk.CTkScrollableFrame) -> None:
        body = self._create_section(parent, 4, "Execucao")
        content = ctk.CTkFrame(body, fg_color="transparent")
        content.pack(fill="both", expand=True, padx=18, pady=(0, 18))
        content.grid_columnconfigure(0, weight=1)
        content.grid_rowconfigure(1, weight=1)

        self.progress = ctk.CTkProgressBar(
            content,
            progress_color=self.BUTTON_BG,
            fg_color="#f1dfdb",
            corner_radius=20,
            height=18,
        )
        self.progress.set(0)
        self.progress.grid(row=1, column=0, sticky="ew", padx=18, pady=(0, 10))

        self.log_box = ctk.CTkTextbox(
            content,
            corner_radius=14,
            fg_color="#fffdfc",
            border_width=1,
            border_color=self.CARD_BORDER,
            text_color="#303030",
        )
        self.log_box.grid(row=2, column=0, sticky="nsew", padx=18, pady=(0, 10))

        actions = ctk.CTkFrame(content, fg_color="transparent")
        actions.grid(row=3, column=0, sticky="ew", padx=18, pady=(0, 0))
        actions.grid_columnconfigure(0, weight=1)

        self.validate_button = self._secondary_button(actions, "Validar Planilha", self.validate_selected_workbook)
        self.validate_button.grid(row=0, column=1, padx=6, pady=6)

        self.pause_button = self._secondary_button(actions, "Pausar", self.pause_processing, width=140)
        self.pause_button.grid(row=0, column=2, padx=6, pady=6)

        self.stop_button = self._secondary_button(actions, "Parar", self.stop_processing, width=140)
        self.stop_button.grid(row=0, column=3, padx=6, pady=6)

        self.start_button = self._primary_button(actions, "Iniciar", self.start_processing)
        self.start_button.grid(row=0, column=4, padx=6, pady=6)

    def _create_section(self, parent: ctk.CTkScrollableFrame, row: int, title: str) -> ctk.CTkFrame:
        frame = ctk.CTkFrame(
            parent,
            fg_color=self.CARD_BG,
            corner_radius=22,
            border_width=1,
            border_color=self.CARD_BORDER,
        )
        frame.grid(row=row, column=0, sticky="ew", padx=8, pady=(0, 14))
        frame.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(
            frame,
            text=title,
            text_color=self.PRIMARY_TEXT,
            font=("Segoe UI", 18, "bold"),
        ).pack(anchor="w", padx=18, pady=(16, 12))
        return frame

    def _entry(self, parent: ctk.CTkBaseClass, variable: ctk.StringVar, show: str | None = None) -> ctk.CTkEntry:
        return ctk.CTkEntry(
            parent,
            textvariable=variable,
            show=show,
            height=44,
            corner_radius=12,
            fg_color="#ffffff",
            border_color="#eadfdb",
            text_color="#303030",
        )

    def _form_label(self, parent: ctk.CTkBaseClass, text: str) -> ctk.CTkLabel:
        return ctk.CTkLabel(
            parent,
            text=text,
            font=("Segoe UI", 13, "bold"),
            text_color="#303030",
        )

    def _primary_button(self, parent: ctk.CTkBaseClass, text: str, command, width: int = 160) -> ctk.CTkButton:
        return ctk.CTkButton(
            parent,
            text=text,
            command=command,
            height=44,
            width=width,
            corner_radius=14,
            fg_color=self.BUTTON_BG,
            hover_color=self.BUTTON_ACTIVE_BG,
            text_color="#ffffff",
            font=("Segoe UI", 14, "bold"),
        )

    def _secondary_button(self, parent: ctk.CTkBaseClass, text: str, command, width: int = 170) -> ctk.CTkButton:
        return ctk.CTkButton(
            parent,
            text=text,
            command=command,
            height=44,
            width=width,
            corner_radius=14,
            fg_color="#ffffff",
            text_color=self.PRIMARY_TEXT,
            hover_color=self.SOFT_RED,
            border_width=1,
            border_color="#f0d7d2",
            font=("Segoe UI", 14, "bold"),
        )

    def _counter_card(self, parent: ctk.CTkFrame, column: int, title: str, key: str, variable: ctk.StringVar) -> None:
        card = ctk.CTkFrame(
            parent,
            fg_color="#fffdfc",
            corner_radius=16,
            border_width=1,
            border_color="#f0d7d2",
        )
        card.grid(row=0, column=column, sticky="ew", padx=6, pady=6)
        ctk.CTkLabel(
            card,
            text=title,
            text_color=self.MUTED_TEXT,
            font=("Segoe UI", 12, "bold"),
        ).pack(anchor="w", padx=12, pady=(10, 2))
        value_label = ctk.CTkLabel(
            card,
            text=variable.get(),
            text_color=self.PRIMARY_TEXT,
            font=ctk.CTkFont(size=24, weight="bold"),
        )
        value_label.pack(anchor="w", padx=12, pady=(0, 10))
        self.counter_value_labels[key] = value_label

    def _refresh_counter_cards(self) -> None:
        values = {
            "total": self.total_var.get(),
            "ready": self.ready_var.get(),
            "success": self.success_var.get(),
            "ignored": self.ignored_var.get(),
            "error": self.error_var.get(),
        }
        for key, value in values.items():
            label = self.counter_value_labels.get(key)
            if label is not None:
                label.configure(text=value)

    def select_excel_file(self) -> None:
        selected = filedialog.askopenfilename(
            title="Selecione a planilha de faturamento",
            filetypes=[("Excel", "*.xlsx *.xls")],
        )
        if selected:
            self.file_path_var.set(selected)
            self.workbook_context = None
            self.log(f"Planilha selecionada: {selected}")
            self._update_action_buttons()

    def validate_selected_workbook(self) -> None:
        path = self.file_path_var.get().strip()
        if not path:
            messagebox.showwarning(APP_TITLE, "Selecione uma planilha antes de validar.")
            return

        context = load_workbook(path)
        missing = validate_workbook(context)
        if missing:
            messagebox.showerror(APP_TITLE, f"Colunas obrigatorias ausentes: {', '.join(missing)}")
            return

        apt, ignored = prepare_rows(context)
        self.workbook_context = context
        self.total_var.set(str(len(context.dataframe)))
        self.ready_var.set(str(len(apt)))
        self.ignored_var.set(str(len(ignored)))
        self.success_var.set("0")
        self.error_var.set("0")
        self._refresh_counter_cards()
        self.progress.set(0)
        self.log("Planilha validada com sucesso.")
        self.log(f"Linhas aptas: {len(apt)} | Ignoradas na triagem: {len(ignored)}")
        self._update_action_buttons()

    def registrar_abertura(self) -> None:
        try:
            data = {
                "entry.846583903": getpass.getuser(),
                "entry.1509395143": datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
            }
            requests.post(URL_PING_ABERTURA, data=data, timeout=5)
            self.log("Abertura liberada.")
        except Exception as exc:
            self.log(f"Falha ao registrar abertura: {exc}")

    def verificar_chave(self) -> bool:
        try:
            response = requests.get(URL_VALIDACAO, timeout=10)
            status = response.text.strip().upper()
            self.log(f"STATUS DO ROBÔ: {status or 'INDEFINIDO'}")
            return status == "ATIVO"
        except Exception as exc:
            self.log(f"Falha ao validar chave remota: {exc}. Mantendo execução liberada.")
            return True

    def start_processing(self) -> None:
        if self.processing_thread and self.processing_thread.is_alive():
            return

        if not self.workbook_context:
            self.validate_selected_workbook()
            if not self.workbook_context:
                return

        self.pause_requested.clear()
        self.stop_requested.clear()
        self.is_paused = False
        self.is_processing = True
        self.log("Thread de processamento iniciada.")
        self.processing_thread = threading.Thread(target=self._run_processing, daemon=True)
        self.processing_thread.start()
        self._update_action_buttons()

    def pause_processing(self) -> None:
        if not self.is_processing:
            return
        self.pause_requested.set()
        self.log("Pausa solicitada. A automacao vai parar com seguranca ao concluir a linha atual.")
        self._update_action_buttons()

    def stop_processing(self) -> None:
        if not self.is_processing:
            return
        self.stop_requested.set()
        self.log("Parada solicitada. A automacao vai encerrar com seguranca ao concluir a linha atual.")
        self._update_action_buttons()

    def _run_processing(self) -> None:
        assert self.workbook_context is not None

        try:
            self.registrar_abertura()
            if not self.verificar_chave():
                self.log("Robô bloqueado por validação remota.")
                self._post_ui(messagebox.showerror, "Bloqueado", "Este robô está temporariamente desativado.")
                return

            self.log("Iniciando triagem das linhas aptas...")
            apt, ignored = prepare_rows(self.workbook_context)
            self.log(f"Triagem concluída. Aptas: {len(apt)} | Ignoradas: {len(ignored)}")
            status_series = self.workbook_context.dataframe["STATUS_PROCESSAMENTO"].fillna("").astype(str).str.strip()

            for row_index in ignored.index:
                if not status_series.loc[row_index]:
                    mark_row(
                        self.workbook_context,
                        row_index,
                        "IGNORADO_TRIAGEM",
                        "Linha ignorada na triagem inicial por ausencia de dados minimos, faturamento ja existente ou CPF invalido.",
                    )

            pending = apt.loc[status_series.loc[apt.index].eq("")].copy()
            self._refresh_counters()

            if pending.empty:
                self.log("Nenhuma linha pendente apta para processamento.")
                return

            if self.bot_instance is None:
                self.log("Preparando instância do bot Protheus...")
                self.bot_instance = ProtheusBot(
                    username=self.username_var.get().strip(),
                    password=self.password_var.get().strip(),
                    log=self.log,
                    headless=self.headless_var.get(),
                )

            if not self.session_ready:
                self.log("Inicializando navegador e sessão do Protheus...")
                self.bot_instance.start()
                self.log("Sessão iniciada. Executando login...")
                self.bot_instance.login()
                self.log("Login concluído. Navegando para faturamento...")
                self.bot_instance.navigate_to_faturamento()
                self.session_ready = True
                self.log("Tela de faturamento pronta para processamento.")

            total = len(pending)
            for position, (row_index, row) in enumerate(pending.iterrows(), start=1):
                if self.stop_requested.is_set():
                    self.log("Processamento interrompido com seguranca antes da proxima linha.")
                    break
                if self.pause_requested.is_set():
                    self.log("Processamento pausado com seguranca antes da proxima linha.")
                    break

                self.log(f"Processando linha {position} de {total}.")
                try:
                    result = self.bot_instance.process_row(row.to_dict())
                    mark_row(
                        self.workbook_context,
                        row_index,
                        result["status"],
                        result["detail"],
                        result.get("invoice", ""),
                    )
                    if result["status"] == "SALVO" and result.get("invoice", "").strip():
                        report_path = append_report_entry(
                            base_dir=Path(self.file_path_var.get().strip()).parent,
                            contract=row.get("CONTRATO", ""),
                            value=row.get("VALOR", ""),
                            invoice_number=result.get("invoice", ""),
                        )
                        self.log(f"Relatorio atualizado em tempo real: {report_path}")
                except Exception as exc:  # pragma: no cover
                    mark_row(self.workbook_context, row_index, "ERRO_AUTOMACAO", str(exc))
                    self.log(f"Erro na linha {row_index}: {exc}")

                self._set_progress(position / total)
                save_workbook(self.workbook_context, self._build_output_path())
                self._refresh_counters()

            if not self.pause_requested.is_set() and not self.stop_requested.is_set():
                self.log(f"Processamento finalizado. Arquivo salvo em: {self._build_output_path()}")
        except Exception as exc:
            print(f"[FATURAMENTO][ERRO_THREAD] {exc}", file=sys.stderr)
            try:
                self.log(f"Erro fatal no processamento: {exc}")
            except Exception:
                pass
        finally:
            self.is_processing = False
            self.is_paused = self.pause_requested.is_set() and not self.stop_requested.is_set()

            if self.stop_requested.is_set():
                if self.bot_instance is not None:
                    self.bot_instance.stop()
                self.bot_instance = None
                self.session_ready = False
                self.log("Sessao encerrada apos a solicitacao de parada.")
            elif not self.is_paused and self.bot_instance is not None:
                self.bot_instance.stop()
                self.bot_instance = None
                self.session_ready = False

            self.pause_requested.clear()
            self.stop_requested.clear()
            self._post_ui(self._refresh_counters)
            self._post_ui(self._update_action_buttons)

    def _build_output_path(self) -> Path:
        source = Path(self.file_path_var.get().strip())
        if not str(source).strip():
            return Path.cwd() / "planilha_processada.xlsx"
        return source.parent / f"{source.stem}_processada.xlsx"

    def _refresh_counters(self) -> None:
        if threading.current_thread() is not threading.main_thread():
            self._post_ui(self._refresh_counters)
            return

        if not self.workbook_context:
            return

        apt, _ = prepare_rows(self.workbook_context)
        dataframe = self.workbook_context.dataframe
        status_series = dataframe["STATUS_PROCESSAMENTO"].fillna("").astype(str).str.strip()

        success = ((status_series != "") & ~status_series.str.startswith("IGNORADO") & ~status_series.str.startswith("ERRO")).sum()
        ignored = status_series.str.startswith("IGNORADO").sum()
        errors = status_series.str.startswith("ERRO").sum()
        remaining = status_series.loc[apt.index].eq("").sum() if not apt.empty else 0

        self.total_var.set(str(len(dataframe)))
        self.ready_var.set(str(int(remaining)))
        self.success_var.set(str(int(success)))
        self.ignored_var.set(str(int(ignored)))
        self.error_var.set(str(int(errors)))
        self._refresh_counter_cards()

    def _set_progress(self, value: float) -> None:
        value = max(0.0, min(1.0, float(value)))
        if threading.current_thread() is threading.main_thread():
            self.progress.set(value)
        else:
            self._post_ui(self.progress.set, value)

    def _has_valid_execution_context(self) -> bool:
        return self.workbook_context is not None and int(self.ready_var.get() or "0") > 0

    def _update_action_buttons(self) -> None:
        if self.validate_button is None or self.start_button is None or self.pause_button is None or self.stop_button is None:
            return

        is_running = self.is_processing
        can_start = (not is_running) and self._has_valid_execution_context()

        self.validate_button.configure(state="disabled" if is_running else "normal")
        self.start_button.configure(state="normal" if can_start else "disabled")
        self.start_button.configure(text="Continuar" if self.is_paused else "Iniciar")
        self.pause_button.configure(state="normal" if is_running and not self.pause_requested.is_set() else "disabled")
        self.stop_button.configure(state="normal" if is_running and not self.stop_requested.is_set() else "disabled")

    def log(self, message: str) -> None:
        if threading.current_thread() is threading.main_thread():
            self._append_log(message)
        else:
            self._post_ui(self._append_log, message)

    def _append_log(self, message: str) -> None:
        self.log_box.insert("end", f"{message}\n")
        self.log_box.see("end")

    def _post_ui(self, callback, *args, **kwargs) -> None:
        self.ui_queue.put((callback, args, kwargs))

    def _process_ui_queue(self) -> None:
        try:
            while True:
                callback, args, kwargs = self.ui_queue.get_nowait()
                try:
                    callback(*args, **kwargs)
                except Exception as exc:
                    print(f"[FATURAMENTO][ERRO_UI] {exc}", file=sys.stderr)
        except queue.Empty:
            pass
        self.after(100, self._process_ui_queue)


if __name__ == "__main__":
    app = BillingApp()
    app.run()
