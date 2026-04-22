from __future__ import annotations

import os
import unicodedata
from dataclasses import dataclass
from pathlib import Path
from typing import Dict

import customtkinter as ctk
import pandas as pd
from PIL import Image
from tkinter import filedialog, messagebox


APP_TITLE = "Robô de Lançamentos de Caixa"
APP_GEOMETRY = "1180x760"

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


FLOW_SPECS = {
    "cash": {
        "label": "Cash",
        "sheet_name": "CASH",
        "header_row": 1,
        "pending_column": "BAIXADO",
        "required_columns": [
            "CONTRATO",
            "LOJA",
            "VALOR",
            "DATA PAGAMENTO",
            "BAIXADO",
        ],
        "aliases": {
            "CONTRATO": {"CONTRATO"},
            "LOJA": {"LOJA", "LOJA "},
            "VALOR": {"VALOR", "VALOR ", "VALOR PAGO"},
            "DATA PAGAMENTO": {"DATA PAGAMENTO", "DATA_PAGAMENTO"},
            "BAIXADO": {"BAIXADO", "STATUS"},
        },
        "group_hint": "Agrupamento previsto: somar cash por loja.",
    },
    "despesas": {
        "label": "Despesas",
        "sheet_name": "DESPESAS",
        "header_row": 0,
        "pending_column": "STATUS",
        "required_columns": [
            "LOJA",
            "DATA DE DESPESA",
            "TIPO DA DESPESA",
            "DESCRICAO",
            "VALOR",
            "STATUS",
        ],
        "aliases": {
            "LOJA": {"LOJA"},
            "DATA DE DESPESA": {"DATA DE DESPESA", "DATA DESPESA"},
            "TIPO DA DESPESA": {"TIPO DA DESPESA", "TIPO DESPESA"},
            "DESCRICAO": {"DESCRICAO", "DESCRIÇÃO", "DESCRICAO ", "DESCRIÇÃO "},
            "VALOR": {"VALOR"},
            "STATUS": {"STATUS", "BAIXADO"},
        },
        "group_hint": "Agrupamento previsto: separar despesas por loja e tipo para rateio.",
    },
    "depositos": {
        "label": "Depósitos",
        "sheet_name": "DEPÓSITOS",
        "header_row": 1,
        "pending_column": "BAIXADO",
        "required_columns": [
            "DATA",
            "BANCO",
            "AGENCIA",
            "VALOR",
            "LOJA",
            "BAIXADO",
        ],
        "aliases": {
            "DATA": {"DATA"},
            "BANCO": {"BANCO"},
            "AGENCIA": {"AGENCIA", "AGÊNCIA"},
            "VALOR": {"VALOR"},
            "LOJA": {"LOJA"},
            "BAIXADO": {"BAIXADO", "STATUS"},
        },
        "group_hint": "Agrupamento previsto: consolidar depósitos por loja, dia e banco.",
    },
}


def _normalize_text(value: str) -> str:
    text = unicodedata.normalize("NFKD", str(value or "")).encode("ASCII", "ignore").decode("ASCII")
    return " ".join(text.strip().upper().split())


def _resolve_sheet_name(sheet_names: list[str], target_name: str) -> str | None:
    target_norm = _normalize_text(target_name)
    for sheet_name in sheet_names:
        if _normalize_text(sheet_name) == target_norm:
            return sheet_name
    return None


def _load_logo_candidates() -> list[Path]:
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


@dataclass
class ValidationResult:
    flow_key: str
    flow_label: str
    sheet_name: str
    total_rows: int
    pending_rows: int
    launched_rows: int
    missing_columns: list[str]
    workbook_sheets: list[str]


class RoboLancamentosCaixaApp(ctk.CTk):
    PRIMARY_TEXT = PRIMARY_TEXT
    CARD_BG = CARD_BG
    CARD_BORDER = CARD_BORDER
    BUTTON_BG = BUTTON_BG
    BUTTON_ACTIVE_BG = BUTTON_ACTIVE_BG

    def __init__(self) -> None:
        super().__init__()
        ctk.set_appearance_mode("light")
        ctk.set_default_color_theme("blue")

        self.title(APP_TITLE)
        self.geometry(APP_GEOMETRY)
        self.minsize(1080, 720)
        self.configure(fg_color=MAIN_BG)

        self.username_var = ctk.StringVar()
        self.password_var = ctk.StringVar()
        self.file_path_var = ctk.StringVar()
        self.flow_var = ctk.StringVar(value="cash")
        self.headless_var = ctk.BooleanVar(value=False)

        self.total_var = ctk.StringVar(value="0")
        self.pending_var = ctk.StringVar(value="0")
        self.ready_var = ctk.StringVar(value="0")
        self.error_var = ctk.StringVar(value="0")
        self.sheet_var = ctk.StringVar(value="-")
        self.status_var = ctk.StringVar(value="Aguardando validação")

        self.validation_result: ValidationResult | None = None
        self.logo_image = self._load_logo()

        self._build_layout()
        self._update_flow_description()
        self._update_action_buttons()

    def _load_logo(self):
        for candidate in _load_logo_candidates():
            try:
                if candidate.exists():
                    image = Image.open(candidate)
                    return ctk.CTkImage(light_image=image, dark_image=image, size=(86, 52))
            except Exception:
                continue
        return None

    def _build_layout(self) -> None:
        container = ctk.CTkScrollableFrame(self, fg_color="transparent", corner_radius=0)
        container.pack(fill="both", expand=True, padx=22, pady=22)
        container.grid_columnconfigure(0, weight=1)

        self._build_header(container)
        self._build_access_section(container)
        self._build_plan_section(container)
        self._build_indicator_section(container)
        self._build_execution_section(container)

        footer = ctk.CTkLabel(
            container,
            text="Criado por Diogo Medeiros © 2026",
            text_color="#ef574f",
            font=("Segoe UI", 12),
        )
        footer.grid(row=5, column=0, sticky="w", padx=6, pady=(6, 0))

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
        ctk.CTkLabel(texts, text="Lançamentos de Caixa", text_color=PRIMARY_TEXT, font=("Segoe UI", 30, "bold")).pack(anchor="w")
        ctk.CTkLabel(
            texts,
            text="Automação de cash, despesas e depósitos com leitura da planilha padrão e validação prévia.",
            text_color="#4b5563",
            font=("Segoe UI", 15),
        ).pack(anchor="w", pady=(8, 0))
        ctk.CTkLabel(
            texts,
            text="ESCOP0 1: INTERFACE + VALIDAÇÃO",
            text_color="#b75b2d",
            font=("Segoe UI", 13, "bold"),
        ).pack(anchor="w", pady=(16, 0))

    def _build_access_section(self, parent) -> None:
        section = self._create_section(parent, 1, "Credenciais do TOTVS")
        grid = ctk.CTkFrame(section, fg_color="transparent")
        grid.pack(fill="x", padx=18, pady=(0, 18))
        grid.grid_columnconfigure((0, 1), weight=1)

        self._form_label(grid, "Usuário").grid(row=0, column=0, sticky="w", padx=(0, 10), pady=(0, 6))
        self._form_label(grid, "Senha").grid(row=0, column=1, sticky="w", padx=(10, 0), pady=(0, 6))
        self._entry(grid, self.username_var).grid(row=1, column=0, sticky="ew", padx=(0, 10))
        self._entry(grid, self.password_var, show="*").grid(row=1, column=1, sticky="ew", padx=(10, 0))

    def _build_plan_section(self, parent) -> None:
        section = self._create_section(parent, 2, "Planilha e Fluxo")
        content = ctk.CTkFrame(section, fg_color="transparent")
        content.pack(fill="x", padx=18, pady=(0, 18))
        content.grid_columnconfigure(0, weight=1)

        self._form_label(content, "Planilha padrão").grid(row=0, column=0, sticky="w", pady=(0, 6))
        self._entry(content, self.file_path_var).grid(row=1, column=0, sticky="ew", padx=(0, 10), pady=(0, 10))
        self._secondary_button(content, "Selecionar", self.select_excel_file, width=150).grid(row=1, column=1, sticky="e", pady=(0, 10))

        self._form_label(content, "Fluxo de lançamento").grid(row=2, column=0, sticky="w", pady=(4, 6))
        self.flow_selector = ctk.CTkSegmentedButton(
            content,
            values=[FLOW_SPECS[key]["label"] for key in FLOW_SPECS],
            fg_color="#f2e7e3",
            selected_color=BUTTON_BG,
            selected_hover_color=BUTTON_ACTIVE_BG,
            unselected_color="#fff9f8",
            unselected_hover_color="#fff0ec",
            text_color="#8e2d2d",
            command=self._on_flow_segment_change,
        )
        self.flow_selector.grid(row=3, column=0, columnspan=2, sticky="ew")
        self.flow_selector.set(FLOW_SPECS[self.flow_var.get()]["label"])

        self.flow_description = ctk.CTkTextbox(
            content,
            height=92,
            corner_radius=14,
            fg_color=SOFT_RED,
            border_width=1,
            border_color=CARD_BORDER,
            text_color="#5a4137",
        )
        self.flow_description.grid(row=4, column=0, columnspan=2, sticky="ew", pady=(12, 10))
        self.flow_description.configure(state="disabled")

        ctk.CTkCheckBox(
            content,
            text="Executar em modo invisível",
            variable=self.headless_var,
            text_color="#303030",
            fg_color=BUTTON_BG,
            hover_color=BUTTON_ACTIVE_BG,
            border_color="#d9c9c3",
        ).grid(row=5, column=0, sticky="w", pady=(0, 6))

    def _build_indicator_section(self, parent) -> None:
        section = self._create_section(parent, 3, "Resumo da Validação")
        cards = ctk.CTkFrame(section, fg_color="transparent")
        cards.pack(fill="x", padx=18, pady=(0, 18))
        for idx in range(5):
            cards.grid_columnconfigure(idx, weight=1)

        self._counter_card(cards, 0, "Linhas na aba", self.total_var)
        self._counter_card(cards, 1, "Pendentes", self.pending_var)
        self._counter_card(cards, 2, "Prontos para lançar", self.ready_var)
        self._counter_card(cards, 3, "Erros de validação", self.error_var)
        self._counter_card(cards, 4, "Aba selecionada", self.sheet_var)

        status_frame = ctk.CTkFrame(section, fg_color=SOFT_RED, corner_radius=18, border_width=1, border_color=CARD_BORDER)
        status_frame.pack(fill="x", padx=18, pady=(0, 18))
        ctk.CTkLabel(status_frame, text="Status", text_color=PRIMARY_TEXT, font=("Segoe UI", 14, "bold")).pack(anchor="w", padx=16, pady=(12, 4))
        ctk.CTkLabel(status_frame, textvariable=self.status_var, text_color="#5c5c5c", font=("Segoe UI", 14)).pack(anchor="w", padx=16, pady=(0, 12))

    def _build_execution_section(self, parent) -> None:
        section = self._create_section(parent, 4, "Execução")
        content = ctk.CTkFrame(section, fg_color="transparent")
        content.pack(fill="both", expand=True, padx=18, pady=(0, 18))
        content.grid_columnconfigure(0, weight=1)
        content.grid_rowconfigure(1, weight=1)

        self.progress = ctk.CTkProgressBar(
            content,
            progress_color=BUTTON_BG,
            fg_color="#f1dfdb",
            corner_radius=20,
            height=18,
        )
        self.progress.set(0)
        self.progress.grid(row=0, column=0, sticky="ew", padx=18, pady=(0, 12))

        self.log_box = ctk.CTkTextbox(
            content,
            corner_radius=14,
            fg_color="#fffdfc",
            border_width=1,
            border_color=CARD_BORDER,
            text_color="#303030",
            height=240,
        )
        self.log_box.grid(row=1, column=0, sticky="nsew", padx=18, pady=(0, 12))

        actions = ctk.CTkFrame(content, fg_color="transparent")
        actions.grid(row=2, column=0, sticky="ew", padx=18)
        for idx in range(4):
            actions.grid_columnconfigure(idx, weight=1)

        self.validate_button = self._primary_button(actions, "Validar planilha", self.validate_workbook)
        self.validate_button.grid(row=0, column=0, padx=(0, 8), sticky="ew")
        self.start_button = self._primary_button(actions, "Iniciar", self.start_processing)
        self.start_button.grid(row=0, column=1, padx=8, sticky="ew")
        self.pause_button = self._secondary_button(actions, "Pausar", self.pause_processing)
        self.pause_button.grid(row=0, column=2, padx=8, sticky="ew")
        self.stop_button = self._secondary_button(actions, "Parar", self.stop_processing)
        self.stop_button.grid(row=0, column=3, padx=(8, 0), sticky="ew")

    def _create_section(self, parent, row: int, title: str) -> ctk.CTkFrame:
        section = ctk.CTkFrame(parent, fg_color=CARD_BG, corner_radius=24, border_width=1, border_color=CARD_BORDER)
        section.grid(row=row, column=0, sticky="ew", pady=(0, 16))
        ctk.CTkLabel(section, text=title, text_color=PRIMARY_TEXT, font=("Segoe UI", 16, "bold")).pack(anchor="w", padx=18, pady=(16, 12))
        return section

    def _entry(self, parent, variable, show: str | None = None) -> ctk.CTkEntry:
        return ctk.CTkEntry(
            parent,
            textvariable=variable,
            height=42,
            corner_radius=14,
            border_color=CARD_BORDER,
            fg_color="#fffdfc",
            text_color="#303030",
            placeholder_text="",
            show=show,
        )

    def _form_label(self, parent, text: str) -> ctk.CTkLabel:
        return ctk.CTkLabel(parent, text=text, text_color="#1f1f1f", font=("Segoe UI", 14, "bold"))

    def _primary_button(self, parent, text: str, command):
        return ctk.CTkButton(
            parent,
            text=text,
            command=command,
            height=42,
            corner_radius=16,
            fg_color=BUTTON_BG,
            hover_color=BUTTON_ACTIVE_BG,
            text_color="#ffffff",
            font=("Segoe UI", 15, "bold"),
        )

    def _secondary_button(self, parent, text: str, command, width: int | None = None):
        kwargs = {
            "text": text,
            "command": command,
            "height": 42,
            "corner_radius": 16,
            "fg_color": "#fff9f8",
            "hover_color": "#fff0ec",
            "border_width": 1,
            "border_color": CARD_BORDER,
            "text_color": PRIMARY_TEXT,
            "font": ("Segoe UI", 14, "bold"),
        }
        if width is not None:
            kwargs["width"] = width
        return ctk.CTkButton(parent, **kwargs)

    def _counter_card(self, parent, column: int, title: str, value_var: ctk.StringVar) -> None:
        card = ctk.CTkFrame(parent, fg_color="#fffdfc", corner_radius=18, border_width=1, border_color=CARD_BORDER)
        card.grid(row=0, column=column, padx=6, sticky="ew")
        ctk.CTkLabel(card, text=title, text_color="#5a4137", font=("Segoe UI", 13, "bold")).pack(anchor="w", padx=14, pady=(14, 4))
        ctk.CTkLabel(card, textvariable=value_var, text_color=PRIMARY_TEXT, font=("Segoe UI", 18, "bold")).pack(anchor="w", padx=14, pady=(0, 14))

    def _on_flow_segment_change(self, label: str) -> None:
        for key, spec in FLOW_SPECS.items():
            if spec["label"] == label:
                self.flow_var.set(key)
                break
        self._update_flow_description()
        self.validation_result = None
        self.status_var.set("Fluxo alterado. Valide a planilha novamente.")
        self.progress.set(0)
        self._update_action_buttons()

    def _update_flow_description(self) -> None:
        flow = FLOW_SPECS[self.flow_var.get()]
        description = [
            f"Fluxo selecionado: {flow['label']}",
            f"Aba esperada: {flow['sheet_name']}",
            flow["group_hint"],
            f"Regra de pendência: coluna {flow['pending_column']} em branco.",
        ]
        self.flow_description.configure(state="normal")
        self.flow_description.delete("1.0", "end")
        self.flow_description.insert("1.0", "\n".join(description))
        self.flow_description.configure(state="disabled")

    def log(self, message: str) -> None:
        timestamp = pd.Timestamp.now().strftime("%H:%M:%S")
        self.log_box.insert("end", f"[{timestamp}] {message}\n")
        self.log_box.see("end")
        self.update_idletasks()

    def select_excel_file(self) -> None:
        file_path = filedialog.askopenfilename(
            title="Selecionar planilha de caixa",
            filetypes=[("Excel", "*.xlsx *.xlsm *.xls")],
        )
        if not file_path:
            return
        self.file_path_var.set(file_path)
        self.validation_result = None
        self.status_var.set("Planilha selecionada. Pronta para validação.")
        self.log(f"Planilha selecionada: {file_path}")
        self.progress.set(0)
        self._update_action_buttons()

    def validate_workbook(self) -> None:
        path_text = self.file_path_var.get().strip()
        if not path_text:
            messagebox.showwarning("Planilha", "Selecione a planilha antes de validar.")
            return

        workbook_path = Path(path_text)
        if not workbook_path.exists():
            messagebox.showerror("Arquivo não encontrado", f"A planilha não foi encontrada:\n{workbook_path}")
            return

        flow = FLOW_SPECS[self.flow_var.get()]
        self.log(f"Iniciando validação da aba {flow['sheet_name']}...")
        self.progress.set(0.15)

        try:
            excel_file = pd.ExcelFile(workbook_path)
            workbook_sheets = list(excel_file.sheet_names)
            actual_sheet_name = _resolve_sheet_name(workbook_sheets, flow["sheet_name"])
            if not actual_sheet_name:
                raise RuntimeError(f"A aba {flow['sheet_name']} não existe nesta planilha.")

            dataframe = pd.read_excel(
                workbook_path,
                sheet_name=actual_sheet_name,
                header=flow.get("header_row", 0),
            )
            renamed, missing_columns = self._normalize_flow_dataframe(dataframe, flow)

            total_rows = len(renamed.index)
            pending_column = flow["pending_column"]
            pending_rows = 0
            launched_rows = 0
            if pending_column in renamed.columns:
                status_series = renamed[pending_column].fillna("").astype(str).str.strip()
                pending_rows = int(status_series.eq("").sum())
                launched_rows = int((~status_series.eq("")).sum())

            self.validation_result = ValidationResult(
                flow_key=self.flow_var.get(),
                flow_label=flow["label"],
                sheet_name=actual_sheet_name,
                total_rows=total_rows,
                pending_rows=pending_rows,
                launched_rows=launched_rows,
                missing_columns=missing_columns,
                workbook_sheets=workbook_sheets,
            )

            self.total_var.set(str(total_rows))
            self.pending_var.set(str(pending_rows))
            self.ready_var.set(str(pending_rows if not missing_columns else 0))
            self.error_var.set(str(len(missing_columns)))
            self.sheet_var.set(actual_sheet_name)

            if missing_columns:
                self.status_var.set("Validação concluída com pendências de estrutura")
                self.log(f"Colunas obrigatórias ausentes: {', '.join(missing_columns)}")
                messagebox.showwarning(
                    "Validação com pendências",
                    "A planilha foi lida, mas faltam colunas obrigatórias:\n\n- " + "\n- ".join(missing_columns),
                )
            else:
                self.status_var.set(f"Validação concluída. {pending_rows} lançamento(s) pendente(s).")
                self.log(f"Aba {actual_sheet_name} validada com sucesso.")
                self.log(f"Linhas totais: {total_rows} | Pendentes: {pending_rows} | Já lançadas: {launched_rows}")

            self.progress.set(1)
        except Exception as exc:
            self.validation_result = None
            self.total_var.set("0")
            self.pending_var.set("0")
            self.ready_var.set("0")
            self.error_var.set("1")
            self.sheet_var.set("-")
            self.status_var.set("Falha na validação")
            self.progress.set(0)
            self.log(f"Erro de validação: {exc}")
            messagebox.showerror("Falha na validação", str(exc))
        finally:
            self._update_action_buttons()

    def _normalize_flow_dataframe(self, dataframe: pd.DataFrame, flow: Dict) -> tuple[pd.DataFrame, list[str]]:
        rename_map: Dict[str, str] = {}
        normalized_columns = {_normalize_text(column): column for column in dataframe.columns}

        for canonical, aliases in flow["aliases"].items():
            for alias in aliases:
                alias_norm = _normalize_text(alias)
                if alias_norm in normalized_columns:
                    rename_map[normalized_columns[alias_norm]] = canonical
                    break

        renamed = dataframe.rename(columns=rename_map).copy()
        missing_columns = [column for column in flow["required_columns"] if column not in renamed.columns]
        return renamed, missing_columns

    def start_processing(self) -> None:
        if self.validation_result is None:
            messagebox.showwarning("Validação", "Valide a planilha antes de iniciar.")
            return
        if self.validation_result.missing_columns:
            messagebox.showwarning("Validação", "Corrija as colunas obrigatórias antes de iniciar.")
            return
        messagebox.showinfo(
            "Escopo 1",
            "A interface e a validação já estão prontas.\n\n"
            "A automação operacional do fluxo selecionado será implementada no próximo passo.",
        )
        self.log(f"Início bloqueado: automação do fluxo {self.validation_result.flow_label} ainda não foi implementada neste escopo.")

    def pause_processing(self) -> None:
        messagebox.showinfo("Escopo 1", "O controle de pausa será habilitado junto com a automação operacional.")

    def stop_processing(self) -> None:
        messagebox.showinfo("Escopo 1", "O controle de parada será habilitado junto com a automação operacional.")

    def _update_action_buttons(self) -> None:
        validated = self.validation_result is not None and not self.validation_result.missing_columns
        self.start_button.configure(state="normal" if validated else "disabled")
        self.pause_button.configure(state="disabled")
        self.stop_button.configure(state="disabled")


if __name__ == "__main__":
    app = RoboLancamentosCaixaApp()
    app.mainloop()
