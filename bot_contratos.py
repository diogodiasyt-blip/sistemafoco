import base64
import binascii
import json
import os
import queue
import re
import sys
import threading
import time
from datetime import datetime
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox
from urllib.error import HTTPError, URLError
from urllib.parse import urlencode, urlparse
from urllib.request import Request, urlopen

import customtkinter as ctk
import multiprocessing
import pandas as pd
import winsound
try:
    import keyring
except Exception:
    keyring = None
from PIL import Image


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
CORAL_API_BASE_URL = os.environ.get("CORAL_API_BASE_URL", "https://servicescoral.aluguefoco.com.br").rstrip("/")
CORAL_API_LOGIN_URL = f"{CORAL_API_BASE_URL}/api/auth/login"
CORAL_API_CONTRACT_PDF_URL = f"{CORAL_API_BASE_URL}/api/pdf/management/voucher/pdf"
CORAL_API_HTTP_TIMEOUT_SECONDS = int(os.environ.get("CORAL_API_HTTP_TIMEOUT_SECONDS", "60"))
APP_CREDENTIAL_SERVICE = "SistemaFOCO"
CREDENTIAL_MODULE_KEY = "contratos_coral"


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


class CoralApiUnauthorizedError(RuntimeError):
    """Token da API do Coral expirado ou invalido."""


def _api_error_detail(raw):
    try:
        payload = json.loads(raw.decode("utf-8", errors="replace")) if raw else {}
    except Exception:
        payload = {}
    if isinstance(payload, dict):
        return str(payload.get("message") or payload.get("error") or "").strip()
    return ""


def login_coral_api(usuario, senha, opener=urlopen):
    usuario = str(usuario or "").strip()
    senha = str(senha or "")
    if not usuario or not senha:
        raise RuntimeError("Informe usuario e senha do Coral.")

    body = json.dumps({"login": usuario, "password": senha}, separators=(",", ":")).encode("utf-8")
    request = Request(
        CORAL_API_LOGIN_URL,
        data=body,
        headers={
            "Accept": "application/json, text/plain, */*",
            "Content-Type": "application/json",
            "User-Agent": "SistemaFOCO-RoboContratos/1.0",
        },
        method="POST",
    )
    try:
        with opener(request, timeout=CORAL_API_HTTP_TIMEOUT_SECONDS) as response:
            raw = response.read()
    except HTTPError as exc:
        detail = _api_error_detail(exc.read())
        raise RuntimeError(f"Login Coral API HTTP {exc.code}: {detail or 'acesso recusado'}") from exc
    except URLError as exc:
        raise RuntimeError(f"Falha de comunicacao no login da API do Coral: {exc.reason}") from exc
    except TimeoutError as exc:
        raise RuntimeError("Timeout no login da API do Coral.") from exc

    try:
        payload = json.loads(raw.decode("utf-8"))
    except Exception as exc:
        raise RuntimeError("Login da API do Coral retornou resposta invalida.") from exc
    data = payload.get("data") if isinstance(payload, dict) else None
    token = str(data.get("token") or "").strip() if isinstance(data, dict) else ""
    if not token:
        raise RuntimeError("Login da API do Coral nao retornou data.token.")
    return token


def _decode_pdf_candidate(value):
    if isinstance(value, bytes):
        return value if value.lstrip().startswith(b"%PDF-") else None
    if isinstance(value, list) and all(isinstance(item, int) and 0 <= item <= 255 for item in value):
        decoded = bytes(value)
        return decoded if decoded.lstrip().startswith(b"%PDF-") else None
    if isinstance(value, str):
        text = value.strip()
        if text.lower().startswith(("http://", "https://")):
            return text
        if text.lower().startswith("data:application/pdf;base64,"):
            text = text.split(",", 1)[1]
        try:
            decoded = base64.b64decode(text, validate=False)
        except (ValueError, binascii.Error):
            return None
        return decoded if decoded.lstrip().startswith(b"%PDF-") else None
    if isinstance(value, dict):
        for key in ("pdf", "base64", "content", "file", "body", "url", "downloadUrl", "data"):
            if key not in value:
                continue
            decoded = _decode_pdf_candidate(value[key])
            if decoded is not None:
                return decoded
    return None


def _extract_pdf_response(raw, content_type=""):
    if raw.lstrip().startswith(b"%PDF-"):
        return raw
    try:
        payload = json.loads(raw.decode("utf-8"))
    except Exception as exc:
        raise RuntimeError(
            f"Endpoint de PDF retornou conteudo invalido ({content_type or 'tipo nao informado'})."
        ) from exc
    decoded = _decode_pdf_candidate(payload)
    if decoded is None:
        raise RuntimeError("Endpoint de PDF nao retornou arquivo, base64 ou URL reconhecivel.")
    return decoded


def _download_pdf_request(token, contrato, opener=urlopen):
    query = urlencode(
        {
            "reservationId": contrato,
            "language": "PORTUGUESE",
            "isRentAgreement": "true",
        }
    )
    request = Request(
        f"{CORAL_API_CONTRACT_PDF_URL}?{query}",
        headers={
            "Accept": "application/json, text/plain, */*",
            "Authorization": f"Bearer {token}",
            "Origin": "https://coral.aluguefoco.com.br",
            "Referer": "https://coral.aluguefoco.com.br/",
            "User-Agent": "SistemaFOCO-RoboContratos/1.0",
        },
        method="GET",
    )
    try:
        with opener(request, timeout=CORAL_API_HTTP_TIMEOUT_SECONDS) as response:
            raw = response.read()
            content_type = str(response.headers.get("Content-Type") or "")
    except HTTPError as exc:
        raw = exc.read()
        if exc.code == 401:
            raise CoralApiUnauthorizedError("Token da API do Coral expirou.") from exc
        detail = _api_error_detail(raw)
        raise RuntimeError(f"PDF Coral API HTTP {exc.code}: {detail or 'download recusado'}") from exc
    except URLError as exc:
        raise RuntimeError(f"Falha de comunicacao ao baixar PDF: {exc.reason}") from exc
    except TimeoutError as exc:
        raise RuntimeError("Timeout ao baixar PDF do Coral.") from exc
    return _extract_pdf_response(raw, content_type)


def baixar_contrato_pdf_api(token, contrato, pasta_download, opener=urlopen):
    contrato = str(contrato or "").strip()
    if not contrato:
        raise ValueError("Numero do contrato vazio.")
    pdf_result = _download_pdf_request(token, contrato, opener=opener)
    if isinstance(pdf_result, str):
        headers = {"User-Agent": "SistemaFOCO-RoboContratos/1.0"}
        if urlparse(pdf_result).netloc.casefold() == urlparse(CORAL_API_BASE_URL).netloc.casefold():
            headers["Authorization"] = f"Bearer {token}"
        request = Request(
            pdf_result,
            headers=headers,
            method="GET",
        )
        try:
            with opener(request, timeout=CORAL_API_HTTP_TIMEOUT_SECONDS) as response:
                pdf_bytes = _extract_pdf_response(
                    response.read(),
                    str(response.headers.get("Content-Type") or ""),
                )
        except HTTPError as exc:
            if exc.code == 401:
                raise CoralApiUnauthorizedError("Token expirou ao abrir a URL do PDF.") from exc
            raise RuntimeError(f"URL do PDF retornou HTTP {exc.code}.") from exc
        except URLError as exc:
            raise RuntimeError(f"Falha ao abrir a URL retornada para o PDF: {exc.reason}") from exc
    else:
        pdf_bytes = pdf_result

    if not isinstance(pdf_bytes, bytes) or not pdf_bytes.lstrip().startswith(b"%PDF-"):
        raise RuntimeError("Arquivo recebido nao possui assinatura valida de PDF.")
    output_dir = Path(pasta_download)
    output_dir.mkdir(parents=True, exist_ok=True)
    safe_contract = re.sub(r"[^A-Za-z0-9._-]+", "_", contrato).strip("._") or "contrato"
    filename = f"{safe_contract}-voucher.pdf"
    output_path = output_dir / filename
    temp_path = output_dir / f".{filename}.part"
    try:
        temp_path.write_bytes(pdf_bytes)
        os.replace(temp_path, output_path)
    finally:
        try:
            temp_path.unlink(missing_ok=True)
        except OSError:
            pass
    return output_path


def atualizar_relatorio_contratos(relatorio_path, registros, contrato, status, log_callback):
    registros.append({"contrato": contrato, "status": status})
    try:
        pd.DataFrame(registros, columns=["contrato", "status"]).to_excel(relatorio_path, index=False)
    except Exception as exc:
        log_callback(f"Nao foi possivel atualizar o relatorio: {exc}")


def montar_mensagem_resumo_execucao(success, failed, tempo, relatorio_path=None):
    linhas = [
        f"Sucessos: {len(success)}",
        f"Erros: {len(failed)}",
        f"Tempo total: {tempo}",
    ]
    if relatorio_path:
        linhas.extend(["", f"Relatorio completo: {relatorio_path}"])

    if failed:
        linhas.extend(["", "Contratos com erro:", *failed])
    else:
        linhas.extend(["", "Nenhum contrato com erro."])

    return "\n".join(linhas)


def executar_robo(usuario, senha, planilha_path, pasta_download, log_callback, queue_result, progress_callback):
    start_time = time.time()
    success = []
    failed = []
    relatorio_registros = []
    relatorio_path = os.path.join(
        os.path.dirname(os.path.abspath(planilha_path)),
        f"Relatorio_Contratos_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
    )
    log_callback("Validando pasta de salvamento...")
    try:
        os.makedirs(pasta_download, exist_ok=True)
        log_callback(f"Pasta pronta: {pasta_download}")
    except Exception as e:
        log_callback(f"Erro ao criar pasta: {str(e)}")
        return

    try:
        log_callback("Processamento direto pela API do Coral; o Chrome nao sera aberto.")
        log_callback("Autenticando na API do Coral...")
        token = login_coral_api(usuario, senha)
        log_callback("Autenticacao da API confirmada.")

        df = pd.read_excel(planilha_path)
        if "contrato" not in df.columns:
            raise ValueError("A planilha precisa conter a coluna 'contrato'.")
        contratos = list(dict.fromkeys(df["contrato"].dropna().astype(str).str.strip().tolist()))
        contratos = [contrato for contrato in contratos if contrato]
        total = len(contratos)
        log_callback(f"Iniciando processamento de {total} contratos...")
        log_callback(f"Relatorio em tempo real: {relatorio_path}")

        for i, contrato in enumerate(contratos, 1):
            log_callback(f"\n[{i}/{total}] Processando: {contrato}")
            try:
                try:
                    pdf_path = baixar_contrato_pdf_api(token, contrato, pasta_download)
                except CoralApiUnauthorizedError:
                    log_callback("Token expirado; renovando autenticacao antes de repetir o GET do PDF.")
                    token = login_coral_api(usuario, senha)
                    pdf_path = baixar_contrato_pdf_api(token, contrato, pasta_download)
                success.append(contrato)
                log_callback(f"{contrato} -> Download concluído: {pdf_path.name}")
                atualizar_relatorio_contratos(
                    relatorio_path,
                    relatorio_registros,
                    contrato,
                    "BAIXADO",
                    log_callback,
                )
            except Exception as e:
                erro = str(e)
                log_callback(f"Erro em {contrato}: {erro}")
                failed.append(contrato)
                atualizar_relatorio_contratos(
                    relatorio_path,
                    relatorio_registros,
                    contrato,
                    f"ERRO - {erro}",
                    log_callback,
                )
            finally:
                progress_callback((i / total) * 100 if total else 100)

    except Exception as e:
        log_callback(f"ERRO FATAL: {str(e)}")

    elapsed = time.time() - start_time
    minutes = int(elapsed // 60)
    seconds = int(elapsed % 60)

    queue_result.put(
        {
            "success": success,
            "failed": failed,
            "time": f"{minutes}min {seconds}s",
            "relatorio_path": relatorio_path,
        }
    )
    log_callback(f"\nFINALIZADO em {minutes}min {seconds}s -> Sucessos: {len(success)} | Erros: {len(failed)}")

    try:
        winsound.Beep(1000, 800)
        time.sleep(0.3)
        winsound.Beep(1200, 600)
    except Exception:
        pass


class RoboContratosApp:
    def __init__(self, root=None):
        ctk.set_appearance_mode("light")
        ctk.set_default_color_theme("blue")

        default_root = getattr(tk, "_default_root", None)
        self.owns_root = root is None and default_root is None
        if root is not None:
            self.root = root
            try:
                self.root.configure(bg=MAIN_BG)
            except Exception:
                pass
        elif default_root is not None:
            self.root = ctk.CTkToplevel(default_root)
            self.root.title("Robo de Contratos Coral - Desenvolvido por Diogo Medeiros ?? 2026")
            self.root.geometry("1120x820")
            self.root.minsize(980, 720)
            self.root.configure(fg_color=MAIN_BG)
        else:
            self.root = ctk.CTk()
            self.root.title("Robo de Contratos Coral - Desenvolvido por Diogo Medeiros ?? 2026")
            self.root.geometry("1120x820")
            self.root.minsize(980, 720)
            self.root.configure(fg_color=MAIN_BG)

        desktop = get_desktop_path()
        pasta_padrao = os.path.join(desktop, "Contratos_Foco")

        self.usuario_var = tk.StringVar(value="")
        self.senha_var = tk.StringVar(value="")
        self.planilha_var = tk.StringVar(value="")
        self.pasta_var = tk.StringVar(value=pasta_padrao)

        self.log_queue = queue.Queue()
        self.result_queue = queue.Queue()
        self.logo_image = None
        self.logo_label = None

        self.carregar_credenciais_salvas()
        self.create_widgets()

    def chave_credencial(self, sufixo):
        return f"{CREDENTIAL_MODULE_KEY}:{sufixo}"

    def carregar_credenciais_salvas(self):
        if keyring is None:
            return
        try:
            usuario = keyring.get_password(APP_CREDENTIAL_SERVICE, self.chave_credencial("usuario")) or ""
            senha = keyring.get_password(APP_CREDENTIAL_SERVICE, self.chave_credencial("senha")) or ""
            if usuario:
                self.usuario_var.set(usuario)
            if senha:
                self.senha_var.set(senha)
        except Exception:
            pass

    def salvar_credenciais(self):
        if keyring is None:
            messagebox.showwarning("Salvar acesso", "Biblioteca keyring nao esta disponivel neste ambiente.")
            return
        usuario = self.usuario_var.get().strip()
        senha = self.senha_var.get().strip()
        if not usuario or not senha:
            messagebox.showwarning("Salvar acesso", "Preencha usuario e senha antes de salvar.")
            return
        try:
            keyring.set_password(APP_CREDENTIAL_SERVICE, self.chave_credencial("usuario"), usuario)
            keyring.set_password(APP_CREDENTIAL_SERVICE, self.chave_credencial("senha"), senha)
            messagebox.showinfo("Salvar acesso", "Usuario e senha salvos neste computador.")
        except Exception as exc:
            messagebox.showerror("Salvar acesso", f"Nao foi possivel salvar o acesso:\n{exc}")

    def limpar_credenciais(self):
        if keyring is not None:
            for sufixo in ("usuario", "senha"):
                try:
                    keyring.delete_password(APP_CREDENTIAL_SERVICE, self.chave_credencial(sufixo))
                except Exception:
                    pass
        self.usuario_var.set("")
        self.senha_var.set("")
        messagebox.showinfo("Limpar acesso", "Acesso removido deste computador.")

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
            text="Download de contratos.",
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
        credenciais_box = ctk.CTkFrame(acesso_grid, fg_color="transparent")
        credenciais_box.grid(row=2, column=0, columnspan=2, sticky="w", pady=(12, 0))
        ctk.CTkButton(
            credenciais_box,
            text="Salvar acesso",
            command=self.salvar_credenciais,
            height=36,
            width=145,
            corner_radius=12,
            fg_color="#ffffff",
            text_color=PRIMARY_TEXT,
            hover_color=SOFT_RED,
            border_width=1,
            border_color="#f0d7d2",
            font=("Segoe UI", 13, "bold"),
        ).pack(side="left", padx=(0, 10))
        ctk.CTkButton(
            credenciais_box,
            text="Limpar acesso",
            command=self.limpar_credenciais,
            height=36,
            width=145,
            corner_radius=12,
            fg_color="#ffffff",
            text_color=PRIMARY_TEXT,
            hover_color=SOFT_RED,
            border_width=1,
            border_color="#f0d7d2",
            font=("Segoe UI", 13, "bold"),
        ).pack(side="left")

        config = self.criar_secao(scroll, "Configuracoes")
        api_status = ctk.CTkFrame(
            config,
            fg_color=SOFT_RED,
            corner_radius=12,
        )
        api_status.pack(fill="x", padx=18, pady=(0, 18))
        ctk.CTkLabel(
            api_status,
            text="Api configurada",
            font=("Segoe UI", 13),
            text_color="#303030",
        ).pack(anchor="w", padx=14, pady=12)

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
            text="Desenvolvido por Diogo Medeiros 2026",
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
        relatorio_path = result.get("relatorio_path")

        messagebox.showinfo(
            "Execucao Finalizada",
            montar_mensagem_resumo_execucao(s, f, tempo, relatorio_path),
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
