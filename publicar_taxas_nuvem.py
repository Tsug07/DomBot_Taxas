"""
DomBot - Publicar Taxas via Nuvem
==================================
Automação RPA para publicar PDFs de Taxa Panificação e Taxa Motorista
no Domínio Folha via "Publicação de Documentos Externos" (GMS Nuvem).

Selecione a pasta que contém os PDFs gerados. Os arquivos devem seguir
o padrão de nome: {codigo} - {nome da empresa} - {MM.YYYY}.pdf
Ex: 314 - PADARIA EXEMPLO LTDA - 05.2026.pdf

O código extraído do nome do arquivo é usado como Nº GMS na publicação.

Autor: Hugo L. Almeida
Versão: 2.0
"""

import customtkinter as ctk
from tkinter import filedialog, messagebox
import threading
import os
import re
import time
import traceback
import logging
from datetime import datetime
from typing import Optional

from PIL import Image, ImageDraw
from pywinauto.application import Application
from pywinauto.keyboard import send_keys
from pywinauto import findwindows, timings
from pywinauto.findwindows import ElementNotFoundError
import win32gui
import win32con
import ctypes

try:
    from dotenv import load_dotenv
    load_dotenv()
    import requests
    _requests_available = True
except ImportError:
    _requests_available = False


# ─────────────────────────── Parsing de nome de arquivo ─────────────────────

def parsear_pdf(nome_arquivo: str) -> Optional[dict]:
    """
    Extrai código, nome e período de um arquivo PDF com o padrão:
      {codigo} - {nome empresa} - {MM.YYYY}.pdf
    Retorna dict com 'codigo', 'nome', 'periodo' ou None se não bater.
    """
    base = os.path.splitext(nome_arquivo)[0]
    # Padrão: número - qualquer coisa - MM.YYYY / MM-YYYY / MM/YYYY
    m = re.match(r'^(\d+)\s*-\s*(.+?)\s*-\s*(\d{2}[.\-/]\d{4})$', base.strip())
    if not m:
        return None
    return {
        'codigo':  m.group(1).strip(),
        'nome':    m.group(2).strip(),
        'periodo': m.group(3).strip(),
    }


def listar_pdfs_da_pasta(pasta: str) -> list[dict]:
    """
    Lista todos os PDFs na pasta que seguem o padrão de nome esperado.
    Retorna lista de dicts: {codigo, nome, periodo, caminho}
    """
    arquivos = []
    try:
        for nome in sorted(os.listdir(pasta)):
            if not nome.lower().endswith('.pdf'):
                continue
            info = parsear_pdf(nome)
            if info:
                info['caminho'] = os.path.join(pasta, nome)
                info['arquivo'] = nome
                arquivos.append(info)
    except Exception:
        pass
    return arquivos


# ─────────────────────────────── Logging Handler ────────────────────────────


class GUILogHandler(logging.Handler):
    def __init__(self, gui):
        super().__init__()
        self.gui = gui

    def emit(self, record):
        msg = self.format(record)
        level = record.levelno
        self.gui.window.after(0, lambda: self.gui.adicionar_log(msg, level))


# ──────────────────────────────────── GUI ───────────────────────────────────


class PublicarTaxasGUI:
    CORES = {
        'sucesso':      '#2ECC71',
        'erro':         '#E74C3C',
        'aviso':        '#F39C12',
        'info':         '#3498DB',
        'texto':        '#ECF0F1',
        'fundo_card':   '#2C3E50',
        'fundo_escuro': '#1A252F',
        'destaque':     '#1ABC9C',
        'processando':  '#9B59B6',
    }

    def __init__(self):
        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("green")

        self.window = ctk.CTk()
        self.window.title("DomBot - Publicar Taxas via Nuvem v2.0")
        self.window.geometry("900x580")
        self.window.minsize(750, 480)
        self.window.protocol("WM_DELETE_WINDOW", self.ao_fechar)

        self.executando = False
        self.pausa_solicitada = False
        self.thread_automacao = None

        self.stats = {'processados': 0, 'sucesso': 0, 'erros': 0, 'tempo_inicio': None}

        self.logs_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "logs")
        os.makedirs(self.logs_dir, exist_ok=True)
        self._setup_file_logging()
        self._set_window_icon()

        self.pasta_var = ctk.StringVar()
        self.tipo_taxa_var = ctk.StringVar(value="Panificação")
        self.status_var = ctk.StringVar(value="Aguardando início...")

        self._pdfs_carregados: list[dict] = []

        self.logger = logging.getLogger('PublicarTaxasNuvem')
        self.logger.setLevel(logging.INFO)
        self.logger.handlers = []
        handler = GUILogHandler(self)
        handler.setFormatter(logging.Formatter('%(message)s'))
        self.logger.addHandler(handler)

        self._criar_interface()

    # ── Setup ─────────────────────────────────────────────────────────────────

    def _setup_file_logging(self):
        data = datetime.now().strftime("%Y-%m-%d")

        self.success_logger = logging.getLogger('PubSuccess')
        self.success_logger.setLevel(logging.INFO)
        if not self.success_logger.handlers:
            h = logging.FileHandler(
                os.path.join(self.logs_dir, f'pub_success_{data}.log'), encoding='utf-8')
            h.setFormatter(logging.Formatter('%(asctime)s - %(message)s', '%Y-%m-%d %H:%M:%S'))
            self.success_logger.addHandler(h)

        self.error_logger = logging.getLogger('PubError')
        self.error_logger.setLevel(logging.ERROR)
        if not self.error_logger.handlers:
            h = logging.FileHandler(
                os.path.join(self.logs_dir, f'pub_error_{data}.log'), encoding='utf-8')
            h.setFormatter(logging.Formatter('%(asctime)s - %(message)s', '%Y-%m-%d %H:%M:%S'))
            self.error_logger.addHandler(h)

    def _set_window_icon(self):
        try:
            icon = os.path.join(os.path.dirname(os.path.abspath(__file__)), "assets", "favicon.ico")
            if os.path.exists(icon):
                self.window.iconbitmap(icon)
        except Exception:
            pass

    # ── Interface ─────────────────────────────────────────────────────────────

    def _criar_interface(self):
        self.window.grid_columnconfigure(0, weight=1)
        self.window.grid_rowconfigure(0, weight=1)

        main = ctk.CTkFrame(self.window, fg_color="transparent")
        main.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)
        main.grid_columnconfigure(0, weight=1)
        main.grid_rowconfigure(3, weight=1)

        self._criar_header(main)
        self._criar_config(main)
        self._criar_stats(main)
        self._criar_conteudo(main)

    def _criar_header(self, parent):
        hdr = ctk.CTkFrame(parent, fg_color=self.CORES['fundo_card'], corner_radius=8)
        hdr.grid(row=0, column=0, sticky="ew", pady=(0, 6))
        hdr.grid_columnconfigure(1, weight=1)

        logo_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "assets", "DomBot_New.png")
        if os.path.exists(logo_path):
            sz, csz = 66, 44
            bg = Image.new("RGBA", (sz, sz), (0, 0, 0, 0))
            mask = Image.new("L", (csz, csz), 0)
            ImageDraw.Draw(mask).ellipse((0, 0, csz - 1, csz - 1), fill=255)
            circle = Image.new("RGBA", (csz, csz), (255, 255, 255, 255))
            off = (sz - csz) // 2
            bg.paste(circle, (off, off), mask)
            orig = Image.open(logo_path).convert("RGBA").resize((sz, sz), Image.LANCZOS)
            bg.paste(orig, (0, 0), orig)
            logo_img = ctk.CTkImage(light_image=bg, dark_image=bg, size=(sz, sz))
            ctk.CTkLabel(hdr, image=logo_img, text="").grid(row=0, column=0, padx=10, pady=8)
        else:
            f = ctk.CTkFrame(hdr, fg_color=self.CORES['destaque'], width=44, height=44, corner_radius=22)
            f.grid(row=0, column=0, padx=10, pady=8)
            f.grid_propagate(False)
            ctk.CTkLabel(f, text="☁️", font=("Segoe UI Emoji", 18)).place(relx=0.5, rely=0.5, anchor="center")

        ctk.CTkLabel(
            hdr, text="DomBot - Publicar Taxas via Nuvem",
            font=ctk.CTkFont(size=16, weight="bold"),
            text_color=self.CORES['texto'],
        ).grid(row=0, column=1, sticky="w", padx=5)

        sf = ctk.CTkFrame(hdr, fg_color="transparent")
        sf.grid(row=0, column=2, padx=10)
        self.status_indicator = ctk.CTkFrame(sf, fg_color="#7F8C8D", width=10, height=10, corner_radius=5)
        self.status_indicator.pack(side="left", padx=(0, 6))
        ctk.CTkLabel(sf, textvariable=self.status_var, font=ctk.CTkFont(size=11),
                     text_color="#95A5A6").pack(side="left")

    def _criar_config(self, parent):
        cfg = ctk.CTkFrame(parent, fg_color=self.CORES['fundo_card'], corner_radius=8)
        cfg.grid(row=1, column=0, sticky="ew", pady=(0, 6))
        cfg.grid_columnconfigure(0, weight=1)

        # Linha 1 — seleção de pasta + tipo taxa
        r1 = ctk.CTkFrame(cfg, fg_color="transparent")
        r1.grid(row=0, column=0, sticky="ew", padx=10, pady=(8, 4))
        r1.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(r1, text="📁", font=ctk.CTkFont(size=14)).grid(row=0, column=0, padx=(0, 5))

        self.entry_pasta = ctk.CTkEntry(
            r1, textvariable=self.pasta_var,
            placeholder_text="Selecione a pasta com os PDFs gerados...",
            height=32, font=ctk.CTkFont(size=11),
        )
        self.entry_pasta.grid(row=0, column=1, sticky="ew", padx=(0, 8))

        ctk.CTkButton(
            r1, text="Procurar", command=self._selecionar_pasta,
            width=80, height=32, font=ctk.CTkFont(size=11),
        ).grid(row=0, column=2, padx=(0, 10))

        ctk.CTkLabel(r1, text="Taxa:", font=ctk.CTkFont(size=11)).grid(row=0, column=3, padx=(0, 4))
        self.tipo_menu = ctk.CTkOptionMenu(
            r1, values=["Panificação", "Motorista"],
            variable=self.tipo_taxa_var,
            width=120, height=32, font=ctk.CTkFont(size=11),
            command=self._on_tipo_changed,
        )
        self.tipo_menu.grid(row=0, column=4)

        # Linha 2 — botões controle
        r2 = ctk.CTkFrame(cfg, fg_color="transparent")
        r2.grid(row=1, column=0, sticky="ew", padx=10, pady=(4, 8))

        self.btn_iniciar = ctk.CTkButton(
            r2, text="▶ Publicar", command=self._iniciar,
            width=110, height=32, font=ctk.CTkFont(size=12, weight="bold"),
            fg_color="#27AE60", hover_color="#1E8449",
        )
        self.btn_iniciar.pack(side="left", padx=(0, 6))

        self.btn_pausar = ctk.CTkButton(
            r2, text="⏸ Pausar", command=self._pausar,
            width=100, height=32, font=ctk.CTkFont(size=11),
            fg_color="#F39C12", hover_color="#D68910", state="disabled",
        )
        self.btn_pausar.pack(side="left", padx=(0, 6))

        self.btn_parar = ctk.CTkButton(
            r2, text="■ Parar", command=self._parar,
            width=90, height=32, font=ctk.CTkFont(size=11),
            fg_color="#E74C3C", hover_color="#C0392B", state="disabled",
        )
        self.btn_parar.pack(side="left", padx=(0, 12))

        ctk.CTkButton(
            r2, text="🔍 Pré-visualizar PDFs", command=self._preview_pasta,
            width=160, height=32, font=ctk.CTkFont(size=11),
        ).pack(side="left")

        self.lbl_tipo_ativo = ctk.CTkLabel(
            r2, text="● Taxa Panificação",
            font=ctk.CTkFont(size=11), text_color=self.CORES['destaque'],
        )
        self.lbl_tipo_ativo.pack(side="right", padx=10)

    def _criar_stats(self, parent):
        sf = ctk.CTkFrame(parent, fg_color=self.CORES['fundo_card'], corner_radius=8)
        sf.grid(row=2, column=0, sticky="ew", pady=(0, 6))
        for i in range(4):
            sf.grid_columnconfigure(i, weight=1)

        for col, (titulo, cor_key, attr) in enumerate([
            ("Publicados",     'sucesso',     'lbl_sucesso'),
            ("Erros",          'erro',        'lbl_erros'),
            ("Empresa atual",  'info',        'lbl_empresa'),
            ("Tempo decorrido",'aviso',       'lbl_tempo'),
        ]):
            card = ctk.CTkFrame(sf, fg_color=self.CORES['fundo_escuro'], corner_radius=6)
            card.grid(row=0, column=col, sticky="ew", padx=6, pady=6)
            ctk.CTkLabel(card, text=titulo, font=ctk.CTkFont(size=10), text_color="#95A5A6").pack(pady=(6, 0))
            lbl = ctk.CTkLabel(card, text="0" if col < 2 else "—",
                               font=ctk.CTkFont(size=18, weight="bold"),
                               text_color=self.CORES[cor_key])
            lbl.pack(pady=(0, 6))
            setattr(self, attr, lbl)

    def _criar_conteudo(self, parent):
        self.tabview = ctk.CTkTabview(parent, corner_radius=8)
        self.tabview.grid(row=3, column=0, sticky="nsew")
        self.tabview.add("📋 Log de Publicação")
        self.tabview.add("📄 PDFs encontrados")

        # Aba Log
        log_tab = self.tabview.tab("📋 Log de Publicação")
        log_tab.grid_columnconfigure(0, weight=1)
        log_tab.grid_rowconfigure(1, weight=1)

        top = ctk.CTkFrame(log_tab, fg_color="transparent")
        top.grid(row=0, column=0, sticky="ew", pady=(4, 4))
        ctk.CTkButton(top, text="Limpar", command=self._limpar_log,
                      width=70, height=26, font=ctk.CTkFont(size=10)).pack(side="right", padx=(4, 0))
        ctk.CTkButton(top, text="Exportar", command=self._exportar_log,
                      width=70, height=26, font=ctk.CTkFont(size=10)).pack(side="right")

        self.txt_log = ctk.CTkTextbox(
            log_tab, font=ctk.CTkFont(family="Consolas", size=11),
            fg_color=self.CORES['fundo_escuro'], wrap="word", state="disabled",
        )
        self.txt_log.grid(row=1, column=0, sticky="nsew")

        for tag, cor in [
            ("sucesso", self.CORES['sucesso']), ("erro", self.CORES['erro']),
            ("aviso", self.CORES['aviso']),     ("info", self.CORES['info']),
            ("processando", self.CORES['processando']), ("texto", self.CORES['texto']),
        ]:
            self.txt_log.tag_config(tag, foreground=cor)

        # Aba Preview
        prev_tab = self.tabview.tab("📄 PDFs encontrados")
        prev_tab.grid_columnconfigure(0, weight=1)
        prev_tab.grid_rowconfigure(0, weight=1)

        self.txt_preview = ctk.CTkTextbox(
            prev_tab, font=ctk.CTkFont(family="Consolas", size=11),
            fg_color=self.CORES['fundo_escuro'], wrap="none", state="disabled",
        )
        self.txt_preview.grid(row=0, column=0, sticky="nsew")

    # ── Ações da GUI ──────────────────────────────────────────────────────────

    def _selecionar_pasta(self):
        pasta = self._askdirectory_powershell()
        if pasta:
            self.pasta_var.set(pasta)
            threading.Thread(target=self._carregar_preview, args=(pasta,), daemon=True).start()

    @staticmethod
    def _askdirectory_powershell() -> str:
        """Abre diálogo de pasta via PowerShell — evita travamento com CustomTkinter."""
        import subprocess
        script = (
            "Add-Type -AssemblyName System.Windows.Forms;"
            "$d = New-Object System.Windows.Forms.FolderBrowserDialog;"
            "$d.Description = 'Selecionar pasta com PDFs das taxas';"
            "$d.ShowNewFolderButton = $false;"
            "if ($d.ShowDialog() -eq 'OK') { $d.SelectedPath }"
        )
        try:
            result = subprocess.run(
                ["powershell", "-NoProfile", "-NonInteractive", "-Command", script],
                capture_output=True, text=True, timeout=120,
            )
            pasta = result.stdout.strip()
            return pasta if pasta and os.path.isdir(pasta) else ""
        except Exception:
            return ""

    def _on_tipo_changed(self, valor):
        cor = self.CORES['destaque'] if valor == "Panificação" else self.CORES['processando']
        self.lbl_tipo_ativo.configure(text=f"● Taxa {valor}", text_color=cor)

    def _preview_pasta(self):
        pasta = self.pasta_var.get()
        if not pasta or not os.path.isdir(pasta):
            messagebox.showwarning("Aviso", "Selecione uma pasta válida primeiro.")
            return
        threading.Thread(target=self._carregar_preview, args=(pasta,), daemon=True).start()

    def _carregar_preview(self, pasta: str):
        # Roda em thread — toda atualização de widget via window.after()
        pdfs = listar_pdfs_da_pasta(pasta)
        self._pdfs_carregados = pdfs

        def _atualizar_ui():
            self.txt_preview.configure(state="normal")
            self.txt_preview.delete("1.0", "end")

            if not pdfs:
                self.txt_preview.insert("end",
                    "Nenhum PDF encontrado com o padrão esperado:\n"
                    "  {codigo} - {nome empresa} - {MM.YYYY}.pdf\n\n"
                    f"Pasta: {pasta}\n"
                )
                self.adicionar_log(f"⚠️ Nenhum PDF válido encontrado em: {pasta}", logging.WARNING)
            else:
                header = f"{'Nº GMS':<10} {'Período':<10} {'Arquivo'}\n"
                sep = "─" * 80 + "\n"
                self.txt_preview.insert("end", f"Pasta: {pasta}\n{sep}{header}{sep}")
                for p in pdfs:
                    linha = f"{p['codigo']:<10} {p['periodo']:<10} {p['arquivo']}\n"
                    self.txt_preview.insert("end", linha)
                self.txt_preview.insert("end", f"{sep}Total: {len(pdfs)} PDF(s) encontrado(s)\n")
                self.adicionar_log(
                    f"📂 {len(pdfs)} PDF(s) encontrado(s) em: {os.path.basename(pasta)}", logging.INFO
                )
                self.tabview.set("📄 PDFs encontrados")

            self.txt_preview.configure(state="disabled")

        self.window.after(0, _atualizar_ui)

    def _iniciar(self):
        pasta = self.pasta_var.get()
        if not pasta or not os.path.isdir(pasta):
            messagebox.showwarning("Aviso", "Selecione uma pasta com os PDFs.")
            return

        pdfs = listar_pdfs_da_pasta(pasta)
        if not pdfs:
            messagebox.showerror("Erro",
                "Nenhum PDF encontrado na pasta com o padrão:\n"
                "{codigo} - {nome empresa} - {MM.YYYY}.pdf")
            return

        self.executando = True
        self.pausa_solicitada = False
        self.stats = {'processados': 0, 'sucesso': 0, 'erros': 0, 'tempo_inicio': time.time()}

        self.btn_iniciar.configure(state="disabled")
        self.btn_pausar.configure(state="normal", text="⏸ Pausar")
        self.btn_parar.configure(state="normal")
        self._set_status("Executando", self.CORES['sucesso'])

        self.thread_automacao = threading.Thread(
            target=self._thread_pub, args=(pdfs,), daemon=True
        )
        self.thread_automacao.start()
        self._tick_timer()

    def _pausar(self):
        if not self.executando:
            return
        self.pausa_solicitada = not self.pausa_solicitada
        if self.pausa_solicitada:
            self.btn_pausar.configure(text="▶ Continuar")
            self._set_status("Pausado", self.CORES['aviso'])
            self.adicionar_log("⏸ Pausado pelo usuário", logging.WARNING)
        else:
            self.btn_pausar.configure(text="⏸ Pausar")
            self._set_status("Executando", self.CORES['sucesso'])
            self.adicionar_log("▶ Retomando...", logging.INFO)

    def _parar(self):
        self.executando = False
        self.pausa_solicitada = False
        self._set_status("Parando...", self.CORES['erro'])
        self.adicionar_log("⏹️ Parando processo...", logging.WARNING)

    def _set_status(self, texto, cor):
        self.status_var.set(texto)
        self.status_indicator.configure(fg_color=cor)

    def _thread_pub(self, pdfs: list[dict]):
        try:
            tipo = self.tipo_taxa_var.get()
            automacao = PublicacaoNuvem(logger=self.logger, gui=self)
            automacao.publicar(pdfs, tipo)
        except Exception as e:
            self.adicionar_log(f"❌ Erro crítico: {e}", logging.ERROR)
            self.adicionar_log(traceback.format_exc(), logging.ERROR)
        finally:
            self.window.after(0, self._on_finalizado)

    def _on_finalizado(self):
        self.executando = False
        self.pausa_solicitada = False
        self._set_status("Concluído", self.CORES['destaque'])
        self.btn_iniciar.configure(state="normal")
        self.btn_pausar.configure(state="disabled", text="⏸ Pausar")
        self.btn_parar.configure(state="disabled")

    def _tick_timer(self):
        if self.executando and self.stats['tempo_inicio']:
            elapsed = time.time() - self.stats['tempo_inicio']
            m, s = divmod(int(elapsed), 60)
            self.lbl_tempo.configure(text=f"{m:02d}:{s:02d}")
            self.window.after(1000, self._tick_timer)

    # ── Log helpers ───────────────────────────────────────────────────────────

    def adicionar_log(self, msg: str, level: int = logging.INFO):
        ts = datetime.now().strftime("%H:%M:%S")
        linha = f"[{ts}] {msg}\n"

        if level == logging.ERROR or "❌" in msg:
            tag = "erro"
        elif level == logging.WARNING or "⚠️" in msg:
            tag = "aviso"
        elif "✅" in msg or "publicado" in msg.lower():
            tag = "sucesso"
        elif "📤" in msg or "processando" in msg.lower():
            tag = "processando"
        else:
            tag = "texto"

        self.txt_log.configure(state="normal")
        self.txt_log.insert("end", linha, tag)
        self.txt_log.see("end")
        self.txt_log.configure(state="disabled")

    def atualizar_stats(self, processados: int, sucesso: int, erros: int, empresa: str = ""):
        self.window.after(0, lambda: self.lbl_sucesso.configure(text=str(sucesso)))
        self.window.after(0, lambda: self.lbl_erros.configure(text=str(erros)))
        if empresa:
            curto = empresa[:24] + "…" if len(empresa) > 24 else empresa
            self.window.after(0, lambda: self.lbl_empresa.configure(text=curto))

    def _limpar_log(self):
        self.txt_log.configure(state="normal")
        self.txt_log.delete("1.0", "end")
        self.txt_log.configure(state="disabled")

    def _exportar_log(self):
        path = filedialog.asksaveasfilename(
            defaultextension=".txt",
            filetypes=[("Texto", "*.txt"), ("Todos", "*.*")],
            initialfile=f"pub_log_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt",
        )
        if path:
            conteudo = self.txt_log.get("1.0", "end")
            with open(path, "w", encoding="utf-8") as f:
                f.write(conteudo)
            self.adicionar_log(f"💾 Log exportado: {path}")

    def ao_fechar(self):
        if self.executando:
            if not messagebox.askyesno("Confirmar", "Automação em andamento. Deseja parar e fechar?"):
                return
            self.executando = False
        self.window.destroy()

    def executar(self):
        self.window.mainloop()


# ─────────────────────────────── Automação Core ─────────────────────────────


class PublicacaoNuvem:
    """Conecta ao Domínio Folha e publica via 'Publicação de Documentos Externos'."""

    def __init__(self, logger: logging.Logger, gui: PublicarTaxasGUI):
        timings.Timings.window_find_timeout = 20
        self.logger = logger
        self.gui = gui
        self.app = None
        self.main_window = None
        self.discord_webhook = os.getenv("DISCORD_WEBHOOK_URL", "")

        log_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "logs")
        self.log_file = os.path.join(log_dir, f"pub_nuvem_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log")

    def log(self, msg: str):
        self.logger.info(msg)
        with open(self.log_file, "a", encoding="utf-8") as f:
            f.write(f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')} - {msg}\n")

    def should_stop(self) -> bool:
        return not self.gui.executando

    def check_pause(self):
        while self.gui.pausa_solicitada and self.gui.executando:
            time.sleep(0.3)

    def smart_sleep(self, seconds: float) -> bool:
        elapsed = 0.0
        while elapsed < seconds:
            if self.should_stop():
                return False
            self.check_pause()
            t = min(0.15, seconds - elapsed)
            time.sleep(t)
            elapsed += t
        return True

    # ── Conexão ───────────────────────────────────────────────────────────────

    def _find_dominio_hwnd(self) -> Optional[int]:
        self.log("🔍 Procurando janela do Domínio Folha...")
        try:
            for hwnd in findwindows.find_windows():
                try:
                    title = win32gui.GetWindowText(hwnd)
                    if "Domínio" in title and "Folha" in title:
                        self.log(f"✅ Janela encontrada: '{title}'")
                        return hwnd
                except Exception:
                    continue
        except Exception:
            pass
        windows = findwindows.find_windows(title_re=".*Domínio Folha.*")
        return windows[0] if windows else None

    def connect_to_dominio(self) -> bool:
        hwnd = self._find_dominio_hwnd()
        if not hwnd:
            self.log("❌ Domínio Folha não encontrado. Abra o sistema antes de iniciar.")
            return False
        try:
            if win32gui.IsIconic(hwnd):
                win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
                time.sleep(1)
            ctypes.windll.user32.keybd_event(0x12, 0, 0, 0)
            ctypes.windll.user32.keybd_event(0x12, 0, 2, 0)
            win32gui.SetForegroundWindow(hwnd)
            time.sleep(0.5)
            self.app = Application(backend="uia").connect(handle=hwnd)
            self.main_window = self.app.window(handle=hwnd)
            self.log("✅ Conectado ao Domínio Folha")
            return True
        except Exception as e:
            self.log(f"❌ Erro ao conectar: {e}")
            return False

    def _is_alive(self) -> bool:
        if not self.app or not self.main_window:
            return False
        try:
            return win32gui.IsWindow(self.main_window.handle)
        except Exception:
            return False

    # ── Publicação de um documento ────────────────────────────────────────────

    def _encontrar_janela_pub(self):
        try:
            pub = self.main_window.child_window(
                title="Publicação de Documentos Externos",
                class_name="FNWND3190",
            )
            if pub.exists() and pub.is_visible():
                return pub
        except Exception:
            pass
        self.log("❌ Janela 'Publicação de Documentos Externos' não encontrada.")
        self.log("   → Abra-a no Domínio: menu GMS > Publicação de Documentos Externos")
        return None

    def _aguardar_dialogo_confirmacao(self, timeout=20) -> Optional[object]:
        inicio = time.time()
        while time.time() - inicio < timeout:
            if self.should_stop():
                return False
            self.check_pause()
            try:
                for hwnd in findwindows.find_windows(class_name="#32770"):
                    try:
                        if win32gui.IsWindowVisible(hwnd):
                            return self.app.window(handle=hwnd)
                    except Exception:
                        continue
            except Exception:
                pass
            try:
                for d in self.main_window.children(class_name="#32770"):
                    if d.exists() and d.is_visible():
                        return d
            except Exception:
                pass
            time.sleep(0.3)
        return None

    def _clicar_ok(self, dialog) -> bool:
        for titulo in ("OK", "Ok", "Sim", "Yes", "Confirmar"):
            try:
                btn = dialog.child_window(title=titulo, control_type="Button")
                if btn.exists(timeout=2):
                    btn.click()
                    return True
            except Exception:
                continue
        for aid in ("1", "6", "1001", "2001"):
            try:
                btn = dialog.child_window(auto_id=aid, control_type="Button")
                if btn.exists(timeout=2):
                    btn.click()
                    return True
            except Exception:
                continue
        try:
            botoes = dialog.children(control_type="Button")
            if botoes:
                botoes[0].click()
                return True
        except Exception:
            pass
        return False

    def _publicar_um(self, pub_window, caminho_pdf: str, numero: str, nome: str) -> bool:
        try:
            if self.should_stop():
                return False

            # Campo Caminho (auto_id=1013)
            campo_caminho = pub_window.child_window(auto_id="1013", class_name="Edit")
            if not campo_caminho.exists(timeout=5):
                self.log("❌ Campo 'Caminho' não encontrado")
                return False
            campo_caminho.set_focus()
            campo_caminho.type_keys("^a{DELETE}")
            time.sleep(0.2)
            campo_caminho.set_text(caminho_pdf)
            self.log(f"   📂 {os.path.basename(caminho_pdf)}")

            if not self.smart_sleep(0.4):
                return False

            # Campo Nº (auto_id=1001)
            campo_num = pub_window.child_window(auto_id="1001", class_name="PBEDIT190")
            if not campo_num.exists(timeout=5):
                self.log("❌ Campo 'Nº' não encontrado")
                return False
            campo_num.set_focus()
            campo_num.type_keys("^a{DELETE}")
            time.sleep(0.2)
            campo_num.set_text(numero)
            self.log(f"   🔢 Nº GMS: {numero}")

            if not self.smart_sleep(0.4):
                return False
            if self.should_stop():
                return False

            # Botão Publicar (auto_id=1003)
            btn = pub_window.child_window(auto_id="1003", class_name="Button")
            if not btn.exists(timeout=5):
                self.log("❌ Botão 'Publicar' não encontrado")
                return False
            self.log(f"📤 Publicando: {nome}")
            btn.click()

            if not self.smart_sleep(2):
                return False
            if self.should_stop():
                return False

            dialog = self._aguardar_dialogo_confirmacao(timeout=20)
            if dialog is False:
                self.log("⏹️ Interrompido")
                return False
            if dialog is None:
                self.log(f"⚠️ Confirmação não apareceu: {nome}")
                return False

            if self._clicar_ok(dialog):
                self.log(f"✅ Publicado: {nome}")
                time.sleep(0.8)
                return True
            else:
                self.log(f"❌ Falha ao clicar OK: {nome}")
                return False

        except ElementNotFoundError as e:
            self.log(f"⚠️ Elemento não encontrado ({nome}): {e}")
            return False
        except Exception as e:
            self.log(f"❌ Erro ({nome}): {e}")
            return False

    # ── Fluxo principal ───────────────────────────────────────────────────────

    def publicar(self, pdfs: list[dict], tipo_taxa: str):
        self.log(f"🚀 Iniciando publicação via nuvem — Taxa {tipo_taxa}")
        self.log(f"   {len(pdfs)} PDF(s) para publicar")

        if not self.connect_to_dominio():
            return

        pub_window = self._encontrar_janela_pub()
        if pub_window is None:
            return

        pub_window.set_focus()
        time.sleep(0.3)

        total = len(pdfs)
        processados = sucesso = erros = 0

        for pdf in pdfs:
            if self.should_stop():
                self.log("⏹️ Processo interrompido pelo usuário")
                break

            self.check_pause()

            processados += 1
            nome = pdf['nome']
            codigo = pdf['codigo']
            caminho = pdf['caminho']
            periodo = pdf['periodo']

            self.log(f"\n[{processados}/{total}] Código {codigo} — {nome} — {periodo}")
            self.gui.atualizar_stats(processados, sucesso, erros, nome)

            if not os.path.exists(caminho):
                self.log(f"   ⚠️ PDF não encontrado: {caminho}")
                erros += 1
                self.gui.error_logger.error(f"PDF_NAO_ENCONTRADO | {codigo} | {nome} | {caminho}")
                continue

            # Reconectar se necessário
            if not self._is_alive():
                self.log("🔄 Reconectando ao Domínio...")
                if not self.connect_to_dominio():
                    self.log("❌ Não foi possível reconectar.")
                    break
                pub_window = self._encontrar_janela_pub()
                if pub_window is None:
                    break
                pub_window.set_focus()
                time.sleep(0.3)

            ok = self._publicar_um(pub_window, caminho, codigo, nome)

            if ok:
                sucesso += 1
                self.gui.success_logger.info(
                    f"OK | Código={codigo} | {nome} | {periodo} | {caminho}"
                )
            else:
                erros += 1
                self.gui.error_logger.error(
                    f"FALHA | Código={codigo} | {nome} | {periodo} | {caminho}"
                )

            self.gui.atualizar_stats(processados, sucesso, erros, nome)

        self.log(f"\n{'='*50}")
        self.log(f"🏁 Publicação Taxa {tipo_taxa} concluída")
        self.log(f"   Total:   {total}")
        self.log(f"   Sucesso: {sucesso}")
        self.log(f"   Erros:   {erros}")
        self.log(f"{'='*50}")

        self._notificar_discord(tipo_taxa, total, sucesso, erros)

    # ── Discord ───────────────────────────────────────────────────────────────

    def _notificar_discord(self, tipo: str, total: int, sucesso: int, erros: int):
        if not _requests_available or not self.discord_webhook:
            return

        if sucesso == total:
            cor, status = 0x00FF00, "Concluído com Sucesso"
        elif sucesso > 0:
            cor, status = 0xFFA500, "Concluído com Avisos"
        else:
            cor, status = 0xFF0000, "Falha na Publicação"

        payload = {
            "content": "<@&1299044385899548752>",
            "embeds": [{
                "title": f"DomBot - Publicação Taxa {tipo} — {status}",
                "color": cor,
                "fields": [
                    {"name": "Total de PDFs",        "value": str(total),   "inline": True},
                    {"name": "Publicados c/ sucesso", "value": str(sucesso), "inline": True},
                    {"name": "Erros",                "value": str(erros),   "inline": True},
                ],
                "footer": {"text": f"DomBot_Taxas • {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}"},
            }],
        }
        try:
            r = requests.post(self.discord_webhook, json=payload, timeout=10)
            if r.status_code == 204:
                self.log("✅ Notificação enviada ao Discord")
            else:
                self.log(f"⚠️ Discord retornou status {r.status_code}")
        except Exception as e:
            self.log(f"⚠️ Erro ao notificar Discord: {e}")


# ─────────────────────────────────── Main ────────────────────────────────────

def main():
    try:
        gui = PublicarTaxasGUI()
        gui.executar()
    except Exception as e:
        print(f"Erro crítico: {e}")
        traceback.print_exc()


if __name__ == "__main__":
    main()
