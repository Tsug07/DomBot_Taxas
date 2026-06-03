"""
DomBot - Publicar Taxas via Nuvem
==================================
Automação RPA para publicar PDFs de Taxa Panificação e Taxa Motorista
no Domínio Folha via "Publicação de Documentos Externos" (GMS Nuvem).

Planilha esperada (mesma usada para emissão):
  Coluna A - Nº       : Número GMS do documento (para publicação)
  Coluna B - Periodo  : Período de competência
  Coluna C - Empresa  : Nome/código da empresa
  Coluna D - Caminho  : Caminho completo do PDF gerado

Autor: Hugo L. Almeida
Versão: 1.0
"""

import customtkinter as ctk
from tkinter import filedialog, messagebox
import threading
import os
import time
import pandas as pd
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
import win32api
import ctypes

try:
    from dotenv import load_dotenv
    load_dotenv()
    import requests
    _requests_available = True
except ImportError:
    _requests_available = False


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
        'sucesso':       '#2ECC71',
        'erro':          '#E74C3C',
        'aviso':         '#F39C12',
        'info':          '#3498DB',
        'texto':         '#ECF0F1',
        'fundo_card':    '#2C3E50',
        'fundo_escuro':  '#1A252F',
        'destaque':      '#1ABC9C',
        'processando':   '#9B59B6',
    }

    def __init__(self):
        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("green")

        self.window = ctk.CTk()
        self.window.title("DomBot - Publicar Taxas via Nuvem v1.0")
        self.window.geometry("860x540")
        self.window.minsize(720, 460)
        self.window.protocol("WM_DELETE_WINDOW", self.ao_fechar)

        self.executando = False
        self.pausa_solicitada = False
        self.thread_automacao = None

        self.stats = {
            'processados': 0,
            'sucesso': 0,
            'erros': 0,
            'tempo_inicio': None,
        }

        self.logs_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "logs")
        os.makedirs(self.logs_dir, exist_ok=True)

        self.setup_file_logging()
        self.set_window_icon()

        self.arquivo_excel = ctk.StringVar()
        self.tipo_taxa_var = ctk.StringVar(value="Panificação")
        self.linha_inicial = ctk.StringVar(value="2")
        self.status_var = ctk.StringVar(value="Aguardando início...")

        self.df_carregado = None

        self.logger = logging.getLogger('PublicarTaxasNuvem')
        self.logger.setLevel(logging.INFO)
        self.logger.handlers = []

        handler = GUILogHandler(self)
        handler.setFormatter(logging.Formatter('%(message)s'))
        self.logger.addHandler(handler)

        self.criar_interface()

    def setup_file_logging(self):
        data_atual = datetime.now().strftime("%Y-%m-%d")

        self.success_logger = logging.getLogger('PublicacaoSuccess')
        self.success_logger.setLevel(logging.INFO)
        if not self.success_logger.handlers:
            h = logging.FileHandler(
                os.path.join(self.logs_dir, f'pub_success_{data_atual}.log'), encoding='utf-8'
            )
            h.setFormatter(logging.Formatter('%(asctime)s - %(message)s', '%Y-%m-%d %H:%M:%S'))
            self.success_logger.addHandler(h)

        self.error_logger = logging.getLogger('PublicacaoError')
        self.error_logger.setLevel(logging.ERROR)
        if not self.error_logger.handlers:
            h = logging.FileHandler(
                os.path.join(self.logs_dir, f'pub_error_{data_atual}.log'), encoding='utf-8'
            )
            h.setFormatter(logging.Formatter('%(asctime)s - %(message)s', '%Y-%m-%d %H:%M:%S'))
            self.error_logger.addHandler(h)

    def set_window_icon(self):
        try:
            icon_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "assets", "favicon.ico")
            if os.name == 'nt' and os.path.exists(icon_path):
                self.window.iconbitmap(icon_path)
        except Exception:
            pass

    # ── Interface ────────────────────────────────────────────────────────────

    def criar_interface(self):
        self.window.grid_columnconfigure(0, weight=1)
        self.window.grid_rowconfigure(0, weight=1)

        main_frame = ctk.CTkFrame(self.window, fg_color="transparent")
        main_frame.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)
        main_frame.grid_columnconfigure(0, weight=1)
        main_frame.grid_rowconfigure(4, weight=1)

        self.criar_header(main_frame)
        self.criar_painel_config(main_frame)
        self.criar_painel_estatisticas(main_frame)
        self.criar_area_logs(main_frame)

    def criar_header(self, parent):
        header = ctk.CTkFrame(parent, fg_color=self.CORES['fundo_card'], corner_radius=8)
        header.grid(row=0, column=0, sticky="ew", pady=(0, 6))
        header.grid_columnconfigure(1, weight=1)

        logo_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "assets", "DomBot_New.png")
        if os.path.exists(logo_path):
            size, circle_size = 66, 44
            bg = Image.new("RGBA", (size, size), (0, 0, 0, 0))
            mask = Image.new("L", (circle_size, circle_size), 0)
            ImageDraw.Draw(mask).ellipse((0, 0, circle_size - 1, circle_size - 1), fill=255)
            circle = Image.new("RGBA", (circle_size, circle_size), (255, 255, 255, 255))
            offset = (size - circle_size) // 2
            bg.paste(circle, (offset, offset), mask)
            original = Image.open(logo_path).convert("RGBA").resize((size, size), Image.LANCZOS)
            bg.paste(original, (0, 0), original)
            logo_img = ctk.CTkImage(light_image=bg, dark_image=bg, size=(size, size))
            ctk.CTkLabel(header, image=logo_img, text="").grid(row=0, column=0, padx=10, pady=8)
        else:
            logo_frame = ctk.CTkFrame(header, fg_color=self.CORES['destaque'],
                                      width=44, height=44, corner_radius=22)
            logo_frame.grid(row=0, column=0, padx=10, pady=8)
            logo_frame.grid_propagate(False)
            ctk.CTkLabel(logo_frame, text="☁️", font=("Segoe UI Emoji", 18)).place(relx=0.5, rely=0.5, anchor="center")

        ctk.CTkLabel(
            header,
            text="DomBot - Publicar Taxas via Nuvem",
            font=ctk.CTkFont(size=16, weight="bold"),
            text_color=self.CORES['texto'],
        ).grid(row=0, column=1, sticky="w", padx=5)

        status_frame = ctk.CTkFrame(header, fg_color="transparent")
        status_frame.grid(row=0, column=2, padx=10)

        self.status_indicator = ctk.CTkFrame(
            status_frame, fg_color="#7F8C8D", width=10, height=10, corner_radius=5
        )
        self.status_indicator.pack(side="left", padx=(0, 6))
        ctk.CTkLabel(
            status_frame,
            textvariable=self.status_var,
            font=ctk.CTkFont(size=11),
            text_color="#95A5A6",
        ).pack(side="left")

    def criar_painel_config(self, parent):
        config = ctk.CTkFrame(parent, fg_color=self.CORES['fundo_card'], corner_radius=8)
        config.grid(row=1, column=0, sticky="ew", pady=(0, 6))
        config.grid_columnconfigure(0, weight=1)

        # Linha 1: Excel + tipo taxa + linha inicial
        row1 = ctk.CTkFrame(config, fg_color="transparent")
        row1.grid(row=0, column=0, sticky="ew", padx=10, pady=(8, 4))
        row1.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(row1, text="📁", font=ctk.CTkFont(size=14)).grid(row=0, column=0, padx=(0, 5))

        self.entry_arquivo = ctk.CTkEntry(
            row1, textvariable=self.arquivo_excel,
            placeholder_text="Selecione a planilha Excel (mesma usada na emissão)...",
            height=32, font=ctk.CTkFont(size=11),
        )
        self.entry_arquivo.grid(row=0, column=1, sticky="ew", padx=(0, 8))

        ctk.CTkButton(
            row1, text="Procurar", command=self.selecionar_arquivo,
            width=80, height=32, font=ctk.CTkFont(size=11),
        ).grid(row=0, column=2, padx=(0, 8))

        # Tipo de taxa
        ctk.CTkLabel(row1, text="Taxa:", font=ctk.CTkFont(size=11)).grid(row=0, column=3, padx=(0, 4))
        tipo_menu = ctk.CTkOptionMenu(
            row1,
            values=["Panificação", "Motorista"],
            variable=self.tipo_taxa_var,
            width=120, height=32, font=ctk.CTkFont(size=11),
            command=self.on_tipo_taxa_changed,
        )
        tipo_menu.grid(row=0, column=4, padx=(0, 8))

        # Linha inicial
        ctk.CTkLabel(row1, text="Da linha:", font=ctk.CTkFont(size=11)).grid(row=0, column=5, padx=(0, 4))
        ctk.CTkEntry(
            row1, textvariable=self.linha_inicial,
            width=55, height=32, font=ctk.CTkFont(size=11),
        ).grid(row=0, column=6)

        # Linha 2: botões controle
        row2 = ctk.CTkFrame(config, fg_color="transparent")
        row2.grid(row=1, column=0, sticky="ew", padx=10, pady=(4, 8))

        self.btn_iniciar = ctk.CTkButton(
            row2, text="▶ Publicar", command=self.iniciar,
            width=110, height=32, font=ctk.CTkFont(size=12, weight="bold"),
            fg_color="#27AE60", hover_color="#1E8449",
        )
        self.btn_iniciar.pack(side="left", padx=(0, 6))

        self.btn_pausar = ctk.CTkButton(
            row2, text="⏸ Pausar", command=self.pausar,
            width=100, height=32, font=ctk.CTkFont(size=11),
            fg_color="#F39C12", hover_color="#D68910", state="disabled",
        )
        self.btn_pausar.pack(side="left", padx=(0, 6))

        self.btn_parar = ctk.CTkButton(
            row2, text="■ Parar", command=self.parar,
            width=90, height=32, font=ctk.CTkFont(size=11),
            fg_color="#E74C3C", hover_color="#C0392B", state="disabled",
        )
        self.btn_parar.pack(side="left", padx=(0, 12))

        ctk.CTkButton(
            row2, text="📋 Pré-visualizar", command=self.preview_excel,
            width=130, height=32, font=ctk.CTkFont(size=11),
        ).pack(side="left")

        # Label de tipo ativo
        self.lbl_tipo_ativo = ctk.CTkLabel(
            row2,
            text="● Taxa Panificação",
            font=ctk.CTkFont(size=11),
            text_color=self.CORES['destaque'],
        )
        self.lbl_tipo_ativo.pack(side="right", padx=10)

    def criar_painel_estatisticas(self, parent):
        stats_frame = ctk.CTkFrame(parent, fg_color=self.CORES['fundo_card'], corner_radius=8)
        stats_frame.grid(row=2, column=0, sticky="ew", pady=(0, 6))

        for i in range(4):
            stats_frame.grid_columnconfigure(i, weight=1)

        labels_info = [
            ("Publicados", 'sucesso', 'lbl_sucesso'),
            ("Erros",       'erro',    'lbl_erros'),
            ("Empresa atual", 'info',  'lbl_empresa'),
            ("Tempo decorrido", 'aviso', 'lbl_tempo'),
        ]

        for col, (titulo, cor_key, attr) in enumerate(labels_info):
            card = ctk.CTkFrame(stats_frame, fg_color=self.CORES['fundo_escuro'], corner_radius=6)
            card.grid(row=0, column=col, sticky="ew", padx=6, pady=6)

            ctk.CTkLabel(card, text=titulo, font=ctk.CTkFont(size=10),
                         text_color="#95A5A6").pack(pady=(6, 0))

            lbl = ctk.CTkLabel(card, text="0" if col < 2 else "—",
                               font=ctk.CTkFont(size=18, weight="bold"),
                               text_color=self.CORES[cor_key])
            lbl.pack(pady=(0, 6))
            setattr(self, attr, lbl)

    def criar_area_logs(self, parent):
        log_frame = ctk.CTkFrame(parent, fg_color=self.CORES['fundo_card'], corner_radius=8)
        log_frame.grid(row=4, column=0, sticky="nsew", pady=(0, 0))
        log_frame.grid_columnconfigure(0, weight=1)
        log_frame.grid_rowconfigure(1, weight=1)

        top = ctk.CTkFrame(log_frame, fg_color="transparent")
        top.grid(row=0, column=0, sticky="ew", padx=10, pady=(8, 4))

        ctk.CTkLabel(top, text="Log de Publicação",
                     font=ctk.CTkFont(size=12, weight="bold")).pack(side="left")

        ctk.CTkButton(top, text="Limpar", command=self.limpar_log,
                      width=70, height=26, font=ctk.CTkFont(size=10)).pack(side="right", padx=(6, 0))
        ctk.CTkButton(top, text="Exportar", command=self.exportar_log,
                      width=70, height=26, font=ctk.CTkFont(size=10)).pack(side="right")

        self.txt_log = ctk.CTkTextbox(
            log_frame, font=ctk.CTkFont(family="Consolas", size=11),
            fg_color=self.CORES['fundo_escuro'], wrap="word", state="disabled",
        )
        self.txt_log.grid(row=1, column=0, sticky="nsew", padx=10, pady=(0, 10))

        self.txt_log.tag_config("sucesso",     foreground=self.CORES['sucesso'])
        self.txt_log.tag_config("erro",        foreground=self.CORES['erro'])
        self.txt_log.tag_config("aviso",       foreground=self.CORES['aviso'])
        self.txt_log.tag_config("info",        foreground=self.CORES['info'])
        self.txt_log.tag_config("processando", foreground=self.CORES['processando'])
        self.txt_log.tag_config("texto",       foreground=self.CORES['texto'])

    # ── Ações GUI ────────────────────────────────────────────────────────────

    def selecionar_arquivo(self):
        path = filedialog.askopenfilename(
            title="Selecionar planilha de taxas",
            filetypes=[("Excel", "*.xlsx *.xls"), ("Todos", "*.*")],
        )
        if path:
            self.arquivo_excel.set(path)
            self.adicionar_log(f"📄 Arquivo selecionado: {os.path.basename(path)}", logging.INFO)
            self._carregar_preview(path)

    def on_tipo_taxa_changed(self, valor):
        cor = self.CORES['destaque'] if valor == "Panificação" else self.CORES['processando']
        self.lbl_tipo_ativo.configure(text=f"● Taxa {valor}", text_color=cor)

    def preview_excel(self):
        path = self.arquivo_excel.get()
        if not path or not os.path.exists(path):
            messagebox.showwarning("Aviso", "Selecione um arquivo Excel válido primeiro.")
            return
        self._carregar_preview(path)

    def _carregar_preview(self, path):
        try:
            df = pd.read_excel(path, header=0)
            self.df_carregado = df
            self.adicionar_log(f"📊 Planilha carregada: {len(df)} linha(s)", logging.INFO)
            self.adicionar_log(f"   Colunas: {', '.join(str(c) for c in df.columns.tolist())}", logging.INFO)

            linha_ini = self._get_linha_inicial()
            df_view = df.iloc[linha_ini - 2:linha_ini + 8] if linha_ini > 1 else df.head(10)
            for i, row in df_view.iterrows():
                vals = [str(v)[:20] for v in row.values]
                self.adicionar_log(f"   L{i+2}: {' | '.join(vals)}", logging.INFO)
        except Exception as e:
            self.adicionar_log(f"❌ Erro ao carregar planilha: {e}", logging.ERROR)

    def _get_linha_inicial(self) -> int:
        try:
            v = int(self.linha_inicial.get())
            return max(2, v)
        except ValueError:
            return 2

    def iniciar(self):
        path = self.arquivo_excel.get()
        if not path or not os.path.exists(path):
            messagebox.showwarning("Aviso", "Selecione um arquivo Excel válido.")
            return

        self.executando = True
        self.pausa_solicitada = False
        self.stats = {'processados': 0, 'sucesso': 0, 'erros': 0, 'tempo_inicio': time.time()}

        self.btn_iniciar.configure(state="disabled")
        self.btn_pausar.configure(state="normal", text="⏸ Pausar")
        self.btn_parar.configure(state="normal")
        self._set_status("Executando", self.CORES['sucesso'])

        self.thread_automacao = threading.Thread(
            target=self._thread_publicacao, args=(path,), daemon=True
        )
        self.thread_automacao.start()
        self._tick_timer()

    def pausar(self):
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

    def parar(self):
        self.executando = False
        self.pausa_solicitada = False
        self._set_status("Parando...", self.CORES['erro'])
        self.adicionar_log("⏹️ Parando processo...", logging.WARNING)

    def _set_status(self, texto, cor):
        self.status_var.set(texto)
        self.status_indicator.configure(fg_color=cor)

    def _thread_publicacao(self, path_excel):
        try:
            tipo = self.tipo_taxa_var.get()
            linha_ini = self._get_linha_inicial()
            automacao = PublicacaoNuvem(logger=self.logger, gui=self)
            automacao.publicar(path_excel, tipo, linha_ini)
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

    # ── Log helpers ──────────────────────────────────────────────────────────

    def adicionar_log(self, msg: str, level: int = logging.INFO):
        ts = datetime.now().strftime("%H:%M:%S")
        linha = f"[{ts}] {msg}\n"

        if level == logging.ERROR:
            tag = "erro"
        elif level == logging.WARNING:
            tag = "aviso"
        elif "✅" in msg or "publicado" in msg.lower():
            tag = "sucesso"
        elif "❌" in msg:
            tag = "erro"
        elif "⚠️" in msg:
            tag = "aviso"
        elif "📤" in msg or "processando" in msg.lower():
            tag = "processando"
        else:
            tag = "texto"

        self.txt_log.configure(state="normal")
        self.txt_log.insert("end", linha, tag)
        self.txt_log.see("end")
        self.txt_log.configure(state="disabled")

    def atualizar_stats(self, processados, sucesso, erros, empresa=""):
        self.stats['processados'] = processados
        self.stats['sucesso'] = sucesso
        self.stats['erros'] = erros
        self.window.after(0, lambda: self.lbl_sucesso.configure(text=str(sucesso)))
        self.window.after(0, lambda: self.lbl_erros.configure(text=str(erros)))
        if empresa:
            nome_curto = empresa[:22] + "…" if len(empresa) > 22 else empresa
            self.window.after(0, lambda: self.lbl_empresa.configure(text=nome_curto))

    def limpar_log(self):
        self.txt_log.configure(state="normal")
        self.txt_log.delete("1.0", "end")
        self.txt_log.configure(state="disabled")

    def exportar_log(self):
        path = filedialog.asksaveasfilename(
            defaultextension=".txt",
            filetypes=[("Texto", "*.txt"), ("Todos", "*.*")],
            initialfile=f"pub_log_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt",
        )
        if path:
            conteudo = self.txt_log.get("1.0", "end")
            with open(path, "w", encoding="utf-8") as f:
                f.write(conteudo)
            self.adicionar_log(f"💾 Log exportado: {path}", logging.INFO)

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
    """Conecta ao Domínio Folha e publica documentos via 'Publicação de Documentos Externos'."""

    COLUNAS_NECESSARIAS = ['Nº', 'Caminho']

    def __init__(self, logger: logging.Logger, gui: PublicarTaxasGUI):
        timings.Timings.window_find_timeout = 20
        self.logger = logger
        self.gui = gui
        self.app = None
        self.main_window = None
        self.discord_webhook = os.getenv("DISCORD_WEBHOOK_URL", "")

        log_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "logs")
        os.makedirs(log_dir, exist_ok=True)
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
        interval = 0.15
        elapsed = 0.0
        while elapsed < seconds:
            if self.should_stop():
                return False
            self.check_pause()
            t = min(interval, seconds - elapsed)
            time.sleep(t)
            elapsed += t
        return True

    # ── Conexão Domínio ───────────────────────────────────────────────────────

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

    # ── Publicação ────────────────────────────────────────────────────────────

    def _encontrar_janela_pub(self):
        """Localiza a sub-janela 'Publicação de Documentos Externos'."""
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
        self.log("   → Abra-a no Domínio antes de iniciar: menu GMS > Publicação de Documentos Externos")
        return None

    def _aguardar_dialogo_confirmacao(self, timeout=20) -> Optional[object]:
        """Aguarda a janela de confirmação após clicar em Publicar. Retorna a janela ou None."""
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
            # Também tenta como filho da janela principal
            try:
                dialogs = self.main_window.children(class_name="#32770")
                for d in dialogs:
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

    def publicar_documento(self, pub_window, caminho_pdf: str, numero: str, nome_exib: str) -> bool:
        """Preenche os campos e clica em Publicar para um documento."""
        try:
            if self.should_stop():
                return False

            # Campo Caminho (auto_id=1013)
            campo_caminho = pub_window.child_window(auto_id="1013", class_name="Edit")
            if not campo_caminho.exists(timeout=5):
                self.log("❌ Campo 'Caminho' (auto_id=1013) não encontrado")
                return False
            campo_caminho.set_focus()
            campo_caminho.type_keys("^a{DELETE}")
            time.sleep(0.2)
            campo_caminho.set_text(caminho_pdf)
            self.log(f"   📂 Caminho: {os.path.basename(caminho_pdf)}")

            if not self.smart_sleep(0.4):
                return False
            if self.should_stop():
                return False

            # Campo Nº (auto_id=1001)
            campo_numero = pub_window.child_window(auto_id="1001", class_name="PBEDIT190")
            if not campo_numero.exists(timeout=5):
                self.log("❌ Campo 'Nº' (auto_id=1001) não encontrado")
                return False
            campo_numero.set_focus()
            campo_numero.type_keys("^a{DELETE}")
            time.sleep(0.2)
            campo_numero.set_text(numero)
            self.log(f"   🔢 Nº GMS: {numero}")

            if not self.smart_sleep(0.4):
                return False
            if self.should_stop():
                return False

            # Botão Publicar (auto_id=1003)
            btn_publicar = pub_window.child_window(auto_id="1003", class_name="Button")
            if not btn_publicar.exists(timeout=5):
                self.log("❌ Botão 'Publicar' (auto_id=1003) não encontrado")
                return False
            self.log(f"📤 Publicando: {nome_exib}")
            btn_publicar.click()

            if not self.smart_sleep(2):
                return False
            if self.should_stop():
                return False

            # Aguarda confirmação
            dialog = self._aguardar_dialogo_confirmacao(timeout=20)
            if dialog is False:
                self.log("⏹️ Interrompido durante espera de confirmação")
                return False
            if dialog is None:
                self.log(f"⚠️ Janela de confirmação não apareceu para '{nome_exib}'")
                return False

            if self._clicar_ok(dialog):
                self.log(f"✅ Publicado com sucesso: {nome_exib}")
                time.sleep(0.8)
                return True
            else:
                self.log(f"❌ Não foi possível clicar OK para: {nome_exib}")
                return False

        except ElementNotFoundError as e:
            self.log(f"⚠️ Elemento não encontrado ({nome_exib}): {e}")
            return False
        except Exception as e:
            self.log(f"❌ Erro ao publicar ({nome_exib}): {e}")
            return False

    # ── Fluxo principal ───────────────────────────────────────────────────────

    def publicar(self, path_excel: str, tipo_taxa: str, linha_inicial: int):
        self.log(f"🚀 Iniciando publicação via nuvem — Taxa {tipo_taxa}")
        self.log(f"   Planilha: {os.path.basename(path_excel)}")
        self.log(f"   A partir da linha: {linha_inicial}")

        # Carregar Excel
        try:
            df_full = pd.read_excel(path_excel, header=0)
        except Exception as e:
            self.log(f"❌ Erro ao abrir planilha: {e}")
            return

        # Verificar colunas mínimas (Nº e Caminho)
        colunas = df_full.columns.tolist()
        self.log(f"   Colunas encontradas: {', '.join(str(c) for c in colunas)}")

        col_numero = None
        col_caminho = None
        col_nome = None

        for c in colunas:
            cs = str(c).strip()
            if cs in ('Nº', 'N°', 'N', 'Numero', 'Número'):
                col_numero = c
            elif cs.lower() in ('caminho', 'path', 'arquivo', 'caminho completo'):
                col_caminho = c
            elif cs.lower() in ('nome', 'empresa', 'salvar como', 'salvar_como', 'nomearquivo'):
                col_nome = c

        if col_numero is None or col_caminho is None:
            self.log("❌ Colunas obrigatórias não encontradas.")
            self.log("   Necessário: coluna 'Nº' e coluna 'Caminho'")
            self.log(f"   Colunas disponíveis: {', '.join(str(c) for c in colunas)}")
            return

        # Filtrar a partir da linha inicial (linha_inicial é número da linha Excel, header=linha 1)
        idx_start = linha_inicial - 2  # header ocupa linha 1, dados começam em 2
        df = df_full.iloc[idx_start:].reset_index(drop=True)
        total = len(df)

        if total == 0:
            self.log("⚠️ Nenhum dado encontrado a partir da linha especificada.")
            return

        self.log(f"📊 {total} documento(s) para publicar")

        # Conectar ao Domínio
        if not self.connect_to_dominio():
            return

        if not self._is_alive():
            self.log("❌ Conexão com Domínio perdida.")
            return

        # Localizar janela de publicação
        pub_window = self._encontrar_janela_pub()
        if pub_window is None:
            return

        pub_window.set_focus()
        time.sleep(0.3)

        processados = 0
        sucesso = 0
        erros = 0

        for i, row in df.iterrows():
            if self.should_stop():
                self.log("⏹️ Processo interrompido pelo usuário")
                break

            self.check_pause()

            numero = str(row[col_numero]).strip() if pd.notnull(row[col_numero]) else ""
            caminho_pdf = str(row[col_caminho]).strip() if pd.notnull(row[col_caminho]) else ""
            nome_exib = str(row[col_nome]).strip() if col_nome and pd.notnull(row[col_nome]) else f"Linha {idx_start + i + 2}"

            processados += 1
            self.log(f"\n[{processados}/{total}] {nome_exib}")
            self.gui.atualizar_stats(processados, sucesso, erros, nome_exib)

            # Validações básicas
            if not numero:
                self.log(f"   ⚠️ Nº vazio — pulando")
                erros += 1
                self.gui.success_logger.info(f"PULADO (sem Nº) | {nome_exib}")
                continue

            if not caminho_pdf:
                self.log(f"   ⚠️ Caminho vazio — pulando")
                erros += 1
                continue

            if not os.path.exists(caminho_pdf):
                self.log(f"   ⚠️ PDF não encontrado: {caminho_pdf}")
                erros += 1
                self.gui.error_logger.error(f"PDF_NAO_ENCONTRADO | {nome_exib} | {caminho_pdf}")
                continue

            # Reconectar se necessário
            if not self._is_alive():
                self.log("🔄 Reconectando ao Domínio...")
                if not self.connect_to_dominio():
                    self.log("❌ Não foi possível reconectar. Abortando.")
                    break
                pub_window = self._encontrar_janela_pub()
                if pub_window is None:
                    break
                pub_window.set_focus()
                time.sleep(0.3)

            ok = self.publicar_documento(pub_window, caminho_pdf, numero, nome_exib)

            if ok:
                sucesso += 1
                self.gui.success_logger.info(f"OK | {nome_exib} | Nº={numero} | {caminho_pdf}")
            else:
                erros += 1
                self.gui.error_logger.error(f"FALHA | {nome_exib} | Nº={numero} | {caminho_pdf}")

            self.gui.atualizar_stats(processados, sucesso, erros, nome_exib)

        # Resumo final
        self.log(f"\n{'='*50}")
        self.log(f"🏁 Publicação Taxa {tipo_taxa} concluída")
        self.log(f"   Total:    {total}")
        self.log(f"   Sucesso:  {sucesso}")
        self.log(f"   Erros:    {erros}")
        self.log(f"   Pulados:  {total - processados}")
        self.log(f"{'='*50}")

        self._notificar_discord(tipo_taxa, total, processados, sucesso)

    # ── Discord ────────────────────────────────────────────────────────────────

    def _notificar_discord(self, tipo_taxa: str, total: int, processados: int, sucesso: int):
        if not _requests_available or not self.discord_webhook:
            return

        erros = processados - sucesso
        if sucesso == processados:
            cor = 0x00FF00
            status = "Concluído com Sucesso"
        elif sucesso > 0:
            cor = 0xFFA500
            status = "Concluído com Avisos"
        else:
            cor = 0xFF0000
            status = "Falha na Publicação"

        payload = {
            "content": "<@&1299044385899548752>",
            "embeds": [{
                "title": f"DomBot - Publicação Taxa {tipo_taxa} — {status}",
                "color": cor,
                "fields": [
                    {"name": "Total na planilha",  "value": str(total),        "inline": True},
                    {"name": "Processados",         "value": str(processados),  "inline": True},
                    {"name": "Publicados c/ sucesso", "value": str(sucesso),    "inline": True},
                    {"name": "Erros",               "value": str(erros),        "inline": True},
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
