"""
DomBot - Taxa Bares
===================
Automação RPA para emissão de relatórios de Taxa Bares no sistema Domínio Folha.

Autor: Hugo L. Almeida
Versão: 1.0
"""

import customtkinter as ctk
import pandas as pd
import time
import logging
import os
import traceback
import threading
import subprocess
import smtplib
import ssl
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from datetime import datetime
from typing import Optional
import tkinter.messagebox as messagebox
from PIL import Image, ImageDraw

from pywinauto.application import Application
from pywinauto.keyboard import send_keys
from pywinauto import findwindows, timings
import win32gui
import win32con
import win32api

try:
    from dotenv import load_dotenv
    load_dotenv()
    import requests
    _requests_available = True
except ImportError:
    _requests_available = False


class GUILogHandler(logging.Handler):
    def __init__(self, gui):
        super().__init__()
        self.gui = gui

    def emit(self, record):
        msg = self.format(record)
        self.gui.window.after(0, lambda: self.gui.adicionar_log(msg, record.levelno))


class AutomacaoGUI:
    CORES = {
        'sucesso': '#2ECC71',
        'erro': '#E74C3C',
        'aviso': '#F39C12',
        'info': '#3498DB',
        'texto': '#ECF0F1',
        'fundo_card': '#2C3E50',
        'fundo_escuro': '#1A252F',
        'destaque': '#1ABC9C',
        'processando': '#9B59B6',
    }

    def __init__(self):
        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("green")

        self.window = ctk.CTk()
        self.window.title("DomBot - Taxa Bares v1.0")
        self.window.geometry("800x480")
        self.window.minsize(700, 430)
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

        # Variáveis da interface
        self.arquivo_excel = ctk.StringVar()
        self.linha_inicial = ctk.StringVar(value="2")
        self.diretorio_salvamento = ctk.StringVar()
        self.periodo_var = ctk.StringVar(value=datetime.now().strftime("%m/%Y"))
        self.status_var = ctk.StringVar(value="Aguardando início...")

        self.df_carregado = None
        self.linhas_processadas = 0
        self.linhas_com_erro = 0

        self.logger = logging.getLogger('AutomacaoTaxaBares')
        self.logger.setLevel(logging.INFO)
        self.logger.handlers = []

        self.gui_handler = GUILogHandler(self)
        self.gui_handler.setFormatter(logging.Formatter('%(message)s'))
        self.logger.addHandler(self.gui_handler)

        self.criar_interface()

    def setup_file_logging(self):
        data_atual = datetime.now().strftime("%Y-%m-%d")

        self.success_logger = logging.getLogger('SuccessLogTaxaBares')
        self.success_logger.setLevel(logging.INFO)
        if not self.success_logger.handlers:
            h = logging.FileHandler(
                os.path.join(self.logs_dir, f'success_{data_atual}.log'), encoding='utf-8'
            )
            h.setFormatter(logging.Formatter('%(asctime)s - %(message)s', '%Y-%m-%d %H:%M:%S'))
            self.success_logger.addHandler(h)

        self.error_logger = logging.getLogger('ErrorLogTaxaBares')
        self.error_logger.setLevel(logging.ERROR)
        if not self.error_logger.handlers:
            h = logging.FileHandler(
                os.path.join(self.logs_dir, f'error_{data_atual}.log'), encoding='utf-8'
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
        self.criar_area_conteudo(main_frame)

    def criar_header(self, parent):
        header_frame = ctk.CTkFrame(parent, fg_color=self.CORES['fundo_card'], corner_radius=8)
        header_frame.grid(row=0, column=0, sticky="ew", pady=(0, 6))
        header_frame.grid_columnconfigure(1, weight=1)

        logo_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "assets", "DomBot_New.png")
        if os.path.exists(logo_path):
            size = 66
            circle_size = 44
            bg = Image.new("RGBA", (size, size), (0, 0, 0, 0))
            circle_mask = Image.new("L", (circle_size, circle_size), 0)
            ImageDraw.Draw(circle_mask).ellipse((0, 0, circle_size - 1, circle_size - 1), fill=255)
            circle = Image.new("RGBA", (circle_size, circle_size), (255, 255, 255, 255))
            circle_offset = (size - circle_size) // 2
            bg.paste(circle, (circle_offset, circle_offset), circle_mask)
            original = Image.open(logo_path).convert("RGBA")
            original = original.resize((size, size), Image.LANCZOS)
            bg.paste(original, (0, 0), original)
            logo_image = ctk.CTkImage(light_image=bg, dark_image=bg, size=(size, size))
            ctk.CTkLabel(header_frame, image=logo_image, text="").grid(row=0, column=0, padx=10, pady=8)
        else:
            logo_frame = ctk.CTkFrame(header_frame, fg_color=self.CORES['destaque'],
                                      width=44, height=44, corner_radius=22)
            logo_frame.grid(row=0, column=0, padx=10, pady=8)
            logo_frame.grid_propagate(False)
            ctk.CTkLabel(logo_frame, text="🤖", font=("Segoe UI Emoji", 18)).place(relx=0.5, rely=0.5, anchor="center")

        ctk.CTkLabel(
            header_frame,
            text="DomBot - Taxa Bares",
            font=ctk.CTkFont(size=16, weight="bold"),
            text_color=self.CORES['texto'],
        ).grid(row=0, column=1, sticky="w", padx=5)

        self.status_frame = ctk.CTkFrame(header_frame, fg_color="transparent")
        self.status_frame.grid(row=0, column=2, padx=10)

        self.status_indicator = ctk.CTkFrame(
            self.status_frame, fg_color="#7F8C8D", width=10, height=10, corner_radius=5
        )
        self.status_indicator.pack(side="left", padx=(0, 6))

        ctk.CTkLabel(
            self.status_frame,
            textvariable=self.status_var,
            font=ctk.CTkFont(size=11),
            text_color="#95A5A6",
        ).pack(side="left")

    def criar_painel_config(self, parent):
        config_frame = ctk.CTkFrame(parent, fg_color=self.CORES['fundo_card'], corner_radius=8)
        config_frame.grid(row=1, column=0, sticky="ew", pady=(0, 6))
        config_frame.grid_columnconfigure(0, weight=1)

        # --- Linha 1: Excel + linha inicial + botões de controle ---
        row1 = ctk.CTkFrame(config_frame, fg_color="transparent")
        row1.grid(row=0, column=0, sticky="ew", padx=10, pady=(8, 4))
        row1.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(row1, text="📁", font=ctk.CTkFont(size=14)).grid(row=0, column=0, padx=(0, 5))

        self.entry_arquivo = ctk.CTkEntry(
            row1, textvariable=self.arquivo_excel,
            placeholder_text="Selecione o arquivo Excel...",
            height=32, font=ctk.CTkFont(size=11),
        )
        self.entry_arquivo.grid(row=0, column=1, sticky="ew", padx=(0, 8))

        ctk.CTkButton(
            row1, text="Procurar", command=self.selecionar_arquivo,
            width=80, height=32, font=ctk.CTkFont(size=11),
            fg_color=self.CORES['info'], hover_color="#2980B9",
        ).grid(row=0, column=2, padx=(0, 15))

        ctk.CTkLabel(row1, text="Linha:", font=ctk.CTkFont(size=11), text_color="#BDC3C7").grid(
            row=0, column=3, padx=(0, 3)
        )
        ctk.CTkEntry(
            row1, textvariable=self.linha_inicial,
            width=50, height=32, font=ctk.CTkFont(size=11), justify="center",
        ).grid(row=0, column=4, padx=(0, 15))

        self.btn_iniciar = ctk.CTkButton(
            row1, text="▶ Iniciar", command=self.iniciar_automacao_thread,
            width=90, height=32, font=ctk.CTkFont(size=11, weight="bold"),
            fg_color=self.CORES['sucesso'], hover_color="#27AE60",
        )
        self.btn_iniciar.grid(row=0, column=5, padx=3)

        self.btn_pausar = ctk.CTkButton(
            row1, text="⏸ Pausar", command=self.pausar_automacao,
            width=90, height=32, font=ctk.CTkFont(size=11, weight="bold"),
            fg_color=self.CORES['aviso'], hover_color="#E67E22", state="disabled",
        )
        self.btn_pausar.grid(row=0, column=6, padx=3)

        self.btn_parar = ctk.CTkButton(
            row1, text="⏹ Parar", command=self.parar_automacao,
            width=90, height=32, font=ctk.CTkFont(size=11, weight="bold"),
            fg_color=self.CORES['erro'], hover_color="#C0392B", state="disabled",
        )
        self.btn_parar.grid(row=0, column=7, padx=(3, 0))

        # --- Linha 2: Pasta de destino ---
        row2 = ctk.CTkFrame(config_frame, fg_color="transparent")
        row2.grid(row=1, column=0, sticky="ew", padx=10, pady=(0, 8))
        row2.grid_columnconfigure(2, weight=1)

        ctk.CTkLabel(row2, text="📂", font=ctk.CTkFont(size=14)).grid(row=0, column=0, padx=(0, 5))
        ctk.CTkLabel(row2, text="Pasta de destino:", font=ctk.CTkFont(size=11), text_color="#BDC3C7").grid(
            row=0, column=1, padx=(0, 5)
        )
        self.entry_diretorio = ctk.CTkEntry(
            row2, textvariable=self.diretorio_salvamento,
            placeholder_text="Pasta onde os arquivos serão salvos...",
            height=32, font=ctk.CTkFont(size=11),
        )
        self.entry_diretorio.grid(row=0, column=2, sticky="ew", padx=(0, 8))

        ctk.CTkButton(
            row2, text="Selecionar", command=self.selecionar_diretorio,
            width=80, height=32, font=ctk.CTkFont(size=11),
            fg_color=self.CORES['info'], hover_color="#2980B9",
        ).grid(row=0, column=5)

    def criar_painel_estatisticas(self, parent):
        stats_frame = ctk.CTkFrame(parent, fg_color=self.CORES['fundo_card'], corner_radius=8)
        stats_frame.grid(row=2, column=0, sticky="ew", pady=(0, 6))

        for i in range(4):
            stats_frame.grid_columnconfigure(i, weight=1)

        self._criar_stat_card(stats_frame, 0, "✅", "Sucesso", "sucesso_label", "0", self.CORES['sucesso'])
        self._criar_stat_card(stats_frame, 1, "❌", "Erros", "erros_label", "0", self.CORES['erro'])
        self._criar_stat_card(stats_frame, 2, "🏢", "Empresa", "empresa_label", "-", self.CORES['destaque'])
        self._criar_stat_card(stats_frame, 3, "⏱", "Tempo", "tempo_label", "00:00:00", self.CORES['aviso'])

    def _criar_stat_card(self, parent, col, icon, titulo, attr_name, valor_inicial, cor=None):
        card = ctk.CTkFrame(parent, fg_color="transparent")
        card.grid(row=0, column=col, padx=5, pady=8)

        ctk.CTkLabel(
            card, text=f"{icon} {titulo}", font=ctk.CTkFont(size=10), text_color="#7F8C8D"
        ).pack()

        label = ctk.CTkLabel(
            card, text=valor_inicial,
            font=ctk.CTkFont(size=14, weight="bold"),
            text_color=cor if cor else self.CORES['texto'],
        )
        label.pack()
        setattr(self, attr_name, label)

    def criar_area_conteudo(self, parent):
        self.tabview = ctk.CTkTabview(
            parent, fg_color=self.CORES['fundo_card'],
            segmented_button_fg_color=self.CORES['fundo_escuro'],
            segmented_button_selected_color=self.CORES['destaque'],
            corner_radius=8, height=25,
        )
        self.tabview.grid(row=4, column=0, sticky="nsew")

        tab_logs = self.tabview.add("📋 Logs")
        tab_preview = self.tabview.add("📊 Preview")
        tab_email = self.tabview.add("📧 Email")

        self._criar_aba_logs(tab_logs)
        self._criar_aba_preview(tab_preview)
        self._criar_aba_email(tab_email)

    def _criar_aba_logs(self, parent):
        parent.grid_columnconfigure(0, weight=1)
        parent.grid_rowconfigure(0, weight=1)

        container = ctk.CTkFrame(parent, fg_color="transparent")
        container.grid(row=0, column=0, sticky="nsew", padx=3, pady=3)
        container.grid_columnconfigure(0, weight=1)
        container.grid_rowconfigure(0, weight=1)

        self.log_text = ctk.CTkTextbox(
            container, font=ctk.CTkFont(family="Consolas", size=11),
            fg_color=self.CORES['fundo_escuro'], corner_radius=6,
        )
        self.log_text.grid(row=0, column=0, sticky="nsew")

        self.log_text._textbox.tag_config("sucesso", foreground=self.CORES['sucesso'])
        self.log_text._textbox.tag_config("erro", foreground=self.CORES['erro'])
        self.log_text._textbox.tag_config("aviso", foreground=self.CORES['aviso'])
        self.log_text._textbox.tag_config("info", foreground=self.CORES['info'])
        self.log_text._textbox.tag_config("processando", foreground=self.CORES['processando'])

        btn_frame = ctk.CTkFrame(container, fg_color="transparent")
        btn_frame.grid(row=1, column=0, sticky="ew", pady=(5, 0))

        ctk.CTkButton(
            btn_frame, text="🗑 Limpar", command=self.limpar_logs,
            width=90, height=26, font=ctk.CTkFont(size=10),
            fg_color="#34495E", hover_color="#2C3E50",
        ).pack(side="left")

        ctk.CTkButton(
            btn_frame, text="💾 Exportar", command=self.exportar_logs,
            width=90, height=26, font=ctk.CTkFont(size=10),
            fg_color="#34495E", hover_color="#2C3E50",
        ).pack(side="left", padx=8)

    def _criar_aba_preview(self, parent):
        parent.grid_columnconfigure(0, weight=1)
        parent.grid_rowconfigure(1, weight=1)

        info_frame = ctk.CTkFrame(parent, fg_color="transparent")
        info_frame.grid(row=0, column=0, sticky="ew", padx=3, pady=3)

        self.preview_info_label = ctk.CTkLabel(
            info_frame, text="Nenhum arquivo carregado",
            font=ctk.CTkFont(size=11), text_color="#95A5A6",
        )
        self.preview_info_label.pack(side="left")

        ctk.CTkButton(
            info_frame, text="🔄 Recarregar", command=self.carregar_preview,
            width=85, height=24, font=ctk.CTkFont(size=10),
            fg_color="#34495E", hover_color="#2C3E50",
        ).pack(side="right")

        self.preview_text = ctk.CTkTextbox(
            parent, font=ctk.CTkFont(family="Consolas", size=10),
            fg_color=self.CORES['fundo_escuro'], corner_radius=6,
        )
        self.preview_text.grid(row=1, column=0, sticky="nsew", padx=3, pady=(0, 3))

    def _criar_aba_email(self, parent):
        parent.grid_columnconfigure(0, weight=1)

        frame = ctk.CTkFrame(parent, fg_color="transparent")
        frame.grid(row=0, column=0, sticky="nsew", padx=12, pady=10)
        frame.grid_columnconfigure(1, weight=1)

        # Período
        ctk.CTkLabel(frame, text="Período:", font=ctk.CTkFont(size=11), text_color="#BDC3C7").grid(
            row=0, column=0, sticky="w", pady=6, padx=(0, 10)
        )
        ctk.CTkEntry(
            frame, textvariable=self.periodo_var,
            width=120, height=32, font=ctk.CTkFont(size=11), justify="center",
            placeholder_text="MM/AAAA",
        ).grid(row=0, column=1, sticky="w")

        # Destinatário
        ctk.CTkLabel(frame, text="Destinatário:", font=ctk.CTkFont(size=11), text_color="#BDC3C7").grid(
            row=1, column=0, sticky="w", pady=6, padx=(0, 10)
        )
        self.email_dest_var = ctk.StringVar(value=os.getenv("EMAIL_DESTINATARIO", "sindembar@uol.com.br"))
        ctk.CTkEntry(
            frame, textvariable=self.email_dest_var,
            height=32, font=ctk.CTkFont(size=11),
        ).grid(row=1, column=1, sticky="ew")

        # Pasta de anexos (usa a pasta de destino já configurada)
        ctk.CTkLabel(frame, text="Anexos em:", font=ctk.CTkFont(size=11), text_color="#BDC3C7").grid(
            row=2, column=0, sticky="w", pady=6, padx=(0, 10)
        )
        self.email_pasta_label = ctk.CTkLabel(
            frame, textvariable=self.diretorio_salvamento,
            font=ctk.CTkFont(size=11), text_color="#95A5A6", anchor="w",
        )
        self.email_pasta_label.grid(row=2, column=1, sticky="ew")

        # Botão enviar
        ctk.CTkButton(
            frame, text="📧 Enviar e-mail", command=self._enviar_email_thread,
            height=36, font=ctk.CTkFont(size=12, weight="bold"),
            fg_color=self.CORES['info'], hover_color="#2980B9",
        ).grid(row=3, column=0, columnspan=2, pady=(16, 4), sticky="ew")

        # Status do envio
        self.email_status_var = ctk.StringVar(value="")
        ctk.CTkLabel(
            frame, textvariable=self.email_status_var,
            font=ctk.CTkFont(size=11), text_color="#95A5A6",
        ).grid(row=4, column=0, columnspan=2, sticky="w")

    # ── Ações da interface ────────────────────────────────────────────────────

    def selecionar_arquivo(self):
        filename = ctk.filedialog.askopenfilename(
            filetypes=[("Excel files", "*.xlsx *.xls")],
            title="Selecione o arquivo Excel",
        )
        if filename:
            self.arquivo_excel.set(filename)
            self.adicionar_log(f"Arquivo selecionado: {os.path.basename(filename)}", logging.INFO, "info")
            self.carregar_preview()

    def carregar_preview(self):
        if not self.arquivo_excel.get():
            return
        try:
            df = pd.read_excel(self.arquivo_excel.get())
            total = len(df)

            self.preview_info_label.configure(
                text=f"📄 {os.path.basename(self.arquivo_excel.get())} | {total} linhas | Colunas: {', '.join(str(c) for c in df.columns[:6])}"
            )

            self.preview_text.delete("1.0", "end")
            header = " | ".join([f"{str(col)[:15]:^15}" for col in df.columns[:6]])
            sep = "─" * len(header)
            self.preview_text.insert("end", f"{sep}\n{header}\n{sep}\n")
            for _, row in df.head(50).iterrows():
                self.preview_text.insert("end", " | ".join([f"{str(v)[:15]:^15}" for v in row.values[:6]]) + "\n")
            if total > 50:
                self.preview_text.insert("end", f"\n... e mais {total - 50} linhas\n")

            colunas_faltando = [c for c in ['Codigo', 'Nome'] if c not in df.columns]
            if colunas_faltando:
                self.adicionar_log(f"Colunas obrigatórias não encontradas: {', '.join(colunas_faltando)}", logging.WARNING, "aviso")
            else:
                self.adicionar_log(f"Preview carregado: {total} linhas. Colunas OK.", logging.INFO, "sucesso")

            self.df_carregado = df
        except Exception as e:
            self.adicionar_log(f"Erro ao carregar preview: {e}", logging.ERROR, "erro")

    def selecionar_diretorio(self):
        script = (
            "Add-Type -AssemblyName System.Windows.Forms; "
            "$f = New-Object System.Windows.Forms.FolderBrowserDialog; "
            "$f.Description = 'Selecione a pasta de destino'; "
            "$f.ShowNewFolderButton = $true; "
            "if ($f.ShowDialog() -eq 'OK') { $f.SelectedPath } else { '' }"
        )
        try:
            result = subprocess.run(
                ["powershell", "-Command", script],
                capture_output=True, text=True, timeout=120,
            )
            diretorio = result.stdout.strip()
            if diretorio:
                self.diretorio_salvamento.set(diretorio)
                self.adicionar_log(f"Pasta de destino: {diretorio}", logging.INFO, "info")
        except Exception as e:
            self.adicionar_log(f"Erro ao selecionar pasta: {e}", logging.ERROR, "erro")

    def limpar_logs(self):
        self.log_text.delete("1.0", "end")
        self.adicionar_log("Log limpo", logging.INFO, "info")

    def exportar_logs(self):
        try:
            filename = ctk.filedialog.asksaveasfilename(
                defaultextension=".txt",
                filetypes=[("Text files", "*.txt"), ("All files", "*.*")],
                initialfilename=f"logs_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt",
            )
            if filename:
                with open(filename, 'w', encoding='utf-8') as f:
                    f.write(self.log_text.get("1.0", "end"))
                self.adicionar_log(f"Logs exportados: {filename}", logging.INFO, "sucesso")
        except Exception as e:
            self.adicionar_log(f"Erro ao exportar logs: {e}", logging.ERROR, "erro")

    def adicionar_log(self, mensagem, level=logging.INFO, tag=None):
        try:
            timestamp = datetime.now().strftime('%H:%M:%S')
            if tag is None:
                if level >= logging.ERROR:
                    tag = "erro"
                elif level >= logging.WARNING:
                    tag = "aviso"
                elif any(w in mensagem.lower() for w in ("sucesso", "processad", "salvo", "concluíd")):
                    tag = "sucesso"
                else:
                    tag = "info"

            prefixos = {"sucesso": "✅", "erro": "❌", "aviso": "⚠️", "info": "ℹ️", "processando": "⏳"}
            prefixo = prefixos.get(tag, "•")

            self.log_text.insert("end", f"[{timestamp}] {prefixo} ", tag)
            self.log_text.insert("end", f"{mensagem}\n", tag)
            self.log_text.see("end")
            self.window.update_idletasks()
        except Exception:
            pass

    def atualizar_status_indicator(self, status):
        cores = {
            'aguardando': '#7F8C8D',
            'executando': self.CORES['sucesso'],
            'pausado': self.CORES['aviso'],
            'erro': self.CORES['erro'],
            'concluido': self.CORES['info'],
        }
        self.status_indicator.configure(fg_color=cores.get(status, '#7F8C8D'))

    def atualizar_tempo(self):
        if self.stats['tempo_inicio']:
            elapsed = datetime.now() - self.stats['tempo_inicio']
            h, rem = divmod(int(elapsed.total_seconds()), 3600)
            m, s = divmod(rem, 60)
            self.tempo_label.configure(text=f"{h:02d}:{m:02d}:{s:02d}")
            if self.executando:
                self.window.after(1000, self.atualizar_tempo)

    def validar_entrada(self):
        if not self.arquivo_excel.get():
            return False, "Selecione um arquivo Excel."
        if not os.path.exists(self.arquivo_excel.get()):
            return False, "Arquivo Excel não encontrado."
        try:
            linha_inicial = int(self.linha_inicial.get())
            if linha_inicial < 1:
                return False, "Linha inicial deve ser maior que 0."
        except ValueError:
            return False, "Linha inicial deve ser um número válido."
        try:
            df = pd.read_excel(self.arquivo_excel.get())
            if len(df) == 0:
                return False, "Arquivo Excel está vazio."
            if linha_inicial > len(df) + 1:
                return False, f"Linha inicial ({linha_inicial}) maior que o total de linhas ({len(df) + 1})."
            colunas_faltando = [c for c in ['Codigo', 'Nome', 'Rubrica'] if c not in df.columns]
            if colunas_faltando:
                return False, f"Colunas obrigatórias não encontradas: {', '.join(colunas_faltando)}"
        except Exception as e:
            return False, f"Erro ao ler arquivo Excel: {e}"
        if not self.diretorio_salvamento.get().strip():
            return False, "Selecione a pasta de destino antes de continuar."
        return True, "OK"

    def iniciar_automacao_thread(self):
        if self.executando:
            self.adicionar_log("Automação já em execução", logging.WARNING, "aviso")
            return

        valido, msg = self.validar_entrada()
        if not valido:
            self.adicionar_log(f"Erro de validação: {msg}", logging.ERROR, "erro")
            messagebox.showerror("Erro de Validação", msg)
            return

        self.linhas_processadas = 0
        self.linhas_com_erro = 0
        self.stats = {'processados': 0, 'sucesso': 0, 'erros': 0, 'tempo_inicio': datetime.now()}
        self.sucesso_label.configure(text="0")
        self.erros_label.configure(text="0")

        self.thread_automacao = threading.Thread(target=self.iniciar_automacao, daemon=True)
        self.thread_automacao.start()

        self.btn_iniciar.configure(state="disabled")
        self.btn_pausar.configure(state="normal")
        self.btn_parar.configure(state="normal")
        self.atualizar_status_indicator('executando')
        self.atualizar_tempo()

    def pausar_automacao(self):
        if self.executando:
            self.pausa_solicitada = not self.pausa_solicitada
            if self.pausa_solicitada:
                self.btn_pausar.configure(text="▶  Retomar")
                self.status_var.set("Pausado")
                self.atualizar_status_indicator('pausado')
                self.adicionar_log("Automação pausada", logging.INFO, "aviso")
            else:
                self.btn_pausar.configure(text="⏸  Pausar")
                self.status_var.set("Em execução...")
                self.atualizar_status_indicator('executando')
                self.adicionar_log("Automação retomada", logging.INFO, "info")

    def parar_automacao(self):
        if self.executando:
            self.executando = False
            self.pausa_solicitada = False
            self.adicionar_log("Solicitação de parada enviada...", logging.INFO, "aviso")
            self.status_var.set("Interrompendo...")
            self.atualizar_status_indicator('erro')

    def ao_fechar(self):
        if self.executando:
            if messagebox.askyesno("Confirmação", "Existe uma automação em execução. Deseja realmente sair?"):
                self.executando = False
                self.pausa_solicitada = False
                self.window.after(1000, self.window.destroy)
        else:
            self.window.destroy()

    # ── Loop principal de automação ───────────────────────────────────────────

    def iniciar_automacao(self):
        diretorio = self.diretorio_salvamento.get().strip()
        linha_inicial = int(self.linha_inicial.get())

        try:
            self.adicionar_log("Iniciando automação...", logging.INFO, "processando")
            self.status_var.set("Em execução...")
            self.executando = True

            os.makedirs(diretorio, exist_ok=True)

            df = pd.read_excel(self.arquivo_excel.get())
            inicio_idx = linha_inicial - 2
            df_processar = df.iloc[inicio_idx:]
            total = len(df_processar)

            self.adicionar_log(f"Arquivo carregado: {total} linhas para processar", logging.INFO, "info")

            automacao = DominioAutomation(self.logger, self)

            if not automacao.connect_to_dominio():
                self.adicionar_log("Não foi possível conectar ao Domínio", logging.ERROR, "erro")
                return

            for idx, (original_index, row) in enumerate(df_processar.iterrows()):
                if not self.executando:
                    self.adicionar_log("Automação interrompida pelo usuário", logging.INFO, "aviso")
                    break

                while self.pausa_solicitada and self.executando:
                    time.sleep(0.5)

                if not self.executando:
                    break

                linha_excel = original_index + 2
                codigo = str(row['Codigo'])
                nome = str(row.get('Nome', 'N/A'))
                rubrica = str(row['Rubrica'])

                self.empresa_label.configure(text=codigo[:20])
                self.status_var.set(f"Processando: {idx + 1}/{total}")
                self.adicionar_log(
                    f"Linha {linha_excel} - Codigo {codigo} - {nome} - Rubrica {rubrica}",
                    logging.INFO, "processando"
                )

                try:
                    success = automacao.processar_taxa_bares(rubrica, codigo, nome, diretorio, linha_excel)

                    if success:
                        self.linhas_processadas += 1
                        self.sucesso_label.configure(text=str(self.linhas_processadas))
                        self.success_logger.info(f"Linha {linha_excel} - Codigo {codigo} - {nome} - OK")
                        self.adicionar_log(f"Linha {linha_excel} processada com sucesso", logging.INFO, "sucesso")
                    else:
                        self.linhas_com_erro += 1
                        self.erros_label.configure(text=str(self.linhas_com_erro))
                        self.error_logger.error(f"Linha {linha_excel} - Codigo {codigo} - {nome} - ERRO")
                        self.adicionar_log(f"Erro na linha {linha_excel}", logging.ERROR, "erro")

                except Exception as e:
                    self.linhas_com_erro += 1
                    self.erros_label.configure(text=str(self.linhas_com_erro))
                    erro_msg = f"Linha {linha_excel} - Erro: {e}"
                    self.error_logger.error(erro_msg)
                    self.adicionar_log(erro_msg, logging.ERROR, "erro")

            if self.executando:
                self.status_var.set("Processamento concluído")
                self.atualizar_status_indicator('concluido')
                self.adicionar_log(
                    f"Automação concluída! {self.linhas_processadas} processadas, {self.linhas_com_erro} com erro.",
                    logging.INFO, "sucesso"
                )

                # Envio automático de e-mail com os PDFs gerados
                if self.linhas_processadas > 0:
                    self.adicionar_log("Iniciando envio automático de e-mail...", logging.INFO, "info")
                    periodo = self.periodo_var.get().strip()
                    destinatario = self.email_dest_var.get().strip()
                    anexos = [
                        os.path.join(diretorio, f)
                        for f in os.listdir(diretorio)
                        if f.lower().endswith(".pdf")
                    ]
                    if anexos and destinatario:
                        self._enviar_email(periodo, destinatario, anexos)
                    else:
                        self.adicionar_log("Nenhum PDF encontrado para enviar ou destinatário não configurado.", logging.WARNING, "aviso")

                if _requests_available:
                    try:
                        webhook_url = os.getenv("DISCORD_WEBHOOK_URL")
                        if webhook_url:
                            mensagem = (
                                f"📋 **Taxa Bares Emitida**\n\n"
                                f"📊 **Processadas:** {self.linhas_processadas}\n"
                                f"📂 **Destino:** `{diretorio}`\n\n"
                                f"✅ Emissão finalizada!\n\n"
                                f"<@&1299044385899548752>"
                            )
                            requests.post(webhook_url, json={"content": mensagem}, timeout=10)
                            self.adicionar_log("Notificação enviada ao Discord", logging.INFO, "sucesso")
                    except Exception as e:
                        self.adicionar_log(f"Erro ao enviar notificação: {e}", logging.WARNING, "aviso")

        except Exception as e:
            erro_msg = f"Erro crítico: {e}"
            self.error_logger.error(erro_msg)
            self.adicionar_log(erro_msg, logging.ERROR, "erro")
            self.status_var.set("Erro no processamento")
            self.atualizar_status_indicator('erro')
        finally:
            self.executando = False
            self.pausa_solicitada = False
            self.atualizar_tempo()
            self.btn_iniciar.configure(state="normal")
            self.btn_pausar.configure(state="disabled", text="⏸ Pausar")
            self.btn_parar.configure(state="disabled")

    def _enviar_email_thread(self):
        periodo = self.periodo_var.get().strip()
        destinatario = self.email_dest_var.get().strip()
        pasta = self.diretorio_salvamento.get().strip()

        if not periodo:
            messagebox.showerror("Erro", "Informe o período antes de enviar.")
            return
        if not destinatario:
            messagebox.showerror("Erro", "Informe o destinatário.")
            return
        if not pasta or not os.path.isdir(pasta):
            messagebox.showerror("Erro", "A pasta de destino não foi configurada ou não existe.")
            return

        anexos = [
            os.path.join(pasta, f)
            for f in os.listdir(pasta)
            if f.lower().endswith(".pdf")
        ]
        if not anexos:
            messagebox.showerror("Erro", f"Nenhum PDF encontrado em:\n{pasta}")
            return

        self.email_status_var.set("Enviando...")
        threading.Thread(
            target=self._enviar_email,
            args=(periodo, destinatario, anexos),
            daemon=True,
        ).start()

    def _enviar_email(self, periodo: str, destinatario: str, anexos: list):
        try:
            gmail_user = os.getenv("GMAIL_USER", "")
            gmail_password = os.getenv("GMAIL_APP_PASSWORD", "")

            if not gmail_user or not gmail_password:
                self.window.after(0, lambda: self.email_status_var.set(
                    "❌ Credenciais não configuradas no .env"
                ))
                self.window.after(0, lambda: messagebox.showerror(
                    "Erro de configuração",
                    "Preencha GMAIL_USER e GMAIL_APP_PASSWORD no arquivo .env"
                ))
                return

            assunto = f"Taxa Assistencial {periodo}"

            corpo = f"""\
<html>
<body style="font-family: Arial, sans-serif; font-size: 14px; color: #333;">
  <p>Prezados,</p>

  <p>
    Encaminhamos em anexo as relações referentes à <strong>Taxa Assistencial</strong>
    do período <strong>{periodo}</strong>, conforme apuração realizada pelo sistema Domínio Folha.
  </p>

  <p>Pedimos que nos enviem a taxa conforme relações em anexo.</p>

  <p>Ficamos à disposição para quaisquer esclarecimentos.</p>

  <br>
  <p>Atenciosamente,<br>
  <strong>Canella e Santos Contabilidade</strong></p>
</body>
</html>"""

            msg = MIMEMultipart()
            msg["From"] = gmail_user
            msg["To"] = destinatario
            msg["Subject"] = assunto
            msg.attach(MIMEText(corpo, "html", "utf-8"))

            for caminho in anexos:
                nome = os.path.basename(caminho)
                with open(caminho, "rb") as f:
                    parte = MIMEBase("application", "octet-stream")
                    parte.set_payload(f.read())
                encoders.encode_base64(parte)
                parte.add_header("Content-Disposition", f'attachment; filename="{nome}"')
                msg.attach(parte)

            context = ssl.create_default_context()
            with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=context) as server:
                server.login(gmail_user, gmail_password.replace(" ", ""))
                server.sendmail(gmail_user, destinatario, msg.as_string())

            total = len(anexos)
            self.window.after(0, lambda: self.email_status_var.set(
                f"✅ E-mail enviado com {total} anexo(s) para {destinatario}"
            ))
            self.adicionar_log(
                f"E-mail enviado — {total} PDF(s) para {destinatario} | Período: {periodo}",
                logging.INFO, "sucesso"
            )

        except smtplib.SMTPAuthenticationError:
            msg_err = "❌ Falha de autenticação. Verifique GMAIL_USER e GMAIL_APP_PASSWORD no .env"
            self.window.after(0, lambda: self.email_status_var.set(msg_err))
            self.window.after(0, lambda: messagebox.showerror("Erro de autenticação", msg_err))
        except Exception as e:
            msg_err = f"❌ Erro ao enviar: {e}"
            self.window.after(0, lambda: self.email_status_var.set(msg_err))
            self.adicionar_log(msg_err, logging.ERROR, "erro")

    def executar(self):
        self.window.mainloop()


# ── Classe de automação Domínio ───────────────────────────────────────────────

class DominioAutomation:
    def __init__(self, logger, gui):
        timings.Timings.window_find_timeout = 20
        self.app = None
        self.main_window = None
        self.logger = logger
        self.gui = gui
        self.empresa_atual = None

    def log(self, message):
        self.logger.info(message)

    def should_stop(self):
        return not self.gui.executando

    def check_pause(self):
        while self.gui.pausa_solicitada and self.gui.executando:
            time.sleep(0.5)

    def smart_sleep(self, seconds: float) -> bool:
        interval = 0.15
        elapsed = 0.0
        while elapsed < seconds:
            if self.should_stop():
                return False
            self.check_pause()
            if self.should_stop():
                return False
            t = min(interval, seconds - elapsed)
            time.sleep(t)
            elapsed += t
        return True

    def wait_for_condition(self, condition_fn, timeout=30.0, poll_interval=0.15, description="") -> bool:
        start = time.time()
        while time.time() - start < timeout:
            if self.should_stop():
                return False
            self.check_pause()
            try:
                if condition_fn():
                    if description:
                        self.log(f"{description} - concluído em {time.time() - start:.1f}s")
                    return True
            except Exception:
                pass
            time.sleep(poll_interval)
        if description:
            self.log(f"{description} - timeout após {timeout}s")
        return False

    def _window_exists(self, title: str, class_name: str) -> bool:
        try:
            result = [False]
            def cb(hwnd, _):
                if win32gui.IsWindowVisible(hwnd):
                    try:
                        if win32gui.GetWindowText(hwnd) == title and win32gui.GetClassName(hwnd) == class_name:
                            result[0] = True
                            return False
                    except Exception:
                        pass
                return True
            win32gui.EnumWindows(cb, None)
            return result[0]
        except Exception:
            return False

    def _save_dialog_exists(self) -> bool:
        try:
            result = [False]
            def cb(hwnd, _):
                if win32gui.IsWindowVisible(hwnd):
                    try:
                        if win32gui.GetClassName(hwnd) == "#32770":
                            child = win32gui.FindWindowEx(hwnd, 0, "Static", "Salvar em:")
                            if child:
                                result[0] = True
                                return False
                    except Exception:
                        pass
                return True
            win32gui.EnumWindows(cb, None)
            return result[0]
        except Exception:
            return False

    def _find_save_window_hwnd(self) -> int:
        result = [0]
        def cb(hwnd, _):
            if win32gui.IsWindowVisible(hwnd):
                try:
                    if win32gui.GetClassName(hwnd) == "#32770":
                        if win32gui.GetWindowText(hwnd) == "Salvar em PDF":
                            result[0] = hwnd
                            return False
                        child = win32gui.FindWindowEx(hwnd, 0, "Static", "Salvar em:")
                        if child:
                            result[0] = hwnd
                            return False
                except Exception:
                    pass
            return True
        try:
            win32gui.EnumWindows(cb, None)
        except Exception:
            pass
        return result[0]

    def _any_error_dialog_visible(self) -> bool:
        keywords = ("erro", "aviso", "atenção", "alerta", "warning", "error", "informação")
        try:
            result = [False]
            def cb(hwnd, _):
                if win32gui.IsWindowVisible(hwnd):
                    try:
                        if win32gui.GetClassName(hwnd) == "#32770":
                            title = win32gui.GetWindowText(hwnd).lower()
                            if any(k in title for k in keywords):
                                result[0] = True
                                return False
                    except Exception:
                        pass
                return True
            win32gui.EnumWindows(cb, None)
            return result[0]
        except Exception:
            return False

    def _is_connection_alive(self) -> bool:
        if self.app is None or self.main_window is None:
            return False
        try:
            hwnd = self.main_window.handle
            if not win32gui.IsWindow(hwnd):
                return False
            win32gui.GetWindowText(hwnd)
            return True
        except Exception:
            return False

    def find_dominio_window(self) -> Optional[int]:
        try:
            self.log("🔍 Procurando janela do Domínio Folha...")
            try:
                all_windows = findwindows.find_windows()
                for hwnd in all_windows:
                    try:
                        title = win32gui.GetWindowText(hwnd)
                        if "Domínio" in title and "Folha" in title:
                            self.log("✅ Janela do Domínio Folha localizada!")
                            return hwnd
                    except Exception:
                        continue
            except Exception as e:
                self.log(f"⚠️ Erro ao listar janelas: {e}")

            windows = findwindows.find_windows(title_re=".*Domínio Folha.*")
            if windows:
                return windows[0]

            self.log("❌ Nenhuma janela do Domínio Folha encontrada")
            return None
        except Exception as e:
            self.log(f"❌ Erro ao procurar janela: {e}")
            return None

    def connect_to_dominio(self) -> bool:
        try:
            handle = self.find_dominio_window()
            if not handle:
                return False

            if win32gui.IsIconic(handle):
                win32gui.ShowWindow(handle, win32con.SW_RESTORE)
                time.sleep(1)

            win32gui.SetForegroundWindow(handle)
            time.sleep(0.5)

            self.app = Application(backend="uia").connect(handle=handle)
            self.main_window = self.app.window(handle=handle)

            self.log("✅ Conectado ao Domínio Folha com sucesso")
            return True
        except Exception as e:
            self.log(f"❌ Erro ao conectar ao Domínio: {e}")
            return False

    def handle_error_dialogs(self) -> bool:
        """Trata diálogos de erro. Retorna True para continuar, False para abortar."""
        try:
            error_titles = {"erro", "erro léxico", "aviso", "atenção", "informação", "alerta", "warning", "error"}
            found_hwnd = None
            found_title = None

            def cb(hwnd, _):
                nonlocal found_hwnd, found_title
                if not win32gui.IsWindowVisible(hwnd):
                    return True
                try:
                    title = win32gui.GetWindowText(hwnd)
                    if title and win32gui.GetClassName(hwnd) == "#32770":
                        t = title.strip().lower()
                        if any(t == e or e in t for e in error_titles):
                            found_hwnd = hwnd
                            found_title = title
                            return False
                except Exception:
                    pass
                return True

            win32gui.EnumWindows(cb, None)

            if found_hwnd is None:
                return True

            try:
                error_window = self.app.window(handle=found_hwnd)
            except Exception:
                win32gui.SetForegroundWindow(found_hwnd)
                send_keys('{ENTER}')
                time.sleep(0.3)
                return True

            message = ""
            try:
                message = error_window.window_text()
                for static in error_window.children(class_name="Static"):
                    text = static.window_text()
                    if text:
                        message += " " + text
            except Exception:
                pass

            self.log(f"⚠️ Diálogo: '{found_title}' - {message[:100]}")

            mensagens_continuar = ["erro na gravação do relatório", "nome do caminho inválido", "caracteres não permitidos"]
            mensagens_abortar = ["sem dados para emitir", "nenhum registro encontrado", "não há dados", "registro não encontrado"]

            msg_lower = message.lower()
            for m in mensagens_continuar:
                if m in msg_lower:
                    self.log(f"⚠️ Aviso não crítico: {m}")
                    error_window.set_focus()
                    send_keys('{ENTER}')
                    time.sleep(0.5)
                    return True

            for m in mensagens_abortar:
                if m in msg_lower:
                    self.log(f"⚠️ Sem dados: {m}")
                    error_window.set_focus()
                    send_keys('{ENTER}')
                    time.sleep(0.5)
                    for _ in range(4):
                        send_keys('{ESC}')
                        time.sleep(0.5)
                    return False

            if "léxico" in found_title.lower():
                error_window.set_focus()
                for _ in range(3):
                    send_keys('{ESC}')
                    time.sleep(0.5)
                return True

            error_window.set_focus()
            time.sleep(0.2)
            try:
                ok_btn = error_window.child_window(title="OK", class_name="Button")
                if ok_btn.exists():
                    ok_btn.click_input()
                    time.sleep(0.5)
                    if found_title.lower() in ("erro", "aviso"):
                        return False
                    return True
            except Exception:
                pass

            send_keys('{ENTER}')
            time.sleep(0.5)
            if found_title.lower() in ("erro", "aviso"):
                return False
            return True

        except Exception as e:
            self.log(f"⚠️ Exceção em handle_error_dialogs: {e}")
            return True

    def cleanup_windows(self):
        try:
            self.log("🧹 Limpando janelas")
            self.main_window.set_focus()
            for _ in range(4):
                send_keys('{ESC}')
                time.sleep(0.5)
        except Exception as e:
            self.log(f"⚠️ Erro durante limpeza: {e}")

    def salvar_pdf(self, nome_arquivo: str, diretorio: str) -> bool:
        try:
            if self.should_stop():
                return False

            self.log("💾 Configurando salvamento do PDF")

            save_hwnd = self._find_save_window_hwnd()
            if save_hwnd:
                save_app = Application(backend="uia").connect(handle=save_hwnd)
                save_window = save_app.window(handle=save_hwnd)
            else:
                self.log("🔍 Procurando janela de salvamento alternativa...")
                save_window = self.main_window.child_window(title="Salvar em PDF", class_name="#32770")
                if not save_window.exists():
                    self.log("❌ Janela de salvamento não encontrada")
                    return False

            if self.should_stop():
                return False
            self.check_pause()

            caminho_completo = os.path.join(diretorio, nome_arquivo) if diretorio else nome_arquivo
            self.log(f"📝 Salvando em: {caminho_completo}")

            time.sleep(0.5)
            win32gui.SetForegroundWindow(save_hwnd)
            time.sleep(0.2)
            send_keys(caminho_completo, with_spaces=True)
            time.sleep(0.3)

            if self.should_stop():
                return False
            self.check_pause()

            self.log("💾 Salvando PDF")
            send_keys('{ENTER}')

            if not self.wait_for_condition(
                lambda: not win32gui.IsWindow(save_hwnd) or not win32gui.IsWindowVisible(save_hwnd),
                timeout=30, poll_interval=0.3,
                description="Aguardando salvamento do PDF",
            ):
                self.log("⚠️ Timeout aguardando janela fechar, continuando mesmo assim")

            self.log(f"✅ PDF salvo: {nome_arquivo}")
            self.cleanup_windows()
            return True

        except Exception as e:
            self.log(f"❌ Erro ao salvar PDF: {e}")
            return False

    def handle_empresa_change(self, empresa_num: str) -> bool:
        try:
            if self.should_stop():
                return False

            self.log("📞 Solicitando troca de empresa (F8)")
            self.main_window.set_focus()
            if not self.smart_sleep(0.3):
                return False
            send_keys('{F8}')
            if not self.smart_sleep(2):
                return False

            troca_window = None
            for attempt in range(10):
                if self.should_stop():
                    return False
                self.check_pause()
                try:
                    troca_window = self.main_window.child_window(
                        title="Troca de empresas", class_name="FNWND3190"
                    )
                    if troca_window.exists():
                        break
                    if not self.handle_error_dialogs():
                        self.cleanup_windows()
                        return False
                    if not self.smart_sleep(0.5):
                        return False
                except Exception:
                    if attempt == 9:
                        self.log("❌ Janela 'Troca de empresas' não encontrada (timeout)")
                        return False
                    if not self.smart_sleep(1):
                        return False

            if not troca_window:
                self.log("❌ Janela 'Troca de empresas' não encontrada")
                return False

            self.log(f"🏢 Alterando para empresa: {empresa_num}")
            troca_window.set_focus()
            if not self.smart_sleep(0.3):
                return False
            send_keys(empresa_num)
            if not self.smart_sleep(0.5):
                return False
            send_keys('{ENTER}')
            if not self.smart_sleep(1.5):
                return False

            if not self.handle_error_dialogs():
                self.cleanup_windows()
                return False

            start = time.time()
            while time.time() - start < 30:
                if self.should_stop():
                    return False
                self.check_pause()
                try:
                    if not troca_window.exists() or not troca_window.is_visible():
                        break
                except Exception:
                    break
                self.handle_error_dialogs()
                time.sleep(0.15)

            self.close_avisos_vencimento()
            return True

        except Exception as e:
            self.log(f"❌ Erro na troca de empresa: {e}")
            return False

    def close_avisos_vencimento(self):
        try:
            aviso_window = self.main_window.child_window(
                title="Avisos de Vencimento", class_name="FNWND3190"
            )
            if aviso_window.exists() and aviso_window.is_visible():
                self.log("📋 Fechando 'Avisos de Vencimento'")
                aviso_window.set_focus()
                send_keys('{ESC}')
                time.sleep(0.5)
                send_keys('{ESC}')
                time.sleep(0.5)
        except Exception:
            pass

    def processar_taxa_bares(self, rubrica: str, codigo: str, nome: str, diretorio: str, linha_excel: int) -> bool:
        """Abre Relatórios > F > M no Domínio e salva o PDF resultante."""
        try:
            if not self._is_connection_alive():
                handle = self.find_dominio_window()
                if not handle:
                    self.log("❌ Janela do Domínio não localizada")
                    return False
                self.app = Application(backend="uia").connect(handle=handle)
                self.main_window = self.app.window(handle=handle)

            handle = self.main_window.handle
            if win32gui.IsIconic(handle):
                win32gui.ShowWindow(handle, win32con.SW_RESTORE)
                if not self.smart_sleep(0.5):
                    return False

            win32gui.SetForegroundWindow(handle)
            time.sleep(0.2)

            self.log(f"👤 Linha {linha_excel} - Código: {codigo} | Nome: {nome} | Rubrica: {rubrica}")

            # Troca de empresa (só troca se for diferente da atual)
            if codigo != self.empresa_atual:
                if not self.handle_empresa_change(codigo):
                    return False
                self.empresa_atual = codigo
            else:
                self.log(f"🏢 Empresa {codigo} já selecionada, pulando troca")

            if self.should_stop():
                return False
            self.check_pause()

            # Abrir menu Relatórios (ALT+R) → F → M
            self.log("📊 Acessando Relatórios > F > M")
            self.main_window.set_focus()
            send_keys('%r')
            if not self.smart_sleep(0.5):
                return False
            send_keys('f')
            if not self.smart_sleep(0.5):
                return False
            send_keys('m')
            if not self.smart_sleep(1):
                return False

            # Aguardar a janela do relatório (FNWND3190) e clicar em Rubricas via Alt+U
            self.log("🔘 Aguardando janela do relatório para clicar em Rubricas...")

            def _find_fnwnd():
                hwnds = []
                def cb(hwnd, _):
                    if win32gui.IsWindowVisible(hwnd):
                        try:
                            if win32gui.GetClassName(hwnd) == "FNWND3190":
                                btn = win32gui.FindWindowEx(hwnd, 0, "Button", None)
                                if btn:
                                    hwnds.append(hwnd)
                                    return False
                        except Exception:
                            pass
                    return True
                win32gui.EnumWindows(cb, None)
                return hwnds[0] if hwnds else None

            rel_hwnd = None
            for _ in range(20):
                if self.should_stop():
                    return False
                self.check_pause()
                rel_hwnd = _find_fnwnd()
                if rel_hwnd:
                    break
                if not self.smart_sleep(0.5):
                    return False

            if not rel_hwnd:
                self.log("❌ Janela do relatório não encontrada")
                return False

            # Focar a janela e usar o atalho Alt+U do botão Rubricas
            win32gui.SetForegroundWindow(rel_hwnd)
            if not self.smart_sleep(0.3):
                return False
            self.log("🔘 Clicando no botão Rubricas (Alt+U)")
            send_keys('%u')
            if not self.smart_sleep(1):
                return False

            # Aguardar a janela de seleção de Rubrica — busca o checkbox "Rubrica:" em todos os níveis
            self.log("☑️ Verificando checkbox Rubrica...")

            def _find_rubrica_checkbox_hwnd():
                """Varre todas as janelas e seus filhos procurando Button com texto 'Rubrica:'."""
                result = [0]

                def enum_children(parent_hwnd):
                    children = []
                    try:
                        win32gui.EnumChildWindows(
                            parent_hwnd,
                            lambda h, _: children.append(h) or True,
                            None
                        )
                    except Exception:
                        pass
                    for h in children:
                        if result[0]:
                            break
                        try:
                            if (win32gui.IsWindowVisible(h)
                                    and win32gui.GetClassName(h) == "Button"
                                    and win32gui.GetWindowText(h) == "Rubrica:"):
                                result[0] = h
                                return
                        except Exception:
                            pass

                def cb(hwnd, _):
                    if result[0]:
                        return False
                    if win32gui.IsWindowVisible(hwnd):
                        enum_children(hwnd)
                    return True

                win32gui.EnumWindows(cb, None)
                return result[0]

            cb_hwnd = None
            for _ in range(20):
                if self.should_stop():
                    return False
                self.check_pause()
                cb_hwnd = _find_rubrica_checkbox_hwnd()
                if cb_hwnd:
                    break
                if not self.smart_sleep(0.5):
                    return False

            if not cb_hwnd:
                self.log("❌ Janela de seleção de Rubrica não encontrada")
                return False

            # Operar na janela de seleção via win32gui direto (pid diferente do processo principal)
            sel_hwnd = win32gui.GetParent(cb_hwnd)

            # Verificar/marcar checkbox Rubrica via BM_GETCHECK / BM_SETCHECK + clique
            import win32api
            BM_GETCHECK = 0x00F0
            checked = win32api.SendMessage(cb_hwnd, BM_GETCHECK, 0, 0)
            if not checked:
                self.log("☑️ Marcando checkbox Rubrica")
                win32gui.SetForegroundWindow(sel_hwnd)
                time.sleep(0.2)
                win32api.SendMessage(cb_hwnd, 0x00F5, 0, 0)  # BM_CLICK
                if not self.smart_sleep(0.3):
                    return False
            else:
                self.log("✅ Checkbox Rubrica já está marcado")

            # Preencher campo Edit (auto_id="1001") via WM_SETTEXT
            self.log(f"✏️ Preenchendo rubrica: {rubrica}")
            edit_hwnd = win32gui.FindWindowEx(sel_hwnd, 0, "Edit", None)
            if not edit_hwnd:
                self.log("❌ Campo de edição da rubrica não encontrado")
                return False
            win32gui.SetForegroundWindow(sel_hwnd)
            time.sleep(0.1)
            win32gui.SendMessage(edit_hwnd, win32con.WM_SETTEXT, 0, rubrica)
            if not self.smart_sleep(0.3):
                return False

            # OK da janela de seleção — foca e envia Alt+O
            self.log("✅ Clicando em OK (seleção de rubrica)")
            win32gui.SetForegroundWindow(sel_hwnd)
            if not self.smart_sleep(0.3):
                return False
            send_keys('%o')
            if not self.smart_sleep(0.5):
                return False

            # OK da janela do relatório — foca e envia Alt+O
            self.log("✅ Clicando em OK (confirmar relatório)")
            win32gui.SetForegroundWindow(rel_hwnd)
            if not self.smart_sleep(0.3):
                return False
            send_keys('%o')
            if not self.smart_sleep(0.5):
                return False

            # Aguardar janela de salvamento
            self.log("📄 Aguardando janela de salvamento...")
            timeout_total = 180
            intervalo_ctrl_d = 5
            inicio = time.time()
            janela_encontrada = False

            while time.time() - inicio < timeout_total:
                if self.should_stop():
                    return False

                if self._any_error_dialog_visible():
                    deve_continuar = self.handle_error_dialogs()
                    if deve_continuar:
                        time.sleep(0.5)
                        continue
                    else:
                        self.cleanup_windows()
                        return False

                if self._window_exists("Salvar em PDF", "#32770") or self._save_dialog_exists():
                    janela_encontrada = True
                    break

                send_keys('^d')
                for _ in range(int(intervalo_ctrl_d / 0.25)):
                    if self.should_stop():
                        return False
                    if self._any_error_dialog_visible():
                        break
                    if self._window_exists("Salvar em PDF", "#32770") or self._save_dialog_exists():
                        janela_encontrada = True
                        break
                    time.sleep(0.25)
                if janela_encontrada:
                    break

            if not janela_encontrada:
                self.log("❌ Timeout aguardando janela de salvamento")
                return False

            if not self.handle_error_dialogs():
                self.cleanup_windows()
                return False

            nome_arquivo = f"taxa_bares_{codigo}_{rubrica}.pdf"
            return self.salvar_pdf(nome_arquivo, diretorio)

        except Exception as e:
            self.log(f"❌ Erro ao processar linha {linha_excel}: {e}\n{traceback.format_exc()}")
            return False


def main():
    try:
        gui = AutomacaoGUI()
        gui.executar()
    except Exception as e:
        print(f"Erro crítico: {e}")
        traceback.print_exc()


if __name__ == "__main__":
    main()
