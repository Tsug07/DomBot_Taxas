"""
DomBot - Taxa Motorista
=======================
Automação RPA para emissão de relatórios de Taxa Motorista no sistema Domínio Folha.

Planilha esperada:
  Coluna A - Codigo da empresa
  Coluna B - Nome da empresa
  Coluna C - Rubrica
  Coluna D - Caminho completo do arquivo (pasta + nome)

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
import ctypes

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
        self.window.title("DomBot - Taxa Motorista v1.0")
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

        self.arquivo_excel = ctk.StringVar()
        self.linha_inicial = ctk.StringVar(value="2")
        self.competencia_var = ctk.StringVar(value=datetime.now().strftime("%m/%Y"))
        self.status_var = ctk.StringVar(value="Aguardando início...")

        self.df_carregado = None
        self.linhas_processadas = 0
        self.linhas_com_erro = 0

        self.logger = logging.getLogger('AutomacaoTaxaMotorista')
        self.logger.setLevel(logging.INFO)
        self.logger.handlers = []

        self.gui_handler = GUILogHandler(self)
        self.gui_handler.setFormatter(logging.Formatter('%(message)s'))
        self.logger.addHandler(self.gui_handler)

        self.criar_interface()

    def setup_file_logging(self):
        data_atual = datetime.now().strftime("%Y-%m-%d")

        self.success_logger = logging.getLogger('SuccessLogTaxaMotorista')
        self.success_logger.setLevel(logging.INFO)
        if not self.success_logger.handlers:
            h = logging.FileHandler(
                os.path.join(self.logs_dir, f'success_{data_atual}.log'), encoding='utf-8'
            )
            h.setFormatter(logging.Formatter('%(asctime)s - %(message)s', '%Y-%m-%d %H:%M:%S'))
            self.success_logger.addHandler(h)

        self.error_logger = logging.getLogger('ErrorLogTaxaMotorista')
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
            text="DomBot - Taxa Motorista",
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

        # --- Linha 2: Competência ---
        row2 = ctk.CTkFrame(config_frame, fg_color="transparent")
        row2.grid(row=1, column=0, sticky="ew", padx=10, pady=(0, 8))

        ctk.CTkLabel(row2, text="📅", font=ctk.CTkFont(size=14)).grid(row=0, column=0, padx=(0, 5))
        ctk.CTkLabel(row2, text="Competência:", font=ctk.CTkFont(size=11), text_color="#BDC3C7").grid(
            row=0, column=1, padx=(0, 5)
        )
        ctk.CTkEntry(
            row2, textvariable=self.competencia_var,
            width=120, height=32, font=ctk.CTkFont(size=11), justify="center",
            placeholder_text="MM/AAAA",
        ).grid(row=0, column=2)

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

        self._criar_aba_logs(tab_logs)
        self._criar_aba_preview(tab_preview)

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
            df = pd.read_excel(self.arquivo_excel.get(), header=None,
                               names=['Codigo', 'Nome', 'Rubrica', 'Caminho'])
            # Pular linha de cabeçalho se existir
            if str(df.iloc[0]['Codigo']).lower() in ('codigo', 'código', 'code'):
                df = df.iloc[1:].reset_index(drop=True)
            total = len(df)

            self.preview_info_label.configure(
                text=f"📄 {os.path.basename(self.arquivo_excel.get())} | {total} linhas | Colunas: Codigo, Nome, Rubrica, Caminho"
            )

            self.preview_text.delete("1.0", "end")
            cols = ['Codigo', 'Nome', 'Rubrica', 'Caminho']
            header = " | ".join([f"{c:^18}" for c in cols])
            sep = "─" * len(header)
            self.preview_text.insert("end", f"{sep}\n{header}\n{sep}\n")
            for _, row in df.head(50).iterrows():
                self.preview_text.insert("end", " | ".join([f"{str(row[c])[:18]:^18}" for c in cols]) + "\n")
            if total > 50:
                self.preview_text.insert("end", f"\n... e mais {total - 50} linhas\n")

            self.adicionar_log(f"Preview carregado: {total} linhas.", logging.INFO, "sucesso")
            self.df_carregado = df
        except Exception as e:
            self.adicionar_log(f"Erro ao carregar preview: {e}", logging.ERROR, "erro")

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
            df = pd.read_excel(self.arquivo_excel.get(), header=None,
                               names=['Codigo', 'Nome', 'Rubrica', 'Caminho'])
            if len(df) == 0:
                return False, "Arquivo Excel está vazio."
        except Exception as e:
            return False, f"Erro ao ler arquivo Excel: {e}"
        if not self.competencia_var.get().strip():
            return False, "Informe a competência antes de continuar."
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
        linha_inicial = int(self.linha_inicial.get())
        competencia = self.competencia_var.get().strip()

        try:
            self.adicionar_log("Iniciando automação...", logging.INFO, "processando")
            self.status_var.set("Em execução...")
            self.executando = True

            df = pd.read_excel(self.arquivo_excel.get(), header=None,
                               names=['Codigo', 'Nome', 'Rubrica', 'Caminho'])

            # Pular linha de cabeçalho se existir
            if str(df.iloc[0]['Codigo']).lower() in ('codigo', 'código', 'code'):
                df = df.iloc[1:].reset_index(drop=True)

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
                codigo = str(row['Codigo']).strip()
                nome = str(row.get('Nome', 'N/A')).strip()
                rubrica = str(row['Rubrica']).strip()
                caminho_raw = str(row['Caminho']).strip()
                # Sanitiza apenas o nome do arquivo (preserva o caminho de pasta)
                pasta = os.path.dirname(caminho_raw)
                nome_arq = os.path.basename(caminho_raw)
                for ch in r'*?"<>|/':
                    nome_arq = nome_arq.replace(ch, '-')
                caminho = os.path.join(pasta, nome_arq) if pasta else nome_arq

                self.empresa_label.configure(text=codigo[:20])
                self.status_var.set(f"Processando: {idx + 1}/{total}")
                self.adicionar_log(
                    f"Linha {linha_excel} - Codigo {codigo} - {nome} - Rubrica {rubrica}",
                    logging.INFO, "processando"
                )

                try:
                    success = automacao.processar_taxa_motorista(
                        rubrica, codigo, nome, caminho, competencia, linha_excel
                    )

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

    def _set_clipboard(self, text: str):
        """Coloca texto no clipboard do Windows via win32clipboard."""
        import win32clipboard
        win32clipboard.OpenClipboard()
        try:
            win32clipboard.EmptyClipboard()
            win32clipboard.SetClipboardText(text, win32clipboard.CF_UNICODETEXT)
        finally:
            win32clipboard.CloseClipboard()

    def _force_focus(self, hwnd: int):
        """Força foco para uma janela contornando a restrição do Windows 10."""
        try:
            ctypes.windll.user32.keybd_event(0x12, 0, 0, 0)
            ctypes.windll.user32.keybd_event(0x12, 0, 2, 0)
            win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
            win32gui.SetForegroundWindow(hwnd)
            time.sleep(0.3)
        except Exception:
            pass

    def salvar_pdf(self, caminho_completo: str) -> bool:
        """Salva o PDF usando o caminho completo (coluna D) via WM_SETTEXT no campo Edit."""
        try:
            if self.should_stop():
                return False

            self.log("💾 Configurando salvamento do PDF")

            save_hwnd = self._find_save_window_hwnd()
            if not save_hwnd:
                self.log("🔍 Procurando janela de salvamento alternativa...")
                save_window = self.main_window.child_window(title="Salvar em PDF", class_name="#32770")
                if not save_window.exists():
                    self.log("❌ Janela de salvamento não encontrada")
                    return False
                save_hwnd = save_window.handle

            if self.should_stop():
                return False
            self.check_pause()

            self.log(f"📝 Salvando em: {caminho_completo}")
            time.sleep(0.5)

            # Cola o caminho via clipboard para garantir que espaços, acentos e
            # caracteres especiais cheguem intactos ao campo de nome do arquivo
            self._force_focus(save_hwnd)
            self._set_clipboard(caminho_completo)

            # Clica diretamente no campo Edit de nome para garantir foco
            filename_edit = win32gui.FindWindowEx(save_hwnd, 0, "Edit", None)
            if filename_edit:
                win32gui.SetFocus(filename_edit)
                time.sleep(0.1)

            send_keys('^a')   # seleciona tudo que estiver no campo
            time.sleep(0.1)
            send_keys('^v')   # cola o caminho
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

            self.log(f"✅ PDF salvo: {os.path.basename(caminho_completo)}")
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

    def processar_taxa_motorista(self, rubrica: str, codigo: str, nome: str,
                                  caminho: str, competencia: str, linha_excel: int) -> bool:
        """
        Abre Relatórios > Gerenciador > pasta Diversos (d x6 + Enter) > 'c' x9 para
        selecionar a guia Taxa Motorista, preenche competência e rubrica, executa e salva.
        """
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

            # Abrir Gerenciador de Relatórios via ALT+R → i → i → ENTER
            self.log("📊 Acessando Gerenciador de Relatórios")
            self.main_window.set_focus()
            send_keys('%r')
            if not self.smart_sleep(0.5):
                return False
            send_keys('i')
            if not self.smart_sleep(0.5):
                return False
            send_keys('i')
            if not self.smart_sleep(0.5):
                return False
            send_keys('{ENTER}')
            if not self.smart_sleep(1):
                return False

            # Aguardar o Gerenciador de Relatórios
            relatorio_window = None
            for attempt in range(15):
                if self.should_stop():
                    return False
                self.check_pause()
                try:
                    relatorio_window = self.main_window.child_window(
                        title="Gerenciador de Relatórios",
                        class_name="FNWND3190"
                    )
                    if relatorio_window.exists():
                        break
                    if not self.handle_error_dialogs():
                        self.cleanup_windows()
                        return False
                    if not self.smart_sleep(1):
                        return False
                except Exception:
                    if attempt == 14:
                        self.log("❌ Gerenciador de Relatórios não encontrado (timeout)")
                        return False

            if not relatorio_window:
                self.log("❌ Gerenciador de Relatórios não encontrado")
                return False

            self.log("📋 Gerenciador de Relatórios localizado")
            relatorio_window.set_focus()
            if not self.smart_sleep(0.5):
                return False

            # Navegar até a pasta Diversos: pressionar 'd' 6 vezes, depois ENTER
            self.log("📁 Navegando para pasta Diversos (d x6 + Enter)")
            for _ in range(6):
                if self.should_stop():
                    return False
                send_keys('d')
                time.sleep(0.2)
            send_keys('{ENTER}')
            if not self.smart_sleep(0.5):
                return False

            # Pressionar 'c' 9 vezes para selecionar a guia Taxa Motorista
            self.log("🎯 Selecionando guia Taxa Motorista (c x10)")
            for _ in range(10):
                if self.should_stop():
                    return False
                send_keys('c')
                time.sleep(0.2)

            if not self.smart_sleep(0.5):
                return False

            # Preencher campos: TAB (pula 1º) → TAB + competência → TAB + rubrica
            send_keys('{TAB}')
            time.sleep(0.2)

            self.log(f"📝 Preenchendo competência: {competencia}")
            send_keys('{TAB}')
            time.sleep(0.2)
            send_keys(competencia, with_spaces=True)
            if not self.smart_sleep(0.3):
                return False

            self.log(f"📝 Preenchendo rubrica: {rubrica}")
            send_keys('{TAB}')
            time.sleep(0.2)
            send_keys(rubrica, with_spaces=True)
            if not self.smart_sleep(0.3):
                return False

            if self.should_stop():
                return False
            self.check_pause()

            # Executar relatório
            self.log("⚡ Executando relatório")
            try:
                button_executar = relatorio_window.child_window(auto_id="1007", class_name="Button")
                button_executar.click_input()
                if not self.smart_sleep(4):
                    return False
            except Exception as e:
                self.log(f"⚠️ Botão executar não encontrado, tentando F5: {e}")
                send_keys('{F5}')
                if not self.smart_sleep(4):
                    return False

            # Aguardar janela de salvamento com reenvio periódico de Ctrl+D
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

            return self.salvar_pdf(caminho)

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
