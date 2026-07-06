"""
DomBot - Taxa Motorista
=======================
Automação RPA para emissão de relatórios de Taxa Motorista no sistema Domínio Folha.

As empresas são cadastradas e gerenciadas pela própria interface (aba "🏢 Empresas")
e ficam salvas automaticamente em 'empresas_motorista.json' ao lado deste script.
É possível importar uma vez a partir do Excel antigo, cujas colunas são:
  Coluna A - Codigo da empresa
  Coluna B - Nome da empresa
  Coluna C - Rubrica

Autor: Hugo L. Almeida
Versão: 2.0
"""

import customtkinter as ctk
import pandas as pd
import time
import logging
import os
import json
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

        self.pasta_saida = ctk.StringVar()
        self.competencia_var = ctk.StringVar(value=datetime.now().strftime("%m/%Y"))
        self.data_vencimento_var = ctk.StringVar()
        self.status_var = ctk.StringVar(value="Aguardando início...")

        # Lista de empresas cadastradas (gerenciada pela interface e salva em JSON).
        # Cada item: {"codigo": str, "nome": str, "rubrica": str, "ativo": bool}
        self.empresas = []
        self.empresas_json = os.path.join(
            os.path.dirname(os.path.abspath(__file__)), "empresas_motorista.json"
        )
        self.empresa_widgets = []  # linhas (frame + checkbox var) renderizadas na aba

        self.linhas_processadas = 0
        self.linhas_com_erro = 0

        self.logger = logging.getLogger('AutomacaoTaxaMotorista')
        self.logger.setLevel(logging.INFO)
        self.logger.handlers = []

        self.gui_handler = GUILogHandler(self)
        self.gui_handler.setFormatter(logging.Formatter('%(message)s'))
        self.logger.addHandler(self.gui_handler)

        self.criar_interface()
        self.carregar_empresas_json()
        self.render_lista_empresas()

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

        # --- Linha 1: informação das empresas + botões de controle ---
        row1 = ctk.CTkFrame(config_frame, fg_color="transparent")
        row1.grid(row=0, column=0, sticky="ew", padx=10, pady=(8, 4))
        row1.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(row1, text="🏢", font=ctk.CTkFont(size=14)).grid(row=0, column=0, padx=(0, 5))

        self.lbl_resumo_empresas = ctk.CTkLabel(
            row1, text="Nenhuma empresa cadastrada",
            font=ctk.CTkFont(size=11), text_color="#BDC3C7", anchor="w",
        )
        self.lbl_resumo_empresas.grid(row=0, column=1, sticky="ew", padx=(0, 15))

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

        # --- Linha 2: Competência + Vencimento + Pasta de saída ---
        row2 = ctk.CTkFrame(config_frame, fg_color="transparent")
        row2.grid(row=1, column=0, sticky="ew", padx=10, pady=(0, 8))
        row2.grid_columnconfigure(7, weight=1)

        ctk.CTkLabel(row2, text="📅", font=ctk.CTkFont(size=14)).grid(row=0, column=0, padx=(0, 5))
        ctk.CTkLabel(row2, text="Competência:", font=ctk.CTkFont(size=11), text_color="#BDC3C7").grid(
            row=0, column=1, padx=(0, 5)
        )
        ctk.CTkEntry(
            row2, textvariable=self.competencia_var,
            width=100, height=32, font=ctk.CTkFont(size=11), justify="center",
            placeholder_text="MM/AAAA",
        ).grid(row=0, column=2, padx=(0, 15))

        ctk.CTkLabel(row2, text="⏳", font=ctk.CTkFont(size=14)).grid(row=0, column=3, padx=(0, 5))
        ctk.CTkLabel(row2, text="Vencimento:", font=ctk.CTkFont(size=11), text_color="#BDC3C7").grid(
            row=0, column=4, padx=(0, 5)
        )
        ctk.CTkEntry(
            row2, textvariable=self.data_vencimento_var,
            width=100, height=32, font=ctk.CTkFont(size=11), justify="center",
            placeholder_text="DD/MM/AAAA",
        ).grid(row=0, column=5, padx=(0, 15))

        ctk.CTkLabel(row2, text="📂", font=ctk.CTkFont(size=14)).grid(row=0, column=6, padx=(0, 5))
        ctk.CTkEntry(
            row2, textvariable=self.pasta_saida,
            height=32, font=ctk.CTkFont(size=11),
            placeholder_text="Pasta de saída dos PDFs...",
        ).grid(row=0, column=7, sticky="ew", padx=(0, 8))
        ctk.CTkButton(
            row2, text="Pasta", command=self.selecionar_pasta_saida,
            width=70, height=32, font=ctk.CTkFont(size=11),
            fg_color=self.CORES['info'], hover_color="#2980B9",
        ).grid(row=0, column=8)

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

        tab_empresas = self.tabview.add("🏢 Empresas")
        tab_logs = self.tabview.add("📋 Logs")

        self._criar_aba_empresas(tab_empresas)
        self._criar_aba_logs(tab_logs)

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

    def _criar_aba_empresas(self, parent):
        parent.grid_columnconfigure(0, weight=1)
        parent.grid_rowconfigure(2, weight=1)

        # --- Barra de ações ---
        acoes = ctk.CTkFrame(parent, fg_color="transparent")
        acoes.grid(row=0, column=0, sticky="ew", padx=3, pady=(3, 2))

        ctk.CTkButton(
            acoes, text="➕ Adicionar", command=self.adicionar_empresa_dialog,
            width=100, height=28, font=ctk.CTkFont(size=11, weight="bold"),
            fg_color=self.CORES['sucesso'], hover_color="#27AE60",
        ).pack(side="left", padx=(0, 6))

        ctk.CTkButton(
            acoes, text="📥 Importar do Excel", command=self.importar_empresas_excel,
            width=150, height=28, font=ctk.CTkFont(size=11),
            fg_color=self.CORES['info'], hover_color="#2980B9",
        ).pack(side="left", padx=(0, 6))

        ctk.CTkButton(
            acoes, text="☑ Todas", command=lambda: self.marcar_todas(True),
            width=70, height=28, font=ctk.CTkFont(size=11),
            fg_color="#34495E", hover_color="#2C3E50",
        ).pack(side="left", padx=(0, 4))

        ctk.CTkButton(
            acoes, text="☐ Nenhuma", command=lambda: self.marcar_todas(False),
            width=90, height=28, font=ctk.CTkFont(size=11),
            fg_color="#34495E", hover_color="#2C3E50",
        ).pack(side="left")

        # --- Cabeçalho da lista ---
        header = ctk.CTkFrame(parent, fg_color=self.CORES['fundo_escuro'], corner_radius=6)
        header.grid(row=1, column=0, sticky="ew", padx=3, pady=(2, 0))
        for col, (txt, w) in enumerate([("", 30), ("Código", 90), ("Nome", 260),
                                        ("Rubrica", 90), ("Ações", 120)]):
            ctk.CTkLabel(header, text=txt, width=w, font=ctk.CTkFont(size=10, weight="bold"),
                         text_color="#7F8C8D", anchor="w").grid(row=0, column=col, padx=6, pady=4, sticky="w")

        # --- Lista rolável de empresas ---
        self.lista_empresas_frame = ctk.CTkScrollableFrame(
            parent, fg_color=self.CORES['fundo_escuro'], corner_radius=6,
        )
        self.lista_empresas_frame.grid(row=2, column=0, sticky="nsew", padx=3, pady=(0, 3))
        self.lista_empresas_frame.grid_columnconfigure(0, weight=1)

    def selecionar_pasta_saida(self):
        import subprocess
        try:
            script = (
                "Add-Type -AssemblyName System.Windows.Forms; "
                "$f = New-Object System.Windows.Forms.FolderBrowserDialog; "
                "$f.Description = 'Selecione a pasta de saida dos PDFs'; "
                "$f.ShowNewFolderButton = $true; "
                "if ($f.ShowDialog() -eq 'OK') { $f.SelectedPath } else { '' }"
            )
            result = subprocess.run(
                ["powershell", "-Command", script],
                capture_output=True, text=True, timeout=120
            )
            pasta = result.stdout.strip()
            if pasta:
                self.pasta_saida.set(pasta)
                self.adicionar_log(f"Pasta de saída: {pasta}", logging.INFO, "info")
        except Exception as e:
            self.adicionar_log(f"Erro ao selecionar pasta: {e}", logging.ERROR, "erro")

    # ── Gerenciamento de empresas (JSON + interface) ──────────────────────────

    def carregar_empresas_json(self):
        """Carrega a lista de empresas do arquivo JSON, se existir."""
        try:
            if os.path.exists(self.empresas_json):
                with open(self.empresas_json, 'r', encoding='utf-8') as f:
                    dados = json.load(f)
                self.empresas = []
                for item in dados:
                    self.empresas.append({
                        'codigo': str(item.get('codigo', '')).strip(),
                        'nome': str(item.get('nome', '')).strip(),
                        'rubrica': str(item.get('rubrica', '')).strip(),
                        'ativo': bool(item.get('ativo', True)),
                    })
                self.adicionar_log(f"{len(self.empresas)} empresa(s) carregada(s) do cadastro.",
                                   logging.INFO, "sucesso")
        except Exception as e:
            self.adicionar_log(f"Erro ao carregar cadastro de empresas: {e}", logging.ERROR, "erro")

    def salvar_empresas_json(self):
        """Grava a lista atual de empresas no arquivo JSON."""
        try:
            with open(self.empresas_json, 'w', encoding='utf-8') as f:
                json.dump(self.empresas, f, ensure_ascii=False, indent=2)
        except Exception as e:
            self.adicionar_log(f"Erro ao salvar cadastro de empresas: {e}", logging.ERROR, "erro")

    def render_lista_empresas(self):
        """Redesenha a lista rolável de empresas com base em self.empresas."""
        for w in self.lista_empresas_frame.winfo_children():
            w.destroy()
        self.empresa_widgets = []

        if not self.empresas:
            ctk.CTkLabel(
                self.lista_empresas_frame,
                text="Nenhuma empresa cadastrada.\nUse ➕ Adicionar ou 📥 Importar do Excel.",
                font=ctk.CTkFont(size=12), text_color="#7F8C8D", justify="center",
            ).grid(row=0, column=0, pady=30)
            self._atualizar_resumo_empresas()
            return

        for idx, emp in enumerate(self.empresas):
            linha = ctk.CTkFrame(self.lista_empresas_frame, fg_color=self.CORES['fundo_card'],
                                 corner_radius=6)
            linha.grid(row=idx, column=0, sticky="ew", padx=2, pady=2)
            linha.grid_columnconfigure(2, weight=1)

            var = ctk.BooleanVar(value=emp.get('ativo', True))
            chk = ctk.CTkCheckBox(
                linha, text="", variable=var, width=24,
                command=lambda i=idx, v=var: self._toggle_ativo(i, v),
                checkbox_width=20, checkbox_height=20,
            )
            chk.grid(row=0, column=0, padx=(8, 2), pady=6)

            ctk.CTkLabel(linha, text=emp['codigo'], width=90, anchor="w",
                         font=ctk.CTkFont(size=11)).grid(row=0, column=1, padx=6, sticky="w")
            ctk.CTkLabel(linha, text=emp['nome'], anchor="w",
                         font=ctk.CTkFont(size=11)).grid(row=0, column=2, padx=6, sticky="w")
            ctk.CTkLabel(linha, text=emp['rubrica'], width=90, anchor="w",
                         font=ctk.CTkFont(size=11)).grid(row=0, column=3, padx=6, sticky="w")

            btns = ctk.CTkFrame(linha, fg_color="transparent")
            btns.grid(row=0, column=4, padx=6)
            ctk.CTkButton(btns, text="✏", width=32, height=26, font=ctk.CTkFont(size=12),
                          fg_color=self.CORES['info'], hover_color="#2980B9",
                          command=lambda i=idx: self.editar_empresa_dialog(i)).pack(side="left", padx=2)
            ctk.CTkButton(btns, text="🗑", width=32, height=26, font=ctk.CTkFont(size=12),
                          fg_color=self.CORES['erro'], hover_color="#C0392B",
                          command=lambda i=idx: self.remover_empresa(i)).pack(side="left", padx=2)

            self.empresa_widgets.append(var)

        self._atualizar_resumo_empresas()

    def _atualizar_resumo_empresas(self):
        total = len(self.empresas)
        ativas = sum(1 for e in self.empresas if e.get('ativo', True))
        if total == 0:
            texto = "Nenhuma empresa cadastrada"
        else:
            texto = f"{ativas} de {total} empresa(s) selecionada(s) para processar"
        try:
            self.lbl_resumo_empresas.configure(text=texto)
        except Exception:
            pass

    def _toggle_ativo(self, idx, var):
        if 0 <= idx < len(self.empresas):
            self.empresas[idx]['ativo'] = bool(var.get())
            self.salvar_empresas_json()
            self._atualizar_resumo_empresas()

    def marcar_todas(self, valor: bool):
        for emp in self.empresas:
            emp['ativo'] = valor
        self.salvar_empresas_json()
        self.render_lista_empresas()

    def _dialog_empresa(self, titulo, codigo="", nome="", rubrica=""):
        """Abre um diálogo modal para cadastrar/editar. Retorna dict ou None se cancelado."""
        dlg = ctk.CTkToplevel(self.window)
        dlg.title(titulo)
        dlg.geometry("420x280")
        dlg.transient(self.window)
        dlg.grab_set()
        dlg.resizable(False, False)

        resultado = {'valor': None}
        var_cod = ctk.StringVar(value=codigo)
        var_nome = ctk.StringVar(value=nome)
        var_rub = ctk.StringVar(value=rubrica)

        frame = ctk.CTkFrame(dlg, fg_color="transparent")
        frame.pack(fill="both", expand=True, padx=20, pady=20)
        frame.grid_columnconfigure(1, weight=1)

        campos = [("Código:", var_cod), ("Nome:", var_nome), ("Rubrica:", var_rub)]
        entradas = []
        for i, (rot, v) in enumerate(campos):
            ctk.CTkLabel(frame, text=rot, font=ctk.CTkFont(size=12), width=70,
                         anchor="w").grid(row=i, column=0, padx=(0, 8), pady=8, sticky="w")
            e = ctk.CTkEntry(frame, textvariable=v, height=34, font=ctk.CTkFont(size=12))
            e.grid(row=i, column=1, sticky="ew", pady=8)
            entradas.append(e)
        entradas[0].focus()

        erro_lbl = ctk.CTkLabel(frame, text="", font=ctk.CTkFont(size=11),
                                text_color=self.CORES['erro'])
        erro_lbl.grid(row=3, column=0, columnspan=2, sticky="w")

        def confirmar():
            c, n, r = var_cod.get().strip(), var_nome.get().strip(), var_rub.get().strip()
            if not c or not n or not r:
                erro_lbl.configure(text="Preencha código, nome e rubrica.")
                return
            resultado['valor'] = {'codigo': c, 'nome': n, 'rubrica': r, 'ativo': True}
            dlg.destroy()

        botoes = ctk.CTkFrame(frame, fg_color="transparent")
        botoes.grid(row=4, column=0, columnspan=2, pady=(12, 0), sticky="e")
        ctk.CTkButton(botoes, text="Cancelar", command=dlg.destroy, width=100, height=32,
                      fg_color="#34495E", hover_color="#2C3E50").pack(side="left", padx=6)
        ctk.CTkButton(botoes, text="Salvar", command=confirmar, width=100, height=32,
                      fg_color=self.CORES['sucesso'], hover_color="#27AE60").pack(side="left")

        dlg.bind('<Return>', lambda _e: confirmar())
        dlg.wait_window()
        return resultado['valor']

    def adicionar_empresa_dialog(self):
        nova = self._dialog_empresa("Adicionar Empresa")
        if nova:
            self.empresas.append(nova)
            self.salvar_empresas_json()
            self.render_lista_empresas()
            self.adicionar_log(f"Empresa adicionada: {nova['codigo']} - {nova['nome']}",
                               logging.INFO, "sucesso")

    def editar_empresa_dialog(self, idx):
        if not (0 <= idx < len(self.empresas)):
            return
        emp = self.empresas[idx]
        editada = self._dialog_empresa("Editar Empresa", emp['codigo'], emp['nome'], emp['rubrica'])
        if editada:
            editada['ativo'] = emp.get('ativo', True)  # preserva a seleção
            self.empresas[idx] = editada
            self.salvar_empresas_json()
            self.render_lista_empresas()
            self.adicionar_log(f"Empresa atualizada: {editada['codigo']} - {editada['nome']}",
                               logging.INFO, "info")

    def remover_empresa(self, idx):
        if not (0 <= idx < len(self.empresas)):
            return
        emp = self.empresas[idx]
        if messagebox.askyesno("Confirmar remoção",
                               f"Remover a empresa:\n\n{emp['codigo']} - {emp['nome']}?"):
            self.empresas.pop(idx)
            self.salvar_empresas_json()
            self.render_lista_empresas()
            self.adicionar_log(f"Empresa removida: {emp['codigo']} - {emp['nome']}",
                               logging.INFO, "aviso")

    def importar_empresas_excel(self):
        """Importa empresas do Excel antigo (colunas: Codigo, Nome, Rubrica)."""
        filename = ctk.filedialog.askopenfilename(
            filetypes=[("Excel files", "*.xlsx *.xls")],
            title="Selecione o Excel para importar as empresas",
        )
        if not filename:
            return
        try:
            df = pd.read_excel(filename, header=None, names=['Codigo', 'Nome', 'Rubrica'])
            # Pular linha de cabeçalho se existir
            if str(df.iloc[0]['Codigo']).lower() in ('codigo', 'código', 'code'):
                df = df.iloc[1:].reset_index(drop=True)

            existentes = {e['codigo'] for e in self.empresas}
            importadas = 0
            duplicadas = 0
            for _, row in df.iterrows():
                codigo = str(row['Codigo']).strip()
                if not codigo or codigo.lower() == 'nan':
                    continue
                if codigo in existentes:
                    duplicadas += 1
                    continue
                self.empresas.append({
                    'codigo': codigo,
                    'nome': str(row.get('Nome', '')).strip(),
                    'rubrica': str(row['Rubrica']).strip(),
                    'ativo': True,
                })
                existentes.add(codigo)
                importadas += 1

            self.salvar_empresas_json()
            self.render_lista_empresas()
            msg = f"Importação concluída: {importadas} adicionada(s)"
            if duplicadas:
                msg += f", {duplicadas} já existente(s) ignorada(s)"
            self.adicionar_log(msg, logging.INFO, "sucesso")
            messagebox.showinfo("Importação", msg)
        except Exception as e:
            self.adicionar_log(f"Erro ao importar do Excel: {e}", logging.ERROR, "erro")
            messagebox.showerror("Erro na importação", str(e))

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
        if not self.empresas:
            return False, "Nenhuma empresa cadastrada. Adicione ou importe do Excel na aba 🏢 Empresas."
        if not any(e.get('ativo', True) for e in self.empresas):
            return False, "Nenhuma empresa selecionada. Marque ao menos uma na aba 🏢 Empresas."
        if not self.competencia_var.get().strip():
            return False, "Informe a competência antes de continuar."
        venc = self.data_vencimento_var.get().strip()
        if not venc:
            return False, "Informe a data de vencimento (DD/MM/AAAA) para a publicação."
        try:
            datetime.strptime(venc, "%d/%m/%Y")
        except ValueError:
            return False, "Data de vencimento inválida. Use o formato DD/MM/AAAA."
        if not self.pasta_saida.get().strip():
            return False, "Selecione a pasta de saída dos PDFs."
        if not os.path.isdir(self.pasta_saida.get().strip()):
            return False, "Pasta de saída não encontrada. Verifique o caminho."
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
        competencia = self.competencia_var.get().strip()

        try:
            self.adicionar_log("Iniciando automação...", logging.INFO, "processando")
            self.status_var.set("Em execução...")
            self.executando = True

            # Snapshot das empresas selecionadas (ativas) no momento do início
            empresas_processar = [e for e in self.empresas if e.get('ativo', True)]
            total = len(empresas_processar)
            pasta_saida = self.pasta_saida.get().strip()

            self.adicionar_log(f"{total} empresa(s) selecionada(s) para processar", logging.INFO, "info")

            automacao = DominioAutomation(self.logger, self)

            if not automacao.connect_to_dominio():
                self.adicionar_log("Não foi possível conectar ao Domínio", logging.ERROR, "erro")
                return

            for idx, emp in enumerate(empresas_processar):
                if not self.executando:
                    self.adicionar_log("Automação interrompida pelo usuário", logging.INFO, "aviso")
                    break

                while self.pausa_solicitada and self.executando:
                    time.sleep(0.5)

                if not self.executando:
                    break

                num = idx + 1
                codigo = str(emp['codigo']).strip()
                nome = str(emp.get('nome', 'N/A')).strip()
                rubrica = str(emp['rubrica']).strip()
                # Monta o nome do arquivo e sanitiza caracteres inválidos
                comp_fmt = competencia.replace('/', '-')
                nome_arq = f"{codigo} - {nome} - {comp_fmt}.pdf"
                for ch in r'*?"<>|':
                    nome_arq = nome_arq.replace(ch, '-')
                caminho = os.path.join(pasta_saida, nome_arq)

                self.empresa_label.configure(text=codigo[:20])
                self.status_var.set(f"Processando: {num}/{total}")
                self.adicionar_log(
                    f"[{num}/{total}] Codigo {codigo} - {nome} - Rubrica {rubrica}",
                    logging.INFO, "processando"
                )

                try:
                    success = automacao.processar_taxa_motorista(
                        rubrica, codigo, nome, caminho, competencia, num
                    )

                    if success:
                        self.linhas_processadas += 1
                        self.sucesso_label.configure(text=str(self.linhas_processadas))
                        self.success_logger.info(f"Codigo {codigo} - {nome} - OK")
                        self.adicionar_log(f"{codigo} - {nome} processada com sucesso", logging.INFO, "sucesso")
                    else:
                        self.linhas_com_erro += 1
                        self.erros_label.configure(text=str(self.linhas_com_erro))
                        self.error_logger.error(f"Codigo {codigo} - {nome} - ERRO")
                        self.adicionar_log(f"Erro ao processar {codigo} - {nome}", logging.ERROR, "erro")

                except Exception as e:
                    self.linhas_com_erro += 1
                    self.erros_label.configure(text=str(self.linhas_com_erro))
                    erro_msg = f"{codigo} - {nome} - Erro: {e}"
                    self.error_logger.error(erro_msg)
                    self.adicionar_log(erro_msg, logging.ERROR, "erro")

            # Publicação em lote: só após salvar todos os PDFs e fechar as janelas
            publicados, falhas_publicacao = 0, 0
            if self.executando:
                self.adicionar_log(
                    f"Salvamento concluído! {self.linhas_processadas} salvos, {self.linhas_com_erro} com erro.",
                    logging.INFO, "sucesso"
                )
                automacao.cleanup_windows()
                self.status_var.set("Publicando no portal...")
                self.atualizar_status_indicator('executando')
                data_venc = self.data_vencimento_var.get().strip()
                publicados, falhas_publicacao = automacao.publicar_lote(pasta_saida, data_venc)
                self.adicionar_log(
                    f"Publicação concluída: {publicados} publicado(s), {falhas_publicacao} falha(s).",
                    logging.INFO, "sucesso" if falhas_publicacao == 0 else "aviso"
                )

            if self.executando:
                self.status_var.set("Processamento concluído")
                self.atualizar_status_indicator('concluido')
                self.adicionar_log(
                    f"Automação concluída! {self.linhas_processadas} salvos, {self.linhas_com_erro} com erro no salvamento.",
                    logging.INFO, "sucesso"
                )

                if _requests_available:
                    try:
                        webhook_url = os.getenv("DISCORD_WEBHOOK_URL")
                        if webhook_url:
                            mensagem = (
                                f"📋 **Taxa Motorista Emitida**\n\n"
                                f"📊 **Salvos:** {self.linhas_processadas}\n"
                                f"❌ **Erros no salvamento:** {self.linhas_com_erro}\n"
                                f"🌐 **Publicados no Onvio:** {publicados}\n"
                                f"⚠️ **Falhas na publicação:** {falhas_publicacao}\n"
                                f"📂 **Destino:** `{pasta_saida}`\n\n"
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

    def executar(self):
        self.window.mainloop()


# ── Classe de automação Domínio ───────────────────────────────────────────────

class DominioAutomation:
    # Pasta padrão de publicação no portal Onvio (janela "Publicação de Documentos")
    PASTA_PUBLICACAO = "Pessoal/Taxa Assistencial"

    # Janela "Publicação de Documentos Externos" (publicação em lote)
    PUB_LOTE_TITULO = "Publicação de Documentos Externos"
    PUB_LOTE_CLASSE = "FNWND3190"

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

    # ── Publicação em lote (Publicação de Documentos Externos) ────────────────

    def _get_pub_lote_window(self):
        """
        Retorna o objeto pywinauto da janela 'Publicação de Documentos Externos'.
        A janela é FILHA (MDI child) da janela do Domínio — por isso é localizada via
        child_window da main_window, e NÃO por EnumWindows (que só vê janelas top-level).
        """
        return self.main_window.child_window(
            title=self.PUB_LOTE_TITULO, class_name=self.PUB_LOTE_CLASSE
        )

    def _pub_lote_window_ok(self) -> bool:
        try:
            pub = self._get_pub_lote_window()
            return pub.exists() and pub.is_visible()
        except Exception:
            return False

    def _abrir_janela_pub_lote(self) -> bool:
        """
        Garante a janela de publicação em lote aberta. Se ela já estiver aberta
        (permanece entre publicações), retorna direto. Caso contrário, clica no
        botão-nuvem para abri-la.
        """
        # 1) Já está aberta?
        if self._pub_lote_window_ok():
            self.log("📋 Janela de Publicação em Lote já está aberta")
            return True

        # 2) Tenta abrir pelo botão-nuvem
        for tentativa in range(3):
            if self.should_stop():
                return False
            try:
                self.main_window.set_focus()
                btn_nuvem = self.main_window.child_window(
                    auto_id="picturePublicacaoDocumentosExternos"
                )
                if btn_nuvem.exists(timeout=2):
                    btn_nuvem.click_input()
                    self.log("☁️ Botão de publicação em lote clicado")
            except Exception as e:
                self.log(f"⚠️ Tentativa {tentativa + 1} de clicar no botão-nuvem falhou: {e}")

            if self.wait_for_condition(
                self._pub_lote_window_ok,
                timeout=10, poll_interval=0.3,
                description="Aguardando janela de Publicação em Lote",
            ):
                return True
        return False

    def _garantir_checkbox(self, pub, auto_id: str, nome: str) -> bool:
        """Garante que um checkbox esteja marcado (ToggleState On), com re-verificação."""
        try:
            chk = pub.child_window(auto_id=auto_id, class_name="Button")
            if not chk.exists(timeout=2):
                self.log(f"⚠️ Checkbox '{nome}' não encontrado")
                return False
            for _ in range(2):
                try:
                    estado = chk.get_toggle_state()  # 1=On, 0=Off
                except Exception:
                    estado = None
                if estado == 1:
                    return True
                # Off ou desconhecido → clicar para marcar
                self.log(f"☑ Marcando '{nome}'")
                try:
                    chk.click_input()
                except Exception:
                    chk.click()
                self.smart_sleep(0.3)
            # Confere uma última vez
            try:
                return chk.get_toggle_state() == 1
            except Exception:
                return True
        except Exception as e:
            self.log(f"⚠️ Não foi possível marcar '{nome}': {e}")
            return False

    def _preencher_data_mascarada(self, edit_data, digitos: str, data_fmt: str) -> bool:
        """
        Preenche um campo de data com máscara (00/00/0000). Garante o cursor no início
        (HOME + setas à esquerda) antes de digitar só os dígitos DDMMAAAA. Confere a
        leitura de volta e faz uma 2ª tentativa alternativa se sair errado.
        """
        def _ler():
            try:
                v = (edit_data.get_value() or "").strip()
            except Exception:
                try:
                    v = (edit_data.window_text() or "").strip()
                except Exception:
                    v = ""
            return v

        def _bate(valor):
            # Considera correto se os dígitos lidos == dígitos esperados
            return ''.join(c for c in valor if c.isdigit()) == digitos

        # Tentativa 1: focar, ir ao início, limpar e digitar dígitos
        try:
            edit_data.set_focus()
            time.sleep(0.1)
            # posiciona no início da máscara
            edit_data.type_keys("{HOME}{LEFT 12}", set_foreground=False)
            time.sleep(0.1)
            edit_data.type_keys(digitos, set_foreground=False)
            time.sleep(0.2)
            if _bate(_ler()):
                return True
        except Exception as e:
            self.log(f"⚠️ Data (tentativa 1) falhou: {e}")

        # Tentativa 2: limpar tudo e usar set_text com a data formatada
        try:
            edit_data.set_focus()
            edit_data.type_keys("{HOME}{LEFT 12}{DELETE 12}", set_foreground=False)
            time.sleep(0.1)
            try:
                edit_data.set_text(data_fmt)
            except Exception:
                edit_data.type_keys(digitos, set_foreground=False)
            time.sleep(0.2)
            valor = _ler()
            if _bate(valor):
                return True
            self.log(f"⚠️ Data lida do campo: '{valor}' (esperado {data_fmt})")
        except Exception as e:
            self.log(f"⚠️ Data (tentativa 2) falhou: {e}")

        return False

    def _configurar_envio_lote(self, pub, data_vencimento: str) -> bool:
        """
        Configuração feita UMA vez (a janela permanece aberta entre publicações):
        pasta de publicação, data de vencimento (checkbox + data) e 'Concluir atividade'.
        """
        try:
            # Pasta de publicação (ComboBox, auto_id 1001) — lista fixa: selecionar o item
            self.log(f"📁 Configurando pasta: {self.PASTA_PUBLICACAO}")
            try:
                combo = pub.child_window(auto_id="1001", class_name="ComboBox")
                combo.set_focus()
                selecionado = False
                # 1ª tentativa: selecionar pelo texto exato do item na lista suspensa
                try:
                    combo.select(self.PASTA_PUBLICACAO)
                    selecionado = True
                except Exception:
                    # 2ª tentativa: casar de forma tolerante com os itens disponíveis
                    try:
                        itens = combo.item_texts()
                        alvo = self.PASTA_PUBLICACAO.strip().lower()
                        for i, txt in enumerate(itens):
                            t = (txt or "").strip().lower()
                            if t == alvo or alvo in t:
                                combo.select(i)
                                selecionado = True
                                self.log(f"📁 Pasta selecionada da lista: '{txt}'")
                                break
                        if not selecionado:
                            self.log(f"⚠️ '{self.PASTA_PUBLICACAO}' não está na lista. Itens: {itens}")
                    except Exception as e2:
                        self.log(f"⚠️ Não foi possível ler os itens da lista de pastas: {e2}")
                if not selecionado:
                    # 3ª tentativa (fallback): digitar o texto no campo editável
                    try:
                        combo.set_edit_text(self.PASTA_PUBLICACAO)
                    except Exception:
                        pass
            except Exception as e:
                self.log(f"⚠️ Não foi possível definir a pasta: {e}")
            self.smart_sleep(0.4)

            # Checkbox 'Data de vencimento' (auto_id 1006) — garantir marcado ANTES da data
            # (o campo da data costuma ficar desabilitado enquanto o checkbox está off)
            self._garantir_checkbox(pub, "1006", "Data de vencimento")
            self.smart_sleep(0.3)

            # Campo da data (PBEDIT190, auto_id 1005) — campo com máscara 00/00/0000.
            # É preciso posicionar o cursor no INÍCIO antes de digitar, senão os dígitos
            # entram no meio da máscara e a data sai embaralhada (ex.: 01/00/7202).
            self.log(f"📅 Definindo data de vencimento: {data_vencimento}")
            digitos = ''.join(ch for ch in data_vencimento if ch.isdigit())  # DDMMAAAA
            try:
                edit_data = pub.child_window(auto_id="1005", class_name="PBEDIT190")
                if not self._preencher_data_mascarada(edit_data, digitos, data_vencimento):
                    self.log("⚠️ Data pode ter ficado incorreta no campo")
                self.smart_sleep(0.3)
            except Exception as e:
                self.log(f"⚠️ Não foi possível definir a data de vencimento: {e}")

            # Checkbox 'Concluir atividade' (auto_id 1004) — garantir marcado
            self._garantir_checkbox(pub, "1004", "Concluir atividade")

            return True
        except Exception as e:
            self.log(f"❌ Erro ao configurar envio em lote: {e}")
            return False

    def _publicar_um_documento(self, pub, caminho_pdf: str, codigo: str) -> bool:
        """
        Publica um único documento na janela já aberta e confirma o OK.
        Usa o idioma comprovado do publicador GMS: set_text + type_keys('^a{DELETE}').
        """
        try:
            # Caminho do documento (Edit, auto_id 1013)
            self.log(f"📄 Caminho: {os.path.basename(caminho_pdf)}")
            try:
                campo_caminho = pub.child_window(auto_id="1013", class_name="Edit")
                if not campo_caminho.exists(timeout=3):
                    self.log("❌ Campo 'Caminho' não encontrado")
                    return False
                campo_caminho.set_focus()
                campo_caminho.type_keys("^a{DELETE}", set_foreground=False)
                time.sleep(0.3)
                campo_caminho.set_text(caminho_pdf)
            except Exception as e:
                self.log(f"❌ Não foi possível preencher o caminho: {e}")
                return False
            self.smart_sleep(0.5)

            # Código da empresa (PBEDIT190, auto_id 1001)
            self.log(f"🏢 Código da empresa: {codigo}")
            try:
                campo_codigo = pub.child_window(auto_id="1001", class_name="PBEDIT190")
                if not campo_codigo.exists(timeout=3):
                    self.log("❌ Campo 'Código' não encontrado")
                    return False
                campo_codigo.set_focus()
                campo_codigo.type_keys("^a{DELETE}", set_foreground=False)
                time.sleep(0.3)
                campo_codigo.set_text(codigo)
            except Exception as e:
                self.log(f"❌ Não foi possível preencher o código da empresa: {e}")
                return False
            self.smart_sleep(0.5)

            # Botão Publicar (auto_id 1003)
            self.log("⚡ Publicando documento")
            try:
                botao_publicar = pub.child_window(auto_id="1003", class_name="Button")
                if not botao_publicar.exists(timeout=3):
                    self.log("❌ Botão 'Publicar' não encontrado")
                    return False
                botao_publicar.click()
                time.sleep(2)
            except Exception as e:
                self.log(f"❌ Erro ao clicar em 'Publicar': {e}")
                return False

            # Aguardar e confirmar o diálogo de confirmação
            dialog = self._aguardar_confirmacao(timeout=15)
            if dialog is False:
                return False  # interrompido
            if dialog:
                if self._clicar_botao_ok(dialog):
                    time.sleep(1)
                    return True
                self.log("⚠️ Falha ao clicar no OK de confirmação")
                return False
            else:
                self.log("⚠️ Janela de confirmação não encontrada")
                return False

        except Exception as e:
            self.log(f"❌ Erro ao publicar documento {os.path.basename(caminho_pdf)}: {e}")
            return False

    def _aguardar_confirmacao(self, timeout=15):
        """
        Aguarda o diálogo de confirmação após 'Publicar'. Retorna o objeto da janela,
        None (timeout) ou False (interrompido). Adaptado do publicador GMS.
        """
        self.log("🔍 Procurando janela de confirmação...")
        inicio = time.time()
        while (time.time() - inicio) < timeout:
            if self.should_stop():
                self.log("⏹️ Busca por confirmação interrompida")
                return False
            self.check_pause()
            try:
                all_windows = findwindows.find_windows()
                for hwnd in all_windows:
                    try:
                        window = self.app.window(handle=hwnd)
                        if window.is_dialog() and window.is_visible():
                            titulo = window.window_text()
                            if titulo and any(p in titulo.lower() for p in
                                              ['atenção', 'confirmação', 'aviso', 'informação', 'sucesso']):
                                self.log(f"✅ Confirmação encontrada: '{titulo}'")
                                return window
                    except Exception:
                        continue
            except Exception:
                pass
            time.sleep(0.5)
        self.log("⚠️ Timeout: nenhuma janela de confirmação encontrada")
        return None

    def _clicar_botao_ok(self, dialog) -> bool:
        """Clica no OK/Confirmar/Sim do diálogo. Adaptado do publicador GMS."""
        for texto in ["OK", "Ok", "Confirmar", "Sim", "Yes"]:
            try:
                botao = dialog.child_window(title=texto, control_type="Button")
                if botao.exists(timeout=2):
                    botao.click()
                    self.log(f"✅ Botão '{texto}' clicado")
                    return True
            except Exception:
                continue
        for auto_id in ["1", "2", "6", "1001", "2001"]:
            try:
                botao = dialog.child_window(auto_id=auto_id, control_type="Button")
                if botao.exists(timeout=2):
                    botao.click()
                    self.log(f"✅ Botão auto_id '{auto_id}' clicado")
                    return True
            except Exception:
                continue
        try:
            botoes = dialog.children(control_type="Button")
            if botoes:
                botoes[0].click()
                self.log("✅ Primeiro botão do diálogo clicado")
                return True
        except Exception:
            pass
        return False

    def publicar_lote(self, pasta_saida: str, data_vencimento: str) -> tuple:
        """
        Publica em lote todos os PDFs da pasta de saída na janela
        'Publicação de Documentos Externos'. A configuração de envio (pasta/data/
        concluir) é feita uma única vez; depois, cada documento é publicado em sequência.
        Retorna (publicados, falhas).
        """
        publicados = 0
        falhas = 0
        try:
            pdfs = sorted(
                os.path.join(pasta_saida, f)
                for f in os.listdir(pasta_saida)
                if f.lower().endswith('.pdf')
            )
            if not pdfs:
                self.log("⚠️ Nenhum PDF encontrado na pasta de saída para publicar")
                return (0, 0)

            self.log(f"🌐 Iniciando publicação em lote: {len(pdfs)} documento(s)")

            if not self._abrir_janela_pub_lote():
                self.log("❌ Não foi possível abrir a janela de Publicação em Lote")
                return (0, len(pdfs))

            pub = self._get_pub_lote_window()
            try:
                pub.set_focus()
            except Exception:
                pass

            # Configuração única do envio
            if not self._configurar_envio_lote(pub, data_vencimento):
                self.log("❌ Falha ao configurar o envio em lote")
                self.cleanup_windows()
                return (0, len(pdfs))

            for caminho_pdf in pdfs:
                if self.should_stop():
                    self.log("Publicação em lote interrompida pelo usuário")
                    break
                self.check_pause()

                # Código = parte antes do primeiro ' - ' no nome do arquivo
                nome_base = os.path.basename(caminho_pdf)
                codigo = nome_base.split(' - ')[0].strip()
                if not codigo:
                    self.log(f"⚠️ Não foi possível extrair o código de: {nome_base}")
                    falhas += 1
                    continue

                if self._publicar_um_documento(pub, caminho_pdf, codigo):
                    publicados += 1
                    self.log(f"✅ Publicado: {nome_base}")
                else:
                    falhas += 1
                    self.log(f"❌ Falha ao publicar: {nome_base}")

            self.log(f"🌐 Publicação em lote concluída: {publicados} publicado(s), {falhas} falha(s)")
            self.cleanup_windows()
            return (publicados, falhas)

        except Exception as e:
            self.log(f"❌ Erro na publicação em lote: {e}\n{traceback.format_exc()}")
            try:
                self.cleanup_windows()
            except Exception:
                pass
            return (publicados, falhas)

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
