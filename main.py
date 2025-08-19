# =============================================================================
# --- ARQUIVO: main.py ---
# (Arquivo principal que inicia e gerencia a aplicação)
# =============================================================================

import os
import sys
import tkinter as tk
from tkinter import messagebox
import configparser
import logging
from datetime import datetime
import ttkbootstrap as ttk
from PIL import Image, ImageTk
import threading
import requests

# Importa as classes de UI dos outros arquivos
from ui_main_frames import (HomeFrame, ExtractionToolsFrame, NFeToolFrame, 
                            WhatsNewFrame)
from ui_windows import (SettingsWindow, ValidatorWindow, KeyParserWindow)
import core_logic


# --- CONSTANTES GLOBAIS ---
APP_NAME = "FiscalFlowCustoms"
CURRENT_VERSION = "2.3"
RELEASE_NOTES = """
Versão 2.3 - Sistema de Atualização Automática

- Verificação de Versão: O programa agora verifica online se há uma nova versão disponível ao ser iniciado.
- Notificação de Atualização: Uma janela informa o usuário sobre a nova versão e pergunta se deseja atualizar.
- Download e Instalação: O programa pode baixar o novo instalador e iniciá-lo para atualizar a si mesmo.
- Lógica de Rede: Adicionada a biblioteca 'requests' para comunicação com a internet.
"""
REPO_OWNER = "seu_usuario_github"
REPO_NAME = "seu_repositorio"

# --- CONFIGURAÇÃO DO LOG ---
try:
    log_dir = ""
    if getattr(sys, 'frozen', False):
        app_data_path = os.environ.get('APPDATA', os.path.expanduser('~'))
        log_dir = os.path.join(app_data_path, APP_NAME)
    else:
        log_dir = os.path.abspath(os.path.dirname(__file__))
    if not os.path.exists(log_dir):
        os.makedirs(log_dir)
    log_file_path = os.path.join(log_dir, 'debug.log')
    logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s', filename=log_file_path, filemode='w')
except Exception as e:
    messagebox.showerror("Erro Crítico de Log", f"Não foi possível criar o arquivo de log.\nErro: {e}")

# --- CLASSE PRINCIPAL DA APLICAÇÃO ---
class App(ttk.Window):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        
        self.CURRENT_VERSION = CURRENT_VERSION
        self.RELEASE_NOTES = RELEASE_NOTES

        logging.info(f"Iniciando {APP_NAME} V{self.CURRENT_VERSION}.")
        self.title(f"Fiscal Flow - Customs V{self.CURRENT_VERSION}")
        self.withdraw()

        self.app_config = configparser.ConfigParser()
        self.config_path = os.path.join(log_dir, 'config.ini')
        
        self.app_style = self.style
        self.load_config()
        
        self.create_menu()

        container = ttk.Frame(self)
        container.pack(side="top", fill="both", expand=True)
        container.grid_rowconfigure(0, weight=1)
        container.grid_columnconfigure(0, weight=1)

        self.frames = {}
        # 1. Cria os frames que não precisam de argumentos extras
        for F in (HomeFrame, ExtractionToolsFrame, NFeToolFrame):
            page_name = F.__name__
            frame = F(container, self)
            self.frames[page_name] = frame
            frame.grid(row=0, column=0, sticky="nsew")

        # 2. Cria o WhatsNewFrame separadamente, passando os argumentos necessários
        page_name = WhatsNewFrame.__name__
        frame = WhatsNewFrame(container, self, 
                              current_version=self.CURRENT_VERSION, 
                              release_notes=self.RELEASE_NOTES)
        self.frames[page_name] = frame
        frame.grid(row=0, column=0, sticky="nsew")
        
        self.start_splash_screen()

    def start_splash_screen(self):
        splash = ttk.Toplevel(self)
        splash.overrideredirect(True)
        w, h = 400, 250
        ws, hs = self.winfo_screenwidth(), self.winfo_screenheight()
        x, y = (ws/2) - (w/2), (hs/2) - (h/2)
        splash.geometry(f'{w}x{h}+{int(x)}+{int(y)}')
        try:
            pil_image = Image.open(self.resource_path('logo_splash.png')).resize((120, 120), Image.LANCZOS)
            self.logo_image_splash = ImageTk.PhotoImage(pil_image, master=splash)
            logo_label = tk.Label(splash, image=self.logo_image_splash, bd=0)
            logo_label.pack(pady=(20, 10))
        except Exception as e:
            logging.warning(f"Não foi possível carregar a logo para a tela de splash: {e}")
            ttk.Label(splash, text="Logo não encontrada").pack(pady=(20, 10))

        ttk.Label(splash, text="Fiscal Flow - Customs", font=("Helvetica", 16, "bold")).pack()
        ttk.Label(splash, text=f"Carregando Versão {self.CURRENT_VERSION}...", font=("Helvetica", 10)).pack(pady=5)
        self.after(3000, lambda: self.process_after_splash(splash))
    
    def process_after_splash(self, splash):
        splash.destroy()
        self.deiconify()
        self.state('zoomed')
        last_version = self.app_config.get('General', 'last_version', fallback='0.0')
        if self.CURRENT_VERSION > last_version:
            self.show_frame("WhatsNewFrame")
        else:
            self.show_frame("HomeFrame")
        logging.info("Interface principal criada. Exibindo janela.")
        
        # Inicia a verificação de atualização em segundo plano
        self.check_for_updates_on_startup()

    def show_frame(self, page_name):
        frame = self.frames[page_name]
        frame.tkraise()

    def resource_path(self, relative_path):
        try:
            base_path = sys._MEIPASS
        except Exception:
            if getattr(sys, 'frozen', False):
                base_path = os.path.dirname(sys.executable)
            else:
                base_path = os.path.abspath(os.path.dirname(__file__))
        return os.path.join(base_path, relative_path)
    
    def load_config(self):
        self.app_config.read(self.config_path)
        self.theme = self.app_config.get('Preferences', 'theme', fallback='superhero')
        self.confirm_on_exit = self.app_config.getboolean('Preferences', 'confirm_on_exit', fallback=True)
        self.ask_to_open_excel = self.app_config.getboolean('Preferences', 'ask_to_open_excel', fallback=True)
        self.default_xml_path = self.app_config.get('Paths', 'default_xml_path', fallback='')
        self.default_output_path = self.app_config.get('Paths', 'default_output_path', fallback='')
        self.output_filename_pattern = self.app_config.get('Paths', 'output_filename_pattern', fallback='Relatorio_NFe_{data}')
        self.app_style.theme_use(self.theme)

    def save_config(self):
        if not self.app_config.has_section('General'): self.app_config.add_section('General')
        self.app_config.set('General', 'last_version', self.CURRENT_VERSION)
        if not self.app_config.has_section('Preferences'): self.app_config.add_section('Preferences')
        self.app_config.set('Preferences', 'theme', self.app_style.theme.name)
        self.app_config.set('Preferences', 'confirm_on_exit', str(self.confirm_on_exit))
        self.app_config.set('Preferences', 'ask_to_open_excel', str(self.ask_to_open_excel))
        if not self.app_config.has_section('Paths'): self.app_config.add_section('Paths')
        self.app_config.set('Paths', 'default_xml_path', self.default_xml_path)
        self.app_config.set('Paths', 'default_output_path', self.default_output_path)
        self.app_config.set('Paths', 'output_filename_pattern', self.output_filename_pattern)
        with open(self.config_path, 'w') as configfile:
            self.app_config.write(configfile)
        logging.info("Configurações salvas.")

    def change_theme(self, theme_name):
        self.app_style.theme_use(theme_name)
        for frame in self.frames.values():
            if hasattr(frame, 'update_watermarks'):
                frame.update_watermarks()
        self.save_config()

    def create_menu(self):
        menubar = ttk.Menu(self)
        self.config(menu=menubar)
        file_menu = ttk.Menu(menubar, tearoff=False)
        menubar.add_cascade(label="Arquivo", menu=file_menu)
        settings_menu = ttk.Menu(menubar, tearoff=False)
        menubar.add_cascade(label="Configurações", menu=settings_menu)
        settings_menu.add_command(label="Preferências Gerais...", command=self.open_settings_window)
        theme_menu = ttk.Menu(settings_menu, tearoff=False)
        settings_menu.add_cascade(label="Tema", menu=theme_menu)
        themes = self.app_style.theme_names()
        themes.sort()
        for theme_name in themes:
            theme_menu.add_command(label=theme_name.capitalize(), command=lambda t=theme_name: self.change_theme(t))
        utils_menu = ttk.Menu(menubar, tearoff=False)
        menubar.add_cascade(label="Utilitários", menu=utils_menu)
        utils_menu.add_command(label="Validador de CNPJ/CPF", command=self.open_validator_window)
        utils_menu.add_command(label="Analisador de Chave de NFe", command=self.open_key_parser_window)
        help_menu = ttk.Menu(menubar, tearoff=False)
        menubar.add_cascade(label="Ajuda", menu=help_menu)
        help_menu.add_command(label="Verificar Atualizações...", command=self.check_for_updates_on_startup)
        help_menu.add_command(label="Notas da Versão", command=lambda: self.show_frame("WhatsNewFrame"))
        help_menu.add_command(label="Abrir Pasta de Logs", command=self.open_log_folder)
        help_menu.add_separator()
        help_menu.add_command(label="Sobre...", command=self.show_about)
        menubar.add_command(label="Sair", command=self.quit_app)

    def open_settings_window(self):
        if not hasattr(self, 'settings_win') or not self.settings_win.winfo_exists():
            self.settings_win = SettingsWindow(self)
        else:
            self.settings_win.lift()

    def open_validator_window(self):
        if not hasattr(self, 'validator_win') or not self.validator_win.winfo_exists():
            self.validator_win = ValidatorWindow(self)
        else:
            self.validator_win.lift()

    def open_key_parser_window(self):
        if not hasattr(self, 'key_parser_win') or not self.key_parser_win.winfo_exists():
            self.key_parser_win = KeyParserWindow(self)
        else:
            self.key_parser_win.lift()
        
    def open_log_folder(self):
        try:
            os.startfile(log_dir)
        except Exception as e:
            messagebox.showerror("Erro", f"Não foi possível abrir a pasta de logs.\nCaminho: {log_dir}\nErro: {e}")

    def show_about(self):
        messagebox.showinfo(
            title="Sobre o Fiscal Flow",
            message=f"Fiscal Flow - Customs V{self.CURRENT_VERSION}\n\nDesenvolvido por Bruno Silva - Analista Contábil\nAssistente de IA: Órion"
        )
    
    def quit_app(self):
        if self.confirm_on_exit:
            if messagebox.askyesno("Sair", "Tem certeza que deseja sair do programa?"):
                self.destroy()
        else:
            self.destroy()

    def pulse_and_navigate(self, button, target_frame_name):
        original_style = button.cget('style')
        pulse_style = 'success.TButton'
        if 'Outline' in original_style: pulse_style = 'success.Outline.TButton'
        if 'Large' in original_style: pulse_style = 'Large.' + pulse_style
        button.config(state="disabled", style=pulse_style)
        def action():
            self.show_frame(target_frame_name)
            button.config(state="normal", style=original_style)
        self.after(75, action)

    def check_for_updates_on_startup(self):
        # Roda a verificação em uma thread separada para não travar a interface
        threading.Thread(target=self._check_for_updates_thread, daemon=True).start()

    def _check_for_updates_thread(self):
        print("--- DEBUG: Verificando atualizações... ---")
        update_info = core_logic.check_for_updates(self.CURRENT_VERSION, REPO_OWNER, REPO_NAME)
        
        if update_info and update_info["update_available"]:
            print(f"--- DEBUG: Atualização encontrada: {update_info['latest_version']} ---")
            # Usa self.after para garantir que a messagebox seja chamada na thread principal da GUI
            self.after(0, self.show_update_notification, update_info)
        else:
            print("--- DEBUG: Nenhuma atualização encontrada. ---")
            
    def show_update_notification(self, update_info):
        # (Esta função será implementada na Fase 2)
        messagebox.showinfo("Atualização Disponível", 
                            f"Uma nova versão ({update_info['latest_version']}) do Fiscal Flow está disponível!",
                            parent=self)


# --- BLOCO DE EXECUÇÃO PRINCIPAL ---
if __name__ == "__main__":
    try:
        app = App(themename="superhero")
        app.mainloop()
    except Exception as e:
        logging.critical("Uma exceção não tratada ocorreu e a aplicação será fechada.", exc_info=True)
        try:
            root_fallback = tk.Tk()
            root_fallback.withdraw()
            messagebox.showerror("Erro Fatal", f"Um erro crítico ocorreu e a aplicação precisa fechar.\n\nVerifique o arquivo debug.log.\n\n{e}")
        except:
            pass