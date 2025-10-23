# =============================================================================
# --- ARQUIVO: main.py (INTEGRADO COM PAINEL DE ATENDIMENTO) ---
# (Versão atualizada para 2.9.2 e Release Notes adicionadas)
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

try:
    from cryptography.fernet import Fernet, InvalidToken
except ImportError:
    messagebox.showerror(
        "Dependência Faltando",
        "A biblioteca 'cryptography' não foi encontrada. "
        "Por favor, instale-a executando: pip install cryptography"
    )
    sys.exit(1)


from ui.frames_auth import (LoginFrame, SetPasswordFrame, ResetPasswordFrame, WhatsNewFrame)
from ui.frames_app import (HomeFrame, ExtractionToolsFrame, NFeToolFrame)
from ui.frames_dp import (DPMainFrame, DPLançamentosToolFrame, 
                          DPLancamentoColabFrame, DPLancamentoRubricaFrame)
from ui.frames_support import (SupportChoiceFrame, UserTicketsFrame, AdminTicketsFrame)

from ui.dialogs_user import (ChangePasswordDialog, RequestAccessDialog, 
                             RequestsWindow, UserManagementWindow, SectorManagementWindow,
                             TemplateEditorWindow, CommunicationWindow)
from ui.dialogs_flow import (UpdateDownloadWindow, DashboardWindow, PreviewWindow)
from ui.dialogs_tools import (SettingsWindow, ValidatorWindow, KeyParserWindow)
from ui.dialogs_dev import (TestRunnerWindow)
from ui.dialogs_clients import (ClientCodeManagerWindow, ClientCodeReportWindow)


import core_logic
import auth_logic
import dp_logic
import support_logic
import client_logic
import report_logic 
# --- [CORRIGIDO] --- Adiciona a importação que estava faltando ---
import drive_logic

# --- CONSTANTES GLOBAIS ---
APP_NAME = "CustomsFlow"
CURRENT_VERSION = "2.9.2" # <--- VERSÃO ATUALIZADA
RELEASE_NOTES = """
Versão 2.9.2 - Correções Críticas na Extração e Geração de Excel

🐛 CORREÇÃO: Resolvido erro "formatCode should be NoneType" que impedia salvar planilhas Excel, garantindo a aplicação correta de formatos de número/texto em todas as células (afetava `core_logic.py` e `frames_app.py`).
🐛 CORREÇÃO: Corrigido erro de sintaxe "invalid non-printable character U+00A0" que podia impedir a inicialização do programa (causado por erro na geração de código anterior).
🐛 CORREÇÃO: Aprimorada a extração do 'Nome do Processo' em NFe. A lógica agora identifica corretamente o final do número do processo (baseado no ano) mesmo quando há texto adicional colado (ex: "...2025ICMS..."), evitando que "lixo" seja incluído (em `core_logic.py`).
🐛 CORREÇÃO: Coluna 'Valor Serviço Trading' restaurada na janela de Pré-Visualização (abas Entrada/Saída), onde havia sido omitida acidentalmente (em `dialogs_flow.py`).

--- VERSÕES ANTERIORES ---

Versão 2.9.1 - Unificação de Suporte e Melhorias de Usabilidade

✨ NOVO: Agora, os módulos de Suporte de T.I. e Suporte do Sistema possuem as mesmas funcionalidades, incluindo a triagem de chamados por prioridade (cores).

🔧 MELHORIA: A janela do "Gerenciador de Códigos de Cliente" agora pode ser minimizada e maximizada como uma janela normal.

🔧 MELHORIA: A funcionalidade de exclusão no Gerenciador de Códigos foi aprimorada e agora também está disponível através do menu de clique-direito.

🔧 MELHORIA: Atalhos para os dois módulos de suporte foram adicionados à tela principal, com ícones que se adaptam dinamicamente à cor do tema.

🐛 CORREÇÃO: A confirmação de saída agora funciona corretamente também ao clicar no botão "X" da janela principal.
"""
REPO_OWNER = "brunosilva706799-arch"
REPO_NAME = "CustomsFlow"

# (Configuração de Log)
try:
    log_dir = ""
    if getattr(sys, 'frozen', False):
        app_data_path = os.environ.get('APPDATA', os.path.expanduser('~'))
        log_dir = os.path.join(app_data_path, APP_NAME)
    else:
        log_dir = os.path.abspath(os.path.dirname(__file__))
    if not os.path.exists(log_dir): os.makedirs(log_dir)
    log_file_path = os.path.join(log_dir, 'debug.log')
    logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s', filename=log_file_path, filemode='w')
except Exception as e:
    messagebox.showerror("Erro Crítico de Log", f"Não foi possível criar o arquivo de log.\nErro: {e}")

class App(ttk.Window):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        
        auth_logic.set_resource_path_getter(self.resource_path)
        drive_logic.set_resource_path_getter(self.resource_path)

        try: 
            auth_logic.initialize_firebase()
        except Exception as e: 
            messagebox.showerror("Erro Crítico de Conexão", f"Não foi possível conectar ao Firebase.\n\nErro: {e}")
            self.destroy()
            return
            
        self.current_user = None; self.CURRENT_VERSION = CURRENT_VERSION; self.RELEASE_NOTES = RELEASE_NOTES
        self.image_cache = {}; self.cache = {"companies": None, "payroll_codes": None, "sectors": None, "employees": {}}
        logging.info(f"Iniciando {APP_NAME} V{self.CURRENT_VERSION}.")
        self.title(f"Customs Flow")
        self.withdraw()
        self.app_config = configparser.ConfigParser(); self.config_path = os.path.join(log_dir, 'config.ini')
        self.app_style = self.style; self.load_config()
        container = ttk.Frame(self); container.pack(side="top", fill="both", expand=True)
        container.grid_rowconfigure(0, weight=1); container.grid_columnconfigure(0, weight=1)
        self.frames = {}
        
        auth_frames = (LoginFrame, SetPasswordFrame, ResetPasswordFrame)
        app_frames = (HomeFrame, ExtractionToolsFrame, NFeToolFrame)
        dp_frames = (DPMainFrame, DPLançamentosToolFrame, DPLancamentoColabFrame, DPLancamentoRubricaFrame)
        support_frames = (SupportChoiceFrame, UserTicketsFrame, AdminTicketsFrame)

        for F in auth_frames + app_frames + dp_frames + support_frames:
            page_name = F.__name__; frame = F(container, self); self.frames[page_name] = frame; frame.grid(row=0, column=0, sticky="nsew")

        page_name = WhatsNewFrame.__name__
        frame = WhatsNewFrame(container, self, current_version=self.CURRENT_VERSION, release_notes=self.RELEASE_NOTES)
        self.frames[page_name] = frame
        frame.grid(row=0, column=0, sticky="nsew")

        self.protocol("WM_DELETE_WINDOW", self.quit_app)
        
        self.start_splash_screen()
    
    def get_from_cache(self, key, sub_key=None, logic_function=None, force_refresh=False):
        cache_target = self.cache[key]; data_to_check = cache_target.get(sub_key) if sub_key else cache_target
        if data_to_check is None or force_refresh:
            data, error = logic_function(sub_key) if sub_key else logic_function()
            if error: messagebox.showerror("Erro de Conexão", error, parent=self); return [], error
            if sub_key: self.cache[key][sub_key] = data
            else: self.cache[key] = data
            return data, None
        else: return data_to_check, None
    def invalidate_cache(self, key, sub_key=None):
        if sub_key:
            if sub_key in self.cache[key]: del self.cache[key][sub_key]
        else: self.cache[key] = None
    def get_companies(self, force_refresh=False): return self.get_from_cache("companies", logic_function=dp_logic.get_all_companies, force_refresh=force_refresh)
    def get_employees(self, company_id, force_refresh=False): return self.get_from_cache("employees", sub_key=company_id, logic_function=dp_logic.get_employees_for_company, force_refresh=force_refresh)
    def get_payroll_codes(self, force_refresh=False): return self.get_from_cache("payroll_codes", logic_function=dp_logic.get_all_payroll_codes, force_refresh=force_refresh)
    def get_sectors(self, force_refresh=False): return self.get_from_cache("sectors", logic_function=auth_logic.get_all_sectors, force_refresh=force_refresh)
    def start_splash_screen(self):
        splash = ttk.Toplevel(self); splash.overrideredirect(True)
        w, h = 400, 250; ws, hs = self.winfo_screenwidth(), self.winfo_screenheight(); x, y = (ws/2) - (w/2), (hs/2) - (h/2)
        splash.geometry(f'{w}x{h}+{int(x)}+{int(y)}')
        try:
            pil_image = Image.open(self.resource_path('logo_splash.png')).resize((120, 120), Image.LANCZOS)
            self.logo_image_splash = ImageTk.PhotoImage(pil_image, master=splash)
            ttk.Label(splash, image=self.logo_image_splash).pack(pady=(20, 10))
        except Exception: ttk.Label(splash, text="Logo não encontrada").pack(pady=(20, 10))
        ttk.Label(splash, text="Customs Flow", font=("Helvetica", 16, "bold")).pack()
        ttk.Label(splash, text=f"Carregando Versão {self.CURRENT_VERSION}...", font=("Helvetica", 10)).pack(pady=5)
        self.after(2000, lambda: self.process_after_splash(splash))
    def process_after_splash(self, splash):
        splash.destroy(); self.deiconify(); self.show_frame("LoginFrame")
    def on_login_success(self, user_data):
        self.current_user = user_data; self.state('zoomed'); self.title(f"Customs Flow V{self.CURRENT_VERSION}"); self.create_menu()
        self.frames['HomeFrame'].update_user_display(self.current_user['username'])
        last_version = self.app_config.get('General', 'last_version', fallback='0.0')
        if self.CURRENT_VERSION > last_version: self.show_frame("WhatsNewFrame")
        else: self.show_main_app()
    def logout(self):
        self.current_user = None; self.config(menu=tk.Menu(self)); self.cache = {"companies": None, "employees": {}, "payroll_codes": None, "sectors": None}
        w, h = 600, 500; ws, hs = self.winfo_screenwidth(), self.winfo_screenheight(); x, y = (ws/2) - (w/2), (hs/2) - (h/2)
        self.geometry(f'{w}x{h}+{int(x)}+{int(y)}'); self.show_frame("LoginFrame")
    def show_main_app(self):
        self.show_frame("HomeFrame"); self.check_for_updates_on_startup(manual_check=False)
    def show_frame(self, page_name):
        frame = self.frames[page_name]
        if hasattr(frame, 'on_frame_activated'): frame.on_frame_activated()
        frame.tkraise()
    def show_support_tickets(self, ticket_type):
        user_level = self.current_user.get('level'); frame_to_show_name = "UserTicketsFrame"
        if ticket_type == 'developer' and user_level in ['Desenvolvedor', 'Admin']: frame_to_show_name = "AdminTicketsFrame"
        elif ticket_type == 'it' and user_level in ['T.I.', 'Admin', 'Desenvolvedor']: frame_to_show_name = "AdminTicketsFrame"
        frame = self.frames[frame_to_show_name]; frame.set_ticket_type(ticket_type); self.show_frame(frame_to_show_name)
    
    def resource_path(self, relative_path):
        """ Retorna o caminho absoluto para o recurso, funcionando para dev e app compilado. """
        try:
            base_path = sys._MEIPASS
        except Exception:
            base_path = os.path.abspath(".")

        if sys.platform == "darwin" and getattr(sys, 'frozen', False):
             return os.path.join(base_path, '..', 'Resources', relative_path)
             
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
        for section in ['General', 'Preferences', 'Paths']:
            if not self.app_config.has_section(section): self.app_config.add_section(section)
        self.app_config.set('General', 'last_version', self.CURRENT_VERSION)
        self.app_config.set('Preferences', 'theme', self.app_style.theme.name)
        self.app_config.set('Preferences', 'confirm_on_exit', str(self.confirm_on_exit))
        self.app_config.set('Preferences', 'ask_to_open_excel', str(self.ask_to_open_excel))
        self.app_config.set('Paths', 'default_xml_path', self.default_xml_path)
        self.app_config.set('Paths', 'default_output_path', self.default_output_path)
        self.app_config.set('Paths', 'output_filename_pattern', self.output_filename_pattern)
        with open(self.config_path, 'w') as configfile: self.app_config.write(configfile)
    def change_theme(self, theme_name):
        self.app_style.theme_use(theme_name); self.image_cache.clear()
        for frame in self.frames.values():
            if hasattr(frame, 'update_watermarks'): frame.update_watermarks()
        self.save_config()
    def create_menu(self):
        menubar = ttk.Menu(self); self.config(menu=menubar)
        user_level = self.current_user.get('level'); client_code_access = self.current_user.get('acesso_codigos_cliente', 'Nenhum')
        if user_level in ['Desenvolvedor', 'Admin']:
            admin_menu = ttk.Menu(menubar, tearoff=False); menubar.add_cascade(label="Administração", menu=admin_menu)
            admin_menu.add_command(label="Gerenciar Solicitações...", command=self.open_requests_window)
            admin_menu.add_command(label="Gerenciar Usuários...", command=self.open_user_management_window)
            admin_menu.add_command(label="Gerenciar Setores...", command=self.open_sector_management_window)
            admin_menu.add_separator()
            admin_menu.add_command(label="Enviar Comunicado Global...", command=self.open_communication_window)
            admin_menu.add_command(label="Editar Templates de E-mail...", command=self.open_template_editor_window)
            if user_level == 'Desenvolvedor':
                admin_menu.add_separator(); admin_menu.add_command(label="Executar Auto-teste...", command=self.open_test_runner)
        file_menu = ttk.Menu(menubar, tearoff=False); menubar.add_cascade(label="Arquivo", menu=file_menu)
        if client_code_access in ['Consulta', 'Total'] or user_level == 'Desenvolvedor':
            file_menu.add_command(label="Gerenciador de Códigos...", command=self.open_client_code_manager); file_menu.add_separator()
        file_menu.add_command(label="Alterar Minha Senha...", command=self.open_change_password_dialog); file_menu.add_separator()
        file_menu.add_command(label="Sair", command=self.quit_app)
        if client_code_access in ['Consulta', 'Total'] or user_level == 'Desenvolvedor':
            reports_menu = ttk.Menu(menubar, tearoff=False); menubar.add_cascade(label="Relatórios", menu=reports_menu)
            reports_menu.add_command(label="Relatório de Códigos de Clientes...", command=self.open_client_code_report)
        support_menu = ttk.Menu(menubar, tearoff=False); menubar.add_cascade(label="Suporte", menu=support_menu)
        support_menu.add_command(label="Abrir Central de Suporte...", command=lambda: self.show_frame("SupportChoiceFrame"))
        settings_menu = ttk.Menu(menubar, tearoff=False); menubar.add_cascade(label="Configurações", menu=settings_menu)
        settings_menu.add_command(label="Preferências Gerais...", command=self.open_settings_window)
        theme_menu = ttk.Menu(settings_menu, tearoff=False); settings_menu.add_cascade(label="Tema", menu=theme_menu)
        themes = self.app_style.theme_names(); themes.sort()
        for theme_name in themes: theme_menu.add_command(label=theme_name.capitalize(), command=lambda t=theme_name: self.change_theme(t))
        utils_menu = ttk.Menu(menubar, tearoff=False); menubar.add_cascade(label="Utilitários", menu=utils_menu)
        utils_menu.add_command(label="Validador de CNPJ/CPF", command=self.open_validator_window)
        utils_menu.add_command(label="Analisador de Chave de NFe", command=self.open_key_parser_window)
        help_menu = ttk.Menu(menubar, tearoff=False); menubar.add_cascade(label="Ajuda", menu=help_menu)
        help_menu.add_command(label="Verificar Atualizações...", command=lambda: self.check_for_updates_on_startup(manual_check=True))
        help_menu.add_command(label="Notas da Versão", command=lambda: self.show_frame("WhatsNewFrame"))
        help_menu.add_command(label="Abrir Pasta de Logs", command=self.open_log_folder); help_menu.add_separator()
        help_menu.add_command(label="Sobre...", command=self.show_about)
    def open_client_code_manager(self):
        if not hasattr(self, 'client_code_win') or not self.client_code_win.winfo_exists(): self.client_code_win = ClientCodeManagerWindow(self, self)
        else: self.client_code_win.lift()
    def open_client_code_report(self):
        if not hasattr(self, 'client_report_win') or not self.client_report_win.winfo_exists(): self.client_report_win = ClientCodeReportWindow(self, self)
        else: self.client_report_win.lift()
    def open_requests_window(self):
        if not hasattr(self, 'requests_win') or not self.requests_win.winfo_exists(): self.requests_win = RequestsWindow(self)
        else: self.requests_win.lift()
    def open_user_management_window(self):
        if not hasattr(self, 'user_mgmt_win') or not self.user_mgmt_win.winfo_exists(): self.user_mgmt_win = UserManagementWindow(self)
        else: self.user_mgmt_win.lift()
    def open_sector_management_window(self):
        if not hasattr(self, 'sector_mgmt_win') or not self.sector_mgmt_win.winfo_exists(): self.sector_mgmt_win = SectorManagementWindow(self)
        else: self.sector_mgmt_win.lift()
    def open_communication_window(self):
        if not hasattr(self, 'comm_win') or not self.comm_win.winfo_exists(): self.comm_win = CommunicationWindow(self)
        else: self.comm_win.lift()
    def open_template_editor_window(self):
        if not hasattr(self, 'template_win') or not self.template_win.winfo_exists(): self.template_win = TemplateEditorWindow(self)
        else: self.template_win.lift()
    def open_change_password_dialog(self):
        if self.current_user: ChangePasswordDialog(self, self.current_user['id'])
        else: messagebox.showerror("Erro", "Nenhum usuário logado.")
    def open_test_runner(self):
        if not hasattr(self, 'test_win') or not self.test_win.winfo_exists(): self.test_win = TestRunnerWindow(self)
        else: self.test_win.lift()
    def open_settings_window(self):
        if not hasattr(self, 'settings_win') or not self.settings_win.winfo_exists(): self.settings_win = SettingsWindow(self)
        else: self.settings_win.lift()
    def open_validator_window(self):
        if not hasattr(self, 'validator_win') or not self.validator_win.winfo_exists(): self.validator_win = ValidatorWindow(self)
        else: self.validator_win.lift()
    def open_key_parser_window(self):
        if not hasattr(self, 'key_parser_win') or not self.key_parser_win.winfo_exists(): self.key_parser_win = KeyParserWindow(self)
        else: self.key_parser_win.lift()
    def open_request_access_dialog(self):
        RequestAccessDialog(self)
    def open_log_folder(self):
        try: os.startfile(log_dir)
        except Exception as e: messagebox.showerror("Erro", f"Não foi possível abrir a pasta.\n{e}")
    def show_about(self):
        messagebox.showinfo(title="Sobre o Customs Flow", message=(f"Customs Flow V{self.CURRENT_VERSION}\n\n" "Desenvolvido por Bruno Silva\n" "Assistente de IA: Órion\n\n" f"Repositório do Projeto:\n" f"github.com/{REPO_OWNER}/{REPO_NAME}"))
    def quit_app(self):
        if self.confirm_on_exit:
            if messagebox.askyesno("Sair", "Tem certeza que deseja sair?"): self.destroy()
        else: self.destroy()
    def pulse_and_navigate(self, button, target_frame_name):
        original_style = button.cget('style'); pulse_style = 'success.TButton'
        if 'Outline' in original_style: pulse_style = 'success.Outline.TButton'
        if 'Large' in original_style: pulse_style = 'Large.' + pulse_style
        button.config(state="disabled", style=pulse_style)
        def action(): self.show_frame(target_frame_name); button.config(state="normal", style=original_style)
        self.after(75, action)
    def check_for_updates_on_startup(self, manual_check=False):
        threading.Thread(target=self._check_for_updates_thread, args=(manual_check,), daemon=True).start()
    def _check_for_updates_thread(self, manual_check):
        update_info = core_logic.check_for_updates(self.CURRENT_VERSION, REPO_OWNER, REPO_NAME)
        if update_info and update_info["update_available"]: self.after(0, self.show_update_notification, update_info)
        elif manual_check: self.after(0, self.show_no_updates_found_message)
    def show_update_notification(self, update_info):
        version, notes = update_info['latest_version'], update_info['release_notes']
        title = f"Atualização Disponível: Versão {version}"; message = f"Uma nova versão do Customs Flow está disponível!\n\n{notes}\n\nDeseja baixar e instalar agora?"
        if messagebox.askyesno(title, message, parent=self): self.start_download_process(update_info)
    def start_download_process(self, update_info):
        download_url, latest_version = update_info.get("download_url"), update_info.get("latest_version")
        if not download_url: messagebox.showerror("Erro", "URL de download não encontrado."); return
        download_window = UpdateDownloadWindow(self, latest_version)
        def p_cb(d, t): download_window.after(0, download_window.update_progress, d, t)
        def c_cb(f): download_window.after(0, download_window.on_download_complete, f)
        def e_cb(e): download_window.after(0, download_window.on_download_error, e)
        threading.Thread(target=core_logic.download_update, args=(download_url, p_cb, c_cb, e_cb), daemon=True).start()
    def show_no_updates_found_message(self):
        messagebox.showinfo("Verificar Atualizações", "Você já está usando a versão mais recente.", parent=self)

if __name__ == "__main__":
    try:
        app = App(themename="superhero")
        app.mainloop()
    except Exception as e:
        logging.critical("Uma exceção não tratada ocorreu.", exc_info=True)
        try:
            root_fallback = tk.Tk()
            root_fallback.withdraw()
            messagebox.showerror("Erro Fatal", f"Um erro crítico ocorreu.\n\nVerifique o debug.log.\n\n{e}")
        except: pass