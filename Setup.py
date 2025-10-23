import sys
from cx_Freeze import setup, Executable

# --- Configurações do Build ---
packages = [
    # Bibliotecas de terceiros
    "lxml", "openpyxl", "PIL", "ttkbootstrap", "configparser", 
    "requests", "firebase_admin", "bcrypt", "cryptography", "pytz",
    "pandas", "reportlab", "docx",
    "googleapiclient", "google_auth_oauthlib", "google_auth_httplib2",
]

include_files = [
    "logo.ico", "logo_light.png", "logo_dark.png",
    "logo_text_light.png", "logo_text_dark.png", "logo_splash.png",
    "firebase_credentials.json", "google_drive_credentials.json",
    "test_assets/"
]

# --- [CORRIGIDO] Adiciona a opção 'includes' para forçar a inclusão dos módulos ---
build_exe_options = {
    "packages": packages,
    "include_files": include_files,
    "includes": [
        "core_logic",
        "auth_logic",
        "drive_logic",
        "dp_logic",
        "client_logic",
        "support_logic",
        "report_logic"
    ],
}

# --- Base da Interface Gráfica ---
base = None
if sys.platform == "win32":
    base = "Win32GUI"

# --- Definição do Executável ---
target_name = "CustomsFlow.exe" if sys.platform == "win32" else "CustomsFlow"

executables = [
    Executable(
        "main.py",
        base=base,
        target_name=target_name,
        icon="logo.ico"
    )
]

# --- Comando Final de Setup ---
setup(
    name="Customs Flow",
    version="2.9.2", # <--- VERSÃO ATUALIZADA
    description="Ferramenta de Comércio Exterior com módulos de Extração, DP e Suporte.",
    options={"build_exe": build_exe_options},
    executables=executables
)