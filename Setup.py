import sys
from cx_Freeze import setup, Executable

# --- Configurações do Build ---
packages = [
    "lxml", "openpyxl", "PIL", "ttkbootstrap", "configparser", 
    "requests", "firebase_admin", "bcrypt", "cryptography", "pytz",
    "pandas",         # Para importação de planilhas
    "reportlab",      # Para exportação em PDF
    "docx",           # Para exportação em Word
    # --- [ADICIONADO] Bibliotecas do Google para o Drive ---
    "googleapiclient", 
    "google_auth_oauthlib",
    "google_auth_httplib2"
]

include_files = [
    "logo.ico",
    "logo_light.png", 
    "logo_dark.png",
    "logo_text_light.png", 
    "logo_text_dark.png",
    "logo_splash.png",
    "firebase_credentials.json",
    "google_drive_credentials.json", # Essencial para os anexos do suporte
    "test_assets/"
]

build_exe_options = {
    "packages": packages,
    "include_files": include_files,
}

# --- Base da Interface Gráfica ---
base = None
if sys.platform == "win32":
    base = "Win32GUI"

# --- Definição do Executável ---
executables = [
    Executable(
        "main.py",
        base=base,
        target_name="CustomsFlow.exe", # <-- [ALTERADO]
        icon="logo.ico"
    )
]

# --- Comando Final de Setup ---
setup(
    name="Customs Flow", # <-- [ALTERADO]
    version="2.9.1",      # <-- [ALTERADO]
    description="Ferramenta de Comércio Exterior com módulos de Extração, DP e Suporte.",
    options={"build_exe": build_exe_options},
    executables=executables
)