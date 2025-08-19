import sys
from cx_Freeze import setup, Executable

# --- Configurações do Build ---
# ADICIONADO 'requests' à lista de pacotes
packages = ["lxml", "openpyxl", "PIL", "ttkbootstrap", "configparser", "requests"]

# Inclui todos os arquivos de imagem necessários
include_files = [
    "logo.ico",
    "logo_light.png", 
    "logo_dark.png",
    "logo_text_light.png", 
    "logo_text_dark.png",
    "logo_splash.png"
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
        target_name="FiscalFlowCustoms.exe",
        icon="logo.ico"
    )
]

# --- Comando Final de Setup ---
setup(
    name="Fiscal Flow - Customs",
    version="2.3", # Versão da nova atualização
    description="Snapshot 2.3 - Sistema de Atualização Automática",
    options={"build_exe": build_exe_options},
    executables=executables
)