# =============================================================================
# --- ARQUIVO: drive_logic.py ---
# (Lida com a autenticação OAuth 2.0 e a comunicação com a API do Google Drive)
# =============================================================================

import os
import pickle
from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload
from googleapiclient.errors import HttpError

# --- [NOVO] Variável global para armazenar a função 'resource_path' ---
_get_resource_path = None

def set_resource_path_getter(getter_func):
    """
    Recebe a função 'resource_path' do main.py para ser usada neste módulo.
    """
    global _get_resource_path
    _get_resource_path = getter_func

# --- CONFIGURAÇÕES ---
SCOPES = ['https://www.googleapis.com/auth/drive']
DRIVE_FOLDER_NAME = 'DataBase CustomsFlow'

FOLDER_ID_CACHE = None
SERVICE_CACHE = None

def get_drive_service():
    """
    Autentica o usuário via OAuth 2.0 e retorna um objeto de serviço para
    interagir com a API. Lida com o fluxo de login via navegador na primeira vez.
    """
    global SERVICE_CACHE
    if SERVICE_CACHE:
        return SERVICE_CACHE, None
        
    creds = None
    
    # --- [MODIFICADO] Usa a nova função 'GPS' para encontrar os arquivos ---
    if not _get_resource_path:
        return None, "Erro crítico: A função de localização de recursos (resource_path) não foi configurada."
        
    token_file = _get_resource_path('token.json')
    credentials_file = _get_resource_path('google_drive_credentials.json')

    if os.path.exists(token_file):
        with open(token_file, 'rb') as token:
            creds = pickle.load(token)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            try:
                creds.refresh(Request())
            except Exception as e:
                print(f"Erro ao atualizar token, solicitando novo login: {e}")
                # Se o refresh falhar, força um novo login limpando as credenciais antigas
                if os.path.exists(token_file):
                    os.remove(token_file)
                creds = None
        
        # Se ainda não tiver credenciais, faz o fluxo de login
        if not creds:
            try:
                flow = InstalledAppFlow.from_client_secrets_file(credentials_file, SCOPES)
                creds = flow.run_local_server(port=0)
            except FileNotFoundError:
                return None, f"Erro Crítico: O arquivo '{os.path.basename(credentials_file)}' não foi encontrado. Por favor, verifique."
            except Exception as e:
                return None, f"Erro ao obter credenciais. Verifique o arquivo de credenciais. Erro: {e}"

        with open(token_file, 'wb') as token:
            pickle.dump(creds, token)
    
    try:
        service = build('drive', 'v3', credentials=creds)
        SERVICE_CACHE = service
        return service, None
    except Exception as e:
        return None, f"Erro ao construir o serviço do Drive: {e}"


def get_folder_id(service):
    """Busca o ID da pasta de anexos. Se não encontrar, cria uma nova."""
    global FOLDER_ID_CACHE
    if FOLDER_ID_CACHE: return FOLDER_ID_CACHE, None
    try:
        response = service.files().list(
            q=f"name='{DRIVE_FOLDER_NAME}' and mimeType='application/vnd.google-apps.folder' and trashed=false",
            spaces='drive', fields='files(id, name)').execute()
        files = response.get('files', [])
        if files:
            FOLDER_ID_CACHE = files[0].get('id')
            return FOLDER_ID_CACHE, None
        else:
            print(f"Pasta '{DRIVE_FOLDER_NAME}' não encontrada. Criando uma nova...")
            file_metadata = {'name': DRIVE_FOLDER_NAME, 'mimeType': 'application/vnd.google-apps.folder'}
            file = service.files().create(body=file_metadata, fields='id').execute()
            FOLDER_ID_CACHE = file.get('id')
            print(f"Pasta criada com ID: {FOLDER_ID_CACHE}")
            return FOLDER_ID_CACHE, None
    except Exception as e:
        return None, f"Ocorreu um erro ao acessar a pasta no Google Drive: {e}"

def upload_attachment(local_file_path, original_filename):
    """Faz o upload de um arquivo para a pasta designada no Google Drive."""
    service, error = get_drive_service()
    if error: return None, None, error

    folder_id, error = get_folder_id(service)
    if error: return None, None, error
        
    try:
        file_metadata = {'name': original_filename, 'parents': [folder_id]}
        media = MediaFileUpload(local_file_path)
        
        file = service.files().create(
            body=file_metadata,
            media_body=media,
            fields='id, webViewLink'
        ).execute()
        
        file_id = file.get('id')
        
        permission = {'type': 'anyone', 'role': 'reader'}
        service.permissions().create(fileId=file_id, body=permission).execute()

        print(f"Arquivo '{original_filename}' enviado com sucesso. Link: {file.get('webViewLink')}")
        return file.get('webViewLink'), original_filename, None
    except Exception as e:
        return None, None, f"Ocorreu um erro ao fazer o upload do anexo: {e}"