# =============================================================================
# --- ARQUIVO: auth_logic.py (ATUALIZADO COM E-MAILS DE SUPORTE) ---
# =============================================================================

import bcrypt
import os
import sys
import smtplib
import ssl
from email.message import EmailMessage
import secrets
from datetime import datetime, timedelta

import firebase_admin
from firebase_admin import credentials, firestore
from google.cloud.firestore_v1.base_query import FieldFilter
from google.api_core import exceptions as google_exceptions

TOKEN_VALIDITY_HOURS = 24
CREDENTIALS_FILE = "firebase_credentials.json"
QUERY_TIMEOUT = 45 
db = None

def initialize_firebase():
    global db
    if firebase_admin._apps: return
    try:
        if getattr(sys, 'frozen', False):
            base_path = os.path.dirname(sys.executable)
        else:
            base_path = os.path.abspath(os.path.dirname(__file__))
        cred_path = os.path.join(base_path, CREDENTIALS_FILE)
        if not os.path.exists(cred_path):
            raise FileNotFoundError(f"Arquivo de credenciais '{CREDENTIALS_FILE}' não encontrado.")
        
        cred = credentials.Certificate(cred_path)
        firebase_admin.initialize_app(cred)
        db = firestore.client()
        
        users_ref = db.collection('users')
        query = users_ref.where(filter=FieldFilter('level', '==', 'Desenvolvedor')).limit(1).get(timeout=QUERY_TIMEOUT)
        if not query:
            password = 'dev'
            hashed_password = bcrypt.hashpw(password.encode('utf-8'), bcrypt.gensalt())
            users_ref.add({
                'username': 'dev', 
                'password_hash': hashed_password.decode('utf-8'), 
                'email': 'dev@fiscalflow.com', 
                'level': 'Desenvolvedor',
                'acesso_codigos_cliente': 'Total'
            })

        templates_to_check = {
            'account_creation': {
                'name': 'E-mail de Criação de Conta',
                'subject': "Configure sua conta no Fiscal Flow",
                'body': """
                <html><body>
                <h2>Olá, {username}!</h2>
                <p>Seu cadastro no sistema Fiscal Flow foi iniciado. Para completar o processo, por favor, defina sua senha de acesso.</p>
                <p>Na tela de login do programa, clique em "Primeiro Acesso / Definir Senha" e utilize o seguinte código de verificação:</p>
                <h1 style="text-align:center; letter-spacing: 5px; font-size: 36px;">{token}</h1>
                <p>Este código é válido por 24 horas.</p>
                <br><p>Atenciosamente,<br>Equipe Fiscal Flow</p>
                </body></html>
                """
            },
            'password_reset': {
                'name': 'E-mail de Redefinição de Senha',
                'subject': "Seu código de redefinição de senha do Fiscal Flow",
                'body': """
                <html><body>
                <h2>Olá, {username}!</h2>
                <p>Recebemos uma solicitação para redefinir sua senha no sistema Fiscal Flow.</p>
                <p>Utilize o código abaixo na tela "Redefinir Senha" do programa:</p>
                <h1 style="text-align:center; letter-spacing: 5px; font-size: 36px;">{token}</h1>
                <p>Se você não solicitou isso, pode ignorar este e-mail com segurança.</p>
                <p>Este código é válido por 24 horas.</p>
                <br><p>Atenciosamente,<br>Equipe Fiscal Flow</p>
                </body></html>
                """
            },
            'support_reply': {
                'name': 'E-mail de Resposta do Suporte',
                'subject': 'Sua solicitação de suporte foi respondida - Chamado #{ticket_id}',
                'body': """
                <html><body>
                <h2>Olá, {username}!</h2>
                <p>O seu chamado de suporte com o assunto "<b>{subject}</b>" recebeu uma nova resposta do atendente <b>{attendant_name}</b>.</p>
                <p>Por favor, abra o Fiscal Flow para visualizar a resposta.</p>
                <br><p>Atenciosamente,<br>Equipe de Suporte</p>
                </body></html>
                """
            },
            'support_closed': {
                'name': 'E-mail de Chamado Concluído',
                'subject': 'Sua solicitação de suporte foi concluída - Chamado #{ticket_id}',
                'body': """
                <html><body>
                <h2>Olá, {username}!</h2>
                <p>O seu chamado de suporte com o assunto "<b>{subject}</b>" foi marcado como "Fechado" pelo nosso time.</p>
                <p>Se o seu problema foi resolvido, nenhuma ação é necessária. Se você ainda precisa de ajuda, sinta-se à vontade para responder no chamado ou abrir um novo.</p>
                <br><p>Atenciosamente,<br>Equipe de Suporte</p>
                </body></html>
                """
            }
        }
        
        for identifier, data in templates_to_check.items():
            template_ref = db.collection('email_templates').document(identifier)
            if not template_ref.get(timeout=QUERY_TIMEOUT).exists:
                template_ref.set({
                    'identifier': identifier, 'name': data['name'],
                    'subject': data['subject'], 'body_html': data['body']
                })

    except google_exceptions.DeadlineExceeded:
        raise ConnectionError(f"A conexão com o banco de dados excedeu o tempo limite de {QUERY_TIMEOUT} segundos.")
    except Exception as e:
        raise e

def get_smtp_config_from_firestore():
    try:
        config_ref = db.collection('app_config').document('smtp_settings')
        config_doc = config_ref.get(timeout=QUERY_TIMEOUT)
        return config_doc.to_dict() if config_doc.exists else None
    except Exception as e:
        print(f"ERRO ao buscar config de SMTP: {e}")
        return None

def verify_user(username, password):
    try:
        users_ref = db.collection('users')
        query = users_ref.where(filter=FieldFilter('username', '==', username.lower())).limit(1).get(timeout=QUERY_TIMEOUT)
        if not query: return None
        
        user_doc = query[0]
        if not user_doc.to_dict().get('password_hash'): return None
        
        user_data = user_doc.to_dict()
        stored_hash = user_data['password_hash'].encode('utf-8')
        
        if bcrypt.checkpw(password.encode('utf-8'), stored_hash):
            user_id = user_doc.id
            user_data['id'] = user_id
            
            if 'acesso_codigos_cliente' not in user_data:
                user_data['acesso_codigos_cliente'] = 'Nenhum'
            
            # --- [NOVO] Busca e anexa os setores do usuário aos dados da sessão ---
            user_data['sector_ids'] = get_user_sectors(user_id)

            return user_data
        else: 
            return None
    except google_exceptions.DeadlineExceeded:
        raise ConnectionError("A verificação de usuário excedeu o tempo limite.")
    except Exception as e:
        print(f"Erro ao verificar usuário: {e}")
        return None

def get_all_users():
    try:
        users_list = []
        users_ref = db.collection('users').where(filter=FieldFilter('level', '!=', 'Desenvolvedor')).order_by('username').stream()
        for user in users_ref:
            user_data = user.to_dict()
            users_list.append({
                'id': user.id, 
                'username': user_data.get('username'), 
                'email': user_data.get('email'), 
                'level': user_data.get('level'),
                'acesso_codigos_cliente': user_data.get('acesso_codigos_cliente', 'Nenhum')
            })
        return users_list, None
    except google_exceptions.DeadlineExceeded:
        return [], "Erro: A busca por usuários excedeu o tempo limite."
    except Exception as e:
        error_message = f"ERRO AO BUSCAR USUÁRIOS NO FIREBASE:\n\n{e}"
        print(error_message)
        return [], error_message

def add_user(username, email, level, sector_ids=None, client_code_access='Nenhum'):
    if not (username and email and level): return None, "Todos os campos são obrigatórios."
    try:
        users_ref = db.collection('users')
        query = users_ref.where(filter=FieldFilter('username', '==', username.lower())).limit(1).get(timeout=QUERY_TIMEOUT)
        if query: return None, "Erro: O nome de usuário já existe."
        
        new_user_data = {
            'username': username.lower(), 'email': email, 'level': level, 
            'password_hash': None, 'acesso_codigos_cliente': client_code_access
        }
        doc_ref = users_ref.add(new_user_data)
        
        new_user_id = doc_ref[1].id
        if sector_ids is not None:
            update_user_sectors(new_user_id, sector_ids)
        return new_user_id, "Usuário adicionado com sucesso. Gerando convite..."
    except google_exceptions.DeadlineExceeded:
        return None, "Erro: A operação excedeu o tempo limite."
    except Exception as e:
        return None, f"Erro ao adicionar usuário: {e}"

def generate_password_setup_token(user_id):
    token = ''.join([secrets.choice('0123456789') for _ in range(6)])
    hashed_token = bcrypt.hashpw(token.encode('utf-8'), bcrypt.gensalt())
    expires_at = datetime.now() + timedelta(hours=TOKEN_VALIDITY_HOURS)
    tokens_ref = db.collection('password_reset_tokens')
    tokens_ref.add({'user_id': user_id, 'token_hash': hashed_token.decode('utf-8'), 'expires_at': expires_at, 'type': 'creation'})
    return token

def set_password_with_token(username, token, new_password):
    try:
        users_ref = db.collection('users')
        query = users_ref.where(filter=FieldFilter('username', '==', username.lower())).limit(1).get(timeout=QUERY_TIMEOUT)
        if not query: return "Usuário não encontrado."
        user_doc = query[0]
        user_id = user_doc.id
        
        tokens_ref = db.collection('password_reset_tokens')
        token_query = tokens_ref.where(filter=FieldFilter('user_id', '==', user_id)).where(filter=FieldFilter('expires_at', '>', datetime.now())).where(filter=FieldFilter('type', '==', 'creation')).stream()
        
        valid_tokens = list(token_query)
        if not valid_tokens: return "Código de verificação inválido ou expirado."
        
        token_match, token_doc_to_delete = False, None
        for token_doc in valid_tokens:
            stored_hash = token_doc.to_dict()['token_hash'].encode('utf-8')
            if bcrypt.checkpw(token.encode('utf-8'), stored_hash):
                token_match, token_doc_to_delete = True, token_doc.reference; break
        
        if not token_match: return "Código de verificação inválido ou expirado."
        
        new_hashed_password = bcrypt.hashpw(new_password.encode('utf-8'), bcrypt.gensalt())
        user_doc.reference.update({'password_hash': new_hashed_password.decode('utf-8')})
        token_doc_to_delete.delete()
        
        return "Senha definida com sucesso! Você já pode fazer o login."
    except google_exceptions.DeadlineExceeded:
        return "Erro: A operação excedeu o tempo limite."
    except Exception as e:
        print(f"ERRO em set_password_with_token: {e}"); return f"Ocorreu um erro inesperado."

def update_user(user_id, username, email, level, sector_ids=None, password=None, client_code_access=None):
    try:
        user_ref = db.collection('users').document(user_id)
        update_data = {'username': username.lower(), 'email': email, 'level': level}
        if password:
            hashed_password = bcrypt.hashpw(password.encode('utf-8'), bcrypt.gensalt())
            update_data['password_hash'] = hashed_password.decode('utf-8')
        if client_code_access is not None:
            update_data['acesso_codigos_cliente'] = client_code_access
        user_ref.update(update_data)
        if sector_ids is not None:
            update_user_sectors(user_id, sector_ids)
        return "Usuário atualizado com sucesso."
    except Exception as e: return f"Erro ao atualizar usuário: {e}"

def delete_user(user_id):
    try:
        db.collection('users').document(user_id).delete()
        db.collection('user_sectors').document(user_id).delete()
        return "Usuário removido com sucesso."
    except Exception as e: return f"Erro ao remover usuário: {e}"

def change_password(user_id, old_password, new_password):
    try:
        user_ref = db.collection('users').document(user_id)
        user_doc = user_ref.get(timeout=QUERY_TIMEOUT)
        if not user_doc.exists: return "Erro: Usuário não encontrado."
        user_data = user_doc.to_dict()
        stored_hash = user_data.get('password_hash', '').encode('utf-8')
        if not bcrypt.checkpw(old_password.encode('utf-8'), stored_hash): return "Senha atual incorreta."
        new_hashed_password = bcrypt.hashpw(new_password.encode('utf-8'), bcrypt.gensalt())
        user_ref.update({'password_hash': new_hashed_password.decode('utf-8')})
        return "Senha alterada com sucesso!"
    except google_exceptions.DeadlineExceeded:
        return "Erro: A operação excedeu o tempo limite."
    except Exception as e:
        return f"Erro inesperado: {e}"

# ... (o restante do arquivo continua o mesmo) ...

def get_email_template(identifier):
    try:
        template_doc = db.collection('email_templates').document(identifier).get(timeout=QUERY_TIMEOUT)
        return template_doc.to_dict() if template_doc.exists else None
    except Exception as e:
        print(f"Erro ao buscar template '{identifier}': {e}"); return None

def send_email(recipient_email, template, format_args):
    smtp_config = get_smtp_config_from_firestore()
    if not smtp_config: return (False, "Configuração de SMTP não encontrada.")
    sender_email, password, server, port = smtp_config.get('user'), smtp_config.get('password'), smtp_config.get('server'), smtp_config.get('port')
    if not all([sender_email, password, server, port]): return (False, "Configuração de SMTP incompleta.")
    msg = EmailMessage()
    msg['Subject'] = template['subject'].format(**format_args); msg['From'] = sender_email; msg['To'] = recipient_email
    msg.set_content(template['body_html'].format(**format_args), subtype='html')
    try:
        context = ssl.create_default_context()
        with smtplib.SMTP(server, int(port), timeout=15) as s:
            s.starttls(context=context); s.login(sender_email, password); s.send_message(msg)
        return (True, f"E-mail enviado para {recipient_email}.")
    except Exception as e: return (False, str(e))

def send_creation_email(recipient_email, new_username, token):
    template = get_email_template('account_creation')
    if not template: return (False, "Erro: Template 'account_creation' não encontrado.")
    return send_email(recipient_email, template, {'username': new_username, 'token': token})
    
def get_all_email_templates():
    templates_list = []
    for template in db.collection('email_templates').order_by('name').stream():
        templates_list.append({'id': template.id, **template.to_dict()})
    return templates_list

def update_email_template(identifier, new_subject, new_body):
    try:
        db.collection('email_templates').document(identifier).update({'subject': new_subject, 'body_html': new_body})
        return "Template atualizado com sucesso."
    except Exception as e: return f"Erro ao atualizar o template: {e}"

def create_access_request(full_name, email, sector):
    if not all([full_name, email, sector]): return "Todos os campos são obrigatórios."
    try:
        db.collection('access_requests').add({'full_name': full_name, 'email': email, 'sector': sector, 'request_date': firestore.SERVER_TIMESTAMP, 'status': 'pendente'})
        return "Solicitação enviada com sucesso!"
    except Exception as e:
        print(f"ERRO ao criar solicitação: {e}"); return "Ocorreu um erro ao enviar sua solicitação."

def get_pending_requests():
    try:
        requests_list = []
        req_ref = db.collection('access_requests').where(filter=FieldFilter('status', '==', 'pendente')).order_by('request_date').stream()
        for req in req_ref:
            req_data = req.to_dict(); req_data['id'] = req.id
            if 'request_date' in req_data and isinstance(req_data.get('request_date'), datetime):
                req_data['request_date'] = req_data['request_date'].strftime("%d/%m/%Y %H:%M")
            requests_list.append(req_data)
        return requests_list, None
    except google_exceptions.DeadlineExceeded:
        return [], "Erro: A busca excedeu o tempo limite."
    except Exception as e:
        print(f"ERRO AO BUSCAR SOLICITAÇÕES: {e}"); return [], f"ERRO: {e}"

def update_request_status(request_id, new_status):
    try:
        db.collection('access_requests').document(request_id).update({'status': new_status})
        return True, "Status da solicitação atualizado."
    except Exception as e:
        print(f"ERRO ao atualizar status: {e}"); return False, "Erro ao atualizar status."
        
def add_sector(name):
    if not name: return "O nome do setor não pode estar vazio."
    try:
        sectors_ref = db.collection('sectors')
        if list(sectors_ref.where(filter=FieldFilter('name_lower', '==', name.lower())).limit(1).stream()):
            return "Erro: Este setor já existe."
        sectors_ref.add({'name': name, 'name_lower': name.lower()})
        return f"Setor '{name}' adicionado com sucesso."
    except Exception as e: return f"Ocorreu um erro ao adicionar o setor: {e}"

def get_all_sectors():
    try:
        sectors_list = []
        for sector in db.collection('sectors').order_by('name').stream():
            sectors_list.append({'id': sector.id, **sector.to_dict()})
        return sectors_list, None
    except Exception as e:
        print(f"ERRO AO BUSCAR SETORES: {e}"); return [], f"ERRO: {e}"

def delete_sector(sector_id):
    try:
        db.collection('sectors').document(sector_id).delete(); return "Setor removido com sucesso."
    except Exception as e: return f"Ocorreu um erro ao remover o setor: {e}"

def get_user_sectors(user_id):
    try:
        doc = db.collection('user_sectors').document(user_id).get(timeout=QUERY_TIMEOUT)
        return doc.to_dict().get('sector_ids', []) if doc.exists else []
    except Exception as e:
        print(f"Erro ao buscar setores do usuário {user_id}: {e}"); return []

def update_user_sectors(user_id, sector_ids):
    try:
        db.collection('user_sectors').document(user_id).set({'sector_ids': sector_ids})
    except Exception as e: print(f"Erro ao atualizar setores do usuário {user_id}: {e}")

def get_user_emails_by_sector(sector_id):
    try:
        user_ids = [doc.id for doc in db.collection('user_sectors').where(filter=FieldFilter('sector_ids', 'array_contains', sector_id)).stream()]
        if not user_ids: return [], None
        email_list = []
        for user_id in user_ids:
            user_doc = db.collection('users').document(user_id).get(timeout=QUERY_TIMEOUT)
            if user_doc.exists and user_doc.to_dict().get('email'): email_list.append(user_doc.to_dict().get('email'))
        return email_list, None
    except google_exceptions.DeadlineExceeded:
        return [], "Erro: A busca excedeu o tempo limite."
    except Exception as e:
        print(f"ERRO AO BUSCAR E-MAILS POR SETOR: {e}"); return [], f"ERRO: {e}"

def get_all_user_emails():
    try:
        email_list = []
        for user in db.collection('users').where(filter=FieldFilter('level', '!=', 'Desenvolvedor')).stream():
            if user.to_dict().get('email'): email_list.append(user.to_dict().get('email'))
        return email_list, None
    except Exception as e:
        print(f"ERRO AO BUSCAR TODOS OS E-MAILS: {e}"); return [], f"ERRO: {e}"

def send_communication_email(recipient_list, subject, body):
    smtp_config = get_smtp_config_from_firestore()
    if not smtp_config: return (False, "Configuração de SMTP não encontrada.")
    sender_email, password, server, port = smtp_config.get('user'), smtp_config.get('password'), smtp_config.get('server'), smtp_config.get('port')
    if not all([sender_email, password, server, port]): return (False, "Configuração de SMTP incompleta.")
    msg = EmailMessage(); msg['Subject'] = subject; msg['From'] = sender_email
    msg['Bcc'] = ", ".join(recipient_list)
    msg.set_content(f"<html><body>{body.replace('\n', '<br>')}</body></html>", subtype='html')
    try:
        context = ssl.create_default_context()
        with smtplib.SMTP(server, int(port), timeout=15) as s:
            s.starttls(context=context); s.login(sender_email, password); s.send_message(msg)
        return (True, f"Comunicado enviado com sucesso para {len(recipient_list)} destinatário(s).")
    except Exception as e: return (False, f"Falha ao enviar comunicado: {e}")

def request_password_reset(email):
    try:
        query = db.collection('users').where(filter=FieldFilter('email', '==', email.lower())).limit(1).get(timeout=QUERY_TIMEOUT)
        if not query: return (False, "Nenhum usuário encontrado com este e-mail.")
        user_doc = query[0]; user_id = user_doc.id; username = user_doc.to_dict().get('username', 'usuário')
        token = ''.join([secrets.choice('0123456789') for _ in range(6)])
        hashed_token = bcrypt.hashpw(token.encode('utf-8'), bcrypt.gensalt())
        expires_at = datetime.now() + timedelta(hours=TOKEN_VALIDITY_HOURS)
        db.collection('password_reset_tokens').add({'user_id': user_id, 'token_hash': hashed_token.decode('utf-8'), 'expires_at': expires_at, 'type': 'reset'})
        template = get_email_template('password_reset')
        if not template: return (False, "Erro: Template 'password_reset' não encontrado.")
        success, message = send_email(email, template, {'username': username, 'token': token})
        return (True, f"Código de redefinição enviado para {email}.") if success else (False, f"Falha ao enviar e-mail: {message}")
    except google_exceptions.DeadlineExceeded:
        return (False, "Erro: A operação excedeu o tempo limite.")
    except Exception as e:
        print(f"ERRO em request_password_reset: {e}"); return (False, "Erro inesperado no servidor.")

def reset_password_with_token(email, token, new_password):
    try:
        query = db.collection('users').where(filter=FieldFilter('email', '==', email.lower())).limit(1).get(timeout=QUERY_TIMEOUT)
        if not query: return "E-mail não encontrado."
        user_doc = query[0]; user_id = user_doc.id
        tokens_ref = db.collection('password_reset_tokens')
        token_query = tokens_ref.where(filter=FieldFilter('user_id', '==', user_id)).where(filter=FieldFilter('expires_at', '>', datetime.now())).where(filter=FieldFilter('type', '==', 'reset')).stream()
        if not list(token_query): return "Código de redefinição inválido ou expirado."
        token_match, token_doc_to_delete = False, None
        for token_doc in list(token_query):
            stored_hash = token_doc.to_dict()['token_hash'].encode('utf-8')
            if bcrypt.checkpw(token.encode('utf-8'), stored_hash):
                token_match, token_doc_to_delete = True, token_doc.reference; break
        if not token_match: return "Código de redefinição inválido."
        new_hashed_password = bcrypt.hashpw(new_password.encode('utf-8'), bcrypt.gensalt())
        user_doc.reference.update({'password_hash': new_hashed_password.decode('utf-8')})
        token_doc_to_delete.delete()
        return "Senha redefinida com sucesso!"
    except google_exceptions.DeadlineExceeded:
        return "Erro: A operação excedeu o tempo limite."
    except Exception as e:
        print(f"ERRO em reset_password_with_token: {e}"); return "Erro inesperado ao redefinir a senha."

def send_support_reply_email(recipient_email, ticket_id, subject, username, attendant_name):
    template = get_email_template('support_reply')
    if not template: return (False, "Erro: Template 'support_reply' não encontrado.")
    return send_email(recipient_email, template, {'ticket_id': ticket_id, 'subject': subject, 'username': username, 'attendant_name': attendant_name})

def send_support_closed_email(recipient_email, ticket_id, subject, username):
    template = get_email_template('support_closed')
    if not template: return (False, "Erro: Template 'support_closed' não encontrado.")
    return send_email(recipient_email, template, {'ticket_id': ticket_id, 'subject': subject, 'username': username})