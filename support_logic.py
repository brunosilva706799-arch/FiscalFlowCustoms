# =============================================================================
# --- ARQUIVO: support_logic.py ---
# (Atualizado com todas as novas lógicas de automação e gerenciamento)
# =============================================================================

from firebase_admin import firestore
from google.api_core import exceptions as google_exceptions
from datetime import datetime
import auth_logic

QUERY_TIMEOUT = 15

def create_ticket(user_id, username, subject, first_message, ticket_type):
    db = firestore.client()
    try:
        if not all([user_id, username, subject, first_message, ticket_type]):
            return None, "Assunto e mensagem são obrigatórios."
        tickets_ref = db.collection('support_tickets')
        new_ticket_data = {
            'user_id': user_id,
            'username': username,
            'subject': subject,
            'ticket_type': ticket_type,
            'status': 'Aberto',
            'created_at': firestore.SERVER_TIMESTAMP,
            'last_updated_at': firestore.SERVER_TIMESTAMP,
            'category_id': None,
            'category_name': 'Não definida'
        }
        if ticket_type == 'it':
            new_ticket_data['flag_color'] = 'Gray'
        new_ticket_ref = tickets_ref.document()
        new_ticket_ref.set(new_ticket_data)
        messages_ref = new_ticket_ref.collection('messages')
        messages_ref.add({
            'sender_id': user_id,
            'sender_name': username,
            'text': first_message,
            'timestamp': firestore.SERVER_TIMESTAMP
        })
        return new_ticket_ref.id, "Chamado aberto com sucesso."
    except google_exceptions.DeadlineExceeded:
        return None, f"Erro: A operação excedeu o tempo limite de {QUERY_TIMEOUT} segundos."
    except Exception as e:
        print(f"ERRO ao criar ticket: {e}")
        return None, f"Ocorreu um erro inesperado ao abrir o chamado: {e}"

def get_tickets_for_user(user_id, ticket_type):
    db = firestore.client()
    try:
        tickets_list = []
        query = db.collection('support_tickets').where(
            'user_id', '==', user_id
        ).where(
            'ticket_type', '==', ticket_type
        ).order_by(
            'last_updated_at', direction=firestore.Query.DESCENDING
        ).stream()
        for ticket in query:
            ticket_data = ticket.to_dict()
            tickets_list.append({
                'id': ticket.id, 'subject': ticket_data.get('subject'),
                'status': ticket_data.get('status'), 'username': ticket_data.get('username'),
                'last_updated_at': ticket_data.get('last_updated_at', 'N/A'),
                'flag_color': ticket_data.get('flag_color', 'Gray'),
                'category_id': ticket_data.get('category_id'),
                'category_name': ticket_data.get('category_name', 'Não definida')
            })
        return tickets_list, None
    except google_exceptions.DeadlineExceeded:
        return [], f"Erro: A busca por chamados excedeu o tempo limite de {QUERY_TIMEOUT} segundos."
    except Exception as e:
        print(f"ERRO ao buscar tickets do usuário: {e}")
        return [], f"Ocorreu um erro inesperado ao buscar seus chamados: {e}"

def get_all_tickets(ticket_type):
    db = firestore.client()
    try:
        tickets_list = []
        query = db.collection('support_tickets').where(
            'ticket_type', '==', ticket_type
        ).order_by(
            'last_updated_at', direction=firestore.Query.DESCENDING
        ).stream()
        for ticket in query:
            ticket_data = ticket.to_dict()
            tickets_list.append({
                'id': ticket.id, 'subject': ticket_data.get('subject'),
                'status': ticket_data.get('status'), 'username': ticket_data.get('username'),
                'last_updated_at': ticket_data.get('last_updated_at', 'N/A'),
                'flag_color': ticket_data.get('flag_color', 'Gray'),
                'category_id': ticket_data.get('category_id'),
                'category_name': ticket_data.get('category_name', 'Não definida')
            })
        return tickets_list, None
    except google_exceptions.DeadlineExceeded:
        return [], f"Erro: A busca por chamados excedeu o tempo limite de {QUERY_TIMEOUT} segundos."
    except Exception as e:
        print(f"ERRO ao buscar todos os tickets: {e}")
        return [], f"Ocorreu um erro inesperado ao buscar os chamados: {e}"

def update_ticket_details(ticket_id, new_status, new_color, category_id, category_name):
    db = firestore.client()
    try:
        if not ticket_id: return False, "ID do ticket é necessário."
        ticket_ref = db.collection('support_tickets').document(ticket_id)
        update_data = {'last_updated_at': firestore.SERVER_TIMESTAMP}
        ticket_doc_before = ticket_ref.get().to_dict()
        status_changed_to_closed = False
        if new_status and new_status != ticket_doc_before.get('status'):
            update_data['status'] = new_status
            if new_status == "Fechado": status_changed_to_closed = True
        if new_color: update_data['flag_color'] = new_color
        update_data['category_id'] = category_id
        update_data['category_name'] = category_name
        if len(update_data) > 1: ticket_ref.update(update_data)
        if status_changed_to_closed:
            user_doc = db.collection('users').document(ticket_doc_before['user_id']).get()
            if user_doc.exists:
                user_email = user_doc.to_dict().get('email')
                if user_email:
                    auth_logic.send_support_closed_email(user_email, ticket_id, ticket_doc_before['subject'], ticket_doc_before['username'])
        return True, "Chamado atualizado com sucesso."
    except Exception as e:
        print(f"ERRO ao atualizar detalhes do ticket {ticket_id}: {e}")
        return False, f"Ocorreu um erro ao atualizar o chamado: {e}"

def delete_ticket(ticket_id):
    db = firestore.client()
    try:
        ticket_ref = db.collection('support_tickets').document(ticket_id)
        messages_ref = ticket_ref.collection('messages')
        docs = messages_ref.limit(500).stream()
        for doc in docs: doc.reference.delete()
        ticket_ref.delete()
        return True, "Chamado e todas as suas mensagens foram removidos."
    except Exception as e:
        print(f"ERRO ao deletar o ticket {ticket_id}: {e}")
        return False, f"Ocorreu um erro inesperado ao remover o chamado: {e}"

def get_messages_for_ticket(ticket_id):
    db = firestore.client()
    try:
        messages_list = []
        messages_ref = db.collection('support_tickets').document(ticket_id).collection('messages').order_by('timestamp', direction=firestore.Query.ASCENDING).stream()
        for msg in messages_ref:
            msg_data = msg.to_dict()
            # --- [MODIFICADO] Adiciona os dados do anexo ao carregar as mensagens ---
            messages_list.append({
                'id': msg.id, 
                'sender_name': msg_data.get('sender_name'), 
                'text': msg_data.get('text'), 
                'timestamp': msg_data.get('timestamp'),
                'attachment_url': msg_data.get('attachment_url'),
                'attachment_filename': msg_data.get('attachment_filename')
            })
        return messages_list, None
    except Exception as e:
        print(f"ERRO ao buscar mensagens do ticket {ticket_id}: {e}")
        return [], f"Ocorreu um erro inesperado ao carregar a conversa: {e}"

# --- PONTO DA CORREÇÃO ---
# A assinatura da função foi atualizada para aceitar os argumentos do anexo.
def add_message_to_ticket(ticket_id, sender_id, sender_name, sender_level, text, attachment_url=None, attachment_filename=None):
    db = firestore.client()
    try:
        # A mensagem de texto pode estar vazia se houver um anexo
        if not text and not attachment_url:
            return None, "A mensagem não pode estar vazia."

        ticket_ref = db.collection('support_tickets').document(ticket_id)
        messages_ref = ticket_ref.collection('messages')
        ticket_data_before = ticket_ref.get().to_dict()

        @firestore.transactional
        def update_in_transaction(transaction, ticket_ref, messages_ref):
            ticket_snapshot = ticket_ref.get(transaction=transaction)
            ticket_data = ticket_snapshot.to_dict()
            new_msg_ref = messages_ref.document()

            # [MODIFICADO] Constrói o dicionário da mensagem dinamicamente
            new_message_data = {
                'sender_id': sender_id,
                'sender_name': sender_name,
                'text': text,
                'timestamp': firestore.SERVER_TIMESTAMP
            }
            # Adiciona os dados do anexo apenas se eles existirem
            if attachment_url and attachment_filename:
                new_message_data['attachment_url'] = attachment_url
                new_message_data['attachment_filename'] = attachment_filename
            
            transaction.set(new_msg_ref, new_message_data)
            
            update_data = {'last_updated_at': firestore.SERVER_TIMESTAMP}
            is_attendant = sender_level in ['Admin', 'Desenvolvedor', 'T.I.']

            if is_attendant and ticket_data.get('status') == 'Aberto':
                update_data['status'] = 'Em Andamento'
            elif not is_attendant and ticket_data.get('status') == 'Aguardando Resposta':
                update_data['status'] = 'Em Andamento'
            
            transaction.update(ticket_ref, update_data)

        transaction = db.transaction()
        update_in_transaction(transaction, ticket_ref, messages_ref)
        
        is_attendant = sender_level in ['Admin', 'Desenvolvedor', 'T.I.']
        if is_attendant:
            user_doc = db.collection('users').document(ticket_data_before['user_id']).get()
            if user_doc.exists:
                user_email = user_doc.to_dict().get('email')
                if user_email:
                    auth_logic.send_support_reply_email(user_email, ticket_id, ticket_data_before['subject'], ticket_data_before['username'], sender_name)

        return True, "Mensagem enviada com sucesso."
    except Exception as e:
        print(f"ERRO ao adicionar mensagem ao ticket {ticket_id}: {e}")
        return None, f"Ocorreu um erro inesperado ao enviar a mensagem: {e}"

# --- FUNÇÕES PARA CATEGORIAS ---

def get_categories_for_type(ticket_type):
    db = firestore.client()
    try:
        query = db.collection('support_categories').where('ticket_type', '==', ticket_type).order_by('name')
        categories = [{'id': doc.id, **doc.to_dict()} for doc in query.stream()]
        return categories, None
    # --- [NOVO] Tratamento de erro específico para índice ---
    except google_exceptions.FailedPrecondition as e:
        print("\n--- AÇÃO NECESSÁRIA NO FIREBASE ---")
        print(e)
        print("--- CLIQUE NO LINK ACIMA PARA CRIAR O ÍNDICE E TENTE NOVAMENTE ---\n")
        return [], "Erro de configuração no banco de dados. Um índice Firestore é necessário. Verifique o terminal (CMD) para o link de criação."
    except Exception as e:
        return [], f"Erro ao buscar categorias: {e}"

def add_category(name, ticket_type):
    db = firestore.client()
    try:
        name_normalized = name.strip().lower()
        query = db.collection('support_categories').where('name_normalized', '==', name_normalized).where('ticket_type', '==', ticket_type).limit(1)
        # Este stream() vai acionar o erro se o índice não existir
        list(query.stream())
        # Se o código continuar, o índice existe. Agora verificamos se há itens.
        if len(list(query.stream())) > 0:
            return None, "Esta categoria já existe."
        new_cat_ref = db.collection('support_categories').document()
        new_cat_ref.set({'name': name.strip(),'name_normalized': name_normalized, 'ticket_type': ticket_type})
        return new_cat_ref.id, "Categoria adicionada com sucesso."
    # --- [NOVO] Tratamento de erro específico para índice ---
    except google_exceptions.FailedPrecondition as e:
        print("\n--- AÇÃO NECESSÁRIA NO FIREBASE ---")
        print(e)
        print("--- CLIQUE NO LINK ACIMA PARA CRIAR O ÍNDICE E TENTE NOVAMENTE ---\n")
        return None, "Erro de configuração no banco de dados. Um índice Firestore é necessário. Verifique o terminal (CMD) para o link de criação."
    except Exception as e:
        return None, f"Erro ao adicionar categoria: {e}"