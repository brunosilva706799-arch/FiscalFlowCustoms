# =============================================================================
# --- ARQUIVO: client_logic.py ---
# (Atualizado com função de resumo de dados)
# =============================================================================

import unicodedata
from firebase_admin import firestore
from google.api_core import exceptions as google_exceptions
import pandas as pd

# --- Constantes do Módulo ---
COLLECTION_NAME = 'client_codes'
PREFIX = "CLI"
VOWELS = "AEIOU"

# =============================================================================
# --- LÓGICA PRINCIPAL DE GERAÇÃO DE CÓDIGO ---
# =============================================================================

def _normalize_text(text):
    """Remove acentos, c, e converte para maiúsculas para comparações."""
    if not text:
        return ""
    nfkd_form = unicodedata.normalize('NFD', text)
    return "".join([c for c in nfkd_form if not unicodedata.combining(c)]).upper()

def generate_acronym(client_name):
    """Gera um acrônimo de 3 letras para um nome de cliente com base na regra definida."""
    if not client_name or not isinstance(client_name, str) or len(client_name.strip()) < 3:
        return ""
    name = _normalize_text(client_name.strip())
    name_no_spaces = "".join(name.split())
    if len(name_no_spaces) == 0:
        return ""
    first_letter = name_no_spaces[0]
    vowel_indices = [i for i, char in enumerate(name_no_spaces) if char in VOWELS]
    second_letter, third_letter = '', ''
    if len(vowel_indices) >= 1:
        start_search_pos = vowel_indices[0] + 1
        for char in name_no_spaces[start_search_pos:]:
            if char not in VOWELS and char.isalpha():
                second_letter = char; break
    if len(vowel_indices) >= 2:
        start_search_pos = vowel_indices[1] + 1
        for char in name_no_spaces[start_search_pos:]:
            if char not in VOWELS and char.isalpha():
                third_letter = char; break
    if not (first_letter and second_letter and third_letter):
        consonants = [c for c in name_no_spaces if c not in VOWELS and c.isalpha()]
        letters = [first_letter]
        for con in consonants:
            if len(letters) < 3 and con not in letters: letters.append(con)
        for char in name_no_spaces:
            if len(letters) < 3 and char not in letters: letters.append(char)
        return "".join(letters).ljust(3, 'X')
    return f"{first_letter}{second_letter}{third_letter}"

# =============================================================================
# --- FUNÇÕES DE INTERAÇÃO COM O BANCO DE DADOS (FIRESTORE) ---
# =============================================================================

def get_next_sequence_for_acronym(db, acronym):
    try:
        docs = db.collection(COLLECTION_NAME).where('acronym', '==', acronym).stream()
        max_seq = 0
        for doc in docs:
            if doc.to_dict().get('sequence', 0) > max_seq: max_seq = doc.to_dict().get('sequence')
        return max_seq + 1
    except Exception as e:
        print(f"ERRO ao buscar sequencial: {e}"); return 1

def check_if_name_exists(db, normalized_name):
    try:
        return len(list(db.collection(COLLECTION_NAME).where('normalized_name', '==', normalized_name).limit(1).stream())) > 0
    except Exception as e:
        print(f"ERRO ao verificar nome: {e}"); return False

def create_client_code(client_name):
    db = firestore.client()
    if not client_name or not client_name.strip():
        return None, "O nome do cliente não pode estar vazio."
    normalized_name = _normalize_text(client_name)
    try:
        if check_if_name_exists(db, normalized_name):
            return None, "Já existe um cliente com este nome cadastrado."
        acronym = generate_acronym(client_name)
        if not acronym: return None, "Não foi possível gerar um acrônimo."
        next_seq = get_next_sequence_for_acronym(db, acronym)
        full_code = f"{PREFIX}.{acronym}.{next_seq:03d}"
        new_client_data = {'name': client_name.strip(), 'normalized_name': normalized_name, 'code': full_code, 'acronym': acronym, 'sequence': next_seq, 'created_at': firestore.SERVER_TIMESTAMP}
        db.collection(COLLECTION_NAME).add(new_client_data)
        return new_client_data, "Código de cliente criado com sucesso."
    except Exception as e:
        print(f"ERRO ao criar código: {e}"); return None, f"Erro inesperado: {e}"

def get_all_clients():
    db = firestore.client()
    try:
        clients_list = []
        for doc in db.collection(COLLECTION_NAME).order_by('name').stream():
            clients_list.append({'id': doc.id, **doc.to_dict()})
        return clients_list, None
    except google_exceptions.FailedPrecondition as e:
        print(f"\n--- AÇÃO NECESSÁRIA NO FIREBASE: {e} ---\n")
        return [], "Erro de BD. Índice necessário. Verifique o terminal."
    except Exception as e:
        print(f"ERRO ao buscar clientes: {e}"); return [], f"Erro inesperado ao buscar clientes: {e}"

def update_client(client_id, new_client_name):
    db = firestore.client()
    if not new_client_name or not new_client_name.strip():
        return None, "O novo nome não pode estar vazio."
    normalized_name = _normalize_text(new_client_name)
    try:
        query = db.collection(COLLECTION_NAME).where('normalized_name', '==', normalized_name).stream()
        for doc in query:
            if doc.id != client_id: return None, "Já existe outro cliente com este nome."
        db.collection(COLLECTION_NAME).document(client_id).update({'name': new_client_name.strip(), 'normalized_name': normalized_name})
        return True, "Nome do cliente atualizado com sucesso."
    except Exception as e:
        print(f"ERRO ao atualizar cliente: {e}"); return None, f"Erro inesperado ao atualizar: {e}"

def delete_clients_batch(client_ids):
    db = firestore.client()
    try:
        batch = db.batch()
        for client_id in client_ids: batch.delete(db.collection(COLLECTION_NAME).document(client_id))
        batch.commit()
        return True, f"{len(client_ids)} cliente(s) removido(s) com sucesso."
    except Exception as e:
        print(f"ERRO ao deletar em lote: {e}"); return False, f"Erro inesperado ao remover: {e}"

def import_clients_from_xlsx(filepath):
    """Lê um arquivo .xlsx e importa os clientes para o Firestore."""
    try:
        df = pd.read_excel(filepath)
        df.columns = [col.strip() for col in df.columns]
        
        if 'Cliente' not in df.columns or 'Nova Nomenclatura' not in df.columns:
            return 0, 0, "O arquivo Excel deve conter as colunas 'Cliente' e 'Nova Nomenclatura'."

        db = firestore.client()
        batch = db.batch()
        
        existing_clients, _ = get_all_clients()
        existing_normalized_names = {c.get('normalized_name') for c in existing_clients}
        
        imported_count, skipped_count = 0, 0

        for index, row in df.iterrows():
            name, code = row.get('Cliente'), row.get('Nova Nomenclatura')

            if pd.isna(name) or pd.isna(code) or not str(name).strip() or not str(code).strip():
                continue

            name, code = str(name).strip(), str(code).strip()
            normalized_name = _normalize_text(name)
            
            if normalized_name in existing_normalized_names:
                skipped_count += 1
                continue

            try:
                parts = code.split('.')
                if len(parts) < 3: raise ValueError(f"Código '{code}' < 3 partes.")
                sequence = int(parts[-1])
                acronym = ".".join(parts[1:-1])
                if not acronym: raise ValueError(f"Código '{code}' tem acrônimo vazio.")
            except (IndexError, ValueError) as e:
                print(f"[CÓDIGO INVÁLIDO] Cliente '{name}' ignorado. Código '{code}'. Erro: {e}")
                skipped_count += 1
                continue

            client_data = {'name': name, 'normalized_name': normalized_name, 'code': code, 'acronym': acronym, 'sequence': sequence, 'created_at': firestore.SERVER_TIMESTAMP}
            doc_ref = db.collection(COLLECTION_NAME).document()
            batch.set(doc_ref, client_data)
            imported_count += 1
            existing_normalized_names.add(normalized_name)
        
        if imported_count > 0:
            batch.commit()
            
        return imported_count, skipped_count, None

    except FileNotFoundError:
        return 0, 0, "Arquivo não encontrado."
    except Exception as e:
        print(f"ERRO ao importar planilha: {e}")
        return 0, 0, f"Ocorreu um erro inesperado ao ler o arquivo: {e}"

# --- [NOVA FUNÇÃO DE RESUMO] ---
def get_clients_summary():
    """Busca dados de resumo da coleção de clientes de forma otimizada."""
    db = firestore.client()
    summary = {'total_clients': 0, 'last_added_date': None}
    try:
        # Pega a contagem total de documentos (muito eficiente)
        count_query = db.collection(COLLECTION_NAME).count()
        count_result = count_query.get()
        if count_result:
            summary['total_clients'] = count_result[0][0].value

        # Pega o último cliente adicionado ordenando pela data de criação
        if summary['total_clients'] > 0:
            last_added_query = db.collection(COLLECTION_NAME).order_by(
                'created_at', direction=firestore.Query.DESCENDING
            ).limit(1).stream()
            
            last_client = next(last_added_query, None)
            if last_client:
                summary['last_added_date'] = last_client.to_dict().get('created_at')
        
        return summary, None
    except Exception as e:
        print(f"ERRO ao buscar resumo de clientes: {e}")
        return summary, f"Ocorreu um erro ao buscar os dados de resumo: {e}"