# =============================================================================
# --- ARQUIVO: test_logic.py (COM TESTES SEPARADOS E CORRIGIDOS) ---
# =============================================================================

import os
import sys
import auth_logic
import core_logic

# Dicionário para guardar os resultados
test_results = {}

def run_all_tests():
    """
    Executa todos os testes em sequência e retorna os resultados.
    """
    print("--- INICIANDO SUÍTE DE AUTO-TESTE ---")
    test_results.clear()

    # --- Roteiro de Testes ---
    test_firebase_connection()
    
    # Executa os dois testes de extração e guarda os resultados combinados
    parsed_data1 = test_parsing_narwal_xml()
    parsed_data2 = test_parsing_dimnfe_xml()
    
    combined_data = []
    if parsed_data1: combined_data.extend(parsed_data1)
    if parsed_data2: combined_data.extend(parsed_data2)
        
    if combined_data:
        test_dashboard_logic(combined_data)
        
    test_user_auth()

    print("--- SUÍTE DE AUTO-TESTE CONCLUÍDA ---")
    return test_results

def resource_path(relative_path):
    """Obtém o caminho absoluto para um recurso."""
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

# --- Funções de Teste ---

def test_firebase_connection():
    """Testa se a inicialização do Firebase ocorre sem erros."""
    test_name = "Conexão com Banco de Dados (Firebase)"
    try:
        auth_logic.initialize_firebase()
        test_results[test_name] = ("OK", "Conexão estabelecida com sucesso.")
    except Exception as e:
        test_results[test_name] = ("FALHOU", f"Não foi possível conectar. Erro: {e}")

def test_parsing_narwal_xml():
    """Testa a extração de dados do XML da Narwal Sistemas."""
    test_name = "Extração (XML Padrão 1 - Narwal)"
    try:
        test_file_path = resource_path(os.path.join("test_assets", "narwal_test_note.xml"))
        if not os.path.exists(test_file_path):
            raise FileNotFoundError("Arquivo 'narwal_test_note.xml' não encontrado.")
        
        data = core_logic.extrair_dados_nf(test_file_path)
        if not data: raise ValueError("A extração retornou um objeto vazio (None).")
        
        # Validações específicas para este arquivo
        assert data.get('numero_nf') == '1315', f"Número da NF esperado '1315', mas veio '{data.get('numero_nf')}'"
        assert data.get('valor_total_nf') == 43613.96, f"Valor Total esperado 43613.96, mas veio '{data.get('valor_total_nf')}'"
        assert data.get('sistema_emissor') == 'Narwal Sistemas', f"Sistema Emissor esperado 'Narwal Sistemas', mas veio '{data.get('sistema_emissor')}'"

        test_results[test_name] = ("OK", "Dados validados com sucesso.")
        return [data]
    except Exception as e:
        test_results[test_name] = ("FALHOU", f"Detalhe: {e}")
        return None

def test_parsing_dimnfe_xml():
    """Testa a extração de dados do XML da DIMNFE."""
    test_name = "Extração (XML Padrão 2 - DIMNFE)"
    try:
        test_file_path = resource_path(os.path.join("test_assets", "dimnfe_test_note.xml"))
        if not os.path.exists(test_file_path):
            raise FileNotFoundError("Arquivo 'dimnfe_test_note.xml' não encontrado.")
        
        data = core_logic.extrair_dados_nf(test_file_path)
        if not data: raise ValueError("A extração retornou um objeto vazio (None).")
        
        # Validações específicas para este arquivo
        assert data.get('numero_nf') == '1356', f"Número da NF esperado '1356', mas veio '{data.get('numero_nf')}'"
        assert data.get('valor_total_nf') == 101581.09, f"Valor Total esperado 101581.09, mas veio '{data.get('valor_total_nf')}'"
        assert data.get('sistema_emissor') == 'DIMNFE-4.00', f"Sistema Emissor esperado 'DIMNFE-4.00', mas veio '{data.get('sistema_emissor')}'"
        
        test_results[test_name] = ("OK", "Dados validados com sucesso.")
        return [data]
    except Exception as e:
        test_results[test_name] = ("FALHOU", f"Detalhe: {e}")
        return None

def test_dashboard_logic(parsed_data_list):
    """Testa a lógica de cálculo do dashboard com os dados combinados."""
    test_name = "Lógica de Cálculos do Dashboard"
    try:
        dashboard_data = core_logic.calcular_dados_dashboard(parsed_data_list)
        if not dashboard_data: raise ValueError("Função de cálculo retornou um objeto vazio.")
        
        resumo = dashboard_data.get('resumo_geral', {})
        assert resumo.get('total_notas') == 2, f"Contagem de notas esperada 2, mas o resultado foi {resumo.get('total_notas')}"
        
        test_results[test_name] = ("OK", "Cálculos processados com sucesso.")
    except Exception as e:
        test_results[test_name] = ("FALHOU", f"Erro nos cálculos. Detalhe: {e}")

def test_user_auth():
    """Testa a lógica de verificação de senha (login)."""
    test_name = "Autenticação de Usuário"
    try:
        user_dev = auth_logic.verify_user('dev', 'dev')
        if not user_dev:
            raise ValueError("Falha ao autenticar o usuário 'dev' com a senha correta.")
        user_wrong_pass = auth_logic.verify_user('dev', 'senha_errada')
        if user_wrong_pass is not None:
            raise ValueError("Sistema permitiu login com senha incorreta.")
        test_results[test_name] = ("OK", "Verificação de senhas funcionou como esperado.")
    except Exception as e:
        test_results[test_name] = ("FALHOU", f"Erro na autenticação. Detalhe: {e}")