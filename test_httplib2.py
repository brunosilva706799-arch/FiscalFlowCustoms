import httplib2
import sys

print("--- Iniciando teste de conexão com a biblioteca 'httplib2' ---")

try:
    # Criamos uma instância do cliente Http, com um timeout de 15 segundos
    http = httplib2.Http(timeout=15)

    url = "https://www.googleapis.com/discovery/v1/apis/drive/v3/rest"
    print(f"Tentando acessar: {url}\n")

    # Fazendo a requisição exatamente como a biblioteca do Google faria
    response_headers, content = http.request(url, 'GET')

    print("\n=====================")
    print(" SUCESSO! (com httplib2)")
    print("=====================")
    print(f"Status da Resposta: {response_headers.status}")
    print("A conexão direta via httplib2 FUNCIONOU.")

except Exception as e:
    print("\n=====================")
    print(" FALHA! (com httplib2)")
    print("=====================")
    print(f"Não foi possível conectar usando httplib2.")
    print(f"O erro foi do tipo: {type(e).__name__}")
    print(f"Mensagem de Erro Completa: {e}")

    # Adiciona um diagnóstico provável para erros comuns de proxy
    if "SSL" in str(e).upper() or "CERTIFICATE_VERIFY_FAILED" in str(e).upper():
        print("\nDIAGNÓSTICO PROVÁVEL: O erro contém 'SSL' ou 'CERTIFICATE'.")
        print("Isso indica quase com certeza que um firewall, antivírus ou proxy corporativo")
        print("está interceptando a conexão SSL que esta biblioteca específica tenta fazer.")

print("\nTeste concluído.")