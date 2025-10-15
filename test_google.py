import requests

print("Tentando conectar ao servidor de API do Google...")
try:
    # Este é o URL que a biblioteca do Google tenta acessar para "aprender" sobre a API do Drive
    url = "https://www.googleapis.com/discovery/v1/apis/drive/v3/rest"

    print(f"Acessando: {url}")

    response = requests.get(url, timeout=20)
    response.raise_for_status()  # Lança um erro se a resposta for um erro HTTP (como 404, 500, etc)

    print("\n=====================")
    print(" SUCESSO! ")
    print("=====================")
    print(f"Conexão estabelecida com sucesso.")
    print(f"Status da Resposta do Servidor: {response.status_code}")

except requests.exceptions.RequestException as e:
    print("\n=====================")
    print(" FALHA! ")
    print("=====================")
    print(f"Não foi possível conectar ao servidor do Google.")
    print(f"O erro foi: {e}")

print("\nTeste concluído.")