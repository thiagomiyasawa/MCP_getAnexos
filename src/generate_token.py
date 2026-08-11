import os
from urllib.parse import parse_qs, urlparse

import msal
from dotenv import load_dotenv


def get_access_token(app_id: str, tenant_id: str, scopes: list):
    """Obtém o token de forma totalmente automática usando o navegador."""
    cache_path = "accessToken.json"
    token_cache = msal.SerializableTokenCache()

    # Carrega o cache se ele existir
    if os.path.exists(cache_path):
        with open(cache_path, "r") as f:
            token_cache.deserialize(f.read())

    authority = f"https://login.microsoftonline.com/{tenant_id}"

    # IMPORTANTE: Mudamos para PublicClientApplication (Não usa Client Secret)
    app = msal.PublicClientApplication(
        client_id=app_id, authority=authority, token_cache=token_cache
    )

    # Tenta obter o token silenciosamente pelo cache primeiro
    accounts = app.get_accounts()
    if accounts:
        print("Obtendo token do cache silenciosamente...")
        token_response = app.acquire_token_silent(scopes, account=accounts[0])
        if token_response:
            return token_response

    # Se não tiver cache ou expirou, abre o navegador e captura o código SOZINHO
    print("Nenhum token válido encontrado. Abrindo o navegador para login...")
    token_response = app.acquire_token_interactive(scopes=scopes)

    if "error" not in token_response:
        # Salva o cache atualizado para não pedir login toda vez
        with open(cache_path, "w") as f:
            f.write(token_cache.serialize())
        return token_response
    else:
        print(f"\n[Erro na Autenticação]: {token_response.get('error')}")
        print(f"Descrição: {token_response.get('error_description')}")
        return None


# --- Execução Principal do Script ---
if __name__ == "__main__":
    # Escopo padrão para o Microsoft Graph
    SCOPES = ["https://graph.microsoft.com/.default"]

    # Carrega as variáveis de ambiente do arquivo .env
    load_dotenv()

    app_id = os.getenv("APPLICATION_ID")
    client_secret = os.getenv("CLIENT_SECRET")
    tenant_id = os.getenv("TENANT_ID")

    # Validação rápida dos dados do .env
    if not all([app_id, client_secret, tenant_id]):
        print(
            "Erro: Certifique-se de que APPLICATION_ID, CLIENT_SECRET e TENANT_ID estão definidos no arquivo .env"
        )
        exit(1)

    token, username = generate_access_token(
        str(app_id),
        str(client_secret),
        str(tenant_id),
        SCOPES,
    )

    if token:
        print("\n Autenticação realizada com sucesso!")
        print(f"Usuário: {username}")
        print(f"Token de Acesso (Truncado): {token.get('access_token')[:30]}...")
    else:
        print("\n Não foi possível gerar o token.")
