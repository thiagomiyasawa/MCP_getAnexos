import os
import base64
import requests
from mcp.server import MCPServer
from dotenv import load_dotenv
from generate_token import get_access_token

# 1. Criar a instância do servidor
mcp = MCPServer(name="Meu Servico MCP")


def get_token():
    SCOPES = ["Mail.Read"]

    load_dotenv()

    app_id = os.getenv("APPLICATION_ID")
    tenant_id = os.getenv("TENANT_ID")  # Use 'consumers' para seu @outlook.com pessoal

    if not all([app_id, tenant_id]):
        print("Erro: APPLICATION_ID ou TENANT_ID faltando no arquivo .env")
        return None

    # Executa o fluxo interativo inteligente
    token_response = get_access_token(str(app_id), str(tenant_id), SCOPES)

    if token_response and "access_token" in token_response:
        return token_response
    print("\nNão foi possível processar a requisição sem um token válido.")
    return None


# 2. Definir uma ferramenta (funcionalidade executável)
@mcp.tool()
def get_attachments(id: str, path: str):
    """Baixa todos os anexos de um e-mail do Outlook/Microsoft 365 para uma pasta local.

    Busca a mensagem pelo seu ID via Microsoft Graph API e salva cada anexo
    encontrado (decodificado de base64) dentro do diretório indicado.

    Args:
        id: ID do e-mail no Microsoft Graph (campo "id" da mensagem), usado para
            montar o endpoint /me/messages/{id}/attachments.
        path: Caminho da pasta local onde os anexos serão salvos. A pasta deve
            existir previamente; cada anexo é gravado como "{path}/{nome_do_anexo}".

    Returns:
        "success" se os anexos foram baixados com sucesso, ou uma string
        "error:{response}" com os detalhes da falha na chamada à API.
    """
    access_token = get_token()["access_token"]
    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json",
    }
    url = f"https://graph.microsoft.com/v1.0/me/messages/{id}/attachments"
    response = requests.get(url, headers=headers, timeout=30)
    if response.status_code == 200:
        anexos = response.json().get("value", [])
        for anexo in anexos:
            anexoPath = f"{path}/{anexo['name']}"
            with open(anexoPath, "wb") as f:
                f.write(base64.b64decode(anexo["contentBytes"]))
        return "success"
    else:
        return f"error:{response}"


# 3. Definir um recurso (dados estáticos ou dinâmicos)
@mcp.resource("info://tools")
async def listar_tools() -> str:
    """Lista todas as tools disponíveis neste servidor MCP, com nome e descrição."""
    tools = await mcp.list_tools()
    if not tools:
        return "Nenhuma tool disponível neste servidor."

    linhas = [f"- {tool.name}: {tool.description}" for tool in tools]
    return "\n".join(linhas)


if __name__ == "__main__":
    mcp.run()
