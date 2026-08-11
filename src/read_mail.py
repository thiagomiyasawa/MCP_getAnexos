import os
from datetime import datetime, timedelta

import requests
from dotenv import load_dotenv
from generate_token import get_access_token
from markdownify import markdownify as md
import base64


def read_mail(access_token: str, dataInicial: datetime, dataFinal: datetime):
    headers = {
        # Garanta que o 'access_token' aqui seja estritamente a String do Token
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json",
    }
    data_inicial_str = dataInicial.strftime("%Y-%m-%d")
    data_final_str = dataFinal.strftime("%Y-%m-%d")

    inicio = f"{data_inicial_str}T00:00:00Z"
    fim = f"{data_final_str}T23:59:59Z"
    print(inicio, fim)
    # CORREÇÃO: Mudamos de '/users/{user_email}/...' para '/me/...'
    url = (
        "https://graph.microsoft.com/v1.0/me/mailFolders/inbox/messages?$top=1"
        # f"?$filter=receivedDateTime ge {inicio} and receivedDateTime le {fim}"
    )

    response = requests.get(url, headers=headers)

    if response.status_code == 200:
        emails = response.json().get("value", [])
        if not emails:
            print("Nenhum e-mail encontrado para a data de ontem.")
            return None
        output = []
        for email in emails:
            anexo = email.get("hasAttachments")
            subject = email.get("subject", "(Sem Assunto)")
            id = email.get("id")
            sender = (
                email.get("from", {})
                .get("emailAddress", {})
                .get("address", "Desconhecido")
            )
            received_date = email.get("receivedDateTime")
            body = email.get("body", {})
            body_type = body.get("contentType", "text")
            body_content = body.get("content", "(E-mail vazio)")
            if body_type == "html":
                content = markdown_texto = md(body_content)
            else:
                content = body_content
            result = {
                "id": id,
                "sender": sender,
                "subject": subject,
                "received_date": received_date,
                "content": content,
                "temAnexo": anexo,
            }
            output.append(result)
        return output
    print(f"Erro {response.status_code}: {response.text}")
    return None


def get_attachments(access_token: str, id: str, path: str):
    headers = {
        # Garanta que o 'access_token' aqui seja estritamente a String do Token
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json",
    }
    url = f"https://graph.microsoft.com/v1.0/me/messages/{id}/attachments"
    response = requests.get(url, headers=headers)
    if response.status_code == 200:
        anexos = response.json().get("value", [])
        for anexo in anexos:
            anexoPath = f"path/{anexo['name']}"
            with open(anexoPath, "wb") as f:
                f.write(base64.b64decode(anexo["contentBytes"]))


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


# --- Execução Principal ---
if __name__ == "__main__":
    # Escopo explícito necessário para leitura

    # Executa o fluxo interativo inteligente
    token_response = get_token()

    if token_response and "access_token" in token_response:
        fim = datetime.today()
        inicio = datetime.today() - timedelta(days=1)
        emails = read_mail(token_response["access_token"], inicio, fim)
        for email in emails:
            if email["temAnexo"]:
                get_attachments(token_response["access_token"], email["id"])

    else:
        print("\nNão foi possível processar a requisição sem um token válido.")
