"""
Executa o fluxo OAuth localmente e salva o token com os escopos
de Calendar + Drive. Rode uma vez:
    python gerar_token_google.py
"""
from google_auth_oauthlib.flow import InstalledAppFlow
from pathlib import Path
import json

SCOPES = [
    'https://www.googleapis.com/auth/calendar.events',
    'https://www.googleapis.com/auth/drive.file',
]

CREDENTIALS_FILE = Path(__file__).parent / 'google_credentials.json'
TOKEN_FILE       = Path(__file__).parent / 'google_token.json'

if not CREDENTIALS_FILE.exists():
    print("ERRO: google_credentials.json não encontrado.")
    exit(1)

flow = InstalledAppFlow.from_client_secrets_file(str(CREDENTIALS_FILE), SCOPES)
creds = flow.run_local_server(port=8080)

TOKEN_FILE.write_text(creds.to_json())
print("\nToken salvo em google_token.json")
print("\nCopie o conteúdo abaixo para a variável GOOGLE_TOKEN_JSON no Vercel:\n")
print(creds.to_json())
