import os
import json
from pathlib import Path

SCOPES = ['https://www.googleapis.com/auth/calendar.events']
CREDENTIALS_FILE = Path(__file__).parent.parent / 'google_credentials.json'
TOKEN_FILE        = Path(__file__).parent.parent / 'google_token.json'


def _get_service():
    from google.oauth2.credentials import Credentials
    from google.auth.transport.requests import Request
    from googleapiclient.discovery import build

    creds = None
    if TOKEN_FILE.exists():
        creds = Credentials.from_authorized_user_file(str(TOKEN_FILE), SCOPES)

    if creds and creds.expired and creds.refresh_token:
        creds.refresh(Request())
        TOKEN_FILE.write_text(creds.to_json())

    if not creds or not creds.valid:
        return None  # precisa autorizar

    return build('calendar', 'v3', credentials=creds)


def is_authorized():
    svc = _get_service()
    return svc is not None


def get_auth_url(redirect_uri):
    from google_auth_oauthlib.flow import Flow
    flow = Flow.from_client_secrets_file(
        str(CREDENTIALS_FILE),
        scopes=SCOPES,
        redirect_uri=redirect_uri,
    )
    url, state = flow.authorization_url(access_type='offline', prompt='consent')
    return url, state


def exchange_code(code, state, redirect_uri):
    from google_auth_oauthlib.flow import Flow
    flow = Flow.from_client_secrets_file(
        str(CREDENTIALS_FILE),
        scopes=SCOPES,
        state=state,
        redirect_uri=redirect_uri,
    )
    flow.fetch_token(code=code)
    TOKEN_FILE.write_text(flow.credentials.to_json())


def criar_eventos_parcelas(registro):
    """Cria um evento no Google Calendar para cada parcela não paga do acordo."""
    svc = _get_service()
    if not svc:
        return False, 'Não autorizado'

    ad       = registro.get('acordo_dados') or {}
    parcelas = ad.get('parcelas') or []
    cliente  = registro.get('cliente', 'Cliente')
    cpf      = registro.get('cpf_cnpj', '')
    total    = len(parcelas)
    criados  = 0

    for p in parcelas:
        if p.get('pago'):
            continue

        data_str = p.get('data', '')
        if not data_str:
            continue

        # normaliza para YYYY-MM-DD
        if '/' in data_str:
            parts = data_str.split('/')
            if len(parts) == 3:
                data_str = f'{parts[2]}-{parts[1]}-{parts[0]}'

        valor = p.get('valor', 0)
        numero = p.get('numero', '?')

        titulo = f'Recebimento — {cliente} · Parcela {numero}/{total}'
        descricao_parts = [f'Valor: R$ {float(valor):,.2f}'.replace(',', 'X').replace('.', ',').replace('X', '.')]
        if cpf:
            descricao_parts.append(f'CPF/CNPJ: {cpf}')
        descricao_parts.append('Carteira Judicializada — Ativuz Veículos')

        evento = {
            'summary': titulo,
            'description': '\n'.join(descricao_parts),
            'start': {'date': data_str},
            'end':   {'date': data_str},
            'reminders': {
                'useDefault': False,
                'overrides': [
                    {'method': 'email',  'minutes': 24 * 60},
                    {'method': 'popup',  'minutes': 60},
                ],
            },
        }

        try:
            svc.events().insert(calendarId='primary', body=evento).execute()
            criados += 1
        except Exception as e:
            return False, str(e)

    return True, criados
