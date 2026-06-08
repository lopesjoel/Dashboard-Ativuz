import os
import json
from pathlib import Path

SCOPES = [
    'https://www.googleapis.com/auth/calendar.events',
    'https://www.googleapis.com/auth/drive.file',
]
CREDENTIALS_FILE = Path(__file__).parent.parent / 'google_credentials.json'
TOKEN_FILE        = Path(__file__).parent.parent / 'google_token.json'


def _credentials_dict():
    """Retorna o dict de credenciais — do arquivo local ou da variável de ambiente."""
    env_val = os.environ.get('GOOGLE_CREDENTIALS_JSON')
    if env_val:
        return json.loads(env_val)
    if CREDENTIALS_FILE.exists():
        return json.loads(CREDENTIALS_FILE.read_text())
    return None


def _token_json():
    """Retorna o token salvo — do arquivo local ou da variável de ambiente."""
    env_val = os.environ.get('GOOGLE_TOKEN_JSON')
    if env_val:
        return env_val
    if TOKEN_FILE.exists():
        return TOKEN_FILE.read_text()
    return None


def _save_token(token_json_str):
    """Salva o token no arquivo local (em produção o Railway precisa de outra estratégia)."""
    try:
        TOKEN_FILE.write_text(token_json_str)
    except Exception:
        pass


def _get_service():
    from google.oauth2.credentials import Credentials
    from google.auth.transport.requests import Request
    from googleapiclient.discovery import build

    token_str = _token_json()
    creds = None
    if token_str:
        creds = Credentials.from_authorized_user_info(json.loads(token_str), SCOPES)

    if creds and creds.expired and creds.refresh_token:
        creds.refresh(Request())
        _save_token(creds.to_json())

    if not creds or not creds.valid:
        return None

    return build('calendar', 'v3', credentials=creds)


def is_authorized():
    return _get_service() is not None


def has_credentials():
    return _credentials_dict() is not None


def get_auth_url(redirect_uri):
    from google_auth_oauthlib.flow import Flow
    creds_dict = _credentials_dict()
    if not creds_dict:
        raise RuntimeError('Credenciais não encontradas')
    flow = Flow.from_client_config(creds_dict, scopes=SCOPES, redirect_uri=redirect_uri)
    url, state = flow.authorization_url(access_type='offline', prompt='consent')
    return url, state


def exchange_code(code, state, redirect_uri):
    from google_auth_oauthlib.flow import Flow
    creds_dict = _credentials_dict()
    flow = Flow.from_client_config(creds_dict, scopes=SCOPES, state=state, redirect_uri=redirect_uri)
    flow.fetch_token(code=code)
    _save_token(flow.credentials.to_json())


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
