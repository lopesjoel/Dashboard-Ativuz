import os
import io
import json
from pathlib import Path

_CREDENTIALS_FILE = Path(__file__).parent.parent / 'google_credentials.json'
_TOKEN_FILE       = Path(__file__).parent.parent / 'google_token.json'

_SCOPES = [
    'https://www.googleapis.com/auth/calendar.events',
    'https://www.googleapis.com/auth/drive.file',
]


def _credentials_dict():
    env_val = os.environ.get('GOOGLE_CREDENTIALS_JSON')
    if env_val:
        return json.loads(env_val)
    if _CREDENTIALS_FILE.exists():
        return json.loads(_CREDENTIALS_FILE.read_text())
    return None


def _token_json():
    env_val = os.environ.get('GOOGLE_TOKEN_JSON')
    if env_val:
        return env_val
    if _TOKEN_FILE.exists():
        return _TOKEN_FILE.read_text()
    return None


def _save_token(token_json_str):
    try:
        _TOKEN_FILE.write_text(token_json_str)
    except Exception:
        pass


def _get_drive_service():
    from google.oauth2.credentials import Credentials
    from google.auth.transport.requests import Request
    from googleapiclient.discovery import build

    token_str = _token_json()
    if not token_str:
        return None

    creds = Credentials.from_authorized_user_info(json.loads(token_str), _SCOPES)
    if creds and creds.expired and creds.refresh_token:
        creds.refresh(Request())
        _save_token(creds.to_json())

    if not creds or not creds.valid:
        return None

    return build('drive', 'v3', credentials=creds)


def docx_bytes_to_pdf(docx_bytes: bytes, filename: str = "documento.docx") -> bytes:
    """
    Converte DOCX para PDF via Google Drive API:
      1. Faz upload do DOCX (Drive converte para Google Docs)
      2. Exporta como PDF
      3. Deleta o arquivo temporário do Drive
    Requer token OAuth com escopo drive.file.
    """
    from googleapiclient.http import MediaIoBaseUpload, MediaIoBaseDownload

    svc = _get_drive_service()
    if not svc:
        raise RuntimeError(
            "Google Drive não autorizado. "
            "Acesse /api/google-calendar/auth para autorizar e obter um token "
            "com o escopo drive.file, depois copie o conteúdo de google_token.json "
            "para a variável de ambiente GOOGLE_TOKEN_JSON no Vercel."
        )

    media = MediaIoBaseUpload(
        io.BytesIO(docx_bytes),
        mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        resumable=False,
    )
    try:
        uploaded = svc.files().create(
            body={
                'name': filename,
                'mimeType': 'application/vnd.google-apps.document',
            },
            media_body=media,
            fields='id',
        ).execute()
    except Exception as e:
        raise RuntimeError(
            f"Falha no upload para Google Drive (verifique se o token tem escopo drive.file): {e}"
        )

    file_id = uploaded['id']
    try:
        request = svc.files().export_media(fileId=file_id, mimeType='application/pdf')
        buf = io.BytesIO()
        downloader = MediaIoBaseDownload(buf, request)
        done = False
        while not done:
            _, done = downloader.next_chunk()
        return buf.getvalue()
    finally:
        try:
            svc.files().delete(fileId=file_id).execute()
        except Exception:
            pass


def is_available() -> bool:
    """Retorna True se o token com escopo drive.file está disponível e válido."""
    try:
        svc = _get_drive_service()
        if not svc:
            return False
        # Testa uma chamada mínima para confirmar o escopo
        svc.files().list(pageSize=1, fields='files(id)').execute()
        return True
    except Exception:
        return False
