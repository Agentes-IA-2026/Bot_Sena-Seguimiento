"""
drive_adapter.py
Adaptador Google Drive para GuíaBot.
Funciona con OAuth2 por docente (cada docente autoriza su propia cuenta).
Mantiene toda la lógica de comparación flexible del código original.
"""

import re
from googleapiclient.discovery import build
from google.oauth2.credentials import Credentials


# ── NORMALIZACIÓN ─────────────────────────────────────────────────────────────

def _normalizar(texto: str) -> str:
    texto = texto.lower()
    reemplazos = {
        'á':'a','é':'e','í':'i','ó':'o','ú':'u',
        'ä':'a','ë':'e','ï':'i','ö':'o','ü':'u',
        'à':'a','è':'e','ì':'i','ò':'o','ù':'u',
        'ñ':'n',
    }
    for origen, destino in reemplazos.items():
        texto = texto.replace(origen, destino)
    texto = re.sub(r'\.(pdf|png|jpg|jpeg|xlsx|xls|pptx|ppt|docx|doc|mp4)(\.|$)', ' ', texto)
    texto = re.sub(r'[^a-z0-9\s]', ' ', texto)
    texto = re.sub(r'\s+', ' ', texto).strip()
    return texto


def _palabras_clave(texto_normalizado: str) -> set:
    IGNORAR = {
        'pdf','png','jpg','xlsx','pptx','docx','doc',
        'con','del','los','las','una','uno','por',
        'que','para','como','the','and',
    }
    palabras = texto_normalizado.split()
    return {p for p in palabras if len(p) > 3 and p not in IGNORAR}


def _coincide(nombre_evidencia: str, nombre_drive: str) -> bool:
    ev_norm    = _normalizar(nombre_evidencia)
    drive_norm = _normalizar(nombre_drive)

    if ev_norm in drive_norm or drive_norm in ev_norm:
        return True

    palabras_ev    = _palabras_clave(ev_norm)
    palabras_drive = _palabras_clave(drive_norm)

    if palabras_ev and palabras_drive:
        comunes    = palabras_ev & palabras_drive
        porcentaje = len(comunes) / len(palabras_ev)
        if porcentaje >= 0.5:
            return True

    def bigramas(palabras):
        lista = sorted(palabras)
        return {f"{lista[i]} {lista[i+1]}" for i in range(len(lista)-1)}

    if len(palabras_ev) >= 2 and len(palabras_drive) >= 2:
        if bigramas(palabras_ev) & bigramas(palabras_drive):
            return True

    return False


# ── CONEXIÓN ──────────────────────────────────────────────────────────────────

def conectar(token_dict: dict):
    """
    Conecta a Drive usando el token OAuth2 del docente.
    token_dict viene de Supabase — guardado cuando el docente autorizó su cuenta.
    """
    creds = Credentials(
        token=token_dict["access_token"],
        refresh_token=token_dict.get("refresh_token"),
        token_uri="https://oauth2.googleapis.com/token",
        client_id=token_dict["client_id"],
        client_secret=token_dict["client_secret"],
    )
    return build('drive', 'v3', credentials=creds)


# ── LISTAR ARCHIVOS ───────────────────────────────────────────────────────────

def _listar_archivos_recursivo(service, folder_id: str) -> list:
    nombres = []
    try:
        page_token = None
        while True:
            results = service.files().list(
                q=f"'{folder_id}' in parents and trashed = false",
                fields="nextPageToken, files(id, name, mimeType)",
                pageSize=100,
                pageToken=page_token
            ).execute()

            for item in results.get('files', []):
                if item['mimeType'] == 'application/vnd.google-apps.folder':
                    nombres.extend(_listar_archivos_recursivo(service, item['id']))
                else:
                    nombres.append(item['name'])

            page_token = results.get('nextPageToken')
            if not page_token:
                break

    except Exception as e:
        print(f"Error al listar carpeta {folder_id}: {e}")

    return nombres


# ── VERIFICACIÓN PRINCIPAL ────────────────────────────────────────────────────

def verificar(token_dict: dict, folder_id: str, lista_evidencias: list) -> dict:
    """
    Verifica qué evidencias están presentes en la carpeta Drive del estudiante.

    Parámetros:
        token_dict       → credenciales OAuth2 del docente (de Supabase)
        folder_id        → ID de la carpeta Drive del estudiante
        lista_evidencias → lista de nombres de evidencias requeridas

    Retorna:
        dict { nombre_evidencia: True/False }
    """
    resultados = {ev: False for ev in lista_evidencias}

    try:
        service = conectar(token_dict)
        archivos = _listar_archivos_recursivo(service, folder_id)

        if not archivos:
            return resultados

        for ev in lista_evidencias:
            for nombre_real in archivos:
                if _coincide(ev, nombre_real):
                    resultados[ev] = True
                    break

    except Exception as e:
        print(f"Error Drive verificar: {e}")

    return resultados


def extraer_id_carpeta(link: str) -> str | None:
    """Extrae el ID de Drive de cualquier formato de URL."""
    if not link or not isinstance(link, str):
        return None
    link = link.strip()
    for patron in [
        r'/folders/([a-zA-Z0-9_-]+)',
        r'/file/d/([a-zA-Z0-9_-]+)',
        r'[?&]id=([a-zA-Z0-9_-]+)',
        r'open\?id=([a-zA-Z0-9_-]+)',
    ]:
        m = re.search(patron, link)
        if m:
            return m.group(1)
    if "http" not in link and len(link) > 10:
        return link
    return None
