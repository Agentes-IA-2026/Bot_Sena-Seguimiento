import os
import re
from google.oauth2 import service_account
from googleapiclient.discovery import build
 
 
def conectar_drive():
    """Conecta con la API usando el JSON de FaroClick"""
    RUTA_JSON = 'assets/credenciales.json'
    SCOPES = ['https://www.googleapis.com/auth/drive.readonly']
 
    if not os.path.exists(RUTA_JSON):
        print(f"❌ No se encontró el archivo: {RUTA_JSON}")
        return None
 
    creds = service_account.Credentials.from_service_account_file(RUTA_JSON, scopes=SCOPES)
    return build('drive', 'v3', credentials=creds)
 
 
def _normalizar(texto: str) -> str:
    """
    Normaliza un texto para comparación flexible:
    - Minúsculas
    - Sin tildes
    - Sin extensiones de archivo
    - Sin caracteres especiales, solo letras y números
    """
    texto = texto.lower()
 
    # Quitar tildes
    reemplazos = {
        'á': 'a', 'é': 'e', 'í': 'i', 'ó': 'o', 'ú': 'u',
        'ä': 'a', 'ë': 'e', 'ï': 'i', 'ö': 'o', 'ü': 'u',
        'à': 'a', 'è': 'e', 'ì': 'i', 'ò': 'o', 'ù': 'u',
        'ñ': 'n',
    }
    for origen, destino in reemplazos.items():
        texto = texto.replace(origen, destino)
 
    # Quitar extensiones comunes
    texto = re.sub(r'\.(pdf|png|jpg|jpeg|xlsx|xls|pptx|ppt|docx|doc|mp4)(\.|$)', ' ', texto)
 
    # Quitar caracteres especiales, dejar solo letras, números y espacios
    texto = re.sub(r'[^a-z0-9\s]', ' ', texto)
 
    # Colapsar espacios múltiples
    texto = re.sub(r'\s+', ' ', texto).strip()
 
    return texto
 
 
def _palabras_clave(texto_normalizado: str) -> set[str]:
    """
    Extrae palabras significativas (más de 3 letras) de un texto normalizado.
    Ignora palabras muy comunes que no aportan al match.
    """
    PALABRAS_IGNORAR = {
        'pdf', 'png', 'jpg', 'xlsx', 'pptx', 'docx', 'doc',
        'con', 'del', 'los', 'las', 'una', 'uno', 'por',
        'que', 'para', 'como', 'the', 'and', 'sena',  # "sena" aparece en casi todo
    }
    palabras = texto_normalizado.split()
    return {p for p in palabras if len(p) > 3 and p not in PALABRAS_IGNORAR}
 
 
def _coincide(nombre_evidencia: str, nombre_drive: str, debug: bool = False) -> bool:
    """
    Determina si un archivo del Drive corresponde a una evidencia requerida.
    Usa múltiples estrategias de comparación flexible.
    """
    ev_norm = _normalizar(nombre_evidencia)
    drive_norm = _normalizar(nombre_drive)
 
    # Estrategia 1: Contención directa (uno contiene al otro)
    if ev_norm in drive_norm or drive_norm in ev_norm:
        return True
 
    # Estrategia 2: Palabras clave en común
    # Calculamos qué porcentaje de palabras clave de la evidencia aparecen en el archivo
    palabras_ev = _palabras_clave(ev_norm)
    palabras_drive = _palabras_clave(drive_norm)
 
    if palabras_ev and palabras_drive:
        comunes = palabras_ev & palabras_drive
        porcentaje = len(comunes) / len(palabras_ev)
 
        if debug and comunes:
            print(f"      palabras_ev={palabras_ev} | drive={palabras_drive} | "
                  f"comunes={comunes} | {porcentaje:.0%}")
 
        # Si el 50% o más de las palabras clave coinciden → match
        if porcentaje >= 0.5:
            return True
 
    # Estrategia 3: Coincidencia por bigramas (pares de palabras consecutivas)
    # Útil para "quien soy" vs "3.1 quien soy.docx"
    def bigramas(palabras):
        lista = sorted(palabras)
        return {f"{lista[i]} {lista[i+1]}" for i in range(len(lista)-1)}
 
    if len(palabras_ev) >= 2 and len(palabras_drive) >= 2:
        bi_ev = bigramas(palabras_ev)
        bi_drive = bigramas(palabras_drive)
        if bi_ev & bi_drive:
            return True
 
    return False
 
 
def _listar_archivos_recursivo(service, folder_id: str) -> list[str]:
    """
    Lista TODOS los archivos dentro de una carpeta y sus subcarpetas.
    Retorna una lista de nombres de archivo originales.
    """
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
 
            items = results.get('files', [])
 
            for item in items:
                if item['mimeType'] == 'application/vnd.google-apps.folder':
                    sub = _listar_archivos_recursivo(service, item['id'])
                    nombres.extend(sub)
                else:
                    nombres.append(item['name'])
 
            page_token = results.get('nextPageToken')
            if not page_token:
                break
 
    except Exception as e:
        print(f"   ⚠️ Error al listar carpeta {folder_id}: {e}")
 
    return nombres
 
 
def verificar_evidencias_en_carpeta(service, folder_id: str,
                                    lista_evidencias: list[str],
                                    debug: bool = True) -> dict:
    """
    Verifica qué evidencias requeridas están presentes en la carpeta del aprendiz.
    Busca recursivamente en subcarpetas y usa comparación flexible.
    """
    resultados = {ev: False for ev in lista_evidencias}
 
    archivos_en_drive = _listar_archivos_recursivo(service, folder_id)
 
    if not archivos_en_drive:
        print(f"   ⚠️ No se encontraron archivos (carpeta vacía o sin acceso)")
        return resultados
 
    print(f"   📂 {len(archivos_en_drive)} archivo(s) en Drive:")
    for a in archivos_en_drive:
        print(f"      • {a}")
 
    print(f"   🔍 Comparando evidencias...")
    for ev in lista_evidencias:
        for nombre_real in archivos_en_drive:
            if _coincide(ev, nombre_real, debug=debug):
                resultados[ev] = True
                print(f"   ✅ '{ev}' → '{nombre_real}'")
                break
 
    return resultados