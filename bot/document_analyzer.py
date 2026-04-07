"""
document_analyzer.py
Extrae evidencias requeridas desde PDFs de guías.
Versión web — lee PDFs desde bytes (Supabase Storage o URL)
en lugar de rutas locales.
Mantiene las 4 estrategias en cascada del código original.
"""

import re
import io
import PyPDF2


# ── CONFIGURACIÓN ─────────────────────────────────────────────────────────────

MATERIAL_APOYO = {
    "test de david kolb.xlsx",
    "valores, defectos y cualidades.pdf",
    "kolb y los estilos de aprendizaje.mp4",
    "portafolio_aprendiz.png",
    "fpi modelo pedagogico sena.pptx",
    "etapa productiva.mp4",
    "acuerdo sena 0009 2024.pdf",
}

EXTENSIONES = r'(?:pdf|png|jpg|jpeg|xlsx|xls|pptx|ppt|docx|doc|mp4)'

TITULOS_TABLA = [
    "PLANTEAMIENTO DE EVIDENCIAS DE APRENDIZAJE",
    "EVIDENCIAS DE APRENDIZAJE PARA LA EVALUACIÓN",
    "PLANTEAMIENTO DE EVIDENCIAS",
    "EVIDENCIAS DE APRENDIZAJE",
    "4. PLANTEAMIENTO",
    "EVALUACIÓN EN EL PROCESO FORMATIVO",
]

MARCADORES_COLUMNA = [
    "evidencias de aprendizaje",
    "evidencia de aprendizaje",
    "evidencias:",
    "evidencia:",
    "producto esperado",
    "productos esperados",
]


# ── EXTRACCIÓN DESDE PDF ──────────────────────────────────────────────────────

def _encontrar_texto_tabla(paginas: list) -> str:
    for i, texto_pagina in enumerate(paginas):
        texto_upper = texto_pagina.upper()
        for titulo in TITULOS_TABLA:
            if titulo.upper() in texto_upper:
                return "\n".join(paginas[i:])
    return ""


def _extraer_archivos_de_texto(texto: str) -> list:
    patron = rf'([\w\s\-_áéíóúÁÉÍÓÚñÑ,()]+\.{EXTENSIONES})'
    hallazgos = re.findall(patron, texto, re.IGNORECASE)

    evidencias = []
    for item in hallazgos:
        item_limpio = item.strip()
        item_lower  = item_limpio.lower()

        if item_lower in MATERIAL_APOYO:
            continue
        if len(item_limpio) < 5:
            continue
        if not any(e.lower() == item_lower for e in evidencias):
            evidencias.append(item_limpio)

    return evidencias


def extraer_desde_bytes(pdf_bytes: bytes) -> list:
    """
    Extrae nombres de evidencias desde los bytes de un PDF.
    Versión web — recibe bytes en lugar de ruta de archivo.
    Úsala cuando el PDF viene de Supabase Storage o de una URL.

    Estrategia en cascada (igual que el original):
      1. Tabla resumen (sección 4 del formato)
      2. Bloques 'Evidencias:' por actividad
      3. Escaneo completo del PDF
      4. Lista manual como respaldo (ver LISTAS_MANUALES abajo)
    """
    try:
        lector  = PyPDF2.PdfReader(io.BytesIO(pdf_bytes))
        paginas = [p.extract_text() or "" for p in lector.pages]

        # Estrategia 1: Tabla resumen
        texto_tabla = _encontrar_texto_tabla(paginas)
        if texto_tabla:
            evidencias = _extraer_archivos_de_texto(texto_tabla)
            if evidencias:
                return evidencias

        # Estrategia 2: Bloques "Evidencias:"
        texto_completo  = "\n".join(paginas)
        bloques         = []
        for marcador in MARCADORES_COLUMNA:
            patron_bloque = rf'{re.escape(marcador)}(.{{0,400}})'
            matches = re.findall(patron_bloque, texto_completo,
                                 re.IGNORECASE | re.DOTALL)
            bloques.extend(matches)

        if bloques:
            evidencias = _extraer_archivos_de_texto("\n".join(bloques))
            if evidencias:
                return evidencias

        # Estrategia 3: Todo el PDF
        evidencias = _extraer_archivos_de_texto(texto_completo)
        if evidencias:
            return evidencias

        return []

    except Exception as e:
        print(f"Error leyendo PDF: {e}")
        return []


def extraer_desde_lista_manual(nombre_guia: str) -> list:
    """
    Retorna la lista manual de evidencias para una guía específica.
    Úsala como respaldo cuando el PDF no extrae bien.
    El docente puede editar estas listas desde el dashboard.
    """
    nombre_lower = nombre_guia.lower()
    for clave, lista in LISTAS_MANUALES.items():
        if clave.lower() in nombre_lower or nombre_lower in clave.lower():
            return lista
    return []


# ── LISTAS MANUALES POR GUÍA ──────────────────────────────────────────────────
# En la versión web estas listas también se pueden guardar en Supabase
# para que el docente las edite desde el dashboard sin tocar código.

LISTAS_MANUALES = {
    "guia_00_induccion.pdf": [
        "Quien soy.jpg",
        "Mi estilo de aprendizaje.xlsx",
        "Info_Identidad.pdf",
        "Plataformas.pdf",
        "Mi Programa de Formacion.pptx",
        "Chat Reglamento Aprendiz.pdf",
        "Propuesta Proyecto.pdf",
        "Linea de tiempo Profesional.png",
    ],
    # Agrega aquí las guías de cada docente según se vayan configurando.
    # En producción estas listas viven en Supabase tabla 'guias_manuales'.
}
