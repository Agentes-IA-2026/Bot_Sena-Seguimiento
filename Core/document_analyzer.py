import PyPDF2
import os
import re
 
 
# ── MATERIAL DE APOYO A IGNORAR (recursos del instructor, no entregables) ────
MATERIAL_APOYO = {
    "acuerdo_sena_0009_2024.pdf",
    "test de david kolb.xlsx",
    "valores, defectos y cualidades.pdf",
    "kolb y los estilos de aprendizaje.mp4",
    "bienvenida aprendices 2025 – sena.mp4",
    "portafolio_aprendiz.png",
    "fpi modelo pedagógico sena.pptx",
    "etapa productiva.mp4",
}
 
# ── EXTENSIONES VÁLIDAS DE EVIDENCIAS ────────────────────────────────────────
EXTENSIONES = r'(?:pdf|png|jpg|jpeg|xlsx|xls|pptx|ppt|docx|doc|mp4)'
 
# ── TÍTULOS QUE MARCAN EL INICIO DE LA TABLA RESUMEN ────────────────────────
# Agregamos todas las variantes posibles que usan las guías del SENA
TITULOS_TABLA = [
    "PLANTEAMIENTO DE EVIDENCIAS DE APRENDIZAJE",
    "EVIDENCIAS DE APRENDIZAJE PARA LA EVALUACIÓN",
    "PLANTEAMIENTO DE EVIDENCIAS",
    "EVIDENCIAS DE APRENDIZAJE",          # más genérico, como fallback
    "4. PLANTEAMIENTO",
    "EVALUACIÓN EN EL PROCESO FORMATIVO",
]
 
# ── FRASES QUE INDICAN QUE ESTAMOS DENTRO DE LA COLUMNA DE EVIDENCIAS ────────
MARCADORES_COLUMNA = [
    "evidencias de aprendizaje",
    "evidencia de aprendizaje",
    "evidencias:",
    "evidencia:",
    "producto esperado",
    "productos esperados",
]
 
 
def _encontrar_texto_tabla(paginas: list[str], debug: bool) -> str:
    """
    Recorre las páginas buscando el encabezado de la tabla resumen de evidencias.
    Devuelve el texto desde esa página hasta el final del PDF.
    Funciona sin importar en qué página esté la tabla ni cuántas páginas tenga el PDF.
    """
    for i, texto_pagina in enumerate(paginas):
        texto_upper = texto_pagina.upper()
        for titulo in TITULOS_TABLA:
            if titulo.upper() in texto_upper:
                if debug:
                    print(f"📌 Tabla de evidencias detectada en página {i + 1} "
                          f"(coincidencia: '{titulo}')")
                # Devolvemos desde esta página hasta el final
                return "\n".join(paginas[i:])
 
    if debug:
        print("⚠️  No se encontró ningún encabezado de tabla conocido.")
    return ""
 
 
def _extraer_archivos_de_texto(texto: str) -> list[str]:
    """
    Extrae nombres de archivo con extensión desde un bloque de texto.
    Filtra material de apoyo y duplicados.
    """
    patron = rf'([\w\s\-_áéíóúÁÉÍÓÚñÑ,()]+\.{EXTENSIONES})'
    hallazgos = re.findall(patron, texto, re.IGNORECASE)
 
    evidencias = []
    for item in hallazgos:
        item_limpio = item.strip()
        item_lower = item_limpio.lower()
 
        if item_lower in MATERIAL_APOYO:
            continue
        if len(item_limpio) < 5:
            continue
        ya_existe = any(e.lower() == item_lower for e in evidencias)
        if not ya_existe:
            evidencias.append(item_limpio)
 
    return evidencias
 
 
def extraer_nombres_evidencias(ruta_pdf: str, debug: bool = True) -> list[str]:
    """
    Extrae los nombres de evidencias ÚNICAMENTE desde la tabla resumen del PDF,
    sin importar en qué página esté ni cuántas páginas tenga el documento.
 
    Estrategia en cascada:
      1. Busca el encabezado de la tabla resumen (sección 4 del formato SENA)
      2. Si no lo encuentra, busca párrafos que contengan "Evidencias:" por actividad
      3. Si tampoco, escanea todo el PDF como último recurso
      4. Si el PDF no se puede leer, usa la lista manual hardcodeada
    """
    if not os.path.exists(ruta_pdf):
        print(f"❌ No se encontró la guía en: {ruta_pdf}")
        return _evidencias_manual(ruta_pdf, debug)
 
    try:
        with open(ruta_pdf, 'rb') as archivo:
            lector = PyPDF2.PdfReader(archivo)
            paginas = [p.extract_text() or "" for p in lector.pages]
 
        if debug:
            print(f"📄 PDF leído: {len(paginas)} páginas — {ruta_pdf}")
 
        # ── ESTRATEGIA 1: Tabla resumen (sección 4) ──────────────────────────
        texto_tabla = _encontrar_texto_tabla(paginas, debug)
        if texto_tabla:
            evidencias = _extraer_archivos_de_texto(texto_tabla)
            if evidencias:
                if debug:
                    _imprimir_lista(evidencias, "tabla resumen")
                return evidencias
            elif debug:
                print("   ⚠️  Tabla encontrada pero sin archivos con extensión.")
 
        # ── ESTRATEGIA 2: Bloques "Evidencias:" por actividad ────────────────
        if debug:
            print("🔄 Intentando estrategia 2: bloques 'Evidencias:' por actividad...")
 
        texto_completo = "\n".join(paginas)
        bloques_evidencia = []
 
        for marcador in MARCADORES_COLUMNA:
            # Buscamos el marcador y tomamos el texto siguiente (hasta 400 chars)
            patron_bloque = rf'{re.escape(marcador)}(.{{0,400}})'
            matches = re.findall(patron_bloque, texto_completo, re.IGNORECASE | re.DOTALL)
            bloques_evidencia.extend(matches)
 
        if bloques_evidencia:
            texto_bloques = "\n".join(bloques_evidencia)
            evidencias = _extraer_archivos_de_texto(texto_bloques)
            if evidencias:
                if debug:
                    _imprimir_lista(evidencias, "bloques de evidencias")
                return evidencias
 
        # ── ESTRATEGIA 3: Todo el PDF (último recurso) ───────────────────────
        if debug:
            print("🔄 Intentando estrategia 3: escaneo completo del PDF...")
 
        evidencias = _extraer_archivos_de_texto(texto_completo)
        if evidencias:
            if debug:
                _imprimir_lista(evidencias, "escaneo completo")
            return evidencias
 
        # ── ESTRATEGIA 4: Lista manual ────────────────────────────────────────
        print("⚠️  No se encontraron evidencias en el PDF. Usando lista manual.")
        return _evidencias_manual(ruta_pdf, debug)
 
    except Exception as e:
        print(f"❌ Error leyendo PDF '{ruta_pdf}': {e}")
        return _evidencias_manual(ruta_pdf, debug)
 
 
def _imprimir_lista(evidencias: list[str], fuente: str):
    print(f"\n📋 Evidencias extraídas desde {fuente} ({len(evidencias)}):")
    for i, ev in enumerate(evidencias, 1):
        print(f"   {i}. '{ev}'")
 
 
# ── LISTAS MANUALES POR GUÍA ─────────────────────────────────────────────────
# Si el PDF de una guía no extrae bien, define su lista aquí.
# La clave es el nombre del archivo PDF (en minúsculas).
LISTAS_MANUALES = {
    "guía_00_inducción.pdf": [
        "Quien soy.jpg",
        "Mi estilo de aprendizaje.xlsx",
        "Info_Identidad SENA.pdf",
        "Plataformas SENA.pdf",
        "Mi Programa de Formación.pptx",
        "Chat Reglamento Aprendiz.pdf",
        "Propuesta Proyecto.pdf",
        "Línea de tiempo Profesional.png",
    ],
    # Agrega aquí las demás guías cuando las necesites:
    # "guía_01_diagnóstico_empresarial.pdf": [...],
    # "guía_02_segmentación.pdf": [...],
}
 
 
def _evidencias_manual(ruta_pdf: str = "", debug: bool = True) -> list[str]:
    """
    Devuelve la lista manual de evidencias para una guía específica.
    Si no hay lista manual para esa guía, devuelve lista vacía.
    """
    nombre_pdf = os.path.basename(ruta_pdf).lower() if ruta_pdf else ""
 
    for clave, lista in LISTAS_MANUALES.items():
        if clave.lower() in nombre_pdf or nombre_pdf in clave.lower():
            if debug:
                _imprimir_lista(lista, f"lista manual ({clave})")
            return lista
 
    if debug:
        print(f"❌ No hay lista manual para: '{nombre_pdf}'")
        print("   Agrega la lista en LISTAS_MANUALES dentro de document_analyzer.py")
    return []
 
 
def extraer_nombres_evidencias_manual(ruta_pdf: str = "") -> list[str]:
    """Alias para usar la lista manual directamente desde main.py."""
    return _evidencias_manual(ruta_pdf, debug=True)