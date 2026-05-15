import os
import re

# ── DEPENDENCIAS PDF (intentamos PyMuPDF primero, PyPDF2 como respaldo) ──────
try:
    import fitz  # PyMuPDF
    _PYMUPDF = True
except ImportError:
    _PYMUPDF = False
    try:
        import PyPDF2
        _PYPDF2 = True
    except ImportError:
        _PYPDF2 = False

# ── MATERIAL DE APOYO A IGNORAR ───────────────────────────────────────────────
MATERIAL_APOYO = {
    "acuerdo_sena_0009_2024.pdf",
    "test de david kolb.xlsx",
    "valores, defectos y cualidades.pdf",
    "kolb y los estilos de aprendizaje.mp4",
    "bienvenida aprendices 2025 – sena.mp4",
    "portafolio_aprendiz.png",
    "fpi modelo pedagógico sena.pptx",
    "etapa productiva.mp4",
    "gfpi-f-135 v04",
}

# ── MARCADORES QUE INICIAN LA SECCIÓN DE EVIDENCIAS EN EL PDF ────────────────
TITULOS_TABLA = [
    "PLANTEAMIENTO DE EVIDENCIAS DE APRENDIZAJE",
    "EVIDENCIAS DE APRENDIZAJE PARA LA EVALUACIÓN",
    "PLANTEAMIENTO DE EVIDENCIAS",
]

# ── ETIQUETAS QUE INTRODUCEN UNA LISTA DE EVIDENCIAS ─────────────────────────
ETIQUETAS_EV = [
    "Evidencias de Conocimiento:",
    "Evidencias de Desempeño:",
    "Evidencias de Producto:",
    "Evidencias de conocimiento:",
    "Evidencias de desempeño:",
    "Evidencias de producto:",
    "Evidencia de Conocimiento:",
    "Evidencia de Desempeño:",
    "Evidencia de Producto:",
    "Evidencia de conocimiento:",
    "Evidencia de desempeño:",
    "Evidencia de producto:",
    "Evidencias de desepeño:",
    "Conocimiento:",
    "Desempeño:",
    "Producto:",
]

# ── PALABRAS QUE SEÑALAN FIN DE BLOQUE ───────────────────────────────────────
FIN_BLOQUE = [
    "criterios de evaluación", "criterios de", "técnicas e instrumentos",
    "técnicas e", "instrumento:", "instrumentos:", "5. glosario",
    "glosario de", "6. referentes", "referentes bib",
    "determina las", "seleccionar ", "aplica ", "elabora ", "registra ",
    "gestiona ", "realiza ", "genera ", "documenta ", "compara ", "prepara ",
    "estructura ", "presenta ", "construye ", "utiliza ", "analiza ",
    "demuestra ", "fomenta ", "evalúa ", "obtiene ", "establece ",
    "indaga ", "reporta ",
]


def _leer_paginas(ruta_pdf: str) -> list:
    if _PYMUPDF:
        try:
            doc = fitz.open(ruta_pdf)
            return [page.get_text() for page in doc]
        except Exception:
            pass
    if _PYPDF2:
        try:
            with open(ruta_pdf, "rb") as f:
                reader = PyPDF2.PdfReader(f)
                return [p.extract_text() or "" for p in reader.pages]
        except Exception:
            pass
    return []


def _norm(texto: str) -> str:
    texto = texto.lower()
    for a, b in [("á","a"),("é","e"),("í","i"),("ó","o"),("ú","u"),("ñ","n")]:
        texto = texto.replace(a, b)
    return texto


def _es_fin_bloque(linea: str) -> bool:
    norm = _norm(linea)
    return any(norm.startswith(f) or f in norm for f in FIN_BLOQUE)


def _limpiar(nombre: str) -> str:
    nombre = re.sub(r'\s+', ' ', nombre).strip()
    nombre = re.sub(r'GFPI-F-\d+\s*V\d+', '', nombre, flags=re.IGNORECASE).strip()
    return nombre


def _reconstruir_fragmentos(lineas: list) -> list:
    resultado = []
    buf = ""
    for linea in lineas:
        linea = linea.strip()
        if not linea:
            if buf:
                resultado.append(buf.strip())
                buf = ""
            continue
        if _es_fin_bloque(linea):
            if buf:
                resultado.append(buf.strip())
                buf = ""
            break
        if not buf:
            buf = linea
        else:
            sigue = (
                buf.endswith("_") or buf.endswith("-") or buf.endswith(",") or
                (len(buf) < 35 and not buf.endswith("."))
            )
            cont = (linea[0].islower() or linea[0] == "_" or linea[0].isdigit())
            if sigue or cont:
                buf = buf + linea
            else:
                resultado.append(buf.strip())
                buf = linea
    if buf:
        resultado.append(buf.strip())
    return resultado


def _extraer_de_texto(texto: str, debug: bool = False) -> list:
    evidencias = []
    vistas = set()

    # Estrategia A: nombres con extensión de archivo
    EXT = r'(?:pdf|png|jpg|jpeg|xlsx|xls|pptx|ppt|docx|doc|mp4)'
    for m in re.finditer(rf'([\w\s\-_áéíóúÁÉÍÓÚñÑ,()\[\]]+\.{EXT})', texto, re.IGNORECASE):
        nombre = _limpiar(m.group(1))
        nl = nombre.lower()
        if nl in MATERIAL_APOYO or len(nombre) < 5 or nl in vistas:
            continue
        vistas.add(nl)
        evidencias.append(nombre)

    # Estrategia B: bloques bajo etiquetas "Evidencias de X:"
    posiciones = []
    for etiq in ETIQUETAS_EV:
        for m in re.finditer(re.escape(etiq), texto, re.IGNORECASE):
            posiciones.append((m.start(), m.end()))
    posiciones.sort()

    for i_etiq, (inicio, fin) in enumerate(posiciones):
        siguiente = posiciones[i_etiq + 1][0] if i_etiq + 1 < len(posiciones) else len(texto)
        bloque = texto[fin:siguiente]
        lineas = [l.strip() for l in bloque.split('\n') if l.strip()]
        for nombre in _reconstruir_fragmentos(lineas):
            nombre = _limpiar(nombre)
            nl = nombre.lower()
            if len(nombre) < 4 or nl in MATERIAL_APOYO or _es_fin_bloque(nombre) or nl in vistas:
                continue
            vistas.add(nl)
            evidencias.append(nombre)

    return evidencias


def extraer_nombres_evidencias(ruta_pdf: str, debug: bool = True) -> list:
    if not os.path.exists(ruta_pdf):
        print(f"❌ No se encontró la guía en: {ruta_pdf}")
        return _evidencias_manual(ruta_pdf, debug)

    paginas = _leer_paginas(ruta_pdf)
    if not paginas:
        print(f"❌ No se pudo leer el PDF: {ruta_pdf}")
        return _evidencias_manual(ruta_pdf, debug)

    if debug:
        print(f"   📄 PDF leído: {len(paginas)} páginas — {os.path.basename(ruta_pdf)}")

    texto_completo = "\n".join(paginas)

    # Estrategia 1: sección de evidencias
    idx = -1
    for titulo in TITULOS_TABLA:
        i = texto_completo.upper().find(titulo.upper())
        if i >= 0:
            idx = i
            if debug:
                print(f"   📌 Sección encontrada: '{titulo}'")
            break

    if idx >= 0:
        evidencias = _extraer_de_texto(texto_completo[idx: idx + 5000], debug)
        if evidencias:
            if debug:
                _imprimir_lista(evidencias, "sección PDF")
            return evidencias
        elif debug:
            print("   ⚠️  Sección encontrada pero sin evidencias. Intentando escaneo completo...")

    # Estrategia 2: todo el PDF
    evidencias = _extraer_de_texto(texto_completo, debug=False)
    if evidencias:
        if debug:
            _imprimir_lista(evidencias, "escaneo completo PDF")
        return evidencias

    # Estrategia 3: lista manual
    print(f"   ⚠️  Sin evidencias en PDF. Usando lista manual: {os.path.basename(ruta_pdf)}")
    return _evidencias_manual(ruta_pdf, debug)


def extraer_nombres_evidencias_manual(ruta_pdf: str = "") -> list:
    """Alias para usar la lista manual directamente desde main.py."""
    return _evidencias_manual(ruta_pdf, debug=False)


# ─────────────────────────────────────────────────────────────────────────────
# LISTAS MANUALES COMPLETAS — RESPALDO PARA TODOS LOS PROGRAMAS
# ─────────────────────────────────────────────────────────────────────────────

LISTAS_MANUALES: dict = {

    # ── GUÍA 00 INDUCCIÓN (común a los 3 programas) ───────────────────────────
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

    # ── ASISTENCIA COMERCIAL ──────────────────────────────────────────────────

    "guía_01_ diagnóstico_empresarial.pdf": [
        "01_3_3_1_TALLER_La importancia de emprender",
        "01_3_3_2_TALLER_Socialización ideas innovadoras",
        "01_3_4_INFORME_Diagnóstico empresarial",
    ],

    "guía_02_segmentación_cliente.pdf": [
        "02_3_2_ACTIVIDAD_Mapa Conceptual Movistar Grooming",
        "02_3_3_1_Ejercicio de Relación",
        "02_3_3_4_Prueba de conocimientos guía 02",
        "02_3_3_2_Matriz estudio con registro de información",
        "02_3_3_3_Tipología y perfil del cliente",
        "02_3_4_TALLER_INFORME_Segmentación de Clientes",
    ],

    "guía_03_prospección_de_clientes.pdf": [
        "03_3_1_TALLER_Lista de Regalos",
        "03_3_2_TALLER_Centro Comercial",
        "03_3_3_TALLER_Presentación Buyer Persona",
        "03_3_3_Torneo de aprendizaje",
        "03_3_3_Diseño encuesta",
        "03_3_3_TALLER_Conociendo el Mundo",
        "03_3_4_INFORME_Prospección de clientes",
    ],

    "guía_04_portafolio_de_productos_y_servicios.pdf": [
        "04_3_3_4_Prueba de conocimientos Guía 04",
        "04_3_3_1_TALLER_Fotografía Producto Fichas Envase-Empaque y embalaje",
        "04_3_3_2_TALLER_Marketing Mix Costos y Precios",
        "04_3_3_3_TALLER_Portafolio de Productos y Servicios y Catálogo",
        "04_3_4_INFORME_Portafolio de Productos o Servicios",
    ],

    "guía_05_herramientas_ofimáticas_de_mercadeo.pdf": [
        "05_3_TALLER_PROPUESTA_HERRAMIENTAS_DE_MERCADEO",
        "05_3_INFOGRAFIA_PROPUESTA_COMERCIAL",
        "05_3_4_INFORME_Propuesta comercial con herramientas ofimáticas",
    ],

    "guía_06_surtido_exhibición_y_merchandising.pdf": [
        "DINÁMICA_Surtido_Exhibición_y_Merchandising",
        "06_3_3_2_PRUEBA_surtido_exhibicion_y_Merchandising",
        "006_3_3_1_TALLER_Surtido",
        "06_3_3_2_Infografía_Supermercados",
        "06_3_3_3_EXPO_trabajo_de_campo",
        "06_3_4_INFORME_Surtido_Exhibicion_Merchandising",
    ],

    "guía_07_negociaciones_de_ventas.pdf": [
        "07_1_ACTIVIDAD_Aprendiendo a Negociar",
        "07_2_1_ACTIVIDAD_Taller Mundo Alterno",
        "07_3_1_ACTIVIDAD_Taller Comunicación Asertiva",
        "07_3_2_ACTIVIDAD_Taller Integral de Ventas",
        "07_3_3_ACTIVIDAD_Taller Rueda de Negocios",
        "07_4_Video o Representacion Física Ejecución Del Proceso De Ventas",
    ],

    "guía_08_postventa_y_satisfacción_ al_cliente.pdf": [
        "08_3_1_ACTIVIDAD_Momentos Claves y Críticos de la verdad",
        "08_3_2_ACTIVIDAD_para Desarrollar en Clase",
        "08_3_3_TALLER_Roles Postventa Acción y Satisfacción al Cliente PQRSF",
        "08_3_4_ACTIVIDAD_Comic Trazabilidad del Servicio",
        "08_3_5_ACTIVIDAD_Encuesta Herramientas Hallazgos Atención al Cliente",
        "08_3_4_INFORME_Postventa y satisfacción al cliente",
    ],

    # ── COMUNICACIÓN Y MARKETING ──────────────────────────────────────────────

    "guía_01_ clasificar_clientes.pdf": [
        "Evidencia transferencia de conocimientos",
        "Evidencia de conocimiento",
    ],

    "guía_02_preparar_exhibición.pdf": [
        "Prueba de conocimiento implantación de superficie de venta",
        "Informe layout punto de venta",
    ],

    "guia_03_publicar_ contenidos_y_anuncios.pdf": [
        "Valoración de conocimientos marketing digital",
        "Propuesta de contenidos y anuncios publicados",
    ],

    "guía_04_organizar_exhibición.pdf": [
        "Evidencia de conocimiento Conceptos claves Clasificación y tipología de productos",
        "Participación en actividades 3.3.4 a 3.3.9",
        "Informe propuesta de exhibición",
    ],

    "guía_05_ ejecutar_acciones_promocionales.pdf": [
        "Diseño BRIEF EVENTO.doc",
        "brief EVENTO.doc",
        "INFORME FINAL EVENTO.docx",
        "LISTA DE CHEQUEO EVENTO.docx",
    ],

    "guía_06_vender_productos.pdf": [
        "Prueba de conocimiento ventas",
        "Feria Comercial ejecución del proceso de venta",
        "Informe de Protocolo de venta con propuesta comercial y material de apoyo",
    ],

    "guia_07_interactuar_ con_ audiencias.pdf": [
        "Valoración de conocimientos interacción con audiencias",
        "Manejo de comunicaciones con la audiencia",
        "Informe interacción con audiencias digitales",
    ],

    "guía_08_realizar_seguimiento.pdf": [
        "Evidencia de conocimiento seguimiento",
        "Evidencia de desempeño seguimiento",
        "Informe de seguimiento",
    ],

    "guía_09_monitorear_métricas.pdf": [
        "Valoración de conocimientos métricas",
        "Informe de métricas de la campaña",
    ],

    # ── VENTAS DE PRODUCTOS EN LÍNEA ──────────────────────────────────────────

    "guía_01_construir_una_relación_con_el_cliente.pdf": [
        "01_3_3_TALLER_Mercado digital",
        "01_3_1_VOCABULARIO_Marketing digital",
        "01_4_INFORME_Cliente y atención",
    ],

    "guía_02_procedimientos_con_el_cliente.pdf": [
        "02_3_TALLER_Requerimientos del cliente",
        "02_3_PRUEBA_Procedimientos con el cliente",
        "02_4_INFORME_Recepcionar requerimientos",
    ],

    "guía_03_optimizar_la_atención_al_cliente.pdf": [
        "02_3_TALLER_Atención cliente",
        "02_3_PRUEBA_Procedimientos con el cliente",
        "02_4_INFORME_Optimizar atención cliente",
    ],

    "guía_04_postventa_digital_y _satisfacción_del_cliente..pdf": [
        "04_3_3_4_PRUEBA_de_conocimientos_Guía 04",
        "04_3_3_1_Marketing_Relacional_En_El_Entorno_Digital",
        "04_3_3_2_Plataformas_CRM_Digital",
        "04_3_3_3_JUEGO_ROLES_Métodos de monitoreo postventa",
        "04_3_4_INFORME_Protocolo postventa",
    ],

    "guía_05_propuesta_comercial_por_canales_digitales.pdf": [
        "05_3_3_2_Mapa_de_empatia_consumidor_digital",
        "05_3_3_1_Caracterización_del_Producto",
        "05_3_3_3_Arquetipo_segmento_objetivo",
        "05_3_3_4_Investigación_básica_de_mercados_digitales",
        "05_3_3_5_Buyer_person",
        "05_3_3_6_Plan_de_Marketing_Digital",
        "05_3_3_7_Propuesta_Comercial_para_buyer_persona",
        "05_3_4_INFORME_Propuesta_comercial_digital_para_buyer_person",
    ],

    "guía_06_efectuar_la_venta_de_productos_en_línea.pdf": [
        "06_3_1_Canales y confidencialidad",
        "06_3_2_Comercio_Online",
        "06_3_3_1_TALLER_Omnicanalidad",
        "06_3_3_2_GRÁFICA_Comunicación Digital",
        "06_3_3_3_GRÁFICA_Objeciones",
        "06_3_3_4_DIAGRAMA_Promoción Marketing Digital",
        "06_3_PRUEBA_Efectuar la venta de productos en línea",
        "06_4_INFORME_Venta de Productos en Línea",
    ],

    "guía_07_aplicar_acciones_de_seguimiento_y_control.pdf": [
        "07_2_TALLER_Ventas_Efectivas_y_Postventa_Inteligente",
        "07_3_1_TALLER_Conquistando_y_Fidelizando_Clientes_Online",
        "07_3_2_TALLER_Seguimiento_concepto_técnicas_instrumentos_y_registros",
        "07_3_3_TALLER_SIMULACIÓN_Datos_y_Análisis",
        "07_3_4_TALLER_Seguimiento_concepto_técnicas_instrumentos_y_registros",
        "07_4_REPORTE_Efectividad_y_Análisis_de_las_Campañas",
    ],
}


def _evidencias_manual(ruta_pdf: str = "", debug: bool = True) -> list:
    nombre_pdf = os.path.basename(ruta_pdf).lower() if ruta_pdf else ""
    for clave, lista in LISTAS_MANUALES.items():
        if clave == nombre_pdf or clave in nombre_pdf or nombre_pdf in clave:
            if debug:
                _imprimir_lista(lista, f"lista manual ({clave})")
            return lista
    if debug:
        print(f"   ❌ Sin lista manual para: '{nombre_pdf}'")
    return []


def _imprimir_lista(evidencias: list, fuente: str):
    print(f"\n   📋 Evidencias desde {fuente} ({len(evidencias)}):")
    for i, ev in enumerate(evidencias, 1):
        print(f"      {i}. '{ev}'")
