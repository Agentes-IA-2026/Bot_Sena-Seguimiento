"""
evaluador.py
Define los criterios de evaluación de evidencias para cada guía del sistema GuíaBot.

Estructura pública:
    EVALUADORES: dict[str, EvaluadorGuia]
        Clave = nombre exacto de la guía (igual que en actividades_parametrizadas).

    get_evaluador(nombre_guia, programa=None) -> EvaluadorGuia | None
        Resolución tolerante: exacto → normalizado-exacto → empieza-por → contiene.
        Devuelve None (con aviso) si no hay evaluador definido; el caller puede
        operar en modo solo ENTREGADO/NO ENTREGADO.

    listar_guias(programa=None) -> list[str]
    contar_criterios(nombre_guia) -> dict

Criterio de aprobación estándar (salvo indicación contraria en el Excel madre):
    cumple  >= 0.60  -> CUMPLE
    >= 0.30          -> PARCIAL
    < 0.30           -> NO CUMPLE

Los criterios de cada evidencia son rúbricas observables: se pueden verificar
leyendo el archivo sin necesidad de IA de contenido.
"""

from __future__ import annotations

import unicodedata
from typing import TypedDict


# ─────────────────────────────────────────────────────────────────────────────
# TIPOS PÚBLICOS
# ─────────────────────────────────────────────────────────────────────────────

class CriterioAprobacion(TypedDict):
    """Umbrales de aprobación para una guía."""
    cumple: float
    parcial: float
    descripcion: str


class Evidencia(TypedDict):
    """Definición de una evidencia evaluable dentro de una guía."""
    nombre: str        # nombre_esperado en actividades_parametrizadas
    tipo: str          # documento, imagen, presentacion, excel, cualquier
    descripcion: str   # descripción legible del entregable
    criterios: list[str]


class EvaluadorGuia(TypedDict):
    """Evaluador completo para una guía: metadatos + lista de evidencias."""
    guia: str
    programa: str
    criterio_aprobacion: CriterioAprobacion
    evidencias: list[Evidencia]


__all__ = [
    "EVALUADORES",
    "EvaluadorGuia",
    "Evidencia",
    "CriterioAprobacion",
    "get_evaluador",
    "listar_guias",
    "contar_criterios",
]


# ─────────────────────────────────────────────────────────────────────────────
# CONSTANTE ESTÁNDAR DE CRITERIO
# ─────────────────────────────────────────────────────────────────────────────

_CRITERIO_STD: CriterioAprobacion = {
    "cumple": 0.6,
    "parcial": 0.3,
    "descripcion": (
        "60%+ de criterios presentes = CUMPLE | "
        "30-59% = PARCIAL | "
        "menos del 30% = NO CUMPLE"
    ),
}

# ─────────────────────────────────────────────────────────────────────────────
# HELPER INTERNO — normalización de claves para coincidencia tolerante
# ─────────────────────────────────────────────────────────────────────────────

def _normalizar_clave(s: str) -> str:
    """
    Normaliza una cadena de nombre de guía para comparación tolerante.

    Aplica en orden:
      1. Descomposición Unicode NFD (separa letra base de diacrítico).
      2. Elimina diacríticos (Mn = Mark, Nonspacing).
      3. Minúsculas.
      4. Reemplaza guiones bajos, guiones medios y espacios múltiples por un espacio.
      5. Strip de blancos extremos.

    Ejemplos:
      "Guía_01_Diagnóstico_Empresarial" -> "guia 01 diagnostico empresarial"
      "GUIA_01_DIAGNOSTICO_EMPRESARIAL" -> "guia 01 diagnostico empresarial"
    """
    s = unicodedata.normalize("NFD", s)
    s = "".join(c for c in s if unicodedata.category(c) != "Mn")
    s = s.lower().replace("_", " ").replace("-", " ")
    return " ".join(s.split())


# ─────────────────────────────────────────────────────────────────────────────
# EVALUADORES
# clave: nombre de la guía tal como aparece en actividades_parametrizadas.guia
# ─────────────────────────────────────────────────────────────────────────────

EVALUADORES: dict[str, EvaluadorGuia] = {

    # ══════════════════════════════════════════════════════════════════════════
    # GUÍA TRANSVERSAL — todos los programas
    # ══════════════════════════════════════════════════════════════════════════

    "GUIA_00_INDUCCION": {
        "guia": "GUIA_00_INDUCCION",
        "programa": "Asistencia_Comercial",
        "criterio_aprobacion": _CRITERIO_STD,
        "evidencias": [
            {
                "nombre": "QUIEN_SOY.png",
                "tipo": "imagen",
                "descripcion": "Nube de palabras con valores, cualidades y defectos del aprendiz",
                "criterios": [
                    "Es una imagen (nube de palabras, dibujo, collage o similar)",
                    "Contiene palabras relacionadas con valores, cualidades o defectos personales",
                    "Tiene mínimo 5 palabras o elementos visibles",
                    "El contenido hace referencia a características personales del aprendiz",
                ],
            },
            {
                "nombre": "MI_ESTILO_DE_APRENDIZAJE.xlsx",
                "tipo": "excel",
                "descripcion": "Test de David Kolb completado para identificar el estilo de aprendizaje",
                "criterios": [
                    "Es un archivo Excel o hoja de cálculo con datos diligenciados",
                    "Contiene puntuaciones, resultados o respuestas del test",
                    "Menciona estilos de aprendizaje o referencia el test de Kolb",
                    "No está completamente vacío; tiene al menos una sección completada",
                ],
            },
            {
                "nombre": "INFO_IDENTIDAD_SENA.pdf",
                "tipo": "documento",
                "descripcion": "Infografía sobre historia, misión, visión y símbolos institucionales del SENA",
                "criterios": [
                    "Menciona la historia o fundación del SENA",
                    "Menciona la misión y/o visión institucional",
                    "Incluye símbolos institucionales (escudo, bandera o himno)",
                    "Tiene formato visual de infografía (no solo texto corrido)",
                ],
            },
            {
                "nombre": "PLATAFORMAS_SENA.pdf",
                "tipo": "documento",
                "descripcion": "Documento con capturas y descripción de las plataformas digitales del SENA",
                "criterios": [
                    "Menciona al menos 2 plataformas (SofiaPlus, Zajuna, Betowa u otras)",
                    "Contiene capturas de pantalla o imágenes de las plataformas",
                    "Incluye descripciones o comentarios sobre cada plataforma explorada",
                    "Tiene mínimo 2 páginas de contenido desarrollado",
                ],
            },
            {
                "nombre": "PRESENTACION_PROGRAMA_FORMACION.pptx",
                "tipo": "presentacion",
                "descripcion": "Presentación con los puntos clave del programa de formación técnica",
                "criterios": [
                    "Es una presentación con diapositivas visibles",
                    "Menciona el nombre del programa de formación técnica",
                    "Incluye al menos 2 de: perfil de egreso, perfil ocupacional o proyección del egresado",
                    "Tiene contenido visual (imágenes, gráficos o diseño de diapositivas)",
                ],
            },
            {
                "nombre": "CHAT_REGLAMENTO_APRENDIZ.pdf",
                "tipo": "documento",
                "descripcion": "Chat generado con IA (ChatPDF, NotebookLM) sobre el reglamento del aprendiz SENA",
                "criterios": [
                    "Contiene preguntas y respuestas en formato conversacional (chat)",
                    "Menciona derechos o deberes del aprendiz SENA",
                    "Tiene mínimo 5 intercambios (pregunta–respuesta) completos",
                    "Hace referencia a temas del reglamento: inasistencia, deserción o llamados de atención",
                ],
            },
            {
                "nombre": "PROPUESTA_PROYECTO.pdf",
                "tipo": "documento",
                "descripcion": "Propuesta de proyecto productivo grupal con línea temática definida",
                "criterios": [
                    "Describe una idea o propuesta de proyecto productivo concreta",
                    "Menciona una línea temática o sector productivo específico",
                    "Hace referencia a trabajo grupal o menciona integrantes del equipo",
                    "Tiene mínimo una página de contenido desarrollado",
                ],
            },
            {
                "nombre": "LINEA_TIEMPO_PROFESIONAL.png",
                "tipo": "imagen",
                "descripcion": "Imagen de línea de tiempo con proyección profesional del aprendiz",
                "criterios": [
                    "Es una imagen que representa una línea de tiempo",
                    "Menciona años o fechas de proyección profesional futura",
                    "Incluye títulos, carreras o instituciones educativas como metas",
                    "Se proyecta más allá del nivel técnico (tecnológico, profesional o especialización)",
                ],
            },
        ],
    },

    # ══════════════════════════════════════════════════════════════════════════
    # PROGRAMA: Asistencia_Comercial / Técnico En Asesoría Comercial
    # ══════════════════════════════════════════════════════════════════════════

    "Guía_01_Diagnóstico_Empresarial": {
        "guia": "Guía_01_Diagnóstico_Empresarial",
        "programa": "Asistencia_Comercial",
        "criterio_aprobacion": _CRITERIO_STD,
        "evidencias": [
            {
                "nombre": "01_3_1_TALLER_JUGANDO_A_CREAR",
                "tipo": "documento",
                "descripcion": "Taller creativo sobre línea de juguetes o serie animada como punto de partida para el diagnóstico empresarial",
                "criterios": [
                    "Desarrolla un ejercicio creativo relacionado con producto, marca o empresa",
                    "Identifica al menos una característica diferenciadora del producto o idea",
                    "Presenta reflexión o análisis sobre el proceso creativo realizado",
                    "El documento está completo, no es solo portada o encabezado",
                ],
            },
            {
                "nombre": "01_3_2_EXPO_CASOS_EMPRENDIMIENTO_EN_COLOMBIA",
                "tipo": "documento",
                "descripcion": "Presentación de casos de éxito de start-ups colombianas analizados en equipo",
                "criterios": [
                    "Menciona al menos un caso real de emprendimiento colombiano (start-up o empresa)",
                    "Describe la problemática que resolvió el emprendimiento presentado",
                    "Identifica factores clave del éxito del caso analizado",
                    "Incluye reflexión del equipo o conclusión propia del aprendiz",
                ],
            },
            {
                "nombre": "01_3_3_1_TALLER_La_importancia_de_emprender",
                "tipo": "documento",
                "descripcion": "Taller reflexivo sobre economía naranja y emprendimiento (referencia: Yokoi Kenji)",
                "criterios": [
                    "Hace referencia al concepto de economía naranja o emprendimiento cultural/creativo",
                    "Menciona el ejemplo de Yokoi Kenji u otro caso de impacto económico por creatividad",
                    "Presenta argumentos sobre la importancia de emprender en Colombia",
                    "Tiene respuestas o reflexiones propias del aprendiz (no solo copia del material)",
                ],
            },
            {
                "nombre": "01_3_3_2_TALLER_Socialización_ideas_innovadoras",
                "tipo": "documento",
                "descripcion": "Taller de socialización de ideas innovadoras entre pares",
                "criterios": [
                    "Presenta una idea de negocio o producto innovador descrita con claridad",
                    "Identifica a qué público o necesidad responde la idea",
                    "Incluye retroalimentación recibida o reflexión sobre la socialización",
                    "El documento evidencia participación activa en la dinámica grupal",
                ],
            },
            {
                "nombre": "01_3_3_3_TALLER_Mentalidad_emprendedora",
                "tipo": "documento",
                "descripcion": "Taller de autodiagnóstico sobre actitudes y motivaciones emprendedoras personales",
                "criterios": [
                    "Identifica al menos 3 características o actitudes propias relacionadas con el emprendimiento",
                    "Relaciona esas características con perfiles emprendedores estudiados",
                    "Incluye una reflexión personal sobre sus fortalezas y áreas de mejora como emprendedor",
                    "Está desarrollado con respuestas propias, no con texto copiado del material",
                ],
            },
            {
                "nombre": "01_3_3_4_TALLER_Mapa_mental_conceptos",
                "tipo": "documento",
                "descripcion": "Mapa mental sobre estructura empresarial, tipos de empresa y conceptos de mercado",
                "criterios": [
                    "Es un mapa mental (visual, ramificado) o esquema gráfico estructurado",
                    "Incluye conceptos clave de estructura empresarial (misión, visión, organigrama o similar)",
                    "Relaciona al menos 3 tipos de empresa o formas jurídicas",
                    "Tiene jerarquía visual clara y no es solo texto lineal",
                ],
            },
            {
                "nombre": "01_3_3_5_TALLER_Caso_empresarial",
                "tipo": "documento",
                "descripcion": "Análisis de un caso empresarial real: misión, visión, estrategia y diagnóstico",
                "criterios": [
                    "Identifica y describe la empresa analizada con nombre y sector",
                    "Incluye misión, visión o valores de la empresa",
                    "Realiza un diagnóstico o análisis de al menos una variable (fortaleza, debilidad, oportunidad o amenaza)",
                    "Presenta conclusiones o recomendaciones basadas en el análisis",
                ],
            },
            {
                "nombre": "01_3_3_6_TALLER_Branding",
                "tipo": "documento",
                "descripcion": "Taller de branding y asesoría de marca: elementos identitarios de una empresa",
                "criterios": [
                    "Analiza o diseña elementos de identidad de marca (nombre, logo, colores, slogan)",
                    "Explica el concepto de branding y su importancia en la asesoría comercial",
                    "Asocia la marca con un público objetivo o segmento de mercado",
                    "Incluye propuesta visual o descripción detallada de la identidad de marca trabajada",
                ],
            },
            {
                "nombre": "01_3_3_7_Dinámica_cultura_emprendedora",
                "tipo": "imagen",
                "descripcion": "Evidencia de participación en dinámica digital sobre cultura emprendedora (Educaplay u otra plataforma)",
                "criterios": [
                    "Es una imagen o captura de pantalla de la actividad completada",
                    "Se puede identificar el nombre del aprendiz o su participación activa",
                    "Muestra resultados, puntaje o progreso en la actividad",
                    "La plataforma o actividad está relacionada con emprendimiento o diagnóstico empresarial",
                ],
            },
            {
                "nombre": "01_3_3_7_PRUEBA_DE_CONOCIMIENTOS",
                "tipo": "documento",
                "descripcion": "Prueba de conocimientos escrita o virtual sobre los temas de la guía",
                "criterios": [
                    "Contiene preguntas y respuestas sobre conceptos de diagnóstico empresarial",
                    "Tiene mínimo 5 ítems respondidos",
                    "Las respuestas están relacionadas con los temas del programa (emprendimiento, empresa, mercado)",
                    "No está en blanco ni solo tiene el encabezado",
                ],
            },
            {
                "nombre": "01_3_4_INFORME_Inicial_empresa_Nombre",
                "tipo": "documento",
                "descripcion": "Informe inicial de diagnóstico de la MIPYME seleccionada como proyecto del programa",
                "criterios": [
                    "Identifica la empresa o MIPYME seleccionada con nombre, sector y actividad económica",
                    "Describe el contexto del negocio (tamaño, años de operación, productos o servicios)",
                    "Incluye un diagnóstico preliminar con al menos 3 variables de análisis",
                    "Concluye con una justificación de por qué se eligió esa empresa para el proyecto",
                    "Tiene estructura de informe: introducción, desarrollo y conclusión",
                ],
            },
        ],
    },

    "Guía_02_Segmentación_cliente": {
        "guia": "Guía_02_Segmentación_cliente",
        "programa": "Asistencia_Comercial",
        "criterio_aprobacion": _CRITERIO_STD,
        "evidencias": [
            {
                "nombre": "02_3_1_TALLER_Gustos_y_Preferencias",
                "tipo": "documento",
                "descripcion": "Taller sobre gustos, preferencias y comportamientos de compra del aprendiz como consumidor",
                "criterios": [
                    "Describe gustos o preferencias de compra propios del aprendiz",
                    "Relaciona esas preferencias con variables de segmentación (edad, ingreso, estilo de vida u otra)",
                    "Reflexiona sobre factores que influyen en sus decisiones de compra",
                    "Tiene respuestas completas, no solo una o dos palabras por pregunta",
                ],
            },
            {
                "nombre": "02_3_2_ACTIVIDAD_Mapa_Conceptual_Movistar_Grooming",
                "tipo": "imagen",
                "descripcion": "Mapa conceptual sobre campaña publicitaria de Movistar y concepto de Grooming",
                "criterios": [
                    "Es un mapa conceptual (visual, con conectores o jerarquías)",
                    "Menciona la campaña de Movistar u otra campaña publicitaria analizada",
                    "Relaciona el concepto de Grooming u otro concepto clave del análisis",
                    "Tiene al menos 3 nodos o conceptos interrelacionados",
                ],
            },
            {
                "nombre": "02_3_3_1_Taller_Fuentes_de_información",
                "tipo": "imagen",
                "descripcion": "Evidencia de ejercicio sobre fuentes de información para segmentación de mercados",
                "criterios": [
                    "Es una imagen o captura que evidencia la actividad realizada",
                    "Identifica o clasifica fuentes de información (primarias, secundarias u otra clasificación)",
                    "Relaciona las fuentes con el proceso de segmentación de mercados",
                    "Se puede identificar que el ejercicio fue completado (no solo en blanco)",
                ],
            },
            {
                "nombre": "02_3_3_2_Ficha_de_estudio_con_registro_de_información",
                "tipo": "documento",
                "descripcion": "Ficha de estudio con síntesis de los conceptos de segmentación de mercados",
                "criterios": [
                    "Contiene definiciones o conceptos clave de segmentación de mercados",
                    "Está organizada como ficha de estudio (campos, secciones o cuadro resumen)",
                    "Menciona al menos 2 criterios de segmentación (demográfico, psicográfico, geográfico o conductual)",
                    "Refleja comprensión del tema con elaboración propia, no solo copia textual",
                ],
            },
            {
                "nombre": "02_3_3_3_Tipología_y_perfil_del_cliente_guión",
                "tipo": "cualquier",
                "descripcion": "Taller, guion y soporte visual sobre tipología y perfil del cliente",
                "criterios": [
                    "Describe al menos 2 tipos de cliente (por comportamiento, actitud de compra u otro criterio)",
                    "Construye o describe un perfil de cliente con características específicas",
                    "Incluye un guion o descripción de cómo tratar cada tipo de cliente",
                    "Tiene soporte visual (esquema, imagen o cuadro) además del texto",
                ],
            },
            {
                "nombre": "02_3_3_4_Prueba_de_conocimientos_guía_02",
                "tipo": "documento",
                "descripcion": "Prueba de conocimientos sobre segmentación de clientes",
                "criterios": [
                    "Contiene preguntas y respuestas sobre segmentación de mercados",
                    "Tiene mínimo 5 ítems respondidos",
                    "Las respuestas se relacionan con criterios de segmentación y perfil del cliente",
                    "No está vacía ni presenta solo el encabezado",
                ],
            },
            {
                "nombre": "02_3_4_TALLER_INFORME_Segmentación_de_Clientes",
                "tipo": "presentacion",
                "descripcion": "Informe de segmentación de los clientes de la MIPYME del proyecto",
                "criterios": [
                    "Identifica el segmento objetivo de la MIPYME con criterios definidos",
                    "Aplica al menos 2 variables de segmentación al caso real de la empresa",
                    "Presenta un perfil de cliente ideal o buyer persona para la MIPYME",
                    "Incluye conclusiones o recomendaciones de estrategia basadas en la segmentación",
                    "Tiene formato de informe o presentación con estructura clara",
                ],
            },
        ],
    },

    "Guía_03_Prospección_de_clientes": {
        "guia": "Guía_03_Prospección_de_clientes",
        "programa": "Asistencia_Comercial",
        "criterio_aprobacion": _CRITERIO_STD,
        "evidencias": [
            {
                "nombre": "03_3_1_Taller_lista_de_regalos",
                "tipo": "documento",
                "descripcion": "Taller sobre dinámicas comerciales usando metáfora de lista de regalos para entender necesidades del cliente",
                "criterios": [
                    "Desarrolla la dinámica de lista de regalos u otra actividad de identificación de necesidades",
                    "Relaciona las necesidades identificadas con el proceso de prospección comercial",
                    "Identifica al menos 3 características de un cliente potencial a partir del ejercicio",
                    "Tiene respuestas completas y reflexivas, no solo palabras sueltas",
                ],
            },
            {
                "nombre": "03_3_2_Taller_Centro_Comercial",
                "tipo": "documento",
                "descripcion": "Taller de observación y análisis de comportamientos de clientes en entornos comerciales",
                "criterios": [
                    "Describe la observación realizada en un entorno comercial real o simulado",
                    "Identifica comportamientos de compra o necesidades de los clientes observados",
                    "Relaciona la observación con técnicas de prospección comercial",
                    "Presenta conclusiones sobre el perfil del cliente en ese entorno",
                ],
            },
            {
                "nombre": "03_3_3_1_Elaboración_Buyer_Persona",
                "tipo": "documento",
                "descripcion": "Construcción del perfil Buyer Persona para la MIPYME del proyecto",
                "criterios": [
                    "Define un Buyer Persona con nombre, edad, ocupación y características demográficas",
                    "Describe motivaciones, metas y frustraciones del perfil construido",
                    "Relaciona el Buyer Persona con el producto o servicio de la MIPYME del proyecto",
                    "Incluye al menos 4 campos del perfil (no solo nombre y edad)",
                ],
            },
            {
                "nombre": "03_3_3_2_Torneo_de_aprendizaje",
                "tipo": "documento",
                "descripcion": "Material didáctico digital preparado para exposición en torneo de aprendizaje",
                "criterios": [
                    "Contiene contenido sobre prospección de clientes o técnicas de venta",
                    "Está diseñado para ser presentado o compartido con otros (diapositivas, quiz, infografía)",
                    "Tiene al menos 3 conceptos clave desarrollados con claridad",
                    "Presenta el tema de manera organizada y visualmente comprensible",
                ],
            },
            {
                "nombre": "03_3_3_3_Taller_conociendo_el_mundo",
                "tipo": "documento",
                "descripcion": "Diseño de encuesta digital y taller sobre bases de datos de clientes potenciales",
                "criterios": [
                    "Incluye un diseño de encuesta con mínimo 5 preguntas relevantes para prospección",
                    "Las preguntas están orientadas a identificar necesidades o perfiles de clientes",
                    "Menciona herramientas digitales para diseñar encuestas (Google Forms, Typeform u otras)",
                    "Relaciona el ejercicio con la construcción de una base de datos de prospectos",
                ],
            },
            {
                "nombre": "03_3_3_4_Taller_ejercicio_de_prospección",
                "tipo": "documento",
                "descripcion": "Taller práctico de ejercicios de prospección aplicados al contexto comercial",
                "criterios": [
                    "Aplica al menos una técnica de prospección (llamada en frío, referidos, redes sociales u otra)",
                    "Describe el proceso seguido para identificar clientes potenciales",
                    "Presenta resultados o evidencia del ejercicio de prospección (lista, registro o simulación)",
                    "Reflexiona sobre la efectividad de la técnica utilizada",
                ],
            },
            {
                "nombre": "03_3_3_5_Prueba_de_conocimientos_Guía_03",
                "tipo": "documento",
                "descripcion": "Prueba de conocimientos sobre prospección de clientes",
                "criterios": [
                    "Contiene preguntas sobre conceptos de prospección comercial",
                    "Tiene mínimo 5 ítems respondidos",
                    "Las respuestas se relacionan con técnicas de prospección, bases de datos o perfil de cliente",
                    "No está en blanco ni presenta solo el encabezado",
                ],
            },
            {
                "nombre": "03_3_4_Informe_Final_de_Prospección",
                "tipo": "documento",
                "descripcion": "Informe final de prospección de clientes para la empresa del proyecto",
                "criterios": [
                    "Identifica y describe los clientes potenciales de la MIPYME con criterios claros",
                    "Presenta la base de datos o lista de prospectos construida",
                    "Describe las técnicas de prospección aplicadas a la empresa real",
                    "Incluye un análisis del mercado objetivo y recomendaciones de abordaje comercial",
                    "Tiene estructura de informe con introducción, desarrollo y conclusión",
                ],
            },
        ],
    },

    "Guía_04_Portafolio_de_productos_y_servicios": {
        "guia": "Guía_04_Portafolio_de_productos_y_servicios",
        "programa": "Asistencia_Comercial",
        "criterio_aprobacion": _CRITERIO_STD,
        "evidencias": [
            {
                "nombre": "04_3_1_Taller_Estructura_Pagina_Amazon",
                "tipo": "documento",
                "descripcion": "Análisis de la estructura de página de producto en Amazon y plataformas similares",
                "criterios": [
                    "Analiza la estructura de una página de producto en Amazon u otra plataforma de e-commerce",
                    "Identifica al menos 4 elementos de la ficha de producto (título, descripción, imágenes, precio, reseñas)",
                    "Relaciona esos elementos con la presentación de productos en asesoría comercial",
                    "Incluye capturas de pantalla o descripción detallada de la página analizada",
                ],
            },
            {
                "nombre": "04_3_2_TALLER_Identificación_Beneficios_Productos_y_Servicios",
                "tipo": "documento",
                "descripcion": "Taller de identificación de características, beneficios y atributos de productos y servicios",
                "criterios": [
                    "Diferencia entre características y beneficios de un producto o servicio",
                    "Aplica la distinción característica–beneficio a al menos 2 productos o servicios concretos",
                    "Identifica el beneficio principal que busca el cliente al adquirir cada producto",
                    "Presenta la información en tabla, cuadro comparativo o estructura organizada",
                ],
            },
            {
                "nombre": "04_3_3_1_TALLER_Fotografía_Producto_Fichas_Envase_Empaque",
                "tipo": "documento",
                "descripcion": "Taller sobre fotografía de producto, ficha técnica, envase y empaque como herramientas de venta",
                "criterios": [
                    "Incluye fotografías propias o analizadas de un producto con criterios de calidad",
                    "Elabora o analiza una ficha técnica con especificaciones del producto",
                    "Describe el envase y empaque como elemento de diferenciación y comunicación",
                    "Relaciona estos elementos con la presentación comercial del producto a los clientes",
                ],
            },
            {
                "nombre": "04_3_3_2_TALLER_Marketing_Mix_Costos_y_Precios",
                "tipo": "documento",
                "descripcion": "Taller sobre las 4P del marketing mix con énfasis en costos y estrategia de precios",
                "criterios": [
                    "Describe las 4P del marketing mix (Producto, Precio, Plaza, Promoción) aplicadas a un ejemplo",
                    "Calcula o analiza costos básicos del producto (costo de producción o adquisición)",
                    "Explica al menos 2 estrategias de fijación de precios",
                    "Aplica el análisis de marketing mix a la MIPYME del proyecto",
                ],
            },
            {
                "nombre": "04_3_3_3_TALLER_Portafolio_de_Productos_y_Servicios_y_Catálogo",
                "tipo": "documento",
                "descripcion": "Taller de diseño de portafolio de productos y catálogo comercial",
                "criterios": [
                    "Contiene un portafolio con al menos 3 productos o servicios descritos",
                    "Cada producto incluye nombre, descripción, precio o rango de precio",
                    "Tiene diseño visual de catálogo (imágenes, diagramación o presentación atractiva)",
                    "Está orientado a facilitar la presentación comercial ante un cliente",
                ],
            },
            {
                "nombre": "04_3_3_4_Prueba_de_conocimientos",
                "tipo": "documento",
                "descripcion": "Prueba de conocimientos sobre portafolio de productos y servicios",
                "criterios": [
                    "Contiene preguntas sobre conceptos de portafolio, marketing mix o fichas de producto",
                    "Tiene mínimo 5 ítems respondidos",
                    "Las respuestas demuestran comprensión de los temas de la guía",
                    "No está en blanco ni presenta solo el encabezado",
                ],
            },
            {
                "nombre": "04_3_4_INFORME_Portafolio_de_Productos_o_Servicios",
                "tipo": "documento",
                "descripcion": "Informe final del portafolio de productos o servicios de la MIPYME del proyecto",
                "criterios": [
                    "Presenta el portafolio completo de la MIPYME con todos sus productos o servicios",
                    "Incluye ficha técnica o descripción detallada de cada producto/servicio",
                    "Analiza el posicionamiento y estrategia de precios de la empresa",
                    "Concluye con recomendaciones para mejorar la presentación del portafolio al cliente",
                    "Tiene estructura de informe con secciones claramente definidas",
                ],
            },
        ],
    },

    "GUIA_05_Herramientas_Ofimáticas_de_Mercadeo": {
        "guia": "GUIA_05_Herramientas_Ofimáticas_de_Mercadeo",
        "programa": "Asistencia_Comercial",
        "criterio_aprobacion": _CRITERIO_STD,
        "evidencias": [
            {
                "nombre": "05_1_ACTIVIDAD_NUEVOS_INFLUENCERS",
                "tipo": "cualquier",
                "descripcion": "Actividad sobre el rol de los nuevos influencers en estrategias de mercadeo digital",
                "criterios": [
                    "Identifica al menos 2 tipos de influencers (macro, micro, nano u otra clasificación)",
                    "Analiza el impacto de los influencers en una estrategia de mercadeo",
                    "Relaciona el fenómeno de los influencers con la asesoría comercial o ventas",
                    "Presenta un ejemplo real o hipotético de uso de influencers en una campaña",
                ],
            },
            {
                "nombre": "05_2_DINAMICA_SALIENDO_DEL_LABERINTO",
                "tipo": "cualquier",
                "descripcion": "Evidencia de participación en la dinámica 'Saliendo del Laberinto' sobre herramientas de mercadeo",
                "criterios": [
                    "Evidencia participación en la dinámica descrita (captura, documento o resultado)",
                    "Contiene respuestas o resolución de los retos planteados en la actividad",
                    "Relaciona el ejercicio con conceptos de mercadeo o herramientas ofimáticas",
                    "El documento está completo y no es solo el enunciado sin resolver",
                ],
            },
            {
                "nombre": "INFOGRAFIA_PROPUESTA_COMERCIAL",
                "tipo": "imagen",
                "descripcion": "Infografía visual sobre los elementos de una propuesta comercial efectiva",
                "criterios": [
                    "Es una infografía con elementos visuales (iconos, colores, diagramas o ilustraciones)",
                    "Presenta los componentes clave de una propuesta comercial",
                    "Está diseñada con herramienta ofimática o digital (Canva, PowerPoint, Word u otra)",
                    "Tiene texto legible y organizado, no solo imágenes sin explicación",
                ],
            },
            {
                "nombre": "INFOGRAFIA_PLAN_DE_VENTAS",
                "tipo": "imagen",
                "descripcion": "Infografía sobre los elementos y estructura de un plan de ventas",
                "criterios": [
                    "Es una infografía con diseño visual sobre plan de ventas",
                    "Incluye al menos 4 elementos del plan de ventas (objetivos, estrategias, metas, indicadores u otros)",
                    "Está diseñada con herramienta ofimática o digital",
                    "Es comprensible de forma autónoma (no requiere explicación externa para entenderse)",
                ],
            },
            {
                "nombre": "05_3_3_2_TALLER_PROPUESTA_HERRAMIENTAS_DE_MERCADEO",
                "tipo": "documento",
                "descripcion": "Taller sobre el uso de herramientas ofimáticas aplicadas al mercadeo y la asesoría comercial",
                "criterios": [
                    "Usa al menos 2 herramientas ofimáticas (Excel, Word, PowerPoint, Canva u otras)",
                    "Aplica esas herramientas a tareas concretas de mercadeo o asesoría comercial",
                    "Describe cómo cada herramienta facilita el trabajo comercial",
                    "Presenta resultados de la aplicación (tablas, gráficos, documentos o diseños)",
                ],
            },
            {
                "nombre": "05_3_3_3_PRUEBA_DE_CONOCIMIENTOS",
                "tipo": "cualquier",
                "descripcion": "Prueba de conocimientos sobre herramientas ofimáticas de mercadeo",
                "criterios": [
                    "Contiene preguntas sobre herramientas ofimáticas y su aplicación en mercadeo",
                    "Tiene mínimo 5 ítems respondidos",
                    "Las respuestas reflejan comprensión de los temas de la guía",
                    "No está en blanco ni solo tiene el encabezado",
                ],
            },
            {
                "nombre": "05_3_4_INFORME_PROPUESTA_DE_MERCADEO",
                "tipo": "cualquier",
                "descripcion": "Informe de propuesta de mercadeo para la MIPYME del proyecto usando herramientas ofimáticas",
                "criterios": [
                    "Presenta una propuesta de mercadeo completa para la empresa del proyecto",
                    "Incluye análisis de situación actual y objetivos de la propuesta",
                    "Usa herramientas ofimáticas para presentar la información (tablas, gráficos, infografías)",
                    "Concluye con un plan de acción o estrategias concretas de mercadeo",
                    "Tiene estructura de informe formal con secciones identificables",
                ],
            },
        ],
    },

    "GUIA_06_Surtido_exhibición_y_merchandising": {
        "guia": "GUIA_06_Surtido_exhibición_y_merchandising",
        "programa": "Asistencia_Comercial",
        "criterio_aprobacion": _CRITERIO_STD,
        "evidencias": [
            {
                "nombre": "06_3_1_EXPO_INDEPENDENCIA_Y_ESTILO",
                "tipo": "documento",
                "descripcion": "Exposición grupal sobre independencia de criterio y estilo en la asesoría comercial y el surtido",
                "criterios": [
                    "Presenta contenido sobre independencia de criterio en selección de productos o surtido",
                    "Relaciona el concepto de estilo con las preferencias del consumidor o las tendencias del mercado",
                    "Tiene evidencia de presentación grupal (guion, diapositivas o registro de la exposición)",
                    "Incluye reflexión sobre la importancia del surtido bien curado en puntos de venta",
                ],
            },
            {
                "nombre": "06_3_2_ACTIVIDAD_MULTIVERSO_RICK_AND_MORTY",
                "tipo": "documento",
                "descripcion": "Actividad creativa usando la metáfora del multiverso para analizar segmentos de mercado y surtido",
                "criterios": [
                    "Desarrolla el ejercicio usando la dinámica propuesta (multiverso, universos alternativos u otra metáfora)",
                    "Aplica la dinámica para analizar diferentes segmentos de clientes o líneas de surtido",
                    "Identifica al menos 3 variables de surtido o exhibición en el ejercicio",
                    "Presenta conclusiones relacionadas con merchandising o surtido",
                ],
            },
            {
                "nombre": "06_3_3_1_TALLER_SURTIDO",
                "tipo": "documento",
                "descripcion": "Taller sobre criterios de surtido: amplitud, profundidad y coherencia de la oferta de productos",
                "criterios": [
                    "Define los conceptos de amplitud y profundidad del surtido",
                    "Analiza el surtido de una empresa o tienda real con esos criterios",
                    "Identifica oportunidades de mejora en el surtido analizado",
                    "Aplica al menos un criterio de surtido a la MIPYME del proyecto",
                ],
            },
            {
                "nombre": "06_3_3_2_INFOGRAFIA_SUPERMERCADOS",
                "tipo": "imagen",
                "descripcion": "Infografía sobre técnicas de surtido y exhibición en supermercados y grandes superficies",
                "criterios": [
                    "Es una infografía con diseño visual sobre surtido o exhibición en supermercados",
                    "Menciona al menos 3 técnicas o estrategias de surtido en grandes superficies",
                    "Tiene iconos, colores o imágenes que refuerzan el contenido",
                    "Es comprensible como pieza autónoma de información",
                ],
            },
            {
                "nombre": "06_3_3_3_EXPO_TRABAJO_DE_CAMPO",
                "tipo": "presentacion",
                "descripcion": "Presentación de los resultados del trabajo de campo en punto de venta real",
                "criterios": [
                    "Presenta evidencia del trabajo de campo realizado (fotos, registros o diario de observación)",
                    "Describe el punto de venta visitado con sus características de exhibición y surtido",
                    "Analiza al menos 2 aspectos de merchandising observados in situ",
                    "Concluye con aprendizajes o recomendaciones basadas en la visita",
                ],
            },
            {
                "nombre": "06_3_3_4_PRUEBA_SURTIDO_EXHIBICION_Y_MERCHANDISING",
                "tipo": "cualquier",
                "descripcion": "Prueba de conocimientos sobre surtido, exhibición y merchandising",
                "criterios": [
                    "Contiene preguntas sobre conceptos de surtido, exhibición o merchandising",
                    "Tiene mínimo 5 ítems respondidos",
                    "Las respuestas demuestran comprensión de los temas de la guía",
                    "No está en blanco ni presenta solo el encabezado",
                ],
            },
            {
                "nombre": "06_3_4_INFORME_SURTIDO_EXHIBICION_MERCHANDISING",
                "tipo": "documento",
                "descripcion": "Informe final de surtido, exhibición y merchandising aplicado a la MIPYME del proyecto",
                "criterios": [
                    "Analiza el estado actual del surtido y exhibición de la empresa del proyecto",
                    "Propone mejoras concretas de merchandising para el punto de venta",
                    "Incluye referencias a técnicas de surtido (planograma, layout, zonificación u otras)",
                    "Presenta recomendaciones de implementación con justificación comercial",
                    "Tiene estructura de informe con secciones identificables",
                ],
            },
        ],
    },

    "Guía 07 _Negociaciones_de_venta": {
        "guia": "Guía 07 _Negociaciones_de_venta",
        "programa": "Asistencia_Comercial",
        "criterio_aprobacion": _CRITERIO_STD,
        "evidencias": [
            {
                "nombre": "07_1_1_ACTIVIDAD_APRENDIENDO_A_NEGOCIAR",
                "tipo": "cualquier",
                "descripcion": "Actividad introductoria sobre principios y estilos de negociación",
                "criterios": [
                    "Identifica al menos 2 estilos de negociación (colaborativo, competitivo, comprometido u otros)",
                    "Describe características del negociador exitoso en contexto comercial",
                    "Relaciona los principios de negociación con la asesoría comercial",
                    "El documento presenta respuestas desarrolladas, no solo palabras clave",
                ],
            },
            {
                "nombre": "07_2_1_TALLER_MUNDO_ALTERNO",
                "tipo": "documento",
                "descripcion": "Taller de negociación usando simulaciones o escenarios alternativos",
                "criterios": [
                    "Desarrolla un escenario o simulación de negociación con roles definidos",
                    "Aplica técnicas de negociación al escenario planteado",
                    "Presenta el resultado o acuerdo alcanzado en la simulación",
                    "Reflexiona sobre las estrategias usadas y su efectividad",
                ],
            },
            {
                "nombre": "07_3_1_TALLER_COMUNICACION_ASERTIVA",
                "tipo": "documento",
                "descripcion": "Taller sobre comunicación asertiva como herramienta clave en la negociación comercial",
                "criterios": [
                    "Define comunicación asertiva y la diferencia de agresiva o pasiva",
                    "Presenta ejemplos de comunicación asertiva en contexto de ventas o negociación",
                    "Aplica técnicas de comunicación asertiva a un caso o diálogo simulado",
                    "Incluye reflexión sobre cómo la asertividad mejora los resultados en negociación",
                ],
            },
            {
                "nombre": "07_3_2_TALLER_INTEGRAL_DE_VENTAS",
                "tipo": "documento",
                "descripcion": "Taller integral que integra técnicas de venta, manejo de objeciones y cierre",
                "criterios": [
                    "Describe el proceso de venta con sus etapas (prospección, presentación, manejo de objeciones, cierre)",
                    "Presenta técnicas de manejo de objeciones con ejemplos concretos",
                    "Muestra técnicas de cierre de venta aplicadas a casos del sector comercial",
                    "Integra la negociación como parte del proceso de venta completo",
                ],
            },
            {
                "nombre": "07_3_3_DINAMICA_NEGOCIACION_EN_VENTAS",
                "tipo": "imagen",
                "descripcion": "Evidencia fotográfica o captura de participación en dinámica de negociación en ventas",
                "criterios": [
                    "Es una imagen o captura que evidencia la participación en la dinámica",
                    "Se identifica el contexto de la actividad de negociación",
                    "Muestra resultados, puntaje o progreso de la dinámica si aplica",
                    "La evidencia está relacionada con temas de negociación en ventas",
                ],
            },
            {
                "nombre": "07_3_3_TALLER_RUEDA_DE_NEGOCIOS",
                "tipo": "documento",
                "descripcion": "Taller sobre rueda de negocios como herramienta de networking y negociación empresarial",
                "criterios": [
                    "Explica el concepto de rueda de negocios y su propósito comercial",
                    "Describe cómo prepararse para una rueda de negocios (propuesta de valor, presentación)",
                    "Presenta un pitch o presentación de negocio simulada para la rueda",
                    "Reflexiona sobre los aprendizajes del ejercicio de negociación",
                ],
            },
            {
                "nombre": "07_3_PRUEBA_NEGOCIACIONES_EN_LA_VENTA",
                "tipo": "cualquier",
                "descripcion": "Prueba de conocimientos sobre técnicas de negociación en ventas",
                "criterios": [
                    "Contiene preguntas sobre conceptos de negociación y técnicas de venta",
                    "Tiene mínimo 5 ítems respondidos",
                    "Las respuestas reflejan comprensión de estilos de negociación y proceso de venta",
                    "No está en blanco ni presenta solo el encabezado",
                ],
            },
            {
                "nombre": "07_4_VIDEO_REPRESENTACION_EJECUCION_PROCESO_DE_VENTAS",
                "tipo": "cualquier",
                "descripcion": "Video o representación del proceso de ventas completo desde prospección hasta cierre",
                "criterios": [
                    "Es un video o evidencia de representación dramática del proceso de ventas",
                    "Cubre al menos 3 etapas del proceso de ventas (prospección, presentación, objeciones o cierre)",
                    "Muestra aplicación de técnicas de venta y negociación estudiadas",
                    "Dura mínimo 2 minutos o equivale en contenido a una representación completa",
                ],
            },
        ],
    },

    # ══════════════════════════════════════════════════════════════════════════
    # PROGRAMA: Comunicacion_y_marketing
    # ══════════════════════════════════════════════════════════════════════════

    "Guía_01_Clasificar_clientes": {
        "guia": "Guía_01_Clasificar_clientes",
        "programa": "Comunicacion_y_marketing",
        "criterio_aprobacion": _CRITERIO_STD,
        "evidencias": [
            {
                "nombre": "01_3_1_GLOSARIO_TERMINOS_MARKETING",
                "tipo": "documento",
                "descripcion": "Glosario de términos básicos de marketing elaborado por el aprendiz",
                "criterios": [
                    "Contiene definiciones de términos de marketing (mercado, cliente, segmento, etc.)",
                    "Tiene mínimo 8 términos definidos",
                    "Las definiciones están redactadas con palabras propias",
                    "El glosario está organizado (alfabético, temático o con formato de tabla)",
                ],
            },
            {
                "nombre": "01_3_2_VIDEO_ORGANIZADOR_GRAFICO_MERCADOS",
                "tipo": "cualquier",
                "descripcion": "Video y/o organizador gráfico sobre tipos de mercados",
                "criterios": [
                    "Es un video, presentación o esquema sobre mercados",
                    "Identifica al menos 3 tipos de mercado (consumidor, industrial, internacional u otros)",
                    "Tiene contenido desarrollado, no solo el enunciado",
                    "El organizador gráfico muestra relaciones entre conceptos",
                ],
            },
            {
                "nombre": "01_3_3_STORYBOARD_PERFIL_CLIENTE",
                "tipo": "imagen",
                "descripcion": "Storyboard o representación visual del perfil del cliente segmentado",
                "criterios": [
                    "Es una imagen o representación visual con secuencia narrativa",
                    "Describe un perfil de cliente con características demográficas y/o psicográficas",
                    "Relaciona el perfil con un producto o servicio específico",
                    "Tiene al menos 4 elementos o cuadros de secuencia visibles",
                ],
            },
            {
                "nombre": "01_3_4_INFORME_PSICOLOGIA_CONSUMIDOR",
                "tipo": "presentacion",
                "descripcion": "Informe o presentación sobre psicología del consumidor y factores de decisión de compra",
                "criterios": [
                    "Es una presentación o informe sobre psicología del consumidor",
                    "Identifica al menos 3 factores que influyen en la decisión de compra",
                    "Relaciona los factores psicológicos con estrategias de marketing",
                    "Tiene estructura lógica con introducción, desarrollo y conclusión",
                ],
            },
            {
                "nombre": "01_3_5_ACTIVIDAD_FUENTES_INFORMACION",
                "tipo": "cualquier",
                "descripcion": "Actividad o juego sobre fuentes de información de mercado (opcional)",
                "criterios": [
                    "Evidencia participación en la actividad de fuentes de información",
                    "Identifica al menos 2 tipos de fuentes (primarias y secundarias)",
                    "Relaciona las fuentes con la investigación de mercados",
                    "El documento no está en blanco",
                ],
            },
            {
                "nombre": "01_3_6_PRUEBA_EVIDENCIA_CONOCIMIENTO",
                "tipo": "cualquier",
                "descripcion": "Prueba de evidencia de conocimiento sobre clasificación de clientes y mercados",
                "criterios": [
                    "Contiene preguntas y respuestas sobre clasificación de clientes o mercados",
                    "Tiene mínimo 5 ítems respondidos",
                    "Las respuestas reflejan comprensión de los temas de la guía",
                    "No está en blanco ni presenta solo el encabezado",
                ],
            },
            {
                "nombre": "01_4_BASE_DATOS_CLIENTES_PROSPECTOS",
                "tipo": "documento",
                "descripcion": "Base de datos de clientes prospectos con informes del proyecto",
                "criterios": [
                    "Contiene una lista o base de datos con información de clientes potenciales",
                    "Incluye al menos 5 registros con datos de contacto o perfil del cliente",
                    "Clasifica los prospectos según algún criterio de segmentación",
                    "Acompaña la base con un análisis o informe descriptivo",
                    "Tiene estructura de tabla o base de datos ordenada",
                ],
            },
        ],
    },

    "GUIA_02_PREPARAR_EXHIBICION": {
        "guia": "GUIA_02_PREPARAR_EXHIBICION",
        "programa": "Comunicacion_y_marketing",
        "criterio_aprobacion": _CRITERIO_STD,
        "evidencias": [
            {
                "nombre": "02_3_1_MAPA_MENTAL_MERCADEO",
                "tipo": "imagen",
                "descripcion": "Mapa mental sobre conceptos de mercadeo y merchandising",
                "criterios": [
                    "Es un mapa mental con estructura radial o ramificada",
                    "Incluye conceptos de mercadeo (merchandising, exhibición, vitrina u otros)",
                    "Tiene al menos 4 ramas o conceptos interrelacionados",
                    "Tiene jerarquía visual clara con nodo central y subnodos",
                ],
            },
            {
                "nombre": "02_3_3_PRESENTACION_FORMATOS_COMERCIALES",
                "tipo": "presentacion",
                "descripcion": "Presentación sobre formatos comerciales (tiendas, supermercados, e-commerce, etc.)",
                "criterios": [
                    "Es una presentación con diapositivas sobre formatos comerciales",
                    "Describe al menos 3 formatos comerciales con sus características",
                    "Relaciona los formatos con estrategias de exhibición y merchandising",
                    "Tiene contenido visual (imágenes, gráficos o diseño de diapositivas)",
                ],
            },
            {
                "nombre": "02_3_4_CUADRO_COMPARATIVO_MERCHANDISING",
                "tipo": "documento",
                "descripcion": "Cuadro comparativo de técnicas y tipos de merchandising",
                "criterios": [
                    "Presenta un cuadro comparativo con estructura de tabla",
                    "Compara al menos 2 tipos o técnicas de merchandising",
                    "Identifica características, ventajas o aplicaciones de cada técnica",
                    "El cuadro está completo, no solo tiene encabezados sin contenido",
                ],
            },
            {
                "nombre": "02_3_5_HERRAMIENTA_GRAFICA_EXHIBICION",
                "tipo": "imagen",
                "descripcion": "Herramienta gráfica o infografía sobre técnicas de exhibición en punto de venta",
                "criterios": [
                    "Es una imagen, infografía o esquema visual sobre exhibición",
                    "Presenta técnicas, principios o elementos de una exhibición efectiva",
                    "Tiene elementos visuales (iconos, colores, fotografías o diagramas)",
                    "Es comprensible como pieza autónoma de información",
                ],
            },
            {
                "nombre": "02_3_6_JUEGO_ESCAPARATISMO",
                "tipo": "cualquier",
                "descripcion": "Evidencia de juego digital sobre escaparatismo (enlace, captura o resultado)",
                "criterios": [
                    "Evidencia participación en un juego o actividad interactiva sobre escaparatismo",
                    "Muestra resultado, puntaje o avance en la actividad",
                    "Relaciona la actividad con conceptos de exhibición o diseño de vitrinas",
                    "El documento o captura no está vacío",
                ],
            },
            {
                "nombre": "02_3_7_PRESENTACION_SOCIALIZACION",
                "tipo": "presentacion",
                "descripcion": "Presentación de socialización del plan de exhibición propuesto",
                "criterios": [
                    "Es una presentación con propuesta de plan de exhibición",
                    "Describe el layout o distribución propuesta para un punto de venta",
                    "Justifica las decisiones de exhibición con criterios de merchandising",
                    "Tiene diseño visual adecuado para presentar ante una audiencia",
                ],
            },
            {
                "nombre": "02_3_8_ACTIVIDAD_MATERIAL_POP",
                "tipo": "documento",
                "descripcion": "Análisis e investigación sobre material POP (Point of Purchase) — opcional",
                "criterios": [
                    "Define el concepto de material POP y sus tipos",
                    "Presenta ejemplos reales de material POP de distintas marcas",
                    "Analiza la efectividad del material POP en el punto de venta",
                    "Tiene contenido desarrollado, no solo el enunciado",
                ],
            },
            {
                "nombre": "02_4_INFORME_LAYOUT_PUNTO_DE_VENTA",
                "tipo": "documento",
                "descripcion": "Informe de layout y diseño del punto de venta para el proyecto",
                "criterios": [
                    "Presenta un diseño o propuesta de layout para el punto de venta del proyecto",
                    "Incluye plano, esquema o descripción detallada de la distribución",
                    "Justifica las decisiones con criterios de tráfico, zonas calientes y frías",
                    "Propone estrategias de exhibición alineadas con el perfil del cliente",
                    "Tiene estructura de informe con secciones identificables",
                ],
            },
        ],
    },

    "GUIA_03_PUBLICAR_CONTENIDOS_ANUNCIOS": {
        "guia": "GUIA_03_PUBLICAR_CONTENIDOS_ANUNCIOS",
        "programa": "Comunicacion_y_marketing",
        "criterio_aprobacion": _CRITERIO_STD,
        "evidencias": [
            {
                "nombre": "03_3_1_DOCUMENTO_MAPA_CONCEPTUAL_MARKETING_DIGITAL",
                "tipo": "documento",
                "descripcion": "Documento y mapa conceptual sobre marketing digital y sus componentes",
                "criterios": [
                    "Incluye un mapa conceptual sobre marketing digital con relaciones entre conceptos",
                    "Define marketing digital y lo diferencia del marketing tradicional",
                    "Menciona al menos 3 componentes o canales del marketing digital",
                    "Combina texto explicativo y representación visual",
                ],
            },
            {
                "nombre": "03_3_2_TABLA_PRESENTACION_PLATAFORMAS_DIGITALES",
                "tipo": "documento",
                "descripcion": "Tabla comparativa y presentación de las principales plataformas digitales",
                "criterios": [
                    "Presenta una tabla comparativa de plataformas digitales (Instagram, Facebook, TikTok u otras)",
                    "Compara al menos 3 plataformas con criterios definidos",
                    "Identifica cuál plataforma es más adecuada para distintos objetivos",
                    "Tiene estructura de tabla o cuadro comparativo completo",
                ],
            },
            {
                "nombre": "03_3_3_PLAN_CONTENIDO_PIEZA_DIGITAL",
                "tipo": "cualquier",
                "descripcion": "Plan de contenidos con pieza de contenido digital creada",
                "criterios": [
                    "Presenta un plan de contenidos con calendario o frecuencia de publicación",
                    "Incluye al menos 4 piezas o temáticas de contenido planificadas",
                    "Adjunta al menos una pieza de contenido digital creada",
                    "El plan está alineado con un público objetivo definido",
                ],
            },
            {
                "nombre": "03_4_1_DOCUMENTO_ESTRATEGIA_CAMPAÑA",
                "tipo": "documento",
                "descripcion": "Documento de estrategia de campaña digital completa",
                "criterios": [
                    "Define el objetivo de la campaña digital con claridad",
                    "Describe el público objetivo y segmentación para la campaña",
                    "Presenta el presupuesto estimado o la distribución de recursos",
                    "Incluye cronograma o plan de ejecución de la campaña",
                ],
            },
            {
                "nombre": "03_4_2_CONTENIDOS_ANUNCIOS_DIGITALES",
                "tipo": "cualquier",
                "descripcion": "Contenidos y anuncios digitales creados para la campaña",
                "criterios": [
                    "Presenta al menos 2 piezas de contenido o anuncios digitales creados",
                    "Las piezas tienen un objetivo claro de comunicación comercial",
                    "El diseño y mensaje están alineados con la estrategia de campaña",
                    "Los formatos son adecuados para la plataforma digital seleccionada",
                ],
            },
            {
                "nombre": "03_4_3_PARRILLA_CONFIGURACION",
                "tipo": "imagen",
                "descripcion": "Parrilla de contenidos y evidencia de configuración de anuncios",
                "criterios": [
                    "Presenta una parrilla o calendario de contenidos en formato visual",
                    "La parrilla cubre al menos 1 semana de publicaciones planificadas",
                    "Incluye captura o evidencia de la configuración del anuncio en la plataforma",
                    "Organiza los contenidos por fecha, temática o formato",
                ],
            },
            {
                "nombre": "03_4_4_PUBLICACION_PROGRAMACION",
                "tipo": "cualquier",
                "descripcion": "Evidencia de publicación o programación de contenidos en plataformas digitales",
                "criterios": [
                    "Muestra evidencia real de publicación (enlace, captura de pantalla o comprobante)",
                    "El contenido está publicado o programado en al menos una plataforma",
                    "La publicación está relacionada con la campaña o plan de contenidos",
                    "La evidencia es verificable (URL, captura con fecha visible o número de publicación)",
                ],
            },
            {
                "nombre": "03_4_5_PORTAFOLIO_DIGITAL_CAMPAÑA",
                "tipo": "documento",
                "descripcion": "Portafolio digital que consolida toda la campaña de contenidos y anuncios",
                "criterios": [
                    "Consolida todas las piezas y acciones de la campaña en un solo documento",
                    "Incluye la estrategia, los contenidos creados y la evidencia de publicación",
                    "Presenta los resultados o proyección de alcance de la campaña",
                    "Tiene estructura de portafolio con secciones claramente identificadas",
                    "Concluye con aprendizajes sobre el proceso de publicación",
                ],
            },
        ],
    },

    "GUIA_04_ORGANIZAR_EXHIBICION": {
        "guia": "GUIA_04_ORGANIZAR_EXHIBICION",
        "programa": "Comunicacion_y_marketing",
        "criterio_aprobacion": _CRITERIO_STD,
        "evidencias": [
            {
                "nombre": "DIORAMA_ELEMENTOS_BASICOS_EXHIBICION",
                "tipo": "imagen",
                "descripcion": "Fotografía de diorama con los elementos básicos de exhibición construido",
                "criterios": [
                    "Es una imagen que muestra un diorama o maqueta de exhibición construida",
                    "Evidencia los elementos básicos de exhibición (iluminación, señalización, disposición u otros)",
                    "La imagen es clara y permite identificar los elementos de la exhibición",
                    "Refleja un ejercicio práctico realizado por el aprendiz",
                ],
            },
            {
                "nombre": "04_3_3_1_DOCUMENTO_BATERIA_PREGUNTAS",
                "tipo": "documento",
                "descripcion": "Batería de preguntas y reglas de juego para clasificación de productos",
                "criterios": [
                    "Contiene una batería de preguntas sobre clasificación de productos",
                    "Incluye reglas de juego o dinámica asociada a la clasificación",
                    "Las preguntas están relacionadas con criterios de organización de productos",
                    "El documento está completo con preguntas y respuestas",
                ],
            },
            {
                "nombre": "04_3_3_2_ANALISIS_EMPAQUES_PRODUCTOS",
                "tipo": "documento",
                "descripcion": "Análisis de empaques de productos elaborado con Canva u otra herramienta",
                "criterios": [
                    "Analiza el empaque de al menos 2 productos reales",
                    "Evalúa aspectos del empaque: material, diseño, información legal y marca",
                    "Usa herramienta digital (Canva u otra) para presentar el análisis",
                    "Incluye conclusiones sobre la efectividad del empaque analizado",
                ],
            },
            {
                "nombre": "04_3_3_3_REJILLA_CONCEPTOS_LOGISTICA",
                "tipo": "documento",
                "descripcion": "Rejilla de conceptos sobre logística y almacenamiento de productos",
                "criterios": [
                    "Presenta una rejilla o tabla con conceptos de logística y almacenamiento",
                    "Define términos clave: inventario, stock, rotación, FIFO/LIFO u otros",
                    "Relaciona los conceptos con la organización del punto de venta",
                    "La rejilla está completa con definiciones y ejemplos o aplicaciones",
                ],
            },
            {
                "nombre": "INFOGRAFIA_CONCEPTOS_IMPLANTACION",
                "tipo": "imagen",
                "descripcion": "Infografía sobre los conceptos de implantación de productos en el punto de venta",
                "criterios": [
                    "Es una infografía con diseño visual sobre implantación de productos",
                    "Presenta conceptos clave (planograma, lineal, zona de compra u otros)",
                    "Tiene elementos visuales que refuerzan los conceptos",
                    "Es comprensible como pieza autónoma de información",
                ],
            },
            {
                "nombre": "04_3_3_6_ORGANIZADOR_GRAFICO_LINEAL",
                "tipo": "imagen",
                "descripcion": "Organizador gráfico sobre el lineal del punto de venta (puede estar escaneado)",
                "criterios": [
                    "Es un organizador gráfico o esquema sobre el lineal de un punto de venta",
                    "Muestra la distribución o planificación del espacio en el lineal",
                    "Identifica zonas, niveles o categorías del lineal",
                    "El esquema es legible aunque esté escaneado del cuaderno",
                ],
            },
            {
                "nombre": "04_3_3_7_DISENO_LINEAL_INFORME",
                "tipo": "documento",
                "descripcion": "Diseño del lineal con informe ejecutivo de implantación de productos",
                "criterios": [
                    "Presenta el diseño del lineal con distribución de productos definida",
                    "Incluye un informe ejecutivo que justifica las decisiones de implantación",
                    "Aplica criterios de rentabilidad por espacio, rotación o margen",
                    "Tiene estructura de documento formal con secciones identificables",
                ],
            },
            {
                "nombre": "04_3_3_8_MAPA_MENTAL_GLOSARIO_SEGURIDAD",
                "tipo": "documento",
                "descripcion": "Mapa mental y glosario sobre seguridad en el punto de venta",
                "criterios": [
                    "Incluye un mapa mental sobre seguridad en el punto de venta",
                    "El glosario define términos de seguridad: prevención de pérdidas, hurto, EHS u otros",
                    "Relaciona la seguridad con la organización y exhibición de productos",
                    "Combina representación visual con definiciones escritas",
                ],
            },
            {
                "nombre": "04_3_3_9_TALLER_INVENTARIOS",
                "tipo": "imagen",
                "descripcion": "Taller de inventarios con evidencia escaneada del cuaderno o actividad",
                "criterios": [
                    "Es una imagen o escáner de un taller sobre inventarios",
                    "Contiene ejercicios o actividades sobre control de inventario",
                    "Aplica conceptos de conteo, registro o gestión de stock",
                    "La imagen es legible y muestra el trabajo realizado",
                ],
            },
            {
                "nombre": "04_3_4_PROPUESTA_EXHIBICION_MAQUETA_INFORME",
                "tipo": "cualquier",
                "descripcion": "Maqueta o propuesta de exhibición con informe para el proyecto",
                "criterios": [
                    "Presenta una maqueta física o propuesta visual de exhibición para la empresa",
                    "El informe describe los criterios de diseño usados en la exhibición",
                    "Aplica conceptos de planograma, zonificación y técnicas de merchandising",
                    "La propuesta está adaptada al sector y perfil de clientes de la empresa",
                    "Incluye conclusiones y recomendaciones de implementación",
                ],
            },
        ],
    },

    "GUIA_05_EJECUTAR_ACCIONES_PROMOCIONALES": {
        "guia": "GUIA_05_EJECUTAR_ACCIONES_PROMOCIONALES",
        "programa": "Comunicacion_y_marketing",
        "criterio_aprobacion": _CRITERIO_STD,
        "evidencias": [
            {
                "nombre": "05_3_1_PLAN_PROMOCIONAL_CARTELERA_PRESENTACION",
                "tipo": "presentacion",
                "descripcion": "Cartelera o presentación con el plan promocional",
                "criterios": [
                    "Es una presentación o cartelera visual sobre un plan promocional",
                    "Describe los objetivos de la promoción y el público objetivo",
                    "Presenta las acciones o actividades promocionales planificadas",
                    "Tiene diseño visual atractivo adecuado para comunicar la propuesta",
                ],
            },
            {
                "nombre": "05_3_2_FICHA_TECNICA_BANNER_PROMOCION",
                "tipo": "cualquier",
                "descripcion": "Ficha técnica, banner y presentación del plan de promoción",
                "criterios": [
                    "Incluye una ficha técnica con especificaciones de la promoción",
                    "Presenta un banner o material gráfico de la promoción diseñado",
                    "El banner tiene mensaje claro, imagen impactante y llamado a la acción",
                    "La presentación integra la ficha técnica y el material creativo",
                ],
            },
            {
                "nombre": "05_3_3_MAPA_MENTAL_EVENTOS",
                "tipo": "imagen",
                "descripcion": "Mapa mental sobre organización y tipos de eventos comerciales y promocionales",
                "criterios": [
                    "Es un mapa mental sobre eventos comerciales o promocionales",
                    "Identifica tipos de eventos: ferias, activaciones, lanzamientos u otros",
                    "Describe etapas o elementos clave en la organización de eventos",
                    "Tiene estructura de mapa mental con nodo central y ramificaciones",
                ],
            },
            {
                "nombre": "05_3_4_INFORME_OBSERVACION_PUNTO_VENTA",
                "tipo": "documento",
                "descripcion": "Informe de observación de estrategias promocionales en punto de venta real",
                "criterios": [
                    "Describe la observación realizada en un punto de venta real",
                    "Identifica al menos 3 acciones o estrategias promocionales observadas",
                    "Analiza la efectividad de las estrategias identificadas",
                    "Concluye con aprendizajes aplicables al proyecto",
                ],
            },
            {
                "nombre": "05_3_5_PRE_BRIEF_CRONOGRAMA_EVENTO",
                "tipo": "documento",
                "descripcion": "Brief y cronograma de planeación del evento promocional",
                "criterios": [
                    "Presenta un brief con los elementos del evento: objetivo, fecha, lugar y público",
                    "Incluye un cronograma con actividades y responsables definidos",
                    "El brief está orientado a un evento real o simulado con detalles concretos",
                    "El cronograma tiene fechas o tiempos de ejecución definidos",
                ],
            },
            {
                "nombre": "05_3_5_EJECUCION_EVENTO_REGISTROS",
                "tipo": "imagen",
                "descripcion": "Evidencia fotográfica o registros de la ejecución de la activación promocional",
                "criterios": [
                    "Es una imagen o captura que evidencia la realización del evento o activación",
                    "Se puede identificar el contexto del evento (lugar, participantes o materiales)",
                    "La evidencia está relacionada con el plan de evento preparado",
                    "Muestra el desarrollo real o simulado de la activación",
                ],
            },
            {
                "nombre": "05_3_5_POST_INFORME_INDICADORES_ENCUESTA",
                "tipo": "documento",
                "descripcion": "Informe final del evento con indicadores de gestión y encuesta de satisfacción",
                "criterios": [
                    "Presenta un informe post-evento con evaluación de resultados",
                    "Incluye indicadores de gestión (asistencia, alcance, ventas generadas u otros)",
                    "Adjunta o describe una encuesta de satisfacción aplicada a los participantes",
                    "Concluye con aprendizajes y recomendaciones para futuros eventos",
                ],
            },
            {
                "nombre": "05_4_EVIDENCIA_INTEGRAL_ACTIVACION_PROMOCIONAL",
                "tipo": "cualquier",
                "descripcion": "Evidencia integral que consolida toda la activación promocional del proyecto",
                "criterios": [
                    "Consolida todas las etapas de la activación: planeación, ejecución y evaluación",
                    "Presenta resultados medibles de la activación realizada",
                    "Incluye reflexión sobre el impacto y aprendizajes del proceso",
                    "Tiene formato de portafolio, video resumen o presentación completa",
                    "Está alineado con los objetivos del proyecto empresarial",
                ],
            },
        ],
    },

    "GUIA_06_VENDER_PRODUCTOS": {
        "guia": "GUIA_06_VENDER_PRODUCTOS",
        "programa": "Comunicacion_y_marketing",
        "criterio_aprobacion": _CRITERIO_STD,
        "evidencias": [
            {
                "nombre": "06_3_1_COMIC_SERVICIO_CLIENTE",
                "tipo": "imagen",
                "descripcion": "Cómic que ilustra una experiencia de servicio al cliente en contexto comercial",
                "criterios": [
                    "Es un cómic con viñetas o secuencia narrativa visual",
                    "Ilustra una interacción de servicio al cliente en un contexto de ventas",
                    "Refleja actitudes positivas o negativas de servicio al cliente",
                    "Tiene al menos 4 viñetas con diálogos o situaciones identificables",
                ],
            },
            {
                "nombre": "06_3_2_LISTADO_CARACTERISTICAS_VENDEDOR",
                "tipo": "documento",
                "descripcion": "Listado de características de vendedores exitosos (panel de discusión) — opcional",
                "criterios": [
                    "Presenta un listado de características de vendedores exitosos",
                    "Incluye al menos 5 características con descripción o justificación",
                    "Refleja reflexión del aprendiz sobre las cualidades de un buen vendedor",
                    "El documento no está en blanco",
                ],
            },
            {
                "nombre": "06_3_3_1_VIDEO_JUEGO_ROLES_COMUNICACION",
                "tipo": "cualquier",
                "descripcion": "Video de juego de roles que simula una comunicación comercial",
                "criterios": [
                    "Es un video o evidencia de un juego de roles de comunicación comercial",
                    "Simula una interacción real de venta o atención al cliente",
                    "Aplica técnicas de comunicación verbal y no verbal en ventas",
                    "Los roles están claramente definidos (vendedor y cliente)",
                ],
            },
            {
                "nombre": "06_3_3_2_ANALISIS_PRODUCTOS_ORGANIZADOR",
                "tipo": "documento",
                "descripcion": "Análisis de productos con organizador gráfico del portafolio",
                "criterios": [
                    "Analiza al menos 3 productos del portafolio de la empresa del proyecto",
                    "Incluye un organizador gráfico que estructura la información del portafolio",
                    "Identifica características, beneficios y argumentos de venta de cada producto",
                    "El análisis está orientado a facilitar la presentación comercial",
                ],
            },
            {
                "nombre": "06_3_3_3_DOCUMENTO_CONCEPTOS_PRODUCTO",
                "tipo": "documento",
                "descripcion": "Documento sobre conceptos de producto y portafolio comercial",
                "criterios": [
                    "Define conceptos clave: producto, portafolio, catálogo, ficha técnica",
                    "Distingue entre características y beneficios de un producto",
                    "Aplica los conceptos al portafolio de la empresa del proyecto",
                    "El documento tiene contenido desarrollado, no solo definiciones copiadas",
                ],
            },
            {
                "nombre": "06_3_3_4_REJILLA_CONCEPTOS_VENTAS",
                "tipo": "documento",
                "descripcion": "Rejilla de conceptos sobre ventas y observación del entorno comercial",
                "criterios": [
                    "Presenta una rejilla o tabla con conceptos clave del proceso de ventas",
                    "Incluye observación del entorno comercial del proyecto",
                    "Define términos: prospección, argumentación, objeción, cierre, postventa",
                    "La rejilla está completa con definiciones y aplicaciones al contexto real",
                ],
            },
            {
                "nombre": "06_3_3_5_PRESENTACION_EMPRESA_PORTAFOLIO",
                "tipo": "presentacion",
                "descripcion": "Presentación de la empresa con portafolio, ficha técnica, catálogo y brochure",
                "criterios": [
                    "Es una presentación formal de la empresa del proyecto",
                    "Incluye el portafolio de productos con ficha técnica o catálogo",
                    "Tiene diseño visual de brochure o material comercial profesional",
                    "Está diseñada para presentar ante clientes potenciales",
                ],
            },
            {
                "nombre": "06_3_3_6_HERRAMIENTA_GRAFICA_PROCESO_VENTAS",
                "tipo": "documento",
                "descripcion": "Herramienta gráfica del proceso de ventas con informe",
                "criterios": [
                    "Presenta una herramienta gráfica (flujograma, mapa o infografía) del proceso de ventas",
                    "Describe las etapas del proceso de ventas de la empresa del proyecto",
                    "El informe justifica cada etapa con criterios técnicos o comerciales",
                    "Combina representación visual y explicación escrita",
                ],
            },
            {
                "nombre": "06_3_3_7_INFOGRAFIA_ESTRATEGIAS_VENTAS",
                "tipo": "cualquier",
                "descripcion": "Infografía o presentación sobre estrategias de ventas",
                "criterios": [
                    "Es una infografía, presentación o video sobre estrategias de ventas",
                    "Presenta al menos 3 estrategias de venta con descripción y aplicación",
                    "Relaciona las estrategias con el contexto del proyecto empresarial",
                    "Tiene diseño visual atractivo y contenido organizado",
                ],
            },
            {
                "nombre": "06_3_3_8_TECNICA_SOMBREROS_MAPA_CONCLUSIONES",
                "tipo": "documento",
                "descripcion": "Actividad con técnica de los seis sombreros, mapa mental y conclusiones",
                "criterios": [
                    "Desarrolla la técnica de los seis sombreros para analizar un tema de ventas",
                    "Identifica la perspectiva de cada sombrero con argumentos concretos",
                    "Incluye un mapa mental con los resultados del análisis",
                    "Presenta conclusiones integrando las perspectivas de todos los sombreros",
                ],
            },
            {
                "nombre": "06_3_3_9_COMIC_ARGUMENTACION_CIERRE_VENTAS",
                "tipo": "imagen",
                "descripcion": "Cómic sobre argumentación de venta y técnicas de cierre",
                "criterios": [
                    "Es un cómic con secuencia visual sobre argumentación y cierre de ventas",
                    "Ilustra el manejo de objeciones por parte del vendedor",
                    "Muestra al menos una técnica de cierre de venta aplicada en el diálogo",
                    "Tiene al menos 4 viñetas con diálogos identificables",
                ],
            },
            {
                "nombre": "06_4_PROTOCOLO_VENTA_MATERIAL_COMERCIAL",
                "tipo": "documento",
                "descripcion": "Protocolo de venta y material comercial de la empresa del proyecto",
                "criterios": [
                    "Presenta un protocolo formal de venta con pasos y guiones definidos",
                    "El protocolo está adaptado a la empresa y productos del proyecto",
                    "Incluye material comercial de apoyo (argumentarios, manejo de objeciones, script de cierre)",
                    "Tiene estructura de protocolo con secciones claramente identificadas",
                    "Concluye con indicadores para evaluar el desempeño en ventas",
                ],
            },
        ],
    },

    "GUIA_07_INTERACTUAR_CON_AUDIENCIAS": {
        "guia": "GUIA_07_INTERACTUAR_CON_AUDIENCIAS",
        "programa": "Comunicacion_y_marketing",
        "criterio_aprobacion": _CRITERIO_STD,
        "evidencias": [
            {
                "nombre": "07_3_1_TEXTO_REFLEXIVO_COMUNIDADES_DIGITALES",
                "tipo": "documento",
                "descripcion": "Texto reflexivo sobre comunidades digitales y su rol en el marketing",
                "criterios": [
                    "Es un texto con reflexión propia sobre comunidades digitales",
                    "Define el concepto de comunidad digital y sus características",
                    "Analiza el rol de las comunidades digitales en el marketing y las ventas",
                    "Presenta ejemplos de comunidades digitales relevantes para el sector",
                ],
            },
            {
                "nombre": "07_3_2_ANALISIS_CASO_INTERACCION_DIGITAL",
                "tipo": "documento",
                "descripcion": "Análisis de caso sobre interacción digital entre marca y audiencia — opcional",
                "criterios": [
                    "Analiza un caso real de interacción digital entre una marca y su audiencia",
                    "Identifica estrategias de engagement usadas en el caso",
                    "Evalúa el impacto de la interacción en la relación marca-comunidad",
                    "Presenta aprendizajes aplicables al proyecto",
                ],
            },
            {
                "nombre": "07_3_3_1_BUYER_PERSONA_FICHA_PRESENTACION",
                "tipo": "documento",
                "descripcion": "Buyer Persona completo con ficha y presentación",
                "criterios": [
                    "Define un Buyer Persona con nombre, demografía, motivaciones y frustraciones",
                    "La ficha incluye comportamiento digital y canales de comunicación preferidos",
                    "La presentación adapta el contenido al perfil del Buyer Persona definido",
                    "El Buyer Persona está orientado a la empresa del proyecto",
                ],
            },
            {
                "nombre": "07_3_3_2_COMUNIDAD_DIGITAL_MODERADA",
                "tipo": "presentacion",
                "descripcion": "Diseño de comunidad digital: reglas de moderación, roles y ética digital",
                "criterios": [
                    "Presenta el diseño de una comunidad digital con propósito definido",
                    "Incluye reglas de moderación y netiqueta para la comunidad",
                    "Define roles de participación (administrador, moderador, miembro u otros)",
                    "Incorpora criterios de ética digital y manejo de conflictos en la comunidad",
                ],
            },
            {
                "nombre": "07_3_3_3_PROTOCOLO_COMUNICACION_DIGITAL",
                "tipo": "documento",
                "descripcion": "Protocolo de comunicación digital para la marca o empresa del proyecto",
                "criterios": [
                    "Presenta un protocolo de comunicación digital con tono de voz y lineamientos",
                    "Define el estilo de comunicación para distintas situaciones",
                    "Incluye plantillas o ejemplos de respuestas para situaciones comunes",
                    "Está adaptado al perfil de audiencia de la empresa del proyecto",
                ],
            },
            {
                "nombre": "07_3_4_CONTENIDO_CORREGIDO_ADAPTADO",
                "tipo": "documento",
                "descripcion": "Contenido digital corregido y adaptado para la audiencia objetivo",
                "criterios": [
                    "Presenta contenido digital revisado y adaptado para una audiencia específica",
                    "La adaptación considera el tono, formato y plataforma de destino",
                    "Muestra la versión antes y después de la corrección o adaptación",
                    "El contenido adaptado está listo para publicación en la plataforma definida",
                ],
            },
            {
                "nombre": "07_4_PORTAFOLIO_CAMPANA_DIGITAL",
                "tipo": "cualquier",
                "descripcion": "Portafolio completo de la campaña digital: estrategia, piezas y resultados",
                "criterios": [
                    "Consolida la estrategia de interacción con audiencias del proyecto",
                    "Incluye las piezas de contenido creadas y el plan de comunidad digital",
                    "Presenta los protocolos de comunicación y los contenidos adaptados",
                    "Incluye análisis de resultados o proyección de impacto de la campaña",
                    "Tiene estructura de portafolio con secciones claramente identificadas",
                ],
            },
        ],
    },

    "GUIA_08_REALIZAR_SEGUIMIENTO": {
        "guia": "GUIA_08_REALIZAR_SEGUIMIENTO",
        "programa": "Comunicacion_y_marketing",
        "criterio_aprobacion": _CRITERIO_STD,
        "evidencias": [
            {
                "nombre": "08_3_1_CUADRO_COMPARATIVO_SERVICIO_CLIENTE",
                "tipo": "documento",
                "descripcion": "Cuadro comparativo con conclusiones sobre modelos de servicio al cliente",
                "criterios": [
                    "Presenta un cuadro comparativo de al menos 2 modelos de servicio al cliente",
                    "Compara criterios como enfoque, herramientas, métricas o resultados",
                    "Incluye conclusiones propias sobre el mejor enfoque según el contexto",
                    "El cuadro está completo y organizado con criterios claramente definidos",
                ],
            },
            {
                "nombre": "08_3_2_RUTA_CLIENTE_PRESENTACION",
                "tipo": "cualquier",
                "descripcion": "Ruta del cliente presentada en formato de historieta, video o mural",
                "criterios": [
                    "Presenta la ruta del cliente (Customer Journey) de la empresa del proyecto",
                    "El formato es visual: historieta, video, infografía o mural",
                    "Identifica los puntos de contacto principales entre cliente y empresa",
                    "Describe emociones o expectativas del cliente en cada etapa",
                ],
            },
            {
                "nombre": "08_3_3_ANALISIS_MOMENTOS_VERDAD_CLIENTE",
                "tipo": "documento",
                "descripcion": "Análisis de momentos de la verdad y estrategias de mejora del servicio",
                "criterios": [
                    "Define el concepto de 'momento de la verdad' y lo aplica al caso del proyecto",
                    "Identifica al menos 3 momentos de la verdad críticos en la experiencia del cliente",
                    "Propone estrategias concretas para mejorar cada momento identificado",
                    "Concluye con un plan de mejora del servicio al cliente",
                ],
            },
            {
                "nombre": "08_3_4_1_REPRESENTACION_PRINCIPIOS_CALIDAD",
                "tipo": "cualquier",
                "descripcion": "Representación creativa de un principio de la norma ISO sobre calidad",
                "criterios": [
                    "Presenta un principio de la norma ISO 9001 o de gestión de calidad",
                    "La representación es creativa: video, póster, historieta o infografía",
                    "Explica cómo se aplica el principio al servicio al cliente",
                    "El contenido demuestra comprensión del principio elegido",
                ],
            },
            {
                "nombre": "08_3_4_2_PROTOCOLO_PQRSF_RESOLUCION",
                "tipo": "documento",
                "descripcion": "Protocolo de PQRSF con resolución de casos aplicados",
                "criterios": [
                    "Presenta un protocolo formal para la gestión de PQRSF",
                    "Incluye los pasos para recepcionar, tramitar y resolver cada tipo de PQRSF",
                    "Aplica el protocolo a al menos 2 casos prácticos resueltos",
                    "El protocolo está adaptado a la empresa del proyecto",
                ],
            },
            {
                "nombre": "08_3_4_3_DIAGRAMA_CAUSA_EFECTO_PLAN_ACCION",
                "tipo": "imagen",
                "descripcion": "Diagrama de causa-efecto (Ishikawa) con plan de acción para problema de servicio",
                "criterios": [
                    "Es un diagrama de Ishikawa (espina de pescado) sobre un problema de servicio",
                    "Identifica al menos 4 categorías de causas del problema analizado",
                    "El plan de acción propone soluciones para las causas raíz identificadas",
                    "El diagrama y el plan están relacionados coherentemente",
                ],
            },
            {
                "nombre": "08_3_4_4_ESTRATEGIA_FIDELIZACION_CLIENTES",
                "tipo": "documento",
                "descripcion": "Estrategia de fidelización de clientes para la empresa del proyecto",
                "criterios": [
                    "Presenta una estrategia de fidelización con objetivo y acciones definidas",
                    "Incluye al menos 2 programas o mecanismos de fidelización",
                    "Describe los beneficios para el cliente y el impacto esperado para la empresa",
                    "La estrategia está adaptada al segmento de clientes de la empresa",
                ],
            },
            {
                "nombre": "08_3_4_5_CRM_BASE_DATOS_CLIENTES",
                "tipo": "excel",
                "descripcion": "CRM o base de datos de clientes en hoja de cálculo",
                "criterios": [
                    "Es una hoja de cálculo con estructura de CRM o base de datos de clientes",
                    "Contiene al menos 5 registros de clientes con información relevante",
                    "Incluye campos como nombre, contacto, historial de compras o segmento",
                    "Está organizada para facilitar el seguimiento de clientes",
                ],
            },
            {
                "nombre": "08_3_4_6_TRAZABILIDAD_PROCESOS_PRESENTACION",
                "tipo": "presentacion",
                "descripcion": "Presentación sobre la trazabilidad de procesos de atención al cliente",
                "criterios": [
                    "Es una presentación sobre trazabilidad de procesos de atención",
                    "Describe el flujo de atención al cliente desde el primer contacto hasta el cierre",
                    "Identifica los registros o evidencias de trazabilidad en cada etapa",
                    "Tiene diseño visual de presentación formal con estructura lógica",
                ],
            },
            {
                "nombre": "08_3_4_7_CUESTIONARIO_CONOCIMIENTO",
                "tipo": "cualquier",
                "descripcion": "Cuestionario de evidencia de conocimiento sobre seguimiento al cliente",
                "criterios": [
                    "Contiene preguntas sobre seguimiento al cliente y gestión de calidad",
                    "Tiene mínimo 5 ítems respondidos",
                    "Las respuestas reflejan comprensión de los temas de la guía",
                    "No está en blanco ni presenta solo el encabezado",
                ],
            },
            {
                "nombre": "08_4_EVIDENCIA_TRANSFERENCIA_SEGUIMIENTO_CLIENTE",
                "tipo": "cualquier",
                "descripcion": "Evidencia de transferencia que integra PQRSF, fidelización, CRM y encuesta del proyecto",
                "criterios": [
                    "Consolida las evidencias de seguimiento al cliente del proyecto",
                    "Incluye el sistema de PQRSF, la estrategia de fidelización y el CRM",
                    "Presenta resultados de una encuesta de satisfacción aplicada",
                    "Tiene formato de portafolio, presentación o informe integrado",
                    "Concluye con aprendizajes y recomendaciones sobre la gestión del cliente",
                ],
            },
        ],
    },

    "GUIA_09_MONITOREAR_METRICAS": {
        "guia": "GUIA_09_MONITOREAR_METRICAS",
        "programa": "Comunicacion_y_marketing",
        "criterio_aprobacion": _CRITERIO_STD,
        "evidencias": [
            {
                "nombre": "09_3_1_TABLA_INDICADORES_VIDEO_ESTRATEGIA",
                "tipo": "cualquier",
                "descripcion": "Tabla de indicadores de gestión y video corto sobre estrategia de métricas",
                "criterios": [
                    "Presenta una tabla con indicadores de gestión relevantes para el proyecto",
                    "Define al menos 4 indicadores con su fórmula, meta y periodicidad",
                    "El video o evidencia complementaria explica la estrategia de monitoreo",
                    "Los indicadores están relacionados con los objetivos del proyecto",
                ],
            },
            {
                "nombre": "09_3_2_GLOSARIO_METRICAS_DIGITALES",
                "tipo": "documento",
                "descripcion": "Glosario de términos y métricas del marketing digital",
                "criterios": [
                    "Contiene definiciones de métricas digitales: CTR, CPM, ROI, engagement, alcance u otras",
                    "Tiene mínimo 10 términos definidos",
                    "Las definiciones incluyen ejemplos o fórmulas donde aplica",
                    "El glosario está organizado de forma clara (alfabética o temática)",
                ],
            },
            {
                "nombre": "09_3_3_1_CARTELERA_GESTION_INFORMACION",
                "tipo": "imagen",
                "descripcion": "Cartelera visual sobre gestión de la información y métricas",
                "criterios": [
                    "Es una cartelera o infografía visual sobre gestión de la información",
                    "Presenta el proceso o flujo de captura, análisis y uso de datos",
                    "Incluye elementos visuales que faciliten la comprensión del tema",
                    "Es comprensible como pieza autónoma de comunicación",
                ],
            },
            {
                "nombre": "09_3_3_2_CONTENIDO_ANALISIS_METRICAS_INFORME",
                "tipo": "cualquier",
                "descripcion": "Contenido digital con análisis de métricas e informe ejecutivo",
                "criterios": [
                    "Presenta un análisis real o simulado de métricas de una campaña o plataforma digital",
                    "Interpreta los datos obtenidos y extrae conclusiones estratégicas",
                    "El informe ejecutivo resume los hallazgos y recomendaciones",
                    "Usa datos concretos (reales o simulados) con gráficos o tablas",
                ],
            },
            {
                "nombre": "09_3_3_3_PRUEBA_CONOCIMIENTO_METRICAS",
                "tipo": "cualquier",
                "descripcion": "Prueba de evidencia de conocimiento sobre métricas digitales",
                "criterios": [
                    "Contiene preguntas sobre métricas y monitoreo de campañas digitales",
                    "Tiene mínimo 5 ítems respondidos",
                    "Las respuestas reflejan comprensión de los indicadores y métricas del marketing digital",
                    "No está en blanco ni presenta solo el encabezado",
                ],
            },
            {
                "nombre": "09_4_INFORME_METRICAS_CAMPANA",
                "tipo": "documento",
                "descripcion": "Informe final de métricas de la campaña del proyecto con gráficas",
                "criterios": [
                    "Presenta el informe final de métricas de la campaña digital del proyecto",
                    "Incluye al menos 4 indicadores con sus valores reales o proyectados",
                    "Usa gráficas o visualizaciones para presentar los datos",
                    "Compara los resultados con los objetivos iniciales de la campaña",
                    "Concluye con recomendaciones de optimización para futuras campañas",
                ],
            },
        ],
    },

    # ══════════════════════════════════════════════════════════════════════════
    # PROGRAMA: Ventas_de_productos_en_linea
    # ══════════════════════════════════════════════════════════════════════════

    "GUIA_01_CONSTRUIR_RELACION_CLIENTE": {
        "guia": "GUIA_01_CONSTRUIR_RELACION_CLIENTE",
        "programa": "Ventas_de_productos_en_linea",
        "criterio_aprobacion": _CRITERIO_STD,
        "evidencias": [
            {
                "nombre": "01_1_CUADRO_TECNOLOGIAS",
                "tipo": "documento",
                "descripcion": "Cuadro comparativo sobre el impacto de las tecnologías en las ventas en línea",
                "criterios": [
                    "Presenta un cuadro comparativo de tecnologías aplicadas al comercio digital",
                    "Compara al menos 3 tecnologías con criterios definidos (impacto, ventaja, uso u otros)",
                    "Relaciona cada tecnología con la construcción de relaciones con clientes en línea",
                    "El cuadro está completo con contenido en todas las celdas",
                ],
            },
            {
                "nombre": "01_2_ROLES_EXPERIENCIAS_EN_LINEA",
                "tipo": "documento",
                "descripcion": "Documento de roles y experiencias en línea (chat simulado o juego de roles digital)",
                "criterios": [
                    "Presenta un juego de roles o chat simulado de interacción digital con cliente",
                    "Define los roles participantes: vendedor en línea y cliente",
                    "El diálogo refleja habilidades de comunicación digital y atención al cliente",
                    "El documento está desarrollado con contenido completo",
                ],
            },
            {
                "nombre": "01_3_1_VOCABULARIO_MARKETING_DIGITAL",
                "tipo": "documento",
                "descripcion": "Glosario o vocabulario de términos de marketing digital aplicado a ventas en línea",
                "criterios": [
                    "Contiene definiciones de términos de marketing digital (e-commerce, omnicanalidad, UX, conversión u otros)",
                    "Tiene mínimo 8 términos definidos",
                    "Las definiciones relacionan los términos con las ventas de productos en línea",
                    "El glosario está organizado de forma clara",
                ],
            },
            {
                "nombre": "01_3_2_FOLLETO_CANALES_DIGITALES",
                "tipo": "documento",
                "descripcion": "Folleto informativo sobre canales digitales de venta en línea",
                "criterios": [
                    "Presenta un folleto con información sobre canales digitales de venta",
                    "Describe al menos 3 canales (redes sociales, marketplace, tienda propia u otros)",
                    "Para cada canal identifica ventajas, limitaciones y tipo de audiencia",
                    "El folleto tiene diseño visual atractivo aunque sea sencillo",
                ],
            },
            {
                "nombre": "01_3_3_TALLER_MERCADO_DIGITAL",
                "tipo": "documento",
                "descripcion": "Taller sobre el mercado digital y sus características",
                "criterios": [
                    "Describe las características del mercado digital en el contexto colombiano",
                    "Identifica tendencias del e-commerce relevantes para el sector del proyecto",
                    "Analiza al menos un dato estadístico sobre comercio electrónico en Colombia",
                    "El taller tiene respuestas completas, no solo palabras clave",
                ],
            },
            {
                "nombre": "01_4_INFORME_CLIENTE_Y_ATENCION",
                "tipo": "presentacion",
                "descripcion": "Informe sobre el cliente digital y estrategias de atención en línea",
                "criterios": [
                    "Es una presentación formal sobre el cliente digital y su comportamiento",
                    "Identifica el perfil del cliente digital de la empresa del proyecto",
                    "Describe las estrategias de atención en línea más adecuadas para ese perfil",
                    "Tiene estructura de presentación con secciones claramente identificadas",
                    "Concluye con recomendaciones de atención digital para el proyecto",
                ],
            },
        ],
    },

    "GUIA_02_PROCEDIMIENTOS_CON_EL_CLIENTE": {
        "guia": "GUIA_02_PROCEDIMIENTOS_CON_EL_CLIENTE",
        "programa": "Ventas_de_productos_en_linea",
        "criterio_aprobacion": _CRITERIO_STD,
        "evidencias": [
            {
                "nombre": "02_1_REFLEXION_CANALES_CONFIDENCIALIDAD",
                "tipo": "documento",
                "descripcion": "Reflexión sobre el uso de canales digitales y la confidencialidad de datos del cliente",
                "criterios": [
                    "Presenta una reflexión sobre el manejo de información confidencial en canales digitales",
                    "Menciona la normatividad de protección de datos (Ley 1581 u otras)",
                    "Identifica riesgos de confidencialidad en el uso de canales digitales de venta",
                    "Incluye compromisos o prácticas para garantizar la confidencialidad",
                ],
            },
            {
                "nombre": "02_2_CASOS_CONFIDENCIALIDAD",
                "tipo": "documento",
                "descripcion": "Resolución de casos sobre confidencialidad y protección de datos del cliente",
                "criterios": [
                    "Presenta al menos 2 casos sobre manejo de datos confidenciales en ventas digitales",
                    "Analiza cada caso identificando el dilema ético o legal",
                    "Propone la solución correcta con justificación normativa o ética",
                    "Las respuestas son elaboradas, no de una sola línea",
                ],
            },
            {
                "nombre": "02_3_TALLER_REQUERIMIENTOS_CLIENTE",
                "tipo": "documento",
                "descripcion": "Taller sobre identificación y gestión de requerimientos del cliente en línea",
                "criterios": [
                    "Define el concepto de requerimiento del cliente en el contexto digital",
                    "Describe el proceso para identificar y registrar requerimientos en canales en línea",
                    "Aplica el proceso a un caso o escenario concreto de venta digital",
                    "El taller tiene respuestas completas y argumentadas",
                ],
            },
            {
                "nombre": "02_3_PRUEBA_PROCEDIMIENTOS_CLIENTE",
                "tipo": "cualquier",
                "descripcion": "Prueba de conocimiento sobre procedimientos con el cliente",
                "criterios": [
                    "Contiene preguntas sobre procedimientos de atención y gestión de clientes en línea",
                    "Tiene mínimo 5 ítems respondidos",
                    "Las respuestas reflejan comprensión de los procedimientos estudiados",
                    "No está en blanco ni presenta solo el encabezado",
                ],
            },
            {
                "nombre": "02_4_INFORME_RECEPCIONAR_REQUERIMIENTOS",
                "tipo": "documento",
                "descripcion": "Informe de recepción de requerimientos del cliente para el proyecto",
                "criterios": [
                    "Presenta el proceso de recepción de requerimientos de la empresa del proyecto",
                    "Describe el flujo de atención desde la solicitud hasta la respuesta al cliente",
                    "Incluye formatos o plantillas de registro de requerimientos",
                    "Concluye con recomendaciones para mejorar el proceso de recepción",
                    "Tiene estructura de informe con secciones identificables",
                ],
            },
        ],
    },

    "GUIA_03_OPTIMIZAR_ATENCION_CLIENTE": {
        "guia": "GUIA_03_OPTIMIZAR_ATENCION_CLIENTE",
        "programa": "Ventas_de_productos_en_linea",
        "criterio_aprobacion": _CRITERIO_STD,
        "evidencias": [
            {
                "nombre": "03_3_TALLER_ATENCION_CLIENTE",
                "tipo": "documento",
                "descripcion": "Taller sobre optimización de la atención al cliente en canales digitales",
                "criterios": [
                    "Identifica cuellos de botella o problemas en el proceso de atención al cliente digital",
                    "Propone mejoras concretas para optimizar la atención en cada canal",
                    "Aplica herramientas o metodologías de mejora (tiempo de respuesta, chatbot, FAQ u otras)",
                    "El taller tiene respuestas completas con argumentos",
                ],
            },
            {
                "nombre": "02_3_PRUEBA_PROCEDIMIENTOS_CLIENTE",
                "tipo": "cualquier",
                "descripcion": "Prueba de conocimiento sobre procedimientos y optimización de atención al cliente",
                "criterios": [
                    "Contiene preguntas sobre procedimientos y optimización de atención al cliente",
                    "Tiene mínimo 5 ítems respondidos",
                    "Las respuestas reflejan comprensión de los temas de optimización",
                    "No está en blanco ni presenta solo el encabezado",
                ],
            },
            {
                "nombre": "02_4_INFORME_OPTIMIZAR_ATENCION_CLIENTE",
                "tipo": "documento",
                "descripcion": "Informe de optimización de la atención al cliente del proyecto",
                "criterios": [
                    "Presenta un diagnóstico del estado actual de la atención al cliente de la empresa",
                    "Propone un plan de optimización con acciones concretas y plazos",
                    "Describe las herramientas o tecnologías que se usarán para mejorar la atención",
                    "Concluye con indicadores para medir el impacto de la optimización",
                    "Tiene estructura de informe con secciones claramente definidas",
                ],
            },
        ],
    },

    "GUIA_04_POSTVENTA_DIGITAL_SATISFACCION_CLIENTE": {
        "guia": "GUIA_04_POSTVENTA_DIGITAL_SATISFACCION_CLIENTE",
        "programa": "Ventas_de_productos_en_linea",
        "criterio_aprobacion": _CRITERIO_STD,
        "evidencias": [
            {
                "nombre": "04_3_1_CASO_RUNNING_EXPRESS",
                "tipo": "documento",
                "descripcion": "Análisis del caso 'Running Express' con respuestas al cuestionario",
                "criterios": [
                    "Analiza el caso Running Express u otro caso equivalente de postventa digital",
                    "Responde las preguntas del cuestionario con argumentos concretos",
                    "Identifica los errores o aciertos de la empresa en el manejo del postventa",
                    "Propone soluciones basadas en el análisis del caso",
                ],
            },
            {
                "nombre": "04_3_2_IMPORTANCIA_PROTOCOLOS_POSTVENTA",
                "tipo": "documento",
                "descripcion": "Documento sobre la importancia de los protocolos de postventa digital",
                "criterios": [
                    "Define el postventa digital y explica su importancia en el e-commerce",
                    "Describe al menos 2 protocolos de postventa usados en comercio digital",
                    "Relaciona los protocolos con la satisfacción y fidelización del cliente",
                    "El documento tiene argumentos propios desarrollados",
                ],
            },
            {
                "nombre": "04_3_3_1_MARKETING_RELACIONAL_ENTORNO_DIGITAL",
                "tipo": "documento",
                "descripcion": "Informe sobre marketing relacional en el entorno digital",
                "criterios": [
                    "Define el marketing relacional y lo diferencia del marketing transaccional",
                    "Describe estrategias de marketing relacional aplicables al e-commerce",
                    "Analiza cómo el marketing relacional impacta la retención de clientes en línea",
                    "Aplica los conceptos al proyecto de ventas en línea",
                ],
            },
            {
                "nombre": "04_3_3_2_INVESTIGACION_CRM_DIGITAL",
                "tipo": "documento",
                "descripcion": "Cuadro comparativo de plataformas CRM para gestión digital de clientes",
                "criterios": [
                    "Presenta un cuadro comparativo de al menos 3 plataformas CRM",
                    "Compara características como precio, funcionalidades y facilidad de uso",
                    "Recomienda una plataforma CRM justificando la elección para el proyecto",
                    "El cuadro está completo con información de cada plataforma comparada",
                ],
            },
            {
                "nombre": "04_3_3_3_JUEGO_ROLES_MONITOREO_POSTVENTA",
                "tipo": "cualquier",
                "descripcion": "Juego de roles o evidencia de monitoreo postventa digital",
                "criterios": [
                    "Presenta un juego de roles o simulación de seguimiento postventa",
                    "El escenario incluye atención de quejas, devoluciones o seguimiento a la entrega",
                    "Aplica protocolos de postventa digital al escenario planteado",
                    "La evidencia muestra el rol del vendedor en línea en el proceso",
                ],
            },
            {
                "nombre": "04_3_3_4_PRUEBA_CONOCIMIENTOS_GUIA_04",
                "tipo": "cualquier",
                "descripcion": "Prueba de conocimientos sobre postventa digital y satisfacción del cliente",
                "criterios": [
                    "Contiene preguntas sobre postventa digital, CRM y marketing relacional",
                    "Tiene mínimo 5 ítems respondidos",
                    "Las respuestas reflejan comprensión de los temas de la guía",
                    "No está en blanco ni presenta solo el encabezado",
                ],
            },
            {
                "nombre": "04_3_4_INFORME_PROTOCOLO_POSTVENTA",
                "tipo": "documento",
                "descripcion": "Informe del protocolo de postventa digital para la empresa del proyecto",
                "criterios": [
                    "Presenta el protocolo de postventa diseñado para la empresa del proyecto",
                    "Incluye los pasos de seguimiento post-compra (confirmación, entrega, satisfacción)",
                    "Describe los canales digitales usados para el seguimiento postventa",
                    "Concluye con indicadores de satisfacción y métricas de postventa",
                    "Tiene estructura de informe con secciones claramente definidas",
                ],
            },
        ],
    },

    "GUIA_05_PROPUESTA_COMERCIAL_CANALES_DIGITALES": {
        "guia": "GUIA_05_PROPUESTA_COMERCIAL_CANALES_DIGITALES",
        "programa": "Ventas_de_productos_en_linea",
        "criterio_aprobacion": _CRITERIO_STD,
        "evidencias": [
            {
                "nombre": "05_3_1_PRUEBA_DIAGNOSTICA",
                "tipo": "documento",
                "descripcion": "Prueba diagnóstica sobre investigación de mercados digitales",
                "criterios": [
                    "Contiene preguntas diagnósticas sobre investigación de mercados digitales",
                    "Tiene respuestas propias que reflejan el nivel de conocimiento previo del aprendiz",
                    "Cubre temas como mercado digital, competencia en línea y comportamiento del consumidor digital",
                    "No está en blanco ni presenta solo el encabezado",
                ],
            },
            {
                "nombre": "05_3_2_CONTEXTO_ORGANIZACION",
                "tipo": "documento",
                "descripcion": "Folleto sobre el contexto de la organización y su entorno digital",
                "criterios": [
                    "Presenta el contexto de la empresa del proyecto en el entorno digital",
                    "Describe la presencia digital actual de la organización (plataformas, canales, herramientas)",
                    "Identifica oportunidades y desafíos del entorno digital para la empresa",
                    "El folleto tiene diseño visual y contenido organizado",
                ],
            },
            {
                "nombre": "05_3_3_1_CARACTERIZACION_PRODUCTO",
                "tipo": "imagen",
                "descripcion": "Mapa conceptual de caracterización del producto para ventas digitales",
                "criterios": [
                    "Es un mapa conceptual sobre las características del producto del proyecto",
                    "Identifica atributos, beneficios y diferenciadores del producto",
                    "Relaciona las características con la propuesta de valor para el cliente digital",
                    "Tiene estructura de mapa conceptual con nodos y relaciones visibles",
                ],
            },
            {
                "nombre": "05_3_3_2_MAPA_EMPÁTIA_CONSUMIDOR_DIGITAL",
                "tipo": "imagen",
                "descripcion": "Mapa de empatía del consumidor digital del proyecto",
                "criterios": [
                    "Es un mapa de empatía con las secciones: piensa, siente, ve, escucha, dice, hace",
                    "Está orientado al consumidor digital de la empresa del proyecto",
                    "Tiene información relevante en cada sección del mapa",
                    "Es comprensible como herramienta de entendimiento del cliente",
                ],
            },
            {
                "nombre": "05_3_3_3_ARQUETIPO_SEGMENTO_OBJETIVO",
                "tipo": "documento",
                "descripcion": "Documento de arquetipo del segmento objetivo del proyecto",
                "criterios": [
                    "Define el arquetipo o perfil detallado del segmento objetivo",
                    "Describe características demográficas, psicográficas y comportamentales del segmento",
                    "Relaciona el arquetipo con el producto o servicio ofrecido",
                    "El documento tiene contenido desarrollado con argumentos",
                ],
            },
            {
                "nombre": "05_3_3_4_INVESTIGACION_MERCADOS_DIGITALES",
                "tipo": "documento",
                "descripcion": "Investigación básica de mercados digitales para la empresa del proyecto",
                "criterios": [
                    "Presenta una investigación sobre el mercado digital del sector del proyecto",
                    "Analiza la competencia digital (al menos 2 competidores) con sus fortalezas",
                    "Identifica tendencias del mercado digital relevantes para el proyecto",
                    "Concluye con oportunidades de posicionamiento en el mercado digital",
                ],
            },
            {
                "nombre": "05_3_3_5_BUYER_PERSON",
                "tipo": "imagen",
                "descripcion": "Buyer Persona visual del cliente objetivo del proyecto",
                "criterios": [
                    "Es una representación visual del Buyer Persona con datos del perfil",
                    "Incluye nombre ficticio, foto o ilustración, y características clave",
                    "Describe motivaciones, metas y frustraciones del Buyer Persona",
                    "Está orientado al cliente digital de la empresa del proyecto",
                ],
            },
            {
                "nombre": "05_3_3_6_PLAN_MARKETING_DIGITAL",
                "tipo": "imagen",
                "descripcion": "Infografía del plan de marketing digital para el proyecto",
                "criterios": [
                    "Es una infografía visual sobre el plan de marketing digital",
                    "Presenta los elementos del plan: objetivo, canales, contenidos, presupuesto e indicadores",
                    "Tiene diseño visual con iconos, colores y estructura clara",
                    "Es comprensible como síntesis del plan de marketing digital",
                ],
            },
            {
                "nombre": "05_3_3_7_PROPUESTA_COMERCIAL",
                "tipo": "cualquier",
                "descripcion": "Propuesta comercial presentada como presentación y video pitch",
                "criterios": [
                    "Presenta una propuesta comercial formal para la empresa del proyecto",
                    "Incluye presentación y/o video pitch de la propuesta",
                    "La propuesta describe el producto, el valor diferencial y el precio",
                    "El pitch es convincente y está orientado a un cliente o inversionista digital",
                ],
            },
            {
                "nombre": "05_3_4_INFORME_PROPUESTA_COMERCIAL_DIGITAL",
                "tipo": "documento",
                "descripcion": "Informe completo de la propuesta comercial por canales digitales",
                "criterios": [
                    "Presenta la propuesta comercial completa adaptada a los canales digitales",
                    "Incluye el análisis de mercado, el Buyer Persona y el plan de marketing digital",
                    "Describe la estrategia de precios y canales de distribución digital",
                    "Concluye con proyecciones de ventas e indicadores de éxito",
                    "Tiene estructura de informe formal con secciones claramente identificadas",
                ],
            },
        ],
    },

    "GUIA_06_EFECTUAR_VENTA_PRODUCTOS_EN_LINEA": {
        "guia": "GUIA_06_EFECTUAR_VENTA_PRODUCTOS_EN_LINEA",
        "programa": "Ventas_de_productos_en_linea",
        "criterio_aprobacion": _CRITERIO_STD,
        "evidencias": [
            {
                "nombre": "06_3_1_CANALES_CONFIDENCIALIDAD",
                "tipo": "documento",
                "descripcion": "Reflexión sobre el manejo de canales digitales y confidencialidad en ventas en línea",
                "criterios": [
                    "Reflexiona sobre la protección de datos del cliente en ventas en línea",
                    "Identifica riesgos de seguridad en los canales digitales de venta",
                    "Menciona normas o buenas prácticas para garantizar la confidencialidad",
                    "El documento tiene argumentos propios desarrollados",
                ],
            },
            {
                "nombre": "06_3_2_COMERCIO_ONLINE",
                "tipo": "documento",
                "descripcion": "Análisis del comercio online y la omnicanalidad en ventas de productos",
                "criterios": [
                    "Define el comercio online y sus modelos (B2C, B2B, C2C u otros)",
                    "Explica el concepto de omnicanalidad y su importancia en ventas",
                    "Analiza cómo la empresa del proyecto puede implementar la omnicanalidad",
                    "Incluye ejemplos de empresas que usan comercio online exitosamente",
                ],
            },
            {
                "nombre": "06_3_3_1_TALLER_OMNICANALIDAD",
                "tipo": "documento",
                "descripcion": "Taller sobre omnicanalidad con cuadro comparativo y organizador gráfico",
                "criterios": [
                    "Presenta un cuadro comparativo de canales de venta online y offline",
                    "Describe cómo integrar los canales en una estrategia omnicanal",
                    "El organizador gráfico muestra la integración de canales de forma visual",
                    "Aplica el concepto de omnicanalidad al proyecto de ventas en línea",
                ],
            },
            {
                "nombre": "06_3_3_2_GRAFICA_COMUNICACION_DIGITAL",
                "tipo": "imagen",
                "descripcion": "Gráfica sobre estrategias de comunicación digital en ventas en línea",
                "criterios": [
                    "Es una imagen o infografía sobre comunicación digital en ventas",
                    "Presenta estrategias de comunicación por canales digitales",
                    "Diferencia entre tipos de mensajes o estrategias según el canal",
                    "Es comprensible como pieza visual autónoma",
                ],
            },
            {
                "nombre": "06_3_3_3_GRAFICA_OBJECIONES",
                "tipo": "imagen",
                "descripcion": "Gráfica sobre técnicas de manejo de objeciones en ventas digitales",
                "criterios": [
                    "Es una imagen o esquema visual sobre manejo de objeciones en ventas en línea",
                    "Presenta al menos 3 técnicas de respuesta a objeciones frecuentes",
                    "Incluye ejemplos de objeciones típicas en e-commerce y sus respuestas",
                    "Tiene diseño visual claro y organizado",
                ],
            },
            {
                "nombre": "06_3_3_4_DIAGRAMA_PROMOCION_MARKETING_DIGITAL",
                "tipo": "imagen",
                "descripcion": "Diagrama sobre promoción y marketing digital en ventas en línea",
                "criterios": [
                    "Es un diagrama o esquema sobre estrategias de promoción digital",
                    "Muestra la relación entre herramientas de marketing digital (SEO, SEM, redes, email u otras)",
                    "Está orientado a la promoción de productos en línea",
                    "El diagrama es comprensible y tiene estructura lógica",
                ],
            },
            {
                "nombre": "06_3_PRUEBA_EFECTUAR_VENTA_PRODUCTOS_EN_LINEA",
                "tipo": "cualquier",
                "descripcion": "Prueba de conocimientos sobre ventas de productos en línea",
                "criterios": [
                    "Contiene preguntas sobre ventas en línea, omnicanalidad y marketing digital",
                    "Tiene mínimo 5 ítems respondidos",
                    "Las respuestas reflejan comprensión de los temas de la guía",
                    "No está en blanco ni presenta solo el encabezado",
                ],
            },
            {
                "nombre": "06_4_INFORME_VENTA_PRODUCTOS_EN_LINEA",
                "tipo": "documento",
                "descripcion": "Informe del proyecto de venta de productos en línea",
                "criterios": [
                    "Presenta el proyecto completo de venta de productos en línea de la empresa",
                    "Describe la estrategia de omnicanalidad y los canales usados",
                    "Incluye el proceso de venta desde la captación hasta el cierre y postventa",
                    "Presenta resultados o proyecciones de ventas en línea",
                    "Tiene estructura de informe formal con secciones claramente identificadas",
                ],
            },
        ],
    },

    "GUIA_07_APLICAR_ACCIONES_SEGUIMIENTO_CONTROL": {
        "guia": "GUIA_07_APLICAR_ACCIONES_SEGUIMIENTO_CONTROL",
        "programa": "Ventas_de_productos_en_linea",
        "criterio_aprobacion": _CRITERIO_STD,
        "evidencias": [
            {
                "nombre": "07_1_ANALISIS_CASO_SEGUIMIENTO_VENTAS",
                "tipo": "imagen",
                "descripcion": "Análisis del caso 'El regalo prometido' con lluvia de ideas sobre seguimiento de ventas",
                "criterios": [
                    "Analiza el caso 'El regalo prometido' u otro caso equivalente",
                    "La lluvia de ideas presenta múltiples estrategias o soluciones de seguimiento",
                    "Identifica los errores o aprendizajes del caso analizado",
                    "La evidencia es una imagen o esquema visual que muestra el análisis",
                ],
            },
            {
                "nombre": "07_2_TALLER_VENTAS_EFECTIVAS_POSTVENTA_INTELIGENTE",
                "tipo": "imagen",
                "descripcion": "Taller sobre ventas efectivas y postventa inteligente en canales digitales",
                "criterios": [
                    "Presenta el concepto de venta efectiva en el contexto digital",
                    "Define postventa inteligente y sus herramientas (automatización, CRM, seguimiento predictivo u otras)",
                    "Relaciona las ventas efectivas con estrategias de postventa para maximizar la satisfacción",
                    "La evidencia visual muestra el desarrollo del taller",
                ],
            },
            {
                "nombre": "07_3_1_TALLER_CONQUISTANDO_FIDELIZANDO_CLIENTES_ONLINE",
                "tipo": "documento",
                "descripcion": "Taller sobre conquista y fidelización de clientes en línea con caso práctico",
                "criterios": [
                    "Presenta estrategias de conquista de nuevos clientes en canales digitales",
                    "Describe al menos 2 mecanismos de fidelización digital (programa de puntos, email marketing u otros)",
                    "Aplica las estrategias a un caso práctico relacionado con el proyecto",
                    "Incluye reflexión sobre los resultados esperados de fidelización",
                ],
            },
            {
                "nombre": "07_3_2_TALLER_SEGUIMIENTO_TECNICAS_INSTRUMENTOS_REGISTROS",
                "tipo": "documento",
                "descripcion": "Taller sobre seguimiento de ventas con técnicas, instrumentos y registros de control",
                "criterios": [
                    "Define qué es el seguimiento de ventas y su importancia en el comercio digital",
                    "Describe al menos 2 técnicas de seguimiento (llamadas, email, CRM u otras)",
                    "Presenta instrumentos o plantillas de registro de seguimiento",
                    "Aplica las técnicas al proceso de ventas de la empresa del proyecto",
                ],
            },
            {
                "nombre": "07_3_3_TALLER_SIMULACION_DATOS_ANALISIS",
                "tipo": "documento",
                "descripcion": "Taller de simulación con datos y análisis de comportamiento de clientes",
                "criterios": [
                    "Presenta un ejercicio de simulación con datos de clientes o ventas",
                    "Analiza los datos para identificar patrones de comportamiento del cliente",
                    "Extrae conclusiones sobre el desempeño del proceso de ventas",
                    "El análisis incluye tablas, gráficos o representaciones visuales de los datos",
                ],
            },
            {
                "nombre": "07_3_4_TALLER_REPORTE_VENTA_POSTVENTA",
                "tipo": "documento",
                "descripcion": "Taller de elaboración de reporte de venta y postventa para el proyecto",
                "criterios": [
                    "Elabora un reporte de ventas con datos del proyecto (reales o simulados)",
                    "Incluye métricas de postventa: tiempo de respuesta, satisfacción, devoluciones u otras",
                    "El reporte está organizado con gráficos o tablas que facilitan su lectura",
                    "Concluye con recomendaciones de mejora basadas en los datos del reporte",
                ],
            },
            {
                "nombre": "07_4_REPORTE_EFECTIVIDAD_ANALISIS_CAMPANAS",
                "tipo": "documento",
                "descripcion": "Reporte de efectividad y análisis de campañas de ventas en línea del proyecto",
                "criterios": [
                    "Presenta el reporte final de efectividad de las acciones de seguimiento del proyecto",
                    "Analiza las campañas de ventas en línea con datos e indicadores",
                    "Compara los resultados obtenidos con los objetivos planteados",
                    "Propone acciones de control y mejora continua para el proceso de ventas",
                    "Tiene estructura de reporte formal con secciones claramente identificadas",
                ],
            },
        ],
    },

    # ── Asistencia_Comercial (continued below) ─────────────────────────────────

    "Guía_08_Postventa_y_ Satisfacción_del_Cliente": {
        "guia": "Guía_08_Postventa_y_ Satisfacción_del_Cliente",
        "programa": "Asistencia_Comercial",
        "criterio_aprobacion": _CRITERIO_STD,
        "evidencias": [
            {
                "nombre": "08_3_ACTIVIDAD_INFOGRAFIA_SABERES_PREVIOS_POSTVENTA_SAT",
                "tipo": "imagen",
                "descripcion": "Infografía de activación de saberes previos sobre postventa y satisfacción del cliente",
                "criterios": [
                    "Es una infografía con diseño visual sobre postventa o satisfacción del cliente",
                    "Presenta conceptos previos del aprendiz sobre el tema (autodiagnóstico o lluvia de ideas)",
                    "Tiene al menos 4 elementos visuales (iconos, colores, cuadros o flechas)",
                    "Es comprensible como pieza autónoma de información",
                ],
            },
            {
                "nombre": "08_3_ACTIVIDAD_DIDACTICA_REFLEXION_POSTVENTA_SATISFACCI",
                "tipo": "documento",
                "descripcion": "Actividad de reflexión sobre la importancia del servicio postventa en la fidelización de clientes",
                "criterios": [
                    "Define el servicio postventa y explica su importancia para la empresa",
                    "Relaciona el postventa con la fidelización y retención de clientes",
                    "Presenta un ejemplo real o caso hipotético de buen o mal servicio postventa",
                    "Incluye reflexión propia sobre el impacto del postventa en la satisfacción del cliente",
                ],
            },
            {
                "nombre": "08_3_1_ACTIVIDAD_MOMENTOS_CLAVES_CRITICOS_VERDAD",
                "tipo": "documento",
                "descripcion": "Análisis de los momentos de verdad y momentos críticos en la experiencia del cliente",
                "criterios": [
                    "Define el concepto de 'momento de la verdad' en la experiencia del cliente",
                    "Identifica al menos 3 momentos de la verdad en el ciclo de servicio de una empresa",
                    "Distingue entre momentos críticos y momentos ordinarios de la experiencia",
                    "Propone mejoras para convertir momentos críticos en experiencias positivas",
                ],
            },
            {
                "nombre": "08_3_2_ACTIVIDAD_PARA_DESARROLLAR_EN_CLASE",
                "tipo": "documento",
                "descripcion": "Actividad práctica sobre gestión de la satisfacción del cliente en el entorno comercial",
                "criterios": [
                    "Desarrolla la actividad propuesta con respuestas completas y elaboradas",
                    "Aplica conceptos de satisfacción del cliente al contexto de la MIPYME del proyecto",
                    "Identifica indicadores de satisfacción del cliente (NPS, CSAT u otros)",
                    "El documento está completo, no solo tiene el enunciado sin resolver",
                ],
            },
            {
                "nombre": "08_3_3_TALLER_ROLES_POSTVENTA_ACCION_SATISFACCION_CLIEN",
                "tipo": "documento",
                "descripcion": "Taller de juego de roles sobre atención de PQRSF y acciones de postventa",
                "criterios": [
                    "Simula o describe un escenario de atención de PQRSF (Petición, Queja, Reclamo, Sugerencia, Felicitación)",
                    "Aplica un protocolo de atención postventa al escenario",
                    "Propone soluciones concretas al caso planteado en el juego de roles",
                    "Reflexiona sobre las habilidades necesarias para gestionar bien el postventa",
                ],
            },
            {
                "nombre": "08_3_4_COMIC_TRAZABILIDAD_SERVICIO",
                "tipo": "imagen",
                "descripcion": "Cómic o línea de tiempo que ilustra la trazabilidad del servicio al cliente",
                "criterios": [
                    "Es un cómic o representación visual con secuencia narrativa",
                    "Ilustra el recorrido del cliente desde la compra hasta el postventa",
                    "Muestra los puntos de contacto entre empresa y cliente en esa trayectoria",
                    "Tiene al menos 4 viñetas o momentos identificables en la trazabilidad",
                ],
            },
            {
                "nombre": "ACTIVIDAD_ENCUESTA_HERRAMIENTAS_HALLAZGOS_ATENCION_CLIE",
                "tipo": "documento",
                "descripcion": "Encuesta de satisfacción al cliente diseñada y analizada para la empresa del proyecto",
                "criterios": [
                    "Presenta una encuesta diseñada con mínimo 6 preguntas sobre satisfacción del cliente",
                    "Las preguntas cubren aspectos como calidad del producto, servicio y postventa",
                    "Incluye análisis de resultados si se aplicó la encuesta, o proyección de uso si es diseño",
                    "Relaciona los hallazgos o el diseño con las estrategias de mejora de la empresa",
                ],
            },
            {
                "nombre": "08_3_6_DINAMICA_POSTVENTA_SATISFACCION_CLIENTE",
                "tipo": "imagen",
                "descripcion": "Evidencia de participación en dinámica digital sobre postventa y satisfacción del cliente",
                "criterios": [
                    "Es una imagen o captura de pantalla de la actividad completada",
                    "Se puede identificar el nombre del aprendiz o su participación activa",
                    "Muestra resultados, puntaje o progreso de la dinámica",
                    "El contenido está relacionado con postventa o satisfacción del cliente",
                ],
            },
            {
                "nombre": "08_3_7_PRUEBA_DE_CONOCIMIENTOS_GUIA_08",
                "tipo": "cualquier",
                "descripcion": "Prueba de conocimientos sobre postventa y satisfacción del cliente",
                "criterios": [
                    "Contiene preguntas sobre postventa, PQRSF y satisfacción del cliente",
                    "Tiene mínimo 5 ítems respondidos",
                    "Las respuestas reflejan comprensión de los temas de la guía",
                    "No está en blanco ni presenta solo el encabezado",
                ],
            },
            {
                "nombre": "08_4_ELABORAR_PROPUESTA_SISTEMA_PQRSF_ACCION_MEJORA_POS",
                "tipo": "documento",
                "descripcion": "Propuesta formal de sistema de PQRSF con acciones de mejora para la MIPYME del proyecto",
                "criterios": [
                    "Presenta el diseño de un sistema de PQRSF adaptado a la empresa del proyecto",
                    "Describe el proceso de recepción, seguimiento y resolución de PQRSF",
                    "Incluye acciones de mejora continua basadas en los hallazgos del postventa",
                    "Propone indicadores para medir la satisfacción del cliente (NPS, tasa de resolución u otros)",
                    "Tiene estructura de propuesta formal con objetivos, desarrollo y conclusiones",
                ],
            },
        ],
    },
}


# ─────────────────────────────────────────────────────────────────────────────
# FUNCIÓN DE ACCESO
# ─────────────────────────────────────────────────────────────────────────────

def get_evaluador(
    nombre_guia: str,
    programa: str | None = None,
) -> EvaluadorGuia | None:
    """
    Retorna el evaluador para una guía, o None si no está definido.

    Estrategia de resolución (primera que coincide gana):
      1. Exacta            — ``EVALUADORES[nombre_guia]``
      2. Normalizada-exacta — normaliza tildes/mayúsculas/separadores en ambos lados.
      3. Empieza-por       — la clave normalizada comienza por la búsqueda o viceversa.
      4. Contiene          — la búsqueda está contenida en la clave o viceversa.

    Si ``programa`` se especifica y no coincide con el programa registrado,
    se emite un aviso pero se devuelve el evaluador igualmente (no bloquea).

    Devuelve ``None`` (con aviso en stdout) cuando no hay evaluador.
    El caller puede operar en modo solo ENTREGADO/NO ENTREGADO con ese None.
    """
    # ── Paso 1: coincidencia exacta ───────────────────────────────────────────
    evaluador: EvaluadorGuia | None = EVALUADORES.get(nombre_guia)

    # ── Paso 2: coincidencia normalizada (tolera tildes, mayúsculas, guiones) ─
    if evaluador is None:
        norm_busca = _normalizar_clave(nombre_guia)
        claves_norm = {_normalizar_clave(c): c for c in EVALUADORES}

        # 2a. Exacta normalizada
        clave_orig = claves_norm.get(norm_busca)
        if clave_orig:
            evaluador = EVALUADORES[clave_orig]

        # 2b. Empieza-por (normalizado)
        if evaluador is None:
            for norm_clave, clave_orig in claves_norm.items():
                if norm_clave.startswith(norm_busca) or norm_busca.startswith(norm_clave):
                    evaluador = EVALUADORES[clave_orig]
                    break

        # 2c. Contiene (normalizado) — coincidencia más laxa
        if evaluador is None:
            for norm_clave, clave_orig in claves_norm.items():
                if norm_busca in norm_clave or norm_clave in norm_busca:
                    evaluador = EVALUADORES[clave_orig]
                    break

    if evaluador is None:
        print(
            f"[WARN] Sin evaluador para '{nombre_guia}' "
            f"-> solo ENTREGADO/NO ENTREGADO, sin criterios."
        )
        return None

    # ── Validar programa (advertencia laxa, no bloquea) ───────────────────────
    if programa and evaluador.get("programa"):
        prog_eval  = _normalizar_clave(evaluador["programa"])
        prog_busca = _normalizar_clave(programa)
        if prog_busca not in prog_eval and prog_eval not in prog_busca:
            print(
                f"[WARN] Evaluador de '{nombre_guia}' pertenece a "
                f"'{evaluador['programa']}' pero se pidio para '{programa}'. "
                f"Usando igualmente."
            )

    return evaluador


# ─────────────────────────────────────────────────────────────────────────────
# UTILIDADES
# ─────────────────────────────────────────────────────────────────────────────

def listar_guias(programa: str | None = None) -> list[str]:
    """Retorna las claves de guías registradas, opcionalmente filtradas por programa."""
    if not programa:
        return sorted(EVALUADORES.keys())
    norm_prog = _normalizar_clave(programa)
    return sorted(
        k for k, v in EVALUADORES.items()
        if norm_prog in _normalizar_clave(v.get("programa") or "")
    )


def contar_criterios(nombre_guia: str) -> dict:
    """Retorna estadísticas de criterios para una guía (para validación interna)."""
    ev = get_evaluador(nombre_guia)
    if not ev:
        return {}
    evidencias = ev.get("evidencias", [])
    return {
        "evidencias": len(evidencias),
        "criterios_total": sum(len(e.get("criterios", [])) for e in evidencias),
        "criterios_por_evidencia": {
            e["nombre"]: len(e.get("criterios", [])) for e in evidencias
        },
    }


# ─────────────────────────────────────────────────────────────────────────────
# DIAGNÓSTICO RÁPIDO (ejecutar este módulo directamente)
# ─────────────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    import sys

    def _p(s: str) -> None:
        """Print tolerante a encodings estrechos (cp1252, ascii)."""
        try:
            print(s)
        except UnicodeEncodeError:
            print(s.encode(sys.stdout.encoding or "ascii", errors="replace").decode(
                sys.stdout.encoding or "ascii"
            ))

    sep = "=" * 62
    _p(f"\n{sep}")
    _p(f"  EVALUADORES CARGADOS: {len(EVALUADORES)} guias")
    _p(sep)

    by_prog: dict[str, list[str]] = {}
    for guia_key, ev in sorted(EVALUADORES.items()):
        prog = ev.get("programa", "?")
        by_prog.setdefault(prog, []).append(guia_key)

    total_ev = total_cr = 0
    for prog, guias in sorted(by_prog.items()):
        _p(f"\n  [{prog}]")
        for guia_key in guias:
            ev   = EVALUADORES[guia_key]
            n_ev = len(ev["evidencias"])
            n_cr = sum(len(e["criterios"]) for e in ev["evidencias"])
            total_ev += n_ev
            total_cr += n_cr
            _p(f"    {guia_key}  ({n_ev} ev / {n_cr} cr)")

    _p(f"\n{sep}")
    _p(f"  TOTAL: {len(EVALUADORES)} guias | {total_ev} evidencias | {total_cr} criterios")
    _p(sep)

    # Prueba de acceso exacta y normalizada
    casos = [
        ("Guia_01_Diagnostico_Empresarial", "Asistencia_Comercial"),  # sin tildes
        ("GUIA_09_MONITOREAR_METRICAS",     "Comunicacion_y_marketing"),
        ("guia_07_aplicar_acciones",        "Ventas_de_productos_en_linea"),
    ]
    _p("\n  -- Prueba de resolucion --")
    for nombre, prog in casos:
        res = get_evaluador(nombre, prog)
        estado = f"OK  -> '{res['guia']}' ({len(res['evidencias'])} ev)" if res else "FALLO"
        _p(f"  get_evaluador({nombre!r}) : {estado}")
