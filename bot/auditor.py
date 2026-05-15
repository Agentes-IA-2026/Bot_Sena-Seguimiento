"""
auditor.py  —  Motor de auditoría enriquecida (Fase 2)

Compara actividades_parametrizadas (de Supabase) contra archivos en Drive.
Retorna un estado detallado por cada actividad.

Estados:
    OK                — nombre + carpeta + tipo correctos
    NOMBRE_DIFERENTE  — archivo encontrado, nombre distinto al esperado
    CARPETA_INCORRECTA— archivo encontrado, carpeta equivocada
    FORMATO_INCORRECTO— archivo encontrado, extensión fuera del tipo esperado
    FALTA             — actividad obligatoria sin archivo candidato
    OPCIONAL_FALTANTE — actividad no obligatoria sin candidato
    EVIDENCIA_CRUZADA — sin match en guía actual, pero archivo coincide fuerte
                        con actividad de otra guía/programa

"DUPLICADO_POSIBLE" aparece en observaciones cuando hay más de un candidato.
"""

import re
import unicodedata
from enum import Enum


# ── ESTADOS ───────────────────────────────────────────────────────────────────

class EstadoAuditoria(str, Enum):
    OK                 = "OK"
    NOMBRE_DIFERENTE   = "NOMBRE_DIFERENTE"
    CARPETA_INCORRECTA = "CARPETA_INCORRECTA"
    FORMATO_INCORRECTO = "FORMATO_INCORRECTO"
    FALTA              = "FALTA"
    OPCIONAL_FALTANTE  = "OPCIONAL_FALTANTE"
    EVIDENCIA_CRUZADA  = "EVIDENCIA_CRUZADA"


# ── TIPOS DE ARCHIVO ──────────────────────────────────────────────────────────

GRUPOS_TIPO: dict[str, set[str] | None] = {
    "CUALQUIER_FORMATO"           : None,
    "CUALQUIER_FORMATO_DOCUMENTO" : {".pdf",".docx",".doc",".odt",".txt",".rtf",
                                      ".xlsx",".xls",".pptx",".ppt"},
    "CUALQUIER_FORMATO_IMAGEN"    : {".jpg",".jpeg",".png",".gif",".webp",
                                      ".bmp",".svg",".heic",".heif"},
    "CUALQUIER_FORMATO_VIDEO"     : {".mp4",".avi",".mov",".mkv",".wmv",".flv",".webm"},
    "CUALQUIER_FORMATO_AUDIO"     : {".mp3",".wav",".ogg",".aac",".m4a",".flac"},
    "PDF"                         : {".pdf"},
    "EXCEL"                       : {".xlsx",".xls",".csv"},
    "WORD"                        : {".docx",".doc",".odt"},
    "IMAGEN"                      : {".jpg",".jpeg",".png",".gif",".webp"},
}

# mimeTypes de Google Apps que no tienen extensión de archivo
_GOOGLE_APPS_DOCUMENTO = frozenset({
    "application/vnd.google-apps.document",
    "application/vnd.google-apps.spreadsheet",
    "application/vnd.google-apps.presentation",
    "application/vnd.google-apps.form",
    "application/vnd.google-apps.drawing",
})


# ── NORMALIZACIÓN ÚNICA ───────────────────────────────────────────────────────
# Una sola función reutilizable para nombres de archivo Y nombres de carpeta.
# Usa NFD + strip combining marks en vez de tabla de transliteración, para
# manejar correctamente nombres en NFC y NFD (Google Drive puede devolver ambas).

_RE_EXT  = re.compile(
    r'\.(pdf|png|jpg|jpeg|xlsx|xls|pptx|ppt|docx|doc|mp4|avi|mov|mp3|wav|ogg|webm)(\.|$)',
    re.IGNORECASE,
)
# Sufijos que Drive añade: "(1)", "(2)", "_copia", "- Copy", "_final", "_borrador", etc.
_RE_COPY = re.compile(
    r'\s*[\(\[]\d+[\)\]]'
    r'|\s*[-_]\s*(v\d+|copia|copy|final|borrador)'
    r'|\s+copia\b',
    re.IGNORECASE,
)
_RE_NOAK = re.compile(r'[^a-z0-9\s]')
_RE_SPC  = re.compile(r'\s+')


def normalizar(texto: str) -> str:
    """
    Normalización canónica para nombres de archivo Y nombres de carpeta.
    Aplica: minúsculas → NFD + strip combining marks (tildes, diacríticos)
            → quitar extensión → quitar sufijos de copia
            → guiones/underscores/puntos → espacio → colapsar espacios.
    NFD es más robusto que una tabla de transliteración: maneja
    correctamente tanto NFC como NFD, incluidos nombres de archivo devueltos
    por la API de Drive que puedan venir descompuestos.
    """
    t = texto.lower()
    t = unicodedata.normalize('NFD', t)
    t = ''.join(c for c in t if unicodedata.category(c) != 'Mn')
    t = _RE_EXT.sub(' ', t)
    t = _RE_COPY.sub(' ', t)
    t = _RE_NOAK.sub(' ', t)
    return _RE_SPC.sub(' ', t).strip()


# Alias corto usado internamente
_norm = normalizar


_STOP = frozenset({
    'pdf','png','jpg','xlsx','pptx','docx','doc',
    # preposiciones y artículos españoles que generan ruido en el matching
    'de','del','la','el','los','las','un','una','uno','y','o','a',
    'en','con','por','al','que','se','su','sus','lo',
    'para','como','the','and','sena',
})


def _tokens(norma: str) -> frozenset[str]:
    toks = frozenset(p for p in norma.split() if len(p) > 3 and p not in _STOP)
    if not toks:
        # Fallback: evidencias cortas (p.ej. "Bit", "Mapa") donde todos los
        # tokens tienen <= 3 chars — relajar a len > 1 para no perder el patrón.
        toks = frozenset(p for p in norma.split() if len(p) > 1 and p not in _STOP)
    return toks


def _singularizar(p: str) -> str:
    if len(p) <= 4:
        return p
    if p.endswith('es') and len(p) > 4:
        return p[:-2]
    if p.endswith(('as', 'os', 's')) and len(p) > 3:
        return p[:-1]
    return p


def _bigramas(tok: frozenset[str]) -> frozenset[str]:
    lst = sorted(tok)
    return frozenset(f"{lst[i]} {lst[i+1]}" for i in range(len(lst) - 1))


# ── MOTOR DE PUNTUACIÓN ───────────────────────────────────────────────────────

def _score_patron(patron_norm: str, drive_norm: str) -> int:
    """
    Puntuación de similitud entre dos nombres ya normalizados.

        100 — exacto
         80 — contención literal O todos los tokens del patrón en el archivo
         75 — reverse containment: todos los tokens del archivo en el patrón
              (archivo con nombre más corto que el patrón esperado)
         70 — 70%+ tokens coinciden
         65 — bigrama común
         60 — 50%+ singulares coinciden
         55 — 50%+ tokens coinciden
    """
    if not patron_norm:
        return 0

    if patron_norm == drive_norm:
        return 100

    if patron_norm in drive_norm or drive_norm in patron_norm:
        return 80

    tok_p = _tokens(patron_norm)
    tok_d = _tokens(drive_norm)
    if not tok_p or not tok_d:
        return 0

    if tok_p <= tok_d:              # todos los tokens del patrón están en el archivo
        return 80
    if tok_d and tok_d <= tok_p:    # reverse: archivo es subconjunto del patrón
        return 75

    comunes = tok_p & tok_d
    pct = len(comunes) / len(tok_p)

    if pct >= 0.70:
        return 70
    if len(tok_p) >= 2 and len(tok_d) >= 2 and _bigramas(tok_p) & _bigramas(tok_d):
        return 65
    raices_p = frozenset(_singularizar(t) for t in tok_p)
    raices_d = frozenset(_singularizar(t) for t in tok_d)
    if raices_p and len(raices_p & raices_d) / len(raices_p) >= 0.50:
        return 60
    if pct >= 0.50:
        return 55

    return 0


def _mejor_score(nombre_esp: str, variantes: list[str], nombre_drive: str) -> tuple[int, str]:
    """
    Prueba nombre_esperado y cada variante.
    Retorna (score_máximo, patrón_ganador).
    Las variantes descuentan 5 puntos para preferir el nombre canónico.
    """
    drive_norm = _norm(nombre_drive)
    best_score, best_pat = 0, ""
    for patron, descuento in [(nombre_esp, 0)] + [(v, 5) for v in variantes if v]:
        s = _score_patron(_norm(patron), drive_norm)
        if s > 0:
            s = max(s - descuento, 1)
        if s > best_score:
            best_score, best_pat = s, patron
    return best_score, best_pat


# ── VALIDACIONES DE TIPO Y CARPETA ────────────────────────────────────────────

def _tipo_ok(tipo: str, ext: str, mime_type: str = "") -> bool:
    grupo = GRUPOS_TIPO.get((tipo or "CUALQUIER_FORMATO").upper())
    if grupo is None:
        return True
    # Google Apps sin extensión → aceptar como documento
    if not ext and mime_type in _GOOGLE_APPS_DOCUMENTO:
        return tipo.upper() in ("CUALQUIER_FORMATO", "CUALQUIER_FORMATO_DOCUMENTO")
    return ext.lower() in grupo


def _carpeta_ok(esperada: str | None, folder_path: str, carpeta_padre: str = "") -> bool:
    """
    Verifica si el archivo está en la carpeta esperada.
    Usa normalizar() para tolerancia cosmética:
        "2. Análisis", "2_ANALISIS", "2 analisis" → todos equivalentes.
    Comprueba folder_path (ruta completa) y carpeta_padre (nombre inmediato).
    """
    if not esperada:
        return True
    esp_norm = _norm(esperada)
    return (
        esp_norm in _norm(folder_path)
        or (carpeta_padre and esp_norm in _norm(carpeta_padre))
    )


# ── UMBRALES ──────────────────────────────────────────────────────────────────

UMBRAL_MATCH   = 50   # score mínimo para ser candidato
UMBRAL_OK      = 75   # score mínimo para OK (antes 80; bajado para cubrir reverse containment)
UMBRAL_CRUZADO = 80   # umbral estricto para EVIDENCIA_CRUZADA


# ── HELPERS INTERNOS ──────────────────────────────────────────────────────────

def _pick(candidatos: list) -> tuple[int, str, dict]:
    return candidatos[0]   # ya están sorted desc


def _obs_dup(candidatos: list) -> str:
    if len(candidatos) > 1:
        return "DUPLICADO_POSIBLE: " + ", ".join(c[2]["nombre"] for c in candidatos[1:3])
    return ""


def _log_resultado(act_id, fuente_tag, nombre_esp, score, archivo, estado, obs=""):
    """Línea de resultado siempre visible por actividad."""
    ext_r  = archivo.get("extension") or (archivo.get("mime_type") or "")[:25]
    carp_r = archivo.get("folder_path") or archivo.get("carpeta_padre") or "(raíz)"
    icons  = {
        "OK"                 : "✓ ",
        "NOMBRE_DIFERENTE"   : "~ ",
        "CARPETA_INCORRECTA" : "📁",
        "FORMATO_INCORRECTO" : "📄",
        "FALTA"              : "✗ ",
        "OPCIONAL_FALTANTE"  : "○ ",
        "EVIDENCIA_CRUZADA"  : "↔ ",
    }
    icon = icons.get(estado, "· ")
    print(
        f"         {icon}[{act_id}] patron='{fuente_tag}{nombre_esp[:45]}'\n"
        f"               ↳ {score:3d}  '{archivo['nombre']}'  [{ext_r}]  /{carp_r}\n"
        f"               → {estado}{(' | '+obs[:60]) if obs else ''}"
    )


def _log_sin_candidato(act_id, fuente_tag, nombre_esp, archivos, nombre_esp_norm):
    """Log cuando no hay ningún candidato; muestra top 5 scores para diagnóstico."""
    top = sorted(
        [(_score_patron(nombre_esp_norm, _norm(a["nombre"])), a["nombre"])
         for a in archivos],
        reverse=True,
    )[:5]
    scores_str = "  |  ".join(f"{s} '{n[:28]}'" for s, n in top if s > 0)
    print(
        f"         ✗ [{act_id}] patron='{fuente_tag}{nombre_esp[:45]}'\n"
        f"               → 0 candidatos (umbral={UMBRAL_MATCH})"
        + (f"\n               top scores: {scores_str}" if scores_str else "  ningún score > 0")
    )


# ── MOTOR PRINCIPAL ───────────────────────────────────────────────────────────

class MotorAuditoria:
    """
    Compara actividades parametrizadas contra archivos en Drive.

    Flujo de _evaluar (5 pasos):
      1. Fallback a actividad_resumen si nombre_esperado vacío.
      2. Candidatos por nombre (score >= UMBRAL_MATCH).
         Sin candidatos → EVIDENCIA_CRUZADA o FALTA.
      3. Preferir candidatos con tipo correcto.
      4. Preferir candidatos en carpeta correcta (no es un veto: es desempate).
         Si todos fuera de carpeta → CARPETA_INCORRECTA.
      5. Clasificar: score >= UMBRAL_OK → OK, else NOMBRE_DIFERENTE.
    """

    def __init__(
        self,
        actividades: list[dict],
        archivos: list[dict],
        todas_actividades: list[dict] | None = None,
        debug: bool = False,
    ):
        self.actividades = actividades
        self.archivos    = archivos
        self.debug       = debug

        _pares = {(a.get("programa"), a.get("guia")) for a in actividades}
        self._otras_acts: list[dict] = [
            a for a in (todas_actividades or [])
            if (a.get("programa"), a.get("guia")) not in _pares
            and a.get("nombre_esperado")
        ]

    def auditar(self) -> list[dict]:
        return [self._evaluar(act) for act in self.actividades]

    def _evaluar(self, act: dict) -> dict:  # noqa: C901
        nombre_esp  = (act.get("nombre_esperado") or "").strip()
        resumen     = (act.get("actividad_resumen") or "").strip()
        variantes   = act.get("variantes_permitidas") or []
        tipo        = act.get("tipo_archivo") or "CUALQUIER_FORMATO"
        obligatoria = act.get("obligatoria", True)
        carpeta_esp = act.get("carpeta_drive")
        act_id      = act.get("actividad_id", "?")

        # ── Paso 1: fallback si nombre_esperado vacío ────────────────────────
        usar_fallback = False
        if not nombre_esp:
            if resumen:
                nombre_esp    = resumen
                usar_fallback = True
                print(f"         ⚠️  [{act_id}] nombre_esperado VACÍO → fallback: '{resumen[:55]}'")
            else:
                obs = "nombre_esperado y actividad_resumen vacíos en Supabase"
                print(f"         ❌ [{act_id}] SIN PATRÓN — {obs}")
                estado = EstadoAuditoria.FALTA if obligatoria else EstadoAuditoria.OPCIONAL_FALTANTE
                return _mk(act, estado, None, 0, "", obs)

        fuente_tag   = "[fallback] " if usar_fallback else ""
        patron_norm  = _norm(nombre_esp)
        patron_toks  = _tokens(patron_norm)

        # ── Debug dirigido: evidencia "Mi programa de Formación" ─────────────
        _dbg = any(kw in patron_norm for kw in ('programa', 'formacion'))
        if _dbg:
            print(
                f"\nDEBUG_MI_PROGRAMA"
                f"  raw={repr(nombre_esp)}"
                f"  norm={repr(patron_norm)}"
                f"  toks={sorted(patron_toks)}"
                f"  variantes={variantes}"
            )

        if self.debug:
            print(f"\n         🔬 [{act_id}] patron='{fuente_tag}{nombre_esp[:55]}'")
            print(f"              norm='{patron_norm[:55]}'  tokens={sorted(patron_toks)}")
            print(f"              tipo={tipo}  carpeta_esp={carpeta_esp}")

        # ── Paso 2: candidatos por nombre ────────────────────────────────────
        candidatos: list[tuple[int, str, dict]] = []
        for archivo in self.archivos:
            score, patron = _mejor_score(nombre_esp, variantes, archivo["nombre"])
            _dn = _norm(archivo["nombre"])
            _dbg_drive = any(kw in _dn for kw in ('programa', 'formacion'))
            # Log si el PATRÓN buscado es la evidencia objetivo, O si el ARCHIVO de
            # Drive contiene las palabras clave — cubre el caso donde nombre_esperado
            # en actividades_parametrizadas tiene otro nombre pero el archivo coincide.
            if _dbg or _dbg_drive:
                print(
                    f"DEBUG_DRIVE_MI_PROGRAMA"
                    f"  score={score:3d}"
                    f"  patron_norm={repr(patron_norm)}"
                    f"  raw_drive={repr(archivo['nombre'])}"
                    f"  norm_drive={repr(_dn)}"
                )
            if self.debug and score > 0:
                drive_norm = _norm(archivo["nombre"])
                drive_toks = _tokens(drive_norm)
                ext_i  = archivo.get("extension") or (archivo.get("mime_type") or "")[:25]
                carp_i = archivo.get("folder_path") or "(raíz)"
                print(f"            ↳ {score:3d}  '{archivo['nombre']}'"
                      f"  [{ext_i}]  /{carp_i}"
                      f"  (norm='{drive_norm[:35]}'  tokens={sorted(drive_toks)})")
            if score >= UMBRAL_MATCH:
                candidatos.append((score, patron, archivo))

        # ── Sin candidatos ───────────────────────────────────────────────────
        if not candidatos:
            _log_sin_candidato(act_id, fuente_tag, nombre_esp, self.archivos, patron_norm)
            cruzada = self._buscar_cruzada(act)
            if cruzada:
                return cruzada
            estado = EstadoAuditoria.FALTA if obligatoria else EstadoAuditoria.OPCIONAL_FALTANTE
            return _mk(act, estado, None, 0, "", "")

        candidatos.sort(key=lambda x: x[0], reverse=True)

        # ── Paso 3: preferir tipo correcto ───────────────────────────────────
        tipo_ok_list  = [(s, p, f) for s, p, f in candidatos
                         if _tipo_ok(tipo, f.get("extension", ""), f.get("mime_type", ""))]
        formato_malo  = not tipo_ok_list
        pool          = tipo_ok_list if tipo_ok_list else candidatos

        # ── Paso 4: preferir carpeta correcta (señal de desempate, no veto) ──
        #
        # REGLA DE VALIDEZ: antes de aplicar el filtro de carpeta, verificar
        # que la carpeta esperada exista en ALGÚN archivo del portafolio.
        # Si carpeta_drive no aparece en ningún archivo del Drive, el valor
        # en Supabase es incorrecto (ej: "Portafolio del Aprendiz" cuando los
        # archivos están en "Análisis"). En ese caso se omite el check.
        if carpeta_esp:
            carpeta_existe = any(
                _carpeta_ok(carpeta_esp,
                             a.get("folder_path", ""),
                             a.get("carpeta_padre", ""))
                for a in self.archivos
            )
            if not carpeta_existe:
                print(f"         ℹ️  [{act_id}] carpeta_drive='{carpeta_esp}' "
                      f"no existe en Drive → check omitido (actualiza carpeta_drive en Supabase)")
                en_carpeta    = pool
                fuera_carpeta = []
            else:
                en_carpeta    = [(s, p, f) for s, p, f in pool
                                 if _carpeta_ok(carpeta_esp,
                                                f.get("folder_path", ""),
                                                f.get("carpeta_padre", ""))]
                fuera_carpeta = [(s, p, f) for s, p, f in pool if (s, p, f) not in en_carpeta]
        else:
            en_carpeta    = pool
            fuera_carpeta = []

        # Carpeta esperada existe en Drive pero ningún candidato está en ella
        if carpeta_esp and fuera_carpeta and not en_carpeta:
            score, patron, archivo = _pick(fuera_carpeta)
            obs = _obs_dup(pool)
            estado = (EstadoAuditoria.FORMATO_INCORRECTO if formato_malo
                      else EstadoAuditoria.CARPETA_INCORRECTA)
            real_carp = archivo.get("folder_path") or archivo.get("carpeta_padre") or "(raíz)"
            obs_carp  = f"esperada='{carpeta_esp}' real='{real_carp}'"
            _log_resultado(act_id, fuente_tag, nombre_esp, score, archivo, estado.value,
                           obs_carp + (f" | {obs}" if obs else ""))
            return _mk(act, estado, archivo, score, patron, obs_carp)

        # Candidatos válidos
        score, patron, archivo = _pick(en_carpeta)
        obs = _obs_dup(en_carpeta + fuera_carpeta)

        if formato_malo:
            # Formato ≠ tipo esperado → advertencia, pero el archivo cuenta como ENTREGADO.
            # Se enriquece obs con el tipo esperado y la extensión real para surfacearlo
            # en el reporte Excel y en logs, sin cambiar el estado de entrega.
            ext_real  = (archivo.get("extension") or "").lstrip(".") or (
                (archivo.get("mime_type") or "").split("/")[-1][:12]
            ) or "sin-ext"
            tipo_norm = (tipo or "CUALQUIER_FORMATO").upper()
            obs_fmt   = f"formato: esperado '{tipo_norm}', encontrado '{ext_real}'"
            obs_full  = (obs_fmt + " | " + obs) if obs else obs_fmt
            _log_resultado(act_id, fuente_tag, nombre_esp, score, archivo,
                           EstadoAuditoria.FORMATO_INCORRECTO.value, obs_full)
            return _mk(act, EstadoAuditoria.FORMATO_INCORRECTO, archivo, score, patron, obs_full)

        # ── Paso 5: calidad del nombre ───────────────────────────────────────
        estado = EstadoAuditoria.OK if score >= UMBRAL_OK else EstadoAuditoria.NOMBRE_DIFERENTE
        _log_resultado(act_id, fuente_tag, nombre_esp, score, archivo, estado.value, obs)
        return _mk(act, estado, archivo, score, patron, obs)

    def _buscar_cruzada(self, act: dict) -> dict | None:
        """
        Para una actividad sin candidatos en la guía actual, busca si algún archivo
        del aprendiz coincide fuertemente con una actividad de otra guía/programa.
        Solo dispara con score >= UMBRAL_CRUZADO (80) para evitar falsos positivos.
        """
        if not self._otras_acts:
            return None

        mejor_score, mejor_archivo, mejor_otra = 0, None, None
        for otra in self._otras_acts:
            n = otra.get("nombre_esperado") or ""
            v = otra.get("variantes_permitidas") or []
            for archivo in self.archivos:
                s, _ = _mejor_score(n, v, archivo["nombre"])
                if s >= UMBRAL_CRUZADO and s > mejor_score:
                    mejor_score, mejor_archivo, mejor_otra = s, archivo, otra

        if not mejor_archivo:
            return None

        obs = (
            f"'{mejor_archivo['nombre']}' coincide con "
            f"'{mejor_otra.get('nombre_esperado')}' "
            f"de {mejor_otra.get('guia')} / {mejor_otra.get('programa')} "
            f"(confianza {mejor_score}%). "
            f"Esperado aquí: '{act.get('nombre_esperado')}'"
        )
        _log_resultado(
            act.get("actividad_id","?"), "", act.get("nombre_esperado",""),
            mejor_score, mejor_archivo, EstadoAuditoria.EVIDENCIA_CRUZADA.value, obs
        )
        return _mk(act, EstadoAuditoria.EVIDENCIA_CRUZADA, mejor_archivo,
                   mejor_score, mejor_otra.get("nombre_esperado",""), obs)


# ── RESULTADO ─────────────────────────────────────────────────────────────────

def _mk(act, estado, archivo, confianza, patron, obs):
    return {
        "actividad_id"      : act.get("actividad_id"),
        "guia"              : act.get("guia"),
        "programa"          : act.get("programa"),
        "actividad_resumen" : act.get("actividad_resumen"),
        "nombre_esperado"   : act.get("nombre_esperado"),
        "obligatoria"       : act.get("obligatoria", True),
        "estado"            : estado.value if hasattr(estado, "value") else estado,
        "archivo_encontrado": archivo["nombre"]          if archivo else None,
        "carpeta_encontrada": archivo.get("folder_path") if archivo else None,
        "extension"         : archivo.get("extension")   if archivo else None,
        "confianza"         : confianza,
        "patron_match"      : patron or None,
        "observaciones"     : obs    or None,
    }


# ── HELPERS PARA CALLERS ──────────────────────────────────────────────────────

def es_entregada(resultado: dict) -> bool:
    return resultado["estado"] in (
        EstadoAuditoria.OK.value,
        EstadoAuditoria.NOMBRE_DIFERENTE.value,
        EstadoAuditoria.CARPETA_INCORRECTA.value,
        EstadoAuditoria.FORMATO_INCORRECTO.value,
    )


def resumen(resultados: list[dict]) -> dict:
    total      = len(resultados)
    ok         = sum(1 for r in resultados if r["estado"] == EstadoAuditoria.OK.value)
    encontradas = sum(1 for r in resultados if es_entregada(r))
    cruzadas   = sum(1 for r in resultados if r["estado"] == EstadoAuditoria.EVIDENCIA_CRUZADA.value)
    return {
        "total"      : total,
        "ok"         : ok,
        "encontradas": encontradas,
        "cruzadas"   : cruzadas,
        "faltantes"  : sum(1 for r in resultados if r["estado"] == EstadoAuditoria.FALTA.value),
        "porcentaje" : round((ok / total * 100) if total else 0),
    }
