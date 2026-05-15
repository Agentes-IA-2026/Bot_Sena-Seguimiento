# CLAUDE.md — GuiaBot Panel Docente

## Identidad del proyecto
GuiaBot es un panel de seguimiento de evidencias para instructores del SENA.
El frontend principal está en `dashboard.html` (HTML + JS embebido), y los datos
vienen de Supabase (tabla `verificaciones` y otras tablas auxiliares).

## Archivos clave
- `dashboard.html` → lógica de fichas, guías, evidencias, modal de aprendiz,
  vistas Resumen, Alertas, Comparativo, Acumulado y Auditoría.
- Carpeta `scripts/` y `api.py` → scripts de soporte (fixes, auditoría, etc.).

## Bugs ya corregidos (no repetir)
- Deduplicación de guías usando `_normGuia()`.
- Regex Unicode en `_normGuia` usando el rango `\u0300-\u036f` en lugar de literales.
- Deduplicación de evidencias en el modal usando `_normEvidencia()`.
- Eliminación de `innerHTML +=` en renders, usando arrays + `.join('')`.
- Race condition en `cargarDatos()` controlada con `_loadId`.
- Apertura del modal protegida con `try/catch`, checks de `null` y `modal.classList.add('open')` fuera del `try`.

## Bug pendiente
- Confirmar que el modal de aprendiz (por ejemplo, Ordoñez Martínez Luna Valentina
  en la ficha 3415087) muestra:
  - número correcto de evidencias por guía (sin duplicados),
  - porcentaje de cumplimiento coherente con lo registrado en Supabase.

## Regla de trabajo con Claude Code
- Siempre revisar la consola del navegador (F12) y pegar el error exacto antes de
  pedir un fix grande.
- No volver a redefinir `_normGuia()` ni `_normEvidencia()`, solo ajustarlas si
  hay nuevos casos de nombres.

  ## Reglas de evidencias (NO modificar sin revisión)

### Dashboard (dashboard.html) — ESTABLE, no tocar
- Toda la deduplicación de evidencias usa `_normEvidencia()` y `_normGuia()` (JS).
- `agruparAprendicesPorGuia()` agrupa por clave normalizada con OR de entregado.
- `renderTablaEvidencias()` construye columnas con `_evMap` (normKey → mejor label).
- `agruparPorAprendizGeneral()` usa clave compuesta `_normGuia(guia)+'|||'+_normEvidencia(evidencia)`.
- Ningún uso de `r.evidencia` como clave cruda de objeto; siempre pasa por `_normEvidencia`.

### Bot y auditoría (Python) — Reglas fijas
- Toda comparación de guías y evidencias usa `_norm_guia_py()` y `_norm_evidencia_py()` 
  (normalización NFD + eliminación de tildes + stopwords alineados con JS).
- `bot/auditor.py` y `bot/drive_adapter.py` usan `unicodedata.normalize('NFD')` 
  para eliminar marcas combinantes; nunca usar `str.maketrans` con caracteres literales.
- `_resolver_archivos_para_guia()` devuelve TODOS los archivos del aprendiz; 
  NO aplicar pre-filtros por subcarpeta antes del motor de auditoría.
  La lógica de carpetas se decide dentro del motor usando `carpeta_drive` y 
  estados como CARPETA_INCORRECTA (que `es_entregada()` trata como entregado).
- `reporte_bool` se construye con:
  `_clave = r.get("nombre_esperado") or r.get("actividad_resumen") or ""`
  Nunca filtrar con `if r.get("nombre_esperado")` porque descarta evidencias 
  encontradas por fallback de `actividad_resumen` (ej. "Mi programa de Formación").
- `es_entregada(r)` cubre OK, FORMATO_INCORRECTO, NOMBRE_DIFERENTE y CARPETA_INCORRECTA.
  No usar `r["estado"] == EstadoAuditoria.OK.value` directamente.

### Umbrales de matching con Drive
- CANDIDATO: score >= 50
- OK (evidencia encontrada): score >= 75
- Si score >= 75 para cualquier archivo de Drive, la evidencia se considera existente.
- No generar "no existe la evidencia" cuando hay un candidato con score >= 75.

### Caso resuelto documentado
- "Mi programa de Formación": fallaba porque `nombre_esperado = NULL` en Supabase
  hacía que `reporte_bool` descartara silenciosamente el resultado del motor.
  Fix: usar `actividad_resumen` como clave de fallback en `reporte_bool`.