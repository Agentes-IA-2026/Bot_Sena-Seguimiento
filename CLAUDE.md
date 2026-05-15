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