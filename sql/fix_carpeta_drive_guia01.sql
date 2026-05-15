-- =============================================================
-- fix_carpeta_drive_guia01.sql
-- Causa raíz confirmada de 0/11 en Guía_01_Diagnóstico_Empresarial:
--
-- La columna carpeta_drive contiene valores descriptivos que NO
-- coinciden con las carpetas reales del portafolio del aprendiz.
-- El matcher encuentra los archivos por nombre (score 80-100),
-- pero el check de carpeta falla para todos → CARPETA_INCORRECTA → 0 OK.
--
-- Valores actuales (incorrectos):
--   "Portafolio del Aprendiz"
--   "GUIA_01_DIAGNÓSTICO_EMPRESARIAL"
--   "Actividades de Transferencia"
--   "Según indicación Instructor"
--
-- Ejecutar en Supabase → SQL Editor
-- =============================================================

-- OPCIÓN A (recomendada si los archivos están en carpeta "Análisis"):
-- Fijar carpeta_drive al nombre real de la carpeta del aprendiz.
UPDATE actividades_parametrizadas
SET carpeta_drive = 'Análisis'
WHERE guia = 'Guía_01_Diagnóstico_Empresarial';

-- OPCIÓN B (si los archivos están directamente en la raíz del portafolio):
-- Poner NULL para que el motor acepte cualquier carpeta.
-- UPDATE actividades_parametrizadas
-- SET carpeta_drive = NULL
-- WHERE guia = 'Guía_01_Diagnóstico_Empresarial';

-- Verificar resultado:
SELECT actividad_id, nombre_esperado, carpeta_drive
FROM actividades_parametrizadas
WHERE guia = 'Guía_01_Diagnóstico_Empresarial'
ORDER BY actividad_id;

-- =============================================================
-- Si el problema se repite en otras guías, usar esta consulta
-- para ver qué carpeta_drive tiene cada guía vs qué carpetas
-- existen realmente en Drive (revisar consola del bot).
-- =============================================================
SELECT guia, carpeta_drive, count(*) as actividades
FROM actividades_parametrizadas
WHERE activa = true AND carpeta_drive IS NOT NULL
GROUP BY guia, carpeta_drive
ORDER BY guia, carpeta_drive;
