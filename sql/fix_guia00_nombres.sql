-- =============================================================
-- fix_guia00_nombres.sql
-- Ejecutar en Supabase → SQL Editor
--
-- Corrige nombre_esperado y variantes_permitidas para GUIA_00_INDUCCION
-- para que el motor de auditoría marque OK cuando los archivos reales
-- del aprendiz coinciden semánticamente con la actividad.
--
-- También limpia carpeta_drive (los archivos de inducción suelen estar
-- en la raíz del portafolio, sin subcarpeta específica).
-- =============================================================


-- ── PASO 0: limpiar carpeta_drive ────────────────────────────────────────────
-- Los archivos de GUIA_00 suelen estar en la raíz del portafolio del aprendiz.
-- Sin carpeta_drive definida, el motor no aplica validación de carpeta.
UPDATE actividades_parametrizadas
SET carpeta_drive = NULL
WHERE guia = 'GUIA_00_INDUCCION';


-- ── PASO 1: Quién soy (nube de palabras) ─────────────────────────────────────
UPDATE actividades_parametrizadas SET
    nombre_esperado      = 'Quien soy',
    variantes_permitidas = ARRAY[
        'quien soy',
        'quién soy',
        'quien_soy',
        'nube de palabras quien soy',
        'nube palabras'
    ]
WHERE guia = 'GUIA_00_INDUCCION'
  AND (
        actividad_resumen ILIKE '%qui_n soy%'
     OR actividad_resumen ILIKE '%quien soy%'
     OR actividad_resumen ILIKE '%nube%palabras%'
  );


-- ── PASO 2: Estilo de aprendizaje ────────────────────────────────────────────
UPDATE actividades_parametrizadas SET
    nombre_esperado      = 'Mi estilo de aprendizaje',
    variantes_permitidas = ARRAY[
        'mi estilo de aprendizaje',
        'estilo de aprendizaje',
        'test kolb',
        'test estilo aprendizaje',
        'estilo_aprendizaje'
    ]
WHERE guia = 'GUIA_00_INDUCCION'
  AND (
        actividad_resumen ILIKE '%estilo%aprendizaje%'
     OR actividad_resumen ILIKE '%kolb%'
  );


-- ── PASO 3: Identidad SENA ───────────────────────────────────────────────────
UPDATE actividades_parametrizadas SET
    nombre_esperado      = 'Info_Identidad SENA',
    variantes_permitidas = ARRAY[
        'info identidad sena',
        'info identidad',
        'infografia identidad sena',
        'identidad sena',
        'Info_Identidad'
    ]
WHERE guia = 'GUIA_00_INDUCCION'
  AND (
        actividad_resumen ILIKE '%identidad%sena%'
     OR actividad_resumen ILIKE '%infograf%identidad%'
  );


-- ── PASO 4: Plataformas SENA ─────────────────────────────────────────────────
UPDATE actividades_parametrizadas SET
    nombre_esperado      = 'Plataformas SENA',
    variantes_permitidas = ARRAY[
        'plataformas sena',
        'plataformas',
        'documento plataformas',
        'plataformas_sena'
    ]
WHERE guia = 'GUIA_00_INDUCCION'
  AND actividad_resumen ILIKE '%plataformas%';


-- ── PASO 5: Programa de formación ────────────────────────────────────────────
UPDATE actividades_parametrizadas SET
    nombre_esperado      = 'Mi programa de formacion',
    variantes_permitidas = ARRAY[
        'mi programa de formacion',
        'programa de formacion',
        'mi programa',
        'presentacion programa',
        'programa_formacion'
    ]
WHERE guia = 'GUIA_00_INDUCCION'
  AND (
        actividad_resumen ILIKE '%programa%formaci%'
     OR actividad_resumen ILIKE '%presentaci%programa%'
  );


-- ── PASO 6: Reglamento del aprendiz ──────────────────────────────────────────
UPDATE actividades_parametrizadas SET
    nombre_esperado      = 'Chat analisis reglamento',
    variantes_permitidas = ARRAY[
        'chat analisis reglamento',
        'chat reglamento',
        'analisis reglamento',
        'reglamento aprendiz',
        'chat_reglamento'
    ]
WHERE guia = 'GUIA_00_INDUCCION'
  AND actividad_resumen ILIKE '%reglamento%';


-- ── PASO 7: Propuesta proyecto productivo ────────────────────────────────────
UPDATE actividades_parametrizadas SET
    nombre_esperado      = 'Propuesta proyecto productivo',
    variantes_permitidas = ARRAY[
        'propuesta proyecto productivo',
        'propuesta proyecto',
        'proyecto productivo',
        'propuesta de proyecto',
        'propuesta_proyecto'
    ]
WHERE guia = 'GUIA_00_INDUCCION'
  AND (
        actividad_resumen ILIKE '%proyecto%productivo%'
     OR actividad_resumen ILIKE '%propuesta%proyecto%'
  );


-- ── PASO 8: Línea de tiempo ───────────────────────────────────────────────────
UPDATE actividades_parametrizadas SET
    nombre_esperado      = 'Linea de tiempo',
    variantes_permitidas = ARRAY[
        'linea de tiempo',
        'linea tiempo',
        'linea_de_tiempo',
        'proyeccion profesional',
        'linea de tiempo proyeccion profesional'
    ]
WHERE guia = 'GUIA_00_INDUCCION'
  AND (
        actividad_resumen ILIKE '%l_nea%tiempo%'
     OR actividad_resumen ILIKE '%proyecci%n%profesional%'
  );


-- ── VERIFICACIÓN ──────────────────────────────────────────────────────────────
-- Ejecutar esto después para confirmar que los cambios quedaron bien:
SELECT
    actividad_id,
    actividad_resumen,
    nombre_esperado,
    variantes_permitidas,
    carpeta_drive
FROM actividades_parametrizadas
WHERE guia = 'GUIA_00_INDUCCION'
ORDER BY actividad_id;
