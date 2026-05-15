-- =============================================================
-- acumulado_migracion.sql
-- Ejecutar en Supabase → SQL Editor (una sola vez)
--
-- Habilita auditoría acumulativa multi-guía:
--   1. Unique constraint en verificaciones → elimina duplicados
--   2. Vista vw_dashboard_acumulado       → detalle por aprendiz + guía
--   3. Vista vw_progreso_aprendiz         → total acumulado por aprendiz
--   4. Vista vw_resumen_ficha             → KPIs por ficha
-- =============================================================


-- ── 1. CONSTRAINT ÚNICO ───────────────────────────────────────────────────────
-- Llave lógica: (docente_id, ficha, estudiante, guia, evidencia)
-- Permite UPSERT futuro y garantiza que re-auditar una guía no duplique.

ALTER TABLE verificaciones
ADD CONSTRAINT IF NOT EXISTS verificaciones_unica_por_actividad
UNIQUE (docente_id, ficha, estudiante, guia, evidencia);

-- Índice de soporte para las consultas del dashboard
CREATE INDEX IF NOT EXISTS idx_verif_docente_ficha
    ON verificaciones (docente_id, ficha);

CREATE INDEX IF NOT EXISTS idx_verif_guia
    ON verificaciones (docente_id, ficha, guia);


-- ── 2. VISTA: detalle por aprendiz y guía ────────────────────────────────────
-- Retorna una fila por (aprendiz, guía) con totales y porcentaje de esa guía.

CREATE OR REPLACE VIEW vw_dashboard_acumulado AS
SELECT
    docente_id,
    ficha,
    colegio,
    estudiante,
    guia,
    COUNT(*)                                                          AS total_actividades,
    SUM(CASE WHEN entregado THEN 1 ELSE 0 END)                       AS total_ok,
    COUNT(*) - SUM(CASE WHEN entregado THEN 1 ELSE 0 END)            AS total_pendientes,
    ROUND(
        SUM(CASE WHEN entregado THEN 1 ELSE 0 END)::numeric
        / NULLIF(COUNT(*), 0)::numeric * 100,
    1)                                                               AS porcentaje_guia,
    MAX(fecha)                                                        AS ultima_auditoria
FROM verificaciones
GROUP BY docente_id, ficha, colegio, estudiante, guia;


-- ── 3. VISTA: progreso acumulado por aprendiz (todas sus guías) ──────────────
-- Agrupa vw_dashboard_acumulado para mostrar el avance total del aprendiz.

CREATE OR REPLACE VIEW vw_progreso_aprendiz AS
SELECT
    docente_id,
    ficha,
    colegio,
    estudiante,
    COUNT(DISTINCT guia)                      AS guias_auditadas,
    SUM(total_actividades)                    AS total_actividades,
    SUM(total_ok)                             AS total_ok,
    SUM(total_pendientes)                     AS total_pendientes,
    ROUND(
        SUM(total_ok)::numeric
        / NULLIF(SUM(total_actividades), 0)::numeric * 100,
    1)                                        AS porcentaje_acumulado,
    MAX(ultima_auditoria)                     AS ultima_auditoria
FROM vw_dashboard_acumulado
GROUP BY docente_id, ficha, colegio, estudiante;


-- ── 4. VISTA: KPIs por ficha ──────────────────────────────────────────────────
-- Una fila por ficha con los números globales.

CREATE OR REPLACE VIEW vw_resumen_ficha AS
SELECT
    docente_id,
    ficha,
    colegio,
    COUNT(DISTINCT estudiante)     AS total_aprendices,
    COUNT(DISTINCT guia)           AS guias_auditadas,
    SUM(total_actividades)         AS total_evidencias_esperadas,
    SUM(total_ok)                  AS total_verificadas,
    SUM(total_pendientes)          AS total_pendientes,
    ROUND(
        SUM(total_ok)::numeric
        / NULLIF(SUM(total_actividades), 0)::numeric * 100,
    1)                             AS porcentaje_general
FROM vw_dashboard_acumulado
GROUP BY docente_id, ficha, colegio;


-- ── 5. VERIFICACIÓN ───────────────────────────────────────────────────────────
-- Ejecutar después para confirmar que las vistas funcionan:

-- SELECT * FROM vw_dashboard_acumulado  LIMIT 20;
-- SELECT * FROM vw_progreso_aprendiz    LIMIT 20;
-- SELECT * FROM vw_resumen_ficha        LIMIT 10;
