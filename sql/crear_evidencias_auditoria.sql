-- =============================================================
-- PASO 2: Ejecutar DESPUÉS de crear_actividades_parametrizadas.sql
-- Tabla que almacena resultados enriquecidos (Fase 2).
-- La tabla verificaciones existente se conserva sin cambios.
-- =============================================================

create table if not exists evidencias_auditoria (
    id                  uuid        default gen_random_uuid() primary key,
    docente_id          uuid,
    ficha               text,
    colegio             text,
    estudiante          text        not null,
    actividad_id        text,
    guia                text        not null,
    programa            text,
    actividad_resumen   text,
    nombre_esperado     text,
    obligatoria         boolean     default true,

    -- Resultado del motor
    estado              text        not null,   -- OK | FALTA | NOMBRE_DIFERENTE | ...
    archivo_encontrado  text,
    carpeta_encontrada  text,
    extension           text,
    confianza           smallint,              -- 0–100
    patron_match        text,
    observaciones       text,

    auditado_en         timestamptz default now()
);

-- Índices para los filtros del dashboard
create index if not exists idx_audit_ficha_guia
    on evidencias_auditoria (ficha, guia);

create index if not exists idx_audit_estudiante
    on evidencias_auditoria (estudiante);

create index if not exists idx_audit_estado
    on evidencias_auditoria (estado);

create index if not exists idx_audit_docente
    on evidencias_auditoria (docente_id);
