-- =============================================================
-- PASO 1: Ejecutar este script en el SQL Editor de Supabase
--         antes de correr sincronizador_tabla.py por primera vez.
-- =============================================================

create table if not exists actividades_parametrizadas (
    id                    uuid          default gen_random_uuid() primary key,
    programa              text          not null,
    guia                  text          not null,
    actividad_id          text          not null,
    actividad_resumen     text,
    texto_fuente          text,
    nombre_esperado       text,
    variantes_permitidas  text[]        not null default '{}',
    tipo_archivo          text          not null default 'CUALQUIER_FORMATO',
    obligatoria           boolean       not null default true,
    carpeta_drive         text,
    regla_validacion      text,
    observaciones         text,
    activa                boolean       not null default true,
    updated_at            timestamptz   not null default now(),

    constraint actividades_clave_unica unique (programa, guia, actividad_id)
);

-- Índices para los filtros más frecuentes
create index if not exists idx_act_programa_guia
    on actividades_parametrizadas (programa, guia);

create index if not exists idx_act_activa
    on actividades_parametrizadas (activa)
    where activa = true;

-- Comentario sobre la columna activa:
-- El sincronizador NUNCA borra filas; solo upserta.
-- Para desactivar una actividad que ya no está en el Excel,
-- ejecutar manualmente:
--   update actividades_parametrizadas
--   set activa = false
--   where programa = 'X' and guia = 'Y' and actividad_id = 'Z';
