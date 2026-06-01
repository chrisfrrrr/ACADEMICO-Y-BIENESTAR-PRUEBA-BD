-- ============================================================
-- AVE Seguimiento Académico PRO v13 - Esquema Supabase
-- Base general multi-curso, multi-sección y multi-asesor
-- Ejecutar en Supabase > SQL Editor > New query > Run
-- ============================================================

-- Extensiones útiles
create extension if not exists pgcrypto;

-- ------------------------------------------------------------
-- 1. Tabla de asesores de bienestar
-- ------------------------------------------------------------
create table if not exists public.asesores_bienestar (
    id uuid primary key default gen_random_uuid(),
    nombre text not null,
    correo text,
    telefono text,
    estado text default 'Activo',
    observaciones text,
    created_at timestamptz default now(),
    updated_at timestamptz default now()
);

create unique index if not exists uq_asesores_bienestar_nombre
on public.asesores_bienestar (lower(trim(nombre)));

-- ------------------------------------------------------------
-- 2. Tabla de estudiantes por curso/sección
-- Un mismo estudiante puede aparecer en varios cursos o secciones.
-- ------------------------------------------------------------
create table if not exists public.estudiantes (
    id uuid primary key default gen_random_uuid(),
    periodo text,
    carne text,
    nombre text not null,
    correo text,
    telefono text,
    carrera text,
    estado text default 'Activo',

    curso text,
    curso_id_canvas text,
    seccion text,
    seccion_id_canvas text,
    semana integer,

    asesor_academico text,
    asesor_bienestar text,

    riesgo text,
    prioridad text,
    ultimo_promedio numeric,
    ultimo_actividades_pct numeric,
    ultimo_dias_inactivo integer,
    ultima_fecha_consulta timestamptz,

    canvas_user_id text,
    login_id text,

    created_at timestamptz default now(),
    updated_at timestamptz default now()
);

-- Evita duplicar al mismo estudiante dentro del mismo contexto académico.
create unique index if not exists uq_estudiantes_contexto
on public.estudiantes (
    coalesce(periodo,''),
    coalesce(carne,''),
    coalesce(curso_id_canvas,''),
    coalesce(seccion,''),
    coalesce(semana,0)
);

create index if not exists idx_estudiantes_curso on public.estudiantes (curso);
create index if not exists idx_estudiantes_seccion on public.estudiantes (seccion);
create index if not exists idx_estudiantes_riesgo on public.estudiantes (riesgo);
create index if not exists idx_estudiantes_asesor_academico on public.estudiantes (asesor_academico);
create index if not exists idx_estudiantes_asesor_bienestar on public.estudiantes (asesor_bienestar);
create index if not exists idx_estudiantes_carne on public.estudiantes (carne);

-- ------------------------------------------------------------
-- 3. Historial de análisis por estudiante, curso, sección y semana
-- Esta es la tabla central para medir mejora, empeoramiento o persistencia.
-- ------------------------------------------------------------
create table if not exists public.historial_estudiantes (
    id uuid primary key default gen_random_uuid(),
    fecha_consulta timestamptz default now(),

    periodo text,
    carne text,
    nombre text not null,
    correo text,

    curso text,
    curso_id_canvas text,
    seccion text,
    seccion_id_canvas text,
    semana integer,

    asesor_academico text,
    asesor_bienestar text,

    actividades_pct numeric,
    promedio numeric,
    entregas_tarde integer,
    entregas_pendientes integer,
    entregas_realizadas integer,
    total_actividades integer,
    ingresos_semana integer,
    dias_inactivo integer,
    ultima_actividad timestamptz,

    riesgo text,
    riesgo_anterior text,
    cambio text,
    motivo_detectado text,

    canvas_user_id text,
    consulta_id uuid,

    created_at timestamptz default now()
);

create index if not exists idx_historial_contexto
on public.historial_estudiantes (periodo, curso_id_canvas, seccion, semana);

create index if not exists idx_historial_estudiante
on public.historial_estudiantes (carne, nombre);

create index if not exists idx_historial_riesgo
on public.historial_estudiantes (riesgo);

create index if not exists idx_historial_fecha
on public.historial_estudiantes (fecha_consulta desc);

-- ------------------------------------------------------------
-- 4. Derivaciones contextualizadas a bienestar
-- ------------------------------------------------------------
create table if not exists public.derivaciones (
    id uuid primary key default gen_random_uuid(),
    id_derivacion text unique,
    fecha_derivacion timestamptz default now(),

    periodo text,
    carne text,
    nombre text not null,
    correo text,
    telefono text,
    carrera text,

    curso text,
    curso_id_canvas text,
    seccion text,
    seccion_id_canvas text,
    semana integer,

    asesor_academico text,
    asesor_bienestar text,

    riesgo text,
    prioridad text,
    motivo_derivacion text,
    acciones_previas text,
    observaciones text,
    estado_derivacion text default 'Pendiente',

    historial_id uuid,
    archivo_derivacion text,
    paquete_derivacion text,

    created_at timestamptz default now(),
    updated_at timestamptz default now()
);

create index if not exists idx_derivaciones_contexto
on public.derivaciones (periodo, curso, seccion, semana);

create index if not exists idx_derivaciones_asesor_bienestar
on public.derivaciones (asesor_bienestar);

create index if not exists idx_derivaciones_asesor_academico
on public.derivaciones (asesor_academico);

create index if not exists idx_derivaciones_riesgo
on public.derivaciones (riesgo, prioridad);

create index if not exists idx_derivaciones_carne
on public.derivaciones (carne);

-- ------------------------------------------------------------
-- 5. Mensajes enviados o generados
-- ------------------------------------------------------------
create table if not exists public.mensajes_enviados (
    id uuid primary key default gen_random_uuid(),
    fecha timestamptz default now(),

    periodo text,
    carne text,
    nombre text not null,
    correo text,

    curso text,
    curso_id_canvas text,
    seccion text,
    seccion_id_canvas text,
    semana integer,

    asesor_academico text,
    asesor_bienestar text,

    riesgo text,
    tipo_mensaje text,
    asunto text,
    mensaje_generado text,
    enviado_canvas boolean default false,
    estado_envio text default 'Generado',

    canvas_message_id text,
    created_at timestamptz default now()
);

create index if not exists idx_mensajes_contexto
on public.mensajes_enviados (periodo, curso, seccion, semana);

create index if not exists idx_mensajes_estudiante
on public.mensajes_enviados (carne, nombre);

-- ------------------------------------------------------------
-- 6. Consultas realizadas a Canvas
-- ------------------------------------------------------------
create table if not exists public.consultas_canvas (
    id uuid primary key default gen_random_uuid(),
    fecha_consulta timestamptz default now(),

    periodo text,
    curso text,
    curso_id_canvas text,
    seccion text,
    seccion_id_canvas text,
    semana integer,
    asesor_academico text,

    total_estudiantes integer default 0,
    bajo integer default 0,
    moderado integer default 0,
    alto integer default 0,
    sin_datos integer default 0,

    fuente text default 'Canvas API',
    observaciones text,
    created_at timestamptz default now()
);

create index if not exists idx_consultas_contexto
on public.consultas_canvas (periodo, curso_id_canvas, seccion, semana, asesor_academico);

-- ------------------------------------------------------------
-- 7. Configuración general
-- ------------------------------------------------------------
create table if not exists public.configuracion (
    id uuid primary key default gen_random_uuid(),
    parametro text unique not null,
    valor text,
    descripcion text,
    updated_at timestamptz default now()
);

insert into public.configuracion (parametro, valor, descripcion)
values
('version_base_datos', 'v13_multicurso', 'Versión del esquema de base de datos'),
('riesgo_bajo_min_actividades_pct', '80', 'Porcentaje mínimo sugerido para riesgo bajo'),
('riesgo_moderado_min_actividades_pct', '50', 'Porcentaje mínimo sugerido para riesgo moderado'),
('riesgo_alto_dias_inactivo', '6', 'Días de inactividad para alerta alta')
on conflict (parametro) do update set
valor = excluded.valor,
descripcion = excluded.descripcion,
updated_at = now();

-- ------------------------------------------------------------
-- 8. Tabla compacta anterior, por compatibilidad con v12
-- No se elimina para evitar pérdida de datos.
-- ------------------------------------------------------------
create table if not exists public.ave_base_datos (
    sheet_name text primary key,
    records jsonb not null default '[]'::jsonb,
    updated_at timestamptz default now(),
    updated_by text
);

-- ------------------------------------------------------------
-- 9. Vista: último registro por estudiante/curso/sección/semana
-- Útil para dashboard.
-- ------------------------------------------------------------
create or replace view public.vw_ultimo_estado_estudiante as
select distinct on (
    coalesce(periodo,''),
    coalesce(carne,''),
    coalesce(nombre,''),
    coalesce(curso_id_canvas,''),
    coalesce(seccion,''),
    coalesce(semana,0)
)
    *
from public.historial_estudiantes
order by
    coalesce(periodo,''),
    coalesce(carne,''),
    coalesce(nombre,''),
    coalesce(curso_id_canvas,''),
    coalesce(seccion,''),
    coalesce(semana,0),
    fecha_consulta desc;

-- ------------------------------------------------------------
-- 10. Vista: resumen institucional por curso/sección/asesor/riesgo
-- ------------------------------------------------------------
create or replace view public.vw_resumen_institucional as
select
    periodo,
    curso,
    curso_id_canvas,
    seccion,
    semana,
    asesor_academico,
    asesor_bienestar,
    riesgo,
    count(*) as total_estudiantes
from public.vw_ultimo_estado_estudiante
group by
    periodo,
    curso,
    curso_id_canvas,
    seccion,
    semana,
    asesor_academico,
    asesor_bienestar,
    riesgo;

-- ------------------------------------------------------------
-- 11. Función para updated_at automático
-- ------------------------------------------------------------
create or replace function public.set_updated_at()
returns trigger as $$
begin
    new.updated_at = now();
    return new;
end;
$$ language plpgsql;

-- Triggers updated_at
create trigger trg_asesores_bienestar_updated_at
before update on public.asesores_bienestar
for each row execute function public.set_updated_at();

drop trigger if exists trg_estudiantes_updated_at on public.estudiantes;
create trigger trg_estudiantes_updated_at
before update on public.estudiantes
for each row execute function public.set_updated_at();

drop trigger if exists trg_derivaciones_updated_at on public.derivaciones;
create trigger trg_derivaciones_updated_at
before update on public.derivaciones
for each row execute function public.set_updated_at();

-- ------------------------------------------------------------
-- 12. RLS
-- Para uso desde Streamlit Cloud con service_role_key, RLS puede quedar activo
-- porque service_role bypasses RLS. Si se usa anon key, deben crearse policies.
-- ------------------------------------------------------------
alter table public.asesores_bienestar enable row level security;
alter table public.estudiantes enable row level security;
alter table public.historial_estudiantes enable row level security;
alter table public.derivaciones enable row level security;
alter table public.mensajes_enviados enable row level security;
alter table public.consultas_canvas enable row level security;
alter table public.configuracion enable row level security;
alter table public.ave_base_datos enable row level security;

-- Nota: No se crean policies públicas. La app debe usar service_role_key desde Streamlit Secrets.
