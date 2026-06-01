-- Esquema mínimo para AVE Seguimiento PRO v12
-- Ejecutar en Supabase > SQL Editor > New query > Run

create table if not exists public.ave_base_datos (
    sheet_name text primary key,
    records jsonb not null default '[]'::jsonb,
    updated_at timestamptz not null default now(),
    updated_by text,
    notes text
);

-- Seguridad recomendada: activar RLS. La app debe usar service_role_key desde Streamlit Secrets.
alter table public.ave_base_datos enable row level security;

comment on table public.ave_base_datos is 'Base en línea AVE Seguimiento PRO. Cada fila guarda una hoja lógica de la base en formato JSONB.';
comment on column public.ave_base_datos.sheet_name is 'Nombre de la hoja lógica: Estudiantes, Historial_Estudiantes, Derivaciones, etc.';
comment on column public.ave_base_datos.records is 'Registros de la hoja almacenados como arreglo JSON.';
