-- POLYTEC Ingeniera Comercialización — esquema inicial
-- Ejecuta este script en el SQL Editor de tu proyecto Supabase

create table if not exists public.projects (
  id          text        primary key,
  name        text        not null,
  messages    jsonb       not null default '[]'::jsonb,
  sales       jsonb,
  budget      jsonb,
  inventory   jsonb,
  created_at  timestamptz not null default now(),
  updated_at  timestamptz not null default now()
);

-- Actualiza updated_at automáticamente en cada UPDATE
create or replace function public.set_updated_at()
returns trigger language plpgsql as $$
begin
  new.updated_at = now();
  return new;
end;
$$;

drop trigger if exists trg_projects_updated_at on public.projects;
create trigger trg_projects_updated_at
  before update on public.projects
  for each row execute procedure public.set_updated_at();

-- Row Level Security: por ahora acceso público (sin auth)
-- Cuando agregues autenticación, reemplaza estas políticas
alter table public.projects enable row level security;

create policy "allow_all" on public.projects
  for all using (true) with check (true);
