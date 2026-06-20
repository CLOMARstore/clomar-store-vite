-- V02.2K Responsive Checkout + Importador Pro
-- Registro opcional y seguro. No borra ni modifica datos operativos.

create table if not exists public.app_updates (
  id uuid primary key default gen_random_uuid(),
  version text not null,
  notes text,
  created_at timestamptz default now()
);

alter table public.app_updates add column if not exists notes text;
alter table public.app_updates enable row level security;

insert into public.app_updates (version, notes)
select
  'V02.2K',
  'Checkout responsivo, cliente compacto y asistente de importación Excel con modos completo, stock y precios.'
where not exists (
  select 1 from public.app_updates where version = 'V02.2K'
);
