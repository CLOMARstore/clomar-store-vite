-- V02.2P Cart Visibility & Navigation Layer Fix
-- Registro opcional de versión. No modifica datos operativos.
create table if not exists app_updates (
  id uuid primary key default gen_random_uuid(),
  version text not null,
  created_at timestamptz default now()
);
alter table app_updates add column if not exists notes text;
alter table app_updates enable row level security;
insert into app_updates(version, notes)
select 'V02.2P', 'Corrige visibilidad del mini carrito, menú móvil y estado de producto agregado.'
where not exists (select 1 from app_updates where version='V02.2P');
