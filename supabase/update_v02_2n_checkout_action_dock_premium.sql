-- V02.2N Checkout Action Dock Premium
-- Registro opcional de versión. No borra ni modifica ventas, productos o clientes.

create table if not exists app_updates (
  id uuid primary key default gen_random_uuid(),
  version text not null,
  created_at timestamptz default now()
);

alter table app_updates add column if not exists notes text;
alter table app_updates enable row level security;

insert into app_updates (version, notes)
select
  'V02.2N',
  'Barra de cobro premium: total visible en dock separado, CTA sin texto cortado y footer móvil ajustado.'
where not exists (
  select 1 from app_updates where version = 'V02.2N'
);
