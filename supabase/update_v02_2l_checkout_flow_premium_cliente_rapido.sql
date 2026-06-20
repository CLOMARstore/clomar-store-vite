-- V02.2L Checkout Flow Premium + Cliente Rápido
-- Registro opcional y seguro de versión. No borra ni modifica ventas, productos o clientes.

create table if not exists app_updates (
  id uuid primary key default gen_random_uuid(),
  version text not null,
  created_at timestamptz default now()
);

alter table app_updates
add column if not exists notes text;

alter table app_updates enable row level security;

insert into app_updates (version, notes)
select
  'V02.2L',
  'Checkout compacto: footer independiente, cliente rápido en modal, descuentos bajo demanda y panel POS refinado.'
where not exists (
  select 1 from app_updates where version = 'V02.2L'
);
