-- V02.2H Checkout + Categories + Inventory UX Polish
-- Registro opcional de version. No borra datos.
create table if not exists app_updates (
  id uuid primary key default gen_random_uuid(),
  version text not null,
  notes text,
  created_at timestamptz default now()
);

alter table app_updates enable row level security;

insert into app_updates (version, notes)
values ('V02.2H', 'Carrito Pro mejorado, flujo de cobro corregido, categorias con detalle en ventana inferior e inventario compacto.')
on conflict do nothing;
