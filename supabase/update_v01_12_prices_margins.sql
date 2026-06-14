-- V01.12 - Control de precios, costos y márgenes para Clomar Store
-- Ejecutar en Supabase SQL Editor antes de reemplazar los archivos de la app.

alter table products add column if not exists price_status text default 'Pendiente';
alter table products add column if not exists margin_target numeric(8,2) default 50;
alter table products add column if not exists min_price numeric(12,2) default 0;
alter table products add column if not exists price_notes text default '';
alter table products add column if not exists price_updated_at timestamptz;
alter table products add column if not exists price_updated_by uuid references auth.users(id);

-- No se asume que los precios actuales son correctos. Todo lo no validado queda Pendiente.
update products
set price_status = 'Pendiente'
where price_status is null or trim(price_status) = '';

update products
set margin_target = 50
where margin_target is null;

update products
set min_price = 0
where min_price is null;

do $$ begin
  alter table products add constraint products_price_status_check
  check (price_status in ('Pendiente', 'Validado', 'Revisar'));
exception when duplicate_object then null;
end $$;

create table if not exists product_price_history (
  id uuid primary key default gen_random_uuid(),
  store_id uuid references stores(id) on delete cascade,
  product_id uuid references products(id) on delete cascade,
  user_id uuid references auth.users(id),
  old_cost numeric(12,2) default 0,
  new_cost numeric(12,2) default 0,
  old_price numeric(12,2) default 0,
  new_price numeric(12,2) default 0,
  old_status text default 'Pendiente',
  new_status text default 'Pendiente',
  note text default '',
  created_at timestamptz default now()
);

alter table product_price_history enable row level security;

do $$ begin
  create policy "authenticated price history read"
  on product_price_history for select
  using (auth.role() = 'authenticated');
exception when duplicate_object then null;
end $$;

do $$ begin
  create policy "authenticated price history insert"
  on product_price_history for insert
  with check (auth.role() = 'authenticated');
exception when duplicate_object then null;
end $$;

alter table sale_items add column if not exists unit_cost numeric(12,2) default 0;
alter table sale_items add column if not exists profit numeric(12,2) default 0;
alter table sale_items add column if not exists margin_percent numeric(8,2) default 0;

create index if not exists idx_products_price_status on products(price_status);
create index if not exists idx_products_price_updated_at on products(price_updated_at);
create index if not exists idx_price_history_product_id on product_price_history(product_id);
create index if not exists idx_price_history_store_id on product_price_history(store_id);

select 'V01.12 listo: control de precios, costos, márgenes e historial preparado. Los precios actuales NO se asumieron como validados.' as status;
