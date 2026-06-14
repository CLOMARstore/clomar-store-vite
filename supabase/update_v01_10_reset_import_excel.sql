-- V01.10 - Reinicio controlado + importación Excel para Clomar Store
-- Ejecutar en Supabase SQL Editor antes de reemplazar los archivos de la app.

create table if not exists product_import_batches (
  id uuid primary key default gen_random_uuid(),
  store_id uuid references stores(id) on delete cascade,
  user_id uuid references auth.users(id),
  file_name text default '',
  total_rows int default 0,
  imported_rows int default 0,
  status text default 'Registrado',
  notes text default '',
  created_at timestamptz default now()
);

alter table product_import_batches enable row level security;

do $$ begin
  create policy "authenticated import batches read" on product_import_batches
  for select using (auth.role() = 'authenticated');
exception when duplicate_object then null; end $$;

do $$ begin
  create policy "authenticated import batches insert" on product_import_batches
  for insert with check (auth.role() = 'authenticated');
exception when duplicate_object then null; end $$;

-- Refuerzos necesarios para importación masiva.
alter table products add column if not exists barcode text default '';
alter table products add column if not exists subcategory text default '';
alter table products add column if not exists category_id uuid references product_categories(id);
alter table products add column if not exists subcategory_id uuid references product_categories(id);
alter table products add column if not exists brand text default '';
alter table products add column if not exists size text default '';
alter table products add column if not exists color text default '';
alter table products add column if not exists description text default '';
alter table products add column if not exists active boolean default true;
alter table products add column if not exists image_path text default '';
alter table products add column if not exists created_by uuid references auth.users(id);
alter table products add column if not exists store_id uuid references stores(id);

create index if not exists idx_products_store_code on products(store_id, code);
create index if not exists idx_products_store_barcode on products(store_id, barcode);
create index if not exists idx_products_category_subcategory on products(category_id, subcategory_id);

-- Función segura para reiniciar datos operativos sin tocar usuarios, tienda ni categorías.
create or replace function public.clomar_reset_operational_data(confirm_text text, store_uuid uuid default '00000000-0000-0000-0000-000000000001')
returns jsonb
language plpgsql
security definer
set search_path = public
as $$
declare
  current_role text;
  deleted_products int := 0;
  deleted_customers int := 0;
  deleted_sales int := 0;
  deleted_cash int := 0;
  deleted_stock int := 0;
  deleted_payments int := 0;
begin
  if confirm_text <> 'REINICIAR CLOMAR' then
    raise exception 'Confirmación inválida. Escribe exactamente REINICIAR CLOMAR.';
  end if;

  select role into current_role
  from profiles
  where id = auth.uid()
    and status = 'Activo'
    and store_id = store_uuid;

  if coalesce(current_role, '') <> 'dueno' then
    raise exception 'Solo el dueño activo puede reiniciar datos operativos.';
  end if;

  delete from credit_payments where store_id = store_uuid;
  get diagnostics deleted_payments = row_count;

  delete from cash_movements where store_id = store_uuid;
  get diagnostics deleted_cash = row_count;

  delete from stock_movements where store_id = store_uuid;
  get diagnostics deleted_stock = row_count;

  delete from sale_items where store_id = store_uuid;

  delete from sales where store_id = store_uuid;
  get diagnostics deleted_sales = row_count;

  delete from customers where store_id = store_uuid;
  get diagnostics deleted_customers = row_count;

  delete from products where store_id = store_uuid;
  get diagnostics deleted_products = row_count;

  return jsonb_build_object(
    'ok', true,
    'deleted_products', deleted_products,
    'deleted_customers', deleted_customers,
    'deleted_sales', deleted_sales,
    'deleted_cash_movements', deleted_cash,
    'deleted_stock_movements', deleted_stock,
    'deleted_credit_payments', deleted_payments
  );
end;
$$;

grant execute on function public.clomar_reset_operational_data(text, uuid) to authenticated;

select 'V01.10 listo: reinicio controlado e importación Excel preparados' as status;
