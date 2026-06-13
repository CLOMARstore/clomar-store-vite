-- V01.8 Productos con imágenes + datos comerciales
-- Ejecutar en Supabase SQL Editor antes de reemplazar el código de la app.

-- 1) Campos nuevos para productos
alter table products add column if not exists image_url text default '';
alter table products add column if not exists image_path text default '';
alter table products add column if not exists brand text default '';
alter table products add column if not exists size text default '';
alter table products add column if not exists color text default '';
alter table products add column if not exists description text default '';
alter table products add column if not exists barcode text default '';
alter table products add column if not exists active boolean default true;

-- 2) Normalizar productos existentes
update products set active = true where active is null;
update products set image_url = '' where image_url is null;
update products set image_path = '' where image_path is null;
update products set brand = '' where brand is null;
update products set size = '' where size is null;
update products set color = '' where color is null;
update products set description = '' where description is null;
update products set barcode = '' where barcode is null;
update products set status = 'Activo' where status is null or status = '';

-- 3) Índices útiles
create index if not exists idx_products_active on products(active);
create index if not exists idx_products_barcode on products(barcode);
create index if not exists idx_products_brand on products(brand);
create index if not exists idx_products_category on products(category);

-- 4) Bucket público para imágenes de productos
insert into storage.buckets (id, name, public, file_size_limit, allowed_mime_types)
values (
  'product-images',
  'product-images',
  true,
  5242880,
  array['image/png','image/jpeg','image/jpg','image/webp']
)
on conflict (id) do update set
  public = true,
  file_size_limit = 5242880,
  allowed_mime_types = array['image/png','image/jpeg','image/jpg','image/webp'];

-- 5) Políticas de Storage para usuarios autenticados
-- Nota: si alguna política ya existe, el bloque evita error por duplicado.
do $$
begin
  create policy "product images public read"
  on storage.objects for select
  using (bucket_id = 'product-images');
exception when duplicate_object then null;
end $$;

do $$
begin
  create policy "product images authenticated insert"
  on storage.objects for insert
  with check (bucket_id = 'product-images' and auth.role() = 'authenticated');
exception when duplicate_object then null;
end $$;

do $$
begin
  create policy "product images authenticated update"
  on storage.objects for update
  using (bucket_id = 'product-images' and auth.role() = 'authenticated')
  with check (bucket_id = 'product-images' and auth.role() = 'authenticated');
exception when duplicate_object then null;
end $$;

do $$
begin
  create policy "product images authenticated delete"
  on storage.objects for delete
  using (bucket_id = 'product-images' and auth.role() = 'authenticated');
exception when duplicate_object then null;
end $$;

-- 6) Verificación rápida
select
  column_name,
  data_type
from information_schema.columns
where table_schema = 'public'
  and table_name = 'products'
  and column_name in ('image_url','image_path','brand','size','color','description','barcode','active')
order by column_name;
