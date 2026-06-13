-- V01.8B - Categorías administrables para Clomar Store

create table if not exists product_categories (
  id uuid primary key default gen_random_uuid(),
  store_id uuid references stores(id) on delete cascade,
  name text not null,
  description text default '',
  parent_id uuid references product_categories(id) on delete cascade,
  sort_order int default 100,
  active boolean default true,
  created_at timestamptz default now(),
  updated_at timestamptz default now(),
  unique(store_id, name, parent_id)
);

alter table products add column if not exists category_id uuid references product_categories(id);
alter table products add column if not exists subcategory_id uuid references product_categories(id);
alter table products add column if not exists subcategory text default '';

alter table product_categories enable row level security;

do $$ begin
  create policy "authenticated product_categories read" on product_categories
  for select using (auth.role() = 'authenticated');
exception when duplicate_object then null; end $$;

do $$ begin
  create policy "authenticated product_categories insert" on product_categories
  for insert with check (auth.role() = 'authenticated');
exception when duplicate_object then null; end $$;

do $$ begin
  create policy "authenticated product_categories update" on product_categories
  for update using (auth.role() = 'authenticated') with check (auth.role() = 'authenticated');
exception when duplicate_object then null; end $$;

create index if not exists idx_product_categories_store_id on product_categories(store_id);
create index if not exists idx_product_categories_parent_id on product_categories(parent_id);
create index if not exists idx_products_category_id on products(category_id);
create index if not exists idx_products_subcategory_id on products(subcategory_id);

-- Categorías principales necesarias según las carpetas actuales.
insert into product_categories (store_id, name, description, parent_id, sort_order, active)
values
('00000000-0000-0000-0000-000000000001','Ropa hombre','Prendas de vestir para hombre',null,10,true),
('00000000-0000-0000-0000-000000000001','Ropa mujer','Prendas de vestir para mujer',null,20,true),
('00000000-0000-0000-0000-000000000001','Calzado','Calzado para hombre y mujer',null,30,true),
('00000000-0000-0000-0000-000000000001','Accesorios de moda','Complementos personales, bolsos, lentes, gorras, joyería y relojes',null,40,true),
('00000000-0000-0000-0000-000000000001','Belleza y cuidado personal','Aseo, cremas, desodorantes, perfumes y cosmética',null,50,true),
('00000000-0000-0000-0000-000000000001','Hogar y decoración','Artículos para el hogar, decoración, espejos y cubrecamas',null,60,true),
('00000000-0000-0000-0000-000000000001','Fiesta y piñatería','Artículos para piñatería, fiestas y celebraciones',null,70,true),
('00000000-0000-0000-0000-000000000001','Bazar y utilitarios','Termos, tomatodos y artículos utilitarios',null,80,true)
on conflict (store_id, name, parent_id) do update set
  description = excluded.description,
  sort_order = excluded.sort_order,
  active = true,
  updated_at = now();

-- Subcategorías basadas en las carpetas mostradas.
with cats as (
  select id, name from product_categories where store_id = '00000000-0000-0000-0000-000000000001' and parent_id is null
)
insert into product_categories (store_id, name, description, parent_id, sort_order, active)
values
-- Ropa hombre
('00000000-0000-0000-0000-000000000001','Abrigos, casacas y chompas hombre','', (select id from cats where name='Ropa hombre'), 11, true),
('00000000-0000-0000-0000-000000000001','Camisas hombre','', (select id from cats where name='Ropa hombre'), 12, true),
('00000000-0000-0000-0000-000000000001','Polos hombre','', (select id from cats where name='Ropa hombre'), 13, true),
('00000000-0000-0000-0000-000000000001','Ropa interior hombre','', (select id from cats where name='Ropa hombre'), 14, true),
('00000000-0000-0000-0000-000000000001','Medias hombre','', (select id from cats where name='Ropa hombre'), 15, true),
-- Ropa mujer
('00000000-0000-0000-0000-000000000001','Abrigos, casacas y chompas mujer','', (select id from cats where name='Ropa mujer'), 21, true),
('00000000-0000-0000-0000-000000000001','Polos mujer','', (select id from cats where name='Ropa mujer'), 22, true),
('00000000-0000-0000-0000-000000000001','Ropa interior mujer','', (select id from cats where name='Ropa mujer'), 23, true),
('00000000-0000-0000-0000-000000000001','Pijamas','', (select id from cats where name='Ropa mujer'), 24, true),
-- Calzado
('00000000-0000-0000-0000-000000000001','Calzado hombre','', (select id from cats where name='Calzado'), 31, true),
('00000000-0000-0000-0000-000000000001','Calzado mujer','', (select id from cats where name='Calzado'), 32, true),
('00000000-0000-0000-0000-000000000001','Sandalias y tacos mujer','', (select id from cats where name='Calzado'), 33, true),
('00000000-0000-0000-0000-000000000001','Pantuflas','', (select id from cats where name='Calzado'), 34, true),
-- Accesorios de moda
('00000000-0000-0000-0000-000000000001','Carteras, bolsos y monederos mujer','', (select id from cats where name='Accesorios de moda'), 41, true),
('00000000-0000-0000-0000-000000000001','Mochilas','', (select id from cats where name='Accesorios de moda'), 42, true),
('00000000-0000-0000-0000-000000000001','Gafas o lentes hombre','', (select id from cats where name='Accesorios de moda'), 43, true),
('00000000-0000-0000-0000-000000000001','Gafas o lentes mujer','', (select id from cats where name='Accesorios de moda'), 44, true),
('00000000-0000-0000-0000-000000000001','Gorras hombre','', (select id from cats where name='Accesorios de moda'), 45, true),
('00000000-0000-0000-0000-000000000001','Gorras mujer','', (select id from cats where name='Accesorios de moda'), 46, true),
('00000000-0000-0000-0000-000000000001','Joyería hombre','', (select id from cats where name='Accesorios de moda'), 47, true),
('00000000-0000-0000-0000-000000000001','Joyería mujer','', (select id from cats where name='Accesorios de moda'), 48, true),
('00000000-0000-0000-0000-000000000001','Relojes hombre','', (select id from cats where name='Accesorios de moda'), 49, true),
-- Belleza y cuidado
('00000000-0000-0000-0000-000000000001','Cosas de aseo personal','', (select id from cats where name='Belleza y cuidado personal'), 51, true),
('00000000-0000-0000-0000-000000000001','Cremas y desodorantes','', (select id from cats where name='Belleza y cuidado personal'), 52, true),
('00000000-0000-0000-0000-000000000001','Cremas, pinturas y más mujer','', (select id from cats where name='Belleza y cuidado personal'), 53, true),
('00000000-0000-0000-0000-000000000001','Perfumes hombre','', (select id from cats where name='Belleza y cuidado personal'), 54, true),
('00000000-0000-0000-0000-000000000001','Perfumes mujer','', (select id from cats where name='Belleza y cuidado personal'), 55, true),
-- Hogar
('00000000-0000-0000-0000-000000000001','Accesorios para el hogar','', (select id from cats where name='Hogar y decoración'), 61, true),
('00000000-0000-0000-0000-000000000001','Cubrecamas','', (select id from cats where name='Hogar y decoración'), 62, true),
('00000000-0000-0000-0000-000000000001','Decoración hogar','', (select id from cats where name='Hogar y decoración'), 63, true),
('00000000-0000-0000-0000-000000000001','Espejos','', (select id from cats where name='Hogar y decoración'), 64, true),
-- Fiesta y bazar
('00000000-0000-0000-0000-000000000001','Piñatería','', (select id from cats where name='Fiesta y piñatería'), 71, true),
('00000000-0000-0000-0000-000000000001','Termos y tomatodos','', (select id from cats where name='Bazar y utilitarios'), 81, true)
on conflict (store_id, name, parent_id) do update set
  sort_order = excluded.sort_order,
  active = true,
  updated_at = now();

-- Asignación básica para productos antiguos por texto de categoría.
update products p
set category_id = c.id
from product_categories c
where p.store_id = c.store_id
  and c.parent_id is null
  and lower(trim(p.category)) = lower(trim(c.name))
  and p.category_id is null;

select
  parent.name as categoria,
  count(child.id) as subcategorias
from product_categories parent
left join product_categories child on child.parent_id = parent.id and child.active = true
where parent.store_id = '00000000-0000-0000-0000-000000000001'
  and parent.parent_id is null
  and parent.active = true
group by parent.name, parent.sort_order
order by parent.sort_order, parent.name;
