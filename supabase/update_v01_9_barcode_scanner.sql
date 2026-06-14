-- V01.9 Código de barras y lector físico/celular
-- Este script no borra datos. Solo refuerza el campo barcode y agrega índices útiles.

alter table products add column if not exists barcode text default '';

-- Evita códigos de barras duplicados cuando el barcode no está vacío.
create unique index if not exists products_barcode_unique_not_empty
on products (barcode)
where barcode is not null and trim(barcode) <> '';

create index if not exists idx_products_barcode on products(barcode);
create index if not exists idx_products_code on products(code);
create index if not exists idx_products_active on products(active);

-- Verificación rápida
select code, barcode, name, stock, active
from products
order by created_at desc
limit 20;
