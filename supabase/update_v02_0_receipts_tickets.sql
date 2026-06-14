-- V02.0 - Comprobantes internos / voucher / ticket para Clomar Store
-- Esta versión no borra datos. Prepara campos e índices para reimpresión y trazabilidad.

alter table sales add column if not exists receipt_notes text default '';
alter table sales add column if not exists receipt_format text default '80mm';
alter table sales add column if not exists receipt_printed_at timestamptz;

create index if not exists idx_sales_receipt_number on sales(receipt_number);
create index if not exists idx_sales_created_at on sales(created_at);
create index if not exists idx_sales_user_id on sales(user_id);
create index if not exists idx_sale_items_sale_id on sale_items(sale_id);
create index if not exists idx_sale_items_store_id on sale_items(store_id);

select 'V02.0 listo: comprobantes internos, tickets, vouchers y PDF A4 preparados.' as status;
