-- V02.2J — Clientes Inteligentes + Fiscal Checkout Pro
-- Recomendado. No borra datos. Prepara ventas y clientes para futura integración fiscal.

create extension if not exists pgcrypto;

create table if not exists app_updates (
  id uuid primary key default gen_random_uuid(),
  version text not null,
  created_at timestamptz default now()
);
alter table app_updates add column if not exists notes text;
alter table app_updates enable row level security;

-- Cliente vinculado a la venta y datos fiscales base.
alter table customers add column if not exists document_type text;
alter table customers add column if not exists email text default '';
alter table customers add column if not exists business_name text default '';
alter table customers add column if not exists updated_at timestamptz default now();

alter table sales add column if not exists customer_id uuid references customers(id);
alter table sales add column if not exists document_type text default 'Interno';
alter table sales add column if not exists fiscal_series text;
alter table sales add column if not exists fiscal_correlative text;
alter table sales add column if not exists sunat_status text default 'Interno';
alter table sales add column if not exists customer_doc_type text;
alter table sales add column if not exists customer_doc_number text;
alter table sales add column if not exists electronic_provider text;
alter table sales add column if not exists xml_url text;
alter table sales add column if not exists cdr_url text;
alter table sales add column if not exists sunat_response text;

-- Inferir DNI/RUC cuando el cliente ya tiene documento registrado.
update customers
set document_type = case
  when regexp_replace(coalesce(document, ''), '\\D', '', 'g') ~ '^\\d{11}$' then 'RUC'
  when regexp_replace(coalesce(document, ''), '\\D', '', 'g') ~ '^\\d{8}$' then 'DNI'
  else coalesce(nullif(document_type, ''), 'DNI')
end
where coalesce(document_type, '') = '';

create index if not exists idx_customers_store_document on customers(store_id, document);
create index if not exists idx_sales_store_sunat_status on sales(store_id, sunat_status);
create index if not exists idx_sales_customer_id on sales(customer_id);

insert into app_updates (version, notes)
select
  'V02.2J',
  'Clientes creados desde ventas, datos fiscales dinámicos, factura con RUC obligatorio, pre-emisión SUNAT y persistencia de trazabilidad fiscal.'
where not exists (
  select 1 from app_updates where version = 'V02.2J'
);
