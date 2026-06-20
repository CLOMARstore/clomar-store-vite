-- V02.2I Checkout Pro + SUNAT Ready UX
-- Seguro: no borra datos. Agrega columnas opcionales para preparar futura integración SUNAT/PSE/OSE.

create table if not exists app_updates (
  id uuid primary key default gen_random_uuid(),
  version text not null,
  created_at timestamptz default now()
);

alter table app_updates add column if not exists notes text;
alter table app_updates enable row level security;

-- Campos opcionales en ventas para preparar comprobante electrónico futuro.
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

insert into app_updates (version, notes)
select
  'V02.2I',
  'Checkout Pro con selector Interno/Boleta/Factura, estados SUNAT-ready, datos fiscales de cliente y base preparada para PSE/OSE.'
where not exists (
  select 1 from app_updates where version = 'V02.2I'
);
