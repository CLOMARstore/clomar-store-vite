-- V02.2F — Mobile Commerce UX Premium
-- Registro opcional de actualización. No borra datos ni altera tablas operativas.

create table if not exists app_updates (
  id uuid primary key default gen_random_uuid(),
  version text not null,
  description text,
  created_at timestamptz default now()
);

insert into app_updates (version, description)
values ('V02.2F', 'Mobile Commerce UX Premium: carrito premium, KPIs 2x2, reportes compactos, precios, etiquetas, compras y herramientas optimizadas para celular.');
