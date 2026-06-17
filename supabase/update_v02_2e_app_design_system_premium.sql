-- V02.2E — App Design System Premium
-- Registro opcional de versión. No borra datos ni cambia tablas críticas.
create table if not exists app_updates (
  id uuid primary key default gen_random_uuid(),
  version text not null,
  description text,
  created_at timestamptz default now()
);
insert into app_updates(version, description) values ('V02.2E', 'App Design System Premium: modales simétricos, comprobante compacto, categorías en bottom sheet y mejor UX móvil.');
