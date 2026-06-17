-- CLOMAR STORE V02.2G — Native Mobile Commerce Layout
-- Registro opcional de versión. No elimina datos.
create table if not exists public.app_updates (
  id bigserial primary key,
  version text not null,
  description text,
  created_at timestamptz default now()
);

alter table public.app_updates enable row level security;

insert into public.app_updates (version, description)
values (
  'V02.2G',
  'Native Mobile Commerce Layout: carrito bottom sheet real, grids 2 por fila, etiquetas compactas, categorías 2 columnas, caja compacta, precios con ayuda tipo burbuja y módulos administrativos con menos scroll.'
);
