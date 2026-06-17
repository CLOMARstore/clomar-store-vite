-- CLOMAR STORE V02.2D — MOBILE APP UX PRO
-- Actualización visual y de experiencia móvil.
-- No borra datos ni modifica estructura crítica.

create table if not exists app_updates (
  id bigserial primary key,
  version text not null,
  description text,
  created_at timestamptz default now()
);

insert into app_updates (version, description)
values ('V02.2D', 'Mobile App UX Pro: carrito emergente, formularios móviles, listas compactas y mejor experiencia táctil.');
