# Clomar Store Pro Free v01

Primera base de migración de Clomar Store desde Streamlit hacia una app más fluida con React + Vite + Supabase Free.

## Qué incluye
- PWA real instalable.
- Login con Supabase Auth.
- POS rápido con búsqueda instantánea.
- Carrito local fluido.
- Registro de venta en Supabase.
- Descuento automático de stock.
- Productos e inventario básico.
- Panel dueño compacto.
- Reportes ligeros sin gráficos pesados.

## Pasos
1. Crear proyecto en Supabase.
2. Ejecutar `supabase/schema.sql` en SQL Editor.
3. Crear usuario en Supabase Auth.
4. Copiar `.env.example` a `.env`.
5. Colocar `VITE_SUPABASE_URL` y `VITE_SUPABASE_ANON_KEY`.
6. Ejecutar:

```bash
npm install
npm run dev
```

## Despliegue gratis recomendado
- Vercel Hobby o Cloudflare Pages.
- Variables de entorno en el panel del hosting.

## Importante
Esta versión es fase 1: POS rápido y base Supabase. Streamlit puede quedarse como respaldo mientras se migra el resto.
