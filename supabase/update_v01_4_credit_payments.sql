-- V01.4: abonos de créditos y saldo por comprobante
create table if not exists credit_payments (
  id uuid primary key default gen_random_uuid(),
  sale_id uuid references sales(id) on delete cascade,
  customer_name text default 'Cliente',
  amount numeric(12,2) default 0,
  method text default 'Efectivo',
  note text default '',
  created_at timestamptz default now(),
  user_email text default auth.email()
);

alter table credit_payments enable row level security;

do $$ begin
  create policy "authenticated credit payments" on credit_payments for all using (auth.role() = 'authenticated') with check (auth.role() = 'authenticated');
exception when duplicate_object then null; end $$;
