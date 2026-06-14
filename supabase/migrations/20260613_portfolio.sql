create table if not exists portfolio (
  id integer primary key default 1,
  data jsonb not null default '[]',
  updated_at timestamptz default now()
);

-- Permite solo una fila
create unique index if not exists portfolio_single_row on portfolio (id);

-- App pública: lectura y escritura sin auth
alter table portfolio enable row level security;

create policy "public select" on portfolio
  for select using (true);

create policy "public insert" on portfolio
  for insert with check (true);

create policy "public update" on portfolio
  for update using (true);

-- Fila inicial vacía
insert into portfolio (id, data)
  values (1, '[]')
  on conflict (id) do nothing;
