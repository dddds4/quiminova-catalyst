-- Catalyst · Quiminova — esquema de base de datos
-- Pega TODO este archivo en Supabase: SQL Editor > New query > Run

create table if not exists app_data (
  user_id uuid primary key references auth.users(id) on delete cascade,
  data jsonb not null,
  updated_at timestamptz not null default now()
);

alter table app_data enable row level security;

drop policy if exists "select own data" on app_data;
create policy "select own data"
  on app_data for select
  using (auth.uid() = user_id);

drop policy if exists "insert own data" on app_data;
create policy "insert own data"
  on app_data for insert
  with check (auth.uid() = user_id);

drop policy if exists "update own data" on app_data;
create policy "update own data"
  on app_data for update
  using (auth.uid() = user_id);
