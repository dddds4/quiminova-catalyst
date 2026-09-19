-- Catalyst · Migración a multiusuario (empresa + roles)
-- SEGURA Y ADITIVA: no borra ni modifica tu tabla app_data original.
-- Crea una copia de respaldo automática antes de tocar nada.
-- Pega TODO este archivo en Supabase: SQL Editor > New query > Run

-- ============================================================
-- 0. RESPALDO AUTOMÁTICO de tu información actual, por seguridad
-- ============================================================
create table if not exists app_data_backup_pre_multiuser as
  select * from app_data;

-- ============================================================
-- 1. Empresas
-- ============================================================
create table if not exists companies (
  id uuid primary key default gen_random_uuid(),
  name text not null default 'Mi empresa',
  owner_id uuid not null references auth.users(id) on delete cascade,
  created_at timestamptz not null default now()
);

-- ============================================================
-- 2. Miembros de cada empresa (quién pertenece y con qué rol)
-- ============================================================
create table if not exists company_members (
  company_id uuid not null references companies(id) on delete cascade,
  user_id uuid not null references auth.users(id) on delete cascade,
  role text not null check (role in ('owner', 'vendedor')),
  created_at timestamptz not null default now(),
  primary key (company_id, user_id)
);

-- ============================================================
-- 3. Datos de cada empresa (reemplaza a app_data, pero por empresa)
-- ============================================================
create table if not exists company_data (
  company_id uuid primary key references companies(id) on delete cascade,
  data jsonb not null,
  updated_at timestamptz not null default now()
);

-- ============================================================
-- 4. Migración: crea tu empresa a partir de tu usuario actual,
--    te agrega como dueño, y copia tu información hacia ella.
--    (No borra nada de app_data; solo copia).
-- ============================================================
insert into companies (owner_id, name)
select ad.user_id, coalesce(ad.data->'company'->>'name', 'Mi empresa')
from app_data ad
where not exists (select 1 from companies c where c.owner_id = ad.user_id);

insert into company_members (company_id, user_id, role)
select c.id, c.owner_id, 'owner'
from companies c
where not exists (
  select 1 from company_members m where m.company_id = c.id and m.user_id = c.owner_id
);

insert into company_data (company_id, data, updated_at)
select c.id, ad.data, ad.updated_at
from companies c
join app_data ad on ad.user_id = c.owner_id
where not exists (select 1 from company_data cd where cd.company_id = c.id);

-- ============================================================
-- 5. Seguridad (RLS): cada quien solo ve la empresa a la que pertenece
-- ============================================================
alter table companies enable row level security;
alter table company_members enable row level security;
alter table company_data enable row level security;

drop policy if exists "ver mi empresa" on companies;
create policy "ver mi empresa" on companies for select
  using (id in (select company_id from company_members where user_id = auth.uid()));

drop policy if exists "dueno actualiza empresa" on companies;
create policy "dueno actualiza empresa" on companies for update
  using (owner_id = auth.uid());

drop policy if exists "ver miembros de mi empresa" on company_members;
create policy "ver miembros de mi empresa" on company_members for select
  using (company_id in (select company_id from company_members where user_id = auth.uid()));

drop policy if exists "dueno administra miembros" on company_members;
create policy "dueno administra miembros" on company_members for all
  using (company_id in (select id from companies where owner_id = auth.uid()))
  with check (company_id in (select id from companies where owner_id = auth.uid()));

drop policy if exists "miembros ven datos de su empresa" on company_data;
create policy "miembros ven datos de su empresa" on company_data for select
  using (company_id in (select company_id from company_members where user_id = auth.uid()));

drop policy if exists "miembros actualizan datos de su empresa" on company_data;
create policy "miembros actualizan datos de su empresa" on company_data for update
  using (company_id in (select company_id from company_members where user_id = auth.uid()));

drop policy if exists "miembros insertan datos de su empresa" on company_data;
create policy "miembros insertan datos de su empresa" on company_data for insert
  with check (company_id in (select company_id from company_members where user_id = auth.uid()));

-- Listo. Tu tabla app_data original y su respaldo (app_data_backup_pre_multiuser)
-- quedan intactos y sin usarse por la app a partir de ahora — puedes conservarlos
-- como historial o borrarlos más adelante cuando confirmes que todo funciona bien.
