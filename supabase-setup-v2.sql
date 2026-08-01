-- ============================================================
-- Asset Manager — schema v2 (relational tables)
-- Run this in Supabase Dashboard → SQL Editor → New query → Run
--
-- This ADDS new tables. It does NOT touch or delete your existing
-- `app_data` table, so your current app keeps working untouched
-- until you switch the code over (see the rollout steps you were
-- given in chat).
-- ============================================================

-- ---------- 1. Tables ----------

create table if not exists locations (
  id text primary key,
  name text not null
);

create table if not exists categories (
  id text primary key,
  name text not null,
  type text not null,           -- 'IT' | 'Non-IT'
  useful_life int
);

create table if not exists users (
  id text primary key,
  name text not null,
  username text unique not null,
  email text,
  position text,
  role text not null,           -- 'Admin' | 'Regional Staff'
  location_id text references locations(id) on delete set null,
  password_hash text not null
);

create table if not exists assets (
  id text primary key,
  tag text,
  name text,
  category_id text references categories(id) on delete set null,
  asset_type text,              -- 'IT' | 'Non-IT'
  brand text,
  model text,
  year_model text,
  serial text,
  status text,
  condition text,
  location_id text references locations(id) on delete set null,
  assigned_to text,
  purchase_date date,
  purchase_cost numeric,
  warranty_expiry date,
  requires_calibration boolean default false,
  calibration_date date,
  next_calibration_date date,
  notes text,
  pre_repair_status text,
  transfer_history jsonb not null default '[]'::jsonb,  -- small nested list, fine as jsonb
  updated_at timestamptz not null default now()
);

create table if not exists maintenance (
  id text primary key,
  asset_id text references assets(id) on delete cascade,
  description text,
  cost numeric,
  date date,
  status text                    -- 'Not Started' | 'In Progress' | 'Done'
);

create table if not exists audit_log (
  id text primary key,
  at timestamptz not null default now(),
  user_id text,
  user_name text,
  message text
);

-- Helpful indexes for the filters/lookups the app does constantly
create index if not exists idx_assets_location on assets(location_id);
create index if not exists idx_assets_category on assets(category_id);
create index if not exists idx_maintenance_asset on maintenance(asset_id);
create index if not exists idx_users_location on users(location_id);
create index if not exists idx_audit_log_at on audit_log(at desc);

-- ---------- 2. Row Level Security ----------
-- Same "open" model as before (the app's own login screen — not the DB —
-- separates Admin vs Staff). Tightening this to real per-country RLS is a
-- good next step once everyone's on the new schema; call it out separately
-- when you're ready and I'll set up policies scoped by location_id.

alter table locations   enable row level security;
alter table categories  enable row level security;
alter table users       enable row level security;
alter table assets      enable row level security;
alter table maintenance enable row level security;
alter table audit_log   enable row level security;

do $$
declare
  t text;
begin
  foreach t in array array['locations','categories','users','assets','maintenance','audit_log']
  loop
    execute format('drop policy if exists "public read" on %I', t);
    execute format('create policy "public read" on %I for select using (true)', t);
    execute format('drop policy if exists "public insert" on %I', t);
    execute format('create policy "public insert" on %I for insert with check (true)', t);
    execute format('drop policy if exists "public update" on %I', t);
    execute format('create policy "public update" on %I for update using (true)', t);
    execute format('drop policy if exists "public delete" on %I', t);
    execute format('create policy "public delete" on %I for delete using (true)', t);
  end loop;
end $$;

-- ---------- 3. Realtime ----------

alter publication supabase_realtime add table locations;
alter publication supabase_realtime add table categories;
alter publication supabase_realtime add table users;
alter publication supabase_realtime add table assets;
alter publication supabase_realtime add table maintenance;
alter publication supabase_realtime add table audit_log;

-- ============================================================
-- 4. ONE-TIME MIGRATION — copies your existing data out of the
--    old app_data.payload blob into the new tables above.
--    Safe to re-run: it upserts, so running it twice won't duplicate rows.
-- ============================================================

insert into locations (id, name)
select l->>'id', l->>'name'
from app_data, jsonb_array_elements(payload->'locations') as l
where id = 1
on conflict (id) do update set name = excluded.name;

insert into categories (id, name, type, useful_life)
select c->>'id', c->>'name', c->>'type', (c->>'usefulLife')::int
from app_data, jsonb_array_elements(payload->'categories') as c
where id = 1
on conflict (id) do update set name = excluded.name, type = excluded.type, useful_life = excluded.useful_life;

insert into users (id, name, username, email, position, role, location_id, password_hash)
select u->>'id', u->>'name', u->>'username', u->>'email', u->>'position', u->>'role',
       nullif(u->>'locationId', ''), u->>'passwordHash'
from app_data, jsonb_array_elements(payload->'users') as u
where id = 1
on conflict (id) do update set
  name = excluded.name, username = excluded.username, email = excluded.email,
  position = excluded.position, role = excluded.role,
  location_id = excluded.location_id, password_hash = excluded.password_hash;

insert into assets (
  id, tag, name, category_id, asset_type, brand, model, year_model, serial, status, condition,
  location_id, assigned_to, purchase_date, purchase_cost, warranty_expiry,
  requires_calibration, calibration_date, next_calibration_date, notes,
  pre_repair_status, transfer_history
)
select
  a->>'id', a->>'tag', a->>'name', nullif(a->>'categoryId', ''), a->>'assetType',
  a->>'brand', a->>'model', a->>'yearModel', a->>'serial', a->>'status', a->>'condition',
  nullif(a->>'locationId', ''), a->>'assignedTo',
  nullif(a->>'purchaseDate','')::date, nullif(a->>'purchaseCost','')::numeric,
  nullif(a->>'warrantyExpiry','')::date,
  coalesce((a->>'requiresCalibration')::boolean, false),
  nullif(a->>'calibrationDate','')::date, nullif(a->>'nextCalibrationDate','')::date,
  a->>'notes', a->>'preRepairStatus',
  coalesce(a->'transferHistory', '[]'::jsonb)
from app_data, jsonb_array_elements(payload->'assets') as a
where id = 1
on conflict (id) do update set
  tag = excluded.tag, name = excluded.name, category_id = excluded.category_id,
  asset_type = excluded.asset_type, brand = excluded.brand, model = excluded.model,
  year_model = excluded.year_model,
  serial = excluded.serial, status = excluded.status, condition = excluded.condition,
  location_id = excluded.location_id, assigned_to = excluded.assigned_to,
  purchase_date = excluded.purchase_date, purchase_cost = excluded.purchase_cost,
  warranty_expiry = excluded.warranty_expiry,
  requires_calibration = excluded.requires_calibration,
  calibration_date = excluded.calibration_date, next_calibration_date = excluded.next_calibration_date,
  notes = excluded.notes, pre_repair_status = excluded.pre_repair_status,
  transfer_history = excluded.transfer_history;

insert into maintenance (id, asset_id, description, cost, date, status)
select m->>'id', m->>'assetId', m->>'description', nullif(m->>'cost','')::numeric,
       nullif(m->>'date','')::date, m->>'status'
from app_data, jsonb_array_elements(payload->'maintenance') as m
where id = 1
on conflict (id) do update set
  asset_id = excluded.asset_id, description = excluded.description,
  cost = excluded.cost, date = excluded.date, status = excluded.status;

insert into audit_log (id, at, user_id, user_name, message)
select e->>'id', (e->>'at')::timestamptz, e->>'userId', e->>'userName', e->>'message'
from app_data, jsonb_array_elements(payload->'auditLog') as e
where id = 1
on conflict (id) do update set
  at = excluded.at, user_id = excluded.user_id, user_name = excluded.user_name, message = excluded.message;

-- ---------- 5. Sanity check — run this and compare counts to your old data ----------
select
  (select count(*) from locations)  as locations,
  (select count(*) from categories) as categories,
  (select count(*) from users)      as users,
  (select count(*) from assets)     as assets,
  (select count(*) from maintenance) as maintenance,
  (select count(*) from audit_log)  as audit_log;
