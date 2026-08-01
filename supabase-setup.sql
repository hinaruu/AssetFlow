-- Run this once in your Supabase project's SQL Editor
-- (Dashboard → SQL Editor → New query → paste → Run)

create table if not exists app_data (
  id int primary key,
  payload jsonb not null default '{}'::jsonb,
  updated_at timestamptz default now()
);

-- Enable Row Level Security, then allow the app's public "anon" key
-- to read and write this one table. This is intentionally simple —
-- everyone with the site URL shares the same data, and the app's own
-- login screen is what separates "Admin" from regular staff.
-- Don't put sensitive company secrets in this table.
alter table app_data enable row level security;

create policy "Public read access"
  on app_data for select
  using (true);

create policy "Public write access"
  on app_data for insert
  with check (true);

create policy "Public update access"
  on app_data for update
  using (true);

-- Enables the live feed: without this, changes only show up when someone
-- manually hits Sync. With it, every connected device gets pushed updates
-- the instant anyone saves.
alter publication supabase_realtime add table app_data;
