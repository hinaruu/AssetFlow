-- Run this in Supabase SQL Editor.
-- 1) Adds a "department" column to the existing assets table.
-- 2) Creates a new "comments" table for per-asset comments/clarification
--    threads, with realtime enabled so notifications show up live.
-- Safe to run even if some of this already exists.

alter table assets add column if not exists department text;

create table if not exists comments (
  id text primary key,
  asset_id text references assets(id) on delete cascade,
  at timestamptz not null default now(),
  author_id text,
  author_name text,
  message text,
  target_user_id text,
  read boolean not null default false
);

alter table comments enable row level security;

drop policy if exists "public read" on comments;
create policy "public read" on comments for select using (true);
drop policy if exists "public insert" on comments;
create policy "public insert" on comments for insert with check (true);
drop policy if exists "public update" on comments;
create policy "public update" on comments for update using (true);
drop policy if exists "public delete" on comments;
create policy "public delete" on comments for delete using (true);

-- Enable Supabase Realtime for the new table so comment notifications
-- show up live for the target user, same as the other tables.
alter publication supabase_realtime add table comments;
