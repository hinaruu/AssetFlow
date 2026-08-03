-- Run this in Supabase SQL Editor.
-- 1) Adds a "department" column to the existing assets table.
-- 2) Creates (or upgrades) a "comments" table for per-asset comments —
--    each comment can notify multiple people (the regional users for that
--    asset's location) and tracks read/unread per person.
-- Safe to run even if some of this already exists, including if you
-- already ran an earlier version of this file with a single target_user_id.

alter table assets add column if not exists department text;

create table if not exists comments (
  id text primary key,
  asset_id text references assets(id) on delete cascade,
  at timestamptz not null default now(),
  author_id text,
  author_name text,
  message text
);

-- New multi-target columns (a comment can notify several regional users).
alter table comments add column if not exists target_user_ids jsonb not null default '[]'::jsonb;
alter table comments add column if not exists read_by jsonb not null default '[]'::jsonb;

-- If this table was created by an earlier version of this script with a
-- single target_user_id/read column, migrate that data across, then drop them.
do $$
begin
  if exists (select 1 from information_schema.columns where table_name = 'comments' and column_name = 'target_user_id') then
    update comments
      set target_user_ids = case when target_user_id is not null then jsonb_build_array(target_user_id) else '[]'::jsonb end
      where target_user_ids = '[]'::jsonb;
    update comments
      set read_by = case when read then jsonb_build_array(target_user_id) else '[]'::jsonb end
      where read_by = '[]'::jsonb and read is true;
    alter table comments drop column target_user_id;
    alter table comments drop column read;
  end if;
end $$;

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
-- show up live for the target users, same as the other tables.
-- (Guarded — safe to re-run even if this table is already in the publication.)
do $$
begin
  if not exists (
    select 1 from pg_publication_tables
    where pubname = 'supabase_realtime' and schemaname = 'public' and tablename = 'comments'
  ) then
    alter publication supabase_realtime add table comments;
  end if;
end $$;
