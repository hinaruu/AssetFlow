-- Run this in Supabase SQL Editor.
--
-- Supports the "asset update notifications" feature:
--
-- 1) Adds an "asset_id" column to the existing audit_log table. It's only
--    ever set for entries that relate to a single asset (status changes,
--    assignment changes, disposal, edits, transfers, maintenance, etc —
--    see withLog in App.jsx). It is NOT shown anywhere in the Activity Log
--    UI and doesn't change how that log looks or behaves — it's purely an
--    extra tag used to power the new per-asset notification badges/bell
--    for Overall Admin and Regional Admin.
--
-- 2) Creates a "notification_reads" table: one row per (user, asset)
--    recording the last time that user viewed that asset. Comparing an
--    asset's most recent activity timestamp (from audit_log + comments)
--    against this table is how "unread" is determined — no per-event
--    read/unread bookkeeping needed, and multiple updates naturally
--    accumulate as unread until the asset is opened again.
--
-- Safe to run even if some of this already exists.

alter table audit_log add column if not exists asset_id text;
create index if not exists idx_audit_log_asset on audit_log(asset_id);

create table if not exists notification_reads (
  id text primary key,
  user_id text not null,
  asset_id text not null,
  last_read_at timestamptz not null default now()
);
create index if not exists idx_notification_reads_user on notification_reads(user_id);
create index if not exists idx_notification_reads_asset on notification_reads(asset_id);

alter table notification_reads enable row level security;

drop policy if exists "public read" on notification_reads;
create policy "public read" on notification_reads for select using (true);
drop policy if exists "public insert" on notification_reads;
create policy "public insert" on notification_reads for insert with check (true);
drop policy if exists "public update" on notification_reads;
create policy "public update" on notification_reads for update using (true);
drop policy if exists "public delete" on notification_reads;
create policy "public delete" on notification_reads for delete using (true);

-- Enable Supabase Realtime for the new table so activity badges update
-- live for other signed-in users, same as the other tables.
do $$
begin
  if not exists (
    select 1 from pg_publication_tables
    where pubname = 'supabase_realtime' and schemaname = 'public' and tablename = 'notification_reads'
  ) then
    alter publication supabase_realtime add table notification_reads;
  end if;
end $$;
