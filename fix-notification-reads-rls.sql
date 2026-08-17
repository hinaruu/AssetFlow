-- ============================================================
-- FIX: "notification_reads" 403 errors when opening an asset
--
-- What we found: this table stores the app's own profile id in
-- user_id (e.g. "usr-2"), but the RLS policy compared it against
-- the Auth login UID instead — two different values that can
-- never match, so every read/write was silently rejected.
--
-- Run this in Supabase → SQL Editor → New query → Run (Role: postgres).
-- Safe to re-run.
-- ============================================================

create or replace function auth_user_profile_id() returns text
language sql stable security definer set search_path = public as $$
  select id from users where auth_user_id = auth.uid()
$$;

drop policy if exists "own reads only" on notification_reads;
drop policy if exists "own reads only insert" on notification_reads;
drop policy if exists "own reads only update" on notification_reads;

create policy "own reads only" on notification_reads
  for select using (user_id = auth_user_profile_id());
create policy "own reads only insert" on notification_reads
  for insert with check (user_id = auth_user_profile_id());
create policy "own reads only update" on notification_reads
  for update
  using (user_id = auth_user_profile_id())
  with check (user_id = auth_user_profile_id());
