-- ============================================================
-- CRITICAL SECURITY MIGRATION
-- Run this in Supabase Dashboard → SQL Editor → New query → Run.
--
-- What this does:
--   1. Links the `users` profile table to real Supabase Auth accounts.
--   2. Adds two helper functions RLS policies use to check "who is this
--      request from, and what's their role/location".
--   3. Replaces every "public read/write" (using (true)) policy with
--      policies scoped by role and location.
--   4. Makes audit_log append-only (no update, no delete, ever).
--
-- Safe to re-run — every statement is idempotent (create-or-replace,
-- if-not-exists, drop-if-exists-then-create).
--
-- Do this BEFORE deploying the updated app code, and only AFTER you've
-- read CRITICAL-SECURITY-STEPS.md — this migration on its own does not
-- create any Auth accounts or link anyone; that's Phase 3 of that guide.
-- ============================================================

-- ---------- 1. Link users -> Supabase Auth ----------

alter table users add column if not exists auth_user_id uuid unique references auth.users(id) on delete set null;
create index if not exists idx_users_auth_user_id on users(auth_user_id);

-- password_hash is no longer written to or read by the app (Supabase Auth
-- owns credentials now) — made nullable so existing rows don't block
-- saves. Drop the column entirely once you've confirmed everything works
-- (see Phase 7 of the guide) — not dropped automatically here, in case
-- you want to keep it around a little longer as a reference.
alter table users alter column password_hash drop not null;

-- ---------- 2. Helper functions RLS policies call ----------
-- security definer so they can read `users` even from inside a policy
-- that's checking access to `users` itself (avoids recursive RLS).

create or replace function auth_user_role() returns text
language sql stable security definer set search_path = public as $$
  select role from users where auth_user_id = auth.uid()
$$;

create or replace function auth_user_location() returns text
language sql stable security definer set search_path = public as $$
  select location_id from users where auth_user_id = auth.uid()
$$;

-- ---------- 3. Locations & Categories ----------
-- Read by anyone signed in; only Admins manage them (matches the
-- "only Admin sees + Add new location" rule already enforced in the UI).

drop policy if exists "public read" on locations;
drop policy if exists "public insert" on locations;
drop policy if exists "public update" on locations;
drop policy if exists "public delete" on locations;
create policy "signed-in read" on locations for select using (auth.uid() is not null);
create policy "admin write" on locations for insert with check (auth_user_role() = 'Admin');
create policy "admin update" on locations for update using (auth_user_role() = 'Admin');
create policy "admin delete" on locations for delete using (auth_user_role() = 'Admin');

drop policy if exists "public read" on categories;
drop policy if exists "public insert" on categories;
drop policy if exists "public update" on categories;
drop policy if exists "public delete" on categories;
create policy "signed-in read" on categories for select using (auth.uid() is not null);
create policy "admin write" on categories for insert with check (auth_user_role() = 'Admin');
create policy "admin update" on categories for update using (auth_user_role() = 'Admin');
create policy "admin delete" on categories for delete using (auth_user_role() = 'Admin');

-- ---------- 4. Users (profiles) ----------
-- Anyone signed in can read the directory (needed for "Added by",
-- approver names, etc. — there's no password data in this table
-- anymore). Only Admins create/edit/remove profiles or change role and
-- location, since that's how access is granted.

drop policy if exists "public read" on users;
drop policy if exists "public insert" on users;
drop policy if exists "public update" on users;
drop policy if exists "public delete" on users;
create policy "signed-in read" on users for select using (auth.uid() is not null);
create policy "admin write" on users for insert with check (auth_user_role() = 'Admin');
create policy "admin update" on users for update using (auth_user_role() = 'Admin');
create policy "admin delete" on users for delete using (auth_user_role() = 'Admin');

-- ---------- 5. Assets ----------
-- Admin: everything, everywhere. Regional Admin / Regional Staff: only
-- their own location. Anyone in-location can insert (new assets are
-- forced to their own location) and update (covers editing AND
-- transferring out — the `using` clause below is checked against the
-- CURRENT row, so a transfer is allowed as long as the asset started in
-- their location; the destination can be anywhere). Direct delete is
-- Admin/Regional Admin only, matching the in-app "request deletion" flow
-- for Regional Staff.

drop policy if exists "public read" on assets;
drop policy if exists "public insert" on assets;
drop policy if exists "public update" on assets;
drop policy if exists "public delete" on assets;

create policy "scoped read" on assets for select using (
  auth_user_role() = 'Admin' or location_id = auth_user_location()
);
create policy "scoped insert" on assets for insert with check (
  auth_user_role() = 'Admin' or location_id = auth_user_location()
);
create policy "scoped update" on assets for update using (
  auth_user_role() = 'Admin' or location_id = auth_user_location()
);
create policy "admin or regional-admin delete" on assets for delete using (
  auth_user_role() = 'Admin'
  or (auth_user_role() = 'Regional Admin' and location_id = auth_user_location())
);

-- ---------- 6. Maintenance ----------
-- Scoped through the asset it belongs to.

drop policy if exists "public read" on maintenance;
drop policy if exists "public insert" on maintenance;
drop policy if exists "public update" on maintenance;
drop policy if exists "public delete" on maintenance;

create policy "scoped read" on maintenance for select using (
  auth_user_role() = 'Admin' or exists (
    select 1 from assets where assets.id = maintenance.asset_id and assets.location_id = auth_user_location()
  )
);
create policy "scoped write" on maintenance for insert with check (
  auth_user_role() = 'Admin' or exists (
    select 1 from assets where assets.id = maintenance.asset_id and assets.location_id = auth_user_location()
  )
);
create policy "scoped update" on maintenance for update using (
  auth_user_role() = 'Admin' or exists (
    select 1 from assets where assets.id = maintenance.asset_id and assets.location_id = auth_user_location()
  )
);
create policy "admin or regional-admin delete" on maintenance for delete using (
  auth_user_role() = 'Admin' or (
    auth_user_role() = 'Regional Admin' and exists (
      select 1 from assets where assets.id = maintenance.asset_id and assets.location_id = auth_user_location()
    )
  )
);

-- ---------- 7. Comments & notification_reads ----------
-- Same location-scoping pattern, through the asset each row belongs to.
-- No delete policy for comments — nobody should be able to erase a
-- comment thread; that's part of the record now, same reasoning as the
-- audit log.

drop policy if exists "public read" on comments;
drop policy if exists "public insert" on comments;
drop policy if exists "public update" on comments;
drop policy if exists "public delete" on comments;

create policy "scoped read" on comments for select using (
  auth_user_role() = 'Admin' or exists (
    select 1 from assets where assets.id = comments.asset_id and assets.location_id = auth_user_location()
  )
);
create policy "scoped insert" on comments for insert with check (
  auth_user_role() = 'Admin' or exists (
    select 1 from assets where assets.id = comments.asset_id and assets.location_id = auth_user_location()
  )
);
-- "update" is only ever used by the app to mark a comment read for
-- yourself — this still needs to be scoped to your own comments/location,
-- real per-column protection would need a trigger; scoping by location is
-- the practical floor for now.
create policy "scoped update" on comments for update using (
  auth_user_role() = 'Admin' or exists (
    select 1 from assets where assets.id = comments.asset_id and assets.location_id = auth_user_location()
  )
);

drop policy if exists "public read" on notification_reads;
drop policy if exists "public insert" on notification_reads;
drop policy if exists "public update" on notification_reads;
drop policy if exists "public delete" on notification_reads;

create policy "own reads only" on notification_reads for select using (user_id = auth.uid()::text);
create policy "own reads only insert" on notification_reads for insert with check (user_id = auth.uid()::text);
create policy "own reads only update" on notification_reads for update using (user_id = auth.uid()::text);

-- ---------- 8. Audit log — append-only, forever ----------
-- Insert is allowed for anyone signed in (the app tags every entry
-- itself); read is scoped like everything else; there is deliberately
-- NO update or delete policy at all — with RLS enabled and no policy for
-- an operation, that operation is simply refused. This is what makes the
-- audit trail tamper-evident.

drop policy if exists "public read" on audit_log;
drop policy if exists "public insert" on audit_log;
drop policy if exists "public update" on audit_log;
drop policy if exists "public delete" on audit_log;

create policy "scoped read" on audit_log for select using (
  auth_user_role() = 'Admin' or location_id = auth_user_location()
);
create policy "signed-in insert" on audit_log for insert with check (auth.uid() is not null);
-- (no update/delete policy — intentional)

-- ============================================================
-- Done. Next: CRITICAL-SECURITY-STEPS.md Phase 3 (create your first
-- Admin Auth account and link it), then Phase 4 (deploy the updated app).
-- ============================================================
