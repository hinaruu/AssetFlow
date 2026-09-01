-- ============================================================
-- Multi-location users (Overall Admin only)
--
-- Adds the ability for the Overall Admin to associate a user with
-- one or more ADDITIONAL locations, on top of their existing
-- single Primary Location (users.location_id, unchanged).
--
-- Data model: a new junction table, user_locations, rather than
-- duplicating location records or turning users.location_id into
-- an array — this keeps every existing single-location code path
-- (defaults, "New Asset" location lock, profile summaries, etc.)
-- working exactly as it does today against the primary location,
-- and layers "also has access to" on top as a separate, optional
-- relationship. Each row's id is deterministic (`user_id::location_id`)
-- so assigning the same location twice is a no-op upsert, not a
-- duplicate row.
--
-- Run in Supabase → SQL Editor → New query → Run (Role: postgres).
-- Safe to re-run.
-- ============================================================

create table if not exists user_locations (
  id text primary key,
  user_id text not null references users(id) on delete cascade,
  location_id text not null references locations(id) on delete cascade,
  unique (user_id, location_id)
);

alter table user_locations enable row level security;

-- Anyone signed in can read this — needed so a multi-location user's own
-- session can compute their full location set, and so the User Accounts
-- screen can display everyone's assignments. Only the Overall Admin can
-- write to it, enforced here at the database level — not just by hiding
-- the "Additional Locations" control in the UI. A Regional Admin or
-- Regional Staff account calling this table directly (bypassing the app
-- entirely) gets rejected the same as if they'd used the UI.
drop policy if exists "signed-in read" on user_locations;
drop policy if exists "admin write" on user_locations;
drop policy if exists "admin update" on user_locations;
drop policy if exists "admin delete" on user_locations;
create policy "signed-in read" on user_locations for select using (auth.uid() is not null);
create policy "admin write" on user_locations for insert with check (auth_user_role() = 'Admin');
create policy "admin update" on user_locations for update using (auth_user_role() = 'Admin');
create policy "admin delete" on user_locations for delete using (auth_user_role() = 'Admin');

-- Every location a user has access to: their primary (users.location_id)
-- plus any rows here. Existing single-location users are unaffected —
-- this just returns a one-element array for them, same as before.
create or replace function auth_user_locations() returns text[]
language sql stable security definer set search_path = public as $$
  select coalesce(array_agg(distinct loc), '{}'::text[]) from (
    select location_id as loc from users
      where auth_user_id = auth.uid() and location_id is not null
    union
    select ul.location_id as loc from user_locations ul
      join users u on u.id = ul.user_id
      where u.auth_user_id = auth.uid()
  ) combined
$$;

-- ---------- Assets: read/update/delete now check the full set ----------
-- (insert stays primary-only via auth_user_location(), unchanged — new
-- assets still default and lock to the Primary Location, per the
-- existing "New Asset" behavior. This only affects what an already-
-- existing asset a multi-location user can see, edit, or delete.)

drop policy if exists "scoped read" on assets;
create policy "scoped read" on assets for select using (
  auth_user_role() = 'Admin' or location_id = any(auth_user_locations())
);

drop policy if exists "scoped update" on assets;
create policy "scoped update" on assets for update
  using (auth_user_role() = 'Admin' or location_id = any(auth_user_locations()))
  with check (true);

drop policy if exists "admin or regional-admin delete" on assets;
create policy "admin or regional-admin delete" on assets for delete using (
  auth_user_role() = 'Admin'
  or (auth_user_role() = 'Regional Admin' and location_id = any(auth_user_locations()))
);

-- ---------- Maintenance: tied to an asset the user can already see ----------

drop policy if exists "scoped read" on maintenance;
create policy "scoped read" on maintenance for select using (
  auth_user_role() = 'Admin' or exists (
    select 1 from assets where assets.id = maintenance.asset_id and assets.location_id = any(auth_user_locations())
  )
);
drop policy if exists "scoped write" on maintenance;
create policy "scoped write" on maintenance for insert with check (
  auth_user_role() = 'Admin' or exists (
    select 1 from assets where assets.id = maintenance.asset_id and assets.location_id = any(auth_user_locations())
  )
);
drop policy if exists "scoped update" on maintenance;
create policy "scoped update" on maintenance for update using (
  auth_user_role() = 'Admin' or exists (
    select 1 from assets where assets.id = maintenance.asset_id and assets.location_id = any(auth_user_locations())
  )
);
drop policy if exists "admin or regional-admin delete" on maintenance;
create policy "admin or regional-admin delete" on maintenance for delete using (
  auth_user_role() = 'Admin' or (
    auth_user_role() = 'Regional Admin' and exists (
      select 1 from assets where assets.id = maintenance.asset_id and assets.location_id = any(auth_user_locations())
    )
  )
);

-- ---------- Comments: same reasoning as maintenance ----------

drop policy if exists "scoped read" on comments;
create policy "scoped read" on comments for select using (
  auth_user_role() = 'Admin' or exists (
    select 1 from assets where assets.id = comments.asset_id and assets.location_id = any(auth_user_locations())
  )
);
drop policy if exists "scoped insert" on comments;
create policy "scoped insert" on comments for insert with check (
  auth_user_role() = 'Admin' or exists (
    select 1 from assets where assets.id = comments.asset_id and assets.location_id = any(auth_user_locations())
  )
);
drop policy if exists "scoped update" on comments;
create policy "scoped update" on comments for update using (
  auth_user_role() = 'Admin' or exists (
    select 1 from assets where assets.id = comments.asset_id and assets.location_id = any(auth_user_locations())
  )
);

-- ---------- Audit log: visibility only (insert stays unrestricted) ----------

drop policy if exists "scoped read" on audit_log;
create policy "scoped read" on audit_log for select using (
  auth_user_role() = 'Admin' or location_id = any(auth_user_locations())
);

-- ---------- transfer_asset() / delete_asset_note(): check the full set ----------
-- These run as security definer and did their own explicit permission
-- check rather than relying on RLS (see add-transfer-asset-function.sql
-- and add-delete-note-function.sql) — updated here to check membership
-- in the caller's full location set instead of just their primary.

create or replace function transfer_asset(
  p_asset_id text,
  p_new_location_id text,
  p_new_assigned_to text,
  p_reason text,
  p_by_name text
)
returns void
language plpgsql
security definer
set search_path = public
as $$
declare
  caller_role text;
  caller_locations text[];
  asset_row assets%rowtype;
  from_loc_name text;
  to_loc_name text;
  xfer jsonb;
begin
  select role into caller_role from users where auth_user_id = auth.uid();
  if caller_role is null then
    raise exception 'No matching profile for this account.';
  end if;
  caller_locations := auth_user_locations();

  select * into asset_row from assets where id = p_asset_id;
  if asset_row.id is null then
    raise exception 'Asset not found.';
  end if;

  if caller_role <> 'Admin' and not (asset_row.location_id = any(caller_locations)) then
    raise exception 'You do not have access to this asset.';
  end if;

  select name into from_loc_name from locations where id = asset_row.location_id;
  select name into to_loc_name from locations where id = p_new_location_id;

  xfer := jsonb_build_object(
    'id', 'xfer-' || floor(extract(epoch from clock_timestamp()) * 1000)::text || '-' || substr(md5(random()::text), 1, 5),
    'fromLocationId', asset_row.location_id,
    'fromLocationName', coalesce(from_loc_name, 'Unknown'),
    'toLocationId', p_new_location_id,
    'toLocationName', coalesce(to_loc_name, 'Unknown'),
    'reason', p_reason,
    'by', p_by_name,
    'at', to_char(now() at time zone 'utc', 'YYYY-MM-DD"T"HH24:MI:SS.MS"Z"')
  );

  update assets set
    location_id = p_new_location_id,
    assigned_to = p_new_assigned_to,
    transfer_history = coalesce(transfer_history, '[]'::jsonb) || jsonb_build_array(xfer),
    updated_at = now()
  where id = p_asset_id;

  insert into audit_log (id, at, user_id, user_name, message, location_id, asset_id)
  values (
    'log-' || floor(extract(epoch from clock_timestamp()) * 1000)::text || '-' || substr(md5(random()::text), 1, 5),
    now(),
    auth.uid()::text,
    p_by_name,
    'Transferred asset "' || coalesce(asset_row.name, asset_row.tag) || '" from '
      || coalesce(from_loc_name, 'Unknown') || ' to ' || coalesce(to_loc_name, 'Unknown')
      || ' — reason: ' || p_reason,
    p_new_location_id,
    p_asset_id
  );
end;
$$;

create or replace function delete_asset_note(p_asset_id text, p_note_id text)
returns void
language plpgsql
security definer
set search_path = public
as $$
declare
  caller_role text;
  caller_locations text[];
  caller_profile_id text;
  asset_row assets%rowtype;
  note jsonb;
  note_found boolean := false;
begin
  select id, role into caller_profile_id, caller_role from users where auth_user_id = auth.uid();
  if caller_profile_id is null then
    raise exception 'No matching profile for this account.';
  end if;
  caller_locations := auth_user_locations();

  select * into asset_row from assets where id = p_asset_id;
  if asset_row.id is null then
    raise exception 'Asset not found.';
  end if;

  if caller_role <> 'Admin' and not (asset_row.location_id = any(caller_locations)) then
    raise exception 'You do not have access to this asset.';
  end if;

  for note in select * from jsonb_array_elements(coalesce(asset_row.notes_log, '[]'::jsonb))
  loop
    if note->>'id' = p_note_id then
      note_found := true;
      if caller_role <> 'Admin' and note->>'authorId' <> caller_profile_id then
        raise exception 'You can only delete your own notes.';
      end if;
    end if;
  end loop;

  if not note_found then
    raise exception 'Note not found.';
  end if;

  update assets
  set notes_log = (
    select coalesce(jsonb_agg(elem), '[]'::jsonb)
    from jsonb_array_elements(assets.notes_log) elem
    where elem->>'id' <> p_note_id
  )
  where id = p_asset_id;

  insert into audit_log (id, at, user_id, user_name, message, location_id, asset_id)
  values (
    'log-' || floor(extract(epoch from clock_timestamp()) * 1000)::text || '-' || substr(md5(random()::text), 1, 5),
    now(),
    auth.uid()::text,
    (select name from users where id = caller_profile_id),
    'Deleted a note on asset "' || coalesce(asset_row.name, asset_row.tag) || '"',
    asset_row.location_id,
    p_asset_id
  );
end;
$$;

revoke all on function transfer_asset(text, text, text, text, text) from public;
grant execute on function transfer_asset(text, text, text, text, text) to authenticated;
revoke all on function delete_asset_note(text, text) from public;
grant execute on function delete_asset_note(text, text) to authenticated;
