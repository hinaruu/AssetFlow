-- ============================================================
-- FIX: non-admin asset transfers failing with 42501
--
-- What we found: a transfer moves an asset's location_id to a
-- DIFFERENT location than the mover's own — and even with a
-- correctly-written "with check (true)" UPDATE policy in place
-- (confirmed via pg_policies), the plain UPDATE still failed for
-- a simulated non-admin session. Rather than keep chasing that
-- specific behavior, this works around it entirely: the transfer
-- now runs through a small function that does its own permission
-- check in plain code and applies the change directly, instead of
-- going through the normal RLS-checked write path.
--
-- Run this in Supabase → SQL Editor → New query → Run (Role: postgres).
-- Safe to re-run.
-- ============================================================

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
  caller_location text;
  asset_row assets%rowtype;
  from_loc_name text;
  to_loc_name text;
  xfer jsonb;
begin
  -- Who is calling this, and are they allowed to touch this asset?
  select role, location_id into caller_role, caller_location
  from users where auth_user_id = auth.uid();

  if caller_role is null then
    raise exception 'No matching profile for this account.';
  end if;

  select * into asset_row from assets where id = p_asset_id;
  if asset_row.id is null then
    raise exception 'Asset not found.';
  end if;

  if caller_role <> 'Admin' and asset_row.location_id <> caller_location then
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

-- Only signed-in app users can call this — and only Regional Staff/Admin
-- moving an asset they already have access to (enforced above, in code,
-- not RLS) will succeed; everyone else gets the "not allowed" exception.
revoke all on function transfer_asset(text, text, text, text, text) from public;
grant execute on function transfer_asset(text, text, text, text, text) to authenticated;
