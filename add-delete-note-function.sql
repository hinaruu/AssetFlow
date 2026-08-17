-- ============================================================
-- Adds server-side enforced note deletion.
--
-- Why an RPC function instead of a normal RLS-guarded delete: the
-- Notes log lives as a JSON array inside a single column on the
-- assets row (notes_log), not as its own table with its own rows.
-- RLS can only govern access to the row as a whole ("is this
-- asset in your location") — it has no concept of "which entry
-- inside this JSON array, written by whom." So per-note authorship
-- ("only the author, or an Admin, can delete this specific note")
-- has to be checked in a function, the same way transfer_asset
-- does its own explicit check rather than relying on RLS alone.
--
-- Run in Supabase → SQL Editor → New query → Run (Role: postgres).
-- Safe to re-run.
-- ============================================================

create or replace function delete_asset_note(p_asset_id text, p_note_id text)
returns void
language plpgsql
security definer
set search_path = public
as $$
declare
  caller_role text;
  caller_location text;
  caller_profile_id text;
  asset_row assets%rowtype;
  note jsonb;
  note_found boolean := false;
begin
  select id, role, location_id into caller_profile_id, caller_role, caller_location
  from users where auth_user_id = auth.uid();

  if caller_profile_id is null then
    raise exception 'No matching profile for this account.';
  end if;

  select * into asset_row from assets where id = p_asset_id;
  if asset_row.id is null then
    raise exception 'Asset not found.';
  end if;

  -- Must have ordinary access to this asset at all (own location, or Admin).
  if caller_role <> 'Admin' and asset_row.location_id <> caller_location then
    raise exception 'You do not have access to this asset.';
  end if;

  -- Per-note authorship check: Admin can delete any note; everyone else
  -- only their own.
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

revoke all on function delete_asset_note(text, text) from public;
grant execute on function delete_asset_note(text, text) to authenticated;
