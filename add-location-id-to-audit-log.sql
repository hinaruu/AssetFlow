-- Run this in Supabase SQL Editor.
-- Adds the "location_id" column to the existing audit_log table.
--
-- Without this column, every activity log entry loses its location as
-- soon as it round-trips through the database: the app tags each entry
-- with the location it happened in (see withLog in App.jsx), but that
-- tag was never actually saved or read back, so it came back as
-- undefined every time. Activity Log filters entries by
-- `entry.locationId === scopedLocationId` for any non-Admin (Regional
-- Admin or Regional Staff) — with locationId always undefined, that
-- comparison never matched, so Regional Admins and Regional Staff saw
-- an empty Activity Log even for their own actions.
--
-- Safe to run even if the column already exists.

alter table audit_log add column if not exists location_id text;
create index if not exists idx_audit_log_location on audit_log(location_id);
