-- ============================================================
-- Adds the notes_log column that powers the new multi-entry,
-- timestamped Notes feature in the Asset Details modal (separate
-- from the existing single "notes" text field, which is left
-- untouched for backward compatibility).
--
-- Run in Supabase → SQL Editor → New query → Run.
-- Safe to re-run.
-- ============================================================

alter table assets add column if not exists notes_log jsonb default '[]'::jsonb;
