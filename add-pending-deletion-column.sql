-- Run this in Supabase SQL Editor.
-- Adds the "pending_deletion" column to the existing assets table.
-- This is what powers the Regional Staff "request deletion" / Admin
-- "Approvals" workflow — without it, a deletion request looks like it
-- saved in the requester's own browser but never actually reaches the
-- database, so the Admin never sees it under Approvals.
-- Safe to run even if the column already exists.

alter table assets add column if not exists pending_deletion jsonb;
