-- Run this in Supabase SQL Editor. Your assets table already exists from
-- the earlier setup — this just adds the new "Manufactured Year / Year
-- Model" column to it. Safe to run even if the column already exists.
alter table assets add column if not exists year_model text;
