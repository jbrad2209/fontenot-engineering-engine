-- Saved Quore weeks for admin history. One row per reporting period.
-- Uploading that same week again replaces the saved copy.
create table if not exists fontenot_archive (
  id text primary key,
  period_start text not null,
  period_end text not null,
  label text not null,
  payload text not null,
  updated_at timestamptz not null default now()
);
