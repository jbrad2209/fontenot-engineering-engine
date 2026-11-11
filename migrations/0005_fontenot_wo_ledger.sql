-- Running All Work Orders list. The YTD file is the base. Weekly files are added on top.
create table if not exists fontenot_wo_ledger (
  id integer primary key,
  payload text not null,
  updated_at timestamptz not null default now()
);
