-- One shared hotel week. Anyone with the link can read it. Only the admin can replace it.
create table if not exists fontenot_published (
  id integer primary key,
  payload text not null,
  updated_at timestamptz not null default now()
);
