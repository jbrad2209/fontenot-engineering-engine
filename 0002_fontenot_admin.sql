-- One Fontenot admin account. The first signed-in user to claim it is the only one.
create table if not exists fontenot_admin (
  user_id text primary key,
  created_at timestamptz not null default now()
);
