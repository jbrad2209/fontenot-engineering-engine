-- Invited @ihg.com viewers. The Fontenot admin account is separate and is not required
-- to use an IHG address. status: invited, verified, revoked.
create table if not exists fontenot_invite (
  email text primary key,
  token text not null unique,
  status text not null default 'invited',
  user_id text,
  created_at timestamptz not null default now(),
  verified_at timestamptz
);
