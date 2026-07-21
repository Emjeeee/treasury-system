-- Run this once in the Supabase project's SQL Editor
-- (Dashboard > SQL Editor > New query > paste > Run).
--
-- The app stores its entire shared state (accounts, transactions, activity
-- log) as one JSON blob keyed by "ct_shared_v4" in this table, so it survives
-- page reloads / rebuilds and is shared across every device and user.
--
-- NOTE: this app has no server-side auth layer -- all login is handled in the
-- browser with the Supabase anon key, which is always public in client code.
-- The policies below let anyone holding that anon key read/write this table,
-- matching the app's existing all-client-side trust model. Do not put data
-- here you would not want readable by anyone with the site's URL/anon key.

create table if not exists kv_store (
  key text primary key,
  value text not null,
  updated_at timestamptz not null default now()
);

alter table kv_store enable row level security;

create policy "kv_store anon select" on kv_store
  for select to anon using (true);

create policy "kv_store anon insert" on kv_store
  for insert to anon with check (true);

create policy "kv_store anon update" on kv_store
  for update to anon using (true) with check (true);
