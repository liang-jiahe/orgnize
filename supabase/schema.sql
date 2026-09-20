-- 在 Supabase 的 SQL Editor 中完整执行一次。
create table if not exists public.fortress_state (
  id text primary key,
  members jsonb not null default '[]'::jsonb,
  queues jsonb not null default '{}'::jsonb,
  last_sweep text not null default '',
  contest boolean not null default false,
  updated_at timestamptz not null default now()
);

alter table public.fortress_state enable row level security;
grant select, insert, update on public.fortress_state to anon, authenticated;

drop policy if exists "shared state can be read" on public.fortress_state;
drop policy if exists "shared state can be created" on public.fortress_state;
drop policy if exists "shared state can be updated" on public.fortress_state;

create policy "shared state can be read"
  on public.fortress_state for select to anon, authenticated
  using (id = 'main');

create policy "shared state can be created"
  on public.fortress_state for insert to anon, authenticated
  with check (id = 'main');

create policy "shared state can be updated"
  on public.fortress_state for update to anon, authenticated
  using (id = 'main')
  with check (id = 'main');

-- 允许已打开的页面通过 Supabase Realtime 接收其他设备的更新。
do $$
begin
  if not exists (
    select 1
    from pg_publication_rel pr
    join pg_class c on c.oid = pr.prrelid
    join pg_namespace n on n.oid = c.relnamespace
    join pg_publication p on p.oid = pr.prpubid
    where p.pubname = 'supabase_realtime'
      and n.nspname = 'public'
      and c.relname = 'fortress_state'
  ) then
    alter publication supabase_realtime add table public.fortress_state;
  end if;
end $$;
