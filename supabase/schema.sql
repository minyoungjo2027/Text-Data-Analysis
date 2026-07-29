create table if not exists public.analysis_sessions (
  id uuid primary key default gen_random_uuid(),
  comments jsonb not null,
  vocabulary jsonb not null default '[]'::jsonb,
  similarity_matrix jsonb not null default '[]'::jsonb,
  last_step integer not null default 1 check (last_step between 1 and 7),
  created_at timestamptz not null default now()
);

alter table public.analysis_sessions enable row level security;

grant select, insert on table public.analysis_sessions to anon, authenticated;

drop policy if exists "anonymous users can save analysis sessions" on public.analysis_sessions;
create policy "anonymous users can save analysis sessions"
  on public.analysis_sessions for insert
  to anon, authenticated
  with check (true);

drop policy if exists "anonymous users can read analysis sessions" on public.analysis_sessions;
create policy "anonymous users can read analysis sessions"
  on public.analysis_sessions for select
  to anon, authenticated
  using (true);
