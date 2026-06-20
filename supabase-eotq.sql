-- Employee of the Quarter (EoTQ)
-- Jalankan di Supabase SQL Editor sebelum memakai menu Admin > EoTQ.
-- Aplikasi ini memakai auth custom/localStorage + anon REST key. Supabase
-- tetap diberi RLS policy permisif agar request REST dari aplikasi tidak
-- ditolak oleh policy Storage/Database saat tabel dibuat dengan RLS aktif.

create table if not exists public.eotq_cycles (
  id bigserial primary key,
  title text not null default 'Employee of the Quarter',
  quarter_label text not null,
  description text,
  start_at timestamptz not null,
  end_at timestamptz not null,
  announce_at timestamptz,
  status text not null default 'draft',
  created_by text,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now(),
  constraint eotq_cycles_status_check check (status in ('draft', 'published', 'archived')),
  constraint eotq_cycles_time_check check (end_at > start_at)
);

do $$
begin
  if not exists (
    select 1 from pg_constraint
    where conname = 'eotq_cycles_status_check'
      and conrelid = 'public.eotq_cycles'::regclass
  ) then
    alter table public.eotq_cycles
      add constraint eotq_cycles_status_check
      check (status in ('draft', 'published', 'archived'));
  end if;
end $$;

create table if not exists public.eotq_nominees (
  id bigserial primary key,
  cycle_id bigint not null references public.eotq_cycles(id) on delete cascade,
  pegawai_nip text not null,
  pegawai_nama text not null,
  note text,
  sort_order integer not null default 0,
  created_at timestamptz not null default now(),
  unique (cycle_id, pegawai_nip)
);

create table if not exists public.eotq_questions (
  id bigserial primary key,
  cycle_id bigint not null references public.eotq_cycles(id) on delete cascade,
  question text not null,
  type text not null default 'rating',
  required boolean not null default true,
  weight numeric not null default 1,
  options jsonb not null default '[]'::jsonb,
  sort_order integer not null default 0,
  created_at timestamptz not null default now(),
  constraint eotq_questions_type_check check (type in ('rating','single','multi','text'))
);

create table if not exists public.eotq_responses (
  id bigserial primary key,
  cycle_id bigint not null references public.eotq_cycles(id) on delete cascade,
  nominee_id bigint not null references public.eotq_nominees(id) on delete cascade,
  voter_user_id text not null,
  voter_nip text,
  voter_name text,
  answers jsonb not null default '[]'::jsonb,
  total_score numeric not null default 0,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now(),
  unique (cycle_id, nominee_id, voter_user_id)
);

create index if not exists eotq_nominees_cycle_idx on public.eotq_nominees(cycle_id);
create index if not exists eotq_questions_cycle_idx on public.eotq_questions(cycle_id);
create index if not exists eotq_responses_cycle_idx on public.eotq_responses(cycle_id);
create index if not exists eotq_responses_nominee_idx on public.eotq_responses(nominee_id);

create or replace function public.set_eotq_updated_at()
returns trigger
language plpgsql
as $$
begin
  new.updated_at = now();
  return new;
end;
$$;

drop trigger if exists trg_eotq_cycles_updated_at on public.eotq_cycles;
create trigger trg_eotq_cycles_updated_at
before update on public.eotq_cycles
for each row execute function public.set_eotq_updated_at();

drop trigger if exists trg_eotq_responses_updated_at on public.eotq_responses;
create trigger trg_eotq_responses_updated_at
before update on public.eotq_responses
for each row execute function public.set_eotq_updated_at();

alter table public.eotq_cycles enable row level security;
alter table public.eotq_nominees enable row level security;
alter table public.eotq_questions enable row level security;
alter table public.eotq_responses enable row level security;

drop policy if exists eotq_cycles_select on public.eotq_cycles;
drop policy if exists eotq_cycles_insert on public.eotq_cycles;
drop policy if exists eotq_cycles_update on public.eotq_cycles;
drop policy if exists eotq_cycles_delete on public.eotq_cycles;
create policy eotq_cycles_select on public.eotq_cycles for select to anon, authenticated using (true);
create policy eotq_cycles_insert on public.eotq_cycles for insert to anon, authenticated with check (true);
create policy eotq_cycles_update on public.eotq_cycles for update to anon, authenticated using (true) with check (true);
create policy eotq_cycles_delete on public.eotq_cycles for delete to anon, authenticated using (true);

drop policy if exists eotq_nominees_select on public.eotq_nominees;
drop policy if exists eotq_nominees_insert on public.eotq_nominees;
drop policy if exists eotq_nominees_update on public.eotq_nominees;
drop policy if exists eotq_nominees_delete on public.eotq_nominees;
create policy eotq_nominees_select on public.eotq_nominees for select to anon, authenticated using (true);
create policy eotq_nominees_insert on public.eotq_nominees for insert to anon, authenticated with check (true);
create policy eotq_nominees_update on public.eotq_nominees for update to anon, authenticated using (true) with check (true);
create policy eotq_nominees_delete on public.eotq_nominees for delete to anon, authenticated using (true);

drop policy if exists eotq_questions_select on public.eotq_questions;
drop policy if exists eotq_questions_insert on public.eotq_questions;
drop policy if exists eotq_questions_update on public.eotq_questions;
drop policy if exists eotq_questions_delete on public.eotq_questions;
create policy eotq_questions_select on public.eotq_questions for select to anon, authenticated using (true);
create policy eotq_questions_insert on public.eotq_questions for insert to anon, authenticated with check (true);
create policy eotq_questions_update on public.eotq_questions for update to anon, authenticated using (true) with check (true);
create policy eotq_questions_delete on public.eotq_questions for delete to anon, authenticated using (true);

drop policy if exists eotq_responses_select on public.eotq_responses;
drop policy if exists eotq_responses_insert on public.eotq_responses;
drop policy if exists eotq_responses_update on public.eotq_responses;
drop policy if exists eotq_responses_delete on public.eotq_responses;
create policy eotq_responses_select on public.eotq_responses for select to anon, authenticated using (true);
create policy eotq_responses_insert on public.eotq_responses for insert to anon, authenticated with check (true);
create policy eotq_responses_update on public.eotq_responses for update to anon, authenticated using (true) with check (true);
create policy eotq_responses_delete on public.eotq_responses for delete to anon, authenticated using (true);

grant select, insert, update, delete on public.eotq_cycles to anon, authenticated;
grant select, insert, update, delete on public.eotq_nominees to anon, authenticated;
grant select, insert, update, delete on public.eotq_questions to anon, authenticated;
grant select, insert, update, delete on public.eotq_responses to anon, authenticated;
grant usage, select on sequence public.eotq_cycles_id_seq to anon, authenticated;
grant usage, select on sequence public.eotq_nominees_id_seq to anon, authenticated;
grant usage, select on sequence public.eotq_questions_id_seq to anon, authenticated;
grant usage, select on sequence public.eotq_responses_id_seq to anon, authenticated;
