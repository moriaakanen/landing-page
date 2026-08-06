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

-- Revoke hak akses tulis langsung (INSERT, UPDATE, DELETE) dari peran anon & authenticated.
revoke insert, update, delete on public.eotq_cycles from anon, authenticated;
revoke insert, update, delete on public.eotq_nominees from anon, authenticated;
revoke insert, update, delete on public.eotq_questions from anon, authenticated;
revoke insert, update, delete on public.eotq_responses from anon, authenticated;

-- Berikan izin SELECT agar aplikasi dapat membaca data
grant select on public.eotq_cycles to anon, authenticated;
grant select on public.eotq_nominees to anon, authenticated;
grant select on public.eotq_questions to anon, authenticated;
grant select on public.eotq_responses to anon, authenticated;
grant usage, select on sequence public.eotq_cycles_id_seq to anon, authenticated;
grant usage, select on sequence public.eotq_nominees_id_seq to anon, authenticated;
grant usage, select on sequence public.eotq_questions_id_seq to anon, authenticated;
grant usage, select on sequence public.eotq_responses_id_seq to anon, authenticated;

drop policy if exists eotq_cycles_select on public.eotq_cycles;
drop policy if exists eotq_cycles_insert on public.eotq_cycles;
drop policy if exists eotq_cycles_update on public.eotq_cycles;
drop policy if exists eotq_cycles_delete on public.eotq_cycles;
create policy eotq_cycles_select on public.eotq_cycles for select to anon, authenticated using (true);

drop policy if exists eotq_nominees_select on public.eotq_nominees;
drop policy if exists eotq_nominees_insert on public.eotq_nominees;
drop policy if exists eotq_nominees_update on public.eotq_nominees;
drop policy if exists eotq_nominees_delete on public.eotq_nominees;
create policy eotq_nominees_select on public.eotq_nominees for select to anon, authenticated using (true);

drop policy if exists eotq_questions_select on public.eotq_questions;
drop policy if exists eotq_questions_insert on public.eotq_questions;
drop policy if exists eotq_questions_update on public.eotq_questions;
drop policy if exists eotq_questions_delete on public.eotq_questions;
create policy eotq_questions_select on public.eotq_questions for select to anon, authenticated using (true);

drop policy if exists eotq_responses_select on public.eotq_responses;
drop policy if exists eotq_responses_insert on public.eotq_responses;
drop policy if exists eotq_responses_update on public.eotq_responses;
drop policy if exists eotq_responses_delete on public.eotq_responses;
create policy eotq_responses_select on public.eotq_responses for select to anon, authenticated using (true);

-- RPC Functions untuk Admin & Voting EoTQ
create or replace function public.eotq_save_cycle(
  p_caller_id bigint,
  p_id bigint default null,
  p_title text default 'Employee of the Quarter',
  p_quarter_label text default null,
  p_description text default null,
  p_start_at timestamptz default null,
  p_end_at timestamptz default null,
  p_announce_at timestamptz default null,
  p_status text default 'draft'
) returns jsonb language plpgsql security definer set search_path = public as $$
declare v_is_admin boolean; v_result jsonb;
begin
  select exists (select 1 from public.users u where u.id = p_caller_id and (u.role = 'admin' or coalesce(u.roles, array[]::text[]) @> array['admin']::text[])) into v_is_admin;
  if not coalesce(v_is_admin, false) then raise exception 'admin access required'; end if;
  if p_id is not null and p_id > 0 then
    update public.eotq_cycles set title = coalesce(p_title, title), quarter_label = coalesce(p_quarter_label, quarter_label), description = p_description, start_at = coalesce(p_start_at, start_at), end_at = coalesce(p_end_at, end_at), announce_at = p_announce_at, status = coalesce(p_status, status), updated_at = now() where id = p_id returning to_jsonb(public.eotq_cycles.*) into v_result;
  else
    insert into public.eotq_cycles (title, quarter_label, description, start_at, end_at, announce_at, status) values (coalesce(p_title, 'Employee of the Quarter'), p_quarter_label, p_description, p_start_at, p_end_at, p_announce_at, coalesce(p_status, 'draft')) returning to_jsonb(public.eotq_cycles.*) into v_result;
  end if;
  return v_result;
end; $$;

create or replace function public.eotq_delete_cycle(p_caller_id bigint, p_id bigint) returns boolean language plpgsql security definer set search_path = public as $$
declare v_is_admin boolean;
begin
  select exists (select 1 from public.users u where u.id = p_caller_id and (u.role = 'admin' or coalesce(u.roles, array[]::text[]) @> array['admin']::text[])) into v_is_admin;
  if not coalesce(v_is_admin, false) then raise exception 'admin access required'; end if;
  delete from public.eotq_cycles where id = p_id;
  return true;
end; $$;

create or replace function public.eotq_save_nominee(
  p_caller_id bigint, p_id bigint default null, p_cycle_id bigint default null, p_pegawai_nip text default null, p_pegawai_nama text default null, p_note text default null, p_sort_order integer default 0
) returns jsonb language plpgsql security definer set search_path = public as $$
declare v_is_admin boolean; v_result jsonb;
begin
  select exists (select 1 from public.users u where u.id = p_caller_id and (u.role = 'admin' or coalesce(u.roles, array[]::text[]) @> array['admin']::text[])) into v_is_admin;
  if not coalesce(v_is_admin, false) then raise exception 'admin access required'; end if;
  if p_id is not null and p_id > 0 then
    update public.eotq_nominees set note = p_note, sort_order = coalesce(p_sort_order, sort_order) where id = p_id returning to_jsonb(public.eotq_nominees.*) into v_result;
  else
    insert into public.eotq_nominees (cycle_id, pegawai_nip, pegawai_nama, note, sort_order) values (p_cycle_id, p_pegawai_nip, p_pegawai_nama, p_note, coalesce(p_sort_order, 0)) returning to_jsonb(public.eotq_nominees.*) into v_result;
  end if;
  return v_result;
end; $$;

create or replace function public.eotq_delete_nominee(p_caller_id bigint, p_id bigint) returns boolean language plpgsql security definer set search_path = public as $$
declare v_is_admin boolean;
begin
  select exists (select 1 from public.users u where u.id = p_caller_id and (u.role = 'admin' or coalesce(u.roles, array[]::text[]) @> array['admin']::text[])) into v_is_admin;
  if not coalesce(v_is_admin, false) then raise exception 'admin access required'; end if;
  delete from public.eotq_nominees where id = p_id;
  return true;
end; $$;

create or replace function public.eotq_save_question(
  p_caller_id bigint, p_id bigint default null, p_cycle_id bigint default null, p_question text default null, p_type text default 'rating', p_required boolean default true, p_weight numeric default 1, p_options jsonb default '[]'::jsonb, p_sort_order integer default 0
) returns jsonb language plpgsql security definer set search_path = public as $$
declare v_is_admin boolean; v_result jsonb;
begin
  select exists (select 1 from public.users u where u.id = p_caller_id and (u.role = 'admin' or coalesce(u.roles, array[]::text[]) @> array['admin']::text[])) into v_is_admin;
  if not coalesce(v_is_admin, false) then raise exception 'admin access required'; end if;
  if p_id is not null and p_id > 0 then
    update public.eotq_questions set question = coalesce(p_question, question), type = coalesce(p_type, type), required = coalesce(p_required, required), weight = coalesce(p_weight, weight), options = coalesce(p_options, options), sort_order = coalesce(p_sort_order, sort_order) where id = p_id returning to_jsonb(public.eotq_questions.*) into v_result;
  else
    insert into public.eotq_questions (cycle_id, question, type, required, weight, options, sort_order) values (p_cycle_id, p_question, coalesce(p_type, 'rating'), coalesce(p_required, true), coalesce(p_weight, 1), coalesce(p_options, '[]'::jsonb), coalesce(p_sort_order, 0)) returning to_jsonb(public.eotq_questions.*) into v_result;
  end if;
  return v_result;
end; $$;

create or replace function public.eotq_delete_question(p_caller_id bigint, p_id bigint) returns boolean language plpgsql security definer set search_path = public as $$
declare v_is_admin boolean;
begin
  select exists (select 1 from public.users u where u.id = p_caller_id and (u.role = 'admin' or coalesce(u.roles, array[]::text[]) @> array['admin']::text[])) into v_is_admin;
  if not coalesce(v_is_admin, false) then raise exception 'admin access required'; end if;
  delete from public.eotq_questions where id = p_id;
  return true;
end; $$;

create or replace function public.eotq_submit_vote(
  p_voter_user_id text, p_cycle_id bigint, p_nominee_id bigint, p_voter_nip text default null, p_voter_name text default null, p_answers jsonb default '[]'::jsonb, p_total_score numeric default 0
) returns jsonb language plpgsql security definer set search_path = public as $$
declare v_cycle_status text; v_start_at timestamptz; v_end_at timestamptz; v_result jsonb;
begin
  select status, start_at, end_at into v_cycle_status, v_start_at, v_end_at from public.eotq_cycles where id = p_cycle_id;
  if v_cycle_status is null then raise exception 'Periode EoTQ tidak ditemukan'; end if;
  if v_cycle_status <> 'published' then raise exception 'Periode EoTQ tidak aktif atau belum dipublikasikan'; end if;
  if now() < v_start_at or now() > v_end_at then raise exception 'Sesi penilaian EoTQ sudah ditutup atau belum dimulai'; end if;
  insert into public.eotq_responses (cycle_id, nominee_id, voter_user_id, voter_nip, voter_name, answers, total_score, updated_at)
  values (p_cycle_id, p_nominee_id, p_voter_user_id, p_voter_nip, p_voter_name, coalesce(p_answers, '[]'::jsonb), coalesce(p_total_score, 0), now())
  on conflict (cycle_id, nominee_id, voter_user_id) do update set voter_nip = excluded.voter_nip, voter_name = excluded.voter_name, answers = excluded.answers, total_score = excluded.total_score, updated_at = now()
  returning to_jsonb(public.eotq_responses.*) into v_result;
  return v_result;
end; $$;

grant execute on function public.eotq_save_cycle to anon, authenticated;
grant execute on function public.eotq_delete_cycle to anon, authenticated;
grant execute on function public.eotq_save_nominee to anon, authenticated;
grant execute on function public.eotq_delete_nominee to anon, authenticated;
grant execute on function public.eotq_save_question to anon, authenticated;
grant execute on function public.eotq_delete_question to anon, authenticated;
grant execute on function public.eotq_submit_vote to anon, authenticated;

