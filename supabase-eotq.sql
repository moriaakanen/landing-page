-- Employee of the Quarter (EoTQ)
-- Jalankan di Supabase SQL Editor sebelum memakai menu Admin > EoTQ.
-- Aplikasi ini memakai auth custom/localStorage + anon REST key, sehingga
-- RLS sengaja dinonaktifkan mengikuti pola tabel operasional lain di portal.

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
  constraint eotq_cycles_time_check check (end_at > start_at)
);

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

alter table public.eotq_cycles disable row level security;
alter table public.eotq_nominees disable row level security;
alter table public.eotq_questions disable row level security;
alter table public.eotq_responses disable row level security;

grant select, insert, update, delete on public.eotq_cycles to anon, authenticated;
grant select, insert, update, delete on public.eotq_nominees to anon, authenticated;
grant select, insert, update, delete on public.eotq_questions to anon, authenticated;
grant select, insert, update, delete on public.eotq_responses to anon, authenticated;
grant usage, select on sequence public.eotq_cycles_id_seq to anon, authenticated;
grant usage, select on sequence public.eotq_nominees_id_seq to anon, authenticated;
grant usage, select on sequence public.eotq_questions_id_seq to anon, authenticated;
grant usage, select on sequence public.eotq_responses_id_seq to anon, authenticated;
