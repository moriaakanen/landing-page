-- Jalankan file ini sekali di Supabase SQL Editor.
-- Tujuan: membuat RPC yang dipanggil halaman manajemen-pengguna.html.

create extension if not exists pgcrypto;

drop function if exists public.buat_pengguna(text, boolean, text, text, text[], text);

create or replace function public.buat_pengguna(
  p_full_name text,
  p_must_change boolean,
  p_password text,
  p_role text,
  p_roles text[],
  p_username text
)
returns jsonb
language plpgsql
security definer
set search_path = public
as $$
declare
  v_username text := lower(trim(coalesce(p_username, '')));
  v_full_name text := trim(coalesce(p_full_name, ''));
  v_role text := lower(trim(coalesce(p_role, 'user')));
  v_roles text[] := coalesce(p_roles, array['user']::text[]);
  v_id bigint;
begin
  if v_username = '' then
    raise exception 'Username wajib diisi';
  end if;

  if v_username ~ '\s' then
    raise exception 'Username tidak boleh mengandung spasi';
  end if;

  if v_full_name = '' then
    raise exception 'Nama lengkap wajib diisi';
  end if;

  if coalesce(p_password, '') = '' or length(p_password) < 6 then
    raise exception 'Password minimal 6 karakter';
  end if;

  select array_agg(distinct lower(trim(x)))
  into v_roles
  from unnest(v_roles) as x
  where lower(trim(x)) in ('admin', 'user');

  if v_roles is null or array_length(v_roles, 1) is null then
    v_roles := array['user']::text[];
  end if;

  if not (v_role = any(v_roles)) then
    v_role := v_roles[1];
  end if;

  if exists (select 1 from public.users where lower(username) = v_username) then
    raise exception 'Username sudah digunakan';
  end if;

  insert into public.users (
    username,
    full_name,
    password_hash,
    must_change_password,
    role,
    roles
  )
  values (
    v_username,
    v_full_name,
    crypt(p_password, gen_salt('bf', 10)),
    coalesce(p_must_change, true),
    v_role,
    v_roles
  )
  returning id into v_id;

  return jsonb_build_object(
    'id', v_id,
    'username', v_username,
    'full_name', v_full_name,
    'must_change_password', coalesce(p_must_change, true),
    'role', v_role,
    'roles', v_roles
  );
end;
$$;

revoke all on function public.buat_pengguna(text, boolean, text, text, text[], text) from public;
grant execute on function public.buat_pengguna(text, boolean, text, text, text[], text) to anon, authenticated;
