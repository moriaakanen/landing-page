-- Portal 9201 RPC remediation draft
-- Review and run after backing up the database. This updates vulnerable RPCs
-- identified from the live Supabase audit export.
--
-- This still does not replace real Supabase Auth/JWT. It only prevents the
-- most obvious anonymous calls to admin RPCs by requiring an admin caller id.

begin;

create or replace function public.is_admin_user(p_user_id bigint)
returns boolean
language sql
security definer
set search_path to 'public'
as $$
  select exists (
    select 1
    from public.users u
    where u.id = p_user_id
      and coalesce(u.must_change_password, false) = false
      and (
        u.role = 'admin'
        or 'admin' = any(coalesce(u.roles, array[]::text[]))
      )
  );
$$;

create or replace function public.buat_pengguna(
  p_caller_id bigint,
  p_username text,
  p_full_name text,
  p_password text,
  p_must_change boolean,
  p_role text,
  p_roles text[]
)
returns bigint
language plpgsql
security definer
set search_path to 'public', 'extensions'
as $$
declare
  v_new_id bigint;
  v_roles text[];
begin
  if not public.is_admin_user(p_caller_id) then
    raise exception 'Akses ditolak: hanya admin yang dapat membuat pengguna'
      using errcode = '42501';
  end if;

  if p_username is null or btrim(p_username) = '' or p_username ~ '\s' then
    raise exception 'Username tidak valid' using errcode = '22023';
  end if;
  if p_password is null or length(p_password) < 6 then
    raise exception 'Password minimal 6 karakter' using errcode = '22023';
  end if;

  v_roles := coalesce(p_roles, array['user']::text[]);
  if array_length(v_roles, 1) is null then
    v_roles := array['user']::text[];
  end if;
  if exists (select 1 from unnest(v_roles) r where r not in ('user', 'admin')) then
    raise exception 'Role tidak valid' using errcode = '22023';
  end if;

  insert into public.users(
    username, full_name, password_hash, must_change_password, role, roles
  )
  values (
    lower(btrim(p_username)),
    nullif(btrim(coalesce(p_full_name, '')), ''),
    extensions.crypt(p_password, extensions.gen_salt('bf', 10)),
    coalesce(p_must_change, true),
    case when 'admin' = any(v_roles) then 'admin' else coalesce(p_role, 'user') end,
    v_roles
  )
  returning id into v_new_id;

  return v_new_id;
end;
$$;

create or replace function public.reset_password_admin(
  p_caller_id bigint,
  p_user_id bigint,
  p_new_password text
)
returns boolean
language plpgsql
security definer
set search_path to 'public', 'extensions'
as $$
declare
  v_updated int;
begin
  if not public.is_admin_user(p_caller_id) then
    raise exception 'Akses ditolak: hanya admin yang dapat reset password'
      using errcode = '42501';
  end if;

  if p_new_password is null or length(p_new_password) < 6 then
    return false;
  end if;

  update public.users
     set password_hash = extensions.crypt(p_new_password, extensions.gen_salt('bf', 10)),
         must_change_password = true
   where id = p_user_id;

  get diagnostics v_updated = row_count;
  return v_updated > 0;
end;
$$;

create or replace function public.pengajuan_pak_approve(
  p_caller_id bigint,
  p_pengajuan_id bigint
)
returns public.pengajuan_pak
language plpgsql
security definer
set search_path to 'public'
as $$
declare
  v_pengajuan public.pengajuan_pak;
begin
  if not public.is_admin_user(p_caller_id) then
    raise exception 'Akses ditolak: hanya admin yang boleh approve pengajuan PAK'
      using errcode = '42501';
  end if;

  select * into v_pengajuan
  from public.pengajuan_pak
  where id = p_pengajuan_id
  for update;

  if not found then
    raise exception 'Pengajuan PAK id=% tidak ditemukan', p_pengajuan_id;
  end if;

  if v_pengajuan.status = 'selesai' then
    return v_pengajuan;
  end if;

  update public.pengajuan_pak
     set status = 'selesai',
         approved_by = p_caller_id,
         approved_at = now()
   where id = p_pengajuan_id
   returning * into v_pengajuan;

  insert into public.riwayat_angka_kredit (pegawai_nip, angka_kredit, tmt)
  values (v_pengajuan.pegawai_nip, v_pengajuan.ak_total, v_pengajuan.tgl_pengajuan);

  return v_pengajuan;
end;
$$;

-- Remove old vulnerable signatures after the frontend has been deployed.
drop function if exists public.buat_pengguna(text, text, text, boolean, text, text[]);
drop function if exists public.reset_password_admin(bigint, text);

-- Keep EXECUTE narrow. The current frontend uses anon key, so anon is still
-- granted only for functions that validate credentials/caller internally.
revoke execute on function public.buat_pengguna(bigint, text, text, text, boolean, text, text[]) from public, authenticated;
revoke execute on function public.reset_password_admin(bigint, bigint, text) from public, authenticated;
revoke execute on function public.pengajuan_pak_approve(bigint, bigint) from public, authenticated;

grant execute on function public.buat_pengguna(bigint, text, text, text, boolean, text, text[]) to anon;
grant execute on function public.reset_password_admin(bigint, bigint, text) to anon;
grant execute on function public.pengajuan_pak_approve(bigint, bigint) to anon;

rollback;

-- Replace ROLLBACK with COMMIT only after smoke testing in a transaction or staging project.
