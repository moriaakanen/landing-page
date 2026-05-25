-- RPC untuk edit dan membatalkan pengajuan PAK.
-- Jalankan file ini di Supabase SQL Editor.
--
-- Kenapa perlu RPC?
-- Aplikasi memakai sistem login custom (session di localStorage), bukan Supabase Auth.
-- Karena itu request REST UPDATE/DELETE langsung ke tabel pengajuan_pak tidak punya
-- konteks auth.uid(), sehingga RLS bisa mengembalikan sukses kosong: tidak ada row
-- yang berubah. RPC SECURITY DEFINER ini mengikuti pola pengajuan_pak_create.

create or replace function public.pengajuan_pak_caller_is_admin(p_caller_id bigint)
returns boolean
language sql
security definer
set search_path = public
as $$
  select exists (
    select 1
    from public.users u
    where u.id = p_caller_id
      and (
        u.role = 'admin'
        or coalesce(u.roles, array[]::text[]) @> array['admin']::text[]
      )
  );
$$;

create or replace function public.pengajuan_pak_caller_nip(p_caller_id bigint)
returns text
language sql
security definer
set search_path = public
as $$
  select dp."NIP"
  from public.data_pegawai dp
  where dp.id = p_caller_id
  limit 1;
$$;

create or replace function public.pengajuan_pak_can_mutate(
  p_caller_id bigint,
  p_pengajuan public.pengajuan_pak
)
returns boolean
language plpgsql
security definer
set search_path = public
as $$
declare
  v_caller_nip text;
begin
  if p_caller_id is null then
    return false;
  end if;

  if public.pengajuan_pak_caller_is_admin(p_caller_id) then
    return true;
  end if;

  v_caller_nip := public.pengajuan_pak_caller_nip(p_caller_id);
  return v_caller_nip is not null
     and v_caller_nip = p_pengajuan.pegawai_nip;
end;
$$;

create or replace function public.pengajuan_pak_update(
  p_caller_id bigint,
  p_pengajuan_id bigint,
  p_tahun_periode integer,
  p_bulan_start integer,
  p_bulan_end integer,
  p_mode_hitung text,
  p_ak_didapat numeric,
  p_ak_n1 numeric,
  p_ak_total numeric,
  p_detail_predikat jsonb,
  p_tgl_pengajuan date,
  p_penandatangan_nip text,
  p_penandatangan_nama text default null
)
returns setof public.pengajuan_pak
language plpgsql
security definer
set search_path = public
as $$
declare
  v_row public.pengajuan_pak;
  v_updated public.pengajuan_pak;
begin
  select *
  into v_row
  from public.pengajuan_pak
  where id = p_pengajuan_id;

  if not found then
    raise exception 'Pengajuan PAK tidak ditemukan';
  end if;

  if coalesce(v_row.status, 'menunggu') <> 'menunggu' then
    raise exception 'Pengajuan yang sudah selesai tidak bisa diedit';
  end if;

  if not public.pengajuan_pak_can_mutate(p_caller_id, v_row) then
    raise exception 'Anda tidak berhak mengubah pengajuan ini';
  end if;

  if p_bulan_start is null or p_bulan_end is null
     or p_bulan_start < 1 or p_bulan_start > 12
     or p_bulan_end < 1 or p_bulan_end > 12
     or p_bulan_start > p_bulan_end then
    raise exception 'Periode pengajuan tidak valid';
  end if;

  if p_tahun_periode is null then
    raise exception 'tahun_periode wajib';
  end if;

  if p_tgl_pengajuan is null then
    raise exception 'Tanggal pengajuan wajib';
  end if;

  if p_tgl_pengajuan < make_date(
    case when p_bulan_end = 12 then p_tahun_periode + 1 else p_tahun_periode end,
    case when p_bulan_end = 12 then 1 else p_bulan_end + 1 end,
    1
  ) then
    raise exception 'Tanggal pengajuan minimal harus setelah akhir periode';
  end if;

  if p_penandatangan_nip is null or length(trim(p_penandatangan_nip)) = 0 then
    raise exception 'Penandatangan wajib';
  end if;

  update public.pengajuan_pak
  set tahun_periode = p_tahun_periode,
      bulan_start = p_bulan_start,
      bulan_end = p_bulan_end,
      mode_hitung = p_mode_hitung,
      ak_didapat = p_ak_didapat,
      ak_n1 = p_ak_n1,
      ak_total = p_ak_total,
      detail_predikat = p_detail_predikat,
      tgl_pengajuan = p_tgl_pengajuan,
      penandatangan_nip = p_penandatangan_nip,
      penandatangan_nama = p_penandatangan_nama
  where id = p_pengajuan_id
  returning * into v_updated;

  return next v_updated;
end;
$$;

create or replace function public.pengajuan_pak_delete(
  p_caller_id bigint,
  p_pengajuan_id bigint
)
returns setof public.pengajuan_pak
language plpgsql
security definer
set search_path = public
as $$
declare
  v_row public.pengajuan_pak;
begin
  select *
  into v_row
  from public.pengajuan_pak
  where id = p_pengajuan_id;

  if not found then
    raise exception 'Pengajuan PAK tidak ditemukan';
  end if;

  if coalesce(v_row.status, 'menunggu') <> 'menunggu' then
    raise exception 'Pengajuan yang sudah selesai tidak bisa dibatalkan';
  end if;

  if not public.pengajuan_pak_can_mutate(p_caller_id, v_row) then
    raise exception 'Anda tidak berhak membatalkan pengajuan ini';
  end if;

  delete from public.pengajuan_pak
  where id = p_pengajuan_id
  returning * into v_row;

  return next v_row;
end;
$$;

revoke all on function public.pengajuan_pak_caller_is_admin(bigint) from public;
revoke all on function public.pengajuan_pak_caller_nip(bigint) from public;
revoke all on function public.pengajuan_pak_can_mutate(bigint, public.pengajuan_pak) from public;
revoke all on function public.pengajuan_pak_update(bigint, bigint, integer, integer, integer, text, numeric, numeric, numeric, jsonb, date, text, text) from public;
revoke all on function public.pengajuan_pak_delete(bigint, bigint) from public;

grant execute on function public.pengajuan_pak_update(bigint, bigint, integer, integer, integer, text, numeric, numeric, numeric, jsonb, date, text, text) to anon, authenticated;
grant execute on function public.pengajuan_pak_delete(bigint, bigint) to anon, authenticated;
