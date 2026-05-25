-- RPC untuk edit dan membatalkan pengajuan PAK.
-- Jalankan seluruh file ini di Supabase SQL Editor.
--
-- Catatan:
-- 1. Jangan tambahkan perintah ALTER TABLE apa pun di tengah function.
-- 2. File ini sengaja tidak memakai variabel composite seperti
--    "v_row public.pengajuan_pak" agar tidak mengundang helper Supabase
--    menambahkan RLS ke nama variabel secara keliru.

drop function if exists public.pengajuan_pak_update(bigint, bigint, integer, integer, integer, text, numeric, numeric, numeric, jsonb, date, text, text);
drop function if exists public.pengajuan_pak_delete(bigint, bigint);

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
returns jsonb
as $function$
declare
  v_pegawai_nip text;
  v_status text;
  v_caller_nip text;
  v_is_admin boolean;
  v_result jsonb;
begin
  select p.pegawai_nip, coalesce(p.status, 'menunggu')
    into v_pegawai_nip, v_status
  from public.pengajuan_pak p
  where p.id = p_pengajuan_id;

  if v_pegawai_nip is null then
    raise exception 'Pengajuan PAK tidak ditemukan';
  end if;

  if v_status <> 'menunggu' then
    raise exception 'Pengajuan yang sudah selesai tidak bisa diedit';
  end if;

  select exists (
    select 1
    from public.users u
    where u.id = p_caller_id
      and (
        u.role = 'admin'
        or coalesce(u.roles, array[]::text[]) @> array['admin']::text[]
      )
  ) into v_is_admin;

  select dp.pegawai_nip
    into v_caller_nip
  from public.data_pegawai dp
  where dp.id = p_caller_id
  limit 1;

  if not coalesce(v_is_admin, false)
     and (v_caller_nip is null or v_caller_nip <> v_pegawai_nip) then
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

  update public.pengajuan_pak p
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
  where p.id = p_pengajuan_id
  returning to_jsonb(p.*) into v_result;

  return v_result;
end;
$function$
language plpgsql
security definer
set search_path = public;

create or replace function public.pengajuan_pak_delete(
  p_caller_id bigint,
  p_pengajuan_id bigint
)
returns jsonb
as $function$
declare
  v_pegawai_nip text;
  v_status text;
  v_caller_nip text;
  v_is_admin boolean;
  v_result jsonb;
begin
  select p.pegawai_nip, coalesce(p.status, 'menunggu')
    into v_pegawai_nip, v_status
  from public.pengajuan_pak p
  where p.id = p_pengajuan_id;

  if v_pegawai_nip is null then
    raise exception 'Pengajuan PAK tidak ditemukan';
  end if;

  if v_status <> 'menunggu' then
    raise exception 'Pengajuan yang sudah selesai tidak bisa dibatalkan';
  end if;

  select exists (
    select 1
    from public.users u
    where u.id = p_caller_id
      and (
        u.role = 'admin'
        or coalesce(u.roles, array[]::text[]) @> array['admin']::text[]
      )
  ) into v_is_admin;

  select dp.pegawai_nip
    into v_caller_nip
  from public.data_pegawai dp
  where dp.id = p_caller_id
  limit 1;

  if not coalesce(v_is_admin, false)
     and (v_caller_nip is null or v_caller_nip <> v_pegawai_nip) then
    raise exception 'Anda tidak berhak membatalkan pengajuan ini';
  end if;

  delete from public.pengajuan_pak p
  where p.id = p_pengajuan_id
  returning to_jsonb(p.*) into v_result;

  return v_result;
end;
$function$
language plpgsql
security definer
set search_path = public;

revoke all on function public.pengajuan_pak_update(bigint, bigint, integer, integer, integer, text, numeric, numeric, numeric, jsonb, date, text, text) from public;
revoke all on function public.pengajuan_pak_delete(bigint, bigint) from public;

grant execute on function public.pengajuan_pak_update(bigint, bigint, integer, integer, integer, text, numeric, numeric, numeric, jsonb, date, text, text) to anon, authenticated;
grant execute on function public.pengajuan_pak_delete(bigint, bigint) to anon, authenticated;
