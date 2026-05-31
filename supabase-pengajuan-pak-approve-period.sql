-- Perbaikan RPC approve PAK.
-- Jalankan file ini di Supabase SQL Editor.
--
-- Efek:
-- 1. Saat pengajuan PAK di-approve, riwayat_angka_kredit otomatis berisi:
--    - periode_tahun
--    - periode_bulan_awal
--    - periode_bulan_akhir
--    Catatan: periode_label tidak di-insert manual karena di beberapa database
--    kolom ini berupa generated/default column.
-- 2. Kolom tmt diisi tanggal terakhir bulan akhir periode.
--    Contoh: Januari s.d Desember 2025 -> 2025-12-31.

drop function if exists public.pengajuan_pak_approve(bigint, bigint);

create or replace function public.pengajuan_pak_approve(
  p_caller_id bigint,
  p_pengajuan_id bigint
)
returns jsonb
language plpgsql
security definer
set search_path = public
as $function$
declare
  v_is_admin boolean;
  v_pegawai_nip text;
  v_status text;
  v_tahun integer;
  v_bulan_start integer;
  v_bulan_end integer;
  v_ak_total numeric;
  v_tmt date;
  v_result jsonb;
begin
  select exists (
    select 1
    from public.users u
    where u.id = p_caller_id
      and (
        u.role = 'admin'
        or coalesce(u.roles, array[]::text[]) @> array['admin']::text[]
      )
  ) into v_is_admin;

  if not coalesce(v_is_admin, false) then
    raise exception 'admin access required';
  end if;

  select p.pegawai_nip,
         coalesce(p.status, 'menunggu'),
         p.tahun_periode,
         p.bulan_start,
         p.bulan_end,
         p.ak_total
    into v_pegawai_nip,
         v_status,
         v_tahun,
         v_bulan_start,
         v_bulan_end,
         v_ak_total
  from public.pengajuan_pak p
  where p.id = p_pengajuan_id;

  if v_pegawai_nip is null then
    raise exception 'Pengajuan PAK tidak ditemukan';
  end if;

  if v_status <> 'menunggu' then
    raise exception 'Pengajuan PAK sudah diproses';
  end if;

  if v_tahun is null
     or v_bulan_start is null
     or v_bulan_end is null
     or v_bulan_start < 1 or v_bulan_start > 12
     or v_bulan_end < 1 or v_bulan_end > 12
     or v_bulan_start > v_bulan_end then
    raise exception 'Periode pengajuan tidak valid';
  end if;

  v_tmt := (make_date(v_tahun, v_bulan_end, 1) + interval '1 month - 1 day')::date;

  insert into public.riwayat_angka_kredit (
    pegawai_nip,
    angka_kredit,
    tmt,
    periode_tahun,
    periode_bulan_awal,
    periode_bulan_akhir
  )
  values (
    v_pegawai_nip,
    v_ak_total,
    v_tmt,
    v_tahun,
    v_bulan_start,
    v_bulan_end
  );

  update public.pengajuan_pak p
  set status = 'selesai'
  where p.id = p_pengajuan_id
  returning to_jsonb(p.*) into v_result;

  return v_result;
end;
$function$;

revoke all on function public.pengajuan_pak_approve(bigint, bigint) from public;
grant execute on function public.pengajuan_pak_approve(bigint, bigint) to anon, authenticated;
