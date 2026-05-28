-- RPC import predikat kinerja agar proses simpan tidak membuka RLS tabel
-- predikat_kinerja_bulanan/predikat_kinerja_tahunan secara lebar.
-- Jalankan seluruh file ini di Supabase SQL Editor.

drop function if exists public.predikat_kinerja_import(bigint, text, integer, jsonb);

create or replace function public.predikat_kinerja_import(
  p_caller_id bigint,
  p_mode text,
  p_year integer,
  p_payload jsonb
)
returns jsonb
as $function$
declare
  v_is_admin boolean;
  v_mode text := lower(trim(coalesce(p_mode, '')));
  v_inserted integer := 0;
  v_employee_count integer := 0;
  v_month_count integer := 0;
begin
  if p_caller_id is null then
    raise exception 'Session admin tidak valid';
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

  if not coalesce(v_is_admin, false) then
    raise exception 'Anda tidak berhak import predikat kinerja';
  end if;

  if p_year is null or p_year < 2020 or p_year > 2100 then
    raise exception 'Tahun import tidak valid';
  end if;

  if p_payload is null or jsonb_typeof(p_payload) <> 'array' then
    raise exception 'Payload import harus berupa array';
  end if;

  if jsonb_array_length(p_payload) = 0 then
    return jsonb_build_object('mode', v_mode, 'year', p_year, 'inserted', 0);
  end if;

  if v_mode in ('bulanan', 'periode', 'triwulanan') then
    drop table if exists pg_temp.predikat_import_bulanan;
    create temporary table predikat_import_bulanan on commit drop as
    with raw as (
      select *
      from jsonb_to_recordset(p_payload) as x(
        pegawai_nip text,
        nama text,
        tahun integer,
        bulan integer,
        bulan_nama text,
        hasil_kerja numeric,
        perilaku numeric,
        nilai numeric,
        predikat text,
        status text
      )
    ),
    valid as (
      select
        trim(pegawai_nip) as pegawai_nip,
        nullif(trim(coalesce(nama, '')), '') as nama,
        p_year as tahun,
        bulan,
        nullif(trim(coalesce(bulan_nama, '')), '') as bulan_nama,
        hasil_kerja,
        perilaku,
        nilai,
        nullif(trim(coalesce(predikat, '')), '') as predikat,
        nullif(trim(coalesce(status, '')), '') as status
      from raw
      where nullif(trim(coalesce(pegawai_nip, '')), '') is not null
        and tahun = p_year
        and bulan between 1 and 12
    )
    select distinct on (pegawai_nip, bulan)
      pegawai_nip, nama, tahun, bulan, bulan_nama, hasil_kerja, perilaku, nilai, predikat, status
    from valid
    order by pegawai_nip, bulan;

    select count(distinct pegawai_nip), count(distinct bulan)
      into v_employee_count, v_month_count
    from pg_temp.predikat_import_bulanan;

    delete from public.predikat_kinerja_bulanan p
    where p.tahun = p_year
      and exists (
        select 1
        from pg_temp.predikat_import_bulanan v
        where v.pegawai_nip = p.pegawai_nip
          and v.bulan = p.bulan
      );

    insert into public.predikat_kinerja_bulanan
      (pegawai_nip, nama, tahun, bulan, bulan_nama, hasil_kerja, perilaku, nilai, predikat, status)
    select pegawai_nip, nama, tahun, bulan, bulan_nama, hasil_kerja, perilaku, nilai, predikat, status
    from pg_temp.predikat_import_bulanan;

    get diagnostics v_inserted = row_count;
    drop table if exists pg_temp.predikat_import_bulanan;

    return jsonb_build_object(
      'mode', 'bulanan',
      'year', p_year,
      'inserted', coalesce(v_inserted, 0),
      'employees', coalesce(v_employee_count, 0),
      'months', coalesce(v_month_count, 0)
    );
  end if;

  if v_mode = 'tahunan' then
    drop table if exists pg_temp.predikat_import_tahunan;
    create temporary table predikat_import_tahunan on commit drop as
    with raw as (
      select *
      from jsonb_to_recordset(p_payload) as x(
        pegawai_nip text,
        nama text,
        tahun integer,
        hasil_kerja numeric,
        perilaku numeric,
        nilai numeric,
        predikat text,
        status text
      )
    ),
    valid as (
      select
        trim(pegawai_nip) as pegawai_nip,
        nullif(trim(coalesce(nama, '')), '') as nama,
        p_year as tahun,
        hasil_kerja,
        perilaku,
        nilai,
        nullif(trim(coalesce(predikat, '')), '') as predikat,
        nullif(trim(coalesce(status, '')), '') as status
      from raw
      where nullif(trim(coalesce(pegawai_nip, '')), '') is not null
        and tahun = p_year
    )
    select distinct on (pegawai_nip)
      pegawai_nip, nama, tahun, hasil_kerja, perilaku, nilai, predikat, status
    from valid
    order by pegawai_nip;

    select count(distinct pegawai_nip)
      into v_employee_count
    from pg_temp.predikat_import_tahunan;

    delete from public.predikat_kinerja_tahunan p
    where p.tahun = p_year
      and exists (
        select 1
        from pg_temp.predikat_import_tahunan v
        where v.pegawai_nip = p.pegawai_nip
      );

    insert into public.predikat_kinerja_tahunan
      (pegawai_nip, nama, tahun, hasil_kerja, perilaku, nilai, predikat, status)
    select pegawai_nip, nama, tahun, hasil_kerja, perilaku, nilai, predikat, status
    from pg_temp.predikat_import_tahunan;

    get diagnostics v_inserted = row_count;
    drop table if exists pg_temp.predikat_import_tahunan;

    return jsonb_build_object(
      'mode', 'tahunan',
      'year', p_year,
      'inserted', coalesce(v_inserted, 0),
      'employees', coalesce(v_employee_count, 0)
    );
  end if;

  raise exception 'Mode import tidak dikenal: %', p_mode;
end;
$function$
language plpgsql
security definer
set search_path = public;

revoke all on function public.predikat_kinerja_import(bigint, text, integer, jsonb) from public;
grant execute on function public.predikat_kinerja_import(bigint, text, integer, jsonb) to anon, authenticated;
