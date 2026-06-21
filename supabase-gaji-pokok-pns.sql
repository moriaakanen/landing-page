-- Master gaji pokok PNS untuk dasar Kenaikan Gaji Berkala (KGB).
-- Jalankan file ini di Supabase SQL Editor.
--
-- Sumber data:
--   Peraturan Pemerintah Republik Indonesia Nomor 5 Tahun 2024
--   tentang Perubahan Kesembilan Belas atas PP Nomor 7 Tahun 1977
--   tentang Peraturan Gaji Pegawai Negeri Sipil.
--
-- Catatan regulasi:
--   Lampiran gaji pokok berlaku mulai 1 Januari 2024.
--   PP ditetapkan/diundangkan pada 26 Januari 2024.
--
-- Desain:
--   CSV asal berbentuk wide: mkg, I/a, I/b, ...
--   Tabel Supabase disimpan long: satu baris untuk setiap kombinasi
--   versi peraturan + golongan + mkg.

create table if not exists public.gaji_pokok_pns (
  id bigserial primary key,
  versi text not null,
  jenis_peraturan text not null default 'Peraturan Pemerintah',
  nomor_peraturan integer not null,
  tahun_peraturan integer not null,
  nama_peraturan text not null,
  berlaku_mulai date not null,
  tanggal_penetapan date,
  golongan text not null,
  golongan_induk text not null,
  ruang text not null,
  mkg integer not null,
  gaji_pokok integer not null,
  sumber_file text,
  is_active boolean not null default true,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now(),
  constraint gaji_pokok_pns_mkg_check check (mkg >= 0),
  constraint gaji_pokok_pns_gaji_check check (gaji_pokok > 0),
  constraint gaji_pokok_pns_golongan_check check (golongan ~ '^(I|II|III|IV)/(a|b|c|d|e)$')
);

create unique index if not exists gaji_pokok_pns_versi_golongan_mkg_uidx
  on public.gaji_pokok_pns(versi, golongan, mkg);

create index if not exists gaji_pokok_pns_active_lookup_idx
  on public.gaji_pokok_pns(is_active, golongan, mkg);

create index if not exists gaji_pokok_pns_regulation_idx
  on public.gaji_pokok_pns(tahun_peraturan desc, nomor_peraturan desc, berlaku_mulai desc);

create or replace function public.set_gaji_pokok_pns_updated_at()
returns trigger
language plpgsql
as $$
begin
  new.updated_at = now();
  return new;
end;
$$;

drop trigger if exists trg_gaji_pokok_pns_updated_at on public.gaji_pokok_pns;
create trigger trg_gaji_pokok_pns_updated_at
before update on public.gaji_pokok_pns
for each row execute function public.set_gaji_pokok_pns_updated_at();

with csv_source(raw_csv) as (
  values ($csv$mkg,I/a,I/b,I/c,I/d,II/a,II/b,II/c,II/d,III/a,III/b,III/c,III/d,IV/a,IV/b,IV/c,IV/d,IV/e
0,1.685.700,,,,2.184.000,,,,2.785.700,2.903.600,3.026.400,3.154.400,3.287.800,3.426.900,3.571.900,3.723.000,3.880.400
1,,,,,2.218.400,,,,,,,,,,,,
2,1.738.800,,,,,,,,2.873.500,2.995.000,3.121.700,3.253.700,3.391.400,3.534.800,3.684.400,3.840.200,4.002.700
3,,1.840.800,1.918.700,1.999.900,2.288.200,2.385.000,2.485.900,2.591.100,,,,,,,,,
4,1.793.500,,,,,,,,2.964.000,3.089.300,3.220.000,3.356.200,3.498.200,3.646.200,3.800.400,3.961.200,4.128.700
5,,1.898.800,1.979.100,2.062.900,2.360.300,2.460.100,2.564.200,2.672.700,,,,,,,,,
6,1.850.000,,,,,,,,3.057.300,3.186.600,3.321.400,3.461.900,3.608.400,3.761.000,3.920.100,4.085.900,4.258.700
7,,1.958.600,2.041.500,2.127.800,2.434.600,2.537.600,2.645.000,2.756.800,,,,,,,,,
8,1.908.300,,,,,,,,3.153.600,3.287.000,3.426.000,3.571.000,3.722.000,3.879.500,4.043.600,4.214.600,4.392.900
9,,2.020.300,2.105.800,2.194.800,2.511.300,2.617.500,2.728.300,2.843.700,,,,,,,,,
10,1.968.400,,,,,,,,3.252.900,3.390.500,3.533.900,3.683.400,3.839.200,4.001.600,4.170.900,4.347.300,4.531.200
11,,2.083.900,2.172.100,2.264.000,2.590.400,2.700.000,2.814.200,2.933.200,,,,,,,,,
12,2.030.400,,,,,,,,3.355.400,3.497.300,3.645.200,3.799.400,3.960.200,4.127.700,4.302.300,4.484.300,4.673.900
13,,2.149.600,2.240.500,2.335.300,2.672.000,2.785.000,2.902.800,3.025.600,,,,,,,,,
14,2.094.300,,,,,,,,3.461.100,3.607.500,3.760.100,3.919.100,4.084.900,4.257.700,4.437.800,4.625.500,4.821.100
15,,2.217.300,2.311.100,2.408.800,2.756.200,2.872.700,2.994.300,3.120.900,,,,,,,,,
16,2.160.300,,,,,,,,3.570.100,3.721.100,3.878.500,4.042.500,4.213.500,4.391.800,4.577.500,4.771.200,4.973.000
17,,2.287.100,2.383.900,2.484.700,2.843.000,2.963.200,3.088.600,3.219.200,,,,,,,,,
18,2.228.300,,,,,,,,3.682.500,3.838.300,4.000.600,4.169.900,4.346.200,4.530.100,4.721.700,4.921.400,5.129.600
19,,2.359.100,2.458.900,2.562.900,2.932.500,3.056.500,3.185.800,3.320.600,,,,,,,,,
20,2.298.500,,,,,,,,3.798.500,3.959.200,4.126.600,4.301.200,4.483.100,4.672.800,4.870.400,5.076.400,5.291.200
21,,2.433.400,2.536.400,2.643.700,3.024.900,3.152.800,3.286.200,3.425.200,,,,,,,,,
22,2.370.900,,,,,,,,3.918.100,4.083.900,4.256.600,4.436.700,4.624.300,4.819.900,5.023.800,5.236.300,5.457.800
23,,2.510.100,2.616.300,2.726.900,3.120.100,3.252.100,3.389.700,3.533.100,,,,,,,,,
24,2.445.500,,,,,,,,4.041.500,4.212.500,4.390.700,4.576.400,4.770.000,4.971.700,5.182.000,5.401.200,5.629.700
25,,2.589.100,2.698.700,2.812.800,3.218.400,3.354.500,3.496.400,3.644.300,,,,,,,,,
26,2.522.600,,,,,,,,4.168.800,4.345.100,4.528.900,4.720.500,4.920.200,5.128.300,5.345.200,5.571.400,5.807.000
27,,2.670.700,2.783.700,2.901.400,3.319.800,3.460.200,3.606.500,3.759.100,,,,,,,,,
28,,,,,,,,,4.300.100,4.482.000,4.671.600,4.869.200,5.075.200,5.289.800,5.513.600,5.746.800,5.989.900
29,,,,,3.424.300,3.569.200,3.720.100,3.877.500,,,,,,,,,
30,,,,,,,,,4.435.500,4.623.200,4.818.700,5.022.500,5.235.000,5.456.400,5.687.200,5.927.800,6.178.600
31,,,,,3.532.200,3.681.600,3.837.300,3.999.600,,,,,,,,,
32,,,,,,,,,4.575.200,4.768.800,4.970.500,5.180.700,5.399.900,5.628.300,5.866.400,6.114.500,6.373.200
33,,,,,3.643.400,3.797.500,3.958.200,4.125.600,,,,,,,,,$csv$)
),
csv_lines as (
  select row_number() over () as line_no, line
  from csv_source,
       regexp_split_to_table(trim(both from raw_csv), E'\r?\n') as line
),
csv_header as (
  select regexp_split_to_array(line, ',') as cols
  from csv_lines
  where line_no = 1
),
csv_rows as (
  select line_no, regexp_split_to_array(line, ',') as cols
  from csv_lines
  where line_no > 1
),
normalized as (
  select
    'PP 5 Tahun 2024'::text as versi,
    'Peraturan Pemerintah'::text as jenis_peraturan,
    5::integer as nomor_peraturan,
    2024::integer as tahun_peraturan,
    'Peraturan Pemerintah Republik Indonesia Nomor 5 Tahun 2024'::text as nama_peraturan,
    date '2024-01-01' as berlaku_mulai,
    date '2024-01-26' as tanggal_penetapan,
    h.cols[i]::text as golongan,
    split_part(h.cols[i], '/', 1)::text as golongan_induk,
    split_part(h.cols[i], '/', 2)::text as ruang,
    r.cols[1]::integer as mkg,
    replace(r.cols[i], '.', '')::integer as gaji_pokok,
    'daftar-gaji-pokok-pns.csv'::text as sumber_file,
    true::boolean as is_active
  from csv_rows r
  cross join csv_header h
  cross join lateral generate_subscripts(h.cols, 1) as g(i)
  where i > 1
    and nullif(btrim(coalesce(r.cols[i], '')), '') is not null
)
insert into public.gaji_pokok_pns (
  versi,
  jenis_peraturan,
  nomor_peraturan,
  tahun_peraturan,
  nama_peraturan,
  berlaku_mulai,
  tanggal_penetapan,
  golongan,
  golongan_induk,
  ruang,
  mkg,
  gaji_pokok,
  sumber_file,
  is_active
)
select
  versi,
  jenis_peraturan,
  nomor_peraturan,
  tahun_peraturan,
  nama_peraturan,
  berlaku_mulai,
  tanggal_penetapan,
  golongan,
  golongan_induk,
  ruang,
  mkg,
  gaji_pokok,
  sumber_file,
  is_active
from normalized
on conflict (versi, golongan, mkg) do update set
  jenis_peraturan = excluded.jenis_peraturan,
  nomor_peraturan = excluded.nomor_peraturan,
  tahun_peraturan = excluded.tahun_peraturan,
  nama_peraturan = excluded.nama_peraturan,
  berlaku_mulai = excluded.berlaku_mulai,
  tanggal_penetapan = excluded.tanggal_penetapan,
  golongan_induk = excluded.golongan_induk,
  ruang = excluded.ruang,
  gaji_pokok = excluded.gaji_pokok,
  sumber_file = excluded.sumber_file,
  is_active = excluded.is_active,
  updated_at = now();

alter table public.gaji_pokok_pns enable row level security;

drop policy if exists "gaji_pokok_pns_select_all" on public.gaji_pokok_pns;
create policy "gaji_pokok_pns_select_all"
on public.gaji_pokok_pns
for select
using (true);

grant select on public.gaji_pokok_pns to anon, authenticated;

comment on table public.gaji_pokok_pns is
  'Master gaji pokok PNS format long. Seed awal dari PP RI Nomor 5 Tahun 2024 untuk dasar KGB.';
comment on column public.gaji_pokok_pns.mkg is 'Masa kerja golongan dalam tahun.';
comment on column public.gaji_pokok_pns.gaji_pokok is 'Nilai gaji pokok rupiah tanpa pemisah ribuan.';
comment on column public.gaji_pokok_pns.is_active is
  'Tandai true untuk versi peraturan yang sedang dipakai aplikasi. Jika ada PP baru, versi lama dapat diset false.';

select
  versi,
  count(*) as jumlah_baris,
  min(mkg) as mkg_min,
  max(mkg) as mkg_max,
  min(gaji_pokok) as gaji_min,
  max(gaji_pokok) as gaji_max
from public.gaji_pokok_pns
where versi = 'PP 5 Tahun 2024'
group by versi;
