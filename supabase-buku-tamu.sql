-- Buku Tamu PST BPS Kabupaten Raja Ampat
-- Jalankan file ini di Supabase SQL Editor sebelum memakai form Buku Tamu
-- pada login.html.

create table if not exists public.buku_tamu (
  id bigserial primary key,
  nama text not null check (char_length(trim(nama)) between 2 and 120),
  no_hp text not null check (char_length(trim(no_hp)) between 6 and 30),
  asal_instansi text not null check (char_length(trim(asal_instansi)) between 2 and 160),
  kategori_pengunjung text not null check (
    kategori_pengunjung in (
      'Umum',
      'Pelajar/Mahasiswa',
      'Instansi Pemerintah',
      'Swasta',
      'Media',
      'Internal BPS',
      'Lainnya'
    )
  ),
  layanan text not null check (
    layanan in (
      'Konsultasi Statistik',
      'Permintaan Data',
      'Publikasi/Perpustakaan',
      'Rekomendasi Statistik',
      'Pengaduan Layanan',
      'Lainnya'
    )
  ),
  keperluan text not null check (char_length(trim(keperluan)) between 3 and 500),
  jumlah_pengunjung integer not null default 1 check (jumlah_pengunjung between 1 and 50),
  tanggal_kunjungan date not null default ((now() at time zone 'Asia/Jayapura')::date),
  jam_kunjungan time without time zone not null default ((now() at time zone 'Asia/Jayapura')::time(0)),
  created_at timestamptz not null default now()
);

comment on table public.buku_tamu is 'Catatan kunjungan PST BPS Kabupaten Raja Ampat dari form publik login page.';
comment on column public.buku_tamu.no_hp is 'Nomor kontak pengunjung untuk tindak lanjut layanan PST.';
comment on column public.buku_tamu.asal_instansi is 'Asal instansi, sekolah, kampus, atau alamat pengunjung.';
comment on column public.buku_tamu.tanggal_kunjungan is 'Tanggal kunjungan lokal Asia/Jayapura.';
comment on column public.buku_tamu.jam_kunjungan is 'Jam kunjungan lokal Asia/Jayapura.';

alter table public.buku_tamu enable row level security;

revoke all on public.buku_tamu from anon, authenticated;
grant insert on public.buku_tamu to anon, authenticated;
grant usage, select on sequence public.buku_tamu_id_seq to anon, authenticated;

drop policy if exists "buku_tamu_public_insert" on public.buku_tamu;
create policy "buku_tamu_public_insert"
on public.buku_tamu
for insert
to anon, authenticated
with check (true);

-- Data buku tamu berisi informasi pribadi pengunjung.
-- Sengaja tidak ada grant SELECT/UPDATE/DELETE untuk anon/authenticated.
-- Jika nanti dibutuhkan halaman rekap admin, sebaiknya buat RPC khusus yang
-- memvalidasi session admin server-side sebelum membuka data.
