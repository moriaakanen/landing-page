-- Status kepegawaian dan masa berlaku jabatan.
-- Jalankan di Supabase SQL Editor sebelum memakai fitur pensiun/TMT selesai.

alter table public.data_pegawai
  add column if not exists status_kepegawaian text not null default 'aktif',
  add column if not exists tanggal_pensiun date,
  add column if not exists keterangan_status text;

alter table public.riwayat_jabatan
  add column if not exists tmt_selesai date;

do $$
begin
  if not exists (
    select 1
    from pg_constraint
    where conname = 'data_pegawai_status_kepegawaian_check'
      and conrelid = 'public.data_pegawai'::regclass
  ) then
    alter table public.data_pegawai
      add constraint data_pegawai_status_kepegawaian_check
      check (status_kepegawaian in ('aktif', 'pensiun', 'mutasi', 'meninggal'));
  end if;

  if not exists (
    select 1
    from pg_constraint
    where conname = 'riwayat_jabatan_tmt_selesai_check'
      and conrelid = 'public.riwayat_jabatan'::regclass
  ) then
    alter table public.riwayat_jabatan
      add constraint riwayat_jabatan_tmt_selesai_check
      check (tmt_selesai is null or tmt_selesai >= tmt);
  end if;
end $$;

create index if not exists data_pegawai_status_kepegawaian_idx
  on public.data_pegawai (status_kepegawaian, tanggal_pensiun);

create index if not exists riwayat_jabatan_pegawai_tmt_selesai_idx
  on public.riwayat_jabatan (pegawai_nip, jenis, tmt desc, tmt_selesai);

comment on column public.data_pegawai.status_kepegawaian is
  'Status aktif/pensiun/mutasi/meninggal. Jangan hapus pegawai historis.';

comment on column public.data_pegawai.tanggal_pensiun is
  'Tanggal mulai pensiun. Pada dan setelah tanggal ini pegawai tidak selectable untuk surat baru.';

comment on column public.riwayat_jabatan.tmt_selesai is
  'Tanggal akhir masa berlaku jabatan, terutama untuk jenis=lainnya seperti PPK, Bendahara, atau Plt.';

-- Contoh pensiun:
-- update public.data_pegawai
-- set status_kepegawaian = 'pensiun',
--     tanggal_pensiun = '2026-10-01',
--     keterangan_status = 'Pensiun TMT 1 Oktober 2026'
-- where pegawai_nip = '199903302019121001';

-- Contoh jabatan lainnya selesai:
-- update public.riwayat_jabatan
-- set tmt_selesai = '2026-12-31'
-- where pegawai_nip = '199510152018021001'
--   and jenis = 'lainnya'
--   and jabatan = 'Pejabat Pembuat Komitmen'
--   and tmt = '2026-01-01';
