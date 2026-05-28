-- Tambahkan pemetaan NIP lama untuk import predikat kinerja dari aplikasi Kinerja.
-- File export memakai kolom "Niplama", sedangkan tabel predikat memakai pegawai_nip.
alter table public.data_pegawai
  add column if not exists nip_lama text;

comment on column public.data_pegawai.nip_lama is
  'NIP lama/Niplama dari aplikasi Kinerja. Dipakai untuk mapping import predikat kinerja ke pegawai_nip.';

create unique index if not exists data_pegawai_nip_lama_unique
  on public.data_pegawai (nip_lama)
  where nip_lama is not null and btrim(nip_lama) <> '';
