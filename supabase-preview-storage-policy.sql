-- Policy Supabase Storage untuk file preview sementara.
-- Jalankan seluruh file ini di Supabase SQL Editor.
--
-- Dipakai oleh:
-- 1. Preview Surat Tugas: {surat_id}_{timestamp}.docx
-- 2. Preview PAK: pak_{pengajuan_pak_id}_{timestamp}.docx
--
-- Tujuan utama file ini adalah membuka DELETE yang sebelumnya ditolak
-- "Access denied", supaya cleanup otomatis 10 menit bisa berjalan.
-- Policy tetap dibatasi hanya bucket surat-tugas-preview dan nama file
-- preview .docx berpola timestamp.

-- Pastikan bucket preview ada. Jika bucket sudah ada, pengaturan utama
-- dipertahankan sebagai private bucket dengan akses via signed URL.
insert into storage.buckets (id, name, public, file_size_limit, allowed_mime_types)
values (
  'surat-tugas-preview',
  'surat-tugas-preview',
  false,
  10485760,
  array['application/vnd.openxmlformats-officedocument.wordprocessingml.document']
)
on conflict (id) do update
set public = false,
    file_size_limit = greatest(
      coalesce(storage.buckets.file_size_limit, 0),
      excluded.file_size_limit
    ),
    allowed_mime_types = excluded.allowed_mime_types;

-- Helper pola nama:
-- - 27_1779416745698.docx
-- - pak_3_1780190945865.docx
-- - pak_tmp_1780190945865.docx
--
-- Catatan: policy ditulis langsung sebagai regex agar mudah dibaca
-- di Supabase Dashboard tanpa membuat function tambahan.

drop policy if exists "preview_docx_anon_select" on storage.objects;
create policy "preview_docx_anon_select"
on storage.objects for select
to anon, authenticated
using (
  bucket_id = 'surat-tugas-preview'
  and (
    name ~ '^[0-9]+_[0-9]{10,}\.docx$'
    or name ~ '^pak_([0-9]+|tmp)_[0-9]{10,}\.docx$'
  )
);

drop policy if exists "preview_docx_anon_insert" on storage.objects;
create policy "preview_docx_anon_insert"
on storage.objects for insert
to anon, authenticated
with check (
  bucket_id = 'surat-tugas-preview'
  and (
    name ~ '^[0-9]+_[0-9]{10,}\.docx$'
    or name ~ '^pak_([0-9]+|tmp)_[0-9]{10,}\.docx$'
  )
);

-- Upload preview memakai header x-upsert:true, jadi UPDATE juga dibutuhkan
-- jika nama file yang sama pernah ada.
drop policy if exists "preview_docx_anon_update" on storage.objects;
create policy "preview_docx_anon_update"
on storage.objects for update
to anon, authenticated
using (
  bucket_id = 'surat-tugas-preview'
  and (
    name ~ '^[0-9]+_[0-9]{10,}\.docx$'
    or name ~ '^pak_([0-9]+|tmp)_[0-9]{10,}\.docx$'
  )
)
with check (
  bucket_id = 'surat-tugas-preview'
  and (
    name ~ '^[0-9]+_[0-9]{10,}\.docx$'
    or name ~ '^pak_([0-9]+|tmp)_[0-9]{10,}\.docx$'
  )
);

-- Ini bagian yang memperbaiki error DELETE "Access denied".
drop policy if exists "preview_docx_anon_delete" on storage.objects;
create policy "preview_docx_anon_delete"
on storage.objects for delete
to anon, authenticated
using (
  bucket_id = 'surat-tugas-preview'
  and (
    name ~ '^[0-9]+_[0-9]{10,}\.docx$'
    or name ~ '^pak_([0-9]+|tmp)_[0-9]{10,}\.docx$'
  )
);
