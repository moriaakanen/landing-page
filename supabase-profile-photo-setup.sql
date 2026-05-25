-- Setup foto profil pegawai untuk Portal 9201.
-- Jalankan sekali di Supabase SQL Editor.

-- 1) Kolom URL foto di data pegawai.
alter table public.data_pegawai
  add column if not exists foto_url text;

comment on column public.data_pegawai.foto_url
  is 'Public URL foto profil pegawai dari Supabase Storage bucket foto-profil.';

-- 2) Bucket public khusus avatar.
insert into storage.buckets (id, name, public, file_size_limit, allowed_mime_types)
values (
  'foto-profil',
  'foto-profil',
  true,
  1048576,
  array['image/jpeg', 'image/png', 'image/webp']
)
on conflict (id) do update
set public = excluded.public,
    file_size_limit = excluded.file_size_limit,
    allowed_mime_types = excluded.allowed_mime_types;

-- 3) Policy storage untuk aplikasi statis yang masih memakai anon key.
-- Catatan: karena aplikasi belum memakai Supabase Auth, policy ini bersifat
-- longgar di level storage. Frontend tetap mengompres file menjadi JPEG 512px.
drop policy if exists "foto_profil_public_read" on storage.objects;
create policy "foto_profil_public_read"
on storage.objects for select
to anon
using (bucket_id = 'foto-profil');

drop policy if exists "foto_profil_anon_upload" on storage.objects;
create policy "foto_profil_anon_upload"
on storage.objects for insert
to anon
with check (bucket_id = 'foto-profil');

drop policy if exists "foto_profil_anon_update" on storage.objects;
create policy "foto_profil_anon_update"
on storage.objects for update
to anon
using (bucket_id = 'foto-profil')
with check (bucket_id = 'foto-profil');

-- 4) RPC sempit untuk update hanya kolom foto_url, supaya tidak perlu
-- membuka UPDATE penuh pada tabel data_pegawai.
create or replace function public.update_pegawai_foto(
  p_nip text,
  p_foto_url text
)
returns public.data_pegawai
language plpgsql
security definer
set search_path = public
as $$
declare
  v_row public.data_pegawai;
begin
  update public.data_pegawai
     set foto_url = p_foto_url
   where "NIP" = p_nip
   returning * into v_row;

  if not found then
    raise exception 'Pegawai dengan NIP % tidak ditemukan', p_nip
      using errcode = 'P0002';
  end if;

  return v_row;
end;
$$;

grant execute on function public.update_pegawai_foto(text, text) to anon;
