-- Portal 9201 Supabase hardening draft
-- Review before running. Run in a transaction first, then test login and core flows.
--
-- Important:
-- 1. This project currently uses a custom localStorage session, not Supabase Auth.
-- 2. Without verified server-side caller identity, table policies cannot safely
--    distinguish one browser user from another.
-- 3. The safest short-term approach is:
--    - revoke direct browser write access to sensitive tables
--    - keep only narrowly needed reads
--    - force privileged changes through RPC functions
--    - update RPC functions so they verify caller/admin server-side

begin;

-- Revoke broad direct table privileges from browser-facing roles.
revoke all on table public.users from anon, authenticated;
revoke all on table public.surat_tugas from anon, authenticated;
revoke all on table public.data_pegawai from anon, authenticated;
revoke all on table public.mitra from anon, authenticated;
revoke all on table public.kamus_pok from anon, authenticated;
revoke all on table public.pengajuan_pak from anon, authenticated;
revoke all on table public.riwayat_angka_kredit from anon, authenticated;
revoke all on table public.riwayat_gelar from anon, authenticated;
revoke all on table public.riwayat_jabatan from anon, authenticated;
revoke all on table public.riwayat_pangkat_golongan from anon, authenticated;

-- Drop policies that currently allow public/anon blanket access.
drop policy if exists "Enable read access for all users" on public.data_pegawai;
drop policy if exists kamus_pok_anon_read on public.kamus_pok;
drop policy if exists kamus_pok_anon_write on public.kamus_pok;
drop policy if exists mitra_select_all on public.mitra;
drop policy if exists mitra_write_all on public.mitra;
drop policy if exists pak_select_all on public.pengajuan_pak;
drop policy if exists anon_delete_ak on public.riwayat_angka_kredit;
drop policy if exists anon_insert_ak on public.riwayat_angka_kredit;
drop policy if exists anon_select_ak on public.riwayat_angka_kredit;
drop policy if exists anon_update_ak on public.riwayat_angka_kredit;
drop policy if exists anon_delete_gelar on public.riwayat_gelar;
drop policy if exists anon_insert_gelar on public.riwayat_gelar;
drop policy if exists anon_select_gelar on public.riwayat_gelar;
drop policy if exists anon_update_gelar on public.riwayat_gelar;
drop policy if exists riwayat_pegawai_select_all on public.riwayat_jabatan;
drop policy if exists riwayat_pegawai_write_all on public.riwayat_jabatan;
drop policy if exists anon_delete_pgg on public.riwayat_pangkat_golongan;
drop policy if exists anon_insert_pgg on public.riwayat_pangkat_golongan;
drop policy if exists anon_select_pgg on public.riwayat_pangkat_golongan;
drop policy if exists anon_update_pgg on public.riwayat_pangkat_golongan;
drop policy if exists surat_tugas_select_all on public.surat_tugas;
drop policy if exists surat_tugas_write_all on public.surat_tugas;
drop policy if exists "allow anon update users" on public.users;
drop policy if exists "allow login check" on public.users;
drop policy if exists "allow update password" on public.users;

-- Keep minimal public reads only for data that is truly non-sensitive.
-- Disable these if the data contains private/personal information.
grant select on table public.kamus_pok to anon, authenticated;
create policy kamus_pok_public_read
on public.kamus_pok
for select
to anon, authenticated
using (aktif = true);

-- RPC execution surface.
-- Public EXECUTE grants are redundant and make auditing harder.
revoke execute on function public.buat_pengguna(text, text, text, boolean, text, text[]) from public, anon, authenticated;
revoke execute on function public.change_password(text, text, text) from public, anon, authenticated;
revoke execute on function public.hapus_pengguna(bigint, bigint) from public, anon, authenticated;
revoke execute on function public.mark_notifikasi_read(bigint) from public, anon, authenticated;
revoke execute on function public.pengajuan_pak_approve(bigint, bigint) from public, anon, authenticated;
revoke execute on function public.pengajuan_pak_create(bigint, text, integer, integer, integer, text, numeric, numeric, numeric, jsonb, date, text, text) from public, anon, authenticated;
revoke execute on function public.reset_password_admin(bigint, text) from public, anon, authenticated;
revoke execute on function public.verify_login(text, text) from public, anon, authenticated;

-- Grant back only functions the current frontend must call.
-- These functions must validate credentials/caller/admin internally.
grant execute on function public.verify_login(text, text) to anon;
grant execute on function public.change_password(text, text, text) to anon;
grant execute on function public.pengajuan_pak_create(bigint, text, integer, integer, integer, text, numeric, numeric, numeric, jsonb, date, text, text) to anon;
grant execute on function public.mark_notifikasi_read(bigint) to anon;

-- Admin RPCs should not be executable until their definitions verify caller/admin.
-- Re-enable only after code review of each function body:
-- grant execute on function public.buat_pengguna(text, text, text, boolean, text, text[]) to anon;
-- grant execute on function public.hapus_pengguna(bigint, bigint) to anon;
-- grant execute on function public.pengajuan_pak_approve(bigint, bigint) to anon;
-- grant execute on function public.reset_password_admin(bigint, text) to anon;

-- Inspect before commit:
select
  schemaname,
  tablename,
  policyname,
  roles,
  cmd,
  qual,
  with_check
from pg_policies
where schemaname = 'public'
order by tablename, policyname;

rollback;

-- If the inspection output and app smoke test plan look good, replace
-- the final ROLLBACK above with COMMIT and run again.
