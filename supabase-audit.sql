-- Portal 9201 Supabase audit queries
-- Run this in Supabase SQL Editor. These queries are read-only.
-- Share the result sets for a deeper review; do not share service-role keys.

-- 1. Tables used by the frontend and their RLS status.
select
  n.nspname as schema_name,
  c.relname as table_name,
  c.relrowsecurity as rls_enabled,
  c.relforcerowsecurity as rls_forced
from pg_class c
join pg_namespace n on n.oid = c.relnamespace
where n.nspname = 'public'
  and c.relkind = 'r'
  and c.relname in (
    'users',
    'surat_tugas',
    'data_pegawai',
    'mitra',
    'kamus_pok',
    'pengajuan_pak',
    'riwayat_angka_kredit',
    'riwayat_pangkat_golongan',
    'riwayat_jabatan',
    'riwayat_gelar',
    'notifikasi'
  )
order by c.relname;

-- 2. Policies on those tables.
select
  schemaname,
  tablename,
  policyname,
  permissive,
  roles,
  cmd,
  qual,
  with_check
from pg_policies
where schemaname = 'public'
  and tablename in (
    'users',
    'surat_tugas',
    'data_pegawai',
    'mitra',
    'kamus_pok',
    'pengajuan_pak',
    'riwayat_angka_kredit',
    'riwayat_pangkat_golongan',
    'riwayat_jabatan',
    'riwayat_gelar',
    'notifikasi'
  )
order by tablename, policyname;

-- 3. Direct table privileges granted to browser-facing roles.
select
  table_schema,
  table_name,
  grantee,
  privilege_type
from information_schema.table_privileges
where table_schema = 'public'
  and grantee in ('anon', 'authenticated')
  and table_name in (
    'users',
    'surat_tugas',
    'data_pegawai',
    'mitra',
    'kamus_pok',
    'pengajuan_pak',
    'riwayat_angka_kredit',
    'riwayat_pangkat_golongan',
    'riwayat_jabatan',
    'riwayat_gelar',
    'notifikasi'
  )
order by table_name, grantee, privilege_type;

-- 4. RPC functions used by the frontend.
select
  n.nspname as schema_name,
  p.proname as function_name,
  pg_get_function_arguments(p.oid) as arguments,
  case p.prosecdef when true then 'SECURITY DEFINER' else 'SECURITY INVOKER' end as security_mode,
  array_to_string(p.proacl, E'\n') as grants
from pg_proc p
join pg_namespace n on n.oid = p.pronamespace
where n.nspname = 'public'
  and p.proname in (
    'verify_login',
    'change_password',
    'buat_pengguna',
    'reset_password_admin',
    'hapus_pengguna',
    'pengajuan_pak_create',
    'pengajuan_pak_approve',
    'mark_notifikasi_read'
  )
order by p.proname;

-- 5. Function definitions for manual review.
select
  p.proname as function_name,
  pg_get_functiondef(p.oid) as function_definition
from pg_proc p
join pg_namespace n on n.oid = p.pronamespace
where n.nspname = 'public'
  and p.proname in (
    'verify_login',
    'change_password',
    'buat_pengguna',
    'reset_password_admin',
    'hapus_pengguna',
    'pengajuan_pak_create',
    'pengajuan_pak_approve',
    'mark_notifikasi_read'
  )
order by p.proname;

-- 6. Storage policies and bucket visibility.
select
  id,
  name,
  public,
  file_size_limit,
  allowed_mime_types
from storage.buckets
where id in ('template', 'preview', 'pak-preview', 'surat-preview');

select
  schemaname,
  tablename,
  policyname,
  roles,
  cmd,
  qual,
  with_check
from pg_policies
where schemaname = 'storage'
order by tablename, policyname;
