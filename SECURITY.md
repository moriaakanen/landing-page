# Security Notes

## Supabase Access Model

This app is a static frontend. Any value stored in `localStorage` can be modified by the browser user, so admin authorization must never rely on frontend checks alone.

The frontend now performs a fresh admin-role recheck with `novaVerifyAdminSession()` before loading admin pages, but Supabase must still enforce the real boundary with RLS policies and RPC checks.

## Required Supabase Hardening

1. Enable RLS on every table used by this app, especially:
   - `users`
   - `surat_tugas`
   - `data_pegawai`
   - `mitra`
   - `kamus_pok`
   - `pengajuan_pak`
   - all `riwayat_*` tables

2. Do not allow broad anonymous `select`, `insert`, `update`, or `delete` policies unless the row is intentionally public.

3. Move privileged actions behind RPC functions that verify admin rights inside the database:
   - create user
   - reset password
   - delete user
   - update roles
   - approve PAK
   - approve or bulk-approve surat tugas

4. Keep `SUPABASE_ANON_KEY` public-only. Never use the service-role key in `config.js` or any browser-delivered file.

5. Keep `ADMIN_USERS` in `config.js` empty after bootstrap. Admin roles should come from the `users.roles` column and database-side authorization.

## Suggested RPC Pattern

Privileged RPC functions should validate the caller server-side before changing data. With a custom auth system, pass only an opaque session/caller value that the database can verify against trusted data; do not trust role values sent from the browser.

Example shape:

```sql
-- Pseudocode: adapt to the real auth/session schema.
create or replace function assert_admin(p_user_id bigint)
returns void
language plpgsql
security definer
as $$
begin
  if not exists (
    select 1
    from users
    where id = p_user_id
      and roles @> array['admin']::text[]
  ) then
    raise exception 'admin access required';
  end if;
end;
$$;
```

For stronger security, migrate to Supabase Auth and enforce policies with `auth.uid()` instead of a browser-managed localStorage session.

## Auditing The Live Supabase Project

Run `supabase-audit.sql` in the Supabase SQL Editor and review the result sets. The queries are read-only and are designed to expose:

- RLS status for tables used by this frontend
- policies on those tables
- direct grants to `anon` and `authenticated`
- security mode and definitions of RPC functions used by the app
- storage bucket visibility and policies

Do not paste service-role keys into chat or commit them to this repository.
