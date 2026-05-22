-- Portal 9201 storage remediation draft
-- Current audit shows anon can INSERT, SELECT, and DELETE every object in
-- bucket `surat-tugas-preview`. That allows anonymous deletion of all preview
-- docs and arbitrary uploads to the preview bucket.
--
-- This draft keeps the current frontend mostly working by allowing anonymous
-- upload/sign-read of generated .docx previews, but removes anonymous delete.
-- A stronger design is to create previews server-side and avoid anon storage
-- writes entirely.

begin;

drop policy if exists preview_anon_delete on storage.objects;
drop policy if exists preview_anon_insert on storage.objects;
drop policy if exists preview_anon_select on storage.objects;

create policy preview_anon_insert_docx_only
on storage.objects
for insert
to anon
with check (
  bucket_id = 'surat-tugas-preview'
  and lower(name) like '%.docx'
);

create policy preview_anon_select_docx_only
on storage.objects
for select
to anon
using (
  bucket_id = 'surat-tugas-preview'
  and lower(name) like '%.docx'
);

-- No anon DELETE policy. Frontend cleanup calls should fail harmlessly.
-- Clean orphan previews with a scheduled server-side job or manual SQL job
-- using a privileged role.

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
  and tablename = 'objects'
order by policyname;

rollback;

-- Replace ROLLBACK with COMMIT only after testing preview upload and preview URL creation.
