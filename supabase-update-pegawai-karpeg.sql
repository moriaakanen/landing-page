-- RPC untuk memperbarui Nomor Seri Karpeg milik pegawai.
-- Jalankan sekali di Supabase SQL editor jika RLS mencegah PATCH langsung
-- ke public.data_pegawai dari halaman Profil Saya.

create or replace function public.update_pegawai_karpeg(
  p_nip text,
  p_karpeg text
)
returns public.data_pegawai
as $$
declare
  v_row public.data_pegawai;
begin
  update public.data_pegawai
     set karpeg = nullif(btrim(p_karpeg), '')
   where pegawai_nip = p_nip
   returning * into v_row;

  if not found then
    raise exception 'Pegawai dengan NIP % tidak ditemukan', p_nip
      using errcode = 'P0002';
  end if;

  return v_row;
end;
$$
language plpgsql
security definer
set search_path = public;

grant execute on function public.update_pegawai_karpeg(text, text) to anon, authenticated;
