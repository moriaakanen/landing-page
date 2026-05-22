# Portal 9201

Aplikasi statis HTML/JavaScript untuk portal internal yang terhubung ke Supabase.

## Menjalankan Lokal

Buka `login.html` atau jalankan static server dari folder repo.

## Konfigurasi

Konfigurasi Supabase ada di `config.js`. `SUPABASE_ANON_KEY` aman berada di frontend hanya jika Row Level Security dan RPC di Supabase sudah dikunci dengan benar. Jangan pernah menaruh service-role key di repo atau browser.

## Keamanan

Lihat `SECURITY.md` sebelum deploy. Halaman admin memiliki guard frontend, tetapi otorisasi final wajib ditegakkan di Supabase melalui RLS/RPC.
