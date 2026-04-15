/**
 * config.js — Portal NOVA
 * Konfigurasi terpusat. Jika URL/key berubah, cukup ubah di sini.
 *
 * CATATAN KEAMANAN (#2):
 * Lockout login saat ini masih di sisi klien (localStorage) dan mudah
 * di-bypass. Untuk keamanan sesungguhnya, tambahkan kolom
 * `login_attempts` dan `locked_until` di tabel users, lalu tangani
 * pemblokirannya di dalam RPC `verify_login` di Supabase.
 */
const SUPABASE_URL     = 'https://jsmmtqeoukkgugorrvmg.supabase.co';
const SUPABASE_ANON_KEY = 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImpzbW10cWVvdWtrZ3Vnb3Jydm1nIiwicm9sZSI6ImFub24iLCJpYXQiOjE3NzYwOTg2NzksImV4cCI6MjA5MTY3NDY3OX0.O03NscRcj4RcNx-P3j65hO7XXLRSkbyJcwpcArpqHBQ';
