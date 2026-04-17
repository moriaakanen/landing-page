// ═══════════════════════════════════════════════
// config.js — Portal NOVA
// Konfigurasi terpusat. Ubah nilai di bawah ini
// sesuai project Supabase Anda.
// ═══════════════════════════════════════════════
const SUPABASE_URL      = 'https://jsmmtqeoukkgugorrvmg.supabase.co';
const SUPABASE_ANON_KEY = 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImpzbW10cWVvdWtrZ3Vnb3Jydm1nIiwicm9sZSI6ImFub24iLCJpYXQiOjE3NzYwOTg2NzksImV4cCI6MjA5MTY3NDY3OX0.O03NscRcj4RcNx-P3j65hO7XXLRSkbyJcwpcArpqHBQ';

// ─────────────────────────────────────────────────────────────
// CATATAN PENTING:
// Sumber kebenaran role adalah DATABASE (kolom `roles` text[] di
// tabel users). Jangan gunakan daftar hardcoded untuk menentukan
// siapa admin — atur role melalui halaman Manajemen Pengguna.
//
// ADMIN_USERS di bawah ini HANYA dipakai sebagai safety-net saat
// bootstrap (misalnya DB masih kosong dan Anda perlu login pertama
// kali sebagai admin). Setelah sistem berjalan, array ini boleh
// dikosongkan: `const ADMIN_USERS = [];`
// ─────────────────────────────────────────────────────────────
const ADMIN_USERS = ['rizal.akbar'];

/**
 * Helper universal untuk mengecek role dari object session.
 * Prioritas:
 *   1. session.roles (array) — sumber utama dari DB
 *   2. session.role  (string) — primary role
 *   3. ADMIN_USERS  — bootstrap fallback
 */
function userHasRole(session, role) {
  if (!session) return false;
  if (Array.isArray(session.roles) && session.roles.includes(role)) return true;
  if (session.role === role) return true;
  if (role === 'admin' && ADMIN_USERS.includes(session.username)) return true;
  if (role === 'user') return true; // setiap user yang login setidaknya punya role user
  return false;
}

function getUserRoles(session) {
  if (!session) return ['user'];
  if (Array.isArray(session.roles) && session.roles.length) return session.roles;
  if (session.role) {
    // Jika hanya punya role admin dari DB, asumsikan juga bisa akses sebagai user
    return session.role === 'admin' ? ['admin', 'user'] : [session.role];
  }
  if (ADMIN_USERS.includes(session.username)) return ['admin', 'user'];
  return ['user'];
}
