// ═══════════════════════════════════════════════
// config.js — Portal NOVA
// ═══════════════════════════════════════════════
const SUPABASE_URL      = 'https://jsmmtqeoukkgugorrvmg.supabase.co';
const SUPABASE_ANON_KEY = 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImpzbW10cWVvdWtrZ3Vnb3Jydm1nIiwicm9sZSI6ImFub24iLCJpYXQiOjE3NzYwOTg2NzksImV4cCI6MjA5MTY3NDY3OX0.O03NscRcj4RcNx-P3j65hO7XXLRSkbyJcwpcArpqHBQ';

// Bootstrap fallback. Setelah sistem berjalan normal, boleh dikosongkan: []
const ADMIN_USERS = ['rizal.akbar'];

/**
 * Cek apakah session memiliki role tertentu.
 * Sumber: kolom `roles` (array) atau `role` (string) dari DB.
 */
function userHasRole(session, role) {
  if (!session) return false;
  if (Array.isArray(session.roles) && session.roles.includes(role)) return true;
  if (session.role === role) return true;
  if (role === 'admin' && ADMIN_USERS.includes(session.username)) return true;
  return false;
}

/**
 * Ambil daftar role yang dimiliki user dari object session.
 */
function getUserRoles(session) {
  if (!session) return ['user'];
  if (Array.isArray(session.roles) && session.roles.length) return session.roles;
  if (session.role) {
    return session.role === 'admin' ? ['admin', 'user'] : [session.role];
  }
  if (ADMIN_USERS.includes(session.username)) return ['admin', 'user'];
  return ['user'];
}
