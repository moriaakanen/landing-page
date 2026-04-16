/**
 * ROLE SWITCHER — nova-role-switcher.js
 * Tambahkan file ini ke semua halaman Portal NOVA.
 * Diload SETELAH config.js
 *
 * Cara pakai:
 * 1. Tambahkan <script src="nova-role-switcher.js"></script> setelah config.js
 * 2. Ganti blok topbar-right dengan HTML di bawah (lihat komentar TOPBAR HTML)
 * 3. Hapus tombol <button class="btn-logout"> yang lama (sudah ada di dalam dropdown)
 * 4. Panggil initRoleSwitcher(session, isAdminPage) di akhir fungsi init()
 */

/* ═══════════════════════════════════════════
   TOPBAR HTML — ganti .topbar-right dengan ini
   ═══════════════════════════════════════════

<div class="topbar-right">
  <span class="topbar-time" id="topbar-time"></span>
  <div class="user-switcher" id="user-switcher">
    <button class="user-switcher-btn" id="user-switcher-btn"
      onclick="toggleUserDropdown()" aria-expanded="false" type="button">
      <div class="topbar-avatar" id="topbar-avatar">—</div>
      <span class="topbar-username" id="topbar-username">Pengguna</span>
      <span class="active-role-badge" id="active-role-badge">Admin</span>
      <span class="user-switcher-caret">▼</span>
    </button>
    <div class="user-switcher-dropdown" id="user-switcher-dropdown">
      <div class="usd-header">
        <div class="usd-header-name" id="usd-header-name">Pengguna</div>
        <div class="usd-header-label">Akun aktif · Ganti tampilan role</div>
      </div>
      <div class="usd-section-label">Tampilkan sebagai</div>
      <button class="usd-role-option" id="usd-opt-admin" onclick="switchViewRole('admin')" type="button">
        <div class="usd-role-icon admin">👑</div>
        <div><div class="usd-role-name">Administrator</div><div class="usd-role-desc">Akses penuh semua menu</div></div>
        <span class="usd-check">✓</span>
      </button>
      <button class="usd-role-option" id="usd-opt-user" onclick="switchViewRole('user')" type="button">
        <div class="usd-role-icon user">👤</div>
        <div><div class="usd-role-name">User</div><div class="usd-role-desc">Akses standar pegawai</div></div>
        <span class="usd-check">✓</span>
      </button>
      <div class="usd-divider"></div>
      <button class="usd-logout" onclick="logout()" type="button">
        <span>🚪</span> Keluar dari Portal
      </button>
    </div>
  </div>
  <button class="menu-toggle" onclick="document.getElementById('sidebar').classList.toggle('open')">☰</button>
</div>

*/

/* ═══════════════════════════════════════════
   CSS — tambahkan ke dalam <style> setiap halaman
   ═══════════════════════════════════════════ */
(function injectCSS() {
  const css = `
  .user-switcher{position:relative}
  .user-switcher-btn{display:flex;align-items:center;gap:8px;padding:5px 10px 5px 6px;border-radius:7px;cursor:pointer;transition:background .2s;border:1px solid transparent;background:transparent;font-family:inherit}
  .user-switcher-btn:hover{background:rgba(255,255,255,.08);border-color:rgba(255,255,255,.1)}
  .user-switcher-caret{font-size:9px;color:rgba(255,255,255,.4);transition:transform .2s;margin-left:2px}
  .user-switcher-btn[aria-expanded="true"] .user-switcher-caret{transform:rotate(180deg)}
  .user-switcher-dropdown{position:absolute;top:calc(100% + 8px);right:0;background:#fff;border:1px solid #e2ddd6;border-radius:10px;box-shadow:0 12px 40px rgba(13,35,64,.18);min-width:240px;overflow:hidden;z-index:500;animation:usdDropDown .15s ease;display:none}
  .user-switcher-dropdown.open{display:block}
  @keyframes usdDropDown{from{opacity:0;transform:translateY(-6px)}to{opacity:1;transform:translateY(0)}}
  .usd-header{padding:14px 16px;border-bottom:1px solid #e2ddd6;background:#fafaf8}
  .usd-header-name{font-size:13px;font-weight:700;color:#0d2340}
  .usd-header-label{font-size:11px;color:#6b7280;margin-top:2px}
  .usd-section-label{font-size:10px;font-weight:600;color:#6b7280;letter-spacing:1.2px;text-transform:uppercase;padding:10px 16px 4px}
  .usd-role-option{display:flex;align-items:center;gap:10px;padding:9px 16px;cursor:pointer;transition:background .15s;border:none;background:transparent;width:100%;text-align:left;font-family:inherit}
  .usd-role-option:hover{background:#f5f4f0}
  .usd-role-option.active{background:#f5edda}
  .usd-role-icon{width:32px;height:32px;border-radius:7px;display:flex;align-items:center;justify-content:center;font-size:14px;flex-shrink:0}
  .usd-role-icon.user{background:#f1f5f9}
  .usd-role-icon.admin{background:rgba(200,168,75,.15)}
  .usd-role-name{font-size:13px;font-weight:600;color:#0d2340}
  .usd-role-desc{font-size:11px;color:#6b7280}
  .usd-check{margin-left:auto;font-size:13px;color:#c8a84b;display:none}
  .usd-role-option.active .usd-check{display:block}
  .usd-divider{height:1px;background:#e2ddd6;margin:4px 0}
  .usd-logout{display:flex;align-items:center;gap:8px;padding:9px 16px;cursor:pointer;transition:background .15s;border:none;background:transparent;width:100%;text-align:left;font-family:inherit;font-size:13px;color:#991b1b;font-weight:500}
  .usd-logout:hover{background:#fef2f2}
  .active-role-badge{display:inline-flex;align-items:center;gap:4px;padding:2px 7px;border-radius:100px;font-size:10px;font-weight:600;letter-spacing:.3px;margin-left:4px}
  .active-role-badge.admin{background:rgba(200,168,75,.2);color:#7a5c10;border:1px solid rgba(200,168,75,.4)}
  .active-role-badge.user{background:rgba(255,255,255,.1);color:rgba(255,255,255,.6);border:1px solid rgba(255,255,255,.15)}
  `;
  const style = document.createElement('style');
  style.textContent = css;
  document.head.appendChild(style);
})();

/* ═══════════════════════════════════════════
   FUNGSI UTAMA
   ═══════════════════════════════════════════ */

/**
 * Ambil role aktif dari session
 */
function getActiveRole() {
  try {
    const s = JSON.parse(localStorage.getItem('nova_user') || 'null');
    return s?.active_role || (ADMIN_USERS.includes(s?.username) ? 'admin' : 'user');
  } catch(e) { return 'user'; }
}

/**
 * Ganti role tampilan
 * @param {string} role - 'admin' | 'user'
 */
function switchViewRole(role) {
  try {
    const s = JSON.parse(localStorage.getItem('nova_user') || 'null');
    if (!s) return;
    if (role === 'admin' && !ADMIN_USERS.includes(s.username)) return;
    s.active_role = role;
    localStorage.setItem('nova_user', JSON.stringify(s));
  } catch(e) {}

  closeUserDropdown();

  // Redirect ke halaman yang sesuai
  if (role === 'user') {
    window.location.replace('index.html');
  } else {
    // Jika sudah di halaman admin, cukup reload UI
    applyRoleSidebar(role);
    applyRoleBadge(role);
  }
}

/**
 * Update tampilan sidebar sesuai role
 */
function applyRoleSidebar(role) {
  const isAdmin = role === 'admin';
  const adminNav = document.getElementById('admin-nav');
  if (adminNav) adminNav.style.display = isAdmin ? 'block' : 'none';
  const sidebarBadge = document.getElementById('sidebar-admin-badge');
  if (sidebarBadge) sidebarBadge.style.display = isAdmin ? 'inline-flex' : 'none';
}

/**
 * Update badge role di topbar
 */
function applyRoleBadge(role) {
  const isAdmin = role === 'admin';
  const badge = document.getElementById('active-role-badge');
  if (badge) {
    badge.textContent = isAdmin ? 'Admin' : 'User';
    badge.className = `active-role-badge ${role}`;
  }
  const optAdmin = document.getElementById('usd-opt-admin');
  const optUser = document.getElementById('usd-opt-user');
  if (optAdmin) optAdmin.classList.toggle('active', isAdmin);
  if (optUser) optUser.classList.toggle('active', !isAdmin);
}

/**
 * Inisialisasi role switcher — panggil ini di init() setiap halaman
 * @param {object} session - object session dari localStorage
 * @param {boolean} isAdminPage - apakah halaman ini membutuhkan akses admin
 */
function initRoleSwitcher(session, isAdminPage = false) {
  if (!session) return;
  // Set default active_role jika belum ada
  if (!session.active_role) {
    session.active_role = ADMIN_USERS.includes(session.username) ? 'admin' : 'user';
    localStorage.setItem('nova_user', JSON.stringify(session));
  }

  // Redirect jika halaman admin tapi role aktif adalah user
  if (isAdminPage && session.active_role !== 'admin') {
    window.location.replace('index.html');
    return;
  }

  // Update nama di dropdown
  const name = session.full_name || session.username || 'Pengguna';
  const usdName = document.getElementById('usd-header-name');
  if (usdName) usdName.textContent = name;

  // Sembunyikan opsi admin jika bukan admin
  const optAdmin = document.getElementById('usd-opt-admin');
  if (optAdmin && !ADMIN_USERS.includes(session.username)) {
    optAdmin.style.display = 'none';
  }

  applyRoleBadge(session.active_role);
  applyRoleSidebar(session.active_role);
}

function toggleUserDropdown() {
  const dropdown = document.getElementById('user-switcher-dropdown');
  const btn = document.getElementById('user-switcher-btn');
  if (!dropdown || !btn) return;
  const isOpen = dropdown.classList.contains('open');
  dropdown.classList.toggle('open', !isOpen);
  btn.setAttribute('aria-expanded', String(!isOpen));
}

function closeUserDropdown() {
  const dropdown = document.getElementById('user-switcher-dropdown');
  const btn = document.getElementById('user-switcher-btn');
  if (dropdown) dropdown.classList.remove('open');
  if (btn) btn.setAttribute('aria-expanded', 'false');
}

document.addEventListener('click', function(e) {
  const switcher = document.getElementById('user-switcher');
  if (switcher && !switcher.contains(e.target)) closeUserDropdown();
});
