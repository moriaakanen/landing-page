/**
 * NOVA TOPBAR — komponen shared
 * Diload SETELAH config.js, SEBELUM nova-sidebar.js dan nova-role-switcher.js.
 *
 * Cara pakai di setiap halaman:
 *   1. Pasang   <div id="topbar-mount"></div>   di awal <body>
 *   2. Sertakan <script src="nova-topbar.js"></script>
 *   3. Setelah session valid, panggil  NovaTopbar.setUser(session)
 *
 * Komponen ini meng-inject:
 *   - CSS topbar (struktur, brand, time, role switcher button, menu-toggle)
 *   - HTML topbar lengkap dengan role switcher dropdown
 *   - Clock yg auto-update tiap detik
 *
 * CSS pages tidak perlu lagi mendefinisikan style topbar/menu-toggle/topbar-*.
 * (Boleh tetap ada — selectorless yg di-inject akan tertimpa oleh page CSS
 * jika diperlukan, tapi defaultnya sudah konsisten antar halaman.)
 */
(function () {
  'use strict';

  // ─── CSS injection ───────────────────────────────────────────────
  // Hanya CSS yg menyangkut topbar. Variable warna (--navy, --gold, dst)
  // diasumsikan sudah didefinisikan di :root tiap halaman (sesuai pola
  // existing). Kalau belum, fallback warna hardcoded dipakai.
  var CSS = `
  :where(.topbar){height:52px;background:var(--navy,#0d2340);display:flex;align-items:center;justify-content:space-between;padding:0 24px 0 0;flex-shrink:0;position:sticky;top:0;z-index:200;font-family:'Plus Jakarta Sans',sans-serif;font-size:13px;line-height:1.4}
  :where(.topbar-brand){width:var(--sidebar-w,260px);display:flex;align-items:center;gap:10px;padding:0 20px;border-right:1px solid rgba(255,255,255,.07);height:100%;flex-shrink:0}
  :where(.brand-emblem){width:30px;height:30px;background:var(--gold,#c8a84b);border-radius:5px;display:flex;align-items:center;justify-content:center;font-size:14px;flex-shrink:0;color:#fff}
  :where(.brand-name){font-size:14px;font-weight:700;color:#fff;letter-spacing:.5px;line-height:1}
  :where(.brand-tag){font-size:9px;color:rgba(255,255,255,.35);letter-spacing:1px;text-transform:uppercase;line-height:1.4;margin-top:2px}
  :where(.topbar-right){display:flex;align-items:center;gap:14px;padding-right:0}
  :where(.topbar-time){font-size:11px;color:rgba(255,255,255,.45);font-variant-numeric:tabular-nums;white-space:nowrap}
  :where(.topbar-avatar){width:30px;height:30px;background:var(--navy2,#163358);border:1.5px solid rgba(200,168,75,.4);border-radius:50%;display:flex;align-items:center;justify-content:center;font-size:11px;font-weight:700;color:var(--gold,#c8a84b);flex-shrink:0}
  :where(.topbar-username){font-size:13px;color:rgba(255,255,255,.78);font-weight:500;white-space:nowrap}
  :where(.menu-toggle){display:none;align-items:center;justify-content:center;width:36px;height:36px;background:rgba(255,255,255,.07);border:none;border-radius:6px;cursor:pointer;color:#fff;font-size:16px;margin-left:4px;font-family:inherit}
  :where(.menu-toggle:hover){background:rgba(255,255,255,.12)}
  @media(max-width:768px){:where(.topbar-time){display:none}}
  @media(max-width:600px){:where(.menu-toggle){display:flex!important}}
  `;
  if (!document.getElementById('nova-topbar-css')) {
    var style = document.createElement('style');
    style.id = 'nova-topbar-css';
    style.textContent = CSS;
    document.head.appendChild(style);
  }

  // ─── HTML template ───────────────────────────────────────────────
  // NB: role switcher CSS (.user-switcher-*, .usd-*, dst.) di-inject
  // oleh nova-role-switcher.js. Jadi script itu tetap perlu dimuat.
  var TOPBAR_HTML = `
  <div class="topbar">
    <div class="topbar-brand">
      <div class="brand-emblem">🏛</div>
      <div>
        <div class="brand-name">PORTAL NOVA</div>
        <div class="brand-tag">Sistem Informasi Terpadu</div>
      </div>
    </div>
    <div class="topbar-right">
      <span class="topbar-time" id="topbar-time"></span>
      <div class="user-switcher" id="user-switcher">
        <button class="user-switcher-btn" id="user-switcher-btn" onclick="toggleUserDropdown()" aria-expanded="false" type="button">
          <div class="topbar-avatar" id="topbar-avatar">—</div>
          <span class="topbar-username" id="topbar-username">Pengguna</span>
          <span class="active-role-badge admin" id="active-role-badge">Admin</span>
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
          <button class="usd-logout" onclick="logout()" type="button"><span>🚪</span> Keluar dari Portal</button>
        </div>
      </div>
      <button class="menu-toggle" onclick="document.getElementById('sidebar') && document.getElementById('sidebar').classList.toggle('open')" aria-label="Buka menu">☰</button>
    </div>
  </div>
  `;

  // ─── Render helper ───────────────────────────────────────────────
  function render() {
    // Cari mount point baru terlebih dulu
    var mount = document.getElementById('topbar-mount');
    if (mount) {
      mount.outerHTML = TOPBAR_HTML;
      return true;
    }
    // Backward-compat: kalau halaman lama masih punya <div class="topbar"> manual,
    // biarkan saja (jangan ditimpa) supaya tidak merusak. Hanya log warning.
    var existing = document.querySelector('.topbar');
    if (existing) {
      console.info('[NovaTopbar] Detected legacy .topbar markup. Topbar JS template not applied.');
      return false;
    }
    // Tidak ada mount point sama sekali — diam saja (mis. login.html).
    return false;
  }

  // ─── Clock ───────────────────────────────────────────────────────
  function updateClock() {
    var el = document.getElementById('topbar-time');
    if (!el) return;
    el.textContent = new Date().toLocaleString('id-ID', {
      weekday: 'long', day: 'numeric', month: 'long', year: 'numeric',
      hour: '2-digit', minute: '2-digit'
    });
  }

  // ─── Set user info di topbar ─────────────────────────────────────
  function setUser(session) {
    if (!session) return;
    var name = session.full_name || session.username || 'Pengguna';
    var initials = name.split(' ').filter(Boolean)
      .map(function (w) { return w[0]; }).join('')
      .toUpperCase().slice(0, 2) || '—';
    var avatarEl = document.getElementById('topbar-avatar');
    var unameEl  = document.getElementById('topbar-username');
    if (avatarEl) avatarEl.textContent = initials;
    if (unameEl)  unameEl.textContent  = name;
  }

  // ─── Init ────────────────────────────────────────────────────────
  function init() {
    var rendered = render();
    if (rendered) {
      updateClock();
      setInterval(updateClock, 1000);
    }
  }

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', init);
  } else {
    init();
  }

  // ─── Public API ──────────────────────────────────────────────────
  window.NovaTopbar = {
    setUser: setUser,
    render:  render,
    updateClock: updateClock
  };
})();
