/**
 * NOVA SIDEBAR — komponen shared
 * Diload SETELAH config.js, SEBELUM nova-role-switcher.js dan script init() halaman.
 *
 * Cara pakai di setiap halaman:
 *   1. Pasang   <aside class="sidebar" id="sidebar" data-active="KEY"></aside>
 *   2. Sertakan <script src="nova-sidebar.js"></script> sebelum init()
 *
 * data-active KEY yang valid:
 *   - "profil"            → Profil Saya (index.html)
 *   - "surat-tugas"       → Surat Tugas user (surat-tugas.html)
 *   - "admin-surat"       → Persetujuan Surat (admin-surat-tugas.html)
 *   - "kamus-pok"         → Kamus POK (manajemen-kamus-pok.html)
 *   - "manajemen-pengguna"→ Manajemen Pengguna (manajemen-pengguna.html)
 *
 * Kalau ada menu admin baru di masa depan, cukup tambah row di NAV_ADMIN_ITEMS
 * di bawah ini — semua halaman otomatis dapat link tersebut. TIDAK perlu lagi
 * edit setiap file HTML.
 */
(function () {
  'use strict';

  // Daftar menu untuk USER (selalu terlihat untuk siapa saja yang login)
  var NAV_USER_ITEMS = [
    { key: 'profil',        href: 'index.html',        icon: '👤', label: 'Profil Saya',  group: 'Menu Utama' },
    { key: 'surat-tugas',   href: 'surat-tugas.html',  icon: '📄', label: 'Minta Surat Tugas',  group: 'Persuratan' },
  ];

  // Daftar menu untuk ADMIN (di-show/hide oleh nova-role-switcher.js
  // berdasarkan active_role)
  var NAV_ADMIN_ITEMS = [
    { key: 'admin-surat',         href: 'admin-surat-tugas.html',    icon: '✅', label: 'Surat Tugas' },
    { key: 'kamus-pok',           href: 'manajemen-kamus-pok.html',  icon: '📚', label: 'Kamus POK' },
    { key: 'manajemen-pengguna',  href: 'manajemen-pengguna.html',   icon: '👥', label: 'Manajemen Pengguna' },
  ];

  function escAttr(s) {
    return String(s).replace(/[&<>"']/g, function (c) {
      return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c];
    });
  }

  function buildItem(item, activeKey) {
    var isActive = item.key === activeKey;
    return '<a class="nav-item' + (isActive ? ' active' : '') + '" ' +
           'href="' + escAttr(item.href) + '">' +
           '<div class="nav-item-icon">' + item.icon + '</div>' +
           item.label +
           '</a>';
  }

  function buildSidebar(activeKey) {
    var html = '';

    // Group user items by `group` field, render section labels
    var lastGroup = null;
    NAV_USER_ITEMS.forEach(function (item) {
      if (item.group !== lastGroup) {
        var labelClass = lastGroup === null ? 'sidebar-section-label' : 'nav-group-label';
        html += '<div class="' + labelClass + '">' + item.group + '</div>';
        lastGroup = item.group;
      }
      html += buildItem(item, activeKey);
    });

    // Admin section — defaultnya HIDDEN. nova-role-switcher.js akan
    // unhide via applyRoleSidebar() kalau active_role === 'admin'.
    html += '<div id="admin-nav" style="display:none">';
    html += '<div class="nav-group-label">Administrasi</div>';
    NAV_ADMIN_ITEMS.forEach(function (item) {
      html += buildItem(item, activeKey);
    });
    html += '</div>';

    // Footer — admin-badge juga di-show/hide oleh role switcher
    html += '<div class="sidebar-footer">' +
            '<strong>Portal NOVA</strong><br>' +
            'Sistem Informasi Terpadu<br>' +
            '© 2026 Pemerintah RI' +
            '<div id="admin-footer-badge" style="display:none;margin-top:6px">' +
              '<span class="admin-badge" id="sidebar-admin-badge">🔑 Administrator</span>' +
            '</div>' +
            '</div>';

    return html;
  }

  function render() {
    var aside = document.getElementById('sidebar');
    if (!aside) return;
    var activeKey = aside.getAttribute('data-active') || '';
    aside.innerHTML = buildSidebar(activeKey);
  }

  // Render segera. Karena script ini di-include SETELAH <aside id="sidebar">
  // di markup HTML, element pasti sudah ada saat script ini dieksekusi.
  // Kalau ternyata DOM belum siap (script di-load via async/defer), pakai
  // listener DOMContentLoaded sebagai fallback.
  if (document.getElementById('sidebar')) {
    render();
  } else if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', render);
  } else {
    // DOM sudah siap tapi #sidebar tidak ditemukan — diam saja.
    // Halaman ini mungkin tidak punya sidebar (mis. login.html, ganti-password.html).
  }

  // Expose untuk debugging / re-render manual kalau dibutuhkan
  window.NovaSidebar = { render: render };
})();
