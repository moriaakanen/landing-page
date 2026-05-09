/**
 * 9201 TOPBAR — komponen shared
 * Diload SETELAH config.js, SEBELUM 9201-sidebar.js dan 9201-role-switcher.js.
 *
 * Cara pakai di setiap halaman:
 *   1. Pasang   <div id="topbar-mount"></div>   di awal <body>
 *   2. Sertakan <script src="9201-topbar.js"></script>
 *   3. Setelah session valid, panggil  Topbar9201.setUser(session)
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
  :where(.brand-emblem){width:30px;height:30px;background:var(--gold,#c8a84b);border-radius:5px;display:flex;align-items:center;justify-content:center;flex-shrink:0;color:#fff;padding:5px}
  :where(.brand-emblem) svg{width:100%;height:100%}
  :where(.brand-name){font-size:14px;font-weight:700;color:#fff;letter-spacing:.5px;line-height:1}
  :where(.brand-tag){font-size:9px;color:rgba(255,255,255,.35);letter-spacing:1px;text-transform:uppercase;line-height:1.4;margin-top:2px}
  :where(.topbar-right){display:flex;align-items:center;gap:14px;padding-right:0}
  :where(.topbar-time){font-size:11px;color:rgba(255,255,255,.45);font-variant-numeric:tabular-nums;white-space:nowrap}
  :where(.topbar-avatar){width:30px;height:30px;background:var(--navy2,#163358);border:1.5px solid rgba(200,168,75,.4);border-radius:50%;display:flex;align-items:center;justify-content:center;font-size:11px;font-weight:700;color:var(--gold,#c8a84b);flex-shrink:0}
  :where(.topbar-username){font-size:13px;color:rgba(255,255,255,.78);font-weight:500;white-space:nowrap}

  /* ═══ NOTIFIKASI BELL ═══
     CATATAN PENTING: jangan pakai :where() di sini. :where() punya
     specificity 0 yang gampang di-override secara tidak sengaja oleh
     CSS halaman. Pakai selector langsung supaya konsisten antar
     halaman dan dropdown selalu muncul saat .open class ditambahkan. */
  .notif-wrap{position:relative;display:flex;align-items:center}
  .notif-btn{background:transparent;border:none;cursor:pointer;width:38px;height:38px;border-radius:50%;display:flex;align-items:center;justify-content:center;color:rgba(255,255,255,.7);transition:background .15s,color .15s;position:relative;font-size:17px;padding:0;font-family:inherit}
  .notif-btn:hover{background:rgba(255,255,255,.08);color:#fff}
  .notif-btn.has-unread{color:#fff}
  .notif-bell{display:inline-block;line-height:1;pointer-events:none}
  .notif-btn.has-unread .notif-bell{animation:bellShake .6s ease-in-out}
  @keyframes bellShake{0%,100%{transform:rotate(0)}20%{transform:rotate(-12deg)}40%{transform:rotate(10deg)}60%{transform:rotate(-6deg)}80%{transform:rotate(4deg)}}
  .notif-badge{position:absolute;top:5px;right:4px;min-width:17px;height:17px;padding:0 4px;background:#dc2626;color:#fff;border-radius:100px;font-size:9.5px;font-weight:700;display:flex;align-items:center;justify-content:center;border:1.5px solid var(--navy,#0d2340);font-variant-numeric:tabular-nums;letter-spacing:-.2px;pointer-events:none}
  .notif-badge[hidden]{display:none !important}

  /* Dropdown — pakai !important untuk display agar tidak ditiban CSS halaman */
  .notif-dropdown{position:absolute;top:calc(100% + 8px);right:0;width:360px;max-width:calc(100vw - 32px);background:#fff;border-radius:11px;box-shadow:0 14px 40px rgba(13,35,64,.22);border:1px solid rgba(13,35,64,.08);z-index:1000;display:none;overflow:hidden;font-family:'Plus Jakarta Sans',sans-serif;color:#0d2340}
  .notif-dropdown.open{display:block !important;animation:notifPop .18s cubic-bezier(.2,.7,.3,1) both}
  @keyframes notifPop{from{opacity:0;transform:translateY(-6px) scale(.98)}to{opacity:1;transform:translateY(0) scale(1)}}
  .notif-header{padding:13px 16px;display:flex;align-items:center;justify-content:space-between;border-bottom:1px solid var(--border,#e2ddd6);background:#fafaf6;color:var(--navy,#0d2340);font-weight:700;font-size:13px;letter-spacing:.2px}
  .notif-header-meta{font-size:11px;color:var(--muted,#6b7280);font-weight:500}
  .notif-list{max-height:420px;overflow-y:auto;background:#fff}
  .notif-empty{padding:36px 18px;text-align:center;color:var(--muted,#6b7280);font-size:13px;line-height:1.6}
  .notif-empty-icon{font-size:30px;margin-bottom:8px;opacity:.45}
  .notif-item{display:flex;gap:11px;padding:11px 14px;border-bottom:1px solid var(--border,#e2ddd6);cursor:pointer;transition:background .12s;text-decoration:none;color:inherit;align-items:flex-start}
  .notif-item:hover{background:#fafaf6}
  .notif-item:last-child{border-bottom:none}
  .notif-item.unread{background:#fffbeb}
  .notif-item.unread:hover{background:#fef3c7}
  .notif-item-icon{flex-shrink:0;width:32px;height:32px;border-radius:8px;background:var(--bg,#f5f4f0);display:flex;align-items:center;justify-content:center;font-size:14px;border:1px solid var(--border,#e2ddd6)}
  .notif-item-body{flex:1;min-width:0}
  .notif-item-title{font-size:12.5px;font-weight:600;color:var(--navy,#0d2340);line-height:1.4;margin-bottom:2px}
  .notif-item-desc{font-size:11.5px;color:var(--muted,#6b7280);line-height:1.5;margin-bottom:3px;word-wrap:break-word}
  .notif-item-time{font-size:10.5px;color:var(--muted,#6b7280);font-weight:500;font-variant-numeric:tabular-nums;display:flex;align-items:center;gap:4px}
  .notif-item.unread .notif-item-title{color:var(--navy,#0d2340);font-weight:700}
  .notif-item-dot{width:7px;height:7px;border-radius:50%;background:#dc2626;flex-shrink:0;margin-top:7px;display:none}
  .notif-item.unread .notif-item-dot{display:block}

  .notif-footer{padding:9px 14px;border-top:1px solid var(--border,#e2ddd6);background:#fafaf6;text-align:center}
  .notif-footer-link{background:transparent;border:none;color:var(--navy,#0d2340);font-family:inherit;font-size:11.5px;font-weight:600;cursor:pointer;letter-spacing:.2px;padding:4px 10px;border-radius:5px;transition:background .15s}
  .notif-footer-link:hover{background:rgba(13,35,64,.06)}
  @media(max-width:480px){
    .notif-dropdown{width:calc(100vw - 24px);right:-12px}
  }
  :where(.menu-toggle){display:none;align-items:center;justify-content:center;width:36px;height:36px;background:rgba(255,255,255,.07);border:none;border-radius:6px;cursor:pointer;color:#fff;font-size:16px;margin-left:4px;font-family:inherit}
  :where(.menu-toggle:hover){background:rgba(255,255,255,.12)}
  @media(max-width:768px){:where(.topbar-time){display:none}}
  @media(max-width:600px){:where(.menu-toggle){display:flex!important}}
  `;
  if (!document.getElementById('9201-topbar-css')) {
    var style = document.createElement('style');
    style.id = '9201-topbar-css';
    style.textContent = CSS;
    document.head.appendChild(style);
  }

  // ─── HTML template ───────────────────────────────────────────────
  // NB: role switcher CSS (.user-switcher-*, .usd-*, dst.) di-inject
  // oleh 9201-role-switcher.js. Jadi script itu tetap perlu dimuat.
  var TOPBAR_HTML = `
  <div class="topbar">
    <div class="topbar-brand">
      <div class="brand-emblem" aria-label="Logo">
        <svg viewBox="0 0 32 32" width="20" height="20" fill="currentColor" xmlns="http://www.w3.org/2000/svg" aria-hidden="true">
          <path d="M 5 15 C 2 13, 1 9, 3 4 C 4 8, 6 11, 9 13 Z" opacity="0.55"/>
          <path d="M 6 18 C 3 16, 2 11, 4 7 C 5 10, 7 14, 10 16 Z" opacity="0.85"/>
          <path d="M 8 18 Q 8 12, 13 11 L 19 11 Q 23 12, 23 16 Q 23 20, 19 21 L 14 22 Q 8 22, 8 18 Z"/>
          <circle cx="20" cy="11" r="3.2"/>
          <path d="M 18 8 Q 18 6, 19 6 Q 19 5, 20 5.5 Q 20 4.5, 21 5 Q 21 6, 22 6.5 L 22 8 Z"/>
          <path d="M 23 11 L 25.5 10.5 L 23 12.5 Z"/>
          <path d="M 22 12.5 Q 22.5 14.5, 21.5 14.5 L 21 12.5 Z"/>
          <circle cx="20.2" cy="10.5" r="0.65" fill="white"/>
          <rect x="13" y="22" width="1" height="4" rx="0.3"/>
          <rect x="17.5" y="22" width="1" height="4" rx="0.3"/>
          <path d="M 11.5 26 L 14.5 26 L 14 26.5 L 12 26.5 Z"/>
          <path d="M 16 26 L 19 26 L 18.5 26.5 L 16.5 26.5 Z"/>
        </svg>
      </div>
      <div>
        <div class="brand-name">9201</div>
        <div class="brand-tag">BPS Kabupaten Raja Ampat</div>
      </div>
    </div>
    <div class="topbar-right">
      <span class="topbar-time" id="topbar-time"></span>

      <!-- ═══ Lonceng notifikasi — di-handle oleh 9201-notifikasi.js ═══ -->
      <div class="notif-wrap" id="notif-wrap">
        <button class="notif-btn" id="notif-btn" type="button" aria-label="Notifikasi" aria-expanded="false">
          <span class="notif-bell" aria-hidden="true">🔔</span>
          <span class="notif-badge" id="notif-badge" hidden>0</span>
        </button>
        <div class="notif-dropdown" id="notif-dropdown" role="menu" aria-hidden="true">
          <div class="notif-header">
            <span>Notifikasi</span>
            <span class="notif-header-meta" id="notif-header-meta">—</span>
          </div>
          <div class="notif-list" id="notif-list">
            <div class="notif-empty"><div class="notif-empty-icon">🔕</div>Tidak ada notifikasi.</div>
          </div>
        </div>
      </div>

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
      console.info('[Topbar9201] Detected legacy .topbar markup. Topbar JS template not applied.');
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
  // Update visual berdasarkan apa yg ada di session SAAT ITU (sync, instant).
  //
  // Auto-resolve dari tabel `users.full_name`:
  //   - Kalau session.full_name kosong/null → trigger fetch
  //   - Kalau session.full_name === session.username (hasil fallback saat
  //     login lama waktu RPC return null) → trigger fetch juga
  //
  // opts.skipEnsure = true mencegah infinite loop saat dipanggil dari
  // dalam novaEnsureFullName.
  function setUser(session, opts) {
    if (!session) return;
    var skipEnsure = !!(opts && opts.skipEnsure);

    var name = session.full_name || session.username || '';
    var initials = name
      ? name.split(' ').filter(Boolean)
            .map(function (w) { return w[0]; }).join('')
            .toUpperCase().slice(0, 2)
      : '';
    if (!initials) initials = '—';

    var avatarEl = document.getElementById('topbar-avatar');
    var unameEl  = document.getElementById('topbar-username');
    if (avatarEl) avatarEl.textContent = initials;
    if (unameEl)  unameEl.textContent  = name || 'Pengguna';

    if (skipEnsure) return;

    if (typeof novaEnsureFullName === 'function') {
      // Kondisi yang memerlukan fetch fresh dari users:
      //   - full_name belum ada sama sekali, ATAU
      //   - full_name kebetulan sama dengan username (hasil fallback)
      //
      // Kasus ke-2 muncul ketika user pernah login dengan RPC verify_login
      // yang return user.full_name = null, sehingga login.html menyimpan
      // session.full_name = session.username sebagai fallback. Tanpa
      // deteksi ini, topbar akan permanen menampilkan username walau
      // tabel users sebenarnya punya full_name yang valid.
      var needsFetch = !session.full_name
                    || (session.username && session.full_name === session.username);
      if (needsFetch) {
        novaEnsureFullName(session, { force: true });
      }
    }
  }

  // ─── Init ────────────────────────────────────────────────────────
  // PENTING: render dipanggil DUA KALI:
  //   1. Synchronously saat IIFE jalan — kalau mount point sudah ada di DOM
  //      (kasus umum: <div id="topbar-mount"> di awal body, script di-load
  //      setelahnya). Ini mencegah race condition di mana inline page
  //      script men-call Topbar9201.setUser() / initRoleSwitcher() SEBELUM
  //      topbar di-render — dulu hal ini menyebabkan badge role tetap
  //      "Admin" walau user sudah switch ke "User".
  //   2. Lewat DOMContentLoaded — fallback kalau script di-load via defer/
  //      async atau halaman struktur-nya tidak biasa.
  //
  // render() idempotent: hanya jalan kalau #topbar-mount masih ada.
  // Setelah render pertama, mount sudah outerHTML'd jadi <div class="topbar">,
  // dan render kedua tidak akan apa-apa.
  function init() {
    var rendered = render();
    if (rendered) {
      updateClock();
      setInterval(updateClock, 1000);
      // Notify dependent scripts (notifikasi, role-switcher fallback, dst.)
      try {
        document.dispatchEvent(new CustomEvent('9201:topbar:rendered'));
      } catch (_) {
        // IE fallback (tidak ada di portal target, tapi defensif)
        var ev = document.createEvent('Event');
        ev.initEvent('9201:topbar:rendered', true, true);
        document.dispatchEvent(ev);
      }
    }
  }

  // Render IMMEDIATELY kalau mount sudah ada (script di akhir body / async load).
  if (document.getElementById('topbar-mount')) {
    init();
  } else if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', init);
  } else {
    init();
  }

  // ─── Public API ──────────────────────────────────────────────────
  window.Topbar9201 = {
    setUser: setUser,
    render:  render,
    updateClock: updateClock
  };
})();
