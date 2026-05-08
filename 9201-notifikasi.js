/**
 * 9201 NOTIFIKASI — engine untuk lonceng 🔔 di topbar
 * ──────────────────────────────────────────────────────────────────
 * Cara kerja: derive list notifikasi on-the-fly dari tabel existing
 * (pengajuan_pak, surat_tugas), tidak ada tabel notifikasi terpisah.
 * Unread count = COUNT(notif.timestamp > users.notifikasi_last_read_at).
 *
 * Dipanggil otomatis saat halaman load (event 9201:topbar:rendered).
 * Tidak perlu setup manual di setiap halaman selama 9201-topbar.js
 * sudah di-load.
 *
 * Sumber notifikasi:
 *   USER (active_role='user'):
 *     - Pengajuan PAK miliknya yang status 'selesai' (di-approve)
 *     - Surat tugas miliknya yang status 'selesai'
 *   ADMIN (active_role='admin'):
 *     - Pengajuan PAK status 'menunggu' (semua user)
 *     - Surat tugas status 'menunggu' (semua user)
 *
 * Polling: refresh tiap 60 detik (cukup untuk pace-of-work portal ini,
 * tidak butuh realtime websocket).
 */
(function () {
  'use strict';

  if (window.__notif_init__) return;  // guard double-init
  window.__notif_init__ = true;

  // ─── State ───────────────────────────────────────────────────────
  var state = {
    list: [],              // [{id, type, title, desc, time, href, unread}]
    lastReadAt: '1970-01-01T00:00:00Z',
    pollTimer: null,
    refreshing: false,
    dropdownOpen: false,
    activeRole: null,
    sessionId: null,
    sessionNIP: null,      // jika user.nip ada (utk filter PAK milik user)
  };

  // ─── Utility ─────────────────────────────────────────────────────
  function escHTML(s) {
    return String(s == null ? '' : s).replace(/[&<>"']/g, function (c) {
      return { '&':'&amp;', '<':'&lt;', '>':'&gt;', '"':'&quot;', "'":'&#39;' }[c];
    });
  }

  /** Format timestamp ke "X menit lalu" / "X jam lalu" / "DD Bulan" */
  function fmtRelativeTime(iso) {
    if (!iso) return '—';
    var d = new Date(iso);
    if (isNaN(d)) return '—';
    var now = new Date();
    var diffSec = Math.floor((now - d) / 1000);
    if (diffSec < 60) return 'Baru saja';
    if (diffSec < 3600) return Math.floor(diffSec / 60) + ' menit lalu';
    if (diffSec < 86400) return Math.floor(diffSec / 3600) + ' jam lalu';
    if (diffSec < 86400 * 7) return Math.floor(diffSec / 86400) + ' hari lalu';
    var bulan = ['Jan','Feb','Mar','Apr','Mei','Jun','Jul','Agu','Sep','Okt','Nov','Des'];
    var s = d.getDate() + ' ' + bulan[d.getMonth()];
    if (d.getFullYear() !== now.getFullYear()) s += ' ' + d.getFullYear();
    return s;
  }

  function getSession() {
    try { return JSON.parse(localStorage.getItem('nova_user') || 'null'); }
    catch (_) { return null; }
  }

  // ─── Fetch ───────────────────────────────────────────────────────

  /**
   * Fetch last_read_at dari users table.
   */
  function fetchLastReadAt(sessionId) {
    var url = window.SUPABASE_URL + '/rest/v1/users?id=eq.' + encodeURIComponent(sessionId)
            + '&select=notifikasi_last_read_at&limit=1';
    return fetch(url, { headers: window.SUPABASE_HEADERS })
      .then(function (r) {
        if (!r.ok) throw new Error('HTTP ' + r.status);
        return r.json();
      })
      .then(function (rows) {
        if (rows && rows[0] && rows[0].notifikasi_last_read_at) {
          return rows[0].notifikasi_last_read_at;
        }
        return '1970-01-01T00:00:00Z';
      })
      .catch(function (e) {
        console.warn('[notif] fetchLastReadAt error:', e);
        return '1970-01-01T00:00:00Z';
      });
  }

  /**
   * Fetch notifikasi untuk USER (role 'user').
   * - PAK miliknya yang sudah selesai
   * - Surat tugas miliknya yang sudah selesai
   *
   * Filter PAK pakai pegawai_nip = session.nip (dari users.nip atau
   * users.username, tergantung struktur DB). Frontend coba kedua dulu.
   */
  function fetchUserNotifs() {
    var session = getSession();
    if (!session) return Promise.resolve([]);

    // Filter NIP user — coba beberapa kemungkinan field di session
    var nip = session.nip || session.NIP || session.username || null;

    var promises = [];

    // PAK selesai untuk pegawai ini
    if (nip) {
      var pakUrl = window.SUPABASE_URL + '/rest/v1/pengajuan_pak'
                 + '?pegawai_nip=eq.' + encodeURIComponent(nip)
                 + '&status=eq.selesai'
                 + '&select=id,nomor_urut,tahun_periode,tgl_pengajuan,ak_total,updated_at,created_at'
                 + '&order=updated_at.desc.nullslast,created_at.desc&limit=20';
      promises.push(
        fetch(pakUrl, { headers: window.SUPABASE_HEADERS })
          .then(function (r) { return r.ok ? r.json() : []; })
          .catch(function () { return []; })
      );
    } else {
      promises.push(Promise.resolve([]));
    }

    // Surat tugas selesai untuk user_id ini
    var stUrl = window.SUPABASE_URL + '/rest/v1/surat_tugas'
              + '?user_id=eq.' + encodeURIComponent(session.id)
              + '&status=eq.selesai'
              + '&select=id,nomor_surat,tipe,perihal,updated_at,created_at'
              + '&order=updated_at.desc.nullslast,created_at.desc&limit=20';
    promises.push(
      fetch(stUrl, { headers: window.SUPABASE_HEADERS })
        .then(function (r) { return r.ok ? r.json() : []; })
        .catch(function () { return []; })
    );

    return Promise.all(promises).then(function (results) {
      var pakRows = results[0] || [];
      var stRows  = results[1] || [];
      var notifs = [];

      pakRows.forEach(function (p) {
        var ts = p.updated_at || p.created_at;
        notifs.push({
          id: 'pak-done-' + p.id,
          type: 'pak-done',
          icon: '✓',
          iconBg: '#d1fae5',
          title: 'Pengajuan PAK Disetujui',
          desc: 'Pengajuan No. ' + String(p.nomor_urut).padStart(3,'0') + '/'
              + p.tahun_periode + ' telah disetujui. AK total: '
              + (p.ak_total || '—'),
          time: ts,
          href: 'index.html',
        });
      });

      stRows.forEach(function (s) {
        var ts = s.updated_at || s.created_at;
        notifs.push({
          id: 'st-done-' + s.id,
          type: 'st-done',
          icon: '📄',
          iconBg: '#dbeafe',
          title: 'Surat Tugas Selesai',
          desc: (s.nomor_surat || 'Surat Tugas') + (s.perihal ? ' — ' + s.perihal.slice(0, 60) : ''),
          time: ts,
          href: 'surat-tugas.html',
        });
      });

      return notifs;
    });
  }

  /**
   * Fetch notifikasi untuk ADMIN.
   * - PAK menunggu approval (semua user)
   * - Surat tugas menunggu approval (semua user)
   */
  function fetchAdminNotifs() {
    var pakUrl = window.SUPABASE_URL + '/rest/v1/pengajuan_pak'
               + '?status=eq.menunggu'
               + '&select=id,nomor_urut,tahun_periode,pegawai_nip,penandatangan_nama,ak_total,created_at'
               + '&order=created_at.desc&limit=20';
    var stUrl = window.SUPABASE_URL + '/rest/v1/surat_tugas'
              + '?status=eq.menunggu'
              + '&select=id,nomor_surat,tipe,perihal,user_id,created_at'
              + '&order=created_at.desc&limit=20';

    return Promise.all([
      fetch(pakUrl, { headers: window.SUPABASE_HEADERS }).then(function (r) { return r.ok ? r.json() : []; }).catch(function () { return []; }),
      fetch(stUrl,  { headers: window.SUPABASE_HEADERS }).then(function (r) { return r.ok ? r.json() : []; }).catch(function () { return []; }),
    ]).then(function (results) {
      var pakRows = results[0] || [];
      var stRows  = results[1] || [];
      var notifs = [];

      pakRows.forEach(function (p) {
        notifs.push({
          id: 'pak-pending-' + p.id,
          type: 'pak-pending',
          icon: '⭐',
          iconBg: '#fef3c7',
          title: 'Pengajuan PAK Baru',
          desc: 'NIP ' + (p.pegawai_nip || '—') + ' mengajukan PAK No. '
              + String(p.nomor_urut).padStart(3,'0') + '/' + p.tahun_periode
              + ' (AK total: ' + (p.ak_total || '—') + ')',
          time: p.created_at,
          href: 'admin-kepegawaian.html?nip=' + encodeURIComponent(p.pegawai_nip || ''),
        });
      });

      stRows.forEach(function (s) {
        notifs.push({
          id: 'st-pending-' + s.id,
          type: 'st-pending',
          icon: '📄',
          iconBg: '#dbeafe',
          title: 'Pengajuan Surat Tugas',
          desc: (s.nomor_surat ? 'No. ' + s.nomor_surat + ' — ' : '')
              + (s.perihal ? s.perihal.slice(0, 80) : 'Menunggu persetujuan'),
          time: s.created_at,
          href: 'admin-surat-tugas.html',
        });
      });

      return notifs;
    });
  }

  // ─── Render ──────────────────────────────────────────────────────

  function unreadCount() {
    if (!state.list.length) return 0;
    var lastRead = new Date(state.lastReadAt).getTime();
    return state.list.filter(function (n) {
      var t = new Date(n.time).getTime();
      return !isNaN(t) && t > lastRead;
    }).length;
  }

  function updateBadge() {
    var btn = document.getElementById('notif-btn');
    var badge = document.getElementById('notif-badge');
    if (!btn || !badge) return;
    var count = unreadCount();
    if (count > 0) {
      badge.textContent = count > 99 ? '99+' : String(count);
      badge.hidden = false;
      btn.classList.add('has-unread');
    } else {
      badge.hidden = true;
      btn.classList.remove('has-unread');
    }
  }

  function renderList() {
    var listEl = document.getElementById('notif-list');
    var metaEl = document.getElementById('notif-header-meta');
    if (!listEl) return;

    var count = unreadCount();
    if (metaEl) {
      metaEl.textContent = count > 0 ? (count + ' belum dibaca') : 'Tidak ada baru';
    }

    if (!state.list.length) {
      listEl.innerHTML = ''
        + '<div class="notif-empty">'
        +   '<div class="notif-empty-icon">🔕</div>'
        +   'Belum ada notifikasi.'
        + '</div>';
      return;
    }

    // Sort terbaru dulu
    var sorted = state.list.slice().sort(function (a, b) {
      return String(b.time || '').localeCompare(String(a.time || ''));
    });
    var lastRead = new Date(state.lastReadAt).getTime();

    listEl.innerHTML = sorted.map(function (n) {
      var t = new Date(n.time).getTime();
      var isUnread = !isNaN(t) && t > lastRead;
      var cls = 'notif-item' + (isUnread ? ' unread' : '');
      return ''
        + '<a class="' + cls + '" href="' + escHTML(n.href || '#') + '">'
        +   '<div class="notif-item-icon" style="background:' + (n.iconBg || '#f5f4f0') + '">'
        +     escHTML(n.icon || '🔔')
        +   '</div>'
        +   '<div class="notif-item-body">'
        +     '<div class="notif-item-title">' + escHTML(n.title) + '</div>'
        +     '<div class="notif-item-desc">' + escHTML(n.desc) + '</div>'
        +     '<div class="notif-item-time">' + escHTML(fmtRelativeTime(n.time)) + '</div>'
        +   '</div>'
        +   '<div class="notif-item-dot" aria-hidden="true"></div>'
        + '</a>';
    }).join('');
  }

  // ─── Refresh + open/close ────────────────────────────────────────

  function refresh() {
    if (state.refreshing) return;
    state.refreshing = true;

    var session = getSession();
    if (!session || !session.id) {
      state.refreshing = false;
      return;
    }

    state.sessionId = session.id;
    state.activeRole = session.active_role || 'user';

    Promise.all([
      fetchLastReadAt(session.id),
      state.activeRole === 'admin' ? fetchAdminNotifs() : fetchUserNotifs(),
    ]).then(function (results) {
      state.lastReadAt = results[0];
      state.list = results[1] || [];
      updateBadge();
      // Re-render list kalau dropdown terbuka
      if (state.dropdownOpen) renderList();
    }).catch(function (e) {
      console.warn('[notif] refresh error:', e);
    }).then(function () {
      state.refreshing = false;
    });
  }

  function openDropdown() {
    var dd = document.getElementById('notif-dropdown');
    var btn = document.getElementById('notif-btn');
    if (!dd || !btn) return;
    state.dropdownOpen = true;
    dd.classList.add('open');
    dd.setAttribute('aria-hidden', 'false');
    btn.setAttribute('aria-expanded', 'true');
    renderList();
    // Mark as read on open (kalau ada unread)
    if (unreadCount() > 0) {
      markAllRead();
    }
  }

  function closeDropdown() {
    var dd = document.getElementById('notif-dropdown');
    var btn = document.getElementById('notif-btn');
    if (!dd || !btn) return;
    state.dropdownOpen = false;
    dd.classList.remove('open');
    dd.setAttribute('aria-hidden', 'true');
    btn.setAttribute('aria-expanded', 'false');
  }

  function toggleDropdown(e) {
    if (e) { e.preventDefault(); e.stopPropagation(); }
    if (state.dropdownOpen) closeDropdown(); else openDropdown();
  }

  /**
   * Call RPC mark_notifikasi_read. Optimistic: badge langsung di-clear,
   * baru push update ke server.
   */
  function markAllRead() {
    var session = getSession();
    if (!session || !session.id) return;

    // Optimistic: set lastReadAt ke now lokal
    state.lastReadAt = new Date().toISOString();
    updateBadge();
    renderList();

    // Push ke server
    fetch(window.SUPABASE_URL + '/rest/v1/rpc/mark_notifikasi_read', {
      method: 'POST',
      headers: Object.assign({}, window.SUPABASE_HEADERS, { 'Content-Type': 'application/json' }),
      body: JSON.stringify({ p_caller_id: session.id }),
    }).then(function (r) {
      if (!r.ok) {
        return r.text().then(function (t) {
          console.warn('[notif] markAllRead RPC failed:', r.status, t);
        });
      }
    }).catch(function (e) {
      console.warn('[notif] markAllRead network error:', e);
    });
  }

  // ─── Init ────────────────────────────────────────────────────────

  function attachHandlers() {
    var btn = document.getElementById('notif-btn');
    if (!btn) return false;

    btn.addEventListener('click', toggleDropdown);

    // Close kalau klik luar dropdown
    document.addEventListener('click', function (e) {
      if (!state.dropdownOpen) return;
      var wrap = document.getElementById('notif-wrap');
      if (wrap && !wrap.contains(e.target)) closeDropdown();
    });

    // ESC untuk close
    document.addEventListener('keydown', function (e) {
      if (e.key === 'Escape' && state.dropdownOpen) closeDropdown();
    });

    return true;
  }

  function startPolling() {
    if (state.pollTimer) clearInterval(state.pollTimer);
    state.pollTimer = setInterval(refresh, 60 * 1000); // 60 detik
  }

  function init() {
    // Pastikan SUPABASE_URL/HEADERS tersedia (dari config.js)
    if (!window.SUPABASE_URL || !window.SUPABASE_HEADERS) {
      console.warn('[notif] SUPABASE_URL/HEADERS belum ter-load. Notifikasi disabled.');
      return;
    }
    if (!attachHandlers()) {
      console.warn('[notif] notif-btn element tidak ditemukan. Notifikasi disabled.');
      return;
    }

    // Initial fetch
    refresh();
    startPolling();

    // Re-fetch saat tab kembali fokus (catch missed updates)
    document.addEventListener('visibilitychange', function () {
      if (!document.hidden) refresh();
    });
  }

  // ─── Bootstrap ───────────────────────────────────────────────────
  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', init);
  } else {
    // Sudah loaded → init lewat microtask supaya topbar.js sempat
    // render dulu kalau di-load di order berbeda.
    setTimeout(init, 0);
  }

  // Expose minimal API untuk debugging dari console
  window.NotifikasiPortal = {
    refresh: refresh,
    open: openDropdown,
    close: closeDropdown,
    markAllRead: markAllRead,
    _state: state,
  };
})();
