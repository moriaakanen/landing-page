/**
 * 9201 NOTIFIKASI — engine untuk lonceng 🔔 di topbar (REWRITE 2026-05)
 * ──────────────────────────────────────────────────────────────────
 * SELF-CONTAINED: file ini sekarang menangani SEMUA aspect notifikasi:
 *   - CSS dropdown (di-inject ke <head>)
 *   - HTML dropdown (di-create dan di-append ke <body>, position:fixed)
 *   - Click handler (event delegation di document, anti race-condition)
 *   - Positioning dropdown via JS (relative to button rect)
 *   - Tabs Facebook-style: "Semua" | "Belum Dibaca"
 *   - Empty state card (selalu muncul kalau tidak ada notif)
 *   - Polling refresh tiap 60 detik
 *   - Mark-as-read on open
 *
 * KENAPA REWRITE:
 *   - Versi sebelumnya: dropdown sebagai child notif-wrap di topbar.
 *     Stacking context dari .topbar (z-index:200, position:sticky)
 *     dan body (overflow-x:hidden) bisa bikin dropdown invisible
 *     di kondisi tertentu.
 *   - Versi sebelumnya: click handler attached langsung ke btn element.
 *     Kalau ada race condition (topbar render telat), handler nempel
 *     ke element yg salah / tidak nempel sama sekali.
 *   - Solusi: dropdown jadi body-level child (escape semua stacking),
 *     event delegation di document (anti race), position via JS.
 *
 * Sumber notifikasi (tidak berubah dari versi sebelumnya):
 *   USER (active_role='user'):
 *     - Pengajuan PAK miliknya yang status 'selesai' (di-approve)
 *     - Surat tugas miliknya yang status 'selesai'
 *   ADMIN (active_role='admin'):
 *     - Pengajuan PAK status 'menunggu' (semua user)
 *     - Surat tugas status 'menunggu' (semua user)
 *
 * Polling: refresh tiap 60 detik.
 */
(function () {
  'use strict';

  if (window.__notif_init__) return;  // guard double-init
  window.__notif_init__ = true;

  // ═══════════════════════════════════════════════════════════════════
  // STATE
  // ═══════════════════════════════════════════════════════════════════
  var state = {
    list: [],              // [{id, type, title, desc, time, href, icon, iconCls}]
    lastReadAt: '1970-01-01T00:00:00Z',
    pollTimer: null,
    refreshing: false,
    dropdownOpen: false,
    activeRole: null,
    sessionId: null,
    sessionNIP: null,
    activeTab: 'all',      // 'all' | 'unread'
    initialFetched: false, // flag: sudah pernah fetch sukses?
  };

  // ═══════════════════════════════════════════════════════════════════
  // CSS — Facebook-inspired, self-contained
  // ═══════════════════════════════════════════════════════════════════
  var CSS = ''
    + '#notif-dropdown-fb{position:fixed;top:60px;right:24px;width:380px;max-width:calc(100vw - 24px);'
    + 'background:#fff;border-radius:14px;box-shadow:0 12px 40px rgba(13,35,64,.18),0 2px 6px rgba(13,35,64,.08);'
    + 'border:1px solid rgba(13,35,64,.06);z-index:9999;display:none;overflow:hidden;'
    + 'font-family:\'Plus Jakarta Sans\',system-ui,sans-serif;color:#0d2340;'
    + 'transform-origin:top right}'
    + '#notif-dropdown-fb.open{display:flex !important;flex-direction:column;'
    + 'animation:notifFbPop .18s cubic-bezier(.2,.7,.3,1) both}'
    + '@keyframes notifFbPop{from{opacity:0;transform:translateY(-8px) scale(.97)}'
    + 'to{opacity:1;transform:translateY(0) scale(1)}}'
    // Header
    + '.notif-fb-head{padding:14px 16px 6px;display:flex;align-items:center;justify-content:space-between;flex-shrink:0}'
    + '.notif-fb-title{font-family:\'Fraunces\',\'Plus Jakarta Sans\',serif;font-size:20px;font-weight:600;color:#0d2340;letter-spacing:-.3px;line-height:1.2;margin:0}'
    + '.notif-fb-actions{display:flex;align-items:center;gap:4px}'
    + '.notif-fb-iconbtn{width:30px;height:30px;border-radius:50%;background:#f0eee8;border:none;cursor:pointer;'
    + 'display:flex;align-items:center;justify-content:center;font-size:14px;color:#0d2340;'
    + 'transition:background .12s;font-family:inherit;line-height:1}'
    + '.notif-fb-iconbtn:hover{background:#e2ddd6}'
    + '.notif-fb-iconbtn[disabled]{opacity:.4;cursor:not-allowed}'
    // Tabs
    + '.notif-fb-tabs{display:flex;gap:6px;padding:6px 16px 12px;flex-shrink:0}'
    + '.notif-fb-tab{padding:7px 14px;border-radius:100px;background:#f0eee8;border:none;cursor:pointer;'
    + 'font-family:inherit;font-size:13px;font-weight:600;color:#0d2340;transition:background .12s,color .12s;'
    + 'display:inline-flex;align-items:center;gap:5px;line-height:1.2}'
    + '.notif-fb-tab:hover{background:#e2ddd6}'
    + '.notif-fb-tab.active{background:#fcf3d9;color:#7a5c10}'
    + '.notif-fb-tab-count{display:inline-flex;align-items:center;justify-content:center;'
    + 'min-width:18px;height:18px;padding:0 5px;background:#fff;border-radius:100px;'
    + 'font-size:10.5px;font-weight:700;color:#7a5c10;font-variant-numeric:tabular-nums}'
    + '.notif-fb-tab:not(.active) .notif-fb-tab-count{background:#fff;color:#6b7280}'
    // List
    + '.notif-fb-list{flex:1;overflow-y:auto;background:#fff;max-height:520px;min-height:120px}'
    + '.notif-fb-section{font-size:14px;font-weight:700;color:#0d2340;padding:10px 16px 6px;letter-spacing:-.1px}'
    + '.notif-fb-section:first-child{padding-top:4px}'
    // Items
    + '.notif-fb-item{display:flex;gap:12px;padding:10px 16px;cursor:pointer;'
    + 'transition:background .12s;text-decoration:none;color:inherit;align-items:flex-start;'
    + 'border-radius:8px;margin:0 6px;position:relative}'
    + '.notif-fb-item:hover{background:#f0eee8}'
    + '.notif-fb-item-icon{flex-shrink:0;width:42px;height:42px;border-radius:50%;'
    + 'display:flex;align-items:center;justify-content:center;font-size:18px;'
    + 'background:#f5edda;color:#7a5c10;border:1px solid rgba(200,168,75,.25)}'
    + '.notif-fb-item-icon.is-success{background:#d1fae5;color:#065f46;border-color:#a7f3d0}'
    + '.notif-fb-item-icon.is-info{background:#dbeafe;color:#1e40af;border-color:#bfdbfe}'
    + '.notif-fb-item-icon.is-warn{background:#fef3c7;color:#92400e;border-color:#fde68a}'
    + '.notif-fb-item-body{flex:1;min-width:0;padding-top:2px}'
    + '.notif-fb-item-title{font-size:13px;font-weight:500;color:#0d2340;line-height:1.4;margin-bottom:2px;word-wrap:break-word}'
    + '.notif-fb-item-title strong{font-weight:700}'
    + '.notif-fb-item-time{font-size:12px;color:#c8a84b;font-weight:600;font-variant-numeric:tabular-nums;letter-spacing:.1px}'
    + '.notif-fb-item-time.is-old{color:#6b7280}'
    + '.notif-fb-item-dot{width:10px;height:10px;border-radius:50%;background:#3b82f6;'
    + 'flex-shrink:0;margin-top:18px;display:none}'
    + '.notif-fb-item.unread .notif-fb-item-dot{display:block}'
    + '.notif-fb-item.unread .notif-fb-item-title{color:#0d2340;font-weight:600}'
    // Empty state
    + '.notif-fb-empty{display:flex;flex-direction:column;align-items:center;justify-content:center;'
    + 'padding:48px 24px 56px;text-align:center;color:#6b7280;gap:10px;background:#fff}'
    + '.notif-fb-empty-icon{width:64px;height:64px;border-radius:50%;background:#f0eee8;'
    + 'display:flex;align-items:center;justify-content:center;font-size:28px;opacity:.7}'
    + '.notif-fb-empty-title{font-size:14px;font-weight:600;color:#0d2340;margin-top:4px}'
    + '.notif-fb-empty-desc{font-size:12.5px;color:#6b7280;line-height:1.5;max-width:240px}'
    // Loading state
    + '.notif-fb-loading{display:flex;flex-direction:column;align-items:center;justify-content:center;'
    + 'padding:48px 24px;gap:12px;color:#6b7280;font-size:13px}'
    + '.notif-fb-loading-spin{width:28px;height:28px;border:2.5px solid #e2ddd6;border-top-color:#0d2340;'
    + 'border-radius:50%;animation:notifFbSpin .8s linear infinite}'
    + '@keyframes notifFbSpin{to{transform:rotate(360deg)}}'
    // Mobile: full-width modal-like
    + '@media(max-width:480px){'
    + '#notif-dropdown-fb{width:auto;left:8px;right:8px;top:58px;max-width:none}'
    + '.notif-fb-list{max-height:60vh}'
    + '}';

  function injectCSS() {
    if (document.getElementById('notif-fb-css')) return;
    var s = document.createElement('style');
    s.id = 'notif-fb-css';
    s.textContent = CSS;
    document.head.appendChild(s);
  }

  // ═══════════════════════════════════════════════════════════════════
  // DROPDOWN HTML — created lazily, appended to <body>
  // ═══════════════════════════════════════════════════════════════════
  function ensureDropdownEl() {
    var dd = document.getElementById('notif-dropdown-fb');
    if (dd) return dd;
    if (!document.body) return null; // body belum ada (terlalu awal)

    dd = document.createElement('div');
    dd.id = 'notif-dropdown-fb';
    dd.setAttribute('role', 'menu');
    dd.setAttribute('aria-hidden', 'true');
    dd.innerHTML = ''
      + '<div class="notif-fb-head">'
      +   '<h3 class="notif-fb-title">Notifikasi</h3>'
      +   '<div class="notif-fb-actions">'
      +     '<button class="notif-fb-iconbtn" id="notif-fb-mark-all" type="button" '
      +             'title="Tandai semua sebagai sudah dibaca" aria-label="Tandai semua dibaca">✓</button>'
      +   '</div>'
      + '</div>'
      + '<div class="notif-fb-tabs">'
      +   '<button class="notif-fb-tab active" data-tab="all" type="button">Semua</button>'
      +   '<button class="notif-fb-tab" data-tab="unread" type="button">'
      +     'Belum Dibaca <span class="notif-fb-tab-count" id="notif-fb-unread-count" style="display:none">0</span>'
      +   '</button>'
      + '</div>'
      + '<div class="notif-fb-list" id="notif-fb-list">'
      +   '<div class="notif-fb-loading"><div class="notif-fb-loading-spin"></div><div>Memuat notifikasi...</div></div>'
      + '</div>';

    document.body.appendChild(dd);
    return dd;
  }

  // ═══════════════════════════════════════════════════════════════════
  // UTILITIES
  // ═══════════════════════════════════════════════════════════════════
  function escHTML(s) {
    return String(s == null ? '' : s).replace(/[&<>"']/g, function (c) {
      return { '&':'&amp;', '<':'&lt;', '>':'&gt;', '"':'&quot;', "'":'&#39;' }[c];
    });
  }

  /** Format timestamp ke Facebook-style relative ("8 jam", "Kemarin", dst.) */
  function fmtRelativeTime(iso) {
    if (!iso) return '—';
    var d = new Date(iso);
    if (isNaN(d)) return '—';
    var now = new Date();
    var diffSec = Math.floor((now - d) / 1000);
    if (diffSec < 60) return 'Baru saja';
    if (diffSec < 3600) return Math.floor(diffSec / 60) + ' menit';
    if (diffSec < 86400) return Math.floor(diffSec / 3600) + ' jam';
    if (diffSec < 86400 * 2) return 'Kemarin';
    if (diffSec < 86400 * 7) return Math.floor(diffSec / 86400) + ' hari';
    var bulan = ['Jan','Feb','Mar','Apr','Mei','Jun','Jul','Agu','Sep','Okt','Nov','Des'];
    var s = d.getDate() + ' ' + bulan[d.getMonth()];
    if (d.getFullYear() !== now.getFullYear()) s += ' ' + d.getFullYear();
    return s;
  }

  /** Group notif by section: "Hari ini", "Kemarin", "Minggu ini", "Lebih lama" */
  function groupBySection(items) {
    var now = new Date();
    var startOfToday = new Date(now.getFullYear(), now.getMonth(), now.getDate()).getTime();
    var startOfYesterday = startOfToday - 86400000;
    var startOfWeek = startOfToday - 6 * 86400000;
    var groups = { today: [], yesterday: [], thisWeek: [], older: [] };
    items.forEach(function (n) {
      var t = new Date(n.time).getTime();
      if (isNaN(t)) { groups.older.push(n); return; }
      if (t >= startOfToday)        groups.today.push(n);
      else if (t >= startOfYesterday) groups.yesterday.push(n);
      else if (t >= startOfWeek)    groups.thisWeek.push(n);
      else                          groups.older.push(n);
    });
    return groups;
  }

  function getSession() {
    try { return JSON.parse(localStorage.getItem('nova_user') || 'null'); }
    catch (_) { return null; }
  }

  // ═══════════════════════════════════════════════════════════════════
  // LOOKUP CACHES — pegawai (NIP→NAMA) & users (id→full_name)
  // ─────────────────────────────────────────────────────────────────
  // Admin notif tampilkan NAMA (bukan NIP atau user_id) — admin tidak
  // hapal NIP 18-digit / user_id internal. Kita fetch dua tabel sekali
  // dan cache mapping selama 5 menit:
  //
  //   - data_pegawai (NIP → NAMA)    → untuk pengajuan PAK
  //   - users        (id → full_name) → untuk surat tugas
  //
  // Cache di-refresh otomatis kalau stale (>5 menit). `pendingPromise`
  // dedupe concurrent fetches (kalau bell di-buka berulang cepat, tidak
  // fire fetch berkali-kali). Fetch dijalankan parallel — total cost
  // sama dengan satu request terlama, bukan dua sequential.
  // ═══════════════════════════════════════════════════════════════════
  var _lookupCache = {
    pegawaiByNIP: null,   // { '199903302019121001': 'Joko Susilo', ... }
    usersByID:    null,   // { 5: 'Jane Doe', 7: 'Budi Santoso', ... }
    loadedAt: 0,
    pendingPromise: null
  };
  var LOOKUP_CACHE_TTL_MS = 5 * 60 * 1000;

  function fetchLookups() {
    var url = resolveSupabaseUrl();
    if (!url) return Promise.resolve({ pegawaiByNIP: {}, usersByID: {} });

    if (_lookupCache.pendingPromise) return _lookupCache.pendingPromise;

    var p = Promise.all([
      fetch(url + '/rest/v1/data_pegawai?select=pegawai_nip,nama',
            { headers: window.SUPABASE_HEADERS })
        .then(function (r) { return r.ok ? r.json() : []; })
        .catch(function () { return []; }),
      fetch(url + '/rest/v1/users?select=id,full_name,username',
            { headers: window.SUPABASE_HEADERS })
        .then(function (r) { return r.ok ? r.json() : []; })
        .catch(function () { return []; }),
    ]).then(function (results) {
      var pegawaiByNIP = {};
      (results[0] || []).forEach(function (row) {
        var nip = window.pegawaiNip ? window.pegawaiNip(row) : (row && (row.pegawai_nip || row.NIP));
        if (row && nip) {
          pegawaiByNIP[String(nip)] = (window.pegawaiNama ? window.pegawaiNama(row) : (row.nama || row.NAMA)) || String(nip);
        }
      });
      var usersByID = {};
      (results[1] || []).forEach(function (u) {
        if (u && u.id != null) {
          // Fallback chain: full_name → cek pegawaiByNIP[username] → username → "User #id"
          // Ini handle kasus where users.full_name belum terisi tapi data_pegawai.NAMA
          // ada (mis. user lama yg tidak pernah update profile).
          var nama = u.full_name;
          if (!nama && u.username && pegawaiByNIP[String(u.username)]) {
            nama = pegawaiByNIP[String(u.username)];
          }
          if (!nama) nama = u.username || ('User #' + u.id);
          usersByID[u.id] = nama;
        }
      });
      _lookupCache.pegawaiByNIP = pegawaiByNIP;
      _lookupCache.usersByID    = usersByID;
      _lookupCache.loadedAt     = Date.now();
      _lookupCache.pendingPromise = null;
      return _lookupCache;
    }).catch(function (e) {
      console.warn('[notif] fetchLookups error:', e);
      _lookupCache.pendingPromise = null;
      return _lookupCache;
    });

    _lookupCache.pendingPromise = p;
    return p;
  }

  /** Pastikan cache fresh (<TTL). Re-fetch kalau stale/empty. */
  function ensureLookups() {
    var hasData = _lookupCache.pegawaiByNIP || _lookupCache.usersByID;
    var stale   = (Date.now() - _lookupCache.loadedAt) > LOOKUP_CACHE_TTL_MS;
    if (hasData && !stale) return Promise.resolve(_lookupCache);
    return fetchLookups();
  }

  /** Sync lookup: NIP → NAMA. Fallback ke NIP kalau tidak ditemukan. */
  function getPegawaiName(nip) {
    if (!nip) return '—';
    var key = String(nip);
    if (_lookupCache.pegawaiByNIP && _lookupCache.pegawaiByNIP[key]) {
      return _lookupCache.pegawaiByNIP[key];
    }
    return key; // fallback: tampilkan NIP daripada kosong
  }

  /** Sync lookup: user_id → full_name. Fallback ke "User #id" kalau tidak ada. */
  function getUserName(userId) {
    if (userId == null) return '—';
    if (_lookupCache.usersByID && _lookupCache.usersByID[userId]) {
      return _lookupCache.usersByID[userId];
    }
    return 'User #' + userId; // fallback informatif
  }

  // ═══════════════════════════════════════════════════════════════════
  // FETCH (sumber: pengajuan_pak + surat_tugas)
  // ═══════════════════════════════════════════════════════════════════

  function fetchLastReadAt(sessionId) {
    var url = resolveSupabaseUrl();
    if (!url) return Promise.resolve('1970-01-01T00:00:00Z');
    var u = url + '/rest/v1/users?id=eq.' + encodeURIComponent(sessionId)
          + '&select=notifikasi_last_read_at&limit=1';
    return fetch(u, { headers: window.SUPABASE_HEADERS })
      .then(function (r) { return r.ok ? r.json() : []; })
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

  function fetchUserNotifs() {
    var url = resolveSupabaseUrl();
    if (!url) return Promise.resolve([]);
    var session = getSession();
    if (!session) return Promise.resolve([]);

    var nip = session.username || null;
    var promises = [];

    if (nip) {
      var pakUrl = url + '/rest/v1/pengajuan_pak'
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

    var stUrl = url + '/rest/v1/surat_tugas'
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
          iconCls: 'is-success',
          title: '<strong>Pengajuan PAK No. ' + String(p.nomor_urut).padStart(3,'0') + '/'
               + p.tahun_periode + '</strong> telah disetujui. AK total: '
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
          iconCls: 'is-info',
          title: '<strong>Surat Tugas ' + escHTML(s.nomor_surat || '') + '</strong>'
               + (s.perihal ? ' — ' + escHTML(s.perihal.slice(0, 80)) : '') + ' telah selesai diproses.',
          time: ts,
          href: 'surat-tugas.html',
        });
      });

      return notifs;
    });
  }

  function fetchAdminNotifs() {
    var url = resolveSupabaseUrl();
    if (!url) return Promise.resolve([]);

    // Pre-fetch lookup caches (pegawai + users) supaya bisa lookup NAMA
    // saat build notif title. Kalau cache fail, tetap lanjut — getter
    // helpers akan fallback ke NIP / "User #id".
    return ensureLookups().then(function () {
      var pakUrl = url + '/rest/v1/pengajuan_pak'
                 + '?status=eq.menunggu'
                 + '&select=id,nomor_urut,tahun_periode,pegawai_nip,penandatangan_nama,ak_total,created_at'
                 + '&order=created_at.desc&limit=20';
      var stUrl = url + '/rest/v1/surat_tugas'
                + '?status=eq.menunggu'
                + '&select=id,nomor_surat,tipe,perihal,user_id,created_at'
                + '&order=created_at.desc&limit=20';

      return Promise.all([
        fetch(pakUrl, { headers: window.SUPABASE_HEADERS }).then(function (r) { return r.ok ? r.json() : []; }).catch(function () { return []; }),
        fetch(stUrl,  { headers: window.SUPABASE_HEADERS }).then(function (r) { return r.ok ? r.json() : []; }).catch(function () { return []; }),
      ]);
    }).then(function (results) {
      var pakRows = results[0] || [];
      var stRows  = results[1] || [];
      var notifs = [];

      pakRows.forEach(function (p) {
        var nama = getPegawaiName(p.pegawai_nip);
        notifs.push({
          id: 'pak-pending-' + p.id,
          type: 'pak-pending',
          icon: '⭐',
          iconCls: 'is-warn',
          title: '<strong>' + escHTML(nama) + '</strong> mengajukan PAK '
               + 'No. ' + String(p.nomor_urut).padStart(3,'0') + '/' + p.tahun_periode
               + ' (AK: ' + (p.ak_total || '—') + ').',
          time: p.created_at,
          href: 'admin-kepegawaian.html?nip=' + encodeURIComponent(p.pegawai_nip || ''),
        });
      });

      stRows.forEach(function (s) {
        var nama = getUserName(s.user_id);
        notifs.push({
          id: 'st-pending-' + s.id,
          type: 'st-pending',
          icon: '📄',
          iconCls: 'is-info',
          title: '<strong>' + escHTML(nama) + '</strong> mengajukan Surat Tugas baru'
               + (s.nomor_surat ? ' (No. ' + escHTML(s.nomor_surat) + ')' : '')
               + (s.perihal ? ' — ' + escHTML(s.perihal.slice(0, 80)) : ''),
          time: s.created_at,
          href: 'admin-surat-tugas.html',
        });
      });

      return notifs;
    });
  }

  // ═══════════════════════════════════════════════════════════════════
  // RENDER
  // ═══════════════════════════════════════════════════════════════════

  function unreadCount() {
    if (!state.list.length) return 0;
    var lastRead = new Date(state.lastReadAt).getTime();
    return state.list.filter(function (n) {
      var t = new Date(n.time).getTime();
      return !isNaN(t) && t > lastRead;
    }).length;
  }

  function isUnread(n, lastReadMs) {
    var t = new Date(n.time).getTime();
    return !isNaN(t) && t > lastReadMs;
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

  function renderEmptyState(tab) {
    var icon, title, desc;
    if (tab === 'unread') {
      icon = '✓';
      title = 'Semua sudah dibaca';
      desc = 'Tidak ada notifikasi yang belum dibaca saat ini.';
    } else {
      icon = '🔔';
      title = 'Belum ada notifikasi';
      desc = state.activeRole === 'admin'
        ? 'Notifikasi akan muncul di sini saat ada pengajuan baru yang menunggu persetujuan.'
        : 'Notifikasi akan muncul di sini saat pengajuan Anda diproses.';
    }
    return ''
      + '<div class="notif-fb-empty">'
      +   '<div class="notif-fb-empty-icon">' + icon + '</div>'
      +   '<div class="notif-fb-empty-title">' + escHTML(title) + '</div>'
      +   '<div class="notif-fb-empty-desc">' + escHTML(desc) + '</div>'
      + '</div>';
  }

  function renderItem(n, lastReadMs) {
    var unread = isUnread(n, lastReadMs);
    var cls = 'notif-fb-item' + (unread ? ' unread' : '');
    var iconCls = n.iconCls || '';
    var timeCls = unread ? '' : ' is-old';
    return ''
      + '<a class="' + cls + '" href="' + escHTML(n.href || '#') + '" data-id="' + escHTML(n.id) + '">'
      +   '<div class="notif-fb-item-icon ' + iconCls + '">' + (n.icon || '🔔') + '</div>'
      +   '<div class="notif-fb-item-body">'
      +     '<div class="notif-fb-item-title">' + (n.title || '') + '</div>'
      +     '<div class="notif-fb-item-time' + timeCls + '">' + escHTML(fmtRelativeTime(n.time)) + '</div>'
      +   '</div>'
      +   '<div class="notif-fb-item-dot" aria-hidden="true"></div>'
      + '</a>';
  }

  function renderList() {
    var listEl = document.getElementById('notif-fb-list');
    if (!listEl) return;

    var unreadEl = document.getElementById('notif-fb-unread-count');
    var unreadN = unreadCount();
    if (unreadEl) {
      if (unreadN > 0) {
        unreadEl.textContent = unreadN > 99 ? '99+' : String(unreadN);
        unreadEl.style.display = '';
      } else {
        unreadEl.style.display = 'none';
      }
    }

    // Loading kalau belum pernah fetch sukses
    if (!state.initialFetched) {
      listEl.innerHTML = '<div class="notif-fb-loading">'
                       + '<div class="notif-fb-loading-spin"></div>'
                       + '<div>Memuat notifikasi...</div>'
                       + '</div>';
      return;
    }

    var lastReadMs = new Date(state.lastReadAt).getTime();

    // Filter berdasar tab aktif
    var visible = state.list.slice();
    if (state.activeTab === 'unread') {
      visible = visible.filter(function (n) { return isUnread(n, lastReadMs); });
    }

    // Sort terbaru dulu
    visible.sort(function (a, b) {
      return String(b.time || '').localeCompare(String(a.time || ''));
    });

    if (!visible.length) {
      listEl.innerHTML = renderEmptyState(state.activeTab);
      return;
    }

    // Group ke section
    var g = groupBySection(visible);
    var html = '';
    function appendSection(label, items) {
      if (!items.length) return;
      html += '<div class="notif-fb-section">' + escHTML(label) + '</div>';
      items.forEach(function (n) { html += renderItem(n, lastReadMs); });
    }
    appendSection('Hari ini', g.today);
    appendSection('Kemarin', g.yesterday);
    appendSection('Minggu ini', g.thisWeek);
    appendSection('Lebih lama', g.older);

    listEl.innerHTML = html;
  }

  // ═══════════════════════════════════════════════════════════════════
  // POSITIONING (dropdown is body-level, position relative to button)
  // ═══════════════════════════════════════════════════════════════════
  function positionDropdown() {
    var dd  = document.getElementById('notif-dropdown-fb');
    var btn = document.getElementById('notif-btn');
    if (!dd || !btn) return;
    var r = btn.getBoundingClientRect();
    var ddW = dd.offsetWidth || 380;
    var vw = window.innerWidth;

    // Mobile (≤480): pakai full-width dari CSS, tidak perlu calc
    if (vw <= 480) {
      dd.style.top = (r.bottom + 8) + 'px';
      dd.style.left = '';
      dd.style.right = '';
      return;
    }

    // Default: align right edge of dropdown ke right edge of button
    var top   = r.bottom + 8;
    var right = vw - r.right;
    // Clamp supaya tidak overflow ke kiri (kalau viewport sempit)
    if (right + ddW > vw - 8) {
      right = vw - ddW - 8;
      if (right < 8) right = 8;
    }
    dd.style.top   = top + 'px';
    dd.style.right = right + 'px';
    dd.style.left  = '';
  }

  // ═══════════════════════════════════════════════════════════════════
  // OPEN / CLOSE / TOGGLE
  // ═══════════════════════════════════════════════════════════════════

  function openDropdown() {
    var dd = ensureDropdownEl();
    var btn = document.getElementById('notif-btn');
    if (!dd) return;
    state.dropdownOpen = true;
    positionDropdown();
    dd.classList.add('open');
    dd.setAttribute('aria-hidden', 'false');
    if (btn) btn.setAttribute('aria-expanded', 'true');
    renderList();
    // Auto mark-as-read on open kalau ada unread
    if (unreadCount() > 0) {
      // Delay sedikit supaya user lihat dulu yang unread (visual cue)
      setTimeout(markAllRead, 400);
    }
  }

  function closeDropdown() {
    var dd = document.getElementById('notif-dropdown-fb');
    var btn = document.getElementById('notif-btn');
    if (dd) {
      dd.classList.remove('open');
      dd.setAttribute('aria-hidden', 'true');
    }
    if (btn) btn.setAttribute('aria-expanded', 'false');
    state.dropdownOpen = false;
  }

  function toggleDropdown() {
    if (state.dropdownOpen) closeDropdown(); else openDropdown();
  }

  function markAllRead() {
    var session = getSession();
    if (!session || !session.id) return;
    var url = resolveSupabaseUrl();
    if (!url) return;

    // Optimistic: update UI segera
    state.lastReadAt = new Date().toISOString();
    updateBadge();
    if (state.dropdownOpen) renderList();

    fetch(url + '/rest/v1/rpc/mark_notifikasi_read', {
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

  // ═══════════════════════════════════════════════════════════════════
  // REFRESH
  // ═══════════════════════════════════════════════════════════════════
  function refresh() {
    if (state.refreshing) return;
    state.refreshing = true;

    var session = getSession();
    if (!session || !session.id) {
      state.refreshing = false;
      // Tetap render — kasih empty state
      state.initialFetched = true;
      if (state.dropdownOpen) renderList();
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
      state.initialFetched = true;
      updateBadge();
      if (state.dropdownOpen) renderList();
    }).catch(function (e) {
      console.warn('[notif] refresh error:', e);
      // Tetap mark fetched supaya UI tidak stuck di "Memuat..."
      state.initialFetched = true;
      if (state.dropdownOpen) renderList();
    }).then(function () {
      state.refreshing = false;
    });
  }

  function startPolling() {
    if (state.pollTimer) clearInterval(state.pollTimer);
    state.pollTimer = setInterval(refresh, 60 * 1000);
  }

  // ═══════════════════════════════════════════════════════════════════
  // EVENT DELEGATION (anti race-condition)
  // ─────────────────────────────────────────────────────────────────
  // Pasang ONE listener di document — handle SEMUA klik:
  //   1. Klik #notif-btn (atau element di dalamnya) → toggle dropdown
  //   2. Klik tab di dropdown → switch tab
  //   3. Klik tombol "Mark all" → mark read
  //   4. Klik di luar dropdown & button → close dropdown
  // Pendekatan ini bekerja BAHKAN kalau button belum ter-render saat
  // listener di-pasang — karena event delegation cek target saat klik,
  // bukan saat attachment.
  // ═══════════════════════════════════════════════════════════════════

  function setupGlobalListeners() {
    if (window.__notif_listeners_set__) return;
    window.__notif_listeners_set__ = true;

    document.addEventListener('click', function (e) {
      var t = e.target;
      if (!t || !t.closest) return;

      // 1. Klik pada notif-btn (atau child-nya)
      var btn = t.closest('#notif-btn');
      if (btn) {
        e.preventDefault();
        e.stopPropagation();
        toggleDropdown();
        return;
      }

      // 2. Klik pada tab di dropdown
      var tabBtn = t.closest('.notif-fb-tab');
      if (tabBtn) {
        e.preventDefault();
        e.stopPropagation();
        var tab = tabBtn.getAttribute('data-tab');
        if (tab && tab !== state.activeTab) {
          state.activeTab = tab;
          var allTabs = document.querySelectorAll('.notif-fb-tab');
          allTabs.forEach(function (b) {
            b.classList.toggle('active', b.getAttribute('data-tab') === tab);
          });
          renderList();
        }
        return;
      }

      // 3. Klik tombol "Mark all read"
      var markBtn = t.closest('#notif-fb-mark-all');
      if (markBtn) {
        e.preventDefault();
        e.stopPropagation();
        if (unreadCount() > 0) markAllRead();
        return;
      }

      // 4. Klik item notif → biarkan navigate, tapi tutup dropdown dulu
      var itemEl = t.closest('.notif-fb-item');
      if (itemEl) {
        // Don't preventDefault — biarkan link navigate
        closeDropdown();
        return;
      }

      // 5. Klik di luar dropdown & button → close
      if (state.dropdownOpen) {
        var dd = document.getElementById('notif-dropdown-fb');
        if (dd && !dd.contains(t)) {
          closeDropdown();
        }
      }
    }, false);

    // ESC = close dropdown
    document.addEventListener('keydown', function (e) {
      if (e.key === 'Escape' && state.dropdownOpen) {
        closeDropdown();
      }
    });

    // Re-position dropdown saat scroll/resize (kalau open)
    window.addEventListener('scroll', function () {
      if (state.dropdownOpen) positionDropdown();
    }, { passive: true });
    window.addEventListener('resize', function () {
      if (state.dropdownOpen) positionDropdown();
    });

    // Re-fetch saat tab kembali fokus
    document.addEventListener('visibilitychange', function () {
      if (!document.hidden) refresh();
    });
  }

  // ═══════════════════════════════════════════════════════════════════
  // INIT
  // ═══════════════════════════════════════════════════════════════════

  // Resolve SUPABASE_URL & SUPABASE_HEADERS dengan fallback ke bare
  // identifier. Di classic script, top-level `const SUPABASE_URL` tidak
  // attach ke window — hanya bisa diakses via bare reference. Shared.js
  // versi baru sudah explicit set window.SUPABASE_URL, tapi kita defensive
  // di sini supaya kalau shared.js cache lama, tetap bisa fallback.
  function resolveSupabaseUrl() {
    if (typeof window.SUPABASE_URL === 'string' && window.SUPABASE_URL) return window.SUPABASE_URL;
    try {
      // eslint-disable-next-line no-undef
      if (typeof SUPABASE_URL === 'string' && SUPABASE_URL) {
        // eslint-disable-next-line no-undef
        window.SUPABASE_URL = SUPABASE_URL;
        return SUPABASE_URL;
      }
    } catch (_) { /* ReferenceError → tidak ada di scope */ }
    return null;
  }

  function resolveSupabaseAnonKey() {
    if (typeof window.SUPABASE_ANON_KEY === 'string' && window.SUPABASE_ANON_KEY) return window.SUPABASE_ANON_KEY;
    try {
      // eslint-disable-next-line no-undef
      if (typeof SUPABASE_ANON_KEY === 'string' && SUPABASE_ANON_KEY) {
        // eslint-disable-next-line no-undef
        window.SUPABASE_ANON_KEY = SUPABASE_ANON_KEY;
        return SUPABASE_ANON_KEY;
      }
    } catch (_) {}
    return null;
  }

  function ensureHeaders() {
    if (window.SUPABASE_HEADERS && window.SUPABASE_HEADERS.apikey) return true;
    var key = resolveSupabaseAnonKey();
    if (!key) return false;
    window.SUPABASE_HEADERS = {
      'apikey':        key,
      'Authorization': 'Bearer ' + key,
      'Content-Type':  'application/json',
    };
    return true;
  }

  function init() {
    // CSS, dropdown element, dan listener click — ALWAYS setup, tidak
    // tergantung Supabase. Tujuannya: tombol notifikasi PASTI buka dropdown
    // walaupun config belum lengkap. Dropdown akan tampilkan empty state
    // dengan pesan yang sesuai.
    injectCSS();
    ensureDropdownEl();
    setupGlobalListeners();

    // Resolve Supabase config — kalau gagal, tetap jalan tapi skip fetch
    // (dropdown tetap bisa dibuka, hanya saja list selalu kosong).
    var url = resolveSupabaseUrl();
    var hasHeaders = ensureHeaders();

    if (!url || !hasHeaders) {
      console.warn('[notif] SUPABASE_URL/HEADERS belum ter-load. '
        + 'Pastikan urutan script: config.js → 9201-shared.js → 9201-notifikasi.js. '
        + 'Polling disabled, tapi tombol tetap berfungsi (dropdown akan kosong).');
      // Mark as fetched supaya dropdown tampilkan empty state, bukan loading
      state.initialFetched = true;
      return;
    }

    refresh();          // initial fetch
    startPolling();     // 60s polling
    // Update badge sekali setelah topbar pasti sudah render
    setTimeout(updateBadge, 100);
  }

  // Bootstrap: aman di-call kapan saja karena event delegation tidak butuh
  // button ada di DOM. Listener akan hidup selamanya.
  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', init);
  } else {
    setTimeout(init, 0);
  }

  // ═══════════════════════════════════════════════════════════════════
  // PUBLIC API (untuk debugging dari console)
  // ═══════════════════════════════════════════════════════════════════
  window.NotifikasiPortal = {
    refresh: refresh,
    open: openDropdown,
    close: closeDropdown,
    toggle: toggleDropdown,
    markAllRead: markAllRead,
    setTab: function (tab) {
      if (tab === 'all' || tab === 'unread') {
        state.activeTab = tab;
        var allTabs = document.querySelectorAll('.notif-fb-tab');
        allTabs.forEach(function (b) {
          b.classList.toggle('active', b.getAttribute('data-tab') === tab);
        });
        renderList();
      }
    },
    _state: state,
  };
})();
