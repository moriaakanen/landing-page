/**
 * 9201 SHARED — utilitas umum yang dipakai semua halaman
 * ─────────────────────────────────────────────────────────
 * Diload SETELAH config.js, SEBELUM 9201-topbar.js / 9201-sidebar.js /
 * 9201-role-switcher.js dan script init() halaman.
 *
 * Mengexpose ke global scope:
 *   - SUPABASE_HEADERS              header siap-pakai untuk fetch ke Supabase
 *   - esc(s) / escAttr(s)           escape HTML (alias)
 *   - logout()                      hapus session & redirect ke login
 *   - novaCheckSession({...})       seragam: cek session, expiry, ganti-sandi,
 *                                   set active_role default, gate role admin
 *   - novaRpc(fnName, params)       POST RPC dengan error handling konsisten
 *   - BULAN                         array nama bulan Indonesia
 *
 * Catatan:
 * - Fungsi `logout` SENGAJA di-expose sebagai global karena dipanggil dari
 *   onclick="logout()" di template topbar (9201-topbar.js).
 * - Date utilities (parseISODate, fmtTgl, fmtWaktu, dst.) TIDAK ada di sini
 *   karena beberapa halaman punya behavior berbeda halus (mis. em-dash vs
 *   empty string untuk input null). Tetap didefinisikan per-file.
 */
(function () {
  'use strict';

  // ─── Sanity check: config.js wajib sudah di-load duluan ────────
  if (typeof SUPABASE_URL === 'undefined' || typeof SUPABASE_ANON_KEY === 'undefined') {
    console.error('[9201Shared] config.js belum di-load. Pastikan urutan script: '
                + 'config.js → 9201-shared.js → ...');
  }

  // ─── Headers Supabase ──────────────────────────────────────────
  // Pengganti pola `const H = { 'apikey': ..., 'Authorization': ..., ... }`
  // yang sebelumnya diduplikasi di setiap halaman.
  window.SUPABASE_HEADERS = {
    'apikey':        typeof SUPABASE_ANON_KEY !== 'undefined' ? SUPABASE_ANON_KEY : '',
    'Authorization': typeof SUPABASE_ANON_KEY !== 'undefined' ? `Bearer ${SUPABASE_ANON_KEY}` : '',
    'Content-Type':  'application/json'
  };

  // ─── Escape HTML ───────────────────────────────────────────────
  function esc(s) {
    return String(s ?? '').replace(/[&<>"']/g, c =>
      ({ '&':'&amp;', '<':'&lt;', '>':'&gt;', '"':'&quot;', "'":'&#39;' }[c])
    );
  }
  window.esc = esc;
  // Alias backward-compat: beberapa file lama pakai escAttr() — sama persis.
  window.escAttr = esc;

  // ─── Logout ────────────────────────────────────────────────────
  // Dipanggil dari template topbar (onclick="logout()") DAN dari halaman lain.
  function logout() {
    localStorage.removeItem('nova_user');
    window.location.replace('login.html');
  }
  window.logout = logout;

  // ─── novaCheckSession ──────────────────────────────────────────
  /**
   * Seragam check session untuk semua halaman aplikasi (kecuali login &
   * ganti-password yang punya alur sendiri).
   *
   * Behaviour:
   *   1. Tidak ada session             → redirect login.html
   *   2. Session corrupt (parse fail)  → clear & redirect login.html
   *   3. expires_at lewat              → clear & redirect login.html
   *   4. must_change_password === true → redirect ganti-password.html
   *   5. active_role belum di-set      → set default (admin kalau punya, else user)
   *   6. active_role tidak valid       → reset ke role yang dimiliki
   *   7. requireAdmin = true tapi      → redirect index.html
   *      user bukan admin
   *
   * @param {object}  opts
   * @param {boolean} opts.requireAdmin  true untuk halaman admin-only
   * @returns {object|null} session object (sukses), atau null bila sudah redirect
   */
  function novaCheckSession(opts) {
    opts = opts || {};
    const requireAdmin = !!opts.requireAdmin;

    let s;
    try {
      s = JSON.parse(localStorage.getItem('nova_user') || 'null');
    } catch (e) {
      localStorage.removeItem('nova_user');
      window.location.replace('login.html');
      return null;
    }

    if (!s) {
      window.location.replace('login.html');
      return null;
    }
    if (s.expires_at && Date.now() > s.expires_at) {
      localStorage.removeItem('nova_user');
      window.location.replace('login.html');
      return null;
    }
    if (s.must_change_password) {
      window.location.replace('ganti-password.html');
      return null;
    }

    // Sumber kebenaran role: DB (lewat getUserRoles dari config.js).
    // Kalau config.js belum loaded (seharusnya tidak terjadi), fallback aman.
    const roles = (typeof getUserRoles === 'function')
      ? getUserRoles(s)
      : (Array.isArray(s.roles) && s.roles.length ? s.roles : ['user']);

    if (!s.active_role) {
      s.active_role = roles.includes('admin') ? 'admin' : 'user';
      localStorage.setItem('nova_user', JSON.stringify(s));
    }
    // Kalau active_role yang ter-pin bukan role yang user miliki (mis. admin
    // dicabut admin lain), reset ke role yang sah.
    if (!roles.includes(s.active_role)) {
      s.active_role = roles[0] || 'user';
      localStorage.setItem('nova_user', JSON.stringify(s));
    }

    if (requireAdmin) {
      if (!roles.includes('admin') || s.active_role !== 'admin') {
        window.location.replace('index.html');
        return null;
      }
    }

    // Background refresh full_name dari DB. Tidak di-await — caller dapat
    // session sync segera. Visual topbar akan ter-update saat fetch selesai.
    // Ini menjamin nama lengkap selalu fresh dari users.full_name walaupun
    // session lama menyimpan fallback ke username atau null.
    if (typeof novaEnsureFullName === 'function') {
      novaEnsureFullName(s, { force: true });
    }

    return s;
  }
  window.novaCheckSession = novaCheckSession;

  // ─── novaRpc ───────────────────────────────────────────────────
  /**
   * POST ke endpoint Supabase RPC dengan error handling konsisten.
   *
   * Versi sebelumnya (callRPC di login.html / ganti-password.html) TIDAK
   * memeriksa res.ok — kalau server return 4xx/5xx dengan body JSON error,
   * caller mengira sukses. Versi ini selalu lempar Error dengan pesan yang
   * berasal dari server kalau response non-2xx.
   *
   * @param {string} fnName  nama RPC function
   * @param {object} params  payload (akan di-JSON.stringify)
   * @returns {*}            hasil parse JSON dari body, atau null kalau body kosong
   * @throws  {Error}        kalau status non-2xx
   */
  async function novaRpc(fnName, params) {
    const res = await fetch(`${SUPABASE_URL}/rest/v1/rpc/${fnName}`, {
      method:  'POST',
      headers: window.SUPABASE_HEADERS,
      body:    JSON.stringify(params || {})
    });
    if (!res.ok) {
      let msg = `HTTP ${res.status}`;
      try {
        const err = await res.json();
        msg = err.message || err.hint || err.details || msg;
      } catch (_) { /* body bukan JSON, biarkan msg default */ }
      throw new Error(msg);
    }
    const text = await res.text();
    if (!text) return null;
    try { return JSON.parse(text); } catch (_) { return text; }
  }
  window.novaRpc = novaRpc;

  // ─── Konstanta: nama bulan Indonesia ───────────────────────────
  // Sebelumnya didefinisikan ulang sebagai BULAN / BULAN_ID di banyak file.
  window.BULAN = [
    'Januari','Februari','Maret','April','Mei','Juni',
    'Juli','Agustus','September','Oktober','November','Desember'
  ];

  // ─── novaEnsureFullName ────────────────────────────────────────
  /**
   * Pastikan session.full_name terisi dengan nilai resmi dari tabel `users`.
   * Kalau kosong/null, fetch dari `users.full_name` berdasarkan session.id.
   *
   * Side effect:
   *   - Update session.full_name di-place
   *   - Persist ke localStorage('nova_user')
   *   - Update visual: #topbar-username, #topbar-avatar, #usd-header-name
   *   - Update visual: avoid recursive call ke setUser (passing { skipEnsure:true })
   *
   * Idempotent untuk panggilan biasa: kalau session.full_name sudah ada,
   * return langsung TANPA fetch. Untuk paksa refresh dari DB (mis. saat
   * user login lama yang full_name-nya pernah di-fallback ke username),
   * panggil dengan { force: true }.
   *
   * Anti-retry: kalau fetch gagal/data tidak ada, set fallback = username
   *             supaya tidak retry tiap render.
   *
   * @param {object} session - session object dari novaCheckSession
   * @param {object} [opts]  - { force?: boolean }
   * @returns {Promise<object>} session yang sudah di-update
   */
  async function novaEnsureFullName(session, opts) {
    if (!session) return session;
    const force = !!(opts && opts.force);
    if (!force && session.full_name) return session; // sudah terisi

    const fallback = session.username || '';

    // DEBUG: log state awal supaya bisa diagnose dari console
    console.log('[9201:ensureFullName] start', {
      sessionId: session.id,
      sessionUsername: session.username,
      sessionFullName: session.full_name,
      force: force
    });

    if (!session.id) {
      if (!session.full_name) session.full_name = fallback;
      try { localStorage.setItem('nova_user', JSON.stringify(session)); } catch(_) {}
      console.warn('[9201:ensureFullName] session.id kosong — tidak bisa fetch');
      return session;
    }

    try {
      const url = `${SUPABASE_URL}/rest/v1/users?id=eq.${encodeURIComponent(session.id)}&select=full_name&limit=1`;
      console.log('[9201:ensureFullName] fetching:', url);
      const res = await fetch(url, { headers: window.SUPABASE_HEADERS });
      console.log('[9201:ensureFullName] response status:', res.status);

      let nama = '';
      if (res.ok) {
        const rows = await res.json();
        console.log('[9201:ensureFullName] response rows:', rows);
        nama = (rows && rows[0] && rows[0].full_name) || '';
      } else {
        const errBody = await res.text().catch(() => '');
        console.warn('[9201:ensureFullName] response not ok:', res.status, errBody);
      }

      if (nama) {
        session.full_name = nama;
        console.log('[9201:ensureFullName] ✓ full_name resolved:', nama);
      } else if (!session.full_name) {
        session.full_name = fallback;
        console.warn('[9201:ensureFullName] DB tidak return nama, fallback ke username:', fallback);
      }
    } catch (e) {
      console.warn('[9201:ensureFullName] fetch gagal:', e);
      if (!session.full_name) session.full_name = fallback;
    }

    try { localStorage.setItem('nova_user', JSON.stringify(session)); } catch(_) {}

    // Update visual immediately (kalau topbar/role-switcher sudah ter-render).
    // Pass skipEnsure:true ke setUser supaya tidak rekursif memanggil
    // novaEnsureFullName lagi (mencegah infinite loop).
    if (window.Topbar9201 && typeof Topbar9201.setUser === 'function') {
      Topbar9201.setUser(session, { skipEnsure: true });
    }
    const usdName = document.getElementById('usd-header-name');
    if (usdName && session.full_name) usdName.textContent = session.full_name;

    return session;
  }
  window.novaEnsureFullName = novaEnsureFullName;

})();
