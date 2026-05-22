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

  // ─── CSS Injection: Custom checkbox & radio (rendering full lewat CSS) ─
  // Native <input type="checkbox"> di-render OS / browser dan tampil jelek
  // (terutama di Chrome/Edge — kotak abu-abu pucat yang tidak match dengan
  // tema gold/navy portal). Kita override SEMUA checkbox global supaya
  // konsisten dan rapi tanpa perlu edit tiap file HTML.
  //
  // Dipakai dengan:
  //   <input type="checkbox">                    → varian default (navy)
  //   <input type="checkbox" class="ck-success"> → varian hijau (approve)
  //   <input type="checkbox" class="ck-gold">    → varian gold
  //   <input type="checkbox" class="ck-danger">  → varian merah
  //   <input type="checkbox" class="ck-sm">      → ukuran kecil (14px)
  //   <input type="checkbox" class="ck-lg">      → ukuran besar (20px)
  //
  // Mendukung state: hover, focus-visible, checked, indeterminate, disabled.
  // SVG checkmark & dash dilekatkan sebagai background-image (data URL)
  // sehingga tidak butuh font / asset eksternal.
  if (!document.getElementById('9201-controls-css')) {
    var ctrlCSS = `
    /* ── Custom checkbox (global) ──────────────────────────────── */
    input[type="checkbox"]{
      -webkit-appearance:none;-moz-appearance:none;appearance:none;
      --ck-accent:#0d2340;
      --ck-border:#c9c2b6;
      width:17px;height:17px;
      border:1.5px solid var(--ck-border);
      border-radius:5px;
      background-color:#fff;
      cursor:pointer;
      margin:0;
      vertical-align:middle;
      position:relative;
      flex-shrink:0;
      display:inline-block;
      transition:background-color .15s ease,border-color .15s ease,box-shadow .15s ease,transform .08s ease;
      background-repeat:no-repeat;
      background-position:center;
      /* Pakai persentase supaya checkmark tetap proporsional walaupun
         per-page CSS override width/height (mis. jadi 13px atau 20px). */
      background-size:78% 78%;
      print-color-adjust:exact;
      -webkit-print-color-adjust:exact;
    }
    input[type="checkbox"]:hover:not(:disabled){
      border-color:var(--ck-accent);
    }
    input[type="checkbox"]:active:not(:disabled){
      transform:scale(.92);
    }
    input[type="checkbox"]:focus-visible{
      outline:none;
      box-shadow:0 0 0 3px rgba(200,168,75,.35);
    }
    input[type="checkbox"]:checked{
      background-color:var(--ck-accent);
      border-color:var(--ck-accent);
      background-image:url("data:image/svg+xml;utf8,<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 16 16'><path d='M3.5 8.4 L6.7 11.5 L12.7 4.7' fill='none' stroke='%23ffffff' stroke-width='2.4' stroke-linecap='round' stroke-linejoin='round'/></svg>");
    }
    input[type="checkbox"]:indeterminate{
      background-color:var(--ck-accent);
      border-color:var(--ck-accent);
      background-image:url("data:image/svg+xml;utf8,<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 16 16'><line x1='4' y1='8' x2='12' y2='8' stroke='%23ffffff' stroke-width='2.4' stroke-linecap='round'/></svg>");
    }
    input[type="checkbox"]:disabled{
      cursor:not-allowed;
      opacity:.45;
      background-color:#f1ede5;
      border-color:#d8d3c8;
    }
    input[type="checkbox"]:checked:disabled,
    input[type="checkbox"]:indeterminate:disabled{
      background-color:#9ca3af;
      border-color:#9ca3af;
    }
    /* Variants warna */
    input[type="checkbox"].ck-success{ --ck-accent:#1a7a4a }
    input[type="checkbox"].ck-gold   { --ck-accent:#c8a84b }
    input[type="checkbox"].ck-danger { --ck-accent:#c0392b }
    /* Variants ukuran */
    input[type="checkbox"].ck-sm{ width:14px;height:14px;border-radius:4px }
    input[type="checkbox"].ck-lg{ width:20px;height:20px;border-radius:6px }
    `;
    var ctrlStyle = document.createElement('style');
    ctrlStyle.id = '9201-controls-css';
    ctrlStyle.textContent = ctrlCSS;
    // Sisipkan di awal <head> supaya specificity rules dari per-page CSS
    // tetap bisa override jika benar-benar dibutuhkan.
    if (document.head.firstChild) {
      document.head.insertBefore(ctrlStyle, document.head.firstChild);
    } else {
      document.head.appendChild(ctrlStyle);
    }
  }

  // ─── Sanity check: config.js wajib sudah di-load duluan ────────
  if (typeof SUPABASE_URL === 'undefined' || typeof SUPABASE_ANON_KEY === 'undefined') {
    console.error('[9201Shared] config.js belum di-load. Pastikan urutan script: '
                + 'config.js → 9201-shared.js → ...');
  }

  // ─── Headers Supabase ──────────────────────────────────────────
  // Pengganti pola `const H = { 'apikey': ..., 'Authorization': ..., ... }`
  // yang sebelumnya diduplikasi di setiap halaman.
  //
  // FIX 2026-05: SUPABASE_URL dan SUPABASE_ANON_KEY di config.js declare
  // sebagai top-level `const`. Di classic script (non-module), top-level
  // `const` TIDAK attach ke window. Bisa diakses via bare identifier
  // (mis. `SUPABASE_URL`) dari script lain di realm yang sama, TAPI
  // `window.SUPABASE_URL` adalah undefined.
  //
  // Banyak file (notifikasi.js, dst.) cek `window.SUPABASE_URL` sebelum
  // init — jadi kita explicit attach di sini supaya semua code path bekerja
  // baik pakai bare identifier MAUPUN window-prefix.
  if (typeof SUPABASE_URL !== 'undefined') {
    window.SUPABASE_URL = SUPABASE_URL;
  }
  if (typeof SUPABASE_ANON_KEY !== 'undefined') {
    window.SUPABASE_ANON_KEY = SUPABASE_ANON_KEY;
  }

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
  function jsArg(v) {
    return JSON.stringify(v == null ? '' : String(v))
      .replace(/</g, '\\u003C')
      .replace(/>/g, '\\u003E')
      .replace(/&/g, '\\u0026')
      .replace(/\u2028/g, '\\u2028')
      .replace(/\u2029/g, '\\u2029');
  }
  window.jsArg = jsArg;
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

  async function novaVerifyAdminSession(session) {
    if (!session || !session.id) {
      localStorage.removeItem('nova_user');
      window.location.replace('login.html');
      return null;
    }

    try {
      const res = await fetch(
        `${SUPABASE_URL}/rest/v1/users?id=eq.${encodeURIComponent(session.id)}&select=id,username,full_name,must_change_password,role,roles&limit=1`,
        { headers: window.SUPABASE_HEADERS }
      );
      if (!res.ok) throw new Error(`HTTP ${res.status}`);
      const rows = await res.json();
      const fresh = rows && rows[0];
      if (!fresh || fresh.username !== session.username) {
        throw new Error('Session tidak cocok dengan data pengguna.');
      }

      const roles = (typeof getUserRoles === 'function') ? getUserRoles(fresh) : [];
      if (fresh.must_change_password) {
        session.must_change_password = true;
        localStorage.setItem('nova_user', JSON.stringify(session));
        window.location.replace('ganti-password.html');
        return null;
      }
      if (!roles.includes('admin')) {
        session.role = fresh.role || null;
        session.roles = roles.length ? roles : ['user'];
        session.active_role = 'user';
        localStorage.setItem('nova_user', JSON.stringify(session));
        window.location.replace('index.html');
        return null;
      }

      session.full_name = fresh.full_name || session.full_name || null;
      session.must_change_password = !!fresh.must_change_password;
      session.role = fresh.role || null;
      session.roles = roles;
      session.active_role = 'admin';
      localStorage.setItem('nova_user', JSON.stringify(session));
      return session;
    } catch (e) {
      console.error('[novaVerifyAdminSession] gagal verifikasi admin:', e);
      localStorage.removeItem('nova_user');
      window.location.replace('login.html');
      return null;
    }
  }
  window.novaVerifyAdminSession = novaVerifyAdminSession;

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

    if (!session.id) {
      if (!session.full_name) session.full_name = fallback;
      try { localStorage.setItem('nova_user', JSON.stringify(session)); } catch(_) {}
      return session;
    }

    try {
      const res = await fetch(
        `${SUPABASE_URL}/rest/v1/users?id=eq.${encodeURIComponent(session.id)}&select=full_name&limit=1`,
        { headers: window.SUPABASE_HEADERS }
      );
      let nama = '';
      if (res.ok) {
        const rows = await res.json();
        nama = (rows && rows[0] && rows[0].full_name) || '';
      }
      // Kalau nama dari DB ada, pakai itu. Kalau tidak, jangan timpa
      // session.full_name yang mungkin sudah valid — tapi kalau session
      // juga kosong, fallback ke username.
      if (nama) {
        session.full_name = nama;
      } else if (!session.full_name) {
        session.full_name = fallback;
      }
    } catch (e) {
      console.warn('[novaEnsureFullName] fetch gagal:', e);
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
