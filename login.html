<!DOCTYPE html>
<html lang="id">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>NOVA — Login</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link href="https://fonts.googleapis.com/css2?family=Bebas+Neue&family=DM+Sans:ital,opsz,wght@0,9..40,300;0,9..40,400;0,9..40,500;1,9..40,300&display=swap" rel="stylesheet">
<style>
  :root {
    --bg: #080808;
    --surface: #111111;
    --surface2: #181818;
    --border: rgba(255,255,255,0.07);
    --border-hover: rgba(255,255,255,0.15);
    --accent: #c8f230;
    --accent2: #f23078;
    --text: #f0ede8;
    --muted: #7a7773;
    --error: #f23078;
    --success: #c8f230;
  }

  *, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }
  html, body { height: 100%; }

  body {
    background: var(--bg);
    color: var(--text);
    font-family: 'DM Sans', sans-serif;
    font-weight: 300;
    overflow: hidden;
    cursor: none;
  }

  /* Cursor */
  .cursor {
    width: 10px; height: 10px;
    background: var(--accent); border-radius: 50%;
    position: fixed; top: 0; left: 0;
    pointer-events: none; z-index: 9999;
    transition: transform 0.15s ease, background 0.2s ease;
    mix-blend-mode: difference;
  }
  .cursor.hover { transform: scale(3.5); background: var(--accent2); }
  .cursor-ring {
    width: 36px; height: 36px;
    border: 1px solid rgba(200,242,48,0.35); border-radius: 50%;
    position: fixed; top: 0; left: 0;
    pointer-events: none; z-index: 9998;
    transition: all 0.25s ease;
  }

  /* Noise */
  body::after {
    content: '';
    position: fixed; inset: 0;
    background-image: url("data:image/svg+xml,%3Csvg viewBox='0 0 200 200' xmlns='http://www.w3.org/2000/svg'%3E%3Cfilter id='n'%3E%3CfeTurbulence type='fractalNoise' baseFrequency='0.85' numOctaves='4' stitchTiles='stitch'/%3E%3C/filter%3E%3Crect width='100%25' height='100%25' filter='url(%23n)' opacity='0.035'/%3E%3C/svg%3E");
    pointer-events: none; z-index: 1000;
  }

  /* Layout */
  .wrapper {
    display: grid;
    grid-template-columns: 1fr 1fr;
    height: 100vh;
  }

  /* LEFT PANEL */
  .left-panel {
    position: relative;
    overflow: hidden;
    display: flex; flex-direction: column;
    justify-content: space-between;
    padding: 48px;
    border-right: 1px solid var(--border);
  }

  /* Animated grid bg */
  .left-panel::before {
    content: '';
    position: absolute; inset: 0;
    background-image:
      linear-gradient(rgba(200,242,48,0.05) 1px, transparent 1px),
      linear-gradient(90deg, rgba(200,242,48,0.05) 1px, transparent 1px);
    background-size: 48px 48px;
    animation: gridScroll 15s linear infinite;
  }
  @keyframes gridScroll {
    from { transform: translateY(0); }
    to { transform: translateY(48px); }
  }

  /* Orbs */
  .orb {
    position: absolute; border-radius: 50%;
    filter: blur(80px); pointer-events: none;
  }
  .orb-1 {
    width: 400px; height: 400px;
    background: rgba(200,242,48,0.15);
    top: -100px; left: -100px;
    animation: float 9s ease-in-out infinite;
  }
  .orb-2 {
    width: 250px; height: 250px;
    background: rgba(242,48,120,0.12);
    bottom: 50px; right: -50px;
    animation: float 12s ease-in-out infinite reverse;
  }
  @keyframes float {
    0%, 100% { transform: translate(0, 0); }
    50% { transform: translate(20px, -25px); }
  }

  .left-top { position: relative; z-index: 2; }
  .brand {
    font-family: 'Bebas Neue', sans-serif;
    font-size: 32px; letter-spacing: 6px; color: var(--accent);
    display: inline-block;
    animation: fadeUp 0.7s ease both;
  }

  .left-center {
    position: relative; z-index: 2;
    animation: fadeUp 0.7s ease 0.2s both;
  }
  .left-title {
    font-family: 'Bebas Neue', sans-serif;
    font-size: clamp(56px, 6vw, 88px);
    line-height: 0.92; letter-spacing: 2px;
    margin-bottom: 24px;
  }
  .left-title .stroke {
    -webkit-text-stroke: 1px rgba(240,237,232,0.25);
    color: transparent;
  }
  .left-title .accent { color: var(--accent); }

  .left-desc {
    font-size: 15px; line-height: 1.8;
    color: var(--muted); max-width: 340px;
  }

  .left-bottom {
    position: relative; z-index: 2;
    animation: fadeUp 0.7s ease 0.4s both;
  }
  .left-stats { display: flex; gap: 40px; }
  .l-stat-num {
    font-family: 'Bebas Neue', sans-serif;
    font-size: 36px; color: var(--accent); letter-spacing: 2px;
    display: block;
  }
  .l-stat-label {
    font-size: 11px; letter-spacing: 1.5px;
    text-transform: uppercase; color: var(--muted);
  }

  /* RIGHT PANEL */
  .right-panel {
    display: flex; align-items: center; justify-content: center;
    padding: 48px;
    position: relative;
  }

  .login-box {
    width: 100%; max-width: 420px;
    animation: fadeUp 0.8s ease 0.1s both;
  }

  .login-header { margin-bottom: 40px; }
  .login-greeting {
    font-size: 12px; letter-spacing: 3px;
    text-transform: uppercase; color: var(--accent);
    margin-bottom: 12px; display: block;
  }
  .login-title {
    font-family: 'Bebas Neue', sans-serif;
    font-size: 52px; letter-spacing: 2px; line-height: 1;
    margin-bottom: 8px;
  }
  .login-sub { font-size: 14px; color: var(--muted); line-height: 1.6; }
  .login-sub a { color: var(--accent); text-decoration: none; transition: opacity 0.2s; }
  .login-sub a:hover { opacity: 0.8; }

  /* Tabs */
  .auth-tabs {
    display: flex; gap: 0;
    border: 1px solid var(--border);
    margin-bottom: 32px;
    position: relative;
  }
  .tab-btn {
    flex: 1; padding: 12px;
    background: transparent; border: none;
    font-family: 'DM Sans', sans-serif;
    font-size: 13px; letter-spacing: 1.5px;
    text-transform: uppercase; color: var(--muted);
    cursor: none; transition: color 0.3s;
    position: relative; z-index: 1;
  }
  .tab-btn.active { color: #080808; }
  .tab-indicator {
    position: absolute; top: 0; left: 0;
    width: 50%; height: 100%;
    background: var(--accent);
    transition: transform 0.3s cubic-bezier(0.4, 0, 0.2, 1);
    z-index: 0;
  }
  .tab-indicator.register { transform: translateX(100%); }

  /* Form */
  .form-group { margin-bottom: 20px; position: relative; }
  .form-label {
    display: block; font-size: 11px;
    letter-spacing: 2px; text-transform: uppercase;
    color: var(--muted); margin-bottom: 10px;
  }
  .form-input {
    width: 100%; padding: 14px 18px;
    background: var(--surface);
    border: 1px solid var(--border);
    color: var(--text);
    font-family: 'DM Sans', sans-serif;
    font-size: 14px; font-weight: 300;
    outline: none;
    transition: border-color 0.3s, background 0.3s;
    cursor: none;
  }
  .form-input::placeholder { color: rgba(122,119,115,0.5); }
  .form-input:focus {
    border-color: var(--accent);
    background: var(--surface2);
  }
  .form-input.error { border-color: var(--error); }

  /* Password toggle */
  .input-wrapper { position: relative; }
  .toggle-pass {
    position: absolute; right: 16px; top: 50%;
    transform: translateY(-50%);
    background: none; border: none;
    color: var(--muted); cursor: none;
    font-size: 16px; padding: 4px;
    transition: color 0.3s;
  }
  .toggle-pass:hover { color: var(--text); }

  /* Error message */
  .error-msg {
    font-size: 12px; color: var(--error);
    margin-top: 6px; display: none;
    letter-spacing: 0.3px;
  }
  .error-msg.show { display: block; }

  /* Alert box */
  .alert {
    padding: 14px 16px;
    border: 1px solid;
    font-size: 13px; line-height: 1.6;
    margin-bottom: 20px;
    display: none;
    animation: fadeUp 0.3s ease;
  }
  .alert.show { display: block; }
  .alert.error { border-color: rgba(242,48,120,0.4); background: rgba(242,48,120,0.05); color: var(--error); }
  .alert.success { border-color: rgba(200,242,48,0.4); background: rgba(200,242,48,0.05); color: var(--accent); }

  /* Submit button */
  .btn-submit {
    width: 100%; padding: 16px;
    background: var(--accent); color: #080808;
    border: none; font-family: 'DM Sans', sans-serif;
    font-size: 14px; font-weight: 500;
    letter-spacing: 2px; text-transform: uppercase;
    cursor: none; position: relative; overflow: hidden;
    transition: transform 0.2s, background 0.3s;
    margin-bottom: 20px;
  }
  .btn-submit::after {
    content: ''; position: absolute; inset: 0;
    background: rgba(255,255,255,0.2);
    transform: translateX(-100%);
    transition: transform 0.4s ease;
  }
  .btn-submit:hover::after { transform: translateX(100%); }
  .btn-submit:hover { transform: translateY(-1px); }
  .btn-submit:disabled {
    opacity: 0.6; transform: none; cursor: not-allowed;
  }
  .btn-submit.loading { color: transparent; }
  .btn-submit.loading::before {
    content: '';
    position: absolute; top: 50%; left: 50%;
    width: 18px; height: 18px;
    border: 2px solid rgba(8,8,8,0.3);
    border-top-color: #080808;
    border-radius: 50%;
    transform: translate(-50%, -50%);
    animation: spin 0.7s linear infinite;
  }
  @keyframes spin { to { transform: translate(-50%, -50%) rotate(360deg); } }

  /* Divider */
  .divider {
    display: flex; align-items: center; gap: 16px;
    margin-bottom: 20px;
  }
  .divider-line { flex: 1; height: 1px; background: var(--border); }
  .divider-text { font-size: 11px; letter-spacing: 2px; text-transform: uppercase; color: var(--muted); }

  /* Google button */
  .btn-google {
    width: 100%; padding: 14px;
    background: transparent;
    border: 1px solid var(--border);
    color: var(--text);
    font-family: 'DM Sans', sans-serif;
    font-size: 13px; font-weight: 300; letter-spacing: 1px;
    cursor: none; display: flex; align-items: center;
    justify-content: center; gap: 10px;
    transition: border-color 0.3s, transform 0.2s, background 0.3s;
  }
  .btn-google:hover {
    border-color: var(--border-hover);
    background: rgba(255,255,255,0.03);
    transform: translateY(-1px);
  }
  .google-icon { width: 18px; height: 18px; flex-shrink: 0; }

  /* Footer note */
  .login-footer {
    margin-top: 28px; text-align: center;
    font-size: 12px; color: var(--muted); line-height: 1.8;
  }
  .login-footer a { color: var(--text); text-decoration: none; border-bottom: 1px solid var(--border); transition: border-color 0.3s; }
  .login-footer a:hover { border-color: var(--text); }

  /* Register-only fields */
  .register-only { display: none; }
  .register-only.show { display: block; }

  @keyframes fadeUp {
    from { opacity: 0; transform: translateY(20px); }
    to { opacity: 1; transform: translateY(0); }
  }

  /* Success screen */
  .success-screen {
    display: none;
    flex-direction: column; align-items: center;
    justify-content: center; text-align: center;
    gap: 20px;
    animation: fadeUp 0.5s ease;
  }
  .success-screen.show { display: flex; }
  .success-icon {
    width: 72px; height: 72px;
    border: 2px solid var(--accent);
    border-radius: 50%;
    display: flex; align-items: center; justify-content: center;
    font-size: 28px;
    animation: popIn 0.5s cubic-bezier(0.175, 0.885, 0.32, 1.275) both;
  }
  @keyframes popIn {
    from { transform: scale(0); opacity: 0; }
    to { transform: scale(1); opacity: 1; }
  }
  .success-title {
    font-family: 'Bebas Neue', sans-serif;
    font-size: 40px; letter-spacing: 2px; color: var(--accent);
  }
  .success-msg { font-size: 14px; color: var(--muted); line-height: 1.8; }
  .redirect-bar {
    width: 200px; height: 2px;
    background: rgba(255,255,255,0.1);
    position: relative; overflow: hidden;
  }
  .redirect-fill {
    height: 100%; background: var(--accent);
    animation: fillBar 2s linear forwards;
  }
  @keyframes fillBar { from { width: 0; } to { width: 100%; } }

  @media (max-width: 768px) {
    body { overflow: auto; cursor: auto; }
    .cursor, .cursor-ring { display: none; }
    .wrapper { grid-template-columns: 1fr; }
    .left-panel { display: none; }
    .right-panel { min-height: 100vh; }
  }
</style>
</head>
<body>

<div class="cursor" id="cursor"></div>
<div class="cursor-ring" id="cursorRing"></div>

<div class="wrapper">

  <!-- LEFT PANEL -->
  <div class="left-panel">
    <div class="orb orb-1"></div>
    <div class="orb orb-2"></div>

    <div class="left-top">
      <span class="brand">NOVA</span>
    </div>

    <div class="left-center">
      <h2 class="left-title">
        SELAMAT<br>
        <span class="accent">DATANG</span><br>
        <span class="stroke">KEMBALI</span>
      </h2>
      <p class="left-desc">
        Masuk ke platform dan lanjutkan membangun masa depan bersama tim kamu. Semua yang kamu butuhkan ada di sini.
      </p>
    </div>

    <div class="left-bottom">
      <div class="left-stats">
        <div>
          <span class="l-stat-num">50K+</span>
          <span class="l-stat-label">Pengguna</span>
        </div>
        <div>
          <span class="l-stat-num">99.9%</span>
          <span class="l-stat-label">Uptime</span>
        </div>
        <div>
          <span class="l-stat-num">4.9★</span>
          <span class="l-stat-label">Rating</span>
        </div>
      </div>
    </div>
  </div>

  <!-- RIGHT PANEL -->
  <div class="right-panel">
    <div class="login-box">

      <!-- Header -->
      <div class="login-header">
        <span class="login-greeting">// Akses Platform</span>
        <h1 class="login-title" id="form-title">MASUK</h1>
        <p class="login-sub" id="form-sub">
          Belum punya akun? <a href="#" onclick="switchTab('register'); return false;">Daftar sekarang</a>
        </p>
      </div>

      <!-- Tabs -->
      <div class="auth-tabs">
        <div class="tab-indicator" id="tab-indicator"></div>
        <button class="tab-btn active" id="tab-login" onclick="switchTab('login')">Masuk</button>
        <button class="tab-btn" id="tab-register" onclick="switchTab('register')">Daftar</button>
      </div>

      <!-- Alert -->
      <div class="alert" id="alert-box"></div>

      <!-- Form -->
      <div id="auth-form">

        <!-- Nama (register only) -->
        <div class="form-group register-only" id="field-name">
          <label class="form-label">Nama Lengkap</label>
          <input type="text" class="form-input" id="input-name" placeholder="John Doe" autocomplete="name">
          <span class="error-msg" id="err-name">Nama tidak boleh kosong</span>
        </div>

        <!-- Email -->
        <div class="form-group">
          <label class="form-label">Email</label>
          <input type="email" class="form-input" id="input-email" placeholder="kamu@email.com" autocomplete="email">
          <span class="error-msg" id="err-email">Email tidak valid</span>
        </div>

        <!-- Password -->
        <div class="form-group">
          <label class="form-label">Password</label>
          <div class="input-wrapper">
            <input type="password" class="form-input" id="input-password" placeholder="••••••••" autocomplete="current-password">
            <button class="toggle-pass" onclick="togglePassword('input-password', this)">👁</button>
          </div>
          <span class="error-msg" id="err-password">Password minimal 6 karakter</span>
        </div>

        <!-- Konfirmasi password (register only) -->
        <div class="form-group register-only" id="field-confirm">
          <label class="form-label">Konfirmasi Password</label>
          <div class="input-wrapper">
            <input type="password" class="form-input" id="input-confirm" placeholder="••••••••" autocomplete="new-password">
            <button class="toggle-pass" onclick="togglePassword('input-confirm', this)">👁</button>
          </div>
          <span class="error-msg" id="err-confirm">Password tidak cocok</span>
        </div>

        <!-- Submit -->
        <button class="btn-submit" id="btn-submit" onclick="handleSubmit()">
          <span id="btn-text">MASUK</span>
        </button>

        <!-- Divider -->
        <div class="divider">
          <div class="divider-line"></div>
          <span class="divider-text">atau</span>
          <div class="divider-line"></div>
        </div>

        <!-- Google -->
        <button class="btn-google" onclick="handleGoogle()">
          <svg class="google-icon" viewBox="0 0 24 24">
            <path fill="#4285F4" d="M22.56 12.25c0-.78-.07-1.53-.2-2.25H12v4.26h5.92c-.26 1.37-1.04 2.53-2.21 3.31v2.77h3.57c2.08-1.92 3.28-4.74 3.28-8.09z"/>
            <path fill="#34A853" d="M12 23c2.97 0 5.46-.98 7.28-2.66l-3.57-2.77c-.98.66-2.23 1.06-3.71 1.06-2.86 0-5.29-1.93-6.16-4.53H2.18v2.84C3.99 20.53 7.7 23 12 23z"/>
            <path fill="#FBBC05" d="M5.84 14.09c-.22-.66-.35-1.36-.35-2.09s.13-1.43.35-2.09V7.07H2.18C1.43 8.55 1 10.22 1 12s.43 3.45 1.18 4.93l2.85-2.22.81-.62z"/>
            <path fill="#EA4335" d="M12 5.38c1.62 0 3.06.56 4.21 1.64l3.15-3.15C17.45 2.09 14.97 1 12 1 7.7 1 3.99 3.47 2.18 7.07l3.66 2.84c.87-2.6 3.3-4.53 6.16-4.53z"/>
          </svg>
          Lanjutkan dengan Google
        </button>

      </div>

      <!-- Success screen -->
      <div class="success-screen" id="success-screen">
        <div class="success-icon">✓</div>
        <div class="success-title" id="success-title">BERHASIL!</div>
        <p class="success-msg" id="success-msg">Login berhasil. Mengalihkan ke dashboard...</p>
        <div class="redirect-bar"><div class="redirect-fill"></div></div>
      </div>

      <!-- Footer -->
      <div class="login-footer">
        Dengan masuk, kamu menyetujui <a href="#">Syarat & Ketentuan</a> dan <a href="#">Kebijakan Privasi</a> kami.
      </div>

    </div>
  </div>
</div>

<script>
  // =============================================
  // KONFIGURASI SUPABASE — sama dengan index.html
  // =============================================
  const SUPABASE_URL = 'https://jsmmtqeoukkgugorrvmg.supabase.co';
  const SUPABASE_ANON_KEY = 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImpzbW10cWVvdWtrZ3Vnb3Jydm1nIiwicm9sZSI6ImFub24iLCJpYXQiOjE3NzYwOTg2NzksImV4cCI6MjA5MTY3NDY3OX0.O03NscRcj4RcNx-P3j65hO7XXLRSkbyJcwpcArpqHBQ';

  // =============================================
  // STATE
  // =============================================
  let currentTab = 'login';

  // =============================================
  // SUPABASE AUTH HELPERS
  // =============================================
  async function supabaseRequest(endpoint, body) {
    const res = await fetch(`${SUPABASE_URL}/auth/v1/${endpoint}`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'apikey': SUPABASE_ANON_KEY,
      },
      body: JSON.stringify(body)
    });
    return res.json();
  }

  // =============================================
  // TAB SWITCH
  // =============================================
  function switchTab(tab) {
    currentTab = tab;
    const isRegister = tab === 'register';

    document.getElementById('tab-login').classList.toggle('active', !isRegister);
    document.getElementById('tab-register').classList.toggle('active', isRegister);
    document.getElementById('tab-indicator').classList.toggle('register', isRegister);

    document.querySelectorAll('.register-only').forEach(el => {
      el.classList.toggle('show', isRegister);
    });

    document.getElementById('form-title').textContent = isRegister ? 'DAFTAR' : 'MASUK';
    document.getElementById('btn-text').textContent = isRegister ? 'BUAT AKUN' : 'MASUK';
    document.getElementById('form-sub').innerHTML = isRegister
      ? 'Sudah punya akun? <a href="#" onclick="switchTab(\'login\'); return false;">Masuk di sini</a>'
      : 'Belum punya akun? <a href="#" onclick="switchTab(\'register\'); return false;">Daftar sekarang</a>';

    clearErrors();
    hideAlert();
  }

  // =============================================
  // VALIDASI
  // =============================================
  function validateForm() {
    let valid = true;
    clearErrors();

    const email = document.getElementById('input-email').value.trim();
    const password = document.getElementById('input-password').value;

    if (currentTab === 'register') {
      const name = document.getElementById('input-name').value.trim();
      if (!name) { showError('err-name', 'input-name'); valid = false; }
    }

    if (!email || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) {
      showError('err-email', 'input-email'); valid = false;
    }

    if (password.length < 6) {
      showError('err-password', 'input-password'); valid = false;
    }

    if (currentTab === 'register') {
      const confirm = document.getElementById('input-confirm').value;
      if (password !== confirm) {
        showError('err-confirm', 'input-confirm'); valid = false;
      }
    }

    return valid;
  }

  function showError(errId, inputId) {
    document.getElementById(errId).classList.add('show');
    document.getElementById(inputId).classList.add('error');
  }

  function clearErrors() {
    document.querySelectorAll('.error-msg').forEach(e => e.classList.remove('show'));
    document.querySelectorAll('.form-input').forEach(i => i.classList.remove('error'));
  }

  // =============================================
  // ALERT
  // =============================================
  function showAlert(msg, type = 'error') {
    const box = document.getElementById('alert-box');
    box.textContent = msg;
    box.className = `alert ${type} show`;
  }

  function hideAlert() {
    document.getElementById('alert-box').classList.remove('show');
  }

  // =============================================
  // LOADING STATE
  // =============================================
  function setLoading(loading) {
    const btn = document.getElementById('btn-submit');
    btn.disabled = loading;
    btn.classList.toggle('loading', loading);
  }

  // =============================================
  // HANDLE SUBMIT
  // =============================================
  async function handleSubmit() {
    if (!validateForm()) return;

    const email = document.getElementById('input-email').value.trim();
    const password = document.getElementById('input-password').value;
    hideAlert();
    setLoading(true);

    try {
      if (currentTab === 'login') {
        // LOGIN
        const data = await supabaseRequest('token?grant_type=password', { email, password });

        if (data.error || data.error_description) {
          const msg = data.error_description || data.error;
          if (msg.includes('Invalid login')) showAlert('Email atau password salah. Coba lagi.');
          else if (msg.includes('Email not confirmed')) showAlert('Email belum diverifikasi. Cek inbox kamu.');
          else showAlert(msg);
          setLoading(false);
          return;
        }

        // Simpan session
        localStorage.setItem('nova_session', JSON.stringify(data));
        showSuccess('login');

      } else {
        // REGISTER
        const name = document.getElementById('input-name').value.trim();
        const data = await supabaseRequest('signup', {
          email, password,
          data: { full_name: name }
        });

        if (data.error || data.error_description) {
          const msg = data.error_description || data.error;
          if (msg.includes('already registered')) showAlert('Email ini sudah terdaftar. Coba masuk.');
          else showAlert(msg);
          setLoading(false);
          return;
        }

        // Cek apakah perlu verifikasi email
        if (data.identities && data.identities.length === 0) {
          showAlert('Email ini sudah terdaftar. Silakan masuk.', 'error');
          setLoading(false);
          return;
        }

        if (!data.session) {
          // Perlu verifikasi email
          showSuccess('verify');
        } else {
          localStorage.setItem('nova_session', JSON.stringify(data.session));
          showSuccess('register');
        }
      }
    } catch (err) {
      showAlert('Terjadi kesalahan. Periksa koneksi internet kamu.');
      setLoading(false);
    }
  }

  // =============================================
  // GOOGLE LOGIN
  // =============================================
  async function handleGoogle() {
    const res = await fetch(`${SUPABASE_URL}/auth/v1/authorize?provider=google&redirect_to=${encodeURIComponent(window.location.origin + '/index.html')}`, {
      headers: { 'apikey': SUPABASE_ANON_KEY }
    });
    window.location.href = `${SUPABASE_URL}/auth/v1/authorize?provider=google&redirect_to=${encodeURIComponent(window.location.origin + '/index.html')}`;
  }

  // =============================================
  // SUCCESS SCREEN
  // =============================================
  function showSuccess(type) {
    document.getElementById('auth-form').style.display = 'none';
    document.querySelector('.login-header').style.display = 'none';
    document.querySelector('.auth-tabs').style.display = 'none';

    const screen = document.getElementById('success-screen');
    screen.classList.add('show');

    if (type === 'verify') {
      document.getElementById('success-title').textContent = 'CEK EMAIL!';
      document.getElementById('success-msg').innerHTML =
        'Link verifikasi sudah dikirim ke email kamu.<br>Klik link tersebut untuk mengaktifkan akun.';
      screen.querySelector('.redirect-bar').style.display = 'none';
    } else {
      document.getElementById('success-title').textContent = 'BERHASIL!';
      document.getElementById('success-msg').textContent =
        type === 'login' ? 'Login berhasil. Mengalihkan ke halaman utama...' : 'Akun berhasil dibuat. Mengalihkan...';

      // Redirect ke index.html setelah 2 detik
      setTimeout(() => {
        window.location.href = 'index.html';
      }, 2000);
    }
  }

  // =============================================
  // CEK SESSION — jika sudah login, langsung redirect
  // =============================================
  function checkSession() {
    const session = localStorage.getItem('nova_session');
    if (session) {
      try {
        const parsed = JSON.parse(session);
        const expiresAt = parsed.expires_at;
        if (expiresAt && Date.now() / 1000 < expiresAt) {
          window.location.href = 'index.html';
        }
      } catch(e) {
        localStorage.removeItem('nova_session');
      }
    }
  }
  checkSession();

  // =============================================
  // PASSWORD TOGGLE
  // =============================================
  function togglePassword(inputId, btn) {
    const input = document.getElementById(inputId);
    const isText = input.type === 'text';
    input.type = isText ? 'password' : 'text';
    btn.textContent = isText ? '👁' : '🙈';
  }

  // =============================================
  // CURSOR
  // =============================================
  const cursor = document.getElementById('cursor');
  const ring = document.getElementById('cursorRing');
  let mx = 0, my = 0, rx = 0, ry = 0;
  document.addEventListener('mousemove', e => {
    mx = e.clientX; my = e.clientY;
    cursor.style.transform = `translate(${mx - 5}px, ${my - 5}px)`;
  });
  function animRing() {
    rx += (mx - rx) * 0.12; ry += (my - ry) * 0.12;
    ring.style.transform = `translate(${rx - 18}px, ${ry - 18}px)`;
    requestAnimationFrame(animRing);
  }
  animRing();
  document.querySelectorAll('button, a, input').forEach(el => {
    el.addEventListener('mouseenter', () => cursor.classList.add('hover'));
    el.addEventListener('mouseleave', () => cursor.classList.remove('hover'));
  });

  // Enter key submit
  document.addEventListener('keydown', e => {
    if (e.key === 'Enter') handleSubmit();
  });
</script>
</body>
</html>
