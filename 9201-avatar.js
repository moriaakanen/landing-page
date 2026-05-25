/**
 * 9201 AVATAR
 * Shared profile photo rendering, crop/compress editor, and Supabase upload.
 */
(function () {
  'use strict';

  const BUCKET = 'foto-profil';
  const OUTPUT_SIZE = 512;
  const JPEG_QUALITY = 0.86;
  const MAX_FILE_MB = 8;

  let editorState = null;

  function getPhotoUrl(row) {
    if (!row) return '';
    return row.foto_url || row.photo_url || row.avatar_url || row.profile_photo_url || '';
  }

  function initialsOf(name) {
    return String(name || '')
      .split(/\s+/)
      .filter(Boolean)
      .map(w => w[0])
      .join('')
      .toUpperCase()
      .slice(0, 2) || '-';
  }

  function escHtml(value) {
    if (typeof window.esc === 'function') return window.esc(value);
    return String(value ?? '').replace(/[&<>"']/g, c =>
      ({ '&':'&amp;', '<':'&lt;', '>':'&gt;', '"':'&quot;', "'":'&#39;' }[c])
    );
  }

  function escAttr(value) {
    if (typeof window.escAttr === 'function') return window.escAttr(value);
    return escHtml(value);
  }

  function renderAvatar(row, name, className, attrs) {
    const url = getPhotoUrl(row);
    const label = name || row?.NAMA || row?.full_name || row?.username || 'Pengguna';
    const extra = attrs ? ' ' + attrs : '';
    if (url) {
      return `<span class="${escAttr(className || 'avatar-photo')}"${extra}><img src="${escAttr(url)}" alt="${escAttr(label)}" loading="lazy" decoding="async"></span>`;
    }
    return `<span class="${escAttr(className || 'avatar-photo')}"${extra}>${escHtml(initialsOf(label))}</span>`;
  }

  function setAvatarElement(el, row, name) {
    if (!el) return;
    const url = getPhotoUrl(row);
    const label = name || row?.NAMA || row?.full_name || row?.username || 'Pengguna';
    el.innerHTML = '';
    el.classList.toggle('has-photo', !!url);
    if (url) {
      const img = document.createElement('img');
      img.src = url;
      img.alt = label;
      img.loading = 'lazy';
      img.decoding = 'async';
      el.appendChild(img);
    } else {
      el.textContent = initialsOf(label);
    }
  }

  function storagePublicUrl(path) {
    return `${SUPABASE_URL}/storage/v1/object/public/${BUCKET}/${path}`;
  }

  async function fetchJson(url, opts) {
    const res = await fetch(url, opts);
    if (!res.ok) {
      let msg = `HTTP ${res.status}`;
      try {
        const err = await res.json();
        msg = err.message || err.error || msg;
      } catch (_) {}
      throw new Error(msg);
    }
    if (res.status === 204) return null;
    const text = await res.text();
    return text ? JSON.parse(text) : null;
  }

  async function resolveSessionPhoto(session) {
    if (!session || !window.SUPABASE_URL || !window.SUPABASE_HEADERS) return null;
    if (getPhotoUrl(session)) return session;

    const tryFetch = async (query) => {
      const rows = await fetchJson(
        `${SUPABASE_URL}/rest/v1/data_pegawai?${query}&select=*&limit=1`,
        { headers: SUPABASE_HEADERS }
      ).catch(() => []);
      return Array.isArray(rows) && rows.length ? rows[0] : null;
    };

    let row = null;
    if (session.id != null) row = await tryFetch(`id=eq.${encodeURIComponent(session.id)}`);
    if (!row && session.username) row = await tryFetch(`NIP=eq.${encodeURIComponent(session.username)}`);
    if (!row) return null;

    const url = getPhotoUrl(row);
    if (url) {
      session.foto_url = url;
      try { localStorage.setItem('nova_user', JSON.stringify(session)); } catch (_) {}
    }
    return row;
  }

  async function hydrateTopbar(session) {
    const el = document.getElementById('topbar-avatar');
    if (!el || !session) return;
    setAvatarElement(el, session, session.full_name || session.username || 'Pengguna');
    const row = await resolveSessionPhoto(session).catch(() => null);
    if (row) setAvatarElement(el, row, session.full_name || row.NAMA || session.username || 'Pengguna');
  }

  function ensureCss() {
    if (document.getElementById('9201-avatar-css')) return;
    const style = document.createElement('style');
    style.id = '9201-avatar-css';
    style.textContent = `
      .topbar-avatar img,.hero-avatar img,.pg-avatar img,.pg-detail-avatar img,.pak-person-avatar img,.user-avatar-sm img,.profil-modal-avatar img{width:100%;height:100%;object-fit:cover;border-radius:inherit;display:block}
      .avatar-edit-wrap{position:relative;display:inline-flex}
      .avatar-edit-btn{position:absolute;right:2px;bottom:2px;width:28px;height:28px;border-radius:50%;border:2px solid #fff;background:var(--gold,#c8a84b);color:#fff;display:flex;align-items:center;justify-content:center;cursor:pointer;box-shadow:0 4px 12px rgba(13,35,64,.22);font-size:0;transition:transform .15s,background .15s}
      .avatar-edit-btn::before{content:'';width:14px;height:10px;border:1.8px solid currentColor;border-radius:3px;display:block;box-sizing:border-box}
      .avatar-edit-btn::after{content:'';position:absolute;width:4px;height:4px;border:1.7px solid currentColor;border-radius:50%;left:50%;top:50%;transform:translate(-50%,-37%)}
      .avatar-edit-btn:hover{background:var(--gold-dark,#a78a3a);transform:translateY(-1px)}
      .avatar-edit-btn:focus-visible{outline:none;box-shadow:0 0 0 3px rgba(200,168,75,.35),0 4px 12px rgba(13,35,64,.22)}
      .avatar-admin-btn{height:34px;padding:0 12px;border-radius:8px;border:1px solid rgba(200,168,75,.38);background:#fffdf6;color:var(--gold-dark,#a78a3a);font-family:inherit;font-size:12px;font-weight:700;cursor:pointer;display:inline-flex;align-items:center;gap:7px;transition:background .15s,border-color .15s}
      .avatar-admin-btn:hover{background:#faf4e3;border-color:var(--gold,#c8a84b)}
      .avatar-modal-overlay{position:fixed;inset:0;z-index:900;background:rgba(13,35,64,.58);display:none;align-items:center;justify-content:center;padding:22px;backdrop-filter:blur(3px)}
      .avatar-modal-overlay.show{display:flex}
      .avatar-modal{width:min(720px,100%);max-height:calc(100vh - 44px);background:#fff;border-radius:14px;box-shadow:0 24px 70px rgba(13,35,64,.28);display:flex;flex-direction:column;overflow:hidden}
      .avatar-modal-head{padding:18px 22px;border-bottom:1px solid var(--border,#e2ddd6);display:flex;align-items:center;gap:12px}
      .avatar-modal-title{font-family:'Fraunces',serif;font-size:20px;font-weight:500;color:var(--navy,#0d2340);flex:1}
      .avatar-modal-title em{font-style:italic;color:var(--gold,#c8a84b)}
      .avatar-modal-close{width:32px;height:32px;border:none;border-radius:50%;background:transparent;color:var(--muted,#6b7280);cursor:pointer;font-size:18px}
      .avatar-modal-close:hover{background:var(--bg,#f5f4f0);color:var(--navy,#0d2340)}
      .avatar-modal-body{padding:20px 22px;overflow:auto;display:grid;grid-template-columns:minmax(260px,320px) 1fr;gap:20px;align-items:start}
      .avatar-preview-shell{background:#fafaf6;border:1px solid var(--border,#e2ddd6);border-radius:12px;padding:16px}
      .avatar-crop-stage{width:100%;aspect-ratio:1;border-radius:50%;background:linear-gradient(135deg,var(--navy,#0d2340),var(--navy2,#163358));position:relative;overflow:hidden;touch-action:none;cursor:grab;box-shadow:inset 0 0 0 1px rgba(255,255,255,.28)}
      .avatar-crop-stage:active{cursor:grabbing}
      .avatar-crop-stage img{position:absolute;left:50%;top:50%;max-width:none;user-select:none;-webkit-user-drag:none;transform-origin:center center}
      .avatar-crop-empty{position:absolute;inset:0;display:flex;align-items:center;justify-content:center;color:var(--gold,#c8a84b);font-family:'Fraunces',serif;font-size:46px;font-weight:600}
      .avatar-crop-ring{position:absolute;inset:0;border-radius:50%;box-shadow:inset 0 0 0 2px rgba(255,255,255,.9),inset 0 0 0 999px rgba(13,35,64,.08);pointer-events:none}
      .avatar-tools{display:flex;flex-direction:column;gap:14px}
      .avatar-drop{border:1.5px dashed var(--border2,#d1ccc3);border-radius:12px;background:#fffdf8;padding:18px;text-align:center;color:var(--navy,#0d2340);cursor:pointer;transition:border-color .15s,background .15s}
      .avatar-drop:hover{border-color:var(--gold,#c8a84b);background:#fff8e6}
      .avatar-drop strong{display:block;font-size:13px;margin-bottom:5px}
      .avatar-drop span{display:block;font-size:11.5px;color:var(--muted,#6b7280);line-height:1.5}
      .avatar-file{display:none}
      .avatar-control label{display:block;font-size:11px;font-weight:800;letter-spacing:.7px;text-transform:uppercase;color:var(--muted,#6b7280);margin-bottom:7px}
      .avatar-control input[type="range"]{width:100%;accent-color:var(--gold,#c8a84b)}
      .avatar-help{font-size:12px;line-height:1.65;color:var(--muted,#6b7280);background:#f8fafc;border:1px solid #e2e8f0;border-radius:10px;padding:11px 12px}
      .avatar-alert{display:none;border-radius:9px;padding:10px 12px;font-size:12px;line-height:1.5}
      .avatar-alert.show{display:block}
      .avatar-alert.error{background:#fef2f2;border:1px solid #fecaca;color:#991b1b}
      .avatar-alert.success{background:#ecfdf5;border:1px solid #bbf7d0;color:#166534}
      .avatar-modal-foot{padding:14px 22px;border-top:1px solid var(--border,#e2ddd6);background:#fafaf6;display:flex;justify-content:space-between;gap:10px;align-items:center}
      .avatar-modal-foot small{color:var(--muted,#6b7280);font-size:11px}
      .avatar-actions{display:flex;gap:10px}
      .avatar-btn{padding:9px 15px;border-radius:8px;border:1px solid var(--border,#e2ddd6);font-family:inherit;font-size:13px;font-weight:700;cursor:pointer;background:#fff;color:var(--navy,#0d2340)}
      .avatar-btn.primary{background:var(--navy,#0d2340);border-color:var(--navy,#0d2340);color:#fff}
      .avatar-btn:disabled{opacity:.55;cursor:not-allowed}
      @media(max-width:680px){.avatar-modal-body{grid-template-columns:1fr}.avatar-modal-foot{align-items:stretch;flex-direction:column}.avatar-actions{justify-content:flex-end}}
    `;
    document.head.appendChild(style);
  }

  function ensureModal() {
    ensureCss();
    let overlay = document.getElementById('avatar-modal');
    if (overlay) return overlay;
    overlay = document.createElement('div');
    overlay.id = 'avatar-modal';
    overlay.className = 'avatar-modal-overlay';
    overlay.innerHTML = `
      <div class="avatar-modal" role="dialog" aria-modal="true" aria-labelledby="avatar-modal-title">
        <div class="avatar-modal-head">
          <div class="avatar-modal-title" id="avatar-modal-title">Ganti <em>Foto</em> Profil</div>
          <button type="button" class="avatar-modal-close" id="avatar-close" aria-label="Tutup">x</button>
        </div>
        <div class="avatar-modal-body">
          <div class="avatar-preview-shell">
            <div class="avatar-crop-stage" id="avatar-stage">
              <div class="avatar-crop-empty" id="avatar-empty">-</div>
              <img id="avatar-img" alt="" hidden>
              <div class="avatar-crop-ring"></div>
            </div>
          </div>
          <div class="avatar-tools">
            <label class="avatar-drop" for="avatar-file">
              <strong>Pilih foto dari perangkat</strong>
              <span>Gunakan foto wajah yang jelas. File besar akan dikompres otomatis menjadi 512 x 512 px.</span>
            </label>
            <input class="avatar-file" id="avatar-file" type="file" accept="image/png,image/jpeg,image/webp">
            <div class="avatar-control">
              <label for="avatar-zoom">Zoom</label>
              <input id="avatar-zoom" type="range" min="1" max="3" step="0.01" value="1">
            </div>
            <div class="avatar-help">Geser foto di area lingkaran untuk memilih bagian yang tampil. Rekomendasi foto asli minimal 800 x 800 px; sistem akan crop dan kompres otomatis.</div>
            <div class="avatar-alert" id="avatar-alert"></div>
          </div>
        </div>
        <div class="avatar-modal-foot">
          <small>Output akhir: JPEG 512 x 512 px, kualitas 86%.</small>
          <div class="avatar-actions">
            <button type="button" class="avatar-btn" id="avatar-cancel">Batal</button>
            <button type="button" class="avatar-btn primary" id="avatar-save" disabled>Simpan Foto</button>
          </div>
        </div>
      </div>
    `;
    document.body.appendChild(overlay);
    overlay.addEventListener('click', e => { if (e.target === overlay) closeEditor(); });
    overlay.querySelector('#avatar-close').addEventListener('click', closeEditor);
    overlay.querySelector('#avatar-cancel').addEventListener('click', closeEditor);
    overlay.querySelector('#avatar-file').addEventListener('change', onFileChange);
    overlay.querySelector('#avatar-zoom').addEventListener('input', onZoomChange);
    overlay.querySelector('#avatar-save').addEventListener('click', saveEditor);

    const stage = overlay.querySelector('#avatar-stage');
    stage.addEventListener('pointerdown', onPointerDown);
    stage.addEventListener('pointermove', onPointerMove);
    stage.addEventListener('pointerup', onPointerUp);
    stage.addEventListener('pointercancel', onPointerUp);
    return overlay;
  }

  function showAlert(message, type) {
    const el = document.getElementById('avatar-alert');
    if (!el) return;
    el.className = `avatar-alert show ${type || 'error'}`;
    el.textContent = message;
  }

  function clearAlert() {
    const el = document.getElementById('avatar-alert');
    if (!el) return;
    el.className = 'avatar-alert';
    el.textContent = '';
  }

  function openEditor(opts) {
    const overlay = ensureModal();
    const person = opts?.person || {};
    const name = opts?.name || person.NAMA || person.full_name || person.username || 'Pegawai';
    editorState = {
      nip: opts?.nip || person.NIP || person.username || '',
      name,
      person,
      onSaved: typeof opts?.onSaved === 'function' ? opts.onSaved : null,
      img: null,
      objectUrl: '',
      baseScale: 1,
      zoom: 1,
      x: 0,
      y: 0,
      dragging: false,
      dragStart: null,
    };
    overlay.querySelector('#avatar-modal-title').innerHTML = `Ganti <em>Foto</em> Profil`;
    overlay.querySelector('#avatar-empty').textContent = initialsOf(name);
    overlay.querySelector('#avatar-img').hidden = true;
    overlay.querySelector('#avatar-img').removeAttribute('src');
    overlay.querySelector('#avatar-file').value = '';
    overlay.querySelector('#avatar-zoom').value = '1';
    overlay.querySelector('#avatar-save').disabled = true;
    clearAlert();
    overlay.classList.add('show');
    document.body.style.overflow = 'hidden';
  }

  function closeEditor() {
    const overlay = document.getElementById('avatar-modal');
    if (overlay) overlay.classList.remove('show');
    if (editorState?.objectUrl) URL.revokeObjectURL(editorState.objectUrl);
    editorState = null;
    document.body.style.overflow = '';
  }

  function onFileChange(e) {
    const file = e.target.files && e.target.files[0];
    if (!file || !editorState) return;
    clearAlert();
    if (!/^image\/(jpeg|png|webp)$/.test(file.type)) {
      showAlert('Format foto harus JPG, PNG, atau WebP.');
      return;
    }
    if (file.size > MAX_FILE_MB * 1024 * 1024) {
      showAlert(`Ukuran file maksimal ${MAX_FILE_MB} MB sebelum kompresi.`);
      return;
    }
    if (editorState.objectUrl) URL.revokeObjectURL(editorState.objectUrl);
    const url = URL.createObjectURL(file);
    const img = new Image();
    img.onload = () => {
      editorState.objectUrl = url;
      editorState.img = img;
      editorState.zoom = 1;
      editorState.x = 0;
      editorState.y = 0;
      document.getElementById('avatar-zoom').value = '1';
      document.getElementById('avatar-save').disabled = false;
      document.getElementById('avatar-empty').style.display = 'none';
      const imgEl = document.getElementById('avatar-img');
      imgEl.src = url;
      imgEl.hidden = false;
      computeBaseScale();
      renderCrop();
    };
    img.onerror = () => showAlert('Foto tidak bisa dibaca. Coba file lain.');
    img.src = url;
  }

  function computeBaseScale() {
    const stage = document.getElementById('avatar-stage');
    if (!stage || !editorState?.img) return;
    const box = stage.getBoundingClientRect();
    editorState.baseScale = Math.max(box.width / editorState.img.naturalWidth, box.height / editorState.img.naturalHeight);
  }

  function clampPan() {
    const stage = document.getElementById('avatar-stage');
    if (!stage || !editorState?.img) return;
    const box = stage.getBoundingClientRect();
    const scale = editorState.baseScale * editorState.zoom;
    const w = editorState.img.naturalWidth * scale;
    const h = editorState.img.naturalHeight * scale;
    const maxX = Math.max(0, (w - box.width) / 2);
    const maxY = Math.max(0, (h - box.height) / 2);
    editorState.x = Math.min(maxX, Math.max(-maxX, editorState.x));
    editorState.y = Math.min(maxY, Math.max(-maxY, editorState.y));
  }

  function renderCrop() {
    if (!editorState?.img) return;
    computeBaseScale();
    clampPan();
    const imgEl = document.getElementById('avatar-img');
    const scale = editorState.baseScale * editorState.zoom;
    imgEl.style.width = `${editorState.img.naturalWidth * scale}px`;
    imgEl.style.height = `${editorState.img.naturalHeight * scale}px`;
    imgEl.style.transform = `translate(calc(-50% + ${editorState.x}px), calc(-50% + ${editorState.y}px))`;
  }

  function onZoomChange(e) {
    if (!editorState) return;
    editorState.zoom = Number(e.target.value) || 1;
    renderCrop();
  }

  function onPointerDown(e) {
    if (!editorState?.img) return;
    editorState.dragging = true;
    editorState.dragStart = { px: e.clientX, py: e.clientY, x: editorState.x, y: editorState.y };
    e.currentTarget.setPointerCapture(e.pointerId);
  }

  function onPointerMove(e) {
    if (!editorState?.dragging || !editorState.dragStart) return;
    editorState.x = editorState.dragStart.x + (e.clientX - editorState.dragStart.px);
    editorState.y = editorState.dragStart.y + (e.clientY - editorState.dragStart.py);
    renderCrop();
  }

  function onPointerUp() {
    if (!editorState) return;
    editorState.dragging = false;
    editorState.dragStart = null;
  }

  async function cropToBlob() {
    if (!editorState?.img) throw new Error('Pilih foto terlebih dahulu.');
    const stage = document.getElementById('avatar-stage');
    const box = stage.getBoundingClientRect();
    const scale = editorState.baseScale * editorState.zoom;
    const visibleLeft = (editorState.img.naturalWidth * scale - box.width) / 2 - editorState.x;
    const visibleTop = (editorState.img.naturalHeight * scale - box.height) / 2 - editorState.y;
    const sx = Math.max(0, visibleLeft / scale);
    const sy = Math.max(0, visibleTop / scale);
    const sw = Math.min(editorState.img.naturalWidth - sx, box.width / scale);
    const sh = Math.min(editorState.img.naturalHeight - sy, box.height / scale);

    const canvas = document.createElement('canvas');
    canvas.width = OUTPUT_SIZE;
    canvas.height = OUTPUT_SIZE;
    const ctx = canvas.getContext('2d');
    ctx.fillStyle = '#ffffff';
    ctx.fillRect(0, 0, OUTPUT_SIZE, OUTPUT_SIZE);
    ctx.drawImage(editorState.img, sx, sy, sw, sh, 0, 0, OUTPUT_SIZE, OUTPUT_SIZE);
    return new Promise((resolve, reject) => {
      canvas.toBlob(blob => blob ? resolve(blob) : reject(new Error('Gagal membuat file foto.')), 'image/jpeg', JPEG_QUALITY);
    });
  }

  async function uploadAvatar(blob, nip) {
    if (!window.SUPABASE_URL || !window.SUPABASE_ANON_KEY) throw new Error('Konfigurasi Supabase belum tersedia.');
    const safeNip = String(nip || 'pegawai').replace(/[^\w-]/g, '_');
    const path = `${safeNip}/avatar-${Date.now()}.jpg`;
    const res = await fetch(`${SUPABASE_URL}/storage/v1/object/${BUCKET}/${path}`, {
      method: 'POST',
      headers: {
        apikey: SUPABASE_ANON_KEY,
        Authorization: `Bearer ${SUPABASE_ANON_KEY}`,
        'Content-Type': 'image/jpeg',
        'x-upsert': 'true',
      },
      body: blob,
    });
    if (!res.ok) {
      let msg = `HTTP ${res.status}`;
      try {
        const err = await res.json();
        msg = err.message || err.error || msg;
      } catch (_) {}
      throw new Error(`Upload foto gagal: ${msg}`);
    }
    return storagePublicUrl(path);
  }

  async function updatePegawaiPhoto(nip, url) {
    if (!nip) throw new Error('NIP pegawai tidak ditemukan.');
    const rpcRes = await fetch(`${SUPABASE_URL}/rest/v1/rpc/update_pegawai_foto`, {
      method: 'POST',
      headers: SUPABASE_HEADERS,
      body: JSON.stringify({ p_nip: String(nip), p_foto_url: url }),
    }).catch(() => null);
    if (rpcRes && rpcRes.ok) {
      const payload = await rpcRes.json().catch(() => null);
      if (Array.isArray(payload)) return payload[0] || { NIP: nip, foto_url: url };
      return payload || { NIP: nip, foto_url: url };
    }

    const rows = await fetchJson(
      `${SUPABASE_URL}/rest/v1/data_pegawai?NIP=eq.${encodeURIComponent(nip)}&select=*`,
      {
        method: 'PATCH',
        headers: { ...SUPABASE_HEADERS, Prefer: 'return=representation' },
        body: JSON.stringify({ foto_url: url }),
      }
    );
    return Array.isArray(rows) && rows.length ? rows[0] : { NIP: nip, foto_url: url };
  }

  async function saveEditor() {
    if (!editorState) return;
    const btn = document.getElementById('avatar-save');
    btn.disabled = true;
    btn.textContent = 'Menyimpan...';
    clearAlert();
    try {
      const blob = await cropToBlob();
      const url = await uploadAvatar(blob, editorState.nip);
      const row = await updatePegawaiPhoto(editorState.nip, url);
      showAlert('Foto profil berhasil diperbarui.', 'success');
      if (editorState.onSaved) editorState.onSaved(row, url);
      setTimeout(closeEditor, 450);
    } catch (e) {
      showAlert(e.message || 'Gagal menyimpan foto profil.');
      btn.disabled = false;
    } finally {
      btn.textContent = 'Simpan Foto';
    }
  }

  window.Avatar9201 = {
    bucket: BUCKET,
    outputSize: OUTPUT_SIZE,
    getPhotoUrl,
    initialsOf,
    renderAvatar,
    setAvatarElement,
    hydrateTopbar,
    resolveSessionPhoto,
    openEditor,
  };

  ensureCss();
})();
