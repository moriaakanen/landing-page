/* ═══════════════════════════════════════════════════════════════════════
   PORTAL NOVA — Admin Surat Tugas (Excel-like, multi-kolom)
   ─────────────────────────────────────────────────────────────────────
   Ketergantungan global:
     - SUPABASE_URL, SUPABASE_ANON_KEY (config.js)
     - getUserRoles, ADMIN_USERS (config.js)
     - initRoleSwitcher, toggleUserDropdown, switchViewRole (nova-role-switcher.js)
     - LOGO_BPS_BASE64 (logo-bps.js, opsional)
     - window.docxGen, window.docxPreview, saveAs (libs eksternal)
═══════════════════════════════════════════════════════════════════════ */

const H = {
  'apikey': SUPABASE_ANON_KEY,
  'Authorization': `Bearer ${SUPABASE_ANON_KEY}`,
  'Content-Type': 'application/json'
};

/* ═════════════════════════════════════════════════════════════════════
   TIPE SURAT — fitur multi-template
   ─────────────────────────────────────────────────────────────────────
   7 pilihan tipe yg menentukan template .docx yg dipakai saat generate.
   String value sengaja sama dengan check-constraint surat_tugas_tipe_enum
   di DB (lihat surat_tugas_add_tipe.sql).

   Ada 3 file template di Supabase Storage:
     T1: template-surat-tugas-spd-kendaraan-menginap.docx
         — punya {#has_spd}, {#kendaraan}, {#menginap}
     T2: template-surat-tugas-lampiran.docx
         — punya {#ul} (tabel lampiran)
     T3: template-surat-tugas-lampiran-spd-kendaraan-menginap.docx
         — punya {#ul} (lampiran), {#kendaraan}, {#menginap}; SPD selalu
           render (tidak ada {#has_spd} wrapper)

   Catatan tag yg perlu Anda pastikan ADA di template:
     - T1 & T3: blok kendaraan dibungkus {#kendaraan}...{/kendaraan}
                blok menginap   dibungkus {#menginap}...{/menginap}
                (sebelumnya pakai {#ul} dua kali — perlu di-rename)
═══════════════════════════════════════════════════════════════════════ */

// 7 pilihan tipe (urutan = urutan tampil di dropdown UI).
const TIPE_OPTIONS = [
  { value: 'surat_tugas',                                   label: 'Surat Tugas' },
  { value: 'surat_tugas_kendaraan',                         label: 'Surat Tugas + Kendaraan' },
  { value: 'surat_tugas_lampiran',                          label: 'Surat Tugas + Lampiran' },
  { value: 'surat_tugas_spd_kendaraan',                     label: 'Surat Tugas + SPD + Kendaraan' },
  { value: 'surat_tugas_spd_kendaraan_menginap',            label: 'Surat Tugas + SPD + Kendaraan + Menginap' },
  { value: 'surat_tugas_lampiran_spd_kendaraan',            label: 'Surat Tugas + Lampiran + SPD + Kendaraan' },
  { value: 'surat_tugas_lampiran_spd_kendaraan_menginap',   label: 'Surat Tugas + Lampiran + SPD + Kendaraan + Menginap' },
];

// URL template di Supabase Storage. Pastikan nama file yang Anda upload
// di bucket `template/` sama persis dengan path di sini.
const TEMPLATE_URL_T1 = 'https://jsmmtqeoukkgugorrvmg.supabase.co/storage/v1/object/public/template/template-surat-tugas-spd-kendaraan-menginap.docx';
const TEMPLATE_URL_T2 = 'https://jsmmtqeoukkgugorrvmg.supabase.co/storage/v1/object/public/template/template-surat-tugas-lampiran.docx';
const TEMPLATE_URL_T3 = 'https://jsmmtqeoukkgugorrvmg.supabase.co/storage/v1/object/public/template/template-surat-tugas-lampiran-spd-kendaraan-menginap.docx';

// Mapping tipe → template URL.
const TIPE_TO_TEMPLATE = {
  'surat_tugas':                                  TEMPLATE_URL_T1,
  'surat_tugas_kendaraan':                        TEMPLATE_URL_T1,
  'surat_tugas_lampiran':                         TEMPLATE_URL_T2,
  'surat_tugas_spd_kendaraan':                    TEMPLATE_URL_T1,
  'surat_tugas_spd_kendaraan_menginap':           TEMPLATE_URL_T1,
  'surat_tugas_lampiran_spd_kendaraan':           TEMPLATE_URL_T3,
  'surat_tugas_lampiran_spd_kendaraan_menginap':  TEMPLATE_URL_T3,
};

// Mapping tipe → flags untuk kontrol section di template.
//   has_spd       → toggle {#has_spd}...{/has_spd}  (hanya berpengaruh di T1)
//   has_kendaraan → kalau true, kirim array `kendaraan` (T1 + T3)
//   has_menginap  → kalau true, kirim array `menginap`  (T1 + T3)
//   has_lampiran  → kalau true, kirim array `ul` utk tabel lampiran (T2 + T3)
//
// Catatan: untuk T3, has_spd dikirim true walau template T3 tidak punya
// {#has_spd} wrapper — engine docxtemplater akan ignore tag yg tidak ada
// di template, jadi aman.
const TIPE_TO_FLAGS = {
  'surat_tugas':                                  { has_spd:false, has_kendaraan:false, has_menginap:false, has_lampiran:false },
  'surat_tugas_kendaraan':                        { has_spd:false, has_kendaraan:true,  has_menginap:false, has_lampiran:false },
  'surat_tugas_lampiran':                         { has_spd:false, has_kendaraan:false, has_menginap:false, has_lampiran:true  },
  'surat_tugas_spd_kendaraan':                    { has_spd:true,  has_kendaraan:true,  has_menginap:false, has_lampiran:false },
  'surat_tugas_spd_kendaraan_menginap':           { has_spd:true,  has_kendaraan:true,  has_menginap:true,  has_lampiran:false },
  'surat_tugas_lampiran_spd_kendaraan':           { has_spd:true,  has_kendaraan:true,  has_menginap:false, has_lampiran:true  },
  'surat_tugas_lampiran_spd_kendaraan_menginap':  { has_spd:true,  has_kendaraan:true,  has_menginap:true,  has_lampiran:true  },
};

// Helper: ubah enum value ke label readable. Kalau value tidak dikenal,
// kembalikan apa adanya (untuk debug).
function tipeLabel(tipe) {
  if (!tipe) return '';
  const opt = TIPE_OPTIONS.find(o => o.value === tipe);
  return opt ? opt.label : String(tipe);
}

// Helper: dapatkan flags dari tipe. Return objek dgn semua flag = false
// kalau tipe NULL/invalid (defensive — caller bisa langsung pakai).
function tipeFlags(tipe) {
  return TIPE_TO_FLAGS[tipe] || { has_spd:false, has_kendaraan:false, has_menginap:false, has_lampiran:false };
}

// Helper: dapatkan URL template dari tipe. Return null kalau tipe
// NULL/invalid — caller wajib handle (jangan render docx tanpa tipe valid).
function tipeTemplateUrl(tipe) {
  return TIPE_TO_TEMPLATE[tipe] || null;
}

let SESSION = null;
let allSurat = [];                // descending by created_at (untuk display di table)
let suratMap = {};                // id → object
let suratOrderMap = {};           // id → urutan global ascending (1-based)
let pegawaiList = [];             // dari "data pegawai"
let pegawaiByNIP = {};            // index NIP → object
let riwayatPegawai = [];          // dari "riwayat_pegawai"
let userMap = {};                 // id → { full_name, username } — untuk tampilkan nama pengaju surat
let selectedId = null;

// ─── DEFAULTS yang di-prefill saat baris status menunggu ──────────────
// v3 — format `pembebanan` berubah menjadi kode MAK terstruktur.
// Lihat parseMAK() untuk format yang valid. Bumping versi key agar
// default lama (yang berisi narasi "DIPA BPS...") tidak ter-prefill.
const APPROVE_DEFAULTS_KEY = 'nova_approve_defaults_v3';
const FACTORY_DEFAULTS = {
  alat_angkutan:  'Kendaraan Darat',
  // Format wajib: program(3 segmen).kegiatan.kro.ro.komponen.sub_komponen.akun
  // Contoh:       054.01.GG.2910.BMA.006.054.A.524119
  pembebanan:     '054.01.GG.2910.BMA.006.054.A.524119',
  ttd_nip:        '',
  ttd_nama:       '',
};
function loadApproveDefaults() {
  try {
    const raw = localStorage.getItem(APPROVE_DEFAULTS_KEY);
    if (!raw) return { ...FACTORY_DEFAULTS };
    return { ...FACTORY_DEFAULTS, ...JSON.parse(raw) };
  } catch(e) { return { ...FACTORY_DEFAULTS }; }
}
function saveApproveDefaults(d) {
  try { localStorage.setItem(APPROVE_DEFAULTS_KEY, JSON.stringify(d)); } catch(e) {}
}

/* ════════════════════════════════════════════════════════════════════
   SESSION & TOPBAR
═══════════════════════════════════════════════════════════════════════ */
function checkSession() {
  try {
    const s = JSON.parse(localStorage.getItem('nova_user') || 'null');
    if (!s) { window.location.replace('login.html'); return null; }
    if (s.expires_at && Date.now() > s.expires_at) {
      localStorage.removeItem('nova_user'); window.location.replace('login.html'); return null;
    }
    if (s.must_change_password) { window.location.replace('ganti-password.html'); return null; }
    const roles = getUserRoles(s);
    if (!s.active_role) {
      s.active_role = roles.includes('admin') ? 'admin' : 'user';
      localStorage.setItem('nova_user', JSON.stringify(s));
    }
    if (!roles.includes('admin') || s.active_role !== 'admin') {
      window.location.replace('index.html'); return null;
    }
    return s;
  } catch(e) {
    localStorage.removeItem('nova_user'); window.location.replace('login.html'); return null;
  }
}
function logout() { localStorage.removeItem('nova_user'); window.location.replace('login.html'); }
function updateClock() {
  document.getElementById('topbar-time').textContent =
    new Date().toLocaleString('id-ID', { weekday:'long', day:'numeric', month:'long', year:'numeric', hour:'2-digit', minute:'2-digit' });
}
updateClock(); setInterval(updateClock, 1000);
function setTopbarUser(s) {
  const name = s.full_name || s.username || 'Pengguna';
  const i = name.split(' ').filter(Boolean).map(w => w[0]).join('').toUpperCase().slice(0, 2);
  document.getElementById('topbar-avatar').textContent = i;
  document.getElementById('topbar-username').textContent = name;
}

/* ════════════════════════════════════════════════════════════════════
   HELPERS — escape, format tanggal, badge
═══════════════════════════════════════════════════════════════════════ */
function esc(str) {
  if (str==null) return '';
  return String(str).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;').replace(/'/g,'&#39;');
}
function escAttr(str) { return esc(str); }

const BULAN = ['Januari','Februari','Maret','April','Mei','Juni','Juli','Agustus','September','Oktober','November','Desember'];

function fmtTgl(str) {
  if (!str) return '';
  const d = parseISODate(str);
  if (!d) return '';
  return `${d.getDate()} ${BULAN[d.getMonth()]} ${d.getFullYear()}`;
}
function fmtTglDash(str) { return fmtTgl(str) || '—'; }

function parseISODate(str) {
  if (!str) return null;
  const m = String(str).match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (!m) return null;
  const d = new Date(+m[1], +m[2]-1, +m[3]);
  return isNaN(d) ? null : d;
}
function toISODate(d) {
  if (!d) return '';
  return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')}`;
}
function todayISO() { return toISODate(new Date()); }

function fmtWaktu(mulai, selesai) {
  if (!mulai) return '';
  if (!selesai || selesai === mulai) return fmtTgl(mulai);
  const a = parseISODate(mulai), b = parseISODate(selesai);
  if (!a || !b) return fmtTgl(mulai);
  if (a.getMonth()===b.getMonth() && a.getFullYear()===b.getFullYear())
    return `${a.getDate()} s.d. ${b.getDate()} ${BULAN[b.getMonth()]} ${b.getFullYear()}`;
  if (a.getFullYear()===b.getFullYear())
    return `${a.getDate()} ${BULAN[a.getMonth()]} s.d. ${b.getDate()} ${BULAN[b.getMonth()]} ${b.getFullYear()}`;
  return `${fmtTgl(mulai)} s.d. ${fmtTgl(selesai)}`;
}
function fmtWaktuDash(mulai, selesai) { return fmtWaktu(mulai, selesai) || '—'; }

function badgeHTML(status) {
  const m = { menunggu:['menunggu','⏳ Menunggu'], selesai:['selesai','✅ Selesai'] };
  const [cls, lbl] = m[status] || ['menunggu', status];
  return `<span class="badge ${cls}"><span class="badge-dot"></span>${lbl}</span>`;
}

/* ════════════════════════════════════════════════════════════════════
   PARSER TANGGAL FLEKSIBEL
═══════════════════════════════════════════════════════════════════════ */
const BULAN_KEYS = {
  'januari':1,'jan':1,'1':1,'01':1,
  'februari':2,'feb':2,'2':2,'02':2,
  'maret':3,'mar':3,'3':3,'03':3,
  'april':4,'apr':4,'4':4,'04':4,
  'mei':5,'5':5,'05':5,
  'juni':6,'jun':6,'6':6,'06':6,
  'juli':7,'jul':7,'7':7,'07':7,
  'agustus':8,'agu':8,'aug':8,'8':8,'08':8,
  'september':9,'sep':9,'sept':9,'9':9,'09':9,
  'oktober':10,'okt':10,'oct':10,'10':10,
  'november':11,'nov':11,'11':11,
  'desember':12,'des':12,'dec':12,'12':12,
};
function normYear(y) {
  y = parseInt(y, 10);
  if (isNaN(y)) return null;
  if (y < 100) return y < 50 ? 2000 + y : 1900 + y;
  return y;
}
function parseFlexDate(input) {
  if (!input) return null;
  let s = String(input).trim().toLowerCase();
  if (!s) return null;
  let m = s.match(/^(\d{4})[-\/\.](\d{1,2})[-\/\.](\d{1,2})$/);
  if (m) {
    const y = +m[1], mo = +m[2], d = +m[3];
    if (mo>=1&&mo<=12&&d>=1&&d<=31) return `${y}-${String(mo).padStart(2,'0')}-${String(d).padStart(2,'0')}`;
  }
  const parts = s.split(/[\/\-\.\s]+/).filter(Boolean);
  if (parts.length === 3) {
    let [d, mo, y] = parts;
    const moKey = BULAN_KEYS[mo];
    if (moKey) {
      d = parseInt(d, 10);
      const yy = normYear(y);
      if (d>=1&&d<=31 && yy) return `${yy}-${String(moKey).padStart(2,'0')}-${String(d).padStart(2,'0')}`;
    }
    d = parseInt(d, 10);
    mo = parseInt(mo, 10);
    const yy = normYear(y);
    if (d>=1&&d<=31&&mo>=1&&mo<=12&&yy) {
      return `${yy}-${String(mo).padStart(2,'0')}-${String(d).padStart(2,'0')}`;
    }
  }
  if (parts.length === 2) {
    let [d, mo] = parts;
    const moKey = BULAN_KEYS[mo] || (parseInt(mo,10) >= 1 && parseInt(mo,10) <= 12 ? parseInt(mo,10) : null);
    d = parseInt(d, 10);
    if (d>=1&&d<=31&&moKey) {
      const y = new Date().getFullYear();
      return `${y}-${String(moKey).padStart(2,'0')}-${String(d).padStart(2,'0')}`;
    }
  }
  return null;
}

function parseFlexRange(input) {
  if (!input) return { mulai: '', selesai: '' };
  let s = String(input).trim();
  if (!s) return { mulai: '', selesai: '' };
  const seps = [' s.d. ', ' sd ', ' - ', ' s/d ', ' to ', ' – ', ' — ', '~'];
  for (const sep of seps) {
    const idx = s.toLowerCase().indexOf(sep);
    if (idx > 0) {
      const left = s.slice(0, idx).trim();
      const right = s.slice(idx + sep.length).trim();
      const a = parseFlexDate(left);
      const b = parseFlexDate(right);
      if (a) return { mulai: a, selesai: b || '' };
    }
  }
  const dashSplit = s.split(/\s+-\s+|\s+–\s+|\s+—\s+/);
  if (dashSplit.length === 2) {
    const a = parseFlexDate(dashSplit[0]);
    const b = parseFlexDate(dashSplit[1]);
    if (a) return { mulai: a, selesai: b || '' };
  }
  const single = parseFlexDate(s);
  if (single) return { mulai: single, selesai: '' };
  return { mulai: '', selesai: '' };
}

/* ════════════════════════════════════════════════════════════════════
   LOAD DATA
═══════════════════════════════════════════════════════════════════════ */
async function loadPegawai() {
  try {
    const res = await fetch(`${SUPABASE_URL}/rest/v1/data%20pegawai?select=NIP,NAMA&order=NAMA.asc`, { headers: H });
    if (!res.ok) return;
    pegawaiList = await res.json();
    pegawaiByNIP = {};
    pegawaiList.forEach(p => {
      const nip = String(p.NIP || '').trim();
      if (nip) pegawaiByNIP[nip] = p;
    });
  } catch(e) { console.warn('Gagal load pegawai:', e); }
}

async function loadRiwayatPegawai() {
  try {
    const res = await fetch(`${SUPABASE_URL}/rest/v1/riwayat_pegawai?select=*&order=tmt.desc`, { headers: H });
    if (!res.ok) return;
    riwayatPegawai = await res.json();
  } catch(e) { console.warn('Gagal load riwayat_pegawai:', e); }
}

// Load daftar user → dipakai untuk tampilkan nama pengaju di kolom "Diajukan oleh"
async function loadUsers() {
  try {
    const res = await fetch(`${SUPABASE_URL}/rest/v1/users?select=id,username,full_name`, { headers: H });
    if (!res.ok) return;
    const list = await res.json();
    userMap = {};
    list.forEach(u => { userMap[u.id] = u; });
  } catch(e) { console.warn('Gagal load users:', e); }
}

function getPengajuNama(s) {
  const u = userMap[s.user_id];
  if (!u) return '';
  return u.full_name || u.username || '';
}

async function loadSurat() {
  document.getElementById('table-area').innerHTML = `<div style="padding:24px"><div class="skel" style="height:44px;border-radius:6px;margin-bottom:8px"></div><div class="skel" style="height:44px;border-radius:6px;margin-bottom:8px"></div><div class="skel" style="height:44px;border-radius:6px"></div></div>`;
  document.getElementById('table-count').textContent = 'Memuat...';
  try {
    const res = await fetch(`${SUPABASE_URL}/rest/v1/surat_tugas?select=*&order=created_at.asc`, { headers: H });
    if (!res.ok) throw new Error(`HTTP ${res.status}`);
    const ascList = await res.json();
    suratOrderMap = {};
    ascList.forEach((s, i) => { suratOrderMap[s.id] = i + 1; });
    allSurat = ascList.slice().reverse();
    suratMap = {};
    allSurat.forEach(s => { suratMap[s.id] = s; });
    updateStats();
    filterTable();
  } catch(e) {
    document.getElementById('table-area').innerHTML = `<div class="table-empty"><div class="table-empty-icon">⚠️</div><div class="table-empty-text">Gagal memuat data: ${esc(e.message)}</div></div>`;
    document.getElementById('table-count').textContent = 'Error';
  }
}

/* ════════════════════════════════════════════════════════════════════
   TAMBAH SURAT (Tugas #4) — admin bikin surat baru tanpa lewat user
   ─────────────────────────────────────────────────────────────────────
   Insert row baru ke tabel `surat_tugas` dengan:
     - user_id    = SESSION.id (admin sendiri sebagai pengaju)
     - status     = 'menunggu'
     - field2 lain kosong/default
   Setelah insert, reload tabel (dengan dirty-preserve agar edit di
   baris menunggu lain tidak hilang) lalu scroll & focus ke baris baru
   supaya admin bisa langsung mengisi.
═══════════════════════════════════════════════════════════════════════ */
async function addNewSurat() {
  if (!SESSION || !SESSION.id) {
    showPageAlert('Sesi tidak valid. Silakan refresh halaman.', 'error');
    return;
  }

  // Disable tombol selama proses agar tidak double-click
  const btn = document.querySelector('button[onclick="addNewSurat()"]');
  if (btn) { btn.disabled = true; btn.textContent = '… Menambahkan'; }

  // Payload minimal — sesuai pattern dari surat-tugas.html (user side).
  // Field-field detail (nomor_surat, penandatangan_*, tipe, dll.) di-NULL
  // dulu — admin akan mengisi via tabel inline lalu klik Setujui.
  const todayIso = todayISO();
  const payload = {
    user_id:           SESSION.id,
    perihal:           '',
    tanggal_berangkat: todayIso,
    tanggal_kembali:   null,
    tujuan:            '',
    pegawai_nip:       [],
    pegawai_list:      [],
    status:            'menunggu',
    created_at:        new Date().toISOString(),
  };

  try {
    const res = await fetch(`${SUPABASE_URL}/rest/v1/surat_tugas`, {
      method: 'POST',
      headers: { ...H, 'Prefer': 'return=representation' },
      body: JSON.stringify(payload),
    });
    if (!res.ok) {
      let msg = `HTTP ${res.status}`;
      try { const j = await res.json(); msg = j.message || j.hint || msg; } catch(_) {}
      throw new Error(msg);
    }
    const inserted = await res.json();
    const newId = Array.isArray(inserted) && inserted.length ? inserted[0].id : null;

    showPageAlert('✅ Baris baru ditambahkan. Lengkapi data lalu klik Setujui.', 'success');

    // Reload tabel — pakai pattern dirty-preserve yg sama dgn submitApprove
    // supaya edit di baris menunggu lain (kalau ada) tidak hilang.
    const dirtySnapshot = captureMenungguDirty(null);
    await loadSurat();
    reapplyMenungguDirty(dirtySnapshot);

    // Scroll & focus ke baris baru, isi sel pertama (perihal) supaya admin
    // bisa langsung ngetik.
    if (newId != null) {
      setTimeout(() => {
        const newRow = document.querySelector(`tr[data-surat-id="${newId}"]`);
        if (newRow) {
          newRow.scrollIntoView({ behavior: 'smooth', block: 'center' });
          // Highlight sebentar agar visible
          newRow.classList.add('row-focused');
          setTimeout(() => newRow.classList.remove('row-focused'), 2000);
          // Focus ke sel perihal (text input pertama di baris)
          const firstCell = newRow.querySelector('.xls-cell[data-field="perihal"], .xls-cell, .pg-input');
          if (firstCell) {
            firstCell.focus();
            if (firstCell.select) try { firstCell.select(); } catch(_) {}
          }
        }
      }, 50);
    }
  } catch(e) {
    showPageAlert(`Gagal menambahkan: ${e.message}`, 'error');
  } finally {
    if (btn) { btn.disabled = false; btn.textContent = '+ Tambah Surat'; }
  }
}

function updateStats() {
  document.getElementById('st-total').textContent    = allSurat.length;
  document.getElementById('st-menunggu').textContent = allSurat.filter(s => s.status === 'menunggu').length;
  document.getElementById('st-selesai').textContent  = allSurat.filter(s => s.status === 'selesai').length;
}

function filterTable() {
  const q  = document.getElementById('search-input').value.toLowerCase();
  const st = document.getElementById('filter-status').value;
  let f = allSurat;
  if (q)  f = f.filter(s => (s.perihal||'').toLowerCase().includes(q) || (s.tujuan||'').toLowerCase().includes(q));
  if (st) f = f.filter(s => s.status === st);
  f = sortData(f);
  renderTable(f);
}

/* ════════════════════════════════════════════════════════════════════
   SORTING
═══════════════════════════════════════════════════════════════════════ */
// Default: urut berdasarkan No. descending (terbaru di atas) —
// sama dengan perilaku sebelum fitur sort ditambahkan.
let sortState = { col: 'no', dir: 'desc' };

function sortData(arr) {
  const { col, dir } = sortState;
  const mul = dir === 'asc' ? 1 : -1;

  return arr.slice().sort((a, b) => {
    let va, vb;
    switch (col) {
      case 'no':
        va = suratOrderMap[a.id] || 0;
        vb = suratOrderMap[b.id] || 0;
        return (va - vb) * mul;

      case 'nomor_surat':
        va = (a.nomor_surat || '').toString();
        vb = (b.nomor_surat || '').toString();
        // numeric:true agar "008" < "010" sesuai angka, bukan string
        return va.localeCompare(vb, 'id', { numeric: true, sensitivity: 'base' }) * mul;

      case 'tanggal_surat':
        va = a.tanggal_surat || '';
        vb = b.tanggal_surat || '';
        return va.localeCompare(vb) * mul;

      case 'waktu':
        // Sort by tanggal_berangkat; tie-breaker: tanggal_kembali
        va = a.tanggal_berangkat || '';
        vb = b.tanggal_berangkat || '';
        if (va !== vb) return va.localeCompare(vb) * mul;
        va = a.tanggal_kembali || '';
        vb = b.tanggal_kembali || '';
        return va.localeCompare(vb) * mul;

      case 'perihal':
        va = (a.perihal || '').toLowerCase();
        vb = (b.perihal || '').toLowerCase();
        return va.localeCompare(vb, 'id') * mul;

      case 'tujuan':
        va = (a.tujuan || '').toLowerCase();
        vb = (b.tujuan || '').toLowerCase();
        return va.localeCompare(vb, 'id') * mul;

      case 'status': {
        // Urutan logis: menunggu → selesai
        const order = { menunggu: 0, selesai: 1 };
        va = order[a.status] ?? 99;
        vb = order[b.status] ?? 99;
        return (va - vb) * mul;
      }

      case 'pengaju':
        va = getPengajuNama(a).toLowerCase();
        vb = getPengajuNama(b).toLowerCase();
        return va.localeCompare(vb, 'id') * mul;

      default:
        return 0;
    }
  });
}

function setSort(col) {
  if (sortState.col === col) {
    sortState.dir = sortState.dir === 'asc' ? 'desc' : 'asc';
  } else {
    sortState.col = col;
    // Kolom yang "secara intuitif" lebih bermakna descending
    sortState.dir = ['no', 'tanggal_surat', 'waktu'].includes(col) ? 'desc' : 'asc';
  }
  filterTable();
}

function sortHeader(col, label, cssClass) {
  const isActive = sortState.col === col;
  const arrow = isActive ? (sortState.dir === 'asc' ? '▴' : '▾') : '↕';
  const activeCls = isActive ? ' sorted' : '';
  return `<th class="${cssClass} sortable${activeCls}" onclick="setSort('${col}')" title="Klik untuk mengurutkan">${label}<span class="sort-arrow">${arrow}</span></th>`;
}

/* ════════════════════════════════════════════════════════════════════
   RENDER TABLE
═══════════════════════════════════════════════════════════════════════ */

// State: id surat yang sedang di-edit (status selesai di-unlock secara
// inline). null kalau tidak ada baris yg sedang di-edit. Hanya boleh ada
// 1 baris dalam mode edit pada satu waktu agar UX tidak ambigu.
let editingRowId = null;

/* Build HTML untuk SATU <tr> berdasarkan object surat.
   Dipisah dari renderTable() supaya bisa dipakai juga oleh enableRowEdit()
   dan cancelRowEdit() — keduanya hanya perlu re-render 1 baris, bukan
   seluruh tabel. */
function renderRowHTML(s) {
  const isMenunggu  = s.status === 'menunggu';
  const isSelesai   = s.status === 'selesai';
  const isEditing   = editingRowId === s.id;

  // Field-field unlocked kalau: status menunggu, ATAU baris ini sedang di-edit.
  const editable = isMenunggu || isEditing;

  const todayStr  = todayISO();
  const urutNo    = suratOrderMap[s.id] || '—';
  // Prefill nomor surat untuk status menunggu = urutan No. di-pad 3 digit.
  // Mis. No. = 8 → "008". Admin tetap bisa mengedit via sel input.
  const urutNomor = (typeof urutNo === 'number') ? String(urutNo).padStart(3, '0') : '';

  const nomorSurat   = s.nomor_surat   || (isMenunggu ? urutNomor : '');
  const tanggalSurat = s.tanggal_surat || (isMenunggu ? todayStr  : '');
  const waktuMulai   = s.tanggal_berangkat || '';
  const waktuSelesai = s.tanggal_kembali   || '';
  const perihal      = s.perihal      || '';
  const tujuan       = s.tujuan       || '';
  // Kolom-kolom berikut TIDAK di-prefill untuk status menunggu —
  // admin mengisi manual saat approve. Prefill default disimpan terpisah
  // dan hanya diterapkan di modal approve.
  const menimbang    = s.menimbang_custom || '';
  const alat         = s.alat_angkutan || '';
  const mak          = s.pembebanan    || '';
  const ttdNama      = s.penandatangan_nama || '';
  const ttdNip       = s.penandatangan_nip  || '';

  const pegNips  = Array.isArray(s.pegawai_nip)  ? s.pegawai_nip  : [];
  const pegNames = Array.isArray(s.pegawai_list) ? s.pegawai_list : [];

  let aksi;
  if (isMenunggu) {
    aksi = `
      <button class="btn-approve" onclick="openApprove(${s.id})">✅ Setujui</button>`;
  } else if (isSelesai) {
    if (isEditing) {
      // Mode edit aktif — tampilkan Simpan & Batal, sembunyikan Preview/Download
      // supaya admin tidak preview docx dengan data yg belum dipersist.
      aksi = `
        <button class="btn-save-edit" onclick="saveRowEdit(${s.id})">💾 Simpan</button>
        <button class="btn-cancel-edit" onclick="cancelRowEdit(${s.id})">✕ Batal</button>`;
    } else {
      aksi = `
        <button class="btn-edit-row" onclick="enableRowEdit(${s.id})" title="Edit field surat ini">✏️ Edit</button>
        <button class="btn-preview" onclick="openPreview(${s.id})">👁 Preview</button>
        <button class="btn-download" onclick="downloadSuratTugas(${s.id})">📥</button>`;
    }
  } else {
    // Status tidak dikenal — tampilkan placeholder
    aksi = `<span style="font-size:11px;color:var(--muted);font-style:italic">—</span>`;
  }

  // Checkbox bulk-download — enabled hanya untuk row status 'selesai'
  // (yang siap di-generate ke .docx). Row 'menunggu' dapat checkbox
  // disabled supaya layout konsisten tapi tidak bisa dipilih.
  const checkCell = isSelesai
    ? `<td class="col-check"><input type="checkbox" class="bulk-dl-check" data-surat-id="${s.id}" onchange="updateBulkDownloadCounter()"></td>`
    : `<td class="col-check"><input type="checkbox" disabled title="Hanya surat yang sudah selesai bisa di-download"></td>`;

  // data-editing sebagai marker tambahan agar styling/CSS bisa membedakan
  // baris yg sedang di-edit (ditambah border kuning di seksi CSS).
  return `
    <tr data-surat-id="${s.id}" data-status="${s.status}"${isEditing ? ' data-editing="1"' : ''}>
      ${checkCell}
      <td class="col-no">${urutNo}</td>

      ${cellTextHTML(s.id, 'nomor_surat', nomorSurat, editable, 'cth: 001 / 013A')}
      ${cellDateHTML(s.id, 'tanggal_surat', tanggalSurat, editable, 'tgl/bln/thn')}
      ${cellDateRangeHTML(s.id, 'waktu', waktuMulai, waktuSelesai, editable)}
      ${cellTextareaHTML(s.id, 'perihal', perihal, editable, 'Perihal surat')}
      ${cellTextHTML(s.id, 'tujuan', tujuan, editable, 'Kota/instansi')}
      ${cellPegawaiMultiHTML(s.id, pegNips, pegNames, editable)}
      ${cellTextareaHTML(s.id, 'menimbang_custom', menimbang, editable, 'cth: pelaksanaan Survei...')}
      ${cellTextareaHTML(s.id, 'alat_angkutan', alat, editable, 'cth: Kendaraan Darat')}
      ${cellMAKHTML(s.id, mak, editable)}
      ${cellPenandatanganHTML(s.id, ttdNip, ttdNama, editable)}
      ${cellTipeHTML(s.id, s.tipe, editable)}

      <td class="col-status">${badgeHTML(s.status)}</td>
      <td class="col-aksi"><div class="aksi-wrap">${aksi}</div></td>
      <td class="col-pengaju" title="${esc(getPengajuNama(s))}">${esc(getPengajuNama(s)) || '<span style="color:var(--muted);font-style:italic">—</span>'}</td>
    </tr>`;
}

/* Pasang event listener (autoGrow + clear err) untuk semua input/textarea
   editable di dalam scope tertentu. Default scope = whole document.
   Dipakai oleh renderTable (semua baris) dan setelah re-render 1 baris. */
function attachEditableListeners(scope) {
  const root = scope || document;
  root.querySelectorAll('textarea.xls-cell').forEach(ta => {
    autoGrow(ta);
    ta.addEventListener('input', () => { autoGrow(ta); ta.classList.remove('err'); });
  });
  root.querySelectorAll('input.xls-cell').forEach(inp => {
    inp.addEventListener('input', () => inp.classList.remove('err'));
  });
}

function renderTable(data) {
  document.getElementById('table-count').textContent = `${data.length} surat`;
  if (!data.length) {
    document.getElementById('table-area').innerHTML = `<div class="table-empty"><div class="table-empty-icon">📭</div><div class="table-empty-text">Tidak ada surat tugas ditemukan.</div></div>`;
    return;
  }

  // Kalau baris yg sedang di-edit tidak ada di data (mis. ke-filter out),
  // reset state edit supaya tombol Edit muncul lagi saat baris kembali.
  if (editingRowId != null && !data.some(s => s.id === editingRowId)) {
    editingRowId = null;
  }

  const rows = data.map(renderRowHTML).join('');

  document.getElementById('table-area').innerHTML = `
    <table class="list-table">
      <thead><tr>
        <th class="col-check" title="Centang semua surat selesai untuk bulk download"><input type="checkbox" id="bulk-dl-master" onchange="toggleBulkDownloadAll(this.checked)"></th>
        ${sortHeader('no',            'No',                'col-no')}
        ${sortHeader('nomor_surat',   'Nomor Surat',       'col-nomor-surat')}
        ${sortHeader('tanggal_surat', 'Tgl Surat',         'col-tgl-surat')}
        ${sortHeader('waktu',         'Waktu Pelaksanaan', 'col-waktu')}
        ${sortHeader('perihal',       'Perihal',           'col-perihal')}
        ${sortHeader('tujuan',        'Tempat Tujuan',     'col-tujuan')}
        <th class="col-nama">Nama Pegawai</th>
        <th class="col-menimbang">Menimbang</th>
        <th class="col-alat">Alat Angkutan</th>
        <th class="col-mak">POK</th>
        <th class="col-ttd">Penandatangan</th>
        <th class="col-tipe">Tipe</th>
        ${sortHeader('status',        'Status',            'col-status')}
        <th class="col-aksi">Aksi</th>
        ${sortHeader('pengaju',       'Diajukan oleh',     'col-pengaju')}
      </tr></thead>
      <tbody>${rows}</tbody>
    </table>`;

  requestAnimationFrame(() => {
    attachEditableListeners();
    setupTopScrollbar();
    // Reset state bulk download (master + counter) setelah re-render
    updateBulkDownloadCounter();
  });
}

/* ════════════════════════════════════════════════════════════════════
   TOP SCROLLBAR — sinkron dengan scrollbar bawah
   Memungkinkan admin menggeser tabel horizontal tanpa harus scroll
   page ke baris paling bawah dulu.
═══════════════════════════════════════════════════════════════════════ */
function setupTopScrollbar() {
  const top      = document.getElementById('table-scroll-top');
  const topInner = document.getElementById('table-scroll-top-inner');
  const bottom   = document.getElementById('table-area');
  if (!top || !topInner || !bottom) return;

  const tableEl    = bottom.querySelector('table');
  const tableWidth = tableEl ? tableEl.scrollWidth : bottom.scrollWidth;
  topInner.style.width = tableWidth + 'px';

  // Sembunyikan top scrollbar kalau tabel tidak overflow horizontal
  // (mis. layar lebar / data sedikit kolomnya)
  if (tableWidth <= bottom.clientWidth) {
    top.style.display = 'none';
    return;
  }
  top.style.display = '';

  // Pasang sync listener satu kali saja (idempotent — pakai flag di element)
  if (!top._syncBound) {
    let syncing = false;
    top.addEventListener('scroll', () => {
      if (syncing) return;
      syncing = true;
      bottom.scrollLeft = top.scrollLeft;
      requestAnimationFrame(() => { syncing = false; });
    });
    bottom.addEventListener('scroll', () => {
      if (syncing) return;
      syncing = true;
      top.scrollLeft = bottom.scrollLeft;
      requestAnimationFrame(() => { syncing = false; });
    });
    top._syncBound = true;
  }
}

// Re-sync saat window di-resize (lebar tabel & overflow bisa berubah)
window.addEventListener('resize', () => setupTopScrollbar());

function autoGrow(el) {
  if (!el || !el.style) return;
  el.style.height = 'auto';
  const h = el.scrollHeight;
  if (h > 0) el.style.height = h + 'px';
}

/* ─────────────────────────────────────────────────────────────────────
   CELL BUILDERS
   Semua readonly cell kini punya tabindex="0" dan data-col-field
   sehingga bisa difokus, diselect, dicopy, dan dinavigasi dengan arrow.
───────────────────────────────────────────────────────────────────── */
const FIELD_TO_COL = {
  nomor_surat:      'col-nomor-surat',
  tanggal_surat:    'col-tgl-surat',
  waktu:            'col-waktu',
  perihal:          'col-perihal',
  tujuan:           'col-tujuan',
  pegawai_multi:    'col-nama',
  menimbang_custom: 'col-menimbang',
  alat_angkutan:    'col-alat',
  pembebanan:       'col-mak',
  penandatangan:    'col-ttd',
  tipe:             'col-tipe',
};

function cellTextHTML(id, field, val, editable, placeholder) {
  const cls = FIELD_TO_COL[field] || '';
  if (!editable) {
    return `<td class="${cls}">
      <div class="ro-text${val ? '' : ' muted'}"
           tabindex="0" data-col-field="${field}">${val ? esc(val) : '—'}</div>
    </td>`;
  }
  return `<td class="${cls}">
    <input type="text" class="xls-cell" data-field="${field}" data-id="${id}"
      data-col-field="${field}"
      value="${escAttr(val)}" placeholder="${escAttr(placeholder)}">
  </td>`;
}

function cellTextareaHTML(id, field, val, editable, placeholder) {
  const cls = FIELD_TO_COL[field] || '';
  if (!editable) {
    return `<td class="${cls}">
      <div class="ro-text${val ? '' : ' muted'}"
           tabindex="0" data-col-field="${field}">${val ? esc(val) : '—'}</div>
    </td>`;
  }
  return `<td class="${cls}">
    <textarea class="xls-cell" rows="1" data-field="${field}" data-id="${id}"
      data-col-field="${field}"
      placeholder="${escAttr(placeholder)}">${esc(val)}</textarea>
  </td>`;
}

/* Cell khusus utk kolom Tipe — render <select> dengan 7 pilihan kalau
   editable, atau ro-text kalau status sudah selesai.
   Value NULL/empty di-display sebagai '—' (readonly) atau placeholder
   "— Pilih tipe —" (editable) — admin wajib pilih sebelum approve. */
function cellTipeHTML(id, val, editable) {
  if (!editable) {
    const label = tipeLabel(val);
    return `<td class="col-tipe">
      <div class="ro-text${val ? '' : ' muted'}"
           tabindex="0" data-col-field="tipe">${val ? esc(label) : '—'}</div>
    </td>`;
  }
  const isEmpty = !val;
  const opts = TIPE_OPTIONS.map(o =>
    `<option value="${escAttr(o.value)}"${o.value === val ? ' selected' : ''}>${esc(o.label)}</option>`
  ).join('');
  return `<td class="col-tipe">
    <select class="xls-cell${isEmpty ? ' tipe-empty' : ''}"
            data-field="tipe" data-id="${id}" data-col-field="tipe"
            onchange="onTipeCellChange(this)">
      <option value=""${isEmpty ? ' selected' : ''}>— Pilih tipe —</option>
      ${opts}
    </select>
  </td>`;
}

// Toggle class 'tipe-empty' (placeholder italic) saat user pilih/clear opsi.
function onTipeCellChange(sel) {
  if (sel.value) sel.classList.remove('tipe-empty');
  else           sel.classList.add('tipe-empty');
  sel.classList.remove('err');
}

function cellDateHTML(id, field, isoVal, editable, placeholder) {
  if (!editable) {
    return `<td class="col-tgl-surat">
      <div class="ro-text${isoVal ? '' : ' muted'}"
           tabindex="0" data-col-field="${field}">${isoVal ? esc(fmtTgl(isoVal)) : '—'}</div>
    </td>`;
  }
  const display = isoVal ? fmtTgl(isoVal) : '';
  return `<td class="col-tgl-surat">
    <input type="text" class="xls-cell" data-field="${field}" data-id="${id}"
      data-col-field="${field}"
      data-iso="${escAttr(isoVal)}"
      value="${escAttr(display)}"
      placeholder="${escAttr(placeholder)}"
      onblur="onDateBlur(this)"
      onfocus="onDateFocus(this)"
      ondblclick="openCal(this, false)"
      onkeydown="if(event.altKey&&event.key==='ArrowDown'){event.preventDefault();openCal(this,false);}">
  </td>`;
}

function cellDateRangeHTML(id, field, isoMulai, isoSelesai, editable) {
  if (!editable) {
    return `<td class="col-waktu">
      <div class="ro-text${isoMulai ? '' : ' muted'}"
           tabindex="0" data-col-field="${field}">${isoMulai ? esc(fmtWaktu(isoMulai, isoSelesai)) : '—'}</div>
    </td>`;
  }
  const display = isoMulai ? fmtWaktu(isoMulai, isoSelesai) : '';
  return `<td class="col-waktu">
    <input type="text" class="xls-cell" data-field="${field}" data-id="${id}"
      data-col-field="${field}"
      data-iso-mulai="${escAttr(isoMulai)}"
      data-iso-selesai="${escAttr(isoSelesai || '')}"
      value="${escAttr(display)}"
      placeholder="tgl atau rentang"
      onblur="onRangeBlur(this)"
      onfocus="onDateFocus(this)"
      ondblclick="openCal(this, true)"
      onkeydown="if(event.altKey&&event.key==='ArrowDown'){event.preventDefault();openCal(this,true);}">
  </td>`;
}

function cellPegawaiMultiHTML(id, nips, names, editable) {
  if (!editable) {
    if (!names.length) {
      return `<td class="col-nama">
        <div class="ro-text muted" tabindex="0" data-col-field="pegawai_multi">—</div>
      </td>`;
    }
    return `<td class="col-nama">
      <div class="ro-text" tabindex="0" data-col-field="pegawai_multi">${names.map(esc).join(', ')}</div>
    </td>`;
  }
  const tags = nips.map((nip, i) => buildPegTag(nip, names[i] || nip, false)).join('');
  return `<td class="col-nama">
    <div class="pg-cell" data-field="pegawai_multi" data-col-field="pegawai_multi" data-id="${id}"
      data-nips='${escAttr(JSON.stringify(nips))}'
      data-names='${escAttr(JSON.stringify(names))}'
      onclick="onPgCellClick(event, this)">
      ${tags}
      <input type="text" class="pg-input" placeholder="${tags ? '' : 'Ketik nama...'}"
        oninput="onPgInput(event, this)"
        onkeydown="onPgKeydown(event, this)"
        onfocus="onPgInputFocus(this)"
        onblur="onPgInputBlur(this)">
    </div>
  </td>`;
}

function cellPenandatanganHTML(id, nip, nama, editable) {
  if (!editable) {
    if (!nama && !nip) {
      return `<td class="col-ttd">
        <div class="ro-text muted" tabindex="0" data-col-field="penandatangan">—</div>
      </td>`;
    }
    return `<td class="col-ttd">
      <div class="ro-ttd" tabindex="0" data-col-field="penandatangan">
        ${nama ? `<div class="ro-ttd-name">${esc(nama)}</div>` : ''}
        ${nip  ? `<div class="ro-ttd-nip">NIP. ${esc(nip)}</div>` : ''}
      </div>
    </td>`;
  }
  const tag = nip ? buildPegTag(nip, nama || nip, true) : '';
  return `<td class="col-ttd">
    <div class="pg-cell single" data-field="penandatangan" data-col-field="penandatangan" data-id="${id}"
      data-nip="${escAttr(nip)}"
      data-nama="${escAttr(nama)}"
      onclick="onPgCellClick(event, this)">
      ${tag}
      <input type="text" class="pg-input" placeholder="${tag ? '' : 'Pilih nama...'}"
        oninput="onPgInput(event, this)"
        onkeydown="onPgKeydown(event, this)"
        onfocus="onPgInputFocus(this)"
        onblur="onPgInputBlur(this)">
    </div>
  </td>`;
}

function buildPegTag(nip, nama, single) {
  const cls = single ? 'pg-tag single' : 'pg-tag';
  const display = nama && nama.length > 22 ? nama.slice(0, 20) + '…' : nama;
  return `<span class="${cls}" data-nip="${escAttr(nip)}" title="${escAttr(nama)}">
    <span class="pg-tag-text">${esc(display)}</span>
    <button type="button" class="pg-tag-x" onclick="onPgTagRemove(event, this)">×</button>
  </span>`;
}


/* ════════════════════════════════════════════════════════════════════
   DATE INPUT BEHAVIOR
═══════════════════════════════════════════════════════════════════════ */
function onDateFocus(el) {
  const iso = el.dataset.iso || '';
  if (iso) {
    const d = parseISODate(iso);
    if (d) el.value = `${d.getDate()}/${d.getMonth()+1}/${d.getFullYear()}`;
  }
  setTimeout(() => el.select(), 0);
}

function onDateBlur(el) {
  const cal = document.getElementById('cal-popup');
  if (cal && cal.classList.contains('open')) return;

  const txt = el.value.trim();
  if (!txt) {
    el.dataset.iso = '';
    el.value = '';
    return;
  }
  const iso = parseFlexDate(txt);
  if (iso) {
    el.dataset.iso = iso;
    el.value = fmtTgl(iso);
    el.classList.remove('err');
  } else {
    el.classList.add('err');
  }
}
function onRangeBlur(el) {
  const cal = document.getElementById('cal-popup');
  if (cal && cal.classList.contains('open')) return;

  // Jika dataset sudah terisi dari kalender, pakai langsung — jangan parse ulang teks tampilan
  if (el.dataset.isoMulai) {
    el.value = fmtWaktu(el.dataset.isoMulai, el.dataset.isoSelesai || '');
    el.classList.remove('err');
    return;
  }

  const txt = el.value.trim();
  if (!txt) {
    el.dataset.isoMulai = '';
    el.dataset.isoSelesai = '';
    el.value = '';
    return;
  }
  const r = parseFlexRange(txt);
  if (r.mulai) {
    el.dataset.isoMulai   = r.mulai;
    el.dataset.isoSelesai = r.selesai || '';
    el.value = fmtWaktu(r.mulai, r.selesai);
    el.classList.remove('err');
  } else {
    el.classList.add('err');
  }
}
/* ════════════════════════════════════════════════════════════════════
   CALENDAR POPUP
═══════════════════════════════════════════════════════════════════════ */
let calState = {
  targetEl: null,
  isRange:  false,
  year:     new Date().getFullYear(),
  month:    new Date().getMonth(),
  rangePhase: 0,
  yearMode: false,
  // Tugas #2 — keyboard navigation: tanggal yg sedang di-highlight
  // (di-set saat openCal, digerakkan oleh ←↑→↓, dipilih dgn Enter).
  focusedDay: '',
};

function openCal(el, isRange) {
  closeAllPopups();
  calState.targetEl = el;
  calState.isRange  = isRange;
  calState.rangePhase = 0;
  calState.yearMode = false;
  let initIso = isRange ? (el.dataset.isoMulai || todayISO()) : (el.dataset.iso || todayISO());
  const d = parseISODate(initIso);
  if (d) { calState.year = d.getFullYear(); calState.month = d.getMonth(); }
  // Init focusedDay ke tanggal yg sedang di-pick (atau hari ini kalau kosong)
  calState.focusedDay = initIso;
  const popup = document.getElementById('cal-popup');
  popup.classList.remove('year-mode');
  popup.classList.add('open');
  positionPopup(popup, el);
  renderCal();
}

function closeCal() {
  document.getElementById('cal-popup').classList.remove('open');
  calState.targetEl = null;
}

function renderCal() {
  const popup = document.getElementById('cal-popup');
  popup.classList.toggle('year-mode', calState.yearMode);
  document.getElementById('cal-title').textContent = `${BULAN[calState.month]} ${calState.year}`;
  document.getElementById('cal-day-names').innerHTML = ['Min','Sen','Sel','Rab','Kam','Jum','Sab']
    .map(d => `<div class="cal-day-name">${d}</div>`).join('');

  if (calState.yearMode) { renderYearGrid(); return; }

  const el = calState.targetEl;
  const startIso = calState.isRange ? (el.dataset.isoMulai || '') : (el.dataset.iso || '');
  const endIso   = calState.isRange ? (el.dataset.isoSelesai || '') : '';
  const firstDay = new Date(calState.year, calState.month, 1).getDay();
  const daysInMonth = new Date(calState.year, calState.month + 1, 0).getDate();
  const today = todayISO();

  let html = '';
  for (let i = 0; i < firstDay; i++) html += `<div class="cal-day cal-empty"></div>`;
  for (let d = 1; d <= daysInMonth; d++) {
    const ds = `${calState.year}-${String(calState.month+1).padStart(2,'0')}-${String(d).padStart(2,'0')}`;
    let cls = 'cal-day';
    if (ds === today) cls += ' cal-today';
    if (calState.isRange && startIso && endIso && startIso !== endIso) {
      if (ds === startIso) cls += ' cal-start';
      else if (ds === endIso) cls += ' cal-end';
      else if (ds > startIso && ds < endIso) cls += ' cal-in-range';
    } else if (startIso && ds === startIso) {
      cls += calState.isRange ? (endIso === startIso ? ' cal-single' : ' cal-start') : ' cal-single';
    }
    if (ds === calState.focusedDay) cls += ' cal-focused';
    html += `<div class="${cls}" data-date="${ds}">${d}</div>`;
  }
  document.getElementById('cal-days').innerHTML = html;

  let txt = 'Pilih tanggal';
  if (calState.isRange) {
    if (calState.rangePhase === 1 && startIso) txt = `Mulai: ${fmtTgl(startIso)} — pilih tgl selesai`;
    else if (startIso && endIso) txt = `${fmtTgl(startIso)} s.d. ${fmtTgl(endIso)}`;
    else if (startIso) txt = fmtTgl(startIso);
  } else if (startIso) {
    txt = fmtTgl(startIso);
  }
  document.getElementById('cal-footer-text').textContent = txt;
}

function renderYearGrid() {
  const baseYear = Math.floor(calState.year / 12) * 12;
  let html = '';
  for (let i = 0; i < 12; i++) {
    const y = baseYear + i;
    const cur = y === calState.year ? ' current' : '';
    html += `<div class="cal-year-cell${cur}" data-year="${y}">${y}</div>`;
  }
  document.getElementById('cal-year-grid').innerHTML = html;
}

function positionPopup(popup, anchorEl) {
  popup.style.visibility = 'hidden';
  popup.style.display = 'block';
  const pH = popup.offsetHeight;
  const pW = popup.offsetWidth;
  popup.style.display = '';
  popup.style.visibility = '';
  const rect = anchorEl.getBoundingClientRect();
  const winH = window.innerHeight, winW = window.innerWidth;
  let top = rect.bottom + 4;
  let left = rect.left;
  if (top + pH > winH - 8) top = Math.max(8, rect.top - pH - 4);
  if (left + pW > winW - 8) left = winW - pW - 8;
  if (left < 8) left = 8;
  popup.style.top = top + 'px';
  popup.style.left = left + 'px';
}

document.addEventListener('DOMContentLoaded', () => {
  document.getElementById('cal-prev').onclick = (e) => {
    e.stopPropagation();
    if (calState.yearMode) { calState.year -= 12; renderYearGrid(); return; }
    calState.month--;
    if (calState.month < 0) { calState.month = 11; calState.year--; }
    renderCal();
  };
  document.getElementById('cal-next').onclick = (e) => {
    e.stopPropagation();
    if (calState.yearMode) { calState.year += 12; renderYearGrid(); return; }
    calState.month++;
    if (calState.month > 11) { calState.month = 0; calState.year++; }
    renderCal();
  };
  document.getElementById('cal-title').onclick = (e) => {
    e.stopPropagation();
    calState.yearMode = !calState.yearMode;
    renderCal();
  };
  document.getElementById('cal-clear').onclick = (e) => {
    e.stopPropagation();
    if (!calState.targetEl) return;
    if (calState.isRange) {
      calState.targetEl.dataset.isoMulai = '';
      calState.targetEl.dataset.isoSelesai = '';
    } else {
      calState.targetEl.dataset.iso = '';
    }
    calState.targetEl.value = '';
    calState.rangePhase = 0;
    renderCal();
    closeCal();
  };
  document.getElementById('cal-days').onclick = (e) => {
    const d = e.target.closest('.cal-day');
    if (!d || d.classList.contains('cal-empty')) return;
    calPickDate(d.dataset.date);
  };
  document.getElementById('cal-year-grid').onclick = (e) => {
    const c = e.target.closest('.cal-year-cell');
    if (!c) return;
    calState.year = parseInt(c.dataset.year, 10);
    calState.yearMode = false;
    renderCal();
  };
});

/* ──────────────────────────────────────────────────────────────────
   Tugas #2 — Keyboard navigation di calendar popup
   ──────────────────────────────────────────────────────────────────
   ←↑→↓  navigasi 1 hari / 1 minggu (auto-pindah bulan kalau lewat batas)
   Enter pilih tanggal yg sedang di-focus
   Esc   tutup popup (sudah di-handle di Escape clause global)
─────────────────────────────────────────────────────────────────── */

/**
 * Pick tanggal ds ke target input. Dipanggil oleh klik mouse (event handler
 * di onclick cal-days) DAN keyboard Enter.
 */
function calPickDate(ds) {
  const el = calState.targetEl;
  if (!el) return;
  if (!calState.isRange) {
    el.dataset.iso = ds;
    el.value = fmtTgl(ds);
    el.classList.remove('err');
    closeCal();
    return;
  }
  if (calState.rangePhase === 0) {
    el.dataset.isoMulai = ds;
    el.dataset.isoSelesai = '';
    el.value = fmtTgl(ds) + ' s.d. …';
    calState.rangePhase = 1;
    renderCal();
  } else {
    let mulai = el.dataset.isoMulai;
    let selesai = ds;
    if (selesai < mulai) { [mulai, selesai] = [selesai, mulai]; }
    el.dataset.isoMulai = mulai;
    el.dataset.isoSelesai = (selesai === mulai) ? '' : selesai;
    el.value = fmtWaktu(mulai, el.dataset.isoSelesai);
    el.classList.remove('err');
    calState.rangePhase = 0;
    renderCal();
    setTimeout(closeCal, 150);
  }
}

/**
 * Geser focusedDay sebanyak deltaDays. Auto-pindah bulan kalau pindah ke
 * tanggal di bulan lain. Skip kalau popup tidak terbuka atau di year-mode.
 */
function calMoveFocus(deltaDays) {
  const popup = document.getElementById('cal-popup');
  if (!popup || !popup.classList.contains('open') || calState.yearMode) return;

  let baseIso = calState.focusedDay;
  if (!baseIso) {
    // Fallback: pakai tgl yg sedang di-set di target, atau hari ini
    const el = calState.targetEl;
    baseIso = (el && (calState.isRange ? el.dataset.isoMulai : el.dataset.iso))
           || todayISO();
  }
  const d = parseISODate(baseIso) || new Date();
  d.setDate(d.getDate() + deltaDays);
  const newIso = `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')}`;
  // Pindah bulan kalau perlu
  if (d.getFullYear() !== calState.year || d.getMonth() !== calState.month) {
    calState.year  = d.getFullYear();
    calState.month = d.getMonth();
  }
  calState.focusedDay = newIso;
  renderCal();
  // Scroll ke focused cell kalau ke luar viewport (rare di calendar 6×7)
  const focusedEl = document.querySelector(`#cal-days .cal-day.cal-focused`);
  if (focusedEl && focusedEl.scrollIntoView) {
    focusedEl.scrollIntoView({ block: 'nearest', inline: 'nearest' });
  }
}

function calPickFocused() {
  if (!calState.focusedDay) return;
  // Pastikan tanggal yg di-focus adalah dalam bulan yang sedang di-render
  // (calMoveFocus sudah jaga ini, tapi defensive)
  calPickDate(calState.focusedDay);
}

/* ════════════════════════════════════════════════════════════════════
   AUTOCOMPLETE PEGAWAI
═══════════════════════════════════════════════════════════════════════ */
let acState = {
  cellEl:    null,
  inputEl:   null,
  filtered:  [],
  focusIdx:  -1,
  isSingle:  false,
};

function onPgCellClick(e, cellEl) {
  if (e.target === cellEl) {
    const inp = cellEl.querySelector('.pg-input');
    if (inp) inp.focus();
  }
}

function onPgInputFocus(inp) {
  const cellEl = inp.closest('.pg-cell');
  if (!cellEl) return;
  cellEl.classList.add('focused');
  openAc(cellEl, inp);
}

function onPgInputBlur(inp) {
  const cellEl = inp.closest('.pg-cell');
  if (cellEl) cellEl.classList.remove('focused');
  setTimeout(() => {
    if (!document.activeElement || !document.getElementById('ac-popup').contains(document.activeElement)) {
      if (!document.activeElement || !document.activeElement.classList.contains('pg-input')) {
        closeAc();
      }
    }
  }, 150);
}

function onPgInput(e, inp) {
  acState.inputEl = inp;
  acState.cellEl  = inp.closest('.pg-cell');
  acState.isSingle = acState.cellEl.classList.contains('single');
  acFilter(inp.value);
}

function onPgKeydown(e, inp) {
  const cellEl = inp.closest('.pg-cell');
  if (!cellEl) return;

  if (e.key === 'Backspace' && inp.value === '') {
    const tags = cellEl.querySelectorAll('.pg-tag');
    if (tags.length) {
      e.preventDefault();
      const lastTag = tags[tags.length - 1];
      removePegTag(cellEl, lastTag.dataset.nip);
    }
    return;
  }

  if (!document.getElementById('ac-popup').classList.contains('open')) {
    if (e.key === 'ArrowDown' || e.key === 'Enter') {
      e.preventDefault();
      openAc(cellEl, inp);
      acFilter(inp.value);
    }
    return;
  }

  if (e.key === 'ArrowDown') {
    e.preventDefault();
    acState.focusIdx = Math.min(acState.focusIdx + 1, acState.filtered.length - 1);
    acRenderFocus();
  } else if (e.key === 'ArrowUp') {
    e.preventDefault();
    acState.focusIdx = Math.max(acState.focusIdx - 1, 0);
    acRenderFocus();
  } else if (e.key === 'Enter') {
    e.preventDefault();
    if (acState.focusIdx >= 0 && acState.filtered[acState.focusIdx]) {
      const p = acState.filtered[acState.focusIdx];
      pickPegawai(cellEl, String(p.NIP).trim(), p.NAMA || '');
      inp.value = '';
      acFilter('');
    }
  } else if (e.key === 'Escape') {
    e.preventDefault();
    closeAc();
  } else if (e.key === 'Tab') {
    closeAc();
  }
}

function openAc(cellEl, inp) {
  acState.cellEl  = cellEl;
  acState.inputEl = inp;
  acState.isSingle = cellEl.classList.contains('single');
  acState.focusIdx = -1;
  const popup = document.getElementById('ac-popup');
  popup.classList.add('open');
  positionPopup(popup, cellEl);
  acFilter(inp.value);
}

function closeAc() {
  document.getElementById('ac-popup').classList.remove('open');
  acState.cellEl = null;
  acState.inputEl = null;
  acState.filtered = [];
  acState.focusIdx = -1;
}

function acFilter(q) {
  q = (q || '').toLowerCase().trim();
  const cellEl = acState.cellEl;
  let selectedNips = [];
  if (cellEl) {
    if (acState.isSingle) {
      const nip = cellEl.dataset.nip;
      if (nip) selectedNips = [nip];
    } else {
      try { selectedNips = JSON.parse(cellEl.dataset.nips || '[]'); } catch(_) {}
    }
  }
  acState.filtered = pegawaiList.filter(p => {
    if (!q) return true;
    const nama = (p.NAMA || '').toLowerCase();
    const nip = String(p.NIP || '').toLowerCase();
    return nama.includes(q) || nip.includes(q);
  }).slice(0, 50);
  acRenderList(selectedNips);
  document.getElementById('ac-count').textContent = `${acState.filtered.length} hasil`;
}

function acRenderList(selectedNips) {
  const list = document.getElementById('ac-list');
  if (!pegawaiList.length) {
    list.innerHTML = `<div class="ac-empty">Data pegawai belum dimuat.</div>`;
    return;
  }
  if (!acState.filtered.length) {
    list.innerHTML = `<div class="ac-empty">Tidak ada hasil — coba ketik nama lain</div>`;
    return;
  }
  list.innerHTML = acState.filtered.map((p, i) => {
    const nip = String(p.NIP || '').trim();
    const nama = p.NAMA || '-';
    const isSel = selectedNips.includes(nip);
    return `<div class="ac-item${isSel ? ' selected' : ''}${i === acState.focusIdx ? ' focused' : ''}"
      data-nip="${escAttr(nip)}" data-idx="${i}"
      onmousedown="event.preventDefault();acPick(${i})">
      <div class="ac-check">${isSel ? '✓' : ''}</div>
      <div style="flex:1">
        <div class="ac-name">${esc(nama)}</div>
        <div class="ac-nip">NIP ${esc(nip || '—')}</div>
      </div>
    </div>`;
  }).join('');
}

function acRenderFocus() {
  const items = document.querySelectorAll('#ac-list .ac-item');
  items.forEach((el, i) => el.classList.toggle('focused', i === acState.focusIdx));
  if (acState.focusIdx >= 0 && items[acState.focusIdx]) {
    items[acState.focusIdx].scrollIntoView({ block: 'nearest' });
  }
}

function acPick(idx) {
  const p = acState.filtered[idx];
  if (!p || !acState.cellEl) return;
  pickPegawai(acState.cellEl, String(p.NIP).trim(), p.NAMA || '');
  if (acState.inputEl) {
    acState.inputEl.value = '';
    acState.inputEl.focus();
  }
  acFilter('');
}

function pickPegawai(cellEl, nip, nama) {
  if (acState.isSingle) {
    cellEl.querySelectorAll('.pg-tag').forEach(t => t.remove());
    cellEl.dataset.nip = nip;
    cellEl.dataset.nama = nama;
    const tagHtml = buildPegTag(nip, nama, true);
    const inp = cellEl.querySelector('.pg-input');
    inp.insertAdjacentHTML('beforebegin', tagHtml);
    inp.placeholder = '';
    closeAc();
  } else {
    let nips = [];
    let names = [];
    try { nips  = JSON.parse(cellEl.dataset.nips  || '[]'); } catch(_) {}
    try { names = JSON.parse(cellEl.dataset.names || '[]'); } catch(_) {}
    if (nips.includes(nip)) {
      const existing = cellEl.querySelector(`.pg-tag[data-nip="${CSS.escape(nip)}"]`);
      if (existing) {
        existing.style.transition = 'background .3s';
        existing.style.background = '#fef3c7';
        setTimeout(() => { existing.style.background = ''; }, 400);
      }
      return;
    }
    nips.push(nip);
    names.push(nama);
    cellEl.dataset.nips  = JSON.stringify(nips);
    cellEl.dataset.names = JSON.stringify(names);
    const tagHtml = buildPegTag(nip, nama, false);
    const inp = cellEl.querySelector('.pg-input');
    inp.insertAdjacentHTML('beforebegin', tagHtml);
    inp.placeholder = '';
  }
}

function onPgTagRemove(e, btn) {
  e.stopPropagation();
  const tag = btn.closest('.pg-tag');
  const cellEl = btn.closest('.pg-cell');
  if (!tag || !cellEl) return;
  removePegTag(cellEl, tag.dataset.nip);
}

function removePegTag(cellEl, nip) {
  const isSingle = cellEl.classList.contains('single');
  const tag = cellEl.querySelector(`.pg-tag[data-nip="${CSS.escape(nip)}"]`);
  if (tag) tag.remove();
  if (isSingle) {
    cellEl.dataset.nip = '';
    cellEl.dataset.nama = '';
  } else {
    let nips = [];
    let names = [];
    try { nips  = JSON.parse(cellEl.dataset.nips  || '[]'); } catch(_) {}
    try { names = JSON.parse(cellEl.dataset.names || '[]'); } catch(_) {}
    const idx = nips.indexOf(nip);
    if (idx >= 0) { nips.splice(idx, 1); names.splice(idx, 1); }
    cellEl.dataset.nips  = JSON.stringify(nips);
    cellEl.dataset.names = JSON.stringify(names);
  }
  const inp = cellEl.querySelector('.pg-input');
  if (inp) {
    const stillHasTags = cellEl.querySelectorAll('.pg-tag').length > 0;
    inp.placeholder = stillHasTags ? '' : (isSingle ? 'Pilih nama...' : 'Ketik nama...');
    inp.focus();
  }
}

/* ════════════════════════════════════════════════════════════════════
   AUTOCOMPLETE POK (MAK PEMBEBANAN)

   Sumber data: distinct `pembebanan` dari tabel surat_tugas (history).
   MAK yang sering dipakai akan muncul di atas (sorted by frequency).
   User tetap bisa ketik MAK baru — validasi format pakai parseMAK()
   yang sudah ada (regex MAK_REGEX). Free-text fallback didukung untuk
   MAK yang belum pernah dipakai sebelumnya.
═══════════════════════════════════════════════════════════════════════ */
let makSuggestions = [];   // [{mak, count, ringkasan}]
let makACState = {
  cellEl:   null,
  inputEl:  null,
  filtered: [],
  focusIdx: -1,
};

async function loadMAKSuggestions() {
  try {
    const url = `${SUPABASE_URL}/rest/v1/surat_tugas?select=pembebanan&pembebanan=not.is.null`;
    const res = await fetch(url, { headers: H });
    if (!res.ok) {
      console.warn(`[NOVA] loadMAKSuggestions HTTP ${res.status}`);
      return;
    }
    const rows = await res.json();
    // Hitung frekuensi pemakaian per MAK valid
    const counts = {};
    rows.forEach(r => {
      const m = (r.pembebanan || '').trim();
      if (m && parseMAK(m)) counts[m] = (counts[m] || 0) + 1;
    });
    makSuggestions = Object.entries(counts)
      .map(([mak, count]) => ({ mak, count, ringkasan: '' }))
      .sort((a, b) => b.count - a.count);
    // Bangun ringkasan (deskripsi gabungan) — load kamus_pok kalau ada
    await enrichMakSuggestionsWithDeskripsi();
    console.log(`[NOVA] loadMAKSuggestions: ${makSuggestions.length} unique MAK`);
  } catch (e) {
    console.warn('[NOVA] loadMAKSuggestions error:', e);
  }
}

async function enrichMakSuggestionsWithDeskripsi() {
  if (!makSuggestions.length) return;
  // Pakai tahun saat ini sebagai default lookup. Deskripsi cuma untuk
  // membantu admin mengenali MAK; kalau ada beda antar tahun, anggap saja
  // sebagai hint informatif, bukan otoritatif.
  const tahun = new Date().getFullYear();
  try {
    if (typeof loadKamusPok === 'function') await loadKamusPok(tahun);
  } catch (_) {}

  makSuggestions.forEach(s => {
    const mak = parseMAK(s.mak);
    if (!mak || typeof lookupDeskripsi !== 'function') return;
    const desKgt = lookupDeskripsi('kegiatan', mak.kegiatan);
    const desAkn = lookupDeskripsi('akun',     mak.akun);
    // Format ringkas: "Kegiatan ... · Akun ..."
    const parts = [];
    if (desKgt) parts.push(desKgt);
    if (desAkn) parts.push(desAkn);
    s.ringkasan = parts.join(' · ');
  });
}

function cellMAKHTML(id, val, editable) {
  if (!editable) {
    return `<td class="col-mak">
      <div class="ro-text${val ? '' : ' muted'}" tabindex="0" data-col-field="pembebanan">${val ? esc(val) : '—'}</div>
    </td>`;
  }
  return `<td class="col-mak">
    <input type="text" class="xls-cell mak-input" data-field="pembebanan" data-id="${id}"
      data-col-field="pembebanan"
      value="${escAttr(val || '')}"
      placeholder="cth: 054.01.GG.2910.BMA.006.054.A.524119"
      oninput="onMAKInput(this)"
      onfocus="onMAKFocus(this)"
      onblur="onMAKBlur(this)"
      onkeydown="onMAKKeydown(event, this)"
      autocomplete="off">
  </td>`;
}

function onMAKFocus(inp) {
  openMakAc(inp);
}

function onMAKBlur(inp) {
  // Delay supaya klik pada item popup sempat ke-handle sebelum popup ditutup
  setTimeout(() => {
    const ae = document.activeElement;
    if (!ae || !document.getElementById('mak-ac-popup').contains(ae)) {
      if (!ae || !ae.classList.contains('mak-input')) closeMakAc();
    }
  }, 150);
}

function onMAKInput(inp) {
  makACState.inputEl = inp;
  if (!document.getElementById('mak-ac-popup').classList.contains('open')) {
    openMakAc(inp);
  }
  makAcFilter(inp.value);
}

function onMAKKeydown(e, inp) {
  const popup = document.getElementById('mak-ac-popup');
  const isOpen = popup.classList.contains('open');

  if (e.key === 'Escape') {
    if (isOpen) { e.preventDefault(); closeMakAc(); }
    return;
  }

  if (!isOpen) {
    if (e.key === 'ArrowDown') {
      e.preventDefault();
      openMakAc(inp);
      makAcFilter(inp.value);
    }
    return;
  }

  if (e.key === 'ArrowDown') {
    e.preventDefault();
    makACState.focusIdx = Math.min(makACState.focusIdx + 1, makACState.filtered.length - 1);
    makAcRenderFocus();
  } else if (e.key === 'ArrowUp') {
    e.preventDefault();
    makACState.focusIdx = Math.max(makACState.focusIdx - 1, -1);
    makAcRenderFocus();
  } else if (e.key === 'Enter') {
    // Enter HANYA untuk pilih item dari dropdown — TIDAK pindah cell.
    // (Konsisten dengan permintaan: Enter tidak navigasi cell.)
    if (makACState.focusIdx >= 0 && makACState.filtered[makACState.focusIdx]) {
      e.preventDefault();
      pickMAK(makACState.filtered[makACState.focusIdx].mak);
    }
    // Kalau tidak ada item terpilih → Enter no-op (default browser),
    // input single-line jadi tidak melakukan apa-apa.
  } else if (e.key === 'Tab') {
    closeMakAc(); // biar handler Tab global yang pindah cell
  }
}

function openMakAc(inp) {
  makACState.inputEl = inp;
  makACState.cellEl = inp.closest('td');
  makACState.focusIdx = -1;
  const popup = document.getElementById('mak-ac-popup');
  popup.classList.add('open');
  positionPopup(popup, inp);
  makAcFilter(inp.value);
}

function closeMakAc() {
  document.getElementById('mak-ac-popup').classList.remove('open');
  makACState.cellEl = null;
  makACState.inputEl = null;
  makACState.filtered = [];
  makACState.focusIdx = -1;
}

function makAcFilter(q) {
  q = (q || '').toLowerCase().trim();
  makACState.filtered = makSuggestions.filter(s => {
    if (!q) return true;
    return s.mak.toLowerCase().includes(q) ||
           (s.ringkasan && s.ringkasan.toLowerCase().includes(q));
  }).slice(0, 50);
  makACState.focusIdx = -1;
  makAcRenderList();
  document.getElementById('mak-ac-count').textContent = `${makACState.filtered.length} hasil`;
}

function makAcRenderList() {
  const list = document.getElementById('mak-ac-list');
  if (!makSuggestions.length) {
    list.innerHTML = `<div class="mak-ac-empty">Belum ada riwayat POK.<br><span style="font-size:10.5px">Ketik MAK manual sesuai format.</span></div>`;
    return;
  }
  if (!makACState.filtered.length) {
    list.innerHTML = `<div class="mak-ac-empty">Tidak ada riwayat yang cocok.<br><span style="font-size:10.5px">Anda tetap bisa mengetik MAK baru langsung di field.</span></div>`;
    return;
  }
  list.innerHTML = makACState.filtered.map((s, i) => `
    <div class="mak-ac-item${i === makACState.focusIdx ? ' focused' : ''}"
         data-idx="${i}"
         onmousedown="event.preventDefault()"
         onclick="pickMAK('${escAttr(s.mak)}')">
      <div class="mak-ac-item-code">${esc(s.mak)}<span class="mak-ac-item-count">${s.count}×</span></div>
      ${s.ringkasan ? `<div class="mak-ac-item-desc">${esc(s.ringkasan)}</div>` : ''}
    </div>
  `).join('');
}

function makAcRenderFocus() {
  const items = document.querySelectorAll('#mak-ac-list .mak-ac-item');
  items.forEach((el, i) => el.classList.toggle('focused', i === makACState.focusIdx));
  if (makACState.focusIdx >= 0) {
    const focused = items[makACState.focusIdx];
    if (focused) focused.scrollIntoView({ block: 'nearest' });
  }
}

function pickMAK(mak) {
  const inp = makACState.inputEl;
  if (inp) {
    inp.value = mak;
    inp.classList.remove('err');
    inp.dispatchEvent(new Event('input', { bubbles: true }));
  }
  closeMakAc();
}

/* ════════════════════════════════════════════════════════════════════
   ARROW NAVIGATION GRID — navigasi Excel-like antar sel
   Mendukung: editable (input/textarea/pg-input) + readonly (ro-text/ro-ttd)
═══════════════════════════════════════════════════════════════════════ */

// Urutan kolom navigasi (kiri → kanan, sesuai urutan visual tabel)
const NAV_FIELDS = [
  'nomor_surat',
  'tanggal_surat',
  'waktu',
  'perihal',
  'tujuan',
  'pegawai_multi',
  'menimbang_custom',
  'alat_angkutan',
  'pembebanan',
  'penandatangan',
];

/**
 * Bangun grid 2D: grid[rowIdx][colIdx] = focusable HTMLElement | null
 * Baris: sesuai urutan tampilan tabel (descending created_at)
 * Kolom: sesuai NAV_FIELDS
 */
function buildNavGrid() {
  const rows = Array.from(document.querySelectorAll('tr[data-surat-id]'));
  return rows.map(row => {
    return NAV_FIELDS.map(field => {
      // 1) Editable input atau textarea
      let el = row.querySelector(
        `input.xls-cell[data-col-field="${field}"],
         textarea.xls-cell[data-col-field="${field}"]`
      );
      if (el) return el;

      // 2) Editable pg-cell → kembalikan pg-input di dalamnya
      const pgCell = row.querySelector(`.pg-cell[data-col-field="${field}"]`);
      if (pgCell) {
        return pgCell.querySelector('.pg-input') || pgCell;
      }

      // 3) Readonly div (ro-text atau ro-ttd)
      el = row.querySelector(`[data-col-field="${field}"]`);
      return el || null;
    });
  });
}

/**
 * Cari posisi elemen dalam grid.
 * Jika elemen adalah pg-input, cari parent pg-cell-nya di grid.
 */
function findInGrid(grid, target) {
  for (let r = 0; r < grid.length; r++) {
    for (let c = 0; c < grid[r].length; c++) {
      const el = grid[r][c];
      if (!el) continue;
      if (el === target) return { r, c };
      // pg-input → cek apakah parent pg-cell ada di grid
      if (target.classList && target.classList.contains('pg-input')) {
        const parentPg = target.closest('.pg-cell');
        if (parentPg && el === parentPg) return { r, c };
      }
    }
  }
  return null;
}

/**
 * Fokus ke sel di posisi [r][c] dengan scroll smooth.
 * Jika sel di kolom c null, coba geser ke kolom terdekat yang ada.
 */
function navFocusCell(grid, r, c) {
  r = Math.max(0, Math.min(r, grid.length - 1));
  c = Math.max(0, Math.min(c, NAV_FIELDS.length - 1));

  let el = grid[r][c];
  // Fallback: geser kanan kalau null
  if (!el) {
    for (let cc = c + 1; cc < grid[r].length; cc++) {
      if (grid[r][cc]) { el = grid[r][cc]; break; }
    }
  }
  // Fallback: geser kiri
  if (!el) {
    for (let cc = c - 1; cc >= 0; cc--) {
      if (grid[r][cc]) { el = grid[r][cc]; break; }
    }
  }
  if (!el) return;

  el.focus({ preventScroll: true });
  el.scrollIntoView({ block: 'nearest', inline: 'nearest', behavior: 'smooth' });

  // Select teks jika input/textarea
  if (el.tagName === 'INPUT' || el.tagName === 'TEXTAREA') {
    try { el.select(); } catch(_) {}
  }

  // Highlight baris aktif
  document.querySelectorAll('tr.row-focused').forEach(tr => tr.classList.remove('row-focused'));
  const parentRow = el.closest('tr');
  if (parentRow) parentRow.classList.add('row-focused');
}

/* ════════════════════════════════════════════════════════════════════
   GLOBAL KEYDOWN — Tab/Enter (navigasi linear) + Arrow (navigasi grid)
   + Escape (tutup popup/modal)
═══════════════════════════════════════════════════════════════════════ */
document.addEventListener('keydown', (e) => {

  /* ── Tugas #2 — Calendar popup keyboard nav ──────────────────────
     Saat cal-popup terbuka, arrow keys navigasi tanggal & Enter pilih.
     Listener ini di atas Escape clause supaya tidak konflik dengan
     handler grid di bawah (yg juga consume arrow keys). */
  const calPopup = document.getElementById('cal-popup');
  if (calPopup && calPopup.classList.contains('open') && !calState.yearMode) {
    if (e.key === 'ArrowUp')    { e.preventDefault(); calMoveFocus(-7); return; }
    if (e.key === 'ArrowDown')  { e.preventDefault(); calMoveFocus( 7); return; }
    if (e.key === 'ArrowLeft')  { e.preventDefault(); calMoveFocus(-1); return; }
    if (e.key === 'ArrowRight') { e.preventDefault(); calMoveFocus( 1); return; }
    if (e.key === 'Enter')      { e.preventDefault(); calPickFocused(); return; }
    // Escape tetap fall-through ke clause Escape di bawah (close popup)
  }

  /* ── Escape ─────────────────────────────────────────────────────── */
  if (e.key === 'Escape') {
    closeAllPopups();
    ['modal-approve','modal-preview'].forEach(closeModal);
    document.querySelectorAll('tr.row-focused').forEach(tr => tr.classList.remove('row-focused'));
    return;
  }

  const target = e.target;

  // Tentukan tipe elemen yang sedang difokus
  const isXlsCell  = target.classList && target.classList.contains('xls-cell');
  const isPgInput  = target.classList && target.classList.contains('pg-input');
  const isRoText   = target.classList && (
    target.classList.contains('ro-text') || target.classList.contains('ro-ttd')
  );
  const isNavTarget = isXlsCell || isPgInput || isRoText;

  if (!isNavTarget) return;

  const isTextarea  = target.tagName === 'TEXTAREA';
  const isReadonly  = isRoText;

  /* ── Arrow keys: navigasi grid ──────────────────────────────────── */
  if (['ArrowUp','ArrowDown','ArrowLeft','ArrowRight'].includes(e.key)) {

    // Alt+ArrowDown di date cell = buka kalender, jangan intercept
    if (e.altKey && e.key === 'ArrowDown') return;

    // ArrowUp/Down di autocomplete pegawai = biarkan handler onPgKeydown handle
    if (isPgInput && document.getElementById('ac-popup').classList.contains('open')) {
      if (e.key === 'ArrowUp' || e.key === 'ArrowDown') return;
    }

    // ArrowUp/Down di MAK typeahead = biarkan onMAKKeydown handle navigate item
    if (target.classList && target.classList.contains('mak-input') &&
        document.getElementById('mak-ac-popup').classList.contains('open')) {
      if (e.key === 'ArrowUp' || e.key === 'ArrowDown') return;
    }

    const grid = buildNavGrid();
    const pos  = findInGrid(grid, target);
    if (!pos) return;

    const { r, c } = pos;

    if (e.key === 'ArrowUp') {
      if (r > 0) {
        e.preventDefault();
        closeAllPopups();
        navFocusCell(grid, r - 1, c);
      }
      return;
    }

    if (e.key === 'ArrowDown') {
      if (r < grid.length - 1) {
        e.preventDefault();
        closeAllPopups();
        navFocusCell(grid, r + 1, c);
      }
      return;
    }

    if (e.key === 'ArrowLeft') {
      // Readonly: selalu pindah kolom
      // Editable: pindah kolom hanya jika kursor di posisi paling kiri (0)
      const cursorAtStart = isReadonly || (
        target.selectionStart === 0 && target.selectionEnd === 0
      );
      if (cursorAtStart && c > 0) {
        e.preventDefault();
        closeAllPopups();
        navFocusCell(grid, r, c - 1);
      }
      return;
    }

    if (e.key === 'ArrowRight') {
      // Readonly: selalu pindah kolom
      // Editable: pindah kolom hanya jika kursor di posisi paling kanan (akhir teks)
      const valLen = (target.value || '').length;
      const cursorAtEnd = isReadonly || (
        target.selectionStart === valLen && target.selectionEnd === valLen
      );
      if (cursorAtEnd && c < NAV_FIELDS.length - 1) {
        e.preventDefault();
        closeAllPopups();
        navFocusCell(grid, r, c + 1);
      }
      return;
    }
  }

  /* ── Tab: navigasi linear (hanya untuk sel editable) ─────────────
     ENTER sengaja TIDAK dipakai untuk navigasi cell — pindah kolom
     hanya via Tab (atau arrow keys di handler di atas).
     - Plain Enter di textarea (cell perihal/menimbang/dst.) = newline default browser
     - Plain Enter di input single-line = no-op (tidak melakukan apa-apa)
     - Shift+Enter di textarea juga newline default                      */
  if (isReadonly) return; // readonly: biarkan Tab default browser

  if (e.key === 'Tab') {
    e.preventDefault();
    moveCellFocus(target, e.shiftKey ? -1 : 1);
  }
});

/**
 * Navigasi linear Tab/Enter di sel editable.
 * Meliputi semua xls-cell dan pg-input di baris menunggu.
 */
function moveCellFocus(currentEl, direction) {
  const allCells = Array.from(document.querySelectorAll(
    'tr[data-status="menunggu"] .xls-cell, tr[data-status="menunggu"] .pg-input'
  ));
  const idx  = allCells.indexOf(currentEl);
  if (idx < 0) return;
  const next = allCells[idx + direction];
  if (next) {
    next.focus();
    if (next.tagName === 'INPUT' && next.type === 'text') next.select();
  }
}

/* ════════════════════════════════════════════════════════════════════
   CLOSE POPUPS saat klik di luar
═══════════════════════════════════════════════════════════════════════ */
document.addEventListener('click', (e) => {
  const cal = document.getElementById('cal-popup');
  // Gunakan composedPath() agar tetap akurat walau DOM berubah saat renderCal()
  const path = e.composedPath ? e.composedPath() : [];
  const clickInsideCal = path.includes(cal);
  if (cal.classList.contains('open') && document.contains(e.target) && !cal.contains(e.target) &&
      (!calState.targetEl || !calState.targetEl.contains(e.target)) && calState.targetEl !== e.target) {
    closeCal();
  }
  const ac = document.getElementById('ac-popup');
  if (ac.classList.contains('open') && !ac.contains(e.target)) {
    const inAnyPgCell = e.target.closest && e.target.closest('.pg-cell');
    if (!inAnyPgCell) closeAc();
  }
  const macAc = document.getElementById('mak-ac-popup');
  if (macAc && macAc.classList.contains('open') && !macAc.contains(e.target)) {
    const inMakInput = e.target.closest && e.target.closest('.mak-input');
    if (!inMakInput) closeMakAc();
  }
  // Hapus highlight baris jika klik di luar tabel
  if (!e.target.closest('tr[data-surat-id]')) {
    document.querySelectorAll('tr.row-focused').forEach(tr => tr.classList.remove('row-focused'));
  }
});

/* ════════════════════════════════════════════════════════════════════
   AUTO-SELECT TEKS di readonly cell saat difokus
   Sehingga langsung bisa Ctrl+C tanpa perlu drag manual
═══════════════════════════════════════════════════════════════════════ */
document.addEventListener('focus', (e) => {
  const target = e.target;
  if (!target) return;
  const isRo = target.classList.contains('ro-text') || target.classList.contains('ro-ttd');
  if (!isRo) return;
  try {
    const range = document.createRange();
    range.selectNodeContents(target);
    const sel = window.getSelection();
    sel.removeAllRanges();
    sel.addRange(range);
  } catch(_) {}
}, true); // useCapture=true karena focus tidak bubble secara default

function closeAllPopups() {
  closeCal();
  closeAc();
  closeMakAc();
}

/* ════════════════════════════════════════════════════════════════════
   COLLECT FIELDS dari row sebelum approve
═══════════════════════════════════════════════════════════════════════ */
function collectRowFields(suratId) {
  const row = document.querySelector(`tr[data-surat-id="${suratId}"]`);
  if (!row) return null;

  const get = (field) => {
    const el = row.querySelector(`[data-field="${field}"]`);
    return el ? (el.value || '').trim() : '';
  };
  const getISO = (field) => {
    const el = row.querySelector(`[data-field="${field}"]`);
    return el ? (el.dataset.iso || '') : '';
  };
  const getRange = (field) => {
    const el = row.querySelector(`[data-field="${field}"]`);
    if (!el) return { mulai: '', selesai: '' };
    return { mulai: el.dataset.isoMulai || '', selesai: el.dataset.isoSelesai || '' };
  };
  const getMulti = (field) => {
    const el = row.querySelector(`[data-field="${field}"]`);
    if (!el) return { nips: [], names: [] };
    let nips = [], names = [];
    try { nips  = JSON.parse(el.dataset.nips  || '[]'); } catch(_) {}
    try { names = JSON.parse(el.dataset.names || '[]'); } catch(_) {}
    return { nips, names };
  };
  const getSingle = (field) => {
    const el = row.querySelector(`[data-field="${field}"]`);
    if (!el) return { nip: '', nama: '' };
    return { nip: el.dataset.nip || '', nama: el.dataset.nama || '' };
  };

  const waktu = getRange('waktu');
  const peg   = getMulti('pegawai_multi');
  const ttd   = getSingle('penandatangan');

  return {
    nomor_surat:       get('nomor_surat'),
    tanggal_surat:     getISO('tanggal_surat'),
    tanggal_berangkat: waktu.mulai,
    tanggal_kembali:   waktu.selesai,
    perihal:           get('perihal'),
    tujuan:            get('tujuan'),
    pegawai_nip:       peg.nips,
    pegawai_list:      peg.names,
    menimbang_custom:  get('menimbang_custom'),
    alat_angkutan:     get('alat_angkutan'),
    pembebanan:        get('pembebanan'),
    penandatangan_nip:    ttd.nip,
    penandatangan_nama:   ttd.nama,
    tempat_terbit:        'Waisai',
    tipe:                 get('tipe') || null,
  };
}

function validateApproveFields(values) {
  const checks = [
    ['nomor_surat',        'Nomor Surat'],
    ['tanggal_surat',      'Tgl Surat'],
    ['tanggal_berangkat',  'Waktu Pelaksanaan'],
    ['perihal',            'Perihal'],
    ['tujuan',             'Tempat Tujuan'],
    ['menimbang_custom',   'Menimbang'],
    ['alat_angkutan',      'Alat Angkutan'],
    ['pembebanan',         'POK'],
    ['penandatangan_nama', 'Penandatangan'],
    ['tipe',               'Tipe Surat'],
  ];
  const errors = [], errFields = [];
  checks.forEach(([k, label]) => {
    const v = values[k];
    if (!v || (Array.isArray(v) && !v.length)) { errors.push(label); errFields.push(k); }
  });
  if (!values.pegawai_nip.length && !values.pegawai_list.length) {
    errors.push('Nama Pegawai');
    errFields.push('pegawai_multi');
  }
  // Validasi format MAK Pembebanan — hanya jika field-nya terisi.
  // Kalau kosong, sudah di-tangkap oleh required-check di atas.
  // Format wajib: 054.01.GG.2910.BMA.006.054.A.524119
  if (values.pembebanan && !parseMAK(values.pembebanan)) {
    errors.push('Format POK tidak valid (contoh: 054.01.GG.2910.BMA.006.054.A.524119)');
    if (!errFields.includes('pembebanan')) errFields.push('pembebanan');
  }
  return { errors, errFields };
}

function highlightRowFieldErrors(suratId, fields) {
  const row = document.querySelector(`tr[data-surat-id="${suratId}"]`);
  if (!row) return;
  row.querySelectorAll('.xls-cell.err, .pg-cell.err').forEach(el => el.classList.remove('err'));

  const FIELD_DOM_MAP = {
    tanggal_berangkat: 'waktu',
    pegawai_nip:       'pegawai_multi',
    pegawai_list:      'pegawai_multi',
    penandatangan_nama: 'penandatangan',
    penandatangan_nip:  'penandatangan',
  };

  fields.forEach(f => {
    const domField = FIELD_DOM_MAP[f] || f;
    const el = row.querySelector(`[data-field="${domField}"]`);
    if (el) el.classList.add('err');
  });

  if (fields.length) {
    const first = row.querySelector('.xls-cell.err, .pg-cell.err');
    if (first) {
      first.scrollIntoView({ behavior: 'smooth', block: 'center', inline: 'center' });
      setTimeout(() => {
        if (first.tagName === 'INPUT' || first.tagName === 'TEXTAREA') first.focus();
        else {
          const inp = first.querySelector('input, textarea, .pg-input');
          if (inp) inp.focus();
        }
      }, 300);
    }
  }
}


/* ════════════════════════════════════════════════════════════════════
   MODAL HELPERS
═══════════════════════════════════════════════════════════════════════ */
function openModal(id) { document.getElementById(id).style.display = 'flex'; }
function closeModal(id) {
  const el = document.getElementById(id);
  if (el) el.style.display = 'none';
  // Cleanup file preview di Supabase Storage saat modal preview ditutup.
  // Delay 60 detik — kalau user buka ulang cepat, tidak perlu re-upload.
  if (id === 'modal-preview' && _previewUploadedPath) {
    const path = _previewUploadedPath;
    _previewUploadedPath = null;
    setTimeout(() => deletePreviewFile(path), 60_000);
  }
}

function showPageAlert(msg, type='error') {
  document.getElementById('page-alert-icon').textContent = type === 'success' ? '✅' : '⚠️';
  document.getElementById('page-alert-text').textContent = msg;
  document.getElementById('page-alert').className = `alert ${type} show`;
  document.getElementById('page-alert').scrollIntoView({ behavior: 'smooth', block: 'nearest' });
  setTimeout(() => { document.getElementById('page-alert').className = 'alert'; }, 5500);
}

/* ════════════════════════════════════════════════════════════════════
/* ════════════════════════════════════════════════════════════════════
   APPROVE
═══════════════════════════════════════════════════════════════════════ */
function openApprove(id) {
  selectedId = id;
  const s = suratMap[id]; if (!s) return;
  const values = collectRowFields(id);
  if (!values) return;

  const { errors, errFields } = validateApproveFields(values);
  if (errors.length) {
    highlightRowFieldErrors(id, errFields);
    showPageAlert(`⚠️ Lengkapi dulu di tabel: ${errors.join(', ')}`, 'error');
    return;
  }

  document.getElementById('approve-perihal').textContent = values.perihal || s.perihal || '—';

  const ttdJabatan = lookupJabatan(values.penandatangan_nip, values.tanggal_surat);
  const nomorFull  = buildNomorSuratFull(values.nomor_surat, values.tanggal_surat);

  document.getElementById('approve-preview').innerHTML = `
    <div class="approve-preview-row"><strong>Nomor</strong><span style="font-family:ui-monospace,monospace;font-size:11px">${esc(nomorFull)}</span></div>
    <div class="approve-preview-row"><strong>Tgl Surat</strong><span>${esc(fmtTgl(values.tanggal_surat))}</span></div>
    <div class="approve-preview-row"><strong>Waktu</strong><span>${esc(fmtWaktu(values.tanggal_berangkat, values.tanggal_kembali))}</span></div>
    <div class="approve-preview-row"><strong>Perihal</strong><span>${esc(values.perihal)}</span></div>
    <div class="approve-preview-row"><strong>Tujuan</strong><span>${esc(values.tujuan)}</span></div>
    <div class="approve-preview-row"><strong>Pegawai</strong><span>${esc(values.pegawai_list.join(', '))}</span></div>
    <div class="approve-preview-row"><strong>Menimbang</strong><span>${esc(values.menimbang_custom)}</span></div>
    <div class="approve-preview-row"><strong>Alat Angkutan</strong><span>${esc(values.alat_angkutan)}</span></div>
    <div class="approve-preview-row"><strong>POK</strong><span>${esc(values.pembebanan)}</span></div>
    <div class="approve-preview-row"><strong>Penandatangan</strong><span>
      ${esc(values.penandatangan_nama)}<br>
      <em style="color:var(--muted);font-style:italic;font-size:11px">${ttdJabatan ? esc(ttdJabatan) : '<span style="color:var(--red)">⚠ Jabatan tidak ditemukan di riwayat (akan ditampilkan "-" di docx)</span>'}</em><br>
      <span style="color:var(--muted);font-size:11px">NIP. ${esc(values.penandatangan_nip)}</span>
    </span></div>
    <div class="approve-preview-row"><strong>Tipe Surat</strong><span>${esc(tipeLabel(values.tipe))}</span></div>
  `;
  document.getElementById('inp-catatan-approve').value = s.catatan_admin || '';
  document.getElementById('approve-alert').className = 'alert';
  openModal('modal-approve');
}

/* ════════════════════════════════════════════════════════════════════
   PRESERVE DIRTY EDITS — helper supaya edit DOM di baris-baris
   "menunggu" lain tidak hilang saat tabel di-render ulang setelah
   sebuah baris di-approve / di-save edit. Bug-fix untuk skenario:
     1. Admin edit banyak baris menunggu sekaligus (belum di-approve)
     2. Approve 1 baris → loadSurat() reload semua → re-render tabel
     3. Edit di baris-baris lain hilang karena DOM dibangun ulang
        dari data server (yang belum punya edit tsb).
   Fix: capture nilai DOM dirty SEBELUM reload, lalu re-apply ke DOM
   SETELAH render ulang.
═══════════════════════════════════════════════════════════════════════ */

/**
 * Snapshot nilai semua cell yang "dirty" di baris status="menunggu".
 * @param {*} excludeId — id baris yang baru di-approve (tidak perlu di-snap
 *                       karena status-nya akan berubah jadi 'selesai').
 * @returns {Object<string, {cells:Array, pegawai:?{nips,names}}>}
 */
function captureMenungguDirty(excludeId = null) {
  const out = {};
  const exclStr = excludeId != null ? String(excludeId) : null;

  document.querySelectorAll('tr[data-status="menunggu"]').forEach(tr => {
    const id = tr.dataset.suratId;
    if (!id) return;
    if (exclStr && String(id) === exclStr) return;

    const entry = { cells: [], pegawai: null };

    // Capture semua xls-cell (input/textarea/select)
    tr.querySelectorAll('.xls-cell').forEach(el => {
      const field = el.dataset.field;
      if (!field) return;
      const cell = { field, value: el.value || '' };
      // Date cells: simpan dataset attr juga (single date dan range)
      if (el.dataset.iso        !== undefined) cell.iso        = el.dataset.iso        || '';
      if (el.dataset.isoMulai   !== undefined) cell.isoMulai   = el.dataset.isoMulai   || '';
      if (el.dataset.isoSelesai !== undefined) cell.isoSelesai = el.dataset.isoSelesai || '';
      entry.cells.push(cell);
    });

    // Capture pg-cell (pegawai multi-select)
    const pg = tr.querySelector('.pg-cell');
    if (pg) {
      try {
        entry.pegawai = {
          nips:  JSON.parse(pg.dataset.nips  || '[]'),
          names: JSON.parse(pg.dataset.names || '[]'),
        };
      } catch(_) { entry.pegawai = { nips: [], names: [] }; }
    }

    out[id] = entry;
  });

  return out;
}

/**
 * Restore snapshot ke DOM. Hanya target baris yang masih ada DAN masih
 * berstatus 'menunggu' (defensive — kalau status berubah krn admin lain
 * meng-approve di window yg sama, kita skip).
 */
function reapplyMenungguDirty(snapshot) {
  if (!snapshot) return;
  Object.keys(snapshot).forEach(id => {
    const tr = document.querySelector(`tr[data-surat-id="${id}"][data-status="menunggu"]`);
    if (!tr) return;
    const entry = snapshot[id];

    // Re-apply cells
    entry.cells.forEach(cell => {
      const el = tr.querySelector(`.xls-cell[data-field="${cell.field}"]`);
      if (!el) return;
      el.value = cell.value;
      if (cell.iso        !== undefined) el.dataset.iso        = cell.iso;
      if (cell.isoMulai   !== undefined) el.dataset.isoMulai   = cell.isoMulai;
      if (cell.isoSelesai !== undefined) el.dataset.isoSelesai = cell.isoSelesai;
      // Untuk select tipe: sinkronkan class tipe-empty (placeholder italic)
      if (el.tagName === 'SELECT' && cell.field === 'tipe') {
        el.classList.toggle('tipe-empty', !cell.value);
      }
    });

    // Re-apply pegawai (re-render visual tag-tag)
    if (entry.pegawai) {
      const pg = tr.querySelector('.pg-cell');
      if (pg) {
        const nips  = entry.pegawai.nips  || [];
        const names = entry.pegawai.names || [];
        pg.dataset.nips  = JSON.stringify(nips);
        pg.dataset.names = JSON.stringify(names);

        // Hapus tag-tag lama
        pg.querySelectorAll('.pg-tag').forEach(t => t.remove());
        // Insert tag baru SEBELUM input
        const inp = pg.querySelector('.pg-input');
        const tagsHTML = nips.map((nip, i) => buildPegTag(nip, names[i] || nip, false)).join('');
        if (inp) {
          inp.insertAdjacentHTML('beforebegin', tagsHTML);
          inp.placeholder = nips.length ? '' : 'Ketik nama...';
        } else {
          pg.insertAdjacentHTML('afterbegin', tagsHTML);
        }
      }
    }
  });
}

async function submitApprove() {
  const values = collectRowFields(selectedId);
  if (!values) return;
  const { errors, errFields } = validateApproveFields(values);
  if (errors.length) {
    document.getElementById('approve-alert-icon').textContent = '⚠️';
    document.getElementById('approve-alert-text').textContent = `Lengkapi: ${errors.join(', ')}`;
    document.getElementById('approve-alert').className = 'alert error show';
    highlightRowFieldErrors(selectedId, errFields);
    setTimeout(() => closeModal('modal-approve'), 1500);
    return;
  }

  const jabatan = lookupJabatan(values.penandatangan_nip, values.tanggal_surat);

  const payload = {
    status: 'selesai',
    nomor_surat:           values.nomor_surat,
    tanggal_surat:         values.tanggal_surat,
    tanggal_berangkat:     values.tanggal_berangkat,
    tanggal_kembali:       values.tanggal_kembali || null,
    perihal:               values.perihal,
    tujuan:                values.tujuan,
    pegawai_nip:           values.pegawai_nip,
    pegawai_list:          values.pegawai_list,
    menimbang_custom:      values.menimbang_custom,
    alat_angkutan:         values.alat_angkutan,
    pembebanan:            values.pembebanan,
    tempat_terbit:         values.tempat_terbit,
    penandatangan_nama:    values.penandatangan_nama,
    penandatangan_nip:     values.penandatangan_nip,
    penandatangan_jabatan: jabatan || '',
    catatan_admin:         document.getElementById('inp-catatan-approve').value.trim() || null,
    tipe:                  values.tipe,
  };

  const btn = document.getElementById('btn-approve-submit');
  btn.disabled = true; btn.classList.add('loading');
  try {
    const res = await fetch(`${SUPABASE_URL}/rest/v1/surat_tugas?id=eq.${selectedId}`, {
      method: 'PATCH',
      headers: { ...H, 'Prefer': 'return=minimal' },
      body: JSON.stringify(payload)
    });
    if (!res.ok) {
      let msg = `HTTP ${res.status}`;
      try { const j = await res.json(); msg = j.message || msg; } catch(_) {}
      throw new Error(msg);
    }

    if (document.getElementById('inp-save-default').checked) {
      const newDefaults = {};
      if (values.alat_angkutan)       newDefaults.alat_angkutan = values.alat_angkutan;
      if (values.pembebanan)          newDefaults.pembebanan    = values.pembebanan;
      if (values.penandatangan_nip)   newDefaults.ttd_nip       = values.penandatangan_nip;
      if (values.penandatangan_nama)  newDefaults.ttd_nama      = values.penandatangan_nama;
      const merged = { ...loadApproveDefaults(), ...newDefaults };
      saveApproveDefaults(merged);
    }

    closeModal('modal-approve');
    showPageAlert('✅ Surat tugas berhasil ditandai selesai.', 'success');
    // BUG FIX (#5): preserve edit di baris menunggu lain — capture sebelum
    // reload, re-apply sesudah render ulang.
    const dirtySnapshot = captureMenungguDirty(selectedId);
    await loadSurat();
    reapplyMenungguDirty(dirtySnapshot);
    // Refresh autocomplete POK — kalau MAK yang baru di-approve belum
    // pernah ada di history, tambahkan ke daftar suggestion.
    loadMAKSuggestions();
  } catch(e) {
    document.getElementById('approve-alert-icon').textContent = '⚠️';
    document.getElementById('approve-alert-text').textContent = `Gagal: ${e.message}`;
    document.getElementById('approve-alert').className = 'alert error show';
  } finally {
    btn.disabled = false; btn.classList.remove('loading');
  }
}

/* ════════════════════════════════════════════════════════════════════
   EDIT ROW (untuk surat yang sudah selesai)
   ─────────────────────────────────────────────────────────────────────
   Mengubah baris readonly (status='selesai') menjadi editable inline
   tanpa membuka modal. Tombol Aksi berubah jadi Simpan / Batal.

   Aturan:
     - Hanya 1 baris yg boleh edit pada satu waktu. Klik Edit pada baris
       lain otomatis menutup edit yang sedang berjalan.
     - Status TIDAK berubah setelah save — tetap 'selesai'. Hanya
       field-field yg di-PATCH.
     - Validasi field memakai validateApproveFields() yang sama dengan
       flow approve, supaya tidak ada celah field invalid lolos.
     - Jabatan penandatangan di-recompute dari riwayat_pegawai berdasar
       NIP+tanggal_surat baru (sama seperti saat approve).
═══════════════════════════════════════════════════════════════════════ */

function enableRowEdit(id) {
  // Kalau ada baris lain yg sedang di-edit, batalkan dulu (revert ke
  // readonly). Pakai cancelRowEdit supaya logikanya konsisten.
  if (editingRowId != null && editingRowId !== id) {
    cancelRowEdit(editingRowId);
  }

  const s = suratMap[id];
  if (!s || s.status !== 'selesai') return;

  editingRowId = id;
  const row = document.querySelector(`tr[data-surat-id="${id}"]`);
  if (!row) return;

  // Re-render hanya baris ini dgn editable=true. Pakai template dummy
  // untuk parse string HTML jadi <tr> element, lalu replace.
  const wrapper = document.createElement('tbody');
  wrapper.innerHTML = renderRowHTML(s);
  const newRow = wrapper.firstElementChild;
  row.replaceWith(newRow);

  // Pasang autoGrow + clear-err handler hanya untuk baris baru ini.
  attachEditableListeners(newRow);

  // Fokus ke field pertama yg editable supaya admin langsung bisa ngetik.
  const firstInput = newRow.querySelector('.xls-cell, .pg-input');
  if (firstInput) {
    setTimeout(() => firstInput.focus(), 50);
  }
}

function cancelRowEdit(id) {
  const s = suratMap[id];
  if (!s) { editingRowId = null; return; }

  editingRowId = null;
  const row = document.querySelector(`tr[data-surat-id="${id}"]`);
  if (!row) return;

  // Render ulang baris dgn editable=false (data lama dari suratMap).
  // Tidak ada perubahan ke DB — ini benar-benar discard.
  const wrapper = document.createElement('tbody');
  wrapper.innerHTML = renderRowHTML(s);
  row.replaceWith(wrapper.firstElementChild);
}

async function saveRowEdit(id) {
  if (editingRowId !== id) return;

  // Pakai collectRowFields & validateApproveFields yang sama dgn approve
  // — biar field rules-nya konsisten (mis. format MAK, required field).
  const values = collectRowFields(id);
  if (!values) return;
  const { errors, errFields } = validateApproveFields(values);
  if (errors.length) {
    highlightRowFieldErrors(id, errFields);
    showPageAlert(`⚠️ Lengkapi dulu: ${errors.join(', ')}`, 'error');
    return;
  }

  // Recompute jabatan penandatangan berdasarkan tanggal_surat (mungkin
  // berubah) — sama seperti di submitApprove.
  const jabatan = lookupJabatan(values.penandatangan_nip, values.tanggal_surat);

  // Payload TANPA field 'status' — status tetap 'selesai'.
  const payload = {
    nomor_surat:           values.nomor_surat,
    tanggal_surat:         values.tanggal_surat,
    tanggal_berangkat:     values.tanggal_berangkat,
    tanggal_kembali:       values.tanggal_kembali || null,
    perihal:               values.perihal,
    tujuan:                values.tujuan,
    pegawai_nip:           values.pegawai_nip,
    pegawai_list:          values.pegawai_list,
    menimbang_custom:      values.menimbang_custom,
    alat_angkutan:         values.alat_angkutan,
    pembebanan:            values.pembebanan,
    tempat_terbit:         values.tempat_terbit,
    penandatangan_nama:    values.penandatangan_nama,
    penandatangan_nip:     values.penandatangan_nip,
    penandatangan_jabatan: jabatan || '',
    tipe:                  values.tipe,
  };

  const row = document.querySelector(`tr[data-surat-id="${id}"]`);
  const btn = row ? row.querySelector('.btn-save-edit') : null;
  if (btn) { btn.disabled = true; btn.classList.add('loading'); }

  try {
    const res = await fetch(`${SUPABASE_URL}/rest/v1/surat_tugas?id=eq.${id}`, {
      method: 'PATCH',
      headers: { ...H, 'Prefer': 'return=minimal' },
      body: JSON.stringify(payload)
    });
    if (!res.ok) {
      let msg = `HTTP ${res.status}`;
      try { const j = await res.json(); msg = j.message || msg; } catch(_) {}
      throw new Error(msg);
    }

    // Sukses — exit edit mode & reload data dari server agar in-memory
    // suratMap ter-sinkron (termasuk updated_at, dll).
    editingRowId = null;
    showPageAlert('✅ Perubahan berhasil disimpan.', 'success');
    // BUG FIX (#5): preserve edit di baris menunggu lain — capture sebelum
    // reload, re-apply sesudah render ulang. (Baris yg di-edit di sini
    // sendiri ber-status 'selesai', jadi tidak ter-include di dirty
    // snapshot — exclude id sebagai safety net.)
    const dirtySnapshot = captureMenungguDirty(id);
    await loadSurat();
    reapplyMenungguDirty(dirtySnapshot);
    // Refresh autocomplete POK kalau MAK pembebanan diganti — sama
    // perlakuannya dgn submitApprove.
    loadMAKSuggestions();
  } catch(e) {
    showPageAlert(`Gagal menyimpan: ${e.message}`, 'error');
    if (btn) { btn.disabled = false; btn.classList.remove('loading'); }
  }
}


/* ════════════════════════════════════════════════════════════════════
   LOOKUP JABATAN dari riwayat_pegawai
═══════════════════════════════════════════════════════════════════════ */
function lookupJabatan(nip, tglSuratIso) {
  if (!nip || !tglSuratIso) return '';
  const candidates = riwayatPegawai
    .filter(r => String(r.pegawai_nip || '').trim() === String(nip).trim())
    .filter(r => r.tmt && r.tmt <= tglSuratIso)
    .sort((a, b) => (b.tmt || '').localeCompare(a.tmt || ''));
  if (candidates.length) return candidates[0].jabatan || '';
  const peg = pegawaiByNIP[nip];
  if (peg && peg.NAMA) {
    const candByName = riwayatPegawai
      .filter(r => (r.nama || '').trim().toLowerCase() === (peg.NAMA || '').trim().toLowerCase())
      .filter(r => r.tmt && r.tmt <= tglSuratIso)
      .sort((a, b) => (b.tmt || '').localeCompare(a.tmt || ''));
    if (candByName.length) return candByName[0].jabatan || '';
  }
  return '';
}

function lookupJabatanForSurat(s) {
  if (s.penandatangan_jabatan) return s.penandatangan_jabatan;
  return lookupJabatan(s.penandatangan_nip, s.tanggal_surat);
}

/* ════════════════════════════════════════════════════════════════════
   FORMAT NOMOR SURAT LENGKAP
═══════════════════════════════════════════════════════════════════════ */
function buildNomorSuratFull(nomor, tglSuratIso) {
  if (!nomor) return '....................';
  const d = parseISODate(tglSuratIso);
  const mm = d ? String(d.getMonth() + 1).padStart(2, '0') : '..';
  const yyyy = d ? String(d.getFullYear()) : '....';
  return `B-${nomor}/668870-92800/KP-650/${mm}/${yyyy}`;
}


/* ════════════════════════════════════════════════════════════════════
   PREVIEW & DOWNLOAD DOCX
═══════════════════════════════════════════════════════════════════════ */
let currentPreviewSurat = null;

/* ── Office Online Viewer (preview) ──────────────────────────────────
   Bucket Supabase Storage tempat file .docx sementara disimpan.
   File otomatis dihapus saat modal ditutup (delay 60 detik). Orphan
   files (kalau user crash/close paksa) dibersihkan saat halaman load
   via cleanupOrphanPreviewFiles().
─────────────────────────────────────────────────────────────────── */
const PREVIEW_BUCKET = 'surat-tugas-preview';
let _previewUploadedPath = null;

async function uploadPreviewDocx(blob, suratId) {
  const filename = `${suratId}_${Date.now()}.docx`;
  const res = await fetch(
    `${SUPABASE_URL}/storage/v1/object/${PREVIEW_BUCKET}/${filename}`,
    {
      method: 'POST',
      headers: {
        apikey: SUPABASE_ANON_KEY,
        Authorization: `Bearer ${SUPABASE_ANON_KEY}`,
        'Content-Type': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        'x-upsert': 'true',
      },
      body: blob,
    }
  );
  if (!res.ok) {
    let msg = `HTTP ${res.status}`;
    try { const err = await res.json(); msg = err.message || err.error || msg; } catch(_) {}
    throw new Error(`Upload preview gagal: ${msg}`);
  }
  return filename;
}

async function getPreviewSignedUrl(filename, expiresInSec = 3600) {
  const res = await fetch(
    `${SUPABASE_URL}/storage/v1/object/sign/${PREVIEW_BUCKET}/${filename}`,
    {
      method: 'POST',
      headers: { ...H },
      body: JSON.stringify({ expiresIn: expiresInSec }),
    }
  );
  if (!res.ok) throw new Error(`Gagal membuat signed URL (HTTP ${res.status})`);
  const data = await res.json();
  // signedURL format: "/object/sign/{bucket}/{path}?token=..."
  return `${SUPABASE_URL}/storage/v1${data.signedURL}`;
}

async function deletePreviewFile(filename) {
  if (!filename) return;
  try {
    await fetch(`${SUPABASE_URL}/storage/v1/object/${PREVIEW_BUCKET}/${filename}`, {
      method: 'DELETE',
      headers: { apikey: SUPABASE_ANON_KEY, Authorization: `Bearer ${SUPABASE_ANON_KEY}` },
    });
  } catch (e) { console.warn('[NOVA] Cleanup preview file gagal:', e); }
}

/* Defensive cleanup: hapus file preview lama (>1 jam) yang mungkin
   ter-orphan kalau user close paksa. Fire-and-forget, dipanggil di init(). */
async function cleanupOrphanPreviewFiles() {
  try {
    const res = await fetch(`${SUPABASE_URL}/storage/v1/object/list/${PREVIEW_BUCKET}`, {
      method: 'POST',
      headers: { ...H },
      body: JSON.stringify({ limit: 100, prefix: '' }),
    });
    if (!res.ok) return;
    const files = await res.json();
    const oneHourAgo = Date.now() - 3600_000;
    files.forEach(f => {
      // Filename pattern: "{suratId}_{timestamp}.docx"
      const m = f.name && f.name.match(/_(\d+)\.docx$/);
      if (m && parseInt(m[1]) < oneHourAgo) deletePreviewFile(f.name);
    });
  } catch (e) { /* fire-and-forget */ }
}

function ensureLibrariesLoaded() {
  if (typeof window.docxGen === 'undefined' && typeof window.docx === 'undefined') {
    throw new Error('Library docx gagal dimuat. Periksa koneksi internet/firewall, lalu refresh halaman.');
  }
  if (typeof window.docxPreview === 'undefined' || typeof window.docxPreview.renderAsync !== 'function') {
    throw new Error('Library docx-preview gagal dimuat. Periksa koneksi dan refresh halaman.');
  }
  if (typeof saveAs === 'undefined') {
    throw new Error('Library FileSaver gagal dimuat. Refresh halaman.');
  }
}

function buildFileName(surat) {
  const base = (surat.nomor_surat || surat.id).toString().replace(/[^a-zA-Z0-9-_]/g, '-');
  return `Surat-Tugas_${base}.docx`;
}

async function openPreview(suratId) {
  const surat = suratMap[suratId];
  if (!surat) return;
  if (surat.status !== 'selesai') {
    showPageAlert('Surat hanya bisa di-preview jika sudah selesai.', 'error');
    return;
  }
  currentPreviewSurat = surat;
  openModal('modal-preview');

  // Bersihkan file preview sebelumnya kalau user buka ulang dengan cepat
  if (_previewUploadedPath) {
    deletePreviewFile(_previewUploadedPath);
    _previewUploadedPath = null;
  }

  const container = document.getElementById('preview-container');
  container.innerHTML = `<div class="preview-loading">
    <div class="preview-loading-spin"></div>
    <div>Menyiapkan dokumen…</div>
    <div style="font-size:11px;color:var(--muted);margin-top:6px">
      Render via Microsoft Word Online (5–15 detik pertama kali)
    </div>
  </div>`;

  try {
    ensureLibrariesLoaded();
    const blob = await buildSuratTugasDoc(surat);
    const filename = await uploadPreviewDocx(blob, surat.id);
    _previewUploadedPath = filename;
    const fileUrl = await getPreviewSignedUrl(filename, 3600);

    const viewerUrl = `https://view.officeapps.live.com/op/embed.aspx?src=${encodeURIComponent(fileUrl)}`;
    container.innerHTML = `
      <iframe src="${viewerUrl}"
        style="width:100%;height:80vh;min-height:600px;border:0;display:block;background:#fff"
        allow="fullscreen"
        sandbox="allow-scripts allow-same-origin allow-forms allow-popups"></iframe>`;
  } catch (e) {
    console.error(e);
    container.innerHTML = `<div class="preview-error">
      <div class="preview-error-icon">⚠️</div>
      <div><strong>Gagal memuat preview</strong></div>
      <div style="font-size:12px;margin-top:8px;color:var(--muted)">${esc(e.message)}</div>
    </div>`;
  }
}

async function downloadSuratTugas(suratId) {
  const surat = suratMap[suratId];
  if (!surat) return;
  if (surat.status !== 'selesai') {
    showPageAlert('Surat hanya bisa di-download jika sudah selesai.', 'error'); return;
  }
  try {
    ensureLibrariesLoaded();
    const blob = await buildSuratTugasDoc(surat);
    saveAs(blob, buildFileName(surat));
    showPageAlert(`📥 Berhasil di-download: ${buildFileName(surat)}`, 'success');
  } catch(e) {
    console.error(e);
    showPageAlert(`Gagal download: ${e.message}`, 'error');
  }
}

/* ════════════════════════════════════════════════════════════════════
   BULK DOWNLOAD (Tugas #7) — admin pilih banyak surat selesai
                              dan download semua sekaligus
   ─────────────────────────────────────────────────────────────────────
   Approach: sequential download dengan jeda kecil antar file. Browser
   bisa memunculkan dialog "Site is requesting to download multiple
   files" — admin tinggal Allow.

   Flow:
     1. Master checkbox di header → toggle semua .bulk-dl-check
        (yang TIDAK disabled — yaitu hanya status='selesai').
     2. Setiap perubahan checkbox → updateBulkDownloadCounter()
        update label tombol dengan count selected.
     3. Klik "Download Terpilih" → bulkDownloadSelected() loop
        sequentially: generate doc + saveAs + jeda 400ms.
═══════════════════════════════════════════════════════════════════════ */

/**
 * Update label & disabled state tombol bulk download berdasarkan
 * jumlah checkbox yang dicentang. Juga sinkron master checkbox state.
 */
function updateBulkDownloadCounter() {
  const all      = document.querySelectorAll('.bulk-dl-check');
  const checked  = document.querySelectorAll('.bulk-dl-check:checked');
  const btn      = document.getElementById('btn-bulk-download');
  const master   = document.getElementById('bulk-dl-master');

  if (btn) {
    if (checked.length === 0) {
      btn.textContent = '📥 Download Terpilih';
      btn.disabled    = true;
    } else {
      btn.textContent = `📥 Download Terpilih (${checked.length})`;
      btn.disabled    = false;
    }
  }

  // Master checkbox state: checked kalau semua di-check, indeterminate
  // kalau sebagian, unchecked kalau tidak ada.
  if (master) {
    if (all.length === 0 || checked.length === 0) {
      master.checked       = false;
      master.indeterminate = false;
    } else if (checked.length === all.length) {
      master.checked       = true;
      master.indeterminate = false;
    } else {
      master.checked       = false;
      master.indeterminate = true;
    }
  }
}

/** Toggle semua checkbox baris selesai mengikuti state master checkbox. */
function toggleBulkDownloadAll(checked) {
  document.querySelectorAll('.bulk-dl-check').forEach(c => { c.checked = !!checked; });
  updateBulkDownloadCounter();
}

/**
 * Loop checked surat dan trigger download .docx satu per satu.
 * Dipanggil dari onclick tombol "Download Terpilih".
 */
async function bulkDownloadSelected() {
  const checked = Array.from(document.querySelectorAll('.bulk-dl-check:checked'));
  const ids = checked.map(c => parseInt(c.dataset.suratId, 10)).filter(Number.isFinite);
  if (!ids.length) {
    showPageAlert('Pilih minimal 1 surat untuk di-download.', 'error');
    return;
  }

  const btn = document.getElementById('btn-bulk-download');
  const originalLabel = btn ? btn.textContent : '';
  if (btn) { btn.disabled = true; }

  let success = 0;
  const failures = [];

  // Pre-warm library sekali saja (bukan per-iteration)
  try { ensureLibrariesLoaded(); } catch(_) {}

  for (let i = 0; i < ids.length; i++) {
    const id    = ids[i];
    const surat = suratMap[id];

    // Update progress di tombol
    if (btn) btn.textContent = `📥 ${i + 1}/${ids.length}…`;

    if (!surat || surat.status !== 'selesai') {
      failures.push(`#${id}: status bukan 'selesai'`);
      continue;
    }

    try {
      const blob = await buildSuratTugasDoc(surat);
      saveAs(blob, buildFileName(surat));
      success++;
      // Jeda antar download supaya browser tidak block multi-trigger.
      // 400ms umumnya cukup untuk Chrome/Edge/Firefox.
      if (i < ids.length - 1) {
        await new Promise(r => setTimeout(r, 400));
      }
    } catch(e) {
      console.error(`[NOVA] bulk download gagal id=${id}:`, e);
      failures.push(`${surat.perihal || 'surat #' + id}: ${e.message}`);
    }
  }

  // Restore tombol & uncheck semua checkbox setelah selesai
  if (btn) { btn.textContent = originalLabel; btn.disabled = false; }
  document.querySelectorAll('.bulk-dl-check:checked').forEach(c => { c.checked = false; });
  updateBulkDownloadCounter();

  // Tampilkan hasil
  if (failures.length === 0) {
    showPageAlert(`✅ ${success} surat berhasil di-download.`, 'success');
  } else if (success === 0) {
    showPageAlert(`Gagal download semua: ${failures.slice(0, 3).join('; ')}${failures.length > 3 ? '…' : ''}`, 'error');
  } else {
    showPageAlert(`⚠️ ${success} berhasil, ${failures.length} gagal. Cek console untuk detail.`, 'error');
  }
}

async function downloadFromPreview() {
  if (!currentPreviewSurat) return;
  try {
    ensureLibrariesLoaded();
    const blob = await buildSuratTugasDoc(currentPreviewSurat);
    saveAs(blob, buildFileName(currentPreviewSurat));
    showPageAlert(`📥 Berhasil di-download: ${buildFileName(currentPreviewSurat)}`, 'success');
    closeModal('modal-preview');
  } catch(e) {
    console.error(e);
    showPageAlert(`Gagal download: ${e.message}`, 'error');
  }
}

function printFromPreview() {
  // DEPRECATED: tombol "🖨 Print" sudah diganti dengan "📄 Buka di Word & Print".
  // Fungsi ini di-keep untuk backward compatibility kalau masih ada caller lain.
  // Forward ke implementasi baru.
  openInWordForPrint();
}

/* ════════════════════════════════════════════════════════════════════
   OPEN IN MICROSOFT WORD (untuk print yang reliable)

   Pendekatan: pakai protocol handler "ms-word:" yang di-handle Microsoft
   Word desktop. Word akan download file dari signed URL Supabase, buka
   di mode Protected View. Admin tinggal Ctrl+P.

   Kenapa pakai ini, bukan window.print() biasa?
   - Office Online iframe = cross-origin, window.print() tidak bisa.
   - Office Online native print sering diblokir Edge tracking prevention
     atau ad-blockers (ERR_BLOCKED_BY_CLIENT).
   - Protocol handler ms-word: = pure native handoff, tidak terkena
     security policy browser.

   Syarat: admin harus punya Microsoft Word desktop terinstall (umumnya
   ya di kantor BPS). Browser akan tampilkan dialog "Open Microsoft Word?"
   sekali — admin klik Allow (atau centang "Always allow" supaya tidak
   muncul lagi).
═══════════════════════════════════════════════════════════════════════ */
async function openInWordForPrint() {
  if (!currentPreviewSurat) {
    showPageAlert('Belum ada surat yang sedang di-preview.', 'error');
    return;
  }

  const btn = document.getElementById('btn-open-in-word');
  if (btn) { btn.disabled = true; btn.textContent = '⏳ Menyiapkan…'; }

  try {
    let signedUrl;

    // Re-use file yang sudah di-upload untuk preview kalau masih ada.
    // Kalau tidak ada (mis. user direct call atau preview gagal), upload ulang.
    if (_previewUploadedPath) {
      signedUrl = await getPreviewSignedUrl(_previewUploadedPath, 3600);
    } else {
      ensureLibrariesLoaded();
      const blob = await buildSuratTugasDoc(currentPreviewSurat);
      const filename = await uploadPreviewDocx(blob, currentPreviewSurat.id);
      _previewUploadedPath = filename;
      signedUrl = await getPreviewSignedUrl(filename, 3600);
    }

    // Format protocol handler Word:
    //   ms-word:ofe|u|<URL>      → Open For Editing
    //   ms-word:ofv|u|<URL>      → Open For Viewing (read-only, lebih cocok untuk print)
    // Kami pakai 'ofe' supaya admin bisa langsung edit kecil kalau ada typo
    // sebelum print, tanpa harus klik "Enable Editing".
    const wordProtocolUrl = `ms-word:ofe|u|${signedUrl}`;

    // Trigger protocol handler. Browser akan tampilkan dialog konfirmasi
    // "Open Microsoft Word?" — admin klik Allow.
    window.location.href = wordProtocolUrl;

    // Fallback notice: kalau Word tidak terinstall, dialog tidak akan muncul
    // dan tidak terjadi apa-apa. Kasih hint setelah 2 detik.
    setTimeout(() => {
      showPageAlert(
        'Jika Microsoft Word tidak terbuka, pastikan Word terinstall di komputer Anda. ' +
        'Sebagai alternatif, klik tombol "📥 Download .docx" lalu buka manual.',
        'success'
      );
    }, 2000);

    // Tutup modal preview otomatis setelah 1 detik supaya admin fokus ke Word
    setTimeout(() => closeModal('modal-preview'), 1000);

  } catch (e) {
    console.error('[NOVA] openInWordForPrint error:', e);
    showPageAlert(`Gagal menyiapkan dokumen untuk Word: ${e.message}`, 'error');
  } finally {
    if (btn) {
      setTimeout(() => {
        btn.disabled = false;
        btn.textContent = '📄 Buka di Word & Print';
      }, 1500);
    }
  }
}

/* ════════════════════════════════════════════════════════════════════
   GENERATOR DOCX
═══════════════════════════════════════════════════════════════════════ */
const MENGINGAT_ITEMS = [
  'Undang-Undang Nomor 16 Tahun 1997 tentang Statistik;',
  'Peraturan Pemerintah Nomor 51 Tahun 1999 tentang Penyelenggaraan Statistik;',
  'Peraturan Presiden Nomor 1 Tahun 2025 tentang Perubahan atas Peraturan Presiden Nomor 86 Tahun 2007 tentang Badan Pusat Statistik;',
  'Peraturan Badan Pusat Statistik Nomor 5 Tahun 2019 tentang Tata Naskah Dinas di Lingkungan Badan Pusat Statistik;',
  'Peraturan Badan Pusat Statistik Nomor 3 Tahun 2025 tentang Perubahan atas Peraturan Badan Pusat Statistik Nomor 5 Tahun 2023 tentang Organisasi dan Tata Kerja Badan Pusat Statistik Provinsi dan Badan Pusat Statistik Kabupaten/Kota;',
];

function buildMenimbang(custom) {
  return `bahwa untuk kepentingan administrasi kegiatan ${custom || '[kegiatan]'}, Kepala Badan Pusat Statistik Kabupaten Raja Ampat perlu menetapkan Surat Tugas;`;
}

function getPegawaiInfoForDoc(nip, fallbackName, tglSuratIso) {
  const peg = pegawaiByNIP[String(nip || '').trim()];
  const nama = (peg && peg.NAMA) || fallbackName || '-';
  const jabatan = lookupJabatan(nip, tglSuratIso);
  return { nama, jabatan: jabatan || '' };
}

function base64ToUint8Array(base64) {
  const cleaned = base64.replace(/^data:image\/[a-z]+;base64,/i, '');
  const raw = atob(cleaned);
  const arr = new Uint8Array(raw.length);
  for (let i = 0; i < raw.length; i++) arr[i] = raw.charCodeAt(i);
  return arr;
}

/* ════════════════════════════════════════════════════════════════════
   LEGACY BUILDER — build-from-code pakai docx-js.
   Dipertahankan sebagai fallback kalau template docxtemplater gagal load.
   Tidak dipanggil langsung — lihat buildSuratTugasDoc() di bawah.
════════════════════════════════════════════════════════════════════ */
async function buildSuratTugasDocLegacy(data) {
  const docxLib = window.docxGen || window.docx;
  const {
    Document, Packer, Paragraph, TextRun, ImageRun,
    Table, TableRow, TableCell,
    AlignmentType, BorderStyle, WidthType, VerticalAlign, HeightRule,
    PageBreak,
  } = docxLib;

  const p = (text, opts = {}) => new Paragraph({
    alignment: opts.align || AlignmentType.LEFT,
    spacing: { after: opts.spaceAfter != null ? opts.spaceAfter : 100, line: opts.line || 276 },
    indent: opts.indent,
    children: [new TextRun({
      text: text,
      bold: opts.bold || false,
      italics: opts.italic || false,
      underline: opts.underline || undefined,
      size: opts.size || 22,
      font: opts.font || 'Times New Roman',
    })],
  });
  const empty = (sa) => new Paragraph({ spacing:{after:sa||0}, children:[new TextRun('')] });

  const NO_BORDER = {
    top:{style:BorderStyle.NONE,size:0,color:'FFFFFF'},
    bottom:{style:BorderStyle.NONE,size:0,color:'FFFFFF'},
    left:{style:BorderStyle.NONE,size:0,color:'FFFFFF'},
    right:{style:BorderStyle.NONE,size:0,color:'FFFFFF'},
    insideHorizontal:{style:BorderStyle.NONE,size:0,color:'FFFFFF'},
    insideVertical:{style:BorderStyle.NONE,size:0,color:'FFFFFF'},
  };
  const cell = (children, opts = {}) => new TableCell({
    width: opts.width ? { size: opts.width, type: WidthType.PERCENTAGE } : undefined,
    verticalAlign: opts.valign || VerticalAlign.TOP,
    borders: NO_BORDER,
    children: Array.isArray(children) ? children : [children],
  });

  const tempat = data.tempat_terbit || 'Waisai';

  let ttdJabatan = data.penandatangan_jabatan;
  if (!ttdJabatan) ttdJabatan = lookupJabatan(data.penandatangan_nip, data.tanggal_surat);
  ttdJabatan = ttdJabatan || '-';

  const nomorFull = buildNomorSuratFull(data.nomor_surat, data.tanggal_surat);

  /* ════════════════════════════════════════════════════════════════════
     HALAMAN SPD (Surat Perjalanan Dinas) — TEMPLATE HARDCODE
     Nilai di-hardcode mengikuti template1.docx untuk pengujian visual.
     TODO: nanti ganti nilai-nilai di bawah jadi data dinamis dari `data` & `pegInfo`.
  ════════════════════════════════════════════════════════════════════ */
  const SPD_HC = {
    // ── HEADER & INSTANSI ────────────────────────────────────────────
    nomor:             'B-688/668870-92800/SPPD-PPIS/11/2025',
    lembar:            '',
    instansi_l1:       'Badan Pusat Statistik Kabupaten Raja Ampat',
    instansi_l2:       'Jl. Jend. Ahmad Yani, Waisai',
    instansi_l3:       'Raja Ampat',

    // ── PPK ───────────────────────────────────────────────────────────
    ppk_nama:          'Abdillah Humam, SST',
    ppk_nip:           '199510152018021001',

    // ── BARIS 2-7 TABEL UTAMA ─────────────────────────────────────────
    pegawai_nama:      'Maulana Tahir, SST., M.AP.',
    pangkat_gol:       'Penata / IIIc',
    jabatan_instansi:  'Statistisi Ahli Muda di BPS Kabupaten Raja Ampat',
    tingkat_biaya:     'C',
    maksud:            'Mengikuti Konsultasi Neraca Wilayah dan Analisis Statistik ke BPS Provinsi Papua Barat di Manokwari, Papua Barat',
    alat_angkutan:     'Kendaraan Umum',
    tempat_berangkat:  'Kabupaten Raja Ampat',
    tempat_tujuan:     'Kabupaten Manokwari',
    lama_perjalanan:   '4 (empat) hari',
    tgl_berangkat:     '18 November 2025',
    tgl_kembali:       '21 November 2025',

    // ── BARIS 9 PEMBEBANAN ANGGARAN ──────────────────────────────────
    program_kode:      'GG',
    program_desc:      'Program Penyediaan dan Pelayanan Informasi Statistik',
    kegiatan_kode:     '2899',
    kegiatan_desc:     'Penyediaan dan Pengembangan Statistik Neraca Produksi',
    komponen_kode:     '052',
    komponen_desc:     'Pengumpulan Data',
    instansi_anggaran: 'Badan Pusat Statistik Kabupaten Raja Ampat',
    mata_anggaran:     '524111',
    keterangan_lain:   '',

    // ── FOOTER HALAMAN 2 (DIKELUARKAN) ──────────────────────────────
    dikeluarkan_di:    'Waisai',
    tgl_dikeluarkan:   '14 November 2025',

    // ── HALAMAN 3: BLOK ATAS (KEPALA BPS) ───────────────────────────
    kepala_nama:       'Ir. Nurhaida Sirun',
    kepala_nip:        '196803201994012001',
    hal3_berangkat_dari: 'Kabupaten Raja Ampat',
    hal3_tempat_kedudukan: '(Tempat Kedudukan)',
    hal3_tgl_berangkat:  '18 November 2025',
    hal3_ke:             'Kabupaten Manokwari',

    // ── HALAMAN 3: SEL TABEL 2x2 ────────────────────────────────────
    tiba_tujuan_kota:    'Kabupaten Manokwari',
    tiba_tujuan_tgl:     '19 November 2025',
    berangkat_balik_dari:'Kabupaten Manokwari',
    berangkat_balik_ke:  'Kabupaten Raja Ampat',
    berangkat_balik_tgl: '21 November 2025',
    tiba_kembali_kota:   'Kabupaten Raja Ampat',
    tiba_kembali_tgl:    '21 November 2025',
  };

  const BORDER_ALL = {
    top:              { style: BorderStyle.SINGLE, size: 4, color: '000000' },
    bottom:           { style: BorderStyle.SINGLE, size: 4, color: '000000' },
    left:             { style: BorderStyle.SINGLE, size: 4, color: '000000' },
    right:            { style: BorderStyle.SINGLE, size: 4, color: '000000' },
    insideHorizontal: { style: BorderStyle.SINGLE, size: 4, color: '000000' },
    insideVertical:   { style: BorderStyle.SINGLE, size: 4, color: '000000' },
  };

  const bCell = (children, opts = {}) => new TableCell({
    width: opts.width ? { size: opts.width, type: WidthType.PERCENTAGE } : undefined,
    verticalAlign: opts.valign || VerticalAlign.TOP,
    columnSpan: opts.colSpan || undefined,
    rowSpan:    opts.rowSpan || undefined,
    margins:    { top: 80, bottom: 80, left: 100, right: 100 },
    children:   Array.isArray(children) ? children : [children],
  });

  const pc = (text, opts = {}) => new Paragraph({
    alignment: opts.align || AlignmentType.LEFT,
    spacing:   { after: opts.spaceAfter != null ? opts.spaceAfter : 0, line: opts.line || 240 },
    indent:    opts.indent,
    children:  [new TextRun({
      text:      text || '',
      bold:      opts.bold || false,
      italics:   opts.italic || false,
      underline: opts.underline || undefined,
      size:      opts.size || 20,
      font:      opts.font || 'Times New Roman',
    })],
  });

  function buildHalamanSPD() {
    const ch = [];

    /* ── HEADER: Nomor + Lembar (rata kanan di atas) ───────────────── */
    ch.push(p(`Nomor : ${SPD_HC.nomor}`, { align: AlignmentType.RIGHT, spaceAfter: 0, size: 22 }));
    ch.push(p(`Lembar : ${SPD_HC.lembar}`, { align: AlignmentType.RIGHT, spaceAfter: 160, size: 22 }));

    /* ── BLOK INSTANSI (kiri, 3 baris) ─────────────────────────────── */
    ch.push(pc(SPD_HC.instansi_l1, { spaceAfter: 40, size: 22 }));
    ch.push(pc(SPD_HC.instansi_l2, { spaceAfter: 40, size: 22 }));
    ch.push(pc(SPD_HC.instansi_l3, { spaceAfter: 160, size: 22 }));

    /* ── TABEL UTAMA 10 BARIS ─────────────────────────────────────────
       Layout: kolom kiri (label) 52% | kolom kanan (value) 48% */
    const W_L = 52, W_R = 48;

    const simpleRow = (label, value, opts = {}) => new TableRow({
      children: [
        bCell([pc(label, { spaceAfter: 0, size: 22 })], { width: W_L }),
        bCell([pc(value, { spaceAfter: 0, size: 22, bold: opts.bold })], { width: W_R }),
      ],
    });

    const multiRow = (labelLines, valueLines) => new TableRow({
      children: [
        bCell(labelLines.map(l => pc(l, { spaceAfter: 0, size: 22 })), { width: W_L }),
        bCell(valueLines.map(v => pc(v, { spaceAfter: 0, size: 22 })), { width: W_R }),
      ],
    });

    const rows = [];
    rows.push(simpleRow('1. Pejabat Pembuat Komitmen', SPD_HC.ppk_nama));
    rows.push(simpleRow('2. Nama pegawai yang melaksanakan perjalanan dinas', SPD_HC.pegawai_nama));
    rows.push(multiRow(
      ['3.  a.  Pangkat dan golongan', '     b.  Jabatan/ instansi', '     c.  Tingkat Biaya Perjalanan Dinas'],
      [SPD_HC.pangkat_gol, SPD_HC.jabatan_instansi, SPD_HC.tingkat_biaya],
    ));
    rows.push(simpleRow('4. Maksud perjalanan dinas', SPD_HC.maksud));
    rows.push(simpleRow('5. Alat Angkutan yang dipergunakan', SPD_HC.alat_angkutan));
    rows.push(multiRow(
      ['6.  a.  Tempat keberangkatan', '     b.  Tempat tujuan'],
      [SPD_HC.tempat_berangkat, SPD_HC.tempat_tujuan],
    ));
    rows.push(multiRow(
      ['7.  a.  Lamanya perjalanan Dinas', '     b.  Tanggal Berangkat', '     c.  Tanggal harus kembali/ tiba ditempat baru *)'],
      [SPD_HC.lama_perjalanan, SPD_HC.tgl_berangkat, SPD_HC.tgl_kembali],
    ));

    /* Row 8: Pengikut — kolom kiri = label "8. Pengikut : Nama"
       kolom kanan = sub-tabel 2 kolom "Umur | Hubungan keluarga/keterangan" */
    const pengikutSubTable = new Table({
      borders: NO_BORDER,
      width: { size: 100, type: WidthType.PERCENTAGE },
      rows: [
        new TableRow({ children: [
          cell([pc('Umur',                         { spaceAfter: 0, size: 22 })], { width: 30 }),
          cell([pc('Hubungan keluarga/keterangan', { spaceAfter: 0, size: 22 })], { width: 70 }),
        ]}),
        new TableRow({ children: [
          cell([pc(' ', { spaceAfter: 0, size: 22 })], { width: 30 }),
          cell([pc(' ', { spaceAfter: 0, size: 22 })], { width: 70 }),
        ]}),
      ],
    });
    rows.push(new TableRow({ children: [
      bCell([pc('8. Pengikut :                                Nama', { spaceAfter: 0, size: 22 })], { width: W_L }),
      bCell([pengikutSubTable], { width: W_R }),
    ]}));

    /* Row 9: Pembebanan anggaran — label kiri dengan sub-item,
       value kanan = pasangan kode + desc (Program/Kegiatan/Komponen),
       Instansi, Mata anggaran */
    const labelPembebanan = [
      pc('9. Pembebanan anggaran', { spaceAfter: 40, size: 22 }),
      pc('                                                Program',  { spaceAfter: 40, size: 22 }),
      pc('                                                Kegiatan', { spaceAfter: 40, size: 22 }),
      pc('                                                Komponen', { spaceAfter: 80, size: 22 }),
      pc('    a.  Instansi',      { spaceAfter: 40, size: 22 }),
      pc('    b.  Mata anggaran', { spaceAfter: 0, size: 22 }),
    ];
    // Value kanan: sub-tabel 2 kolom (kode | desc) supaya kode rata kiri, desc wrap
    const valuePembebananTable = new Table({
      borders: NO_BORDER,
      width: { size: 100, type: WidthType.PERCENTAGE },
      rows: [
        new TableRow({ children: [
          cell([pc(SPD_HC.program_kode,  { spaceAfter: 40, size: 22 })], { width: 18 }),
          cell([pc(SPD_HC.program_desc,  { spaceAfter: 40, size: 22 })], { width: 82 }),
        ]}),
        new TableRow({ children: [
          cell([pc(SPD_HC.kegiatan_kode, { spaceAfter: 40, size: 22 })], { width: 18 }),
          cell([pc(SPD_HC.kegiatan_desc, { spaceAfter: 40, size: 22 })], { width: 82 }),
        ]}),
        new TableRow({ children: [
          cell([pc(SPD_HC.komponen_kode, { spaceAfter: 80, size: 22 })], { width: 18 }),
          cell([pc(SPD_HC.komponen_desc, { spaceAfter: 80, size: 22 })], { width: 82 }),
        ]}),
        new TableRow({ children: [
          cell([pc(SPD_HC.instansi_anggaran, { spaceAfter: 40, size: 22 })], { width: 100, colSpan: 2 }),
        ]}),
        new TableRow({ children: [
          cell([pc(SPD_HC.mata_anggaran,     { spaceAfter: 0,  size: 22 })], { width: 100, colSpan: 2 }),
        ]}),
      ],
    });
    rows.push(new TableRow({ children: [
      bCell(labelPembebanan,         { width: W_L }),
      bCell([valuePembebananTable],  { width: W_R }),
    ]}));

    rows.push(simpleRow('10. Keterangan lain-lain', SPD_HC.keterangan_lain || ''));

    ch.push(new Table({
      rows,
      width: { size: 100, type: WidthType.PERCENTAGE },
      borders: BORDER_ALL,
    }));

    /* ── FOOTER: "Tembusan" di kiri, "Dikeluarkan + PPK ttd" di kanan,
       pakai tabel 2 kolom tanpa border supaya sejajar ────────────── */
    ch.push(empty(120));

    const footerKiri = [
      pc('', { spaceAfter: 1400, size: 22 }),                              // spacer — supaya "Tembusan" turun ke bawah
      pc('Tembusan disampaikan kepada', { spaceAfter: 0, size: 22 }),
      pc(': 1.', { spaceAfter: 0, size: 22 }),
      pc('2.',   { spaceAfter: 0, size: 22 }),
    ];
    const footerKanan = [
      pc(`Dikeluarkan di : ${SPD_HC.dikeluarkan_di}`,  { spaceAfter: 40,  size: 22 }),
      pc(`Pada Tanggal   : ${SPD_HC.tgl_dikeluarkan}`, { spaceAfter: 160, size: 22 }),
      pc('PEJABAT PEMBUAT KOMITMEN', { align: AlignmentType.CENTER, spaceAfter: 1400, size: 22 }),
      pc(SPD_HC.ppk_nama, { align: AlignmentType.CENTER, bold: true, underline: { type: 'single' }, spaceAfter: 0, size: 22 }),
      pc(`NIP. ${SPD_HC.ppk_nip}`, { align: AlignmentType.CENTER, spaceAfter: 0, size: 22 }),
    ];
    ch.push(new Table({
      borders: NO_BORDER,
      width: { size: 100, type: WidthType.PERCENTAGE },
      rows: [ new TableRow({ children: [
        cell(footerKiri,  { width: 50 }),
        cell(footerKanan, { width: 50 }),
      ]}) ],
    }));

    ch.push(new Paragraph({ children: [new PageBreak()] }));
    return ch;
  }

  /* ════════════════════════════════════════════════════════════════════
     HALAMAN PENGESAHAN (I, II, III, IV, V) — TEMPLATE HARDCODE
  ════════════════════════════════════════════════════════════════════ */
  function buildHalamanPengesahan() {
    const ch = [];

    /* ── BAGIAN ATAS: Blok "Berangkat dari…" di kolom kanan + TTD Kepala BPS.
       Tanpa border — gunakan tabel 2 kolom (kiri kosong | kanan isi) ── */
    const blokAtasKanan = [
      pc(`Berangkat dari  :   ${SPD_HC.hal3_berangkat_dari}`, { spaceAfter: 40, size: 22 }),
      pc(SPD_HC.hal3_tempat_kedudukan,                         { spaceAfter: 40, size: 22 }),
      pc(`Pada Tanggal     :   ${SPD_HC.hal3_tgl_berangkat}`,  { spaceAfter: 40, size: 22 }),
      pc(`Ke                      :   ${SPD_HC.hal3_ke}`,      { spaceAfter: 200, size: 22 }),
      pc('Kepala BPS Kabupaten Raja Ampat', { spaceAfter: 1400, size: 22 }),
      pc(SPD_HC.kepala_nama, { bold: true, underline: { type: 'single' }, spaceAfter: 0, size: 22 }),
      pc(`NIP. ${SPD_HC.kepala_nip}`, { spaceAfter: 0, size: 22 }),
    ];
    ch.push(new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      borders: NO_BORDER,
      rows: [ new TableRow({ children: [
        cell([pc('', { spaceAfter: 0 })], { width: 50 }),
        cell(blokAtasKanan,                { width: 50 }),
      ]}) ],
    }));

    ch.push(empty(0));

    /* ── TABEL 2×2 (isi laporan tiba/berangkat) ─────────────────────
       Row 1: Tiba tujuan | Berangkat balik
       Row 2: Tiba kembali + TTD PPK | Telah diperiksa + TTD PPK */

    const selTibaTujuan = [
      pc(`Tiba di           :   ${SPD_HC.tiba_tujuan_kota}`, { spaceAfter: 40, size: 22 }),
      pc(`Pada Tanggal  :   ${SPD_HC.tiba_tujuan_tgl}`,      { spaceAfter: 0,  size: 22 }),
    ];
    const selBerangkatBalik = [
      pc(`Berangkat dari  :   ${SPD_HC.berangkat_balik_dari}`, { spaceAfter: 40, size: 22 }),
      pc(`Ke                      :   ${SPD_HC.berangkat_balik_ke}`,  { spaceAfter: 40, size: 22 }),
      pc(`Pada Tanggal     :   ${SPD_HC.berangkat_balik_tgl}`,        { spaceAfter: 0,  size: 22 }),
    ];
    const selTibaKembali = [
      pc(`Tiba di           :   ${SPD_HC.tiba_kembali_kota}`, { spaceAfter: 40, size: 22 }),
      pc('(Tempat Kedudukan)',                                 { spaceAfter: 40, size: 22 }),
      pc(`Pada Tanggal  :   ${SPD_HC.tiba_kembali_tgl}`,       { spaceAfter: 200, size: 22 }),
      pc('Pejabat Pembuat Komitmen', { align: AlignmentType.CENTER, spaceAfter: 1400, size: 22 }),
      pc(SPD_HC.ppk_nama, { align: AlignmentType.CENTER, bold: true, underline: { type: 'single' }, spaceAfter: 0, size: 22 }),
      pc(`NIP. ${SPD_HC.ppk_nip}`, { align: AlignmentType.CENTER, spaceAfter: 0, size: 22 }),
    ];
    const selDiperiksa = [
      pc('Telah diperiksa dengan keterangan bahwa perjalanan tersebut atas perintahnya dan semata-mata untuk kepentingan jabatan dalam waktu yang sesingkat singkatnya',
         { align: AlignmentType.JUSTIFIED, spaceAfter: 200, size: 22 }),
      pc('Pejabat Pembuat Komitmen', { align: AlignmentType.CENTER, spaceAfter: 1400, size: 22 }),
      pc(SPD_HC.ppk_nama, { align: AlignmentType.CENTER, bold: true, underline: { type: 'single' }, spaceAfter: 0, size: 22 }),
      pc(`NIP. ${SPD_HC.ppk_nip}`, { align: AlignmentType.CENTER, spaceAfter: 0, size: 22 }),
    ];

    const rowHeight = HeightRule ? { value: 3000, rule: HeightRule.ATLEAST } : undefined;

    ch.push(new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      borders: BORDER_ALL,
      rows: [
        new TableRow({
          height: rowHeight,
          children: [
            bCell(selTibaTujuan,     { width: 50 }),
            bCell(selBerangkatBalik, { width: 50 }),
          ],
        }),
        new TableRow({
          height: rowHeight,
          children: [
            bCell(selTibaKembali, { width: 50 }),
            bCell(selDiperiksa,   { width: 50 }),
          ],
        }),
      ],
    }));

    /* ── CATATAN LAIN - LAIN + PERHATIAN (di luar tabel) ──────────── */
    ch.push(empty(0));
    ch.push(pc('CATATAN LAIN - LAIN', { spaceAfter: 80, size: 22 }));

    ch.push(new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      borders: NO_BORDER,
      rows: [ new TableRow({ children: [
        cell([pc('PERHATIAN   :', { spaceAfter: 0, size: 22 })], { width: 18 }),
        cell([pc(
          'Pejabat yang berwenang menerbitkan SPD pegawai yang melakukan perjalanan dinas para pejabat yang mengesahkan tanggal berangkat/tiba serta bendaharawan bertanggung jawab berdasarkan peraturan-peraturan keuangan Negara apabila Negara menderita rugi akibat kesalahan, kelalaian dan kealpaannya.',
          { align: AlignmentType.JUSTIFIED, spaceAfter: 0, size: 22 }
        )], { width: 82 }),
      ]}) ],
    }));

    return ch;
  }

  function buildHalamanPegawai(pegInfo, isLast) {
    const ch = [];

    if (typeof LOGO_BPS_BASE64 !== 'undefined' && LOGO_BPS_BASE64) {
      try {
        ch.push(new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { after: 100 },
          children: [new ImageRun({
            data: base64ToUint8Array(LOGO_BPS_BASE64),
            transformation: { width: 90, height: 90 },
          })],
        }));
      } catch(e) { console.warn('Gagal embed logo:', e); }
    }

    ch.push(
      p('BADAN PUSAT STATISTIK', { bold:true, italic:true, align: AlignmentType.CENTER, size:28, spaceAfter:0 }),
      p('KABUPATEN RAJA AMPAT',  { bold:true, italic:true, align: AlignmentType.CENTER, size:28, spaceAfter:300 }),
    );

    ch.push(
      p('SURAT TUGAS', { align: AlignmentType.CENTER, size:24, spaceAfter:60 }),
      p(`NOMOR ${nomorFull}`, { align: AlignmentType.CENTER, size:24, spaceAfter:240 }),
    );

    const menimbangTxt = buildMenimbang(data.menimbang_custom);

    const rowMenimbang = new TableRow({ children: [
      cell(p('Menimbang', { spaceAfter:0 }),  { width:20 }),
      cell(p(':',         { spaceAfter:0 }),  { width:3 }),
      cell(p(menimbangTxt, { align: AlignmentType.JUSTIFIED, spaceAfter:0 }), { width:77 }),
    ]});
    const mengingatRows = MENGINGAT_ITEMS.map((it, idx) => new TableRow({ children: [
      cell(p(idx===0 ? 'Mengingat' : '', { spaceAfter:0 }), { width:20 }),
      cell(p(idx===0 ? ':' : '',          { spaceAfter:0 }), { width:3 }),
      cell(new Paragraph({
        alignment: AlignmentType.JUSTIFIED,
        spacing: { after: 60, line: 276 },
        children: [new TextRun({ text:`${idx+1}. ${it}`, size:22, font:'Times New Roman' })],
      }), { width:77 }),
    ]}));
    ch.push(new Table({ rows: [rowMenimbang, ...mengingatRows], width:{size:100,type:WidthType.PERCENTAGE}, borders:NO_BORDER }));

    ch.push(empty(100));
    ch.push(p('Memberi Tugas', { align: AlignmentType.CENTER, spaceAfter:100 }));

    const untukTxt = `${data.perihal || '-'} di ${data.tujuan || '-'} pada tanggal ${fmtWaktu(data.tanggal_berangkat, data.tanggal_kembali) || '-'}`;

    const detailRows = [
      new TableRow({ children: [
        cell(p('Kepada',  { spaceAfter:0 }), { width:20 }),
        cell(p(':',        { spaceAfter:0 }), { width:3 }),
        cell(p(pegInfo.nama, { spaceAfter:0 }), { width:77 }),
      ]}),
      new TableRow({ children: [
        cell(p('Jabatan', { spaceAfter:0 }), { width:20 }),
        cell(p(':',        { spaceAfter:0 }), { width:3 }),
        cell(p(pegInfo.jabatan || '-', { spaceAfter:0 }), { width:77 }),
      ]}),
      new TableRow({ children: [
        cell(p('Untuk',   { spaceAfter:0 }), { width:20 }),
        cell(p(':',        { spaceAfter:0 }), { width:3 }),
        cell(p(untukTxt, { align: AlignmentType.JUSTIFIED, spaceAfter:0 }), { width:77 }),
      ]}),
    ];
    ch.push(new Table({ rows: detailRows, width:{size:100,type:WidthType.PERCENTAGE}, borders:NO_BORDER }));
    ch.push(empty(100));

    const akhirRows = [
      new TableRow({ children: [
        cell(p('Alat Angkutan', { spaceAfter:0 }),                      { width:20 }),
        cell(p(':',              { spaceAfter:0 }),                      { width:3 }),
        cell(p(data.alat_angkutan || '-', { spaceAfter:0 }),             { width:77 }),
      ]}),
      new TableRow({ children: [
        cell(p('Pembebanan',    { spaceAfter:0 }),                       { width:20 }),
        cell(p(':',              { spaceAfter:0 }),                       { width:3 }),
        cell(p(data.pembebanan || '-', { spaceAfter:0 }),                { width:77 }),
      ]}),
    ];
    ch.push(new Table({ rows: akhirRows, width:{size:100,type:WidthType.PERCENTAGE}, borders:NO_BORDER }));

    ch.push(empty(400));

    const tglSuratFmt = data.tanggal_surat ? fmtTgl(data.tanggal_surat) : fmtTgl(todayISO());
    ch.push(
      p(`${tempat}, ${tglSuratFmt}`,        { align: AlignmentType.RIGHT, spaceAfter:0 }),
      p(ttdJabatan,                          { align: AlignmentType.RIGHT, spaceAfter:900 }),
      p(data.penandatangan_nama || '-',      { align: AlignmentType.RIGHT, underline:{type:'single'}, spaceAfter:0 }),
      p(`NIP. ${data.penandatangan_nip || '-'}`, { align: AlignmentType.RIGHT, spaceAfter:0 }),
    );

    if (!isLast) ch.push(new Paragraph({ children: [new PageBreak()] }));
    return ch;
  }

  const nipList  = Array.isArray(data.pegawai_nip)  ? data.pegawai_nip  : [];
  const nameList = Array.isArray(data.pegawai_list) ? data.pegawai_list : [];

  let pegawaiInfoList = [];
  if (nipList.length) {
    pegawaiInfoList = nipList.map((nip, i) => getPegawaiInfoForDoc(nip, nameList[i], data.tanggal_surat));
  } else if (nameList.length) {
    pegawaiInfoList = nameList.map(n => ({ nama: n, jabatan: '' }));
  } else {
    pegawaiInfoList = [{ nama: '-', jabatan: '' }];
  }

  const allChildren = [];
  pegawaiInfoList.forEach((info, idx) => {
    const isLast = idx === pegawaiInfoList.length - 1;

    // Halaman 1: Surat Tugas (tanpa page break internal — dipindah ke setelah halaman 3)
    buildHalamanPegawai(info, true).forEach(c => allChildren.push(c));
    allChildren.push(new Paragraph({ children: [new PageBreak()] }));

    // Halaman 2: SPD (fungsi ini sudah memasukkan PageBreak di akhir)
    buildHalamanSPD().forEach(c => allChildren.push(c));

    // Halaman 3: Lembar Pengesahan
    buildHalamanPengesahan().forEach(c => allChildren.push(c));

    // Page break antar pegawai (kecuali pegawai terakhir)
    if (!isLast) {
      allChildren.push(new Paragraph({ children: [new PageBreak()] }));
    }
  });

  const doc = new Document({
    creator: 'Portal NOVA',
    title: `Surat Tugas ${data.nomor_surat || data.id}`,
    description: 'Surat Tugas digenerate oleh Portal NOVA',
    styles: { default: { document: { run: { font: 'Times New Roman', size: 22 } } } },
    sections: [{
      properties: { page: { margin: { top:1134, right:1134, bottom:1134, left:1134 } } },
      children: allChildren,
    }],
  });

  return await Packer.toBlob(doc);
}

/* ════════════════════════════════════════════════════════════════════
   HELPERS UNTUK TEMPLATE BARU — MAK Pembebanan, Hari Inclusive, PPK Lookup
   ──────────────────────────────────────────────────────────────────────
   Semua helper di sini dipakai oleh buildTemplateData() di bawah.
   Dipisah agar mudah di-maintain dan ditest secara terpisah.
═══════════════════════════════════════════════════════════════════════ */

/**
 * Format MAK Pembebanan yang valid:
 *   054.01.GG.2910.BMA.006.054.A.524119
 *   └─program─┘ └kgt─┘└kro┘└ro─┘└kmp┘└s┘└─akun─┘
 *
 * Aturan tiap komponen:
 *   - program       : 3 segmen "ddd.dd.GG|WA"
 *   - kegiatan      : 4 digit angka
 *   - kro           : 3 huruf kapital
 *   - ro            : 3 digit angka
 *   - komponen      : 3 digit angka
 *   - sub_komponen  : 1 huruf kapital
 *   - akun          : 6 digit angka
 */
const MAK_REGEX = /^(\d{3}\.\d{2}\.(?:GG|WA))\.(\d{4})\.([A-Z]{3})\.(\d{3})\.(\d{3})\.([A-Z])\.(\d{6})$/;

function parseMAK(str) {
  if (!str) return null;
  const m = String(str).trim().match(MAK_REGEX);
  if (!m) return null;
  return {
    program:      m[1],
    kegiatan:     m[2],
    kro:          m[3],
    ro:           m[4],
    komponen:     m[5],
    sub_komponen: m[6],
    akun:         m[7],
  };
}

/**
 * Bangun string MAK lengkap untuk placeholder {mak_pembebanan} di template.
 * Format: "{program} {kegiatan}.{kro}.{ro}.{komponen}.{sub_komponen}.{akun}"
 * Catatan: ada SPASI antara program dan kegiatan (sesuai spesifikasi).
 */
function formatMAKLengkap(mak) {
  if (!mak) return '';
  return `${mak.program} ${mak.kegiatan}.${mak.kro}.${mak.ro}.${mak.komponen}.${mak.sub_komponen}.${mak.akun}`;
}

/**
 * Bangun nomor SPD: B-{nomor}/668870-92800/SPPD-{kode_mak}/{mm}/{yyyy}
 * kode_mak diturunkan dari kegiatan:
 *   - kegiatan === '2886'  →  'DM2886'
 *   - lainnya              →  'PPIS{kegiatan}'
 */
function buildNomorSPD(nomor, tglSuratIso, mak) {
  if (!nomor || !tglSuratIso || !mak || !mak.kegiatan) return '';
  const d = parseISODate(tglSuratIso);
  if (!d) return '';
  const mm   = String(d.getMonth() + 1).padStart(2, '0');
  const yyyy = String(d.getFullYear());
  const kodeMAK = mak.kegiatan === '2886' ? 'DM2886' : `PPIS${mak.kegiatan}`;
  return `B-${nomor}/668870-92800/SPPD-${kodeMAK}/${mm}/${yyyy}`;
}

/* ════════════════════════════════════════════════════════════════════
   KAMUS POK — lookup deskripsi kode anggaran

   Tabel `kamus_pok` di Supabase menyimpan deskripsi untuk 7 segmen MAK.
   Hierarki: KRO punya parent = kode kegiatan, RO punya parent =
   "kegiatan.kro". Selain itu standalone (parent_kode = NULL).

   Strategi cache: muat sekali per tahun anggaran. Row dengan
   tahun_anggaran NULL = berlaku semua tahun (default). Row dengan tahun
   spesifik = override khusus tahun tersebut. Saat lookup, row tahun
   spesifik selalu menang.
═══════════════════════════════════════════════════════════════════════ */
let _kamusPokCache = null;  // Map<string, string>
let _kamusPokYear  = null;

async function loadKamusPok(tahun) {
  if (_kamusPokCache && _kamusPokYear === tahun) return _kamusPokCache;

  // Ambil row aktif untuk tahun spesifik DAN row default (NULL).
  const url = `${SUPABASE_URL}/rest/v1/kamus_pok` +
    `?aktif=eq.true` +
    `&or=(tahun_anggaran.eq.${tahun},tahun_anggaran.is.null)` +
    `&select=segmen,kode,parent_kode,deskripsi,tahun_anggaran`;

  const res = await fetch(url, { headers: { ...H } });
  if (!res.ok) {
    console.warn(`[NOVA] Gagal load kamus_pok (HTTP ${res.status}). Deskripsi POK akan kosong.`);
    _kamusPokCache = {};
    _kamusPokYear  = tahun;
    return _kamusPokCache;
  }
  const rows = await res.json();

  // Sort: row tahun spesifik di depan supaya menang saat first-wins build.
  rows.sort((a, b) => {
    if (a.tahun_anggaran && !b.tahun_anggaran) return -1;
    if (!a.tahun_anggaran && b.tahun_anggaran) return 1;
    return 0;
  });

  const map = {};
  rows.forEach(r => {
    const key = `${r.segmen}:${r.kode}:${r.parent_kode || ''}`;
    if (!(key in map)) map[key] = r.deskripsi || '';
  });

  _kamusPokCache = map;
  _kamusPokYear  = tahun;
  console.log(`[NOVA] Kamus POK loaded: ${rows.length} rows untuk tahun ${tahun}`);
  return map;
}

function lookupDeskripsi(segmen, kode, parent_kode) {
  if (!_kamusPokCache || !kode) return '';
  const key = `${segmen}:${kode}:${parent_kode || ''}`;
  return _kamusPokCache[key] || '';
}

// Dipanggil dari halaman manajemen-kamus-pok kalau admin edit data.
// Pasang juga listener storage event supaya tab admin lain auto-refresh.
function invalidateKamusPokCache() {
  _kamusPokCache = null;
  _kamusPokYear  = null;
}
window.addEventListener('storage', e => {
  if (e.key === 'nova_kamus_pok_invalidate') invalidateKamusPokCache();
});

/**
 * Hitung lama hari INCLUSIVE.
 *   - Tanggal tunggal     → 1 hari
 *   - 20 s.d. 22 April    → 3 hari (22-20+1)
 */
function hitungHariInclusive(isoMulai, isoSelesai) {
  if (!isoMulai) return 0;
  const ms = parseISODate(isoMulai);
  if (!ms) return 0;
  const se = isoSelesai ? parseISODate(isoSelesai) : ms;
  if (!se) return 1;
  const diff = Math.round((se.getTime() - ms.getTime()) / (1000 * 60 * 60 * 24));
  return Math.max(1, diff + 1);
}

/**
 * Hitung jumlah destinasi dari teks tujuan, untuk loop {#destinasi}
 * di Section II SPD (Tiba di / Berangkat dari … N kali).
 *
 * Logika: count berapa kali kata "Kampung", "Desa", atau "Kelurahan"
 * muncul di teks tujuan (case-insensitive, word boundary).
 *
 * Contoh:
 *   "Kampung A dan Kampung C, Distrik Z; Kampung G, Distrik F;
 *    dan Kampung Y, Distrik X"
 *   → 4 (kata "Kampung" muncul 4 kali)
 *
 * Edge cases:
 * - Word boundary `\b` mencegah false-match (mis. "kampungan" tidak match)
 * - Case-insensitive (KAMPUNG, Kampung, kampung — semua match)
 * - Fallback minimum 1 — kalau tujuan kosong / tidak ada kata kunci,
 *   tetap render 1 baris supaya Section II tidak hilang.
 *
 * Kata kunci ke depan bisa ditambah (Pekon/Nagari/Gampong/dll) dengan
 * meng-extend regex di bawah.
 */
const RE_DESTINASI = /\b(kampung|desa|kelurahan)\b/gi;
function countDestinasi(tujuanText) {
  if (!tujuanText) return 1;
  const matches = String(tujuanText).match(RE_DESTINASI);
  const count = matches ? matches.length : 0;
  return Math.max(1, count);
}

/**
 * Cari NIP berdasarkan nama di tabel riwayat_pegawai.
 * Toleran terhadap suffix gelar setelah koma — mis. "Abdillah Humam, SST"
 * akan match record dengan nama "Abdillah Humam, SST" maupun "Abdillah Humam".
 * Kalau lebih dari satu record (riwayat jabatan banyak), ambil yang pertama
 * — semua record untuk orang yang sama akan punya pegawai_nip yang sama.
 */
function findNipByNama(namaCari) {
  if (!namaCari || !Array.isArray(riwayatPegawai)) return '';
  const target     = String(namaCari).toLowerCase().trim();
  const targetCore = target.split(',')[0].trim(); // tanpa gelar

  const found = riwayatPegawai.find(r => {
    const nama = (r.nama || '').toLowerCase().trim();
    if (!nama) return false;
    const namaCore = nama.split(',')[0].trim();
    return nama === target || namaCore === targetCore;
  });
  return found ? String(found.pegawai_nip || '').trim() : '';
}

/**
 * Konstanta nama PPK — sementara hardcode di kode.
 * Nanti idealnya dipindah ke kolom config / DB agar tidak perlu deploy ulang
 * saat ada pergantian PPK. NIP di-lookup runtime dari riwayat_pegawai.
 */
const PPK_NAMA_DEFAULT = 'Abdillah Humam, SST';


/* ════════════════════════════════════════════════════════════════════
   TEMPLATE-BASED BUILDER — docxtemplater

   URL template & mapping tipe → template ada di TIPE_TO_TEMPLATE
   (didefinisikan di awal file). Loader-nya: loadTemplateBuffer(tipe).

   Tiga template di Supabase Storage:
     - template-surat-tugas-spd-kendaraan-menginap.docx              (T1)
     - template-surat-tugas-lampiran.docx                             (T2)
     - template-surat-tugas-lampiran-spd-kendaraan-menginap.docx     (T3)

   Tipe surat dipilih admin saat approve (kolom `tipe` di tabel persetujuan).
   Lihat TIPE_OPTIONS, TIPE_TO_FLAGS untuk daftar lengkap.
════════════════════════════════════════════════════════════════════ */

/* Helper untuk load script dinamis kalau belum ada di window */
function loadScript(src) {
  return new Promise((resolve, reject) => {
    const s = document.createElement('script');
    s.src = src;
    s.async = false;
    s.onload = () => { console.log('[NOVA] Loaded:', src); resolve(); };
    s.onerror = () => reject(new Error(`Gagal load ${src}`));
    document.head.appendChild(s);
  });
}

async function ensureDocxtemplaterLoaded() {
  // Kalau sudah ada, return langsung
  if ((window.docxtemplater || window.Docxtemplater) && (window.PizZip || window.pizzip)) {
    return true;
  }
  // Coba load dinamis dengan multiple CDN fallback
  const CDNS = [
    { dxt: 'https://unpkg.com/docxtemplater@3.68.5/build/docxtemplater.js',
      pzz: 'https://unpkg.com/pizzip@3.2.0/dist/pizzip.js' },
    { dxt: 'https://cdn.jsdelivr.net/npm/docxtemplater@3.68.5/build/docxtemplater.js',
      pzz: 'https://cdn.jsdelivr.net/npm/pizzip@3.2.0/dist/pizzip.js' },
    { dxt: 'https://cdnjs.cloudflare.com/ajax/libs/docxtemplater/3.68.5/docxtemplater.js',
      pzz: 'https://cdn.jsdelivr.net/npm/pizzip@3.2.0/dist/pizzip.js' },
  ];
  for (const cdn of CDNS) {
    try {
      if (!window.PizZip && !window.pizzip) await loadScript(cdn.pzz);
      if (!window.docxtemplater && !window.Docxtemplater) await loadScript(cdn.dxt);
      if ((window.docxtemplater || window.Docxtemplater) && (window.PizZip || window.pizzip)) {
        return true;
      }
    } catch(e) {
      console.warn('[NOVA] CDN gagal, coba berikutnya:', e.message);
    }
  }
  return false;
}

/* ════════════════════════════════════════════════════════════════════
   TEMPLATE LOADER (tipe-based)
   ─────────────────────────────────────────────────────────────────────
   Cache buffer per-URL supaya template yg sudah pernah di-fetch tidak
   di-download ulang. Mendukung 3 template (T1/T2/T3) yg URL-nya
   ditentukan oleh `data.tipe` lewat helper tipeTemplateUrl() di awal
   file ini.

   Fungsi loadTemplateBuffer(tipe):
     - tipe valid → fetch URL yg sesuai (cache-aware) → return ArrayBuffer
     - tipe null/invalid → throw dengan pesan jelas
═══════════════════════════════════════════════════════════════════════ */

const _templateBufferCache = {};   // url → ArrayBuffer

async function loadTemplateBuffer(tipe) {
  const url = tipeTemplateUrl(tipe);
  if (!url) {
    throw new Error(
      `Tipe surat tidak valid atau belum dipilih (tipe="${tipe}"). ` +
      `Admin wajib memilih tipe sebelum approve. Silakan edit baris ini ` +
      `dan set ulang tipe.`
    );
  }
  if (_templateBufferCache[url]) return _templateBufferCache[url];

  console.log('[NOVA] Memuat template (tipe=' + tipe + ') dari:', url);
  const res = await fetch(url);
  if (!res.ok) {
    throw new Error(
      `Gagal memuat template untuk tipe "${tipe}" (HTTP ${res.status}). ` +
      `Pastikan file ${url} ada di Supabase Storage dan bisa diakses publik.`
    );
  }
  const buf = await res.arrayBuffer();
  _templateBufferCache[url] = buf;
  console.log('[NOVA] Template loaded:', buf.byteLength, 'bytes');
  return buf;
}

/* Helper format tanggal ISO → "DD Nama-Bulan YYYY" (Indonesia) */
function fmtTglId(isoStr) {
  if (!isoStr) return '';
  const d = parseISODate(isoStr);
  if (!d) return '';
  return `${d.getDate()} ${BULAN[d.getMonth()]} ${d.getFullYear()}`;
}

/* Bangun objek mapping data → placeholder template.
   Satu-satunya tempat di mana field dari DB di-mapping ke nama placeholder.
   ──────────────────────────────────────────────────────────────────────
   Daftar placeholder yang dipakai template — terurut sesuai kemunculan:

   HALAMAN 1 (Surat Tugas):
     {nomor_surat}, {menimbang}, {nama}, {jabatan}, {perihal},
     {tempat_tujuan}, {waktu_pelaksanaan}, {tahun}, {mak_pembebanan},
     {tgl_surat}, {jabatan_penandatangan}, {penandatangan},
     {nip_penandatangan}

     CATATAN: jika pegawai_list >= 2, {nama} & {jabatan} otomatis
     terisi "Terlampir" — daftar lengkap dimunculkan di halaman lampiran.

   HALAMAN 2 (SPD):
     {nomor_spd}, {nama_ppk}, {nip}, {pangkat}, {angkutan},
     {hari}, {tanggal_berangkat}, {tanggal_kembali},
     {program}, {kegiatan}, {kro}, {ro}, {komponen}, {sub_komponen}, {akun},
     {des_program}, {des_kegiatan}, {des_kro}, {des_ro},
     {des_komponen}, {des_sub_komponen}, {des_akun}, {nip_ppk}

   HALAMAN 3 (Lampiran — hanya jika ≥2 pegawai):
     {#has_lampiran_st}...{/has_lampiran_st}  ← bungkus seluruh halaman
     {#ul}...{/ul}                            ← bungkus baris tabel
       Field per-iterasi: {no}, {nama_p}, {nip_p},
                          {pangkat_p}, {golongan_p},
                          {jabatan_p}, {bertugas_p}
     Header lampiran ikut pakai: {nomor_surat}, {tgl_surat}, {awalan},
                                 {menimbang},
                                 {jabatan_penandatangan}, {penandatangan},
                                 {nip_penandatangan}

   Catatan tentang kolom yang BELUM ada di Supabase:
     - {pangkat}     → kolom pangkat/golongan belum ada di riwayat_pegawai → ''
     - {pangkat_p}, {golongan_p}, {bertugas_p} → di-kosongkan dulu di lampiran,
                     akan diisi setelah skema DB dilengkapi.
                     {jabatan_p} sudah di-lookup dari riwayat_pegawai.
     - {des_*}       → tabel kamus_pok belum dibuat → semua ''
   Field-field di atas akan otomatis terisi setelah skema DB dilengkapi —
   tinggal lookup di buildTemplateData() ini.
*/
/* ════════════════════════════════════════════════════════════════════
   PATCH 1 — buildTemplateData()
   GANTI fungsi buildTemplateData yang ada di admin-surat-tugas.js
   (sekitar baris 3146-3253) DENGAN versi di bawah ini.
═══════════════════════════════════════════════════════════════════════ */
async function buildTemplateData(data) {
  // ── Flags dari tipe ──────────────────────────────────────────────────
  // Sumber kebenaran sekarang adalah `data.tipe`, bukan jumlah pegawai.
  // Kalau tipe NULL/invalid, tipeFlags() return all-false (defensive).
  const flags = tipeFlags(data.tipe);

  // ── Daftar pegawai ───────────────────────────────────────────────────
  const nipList  = Array.isArray(data.pegawai_nip)  ? data.pegawai_nip  : [];
  const nameList = Array.isArray(data.pegawai_list) ? data.pegawai_list : [];

  // Pegawai pertama — dipakai untuk halaman 1 (kalau cuma 1 orang) dan SPD
  const firstNip = String(nipList[0] || '').trim();
  const firstNm  = nameList[0] || '';
  const peg              = pegawaiByNIP[firstNip];
  const namaPegawai      = (peg && peg.NAMA) || firstNm || '';
  const jabatanPegawai   = lookupJabatan(firstNip, data.tanggal_surat) || '';
  // Kolom pangkat/golongan belum ada di riwayat_pegawai — kosongkan dulu.
  const pangkatPegawai   = '';
  // Satuan kerja: kolom UNIT KERJA dari tabel "data pegawai".
  const skerjaPegawai    = (peg && (peg['UNIT KERJA'] || peg.UNIT_KERJA)) || '';

  // ── Halaman 1: nama, jabatan, pangkat → "Terlampir" kalau ≥2 pegawai
  // Sesuai konfirmasi user: untuk pegawai >1 nama, jabatan, pangkat &
  // golongan, serta jabatan/instansi memang menjadi "Terlampir".
  // (Termasuk di SPD — sengaja, supaya konsisten.)
  // ─────────────────────────────────────────────────────────────────────
  const banyakPegawai = nameList.length >= 2;
  const namaHal1      = banyakPegawai ? 'Terlampir' : namaPegawai;
  const jabatanHal1   = banyakPegawai ? 'Terlampir' : jabatanPegawai;
  const pangkatHal1   = banyakPegawai ? 'Terlampir' : pangkatPegawai;

  // Bangun array pegawai untuk loop tabel di halaman lampiran (T2/T3).
  // Setiap item: { no, nama_p, nip_p, pangkat_p, golongan_p, jabatan_p,
  //                bertugas_p, skerja_p }
  // pangkat_p / golongan_p / bertugas_p sengaja dikosongkan dulu — user
  // akan update belakangan setelah skema DB dilengkapi.
  const pegawaiLampiran = nameList.map((nm, i) => {
    const nip = String(nipList[i] || '').trim();
    const p   = pegawaiByNIP[nip];
    return {
      no:         String(i + 1),
      nama_p:     (p && p.NAMA) || nm || '',
      nip_p:      nip || '-',
      pangkat_p:  '',
      golongan_p: '',
      jabatan_p:  lookupJabatan(nip, data.tanggal_surat) || '',
      bertugas_p: '',
      skerja_p:   (p && (p['UNIT KERJA'] || p.UNIT_KERJA)) || '',
    };
  });

  // Array kendaraan & menginap — sama isinya per pegawai. Saat ini cuma
  // ada satu "iterasi" per pegawai (tabel di template di-loop sekali per
  // orang). Field belum diisi karena belum ada kolom DB-nya.
  const pegawaiBaris = nameList.map((nm, i) => {
    const nip = String(nipList[i] || '').trim();
    const p   = pegawaiByNIP[nip];
    return {
      no:         String(i + 1),
      nama_p:     (p && p.NAMA) || nm || '',
      nip_p:      nip || '-',
      pangkat_p:  '',
      golongan_p: '',
      jabatan_p:  lookupJabatan(nip, data.tanggal_surat) || '',
      bertugas_p: '',
      skerja_p:   (p && (p['UNIT KERJA'] || p.UNIT_KERJA)) || '',
    };
  });

  // ── Penandatangan ────────────────────────────────────────────────────
  const ttdNama    = data.penandatangan_nama || '';
  const ttdNip     = data.penandatangan_nip  || '';
  const ttdJabatan = data.penandatangan_jabatan
                  || lookupJabatan(ttdNip, data.tanggal_surat) || '';

  // ── PPK (Pejabat Pembuat Komitmen) ───────────────────────────────────
  const namaPPK = PPK_NAMA_DEFAULT;
  const nipPPK  = findNipByNama(namaPPK);

  // ── MAK Pembebanan ───────────────────────────────────────────────────
  const mak        = parseMAK(data.pembebanan);
  const makLengkap = formatMAKLengkap(mak);

  // ── Nomor surat & tahun dari tanggal_surat ───────────────────────────
  const nomorSuratFull = buildNomorSuratFull(data.nomor_surat, data.tanggal_surat);
  const tglSuratDate   = parseISODate(data.tanggal_surat);
  const tahun          = tglSuratDate ? String(tglSuratDate.getFullYear()) : '';

  // ── Nomor SPD ────────────────────────────────────────────────────────
  const nomorSPD = buildNomorSPD(data.nomor_surat, data.tanggal_surat, mak);

  // ── Lama hari (inclusive) ────────────────────────────────────────────
  const hari = hitungHariInclusive(data.tanggal_berangkat, data.tanggal_kembali);

  // ── Waktu pelaksanaan dalam format Indonesia (range-aware) ───────────
  const waktuPelaksanaan = fmtWaktu(data.tanggal_berangkat, data.tanggal_kembali) || '';

  // ── Load kamus POK untuk tahun anggaran surat ────────────────────────
  const tahunAnggaran = tglSuratDate ? tglSuratDate.getFullYear() : new Date().getFullYear();
  await loadKamusPok(tahunAnggaran);

  return {
    // ─────────────────────────────────────────────────────────────────
    // HALAMAN 1 — Surat Tugas
    // ─────────────────────────────────────────────────────────────────
    nomor_surat:           nomorSuratFull,
    menimbang:             data.menimbang_custom || '',
    nama:                  namaHal1,        // ← "Terlampir" jika ≥2 pegawai
    jabatan:               jabatanHal1,     // ← "Terlampir" jika ≥2 pegawai
    perihal:               data.perihal || '',
    tempat_tujuan:         data.tujuan  || '',
    waktu_pelaksanaan:     waktuPelaksanaan,
    tahun:                 tahun,
    mak_pembebanan:        makLengkap,
    tgl_surat:             fmtTglId(data.tanggal_surat),
    jabatan_penandatangan: ttdJabatan,
    penandatangan:         ttdNama,
    nip_penandatangan:     ttdNip,

    // ─────────────────────────────────────────────────────────────────
    // HALAMAN 2 — SPD (toggle dengan {#has_spd}...{/has_spd} di T1)
    //   Catatan: T3 tidak punya {#has_spd} wrapper — SPD selalu render
    //   di sana. Aman karena tipe yg pakai T3 memang selalu ada SPD.
    // ─────────────────────────────────────────────────────────────────
    has_spd:               flags.has_spd,
    nomor_spd:             nomorSPD,
    nama_ppk:              namaPPK,
    nip_ppk:               nipPPK,
    nip:                   firstNip,
    pangkat:               pangkatHal1,    // ← "Terlampir" jika ≥2 pegawai
    angkutan:              data.alat_angkutan || '',
    hari:                  hari ? String(hari) : '',
    tanggal_berangkat:     fmtTglId(data.tanggal_berangkat),
    tanggal_kembali:       fmtTglId(data.tanggal_kembali || data.tanggal_berangkat),

    // ── Mata Anggaran (per komponen) ─────────────────────────────────
    program:               mak ? mak.program       : '',
    kegiatan:              mak ? mak.kegiatan      : '',
    kro:                   mak ? `${mak.kegiatan}.${mak.kro}`               : '',
    ro:                    mak ? `${mak.kegiatan}.${mak.kro}.${mak.ro}`     : '',
    komponen:              mak ? mak.komponen      : '',
    sub_komponen:          mak ? mak.sub_komponen  : '',
    akun:                  mak ? mak.akun          : '',

    // ── Deskripsi mata anggaran ──────────────────────────────────────
    des_program:           mak ? lookupDeskripsi('program',      mak.program)                                  : '',
    des_kegiatan:          mak ? lookupDeskripsi('kegiatan',     mak.kegiatan)                                 : '',
    des_kro:               mak ? lookupDeskripsi('kro',          mak.kro,          mak.kegiatan)               : '',
    des_ro:                mak ? lookupDeskripsi('ro',           mak.ro,           `${mak.kegiatan}.${mak.kro}`) : '',
    des_komponen:          mak ? lookupDeskripsi('komponen',     mak.komponen)                                 : '',
    des_sub_komponen:      mak ? lookupDeskripsi('sub_komponen', mak.sub_komponen)                             : '',
    des_akun:              mak ? lookupDeskripsi('akun',         mak.akun)                                     : '',

    // ─────────────────────────────────────────────────────────────────
    // BLOK KENDARAAN — toggle {#kendaraan}...{/kendaraan}
    // BLOK MENGINAP  — toggle {#menginap}...{/menginap}
    //
    // Array kosong → blok hilang dari output.
    // Array berisi → blok render N kali (N = jumlah pegawai).
    // ─────────────────────────────────────────────────────────────────
    kendaraan:             flags.has_kendaraan ? pegawaiBaris : [],
    menginap:              flags.has_menginap  ? pegawaiBaris : [],

    // ─────────────────────────────────────────────────────────────────
    // SECTION II SPD — multi-row loop {#destinasi}...{/destinasi}
    //
    // 4 baris tabel (Tiba di / Berangkat dari / Pada tanggal / spasi)
    // direplikasi N kali. N dihitung dari jumlah kata "Kampung/Desa/
    // Kelurahan" di teks tujuan (lihat countDestinasi() di atas).
    //
    // Setiap iterasi adalah object KOSONG ({}) — tidak ada placeholder
    // di template Section II, semua label-only. User isi manual di Word
    // setelah generate.
    //
    // Hanya aktif di tipe yang punya SPD (T1/T3 dengan {#has_spd}).
    // Untuk tipe non-SPD, array tetap di-build tapi tidak dipakai.
    // ─────────────────────────────────────────────────────────────────
    destinasi:             Array(countDestinasi(data.tujuan)).fill({}),

    // ─────────────────────────────────────────────────────────────────
    // HALAMAN LAMPIRAN — toggle {#ul}...{/ul} (T2 & T3)
    //
    // Untuk T1 (tanpa lampiran), array kosong cukup karena T1 tidak
    // punya tag {#ul} sama sekali — engine ignore field tak terpakai.
    // - has_lampiran_st: legacy field (backward-compat dgn template lama
    //   yg pakai {#has_lampiran_st}, kalau masih ada di production).
    // - awalan: pembuka kalimat sebelum {menimbang} di header lampiran.
    //   Saat ini dikosongkan — akan diisi belakangan.
    // ─────────────────────────────────────────────────────────────────
    has_lampiran_st:       flags.has_lampiran,
    awalan:                '',
    ul:                    flags.has_lampiran ? pegawaiLampiran : [],
  };
}

/* ════════════════════════════════════════════════════════════════════
   POST-PROCESSING: hapus paragraf kosong di akhir body (Tugas #6)
   ─────────────────────────────────────────────────────────────────────
   Konteks: blok loop {#menginap}...{/menginap} di template kadang
   meninggalkan paragraf kosong di akhir document (artefak engine
   docxtemplater) yang men-trigger blank page di Word. Karena meta-module
   `dropLastPageIfEmpty()` adalah PAID module, kita lakukan post-processing
   manual: ambil `word/document.xml` dari zip, parse via DOMParser, hapus
   paragraf kosong yang ada DI AKHIR body (tepat sebelum final <w:sectPr>),
   lalu write back ke zip.

   Defensive:
     - SEMUA error dibungkus try/catch — kalau apapun gagal, document
       asli TIDAK disentuh (no-op fallback).
     - Berhenti iterasi begitu ketemu konten meaningful (text run, gambar,
       atau sectPr nested).
     - Cap maksimal 30 paragraf untuk mencegah loop accidental.

   Aman dipanggil walaupun template tidak punya trailing empty para
   (no-op).
═══════════════════════════════════════════════════════════════════════ */
function dropTrailingEmptyParagraphs(doc) {
  let zip;
  try { zip = doc.getZip(); } catch(_) { return; }

  const xmlPath = 'word/document.xml';
  const file = zip.file(xmlPath);
  if (!file) {
    console.warn('[NOVA] dropTrailingEmpty: word/document.xml not found');
    return;
  }

  let xml;
  try { xml = file.asText(); } catch(_) { return; }
  if (!xml || typeof xml !== 'string') return;

  // Preserve XML declaration kalau ada (XMLSerializer biasanya drop ini)
  let xmlDecl = '';
  const declMatch = xml.match(/^<\?xml[^?]*\?>\s*/);
  if (declMatch) xmlDecl = declMatch[0];

  // Parse
  let xmlDoc;
  try {
    xmlDoc = new DOMParser().parseFromString(xml, 'application/xml');
  } catch(e) {
    console.warn('[NOVA] dropTrailingEmpty: parse failed', e);
    return;
  }

  // Cek error parse (browser inject <parsererror>)
  if (xmlDoc.getElementsByTagName('parsererror').length > 0) {
    console.warn('[NOVA] dropTrailingEmpty: parsererror found, skip');
    return;
  }

  // Cari <w:body> (namespace-agnostic via localName)
  const body = xmlDoc.getElementsByTagNameNS('*', 'body')[0];
  if (!body) {
    console.warn('[NOVA] dropTrailingEmpty: <w:body> not found');
    return;
  }

  // Walk children dari belakang. Hapus paragraf kosong sampai ketemu
  // konten meaningful atau sectPr.
  let removed = 0;
  const MAX_REMOVE = 30;

  while (removed < MAX_REMOVE) {
    const last = body.lastElementChild;
    if (!last) break;
    const tag = last.localName;

    // Final sectPr (di luar paragraf) — STOP, jangan disentuh.
    if (tag === 'sectPr') break;

    // Hanya proses elemen <w:p>
    if (tag !== 'p') break;

    // Kalau paragraf ini carry sectPr di pPr (final section properties
    // dipindah ke last paragraph), JANGAN hapus — dokumen butuh sectPr.
    if (last.getElementsByTagNameNS('*', 'sectPr').length > 0) break;

    // Cek text run dengan content (whitespace-only juga dianggap kosong)
    const texts = last.getElementsByTagNameNS('*', 't');
    let hasText = false;
    for (let i = 0; i < texts.length; i++) {
      if ((texts[i].textContent || '').length > 0) { hasText = true; break; }
    }
    if (hasText) break;

    // Cek konten media (drawing, picture, OLE object)
    if (last.getElementsByTagNameNS('*', 'drawing').length > 0)  break;
    if (last.getElementsByTagNameNS('*', 'pict').length > 0)     break;
    if (last.getElementsByTagNameNS('*', 'object').length > 0)   break;

    // Paragraf kosong (atau cuma berisi page break / properties) — hapus
    body.removeChild(last);
    removed++;
  }

  if (removed === 0) return;

  // Serialize kembali ke string XML
  let newXml;
  try {
    newXml = new XMLSerializer().serializeToString(xmlDoc);
  } catch(e) {
    console.warn('[NOVA] dropTrailingEmpty: serialize failed', e);
    return;
  }

  // Restore XML declaration kalau XMLSerializer drop
  if (xmlDecl && !newXml.startsWith('<?xml')) {
    newXml = xmlDecl + newXml;
  }

  // Write back ke zip
  try {
    zip.file(xmlPath, newXml);
    console.log(`[NOVA] dropTrailingEmpty: removed ${removed} empty trailing paragraph(s)`);
  } catch(e) {
    console.warn('[NOVA] dropTrailingEmpty: zip update failed', e);
  }
}

async function buildSuratTugasDoc(data) {
  console.log('[NOVA] buildSuratTugasDoc() dipanggil', { suratId: data.id });

  // Pastikan docxtemplater sudah ter-load (dengan fallback dynamic loading)
  await ensureDocxtemplaterLoaded();

  // Docxtemplater bisa ter-expose sebagai window.docxtemplater (lowercase, UMD modern)
  // atau window.Docxtemplater (uppercase, versi lama). Kita handle keduanya.
  const DocxtemplaterCtor =
      (window.docxtemplater && (window.docxtemplater.default || window.docxtemplater))
   || (window.Docxtemplater && (window.Docxtemplater.default || window.Docxtemplater));
  const PizZipCtor =
      (window.PizZip && (window.PizZip.default || window.PizZip))
   || (window.pizzip  && (window.pizzip.default  || window.pizzip));

  if (!DocxtemplaterCtor || !PizZipCtor) {
    console.warn('[NOVA] Docxtemplater / PizZip belum dimuat — fallback ke legacy builder.',
      'window.docxtemplater =', typeof window.docxtemplater,
      'window.Docxtemplater =', typeof window.Docxtemplater,
      'window.PizZip =', typeof window.PizZip,
      'window.pizzip =', typeof window.pizzip);
    return buildSuratTugasDocLegacy(data);
  }
  console.log('[NOVA] Docxtemplater OK — akan pakai template-based rendering');

  // Tentukan template berdasarkan `data.tipe`. Validasi sudah dilakukan
  // di sisi UI (admin wajib pilih tipe sebelum approve), tapi defensive
  // check di sini supaya pesan error-nya jelas kalau ada surat lama yg
  // tipe-nya null (mis. row legacy yg belum di-backfill).
  if (!data.tipe) {
    throw new Error(
      'Surat tugas ini belum punya tipe. Buka tabel persetujuan, atur ' +
      'kembali tipe, lalu approve ulang.'
    );
  }

  const buf = await loadTemplateBuffer(data.tipe);
  const zip = new PizZipCtor(buf);
  const doc = new DocxtemplaterCtor(zip, {
    paragraphLoop: true,
    linebreaks: true,
    delimiters: { start: '{', end: '}' },
  });

  const templateData = await buildTemplateData(data);
  try {
    doc.render(templateData);
  } catch (err) {
    // Docxtemplater error biasanya informatif — munculkan ke console
    console.error('Template render error:', err, err.properties);
    throw new Error(`Template error: ${err.message}`);
  }

  // Tugas #6: post-processing untuk hapus blank page yg muncul setelah
  // {#menginap}...{/menginap} (atau loop lain) — wrapped try/catch supaya
  // error apapun di sini TIDAK menggagalkan generate doc.
  try {
    dropTrailingEmptyParagraphs(doc);
  } catch (e) {
    console.warn('[NOVA] dropTrailingEmptyParagraphs error (non-fatal):', e);
  }

  const out = doc.getZip().generate({
    type: 'blob',
    mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    compression: 'DEFLATE',
  });
  return out;
}

/* ════════════════════════════════════════════════════════════════════
   INIT
═══════════════════════════════════════════════════════════════════════ */
function init() {
  SESSION = checkSession();
  if (!SESSION) return;
  setTopbarUser(SESSION);
  initRoleSwitcher(SESSION, true);
  Promise.all([loadPegawai(), loadRiwayatPegawai(), loadUsers(), loadSurat()]);

  // Pre-warm: load library docxtemplater + 3 template di background
  // supaya klik Preview pertama tidak terhambat fetch template.
  // Cache key adalah URL, jadi load 3 tipe (T1/T2/T3) yang URL-nya unik.
  ensureDocxtemplaterLoaded().then(() => {
    loadTemplateBuffer('surat_tugas_spd_kendaraan_menginap').catch(() => {});           // T1
    loadTemplateBuffer('surat_tugas_lampiran').catch(() => {});                         // T2
    loadTemplateBuffer('surat_tugas_lampiran_spd_kendaraan_menginap').catch(() => {});  // T3
  });

  // Load history POK (MAK Pembebanan) untuk autocomplete dropdown.
  // Dijalankan setelah loadSurat selesai biar tidak ada race kalau RLS lambat.
  loadMAKSuggestions();

  // Bersihkan file preview orphan (>1 jam) di Supabase Storage. Fire-and-forget.
  cleanupOrphanPreviewFiles();
}
init();
