/* ═══════════════════════════════════════════════════════════════════════
   PORTAL 9201 — Admin Surat Tugas (Excel-like, multi-kolom)
   ─────────────────────────────────────────────────────────────────────
   Ketergantungan global:
     - SUPABASE_URL, SUPABASE_ANON_KEY (config.js)
     - getUserRoles, ADMIN_USERS (config.js)
     - SUPABASE_HEADERS, esc, novaCheckSession, novaRpc, BULAN, logout
       (9201-shared.js)
     - initRoleSwitcher, toggleUserDropdown, switchViewRole
       (9201-role-switcher.js)
     - window.docxPreview, saveAs, docxtemplater, PizZip (libs eksternal)
═══════════════════════════════════════════════════════════════════════ */

const H = SUPABASE_HEADERS;

/* ═════════════════════════════════════════════════════════════════════
   TIPE SURAT — fitur multi-template
   ─────────────────────────────────────────────────────────────────────
   10 pilihan tipe yg menentukan template .docx yg dipakai saat generate.
   String value sengaja sama dengan check-constraint surat_tugas_tipe_enum
   di DB (lihat surat_tugas_add_tipe.sql).

   Ada 5 file template di Supabase Storage:
     T1:  template-surat-tugas-spd-kendaraan-menginap.docx
          — punya {#has_spd}, {#kendaraan}, {#menginap}
     T2:  template-surat-tugas-lampiran.docx
          — punya {#ul} (tabel lampiran)
     T3:  template-surat-tugas-lampiran-spd-kendaraan-menginap.docx
          — punya {#ul} (lampiran), {#kendaraan}, {#menginap}; SPD selalu
            render (tidak ada {#has_spd} wrapper)
     T1V: template-surat-tugas-spd-visum-kendaraan-menginap.docx
          — superset dari T1 + {#visum}{#r}{/r}{/visum} (lembar visum)
     T3V: template-surat-tugas-lampiran-spd-visum-kendaraan-menginap.docx
          — superset dari T3 + {#visum}{#r}{/r}{/visum} (lembar visum)

   Catatan tag yg perlu Anda pastikan ADA di template:
     - T1 & T3: blok kendaraan dibungkus {#kendaraan}...{/kendaraan}
                blok menginap   dibungkus {#menginap}...{/menginap}
                (sebelumnya pakai {#ul} dua kali — perlu di-rename)
     - T1V & T3V: blok visum dibungkus {#visum}...{/visum} dan baris loop
                  responden di-bungkus {#r}...{/r} di baris terakhir tabel.
═══════════════════════════════════════════════════════════════════════ */

// 10 pilihan tipe (urutan = urutan tampil di dropdown UI).
// Variant Visum diletakkan setelah variant non-visum yang setara.
const TIPE_OPTIONS = [
  { value: 'surat_tugas',                                       label: 'Surat Tugas' },
  { value: 'surat_tugas_kendaraan',                             label: 'Surat Tugas + Kendaraan' },
  { value: 'surat_tugas_visum_kendaraan',                       label: 'Surat Tugas + Visum + Kendaraan' },
  { value: 'surat_tugas_lampiran',                              label: 'Surat Tugas + Lampiran' },
  { value: 'surat_tugas_spd_kendaraan',                         label: 'Surat Tugas + SPD + Kendaraan' },
  { value: 'surat_tugas_spd_kendaraan_menginap',                label: 'Surat Tugas + SPD + Kendaraan + Menginap' },
  { value: 'surat_tugas_spd_visum_kendaraan_menginap',          label: 'Surat Tugas + SPD + Visum + Kendaraan + Menginap' },
  { value: 'surat_tugas_lampiran_spd_kendaraan',                label: 'Surat Tugas + Lampiran + SPD + Kendaraan' },
  { value: 'surat_tugas_lampiran_spd_kendaraan_menginap',       label: 'Surat Tugas + Lampiran + SPD + Kendaraan + Menginap' },
  { value: 'surat_tugas_lampiran_spd_visum_kendaraan_menginap', label: 'Surat Tugas + Lampiran + SPD + Visum + Kendaraan + Menginap' },
];

// URL template di Supabase Storage. Pastikan nama file yang Anda upload
// di bucket `template/` sama persis dengan path di sini.
const TEMPLATE_URL_T1  = 'https://jsmmtqeoukkgugorrvmg.supabase.co/storage/v1/object/public/template/template-surat-tugas-spd-kendaraan-menginap.docx';
const TEMPLATE_URL_T2  = 'https://jsmmtqeoukkgugorrvmg.supabase.co/storage/v1/object/public/template/template-surat-tugas-lampiran.docx';
const TEMPLATE_URL_T3  = 'https://jsmmtqeoukkgugorrvmg.supabase.co/storage/v1/object/public/template/template-surat-tugas-lampiran-spd-kendaraan-menginap.docx';
const TEMPLATE_URL_T1V = 'https://jsmmtqeoukkgugorrvmg.supabase.co/storage/v1/object/public/template/template-surat-tugas-spd-visum-kendaraan-menginap.docx';
const TEMPLATE_URL_T3V = 'https://jsmmtqeoukkgugorrvmg.supabase.co/storage/v1/object/public/template/template-surat-tugas-lampiran-spd-visum-kendaraan-menginap.docx';

// Mapping tipe → template URL.
// Variant visum pakai T1V/T3V (template superset yg punya {#visum}).
const TIPE_TO_TEMPLATE = {
  'surat_tugas':                                       TEMPLATE_URL_T1,
  'surat_tugas_kendaraan':                             TEMPLATE_URL_T1,
  'surat_tugas_visum_kendaraan':                       TEMPLATE_URL_T1V,
  'surat_tugas_lampiran':                              TEMPLATE_URL_T2,
  'surat_tugas_spd_kendaraan':                         TEMPLATE_URL_T1,
  'surat_tugas_spd_kendaraan_menginap':                TEMPLATE_URL_T1,
  'surat_tugas_spd_visum_kendaraan_menginap':          TEMPLATE_URL_T1V,
  'surat_tugas_lampiran_spd_kendaraan':                TEMPLATE_URL_T3,
  'surat_tugas_lampiran_spd_kendaraan_menginap':       TEMPLATE_URL_T3,
  'surat_tugas_lampiran_spd_visum_kendaraan_menginap': TEMPLATE_URL_T3V,
};

// Mapping tipe → flags untuk kontrol section di template.
//   has_spd       → toggle {#has_spd}...{/has_spd}  (hanya berpengaruh di T1/T1V)
//   has_kendaraan → kalau true, kirim array `kendaraan` (T1/T1V/T3/T3V)
//   has_menginap  → kalau true, kirim array `menginap`  (T1/T1V/T3/T3V)
//   has_lampiran  → kalau true, kirim array `ul` utk tabel lampiran (T2/T3/T3V)
//   has_visum     → kalau true, kirim section `visum` & array `r` (T1V/T3V)
//
// Catatan: untuk T3/T3V, has_spd dikirim true walau template tidak punya
// {#has_spd} wrapper — engine docxtemplater akan ignore tag yg tidak ada
// di template, jadi aman.
const TIPE_TO_FLAGS = {
  'surat_tugas':                                       { has_spd:false, has_kendaraan:false, has_menginap:false, has_lampiran:false, has_visum:false },
  'surat_tugas_kendaraan':                             { has_spd:false, has_kendaraan:true,  has_menginap:false, has_lampiran:false, has_visum:false },
  'surat_tugas_visum_kendaraan':                       { has_spd:false, has_kendaraan:true,  has_menginap:false, has_lampiran:false, has_visum:true  },
  'surat_tugas_lampiran':                              { has_spd:false, has_kendaraan:false, has_menginap:false, has_lampiran:true,  has_visum:false },
  'surat_tugas_spd_kendaraan':                         { has_spd:true,  has_kendaraan:true,  has_menginap:false, has_lampiran:false, has_visum:false },
  'surat_tugas_spd_kendaraan_menginap':                { has_spd:true,  has_kendaraan:true,  has_menginap:true,  has_lampiran:false, has_visum:false },
  'surat_tugas_spd_visum_kendaraan_menginap':          { has_spd:true,  has_kendaraan:true,  has_menginap:true,  has_lampiran:false, has_visum:true  },
  'surat_tugas_lampiran_spd_kendaraan':                { has_spd:true,  has_kendaraan:true,  has_menginap:false, has_lampiran:true,  has_visum:false },
  'surat_tugas_lampiran_spd_kendaraan_menginap':       { has_spd:true,  has_kendaraan:true,  has_menginap:true,  has_lampiran:true,  has_visum:false },
  'surat_tugas_lampiran_spd_visum_kendaraan_menginap': { has_spd:true,  has_kendaraan:true,  has_menginap:true,  has_lampiran:true,  has_visum:true  },
};

// Konstanta layout tabel visum (sinkron dengan template T1V/T3V).
// Tabel visum punya 6 baris kosong default; baris terakhir adalah baris loop
// dengan tag {#r}...{/r}. Kalau user input jumlah responden ≤ VISUM_DEFAULT_ROWS,
// loop tetap render minimal 1x → minimal 7 baris terlihat.
// Kalau input > VISUM_DEFAULT_ROWS, loop render (input - VISUM_DEFAULT_ROWS) baris.
const VISUM_DEFAULT_ROWS = 6;

// Hitung jumlah baris yang harus di-loop (`r` array length) berdasarkan
// jumlah responden yang user input. Defensive: input apapun (null/string/
// negatif/desimal) di-coerce ke integer. Lihat juga buildTemplateData().
function calcVisumLoopCount(jumlahResponden) {
  const n = parseInt(jumlahResponden, 10);
  if (!isFinite(n) || n <= VISUM_DEFAULT_ROWS) return 1; // minimal 1 → total 7 baris
  return n - VISUM_DEFAULT_ROWS;
}

// Helper: apakah tipe surat ini butuh input "Jumlah Responden" untuk visum?
function tipeHasVisum(tipe) {
  return tipeFlags(tipe).has_visum === true;
}

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
  return TIPE_TO_FLAGS[tipe] || { has_spd:false, has_kendaraan:false, has_menginap:false, has_lampiran:false, has_visum:false };
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
let pegawaiList = [];             // dari "data_pegawai"
let pegawaiByNIP = {};            // index NIP → object
let riwayatJabatan = [];          // dari "riwayat_jabatan"
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
// `checkSession`, `logout`, `esc`, dan `BULAN` di-load dari 9201-shared.js.
// Halaman ini admin-only — pakai novaCheckSession({ requireAdmin: true })
// di init().
// Topbar (clock + user info) di-handle oleh 9201-topbar.js.
// Panggil Topbar9201.setUser(SESSION) saat init untuk set avatar+username.

/* ════════════════════════════════════════════════════════════════════
   HELPERS — escape, format tanggal, badge
═══════════════════════════════════════════════════════════════════════ */
// `esc` dari 9201-shared.js. `escAttr` adalah alias `esc` (juga shared).

function fmtTgl(str) {
  if (!str) return '';
  const d = parseISODate(str);
  if (!d) return '';
  return `${d.getDate()} ${BULAN[d.getMonth()]} ${d.getFullYear()}`;
}

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
    const res = await fetch(`${SUPABASE_URL}/rest/v1/data_pegawai?select=NIP,NAMA&order=NAMA.asc`, { headers: H });
    if (!res.ok) return;
    pegawaiList = await res.json();
    pegawaiByNIP = {};
    pegawaiList.forEach(p => {
      const nip = String(p.NIP || '').trim();
      if (nip) pegawaiByNIP[nip] = p;
    });
  } catch(e) { console.warn('Gagal load pegawai:', e); }
}

async function loadRiwayatJabatan() {
  try {
    const res = await fetch(`${SUPABASE_URL}/rest/v1/riwayat_jabatan?select=*&order=tmt.desc`, { headers: H });
    if (!res.ok) return;
    riwayatJabatan = await res.json();
  } catch(e) { console.warn('Gagal load riwayat_jabatan:', e); }
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
      ${cellTextareaHTML(s.id, 'tujuan', tujuan, editable, 'Kota/instansi')}
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
           tabindex="0" data-col-field="tipe" data-value="${escAttr(val || '')}">${val ? esc(label) : '—'}</div>
    </td>`;
  }
  const tag = val ? buildTipeTag(val) : '';
  const placeholder = val ? '' : 'Pilih tipe...';
  return `<td class="col-tipe">
    <div class="tp-cell" data-field="tipe" data-col-field="tipe" data-id="${id}"
      data-value="${escAttr(val || '')}"
      onclick="onTpCellClick(event, this)">
      ${tag}
      <input type="text" class="tp-input" placeholder="${placeholder}"
        oninput="onTpInput(event, this)"
        onkeydown="onTpKeydown(event, this)"
        onfocus="onTpInputFocus(this)"
        onblur="onTpInputBlur(this)">
    </div>
  </td>`;
}

function buildTipeTag(val) {
  const label = tipeLabel(val);
  return `<span class="tp-tag" data-value="${escAttr(val)}" title="${escAttr(label)}">
    <span class="tp-tag-text">${esc(label)}</span>
    <button type="button" class="tp-tag-x" onclick="onTpTagRemove(event, this)">×</button>
  </span>`;
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
      console.warn(`[9201] loadMAKSuggestions HTTP ${res.status}`);
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
    console.log(`[9201] loadMAKSuggestions: ${makSuggestions.length} unique MAK`);
  } catch (e) {
    console.warn('[9201] loadMAKSuggestions error:', e);
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
    <textarea class="xls-cell mak-input" rows="1" data-field="pembebanan" data-id="${id}"
      data-col-field="pembebanan"
      placeholder="cth: 054.01.GG.2910.BMA.006.054.A.524119"
      oninput="onMAKInput(this)"
      onfocus="onMAKFocus(this)"
      onblur="onMAKBlur(this)"
      onkeydown="onMAKKeydown(event, this)"
      autocomplete="off">${esc(val || '')}</textarea>
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
    } else if (e.key === 'Enter') {
      // Cegah newline — POK secara semantik adalah satu kode, bukan
      // multi-line. Wrap visual hanya supaya teks panjang terlihat,
      // bukan untuk menampung baris baru.
      e.preventDefault();
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
    // Selalu preventDefault supaya textarea tidak menyisipkan newline,
    // baik ada item terpilih maupun tidak.
    e.preventDefault();
    if (makACState.focusIdx >= 0 && makACState.filtered[makACState.focusIdx]) {
      pickMAK(makACState.filtered[makACState.focusIdx].mak);
    }
    // Kalau tidak ada item terpilih → no-op (sesuai single-line behavior).
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
   AUTOCOMPLETE TIPE — single-select dari TIPE_OPTIONS
   Pola: mirror dari pegawai (pg-cell), tapi infrastructure terpisah
   karena source data statis & value berbeda type.

   Flow:
   1. User klik cell (.tp-cell) → fokus ke .tp-input → openTpAc()
   2. User mengetik → tpAcFilter(query) menyaring TIPE_OPTIONS
   3. ↑↓ Enter atau klik item → pickTipe(value)
   4. Tag terpasang, input kosong, dropdown tutup
   5. Klik × pada tag → clearTipe()
═══════════════════════════════════════════════════════════════════════ */
let tpState = {
  cellEl: null,
  inputEl: null,
  filtered: [],
  focusIdx: -1,
};

function onTpCellClick(e, cellEl) {
  // Klik di dalam tag-x: biarkan handler tag yg jalan, jangan re-fokus.
  if (e.target.closest('.tp-tag-x')) return;
  const inp = cellEl.querySelector('.tp-input');
  if (inp) inp.focus();
}

function onTpInputFocus(inp) {
  const cellEl = inp.closest('.tp-cell');
  if (cellEl) cellEl.classList.add('focused');
  openTpAc(cellEl, inp);
  tpAcFilter(inp.value);
}

function onTpInputBlur(inp) {
  const cellEl = inp.closest('.tp-cell');
  if (cellEl) cellEl.classList.remove('focused');
  // Delay close supaya klik item dropdown tidak kabur duluan.
  setTimeout(() => {
    if (!document.activeElement || !document.activeElement.closest('#tp-ac-popup')) {
      closeTpAc();
    }
  }, 120);
}

function onTpInput(e, inp) {
  const cellEl = inp.closest('.tp-cell');
  if (cellEl) cellEl.classList.remove('err');
  tpState.cellEl  = cellEl;
  tpState.inputEl = inp;
  if (!document.getElementById('tp-ac-popup').classList.contains('open')) {
    openTpAc(cellEl, inp);
  }
  tpAcFilter(inp.value);
}

function onTpKeydown(e, inp) {
  const popup = document.getElementById('tp-ac-popup');
  const isOpen = popup.classList.contains('open');
  const cellEl = inp.closest('.tp-cell');

  // Backspace di input kosong + ada tag → hapus tag (mimik pg-cell behavior)
  if (e.key === 'Backspace' && inp.value === '') {
    const tag = cellEl && cellEl.querySelector('.tp-tag');
    if (tag) {
      e.preventDefault();
      clearTipe(cellEl);
    }
    return;
  }

  if (e.key === 'Escape') {
    if (isOpen) { e.preventDefault(); closeTpAc(); }
    return;
  }

  if (!isOpen) {
    if (e.key === 'ArrowDown' || e.key === 'Enter') {
      e.preventDefault();
      openTpAc(cellEl, inp);
      tpAcFilter(inp.value);
    }
    return;
  }

  if (e.key === 'ArrowDown') {
    e.preventDefault();
    if (tpState.filtered.length) {
      tpState.focusIdx = Math.min(tpState.focusIdx + 1, tpState.filtered.length - 1);
      tpAcRenderFocus();
    }
  } else if (e.key === 'ArrowUp') {
    e.preventDefault();
    tpState.focusIdx = Math.max(tpState.focusIdx - 1, 0);
    tpAcRenderFocus();
  } else if (e.key === 'Enter') {
    e.preventDefault();
    if (tpState.focusIdx >= 0 && tpState.filtered[tpState.focusIdx]) {
      pickTipe(tpState.filtered[tpState.focusIdx].value);
    }
  } else if (e.key === 'Tab') {
    closeTpAc();
  }
}

function onTpTagRemove(e, btn) {
  e.stopPropagation();
  const cellEl = btn.closest('.tp-cell');
  if (cellEl) clearTipe(cellEl);
}

function openTpAc(cellEl, inp) {
  if (!cellEl || !inp) return;
  tpState.cellEl  = cellEl;
  tpState.inputEl = inp;
  tpState.focusIdx = -1;
  const popup = document.getElementById('tp-ac-popup');
  popup.classList.add('open');
  positionTpAcPopup(inp);
}

function closeTpAc() {
  const popup = document.getElementById('tp-ac-popup');
  if (popup) popup.classList.remove('open');
  tpState.cellEl = null;
  tpState.inputEl = null;
  tpState.filtered = [];
  tpState.focusIdx = -1;
}

function positionTpAcPopup(inp) {
  const popup = document.getElementById('tp-ac-popup');
  const cell  = inp.closest('.tp-cell') || inp;
  const rect  = cell.getBoundingClientRect();
  // Posisikan popup tepat di bawah cell, alignment kiri
  const popH  = 280; // estimasi tinggi maksimal popup
  const spaceBelow = window.innerHeight - rect.bottom;
  const above = spaceBelow < popH && rect.top > popH;
  popup.style.left = `${Math.max(8, rect.left)}px`;
  popup.style.minWidth = `${Math.max(rect.width, 260)}px`;
  if (above) {
    popup.style.top = '';
    popup.style.bottom = `${window.innerHeight - rect.top + 4}px`;
  } else {
    popup.style.bottom = '';
    popup.style.top = `${rect.bottom + 4}px`;
  }
}

function tpAcFilter(query) {
  const q = (query || '').toLowerCase().trim();
  // Filter: cocok di label ATAU di value (untuk cari "lampiran", "spd", "menginap", dll.)
  tpState.filtered = TIPE_OPTIONS.filter(o =>
    !q ||
    o.label.toLowerCase().includes(q) ||
    o.value.toLowerCase().includes(q)
  );
  // Default: focus item pertama (kalau ada). Kalau current tipe sudah dipilih,
  // sorot item itu sebagai posisi awal.
  const currentVal = tpState.cellEl ? (tpState.cellEl.dataset.value || '') : '';
  let idx = -1;
  if (currentVal) {
    idx = tpState.filtered.findIndex(o => o.value === currentVal);
  }
  if (idx < 0 && tpState.filtered.length) idx = 0;
  tpState.focusIdx = idx;
  tpAcRender();
}

function tpAcRender() {
  const list = document.getElementById('tp-ac-list');
  if (!list) return;
  if (!tpState.filtered.length) {
    list.innerHTML = `<div class="tp-ac-empty">Tidak ada opsi cocok</div>`;
    return;
  }
  const currentVal = tpState.cellEl ? (tpState.cellEl.dataset.value || '') : '';
  list.innerHTML = tpState.filtered.map((o, i) => {
    const cls = [
      'tp-ac-item',
      o.value === currentVal ? 'selected' : '',
      i === tpState.focusIdx ? 'focused' : '',
    ].filter(Boolean).join(' ');
    return `<div class="${cls}" data-idx="${i}" onmousedown="event.preventDefault()" onclick="pickTipe('${escAttr(o.value)}')">${esc(o.label)}</div>`;
  }).join('');
}

function tpAcRenderFocus() {
  const items = document.querySelectorAll('#tp-ac-list .tp-ac-item');
  items.forEach((el, i) => el.classList.toggle('focused', i === tpState.focusIdx));
  // Auto-scroll item ter-focus ke dalam viewport popup
  if (tpState.focusIdx >= 0 && items[tpState.focusIdx]) {
    items[tpState.focusIdx].scrollIntoView({ block: 'nearest' });
  }
}

function pickTipe(value) {
  const cellEl = tpState.cellEl;
  if (!cellEl) { closeTpAc(); return; }
  cellEl.dataset.value = value || '';
  cellEl.classList.remove('err');
  // Re-render: hapus tag lama, tambahkan tag baru, kosongkan input, hilangkan placeholder
  const existing = cellEl.querySelector('.tp-tag');
  if (existing) existing.remove();
  const inp = cellEl.querySelector('.tp-input');
  if (value) {
    cellEl.insertAdjacentHTML('afterbegin', buildTipeTag(value));
  }
  if (inp) {
    inp.value = '';
    inp.placeholder = value ? '' : 'Pilih tipe...';
  }
  closeTpAc();
}

function clearTipe(cellEl) {
  if (!cellEl) return;
  cellEl.dataset.value = '';
  const tag = cellEl.querySelector('.tp-tag');
  if (tag) tag.remove();
  const inp = cellEl.querySelector('.tp-input');
  if (inp) {
    inp.placeholder = 'Pilih tipe...';
    inp.focus();
  }
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
  'tipe',
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

      // 3) Editable tp-cell (tipe) → kembalikan tp-input di dalamnya
      const tpCell = row.querySelector(`.tp-cell[data-col-field="${field}"]`);
      if (tpCell) {
        return tpCell.querySelector('.tp-input') || tpCell;
      }

      // 4) Readonly div (ro-text atau ro-ttd)
      el = row.querySelector(`[data-col-field="${field}"]`);
      return el || null;
    });
  });
}

/**
 * Cari posisi elemen dalam grid.
 * Jika elemen adalah pg-input/tp-input, cari parent pg-cell/tp-cell-nya di grid.
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
      // tp-input → cek apakah parent tp-cell ada di grid
      if (target.classList && target.classList.contains('tp-input')) {
        const parentTp = target.closest('.tp-cell');
        if (parentTp && el === parentTp) return { r, c };
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
  const isTpInput  = target.classList && target.classList.contains('tp-input');
  const isRoText   = target.classList && (
    target.classList.contains('ro-text') || target.classList.contains('ro-ttd')
  );
  const isNavTarget = isXlsCell || isPgInput || isTpInput || isRoText;

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

    // ArrowUp/Down di Tipe autocomplete = biarkan handler onTpKeydown handle navigate item
    if (isTpInput && document.getElementById('tp-ac-popup').classList.contains('open')) {
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
      if (cursorAtStart) {
        if (c > 0) {
          // Pindah kiri dalam baris yang sama
          e.preventDefault();
          closeAllPopups();
          navFocusCell(grid, r, c - 1);
        } else if (r > 0) {
          // Sudah di kolom pertama → wrap ke kolom terakhir baris sebelumnya
          e.preventDefault();
          closeAllPopups();
          navFocusCell(grid, r - 1, NAV_FIELDS.length - 1);
        }
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
      if (cursorAtEnd) {
        if (c < NAV_FIELDS.length - 1) {
          // Pindah kanan dalam baris yang sama
          e.preventDefault();
          closeAllPopups();
          navFocusCell(grid, r, c + 1);
        } else if (r < grid.length - 1) {
          // Sudah di kolom terakhir → wrap ke kolom pertama baris berikutnya
          e.preventDefault();
          closeAllPopups();
          navFocusCell(grid, r + 1, 0);
        }
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
    'tr[data-status="menunggu"] .xls-cell, tr[data-status="menunggu"] .pg-input, tr[data-status="menunggu"] .tp-input'
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
  const tpAc = document.getElementById('tp-ac-popup');
  if (tpAc && tpAc.classList.contains('open') && !tpAc.contains(e.target)) {
    const inAnyTpCell = e.target.closest && e.target.closest('.tp-cell');
    if (!inAnyTpCell) closeTpAc();
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
  closeTpAc();
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
  // Tipe disimpan di dataset.value pada .tp-cell (bukan value/value attribute)
  // karena cell custom autocomplete, bukan <select> native.
  const getTipe = () => {
    const el = row.querySelector('.tp-cell[data-field="tipe"]');
    if (el) return el.dataset.value || '';
    // Fallback: kalau row sudah selesai (read-only ro-text), baca data-col-field
    const ro = row.querySelector('[data-col-field="tipe"]');
    return ro ? (ro.dataset.value || '') : '';
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
    tipe:                 getTipe() || null,
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
  row.querySelectorAll('.xls-cell.err, .pg-cell.err, .tp-cell.err').forEach(el => el.classList.remove('err'));

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
    const first = row.querySelector('.xls-cell.err, .pg-cell.err, .tp-cell.err');
    if (first) {
      first.scrollIntoView({ behavior: 'smooth', block: 'center', inline: 'center' });
      setTimeout(() => {
        if (first.tagName === 'INPUT' || first.tagName === 'TEXTAREA') first.focus();
        else {
          const inp = first.querySelector('input, textarea, .pg-input, .tp-input');
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
  // Reset state khusus modal preview
  if (id === 'modal-preview') {
    _previewVisumOpts = null;
  }
  // Kalau modal-visum-prompt ditutup tanpa klik "Lanjutkan" (X / Batal /
  // klik backdrop), resolve Promise yang lagi menunggu dengan null
  // supaya caller (withVisumOpts) tahu user cancel — tidak menggantung.
  if (id === 'modal-visum-prompt' && _visumPromptResolver) {
    const fn = _visumPromptResolver;
    _visumPromptResolver = null;
    fn(null);
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

  // Info badge tipe — kalau tipe ada visum, tampilkan reminder agar admin
  // tahu nanti perlu input jumlah responden saat preview/download.
  const tipeBadge = tipeHasVisum(values.tipe)
    ? `${esc(tipeLabel(values.tipe))} <span style="display:inline-block;margin-left:6px;background:rgba(200,168,75,.18);color:#7a5c10;border:1px solid rgba(200,168,75,.4);border-radius:100px;padding:1px 7px;font-size:10px;font-weight:600;letter-spacing:.3px">📋 VISUM</span>`
    : esc(tipeLabel(values.tipe));

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
    <div class="approve-preview-row"><strong>Tipe Surat</strong><span>${tipeBadge}</span></div>
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

    // Capture tp-cell (tipe single-select)
    const tp = tr.querySelector('.tp-cell');
    if (tp) entry.tipe = tp.dataset.value || '';

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
    });

    // Re-apply tp-cell (tipe): re-render tag + sync data-value
    if (entry.tipe !== undefined) {
      const tp = tr.querySelector('.tp-cell');
      if (tp) {
        tp.dataset.value = entry.tipe || '';
        tp.querySelectorAll('.tp-tag').forEach(t => t.remove());
        const inp = tp.querySelector('.tp-input');
        if (entry.tipe) {
          tp.insertAdjacentHTML('afterbegin', buildTipeTag(entry.tipe));
        }
        if (inp) {
          inp.value = '';
          inp.placeholder = entry.tipe ? '' : 'Pilih tipe...';
        }
      }
    }

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
     - Jabatan penandatangan di-recompute dari riwayat_jabatan berdasar
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
   LOOKUP JABATAN dari riwayat_jabatan
═══════════════════════════════════════════════════════════════════════ */
function lookupJabatan(nip, tglSuratIso) {
  if (!nip || !tglSuratIso) return '';
  const candidates = riwayatJabatan
    .filter(r => String(r.pegawai_nip || '').trim() === String(nip).trim())
    .filter(r => r.tmt && r.tmt <= tglSuratIso)
    .sort((a, b) => (b.tmt || '').localeCompare(a.tmt || ''));
  if (candidates.length) return candidates[0].jabatan || '';
  const peg = pegawaiByNIP[nip];
  if (peg && peg.NAMA) {
    const candByName = riwayatJabatan
      .filter(r => (r.nama || '').trim().toLowerCase() === (peg.NAMA || '').trim().toLowerCase())
      .filter(r => r.tmt && r.tmt <= tglSuratIso)
      .sort((a, b) => (b.tmt || '').localeCompare(a.tmt || ''));
    if (candByName.length) return candByName[0].jabatan || '';
  }
  return '';
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
// Opts visum yg sedang aktif untuk modal preview saat ini. Dipakai oleh
// downloadFromPreview() dan openInWordForPrint() agar tombol-tombol di
// modal preview pakai jumlah responden yang sama dengan preview yg sedang
// ditampilkan, tanpa tanya ulang ke user. Reset saat modal preview ditutup.
let _previewVisumOpts = null;

/* ── Visum prompt state ──────────────────────────────────────────────
   Karena jumlah_responden tidak disimpan ke DB, kita pakai modal
   prompt (#modal-visum-prompt) yang muncul on-demand sebelum
   preview/download untuk tipe yang punya visum.

   Cache opsional `_visumLastInput[suratId]` menyimpan input user
   sebelumnya per-session — jadi kalau user buka Preview lalu
   Download untuk surat yg sama, value sebelumnya jadi default
   (tetap bisa diubah). Hilang saat reload halaman.
─────────────────────────────────────────────────────────────────── */
const _visumLastInput = {};
let _visumPromptResolver = null;  // resolver untuk Promise yg lagi pending

/**
 * Tampilkan modal input "Jumlah Responden Visum" dan return Promise
 * yang resolve dengan value input (string atau null kalau di-cancel).
 *
 * @param {object} surat - object surat (untuk default value & info)
 * @returns {Promise<string|null>}  string angka atau '' (kosong) kalau OK
 *                                  null kalau user cancel
 */
function promptVisumResponden(surat) {
  return new Promise(resolve => {
    _visumPromptResolver = resolve;
    const inp = document.getElementById('inp-jumlah-responden-prompt');
    if (inp) {
      // Set default dari last input (kalau ada) untuk surat yang sama
      const cached = _visumLastInput[surat.id];
      inp.value = cached != null ? cached : '';
      // Auto-focus setelah modal terlihat
      setTimeout(() => { try { inp.focus(); inp.select(); } catch(_) {} }, 50);
    }
    openModal('modal-visum-prompt');
  });
}

/** Handler tombol "Lanjutkan" di modal visum prompt. */
function confirmVisumPrompt() {
  const inp = document.getElementById('inp-jumlah-responden-prompt');
  const val = inp ? inp.value.trim() : '';
  closeModal('modal-visum-prompt');
  if (_visumPromptResolver) {
    const fn = _visumPromptResolver;
    _visumPromptResolver = null;
    fn(val);  // string ('' kalau kosong, '10' kalau diisi)
  }
}

/**
 * Wrapper: jalankan callback dengan opts yang sudah di-resolve dari
 * visum prompt (kalau perlu). Kalau tipe surat tidak butuh visum,
 * langsung jalankan callback dengan opts kosong.
 *
 * @param {object}   surat    - object surat
 * @param {Function} callback - async (opts) => any
 * @returns {Promise<any>}    - return value callback, atau null kalau dibatalkan
 */
async function withVisumOpts(surat, callback) {
  if (!tipeHasVisum(surat && surat.tipe)) {
    return callback({});  // tipe non-visum → opts kosong
  }
  const val = await promptVisumResponden(surat);
  if (val === null) return null;  // user cancel
  _visumLastInput[surat.id] = val; // cache untuk next call
  return callback({ jumlahResponden: val === '' ? null : val });
}

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
  } catch (e) { console.warn('[9201] Cleanup preview file gagal:', e); }
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
  // docx-preview dipakai oleh openPreview() untuk render docx blob ke HTML.
  // Library di-load di admin-surat-tugas.html.
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

  // Untuk tipe visum: minta jumlah responden DULU sebelum buka modal
  // preview. Kalau user cancel di prompt, jangan buka modal preview.
  // Untuk tipe non-visum: opts={} langsung lanjut.
  let visumOpts;
  if (tipeHasVisum(surat.tipe)) {
    const val = await promptVisumResponden(surat);
    if (val === null) return;  // user cancel — abort tanpa buka preview
    _visumLastInput[surat.id] = val;
    visumOpts = { jumlahResponden: val === '' ? null : val };
  } else {
    visumOpts = {};
  }

  currentPreviewSurat = surat;
  // Simpan opts visum agar tombol "Download .docx" dan "Buka di Word"
  // di modal preview ini bisa pakai jumlah responden yang sama tanpa
  // tanya ulang. _previewVisumOpts hanya hidup selama modal preview
  // terbuka — di-reset saat modal ditutup.
  _previewVisumOpts = visumOpts;

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
    const blob = await buildSuratTugasDoc(surat, visumOpts);
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
  // Untuk tipe visum: minta jumlah responden dulu. User cancel → abort.
  const result = await withVisumOpts(surat, async (opts) => {
    try {
      ensureLibrariesLoaded();
      const blob = await buildSuratTugasDoc(surat, opts);
      saveAs(blob, buildFileName(surat));
      showPageAlert(`📥 Berhasil di-download: ${buildFileName(surat)}`, 'success');
      return true;
    } catch(e) {
      console.error(e);
      showPageAlert(`Gagal download: ${e.message}`, 'error');
      return false;
    }
  });
  // result === null artinya user cancel di prompt visum — tidak perlu alert.
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

  // Pre-collect: untuk setiap surat dengan tipe visum, prompt jumlah
  // responden DULU (sebelum mulai loop download). Ini supaya admin tidak
  // perlu nunggu di tengah-tengah loop. Kalau admin cancel salah satu
  // prompt, surat tsb di-skip (failures); sisanya tetap lanjut.
  const visumOptsById = {};
  for (let i = 0; i < ids.length; i++) {
    const id = ids[i];
    const surat = suratMap[id];
    if (!surat || surat.status !== 'selesai') continue;
    if (!tipeHasVisum(surat.tipe)) {
      visumOptsById[id] = {};
      continue;
    }
    // Tipe visum — prompt sequentially
    if (btn) btn.textContent = `📋 Visum ${i + 1}/${ids.length}…`;
    const val = await promptVisumResponden(surat);
    if (val === null) {
      // Cancel → tandai null supaya loop di-skip dengan pesan informatif
      visumOptsById[id] = null;
    } else {
      _visumLastInput[surat.id] = val;
      visumOptsById[id] = { jumlahResponden: val === '' ? null : val };
    }
  }

  for (let i = 0; i < ids.length; i++) {
    const id    = ids[i];
    const surat = suratMap[id];

    // Update progress di tombol
    if (btn) btn.textContent = `📥 ${i + 1}/${ids.length}…`;

    if (!surat || surat.status !== 'selesai') {
      failures.push(`#${id}: status bukan 'selesai'`);
      continue;
    }

    // Surat visum yang prompt-nya di-cancel → skip
    if (visumOptsById[id] === null) {
      failures.push(`${surat.perihal || 'surat #' + id}: dibatalkan saat input visum`);
      continue;
    }

    try {
      const blob = await buildSuratTugasDoc(surat, visumOptsById[id] || {});
      saveAs(blob, buildFileName(surat));
      success++;
      // Jeda antar download supaya browser tidak block multi-trigger.
      // 400ms umumnya cukup untuk Chrome/Edge/Firefox.
      if (i < ids.length - 1) {
        await new Promise(r => setTimeout(r, 400));
      }
    } catch(e) {
      console.error(`[9201] bulk download gagal id=${id}:`, e);
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
  // Pakai opts visum yg SAMA dengan yg dipakai untuk preview saat ini —
  // jangan tanya ulang. Kalau modal preview di-buka ulang, opts akan
  // di-set ulang lewat openPreview().
  const opts = _previewVisumOpts || {};
  try {
    ensureLibrariesLoaded();
    const blob = await buildSuratTugasDoc(currentPreviewSurat, opts);
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
    // Kalau tidak ada (mis. user direct call atau preview gagal), upload ulang
    // pakai opts visum yg SAMA dengan yg dipakai di preview saat ini.
    if (_previewUploadedPath) {
      signedUrl = await getPreviewSignedUrl(_previewUploadedPath, 3600);
    } else {
      ensureLibrariesLoaded();
      const blob = await buildSuratTugasDoc(currentPreviewSurat, _previewVisumOpts || {});
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
    console.error('[9201] openInWordForPrint error:', e);
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
    console.warn(`[9201] Gagal load kamus_pok (HTTP ${res.status}). Deskripsi POK akan kosong.`);
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
  console.log(`[9201] Kamus POK loaded: ${rows.length} rows untuk tahun ${tahun}`);
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
const RE_DE = /\b(kampung|desa|kelurahan)\b/gi;
function countDe(tujuanText) {
  if (!tujuanText) return 1;
  const matches = String(tujuanText).match(RE_DE);
  const count = matches ? matches.length : 0;
  return Math.max(1, count);
}

/**
 * Cari NIP berdasarkan nama di tabel riwayat_jabatan.
 * Toleran terhadap suffix gelar setelah koma — mis. "Abdillah Humam, SST"
 * akan match record dengan nama "Abdillah Humam, SST" maupun "Abdillah Humam".
 * Kalau lebih dari satu record (riwayat jabatan banyak), ambil yang pertama
 * — semua record untuk orang yang sama akan punya pegawai_nip yang sama.
 */
function findNipByNama(namaCari) {
  if (!namaCari || !Array.isArray(riwayatJabatan)) return '';
  const target     = String(namaCari).toLowerCase().trim();
  const targetCore = target.split(',')[0].trim(); // tanpa gelar

  const found = riwayatJabatan.find(r => {
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
 * saat ada pergantian PPK. NIP di-lookup runtime dari riwayat_jabatan.
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
    s.onload = () => { console.log('[9201] Loaded:', src); resolve(); };
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
      console.warn('[9201] CDN gagal, coba berikutnya:', e.message);
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

  console.log('[9201] Memuat template (tipe=' + tipe + ') dari:', url);
  const res = await fetch(url);
  if (!res.ok) {
    throw new Error(
      `Gagal memuat template untuk tipe "${tipe}" (HTTP ${res.status}). ` +
      `Pastikan file ${url} ada di Supabase Storage dan bisa diakses publik.`
    );
  }
  const buf = await res.arrayBuffer();
  _templateBufferCache[url] = buf;
  console.log('[9201] Template loaded:', buf.byteLength, 'bytes');
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
     - {pangkat}     → kolom pangkat/golongan belum ada di riwayat_jabatan → ''
     - {pangkat_p}, {golongan_p}, {bertugas_p} → di-kosongkan dulu di lampiran,
                     akan diisi setelah skema DB dilengkapi.
                     {jabatan_p} sudah di-lookup dari riwayat_jabatan.
     - {des_*}       → tabel kamus_pok belum dibuat → semua ''
   Field-field di atas akan otomatis terisi setelah skema DB dilengkapi —
   tinggal lookup di buildTemplateData() ini.
*/
/* ════════════════════════════════════════════════════════════════════
   PATCH 1 — buildTemplateData()
   GANTI fungsi buildTemplateData yang ada di admin-surat-tugas.js
   (sekitar baris 3146-3253) DENGAN versi di bawah ini.
═══════════════════════════════════════════════════════════════════════ */
async function buildTemplateData(data, opts) {
  // opts: { jumlahResponden?: number|string|null }
  // Hanya dipakai untuk tipe yang punya visum. Diabaikan untuk tipe lain.
  const jumlahResponden = (opts && opts.jumlahResponden != null) ? opts.jumlahResponden : null;
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
  // Kolom pangkat/golongan belum ada di riwayat_jabatan — kosongkan dulu.
  const pangkatPegawai   = '';
  // Satuan kerja: kolom UNIT KERJA dari tabel "data_pegawai".
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
    de:             Array(countDe(data.tujuan)).fill({}),

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

    // ─────────────────────────────────────────────────────────────────
    // LEMBAR VISUM — toggle {#visum}...{/visum} (T1V & T3V)
    //
    // Struktur tabel di template:
    //   - 1 baris header (No, Tanggal, Nama, Instansi, TTD, Ket)
    //   - 1 baris angka (1)(2)(3)(4)(5)(6)
    //   - 6 baris kosong default (untuk diisi manual setelah cetak)
    //   - 1 baris loop {#r}...{/r} di akhir
    //
    // Untuk loop {#r}{/r}: kalau jumlah_responden user input ≤ 6 atau
    // kosong → array `r` punya 1 elemen (baris loop tetap render 1x →
    // total 7 baris kosong terlihat). Kalau input > 6 → array `r` punya
    // (input - 6) elemen → total = input baris.
    //
    // Setiap elemen `r` adalah object kosong {} karena tabel ini
    // dirancang untuk dicetak & diisi manual — tidak ada placeholder
    // data di kolom-kolomnya.
    //
    // visum: array dengan 1 elemen kalau has_visum, supaya seluruh
    // section visum render 1x. Kalau false → array kosong → section
    // hilang dari output.
    //
    // Kenapa visum dibungkus array (bukan boolean)?
    //   docxtemplater: {#visum}...{/visum} dengan array berisi N elemen
    //   akan render N kali. Dengan boolean true → render 1x (sama).
    //   Pakai array supaya scope `r` di dalamnya bisa di-reference.
    // ─────────────────────────────────────────────────────────────────
    visum:                 flags.has_visum
                             ? [{ r: Array(calcVisumLoopCount(jumlahResponden)).fill({}) }]
                             : [],
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

   Struktur OOXML body biasanya berakhir salah satu dari pola ini:
     Pola A:  ...<w:p>konten</w:p><w:p>kosong</w:p><w:sectPr>...</w:sectPr></w:body>
                                                    ^^^^^^^^^^
                                                    sectPr SEBAGAI CHILD body
     Pola B:  ...<w:p>konten</w:p><w:p><w:pPr><w:sectPr>...</w:sectPr></w:pPr>...</w:p></w:body>
                                                ^^^^^^^^^^
                                                sectPr DI DALAM pPr paragraf terakhir

   Implementasi lama: cek `body.lastElementChild` — kalau ketemu <w:sectPr>
   (Pola A) langsung break tanpa cleaning. Itu BUG: paragraf-paragraf
   kosong sebelum sectPr tetap ada → blank page.

   Fix: skip <w:sectPr> di posisi paling akhir, lalu walk ke belakang dari
   sibling sebelumnya. Untuk Pola B, paragraf carrier sectPr-nya tetap
   tidak boleh dihapus (sudah di-handle dgn cek nested sectPr).

   Tambahan: paragraf yang isinya hanya page break (<w:br w:type="page"/>)
   tanpa text/media juga di-hapus — ini sumber blank page lain yang
   sebelumnya tidak ter-handle.

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
    console.warn('[9201] dropTrailingEmpty: word/document.xml not found');
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
    console.warn('[9201] dropTrailingEmpty: parse failed', e);
    return;
  }

  // Cek error parse (browser inject <parsererror>)
  if (xmlDoc.getElementsByTagName('parsererror').length > 0) {
    console.warn('[9201] dropTrailingEmpty: parsererror found, skip');
    return;
  }

  // Cari <w:body> (namespace-agnostic via localName)
  const body = xmlDoc.getElementsByTagNameNS('*', 'body')[0];
  if (!body) {
    console.warn('[9201] dropTrailingEmpty: <w:body> not found');
    return;
  }

  // Helper: cek apakah suatu <w:p> dianggap "kosong" — boleh dihapus?
  // Kosong = tidak punya text content, drawing, pict, atau object.
  // Paragraf yang isinya HANYA page break / line break / properties juga
  // dianggap kosong.
  function isEmptyParagraph(p) {
    // Ada text dengan konten?
    const texts = p.getElementsByTagNameNS('*', 't');
    for (let i = 0; i < texts.length; i++) {
      if ((texts[i].textContent || '').length > 0) return false;
    }
    // Ada media?
    if (p.getElementsByTagNameNS('*', 'drawing').length > 0)  return false;
    if (p.getElementsByTagNameNS('*', 'pict').length > 0)     return false;
    if (p.getElementsByTagNameNS('*', 'object').length > 0)   return false;
    // Konten lain (mis. SmartArt, math, dll) — kalau ragu, jangan hapus.
    // Cek tag-tag content lain yang mungkin penting.
    if (p.getElementsByTagNameNS('*', 'oMath').length > 0)    return false;
    return true;
  }

  // Helper: cek apakah <w:p> ini carrier untuk final <w:sectPr> (Pola B).
  // Kalau iya, JANGAN dihapus — dokumen butuh section properties.
  function carriesSectPr(p) {
    return p.getElementsByTagNameNS('*', 'sectPr').length > 0;
  }

  // ── STEP 1: identifikasi posisi awal walk ──────────────────────────
  // Kalau child terakhir body adalah <w:sectPr> (Pola A), skip dia —
  // mulai walk dari sibling sebelumnya.
  let cursor = body.lastElementChild;
  if (cursor && cursor.localName === 'sectPr') {
    cursor = cursor.previousElementSibling;
  }

  // ── STEP 2: walk ke belakang, hapus paragraf kosong ────────────────
  let removed = 0;
  const MAX_REMOVE = 30;

  while (cursor && removed < MAX_REMOVE) {
    // Hanya proses <w:p>. Selain itu (mis. <w:tbl>, <w:sectPr> tak
    // terduga) → STOP.
    if (cursor.localName !== 'p') break;

    // Carrier sectPr (Pola B) → STOP, jangan disentuh.
    if (carriesSectPr(cursor)) break;

    // Punya konten meaningful → STOP.
    if (!isEmptyParagraph(cursor)) break;

    // Hapus paragraf kosong, geser cursor ke sibling sebelumnya.
    const prev = cursor.previousElementSibling;
    cursor.parentNode.removeChild(cursor);
    removed++;
    cursor = prev;
  }

  if (removed === 0) return;

  // Serialize kembali ke string XML
  let newXml;
  try {
    newXml = new XMLSerializer().serializeToString(xmlDoc);
  } catch(e) {
    console.warn('[9201] dropTrailingEmpty: serialize failed', e);
    return;
  }

  // Restore XML declaration kalau XMLSerializer drop
  if (xmlDecl && !newXml.startsWith('<?xml')) {
    newXml = xmlDecl + newXml;
  }

  // Write back ke zip
  try {
    zip.file(xmlPath, newXml);
    console.log(`[9201] dropTrailingEmpty: removed ${removed} empty trailing paragraph(s)`);
  } catch(e) {
    console.warn('[9201] dropTrailingEmpty: zip update failed', e);
  }
}

async function buildSuratTugasDoc(data, opts) {
  // opts: { jumlahResponden?: number|string|null }
  // jumlahResponden tidak disimpan ke DB — di-pass per-call dari UI
  // (modal Approve / dialog input saat preview/download).
  console.log('[9201] buildSuratTugasDoc() dipanggil', { suratId: data.id, opts });

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
    console.error('[9201] Docxtemplater / PizZip belum dimuat.',
      'window.docxtemplater =', typeof window.docxtemplater,
      'window.Docxtemplater =', typeof window.Docxtemplater,
      'window.PizZip =', typeof window.PizZip,
      'window.pizzip =', typeof window.pizzip);
    throw new Error(
      'Library docxtemplater / PizZip gagal dimuat. ' +
      'Periksa koneksi internet/firewall, lalu refresh halaman.'
    );
  }
  console.log('[9201] Docxtemplater OK — akan pakai template-based rendering');

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

  const templateData = await buildTemplateData(data, opts);
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
    console.warn('[9201] dropTrailingEmptyParagraphs error (non-fatal):', e);
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
  SESSION = novaCheckSession({ requireAdmin: true });
  if (!SESSION) return;
  Topbar9201.setUser(SESSION);
  initRoleSwitcher(SESSION, true);
  Promise.all([loadPegawai(), loadRiwayatJabatan(), loadUsers(), loadSurat()]);

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
