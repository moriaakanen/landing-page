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

   Hanya 3 file template di Supabase Storage (per keputusan user — pakai
   T1V & T3V sebagai superset untuk semua tipe non-visum):
     T2:  template-surat-tugas-lampiran.docx
          — punya {#ul} (tabel lampiran). Dipakai HANYA untuk tipe
            surat_tugas_lampiran (tanpa SPD).
     T1V: template-surat-tugas-spd-visum-kendaraan-menginap.docx
          — punya {#has_spd}, {#kendaraan}, {#menginap},
            {#visum}{#r}{/r}{/visum}. Dipakai untuk SEMUA tipe
            non-lampiran (visum maupun non-visum). Untuk non-visum:
            flags.has_visum=false → visum:[] → section visum hilang.
     T3V: template-surat-tugas-lampiran-spd-visum-kendaraan-menginap.docx
          — punya {#ul} (lampiran), {#kendaraan}, {#menginap}, dan
            {#visum}{#r}{/r}{/visum}. SPD selalu render (tidak ada
            {#has_spd} wrapper). Dipakai untuk SEMUA tipe lampiran-with-SPD.

   Catatan tag yg perlu Anda pastikan ADA di template:
     - T1V & T3V: blok kendaraan dibungkus {#kendaraan}...{/kendaraan}
                  blok menginap   dibungkus {#menginap}...{/menginap}
                  blok visum      dibungkus {#visum}...{/visum}
                  baris loop responden di-bungkus {#r}...{/r}
                  di baris terakhir tabel visum.
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
//
// KONSOLIDASI TEMPLATE (per keputusan user):
//   - T1V dipakai untuk SEMUA tipe non-lampiran (bukan hanya visum)
//     karena T1V adalah superset T1: identik dengan T1 + tambahan
//     section {#visum}{/visum} di akhir. Untuk tipe non-visum,
//     `flags.has_visum=false` → array `visum:[]` → seluruh section
//     visum hilang dari output (engine docxtemplater rule untuk
//     section dengan array kosong).
//   - T3V dipakai untuk SEMUA tipe lampiran (bukan hanya visum) dengan
//     alasan yang sama.
//   - T2 tetap dipakai untuk tipe surat_tugas_lampiran (tanpa SPD).
//
// Hanya 3 file yang perlu di-upload ke Storage:
//   - template-surat-tugas-lampiran.docx                              (T2)
//   - template-surat-tugas-spd-visum-kendaraan-menginap.docx          (T1V)
//   - template-surat-tugas-lampiran-spd-visum-kendaraan-menginap.docx (T3V)
const TEMPLATE_URL_T2  = 'https://jsmmtqeoukkgugorrvmg.supabase.co/storage/v1/object/public/template/template-surat-tugas-lampiran.docx';
const TEMPLATE_URL_T1V = 'https://jsmmtqeoukkgugorrvmg.supabase.co/storage/v1/object/public/template/template-surat-tugas-spd-visum-kendaraan-menginap.docx';
const TEMPLATE_URL_T3V = 'https://jsmmtqeoukkgugorrvmg.supabase.co/storage/v1/object/public/template/template-surat-tugas-lampiran-spd-visum-kendaraan-menginap.docx';

// Mapping tipe → template URL.
// Semua tipe non-lampiran → T1V (superset). Tipe lampiran-with-SPD → T3V.
// Hanya tipe surat_tugas_lampiran (tanpa SPD) yang masih pakai T2.
const TIPE_TO_TEMPLATE = {
  'surat_tugas':                                       TEMPLATE_URL_T1V,
  'surat_tugas_kendaraan':                             TEMPLATE_URL_T1V,
  'surat_tugas_visum_kendaraan':                       TEMPLATE_URL_T1V,
  'surat_tugas_lampiran':                              TEMPLATE_URL_T2,
  'surat_tugas_spd_kendaraan':                         TEMPLATE_URL_T1V,
  'surat_tugas_spd_kendaraan_menginap':                TEMPLATE_URL_T1V,
  'surat_tugas_spd_visum_kendaraan_menginap':          TEMPLATE_URL_T1V,
  'surat_tugas_lampiran_spd_kendaraan':                TEMPLATE_URL_T3V,
  'surat_tugas_lampiran_spd_kendaraan_menginap':       TEMPLATE_URL_T3V,
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

// Helper: apakah tipe surat ini punya halaman lampiran?
// Dipakai untuk toggle UI editor "Bertugas Sebagai" — hanya muncul kalau
// tipe ada lampiran (T2/T3V) DAN pegawai ≥ 2 (1 pegawai tidak butuh role
// pembeda karena kolom Jabatan/Bertugas hanya berisi 1 baris).
function tipeHasLampiran(tipe) {
  return tipeFlags(tipe).has_lampiran === true;
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
let mitraList   = [];             // dari tabel "mitra" (semua tahun)
let mitraByNip  = {};             // index NIP placeholder → object mitra
                                  //   format NIP placeholder: MITRA-{tahun}-{id3digit}
                                  //   contoh: MITRA-2026-042
let riwayatJabatan = [];          // dari "riwayat_jabatan"
let riwayatPangkatGolongan = [];  // dari "riwayat_pangkat_golongan"
let riwayatGelar = [];            // dari "riwayat_gelar"
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

function normalizeStatusKepegawaian(p) {
  return String(p && (p.status_kepegawaian || p.status_pegawai || '') || 'aktif').trim().toLowerCase();
}

function tanggalPensiunPegawai(p) {
  return String(p && (p.tanggal_pensiun || p.tmt_pensiun || '') || '').slice(0, 10);
}

function isPegawaiPensiunAt(p, isoRef) {
  if (!p || p._isMitra) return false;
  const status = normalizeStatusKepegawaian(p);
  if (status !== 'pensiun') return false;
  const tglPensiun = tanggalPensiunPegawai(p);
  if (!tglPensiun) return true;
  return String(isoRef || todayISO()).slice(0, 10) >= tglPensiun;
}

function statusPegawaiLabel(p, isoRef) {
  if (!isPegawaiPensiunAt(p, isoRef)) return '';
  const tglPensiun = tanggalPensiunPegawai(p);
  return tglPensiun ? `Pensiun per ${fmtTgl(tglPensiun)}` : 'Pensiun';
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
    // Ambil semua kolom supaya status_kepegawaian/tanggal_pensiun ikut
    // tersedia setelah migration dijalankan, tanpa perlu ubah select lagi.
    const res = await fetch(`${SUPABASE_URL}/rest/v1/data_pegawai?select=*&order=nama.asc`, { headers: H });
    if (!res.ok) return;
    pegawaiList = await res.json();
    pegawaiByNIP = {};
    pegawaiList.forEach(p => {
      const nip = String(pegawaiNip(p) || '').trim();
      if (nip) pegawaiByNIP[nip] = p;
    });
  } catch(e) { console.warn('Gagal load pegawai:', e); }
}

/* ════════════════════════════════════════════════════════════════════
   MITRA — load + helpers untuk integrasi ke picker pegawai
   ─────────────────────────────────────────────────────────────────────
   Mitra direkrut tahunan, di-load semua tahun supaya:
   - Picker form bisa filter ke mitra tahun aktif (tahun saat ini)
   - Surat lama yang sudah me-reference mitra tahun sebelumnya tetap bisa
     resolve nama-nya saat preview/regenerate (audit trail terjaga).

   NIP placeholder format: MITRA-{tahun}-{id_3digit}
     contoh: MITRA-2026-042 → mitra id=42 tahun 2026
   ─────────────────────────────────────────────────────────────────
   Kenapa pakai NIP placeholder?
   - Tabel surat_tugas.pegawai_nip[] paralel index dengan pegawai_list[].
     Mitra HARUS punya entry di pegawai_nip[] supaya tidak rusak indexing.
   - Format readable & deterministik — admin bisa langsung tahu ini mitra
     ID berapa, tahun apa, kalau debug DB.
═══════════════════════════════════════════════════════════════════════ */

// Format ID mitra → NIP placeholder. Pad id ke 3 digit untuk konsistensi
// visual (mis. ID 5 → "MITRA-2026-005").
function formatMitraNip(tahun, id) {
  return `MITRA-${tahun}-${String(id).padStart(3, '0')}`;
}

// Cek apakah NIP adalah placeholder mitra (bukan NIP pegawai asli).
// Pegawai asli NIP-nya numerik 18 digit, mitra prefixed "MITRA-".
function isMitraNip(nip) {
  return typeof nip === 'string' && nip.startsWith('MITRA-');
}

async function loadMitra() {
  try {
    const res = await fetch(`${SUPABASE_URL}/rest/v1/mitra?select=id,tahun,nama,no_hp,jabatan,instansi&order=tahun.desc,nama.asc`, { headers: H });
    if (!res.ok) {
      // Tabel mitra mungkin belum dibuat — fallback ke array kosong, jangan
      // hard-fail. Picker tetap berfungsi untuk pegawai biasa.
      console.warn('[9201] loadMitra: HTTP', res.status, '— skip (tabel mungkin belum ada)');
      return;
    }
    mitraList = await res.json();
    mitraByNip = {};
    mitraList.forEach(m => {
      const nip = formatMitraNip(m.tahun, m.id);
      mitraByNip[nip] = m;
    });
  } catch(e) {
    console.warn('Gagal load mitra:', e);
  }
}

async function loadRiwayatJabatan() {
  try {
    const res = await fetch(`${SUPABASE_URL}/rest/v1/riwayat_jabatan?select=*&order=tmt.desc`, { headers: H });
    if (!res.ok) return;
    riwayatJabatan = await res.json();
  } catch(e) { console.warn('Gagal load riwayat_jabatan:', e); }
}

// Load riwayat pangkat/golongan — dipakai untuk lookup {pangkat_golongan_p}
// per pegawai di buildPegawaiRow. Pattern sama dengan loadRiwayatJabatan:
// load semua sekali, lookup in-memory.
async function loadRiwayatPangkatGolongan() {
  try {
    const res = await fetch(`${SUPABASE_URL}/rest/v1/riwayat_pangkat_golongan?select=*&order=tmt.desc`, { headers: H });
    if (!res.ok) return;
    riwayatPangkatGolongan = await res.json();
  } catch(e) { console.warn('Gagal load riwayat_pangkat_golongan:', e); }
}

async function loadRiwayatGelar() {
  try {
    const res = await fetch(`${SUPABASE_URL}/rest/v1/riwayat_gelar?select=*&order=tmt.desc`, { headers: H });
    if (!res.ok) return;
    riwayatGelar = await res.json();
  } catch(e) { console.warn('Gagal load riwayat_gelar:', e); }
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
    // Tombol utama "Setujui" tetap pakai label — eye-catching karena ini
    // aksi paling penting di row status menunggu.
    aksi = `
      <button class="btn-approve" onclick="openApprove(${s.id})">✅ Setujui</button>`;
  } else if (isSelesai) {
    if (isEditing) {
      // Mode edit aktif — Simpan (label) & Batal (logo only).
      aksi = `
        <button class="btn-save-edit" onclick="saveRowEdit(${s.id})">💾 Simpan</button>
        <button class="btn-cancel-edit" onclick="cancelRowEdit(${s.id})" title="Batal edit">✕</button>`;
    } else {
      // Tombol "Edit Role" muncul kalau tipe ada lampiran (T2/T3V) DAN
      // ≥ 2 pegawai. Sama dengan kondisi show-bertugas di modal Approve.
      // Semua tombol pakai LOGO ONLY supaya kolom aksi compact.
      const showEditRole = tipeHasLampiran(s.tipe)
                           && Array.isArray(s.pegawai_list)
                           && s.pegawai_list.length >= 2;
      const btnEditRole = showEditRole
        ? `<button class="btn-edit-row btn-icon" onclick="openEditBertugas(${s.id})" title="Edit role per pegawai" style="background:#5b3a8a;color:#fff;border-color:#5b3a8a">✏️</button>`
        : '';
      aksi = `
        <button class="btn-edit-row btn-icon" onclick="enableRowEdit(${s.id})" title="Edit field surat ini">✏️</button>
        ${btnEditRole}
        <button class="btn-preview btn-icon" onclick="openPreview(${s.id})" title="Preview surat">👁</button>
        <button class="btn-download btn-icon" onclick="downloadSuratTugas(${s.id})" title="Download .docx">📥</button>`;
    }
  } else {
    // Status tidak dikenal — tampilkan placeholder
    aksi = `<span style="font-size:11px;color:var(--muted);font-style:italic">—</span>`;
  }

  // Checkbox dual-purpose:
  //   - Baris 'selesai'  → bulk-dl-check  (untuk bulk download, navy default)
  //   - Baris 'menunggu' → bulk-approve-check (untuk bulk approve, hijau via .ck-success)
  // Keduanya pakai kolom yang sama supaya layout konsisten.
  // Styling visual checkbox di-handle global oleh 9201-shared.js — di sini
  // cukup tambahkan class .ck-success untuk varian hijau.
  const checkCell = isSelesai
    ? `<td class="col-check"><input type="checkbox" class="bulk-dl-check" data-surat-id="${s.id}" onchange="updateBulkDownloadCounter()"></td>`
    : `<td class="col-check"><input type="checkbox" class="bulk-approve-check ck-success" data-surat-id="${s.id}" onchange="updateBulkApproveCounter()" title="Pilih untuk bulk approve"></td>`;

  // data-editing sebagai marker tambahan agar styling/CSS bisa membedakan
  // baris yg sedang di-edit (ditambah border kuning di seksi CSS).
  return `
    <tr data-surat-id="${s.id}" data-status="${s.status}"${isEditing ? ' data-editing="1"' : ''}>
      ${checkCell}
      <td class="col-no">${urutNo}</td>

      ${cellTextHTML(s.id, 'nomor_surat', nomorSurat, editable, 'cth: 001 / 013A')}
      ${cellDateHTML(s.id, 'tanggal_surat', tanggalSurat, editable, 'tgl/bln/thn')}
      ${cellDateRangeHTML(s.id, 'waktu', waktuMulai, waktuSelesai, editable, s.waktu_pelaksanaan_text)}
      ${cellTextareaHTML(s.id, 'perihal', perihal, editable, 'Perihal surat')}
      ${cellTextareaHTML(s.id, 'tujuan', tujuan, editable, 'Kota/instansi')}
      ${cellPegawaiMultiHTML(s.id, pegNips, pegNames, editable)}
      ${cellTextareaHTML(s.id, 'menimbang_custom', menimbang, editable, 'cth: pelaksanaan Survei...')}
      ${cellTextareaHTML(s.id, 'alat_angkutan', alat, editable, 'cth: Kendaraan Darat')}
      ${cellMAKHTML(s.id, mak, editable)}
      ${cellPenandatanganHTML(s.id, ttdNip, ttdNama, editable)}
      ${cellTipeHTML(s.id, s.tipe, editable)}

      <td class="col-aksi"><div class="aksi-wrap">${aksi}</div></td>
      <td class="col-status">${badgeHTML(s.status)}</td>
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
  // #8: Auto-save draft tiap blur untuk semua field editable.
  // Wire ke semua textarea/input di scope (row atau full table).
  root.querySelectorAll('textarea.xls-cell, input.xls-cell, input.waktu-custom').forEach(el => {
    el.addEventListener('blur', () => {
      const tr = el.closest('tr[data-surat-id]');
      if (tr) snapshotAdminDraft(tr.dataset.suratId);
    });
  });
  // Pegawai multi-tag dan penandatangan: trigger snapshot saat tag berubah
  // via MutationObserver pada attribute data-nips / data-nip.
  root.querySelectorAll('.pg-cell').forEach(pg => {
    if (pg.dataset.draftWired) return;
    pg.dataset.draftWired = '1';
    const obs = new MutationObserver(() => {
      const tr = pg.closest('tr[data-surat-id]');
      if (tr) snapshotAdminDraft(tr.dataset.suratId);
    });
    obs.observe(pg, { attributes: true, attributeFilter: ['data-nips', 'data-nip'] });
  });
}

/* ════════════════════════════════════════════════════════════════════
   #8: ADMIN DRAFT PERSIST
   ─────────────────────────────────────────────────────────────────────
   Strategi:
     - Per-row: key = nova_st_admin_draft_{userId}_{suratId}
     - Snapshot tiap blur field
     - Restore dipanggil setelah loadSurat() selesai render
     - Cleanup: setelah save row sukses (saveRowEdit / submitApprove) atau
       saat user explicit cancel edit (cancelRowEdit)
   ════════════════════════════════════════════════════════════════════ */
function getAdminDraftKey(suratId) {
  let uid = 'anon';
  try {
    const s = JSON.parse(localStorage.getItem('nova_user') || 'null');
    if (s && s.id) uid = String(s.id);
  } catch (_) {}
  return `nova_st_admin_draft_${uid}_${suratId}`;
}
function snapshotAdminDraft(suratId) {
  const values = collectRow(suratId);
  if (!values) return;
  // Cek apakah ada perubahan substansial dari original. Kalau row kosong
  // total (mungkin baru), tetap save supaya state baru di-preserve.
  const sOrig = suratMap[suratId];
  const isEmpty = !values.nomor_surat && !values.tanggal_surat
                  && !values.tanggal_berangkat && !values.perihal
                  && !values.tujuan && !values.menimbang_custom
                  && !values.alat_angkutan && !values.pembebanan
                  && !values.penandatangan_nip && !values.tipe
                  && (!values.pegawai_nip || !values.pegawai_nip.length);
  // Cek diff dengan original — kalau identik, tidak perlu save.
  const isUnchanged = sOrig
    && (values.nomor_surat   || '') === (sOrig.nomor_surat   || '')
    && (values.tanggal_surat || '') === (sOrig.tanggal_surat || '')
    && (values.tanggal_berangkat || '') === (sOrig.tanggal_berangkat || '')
    && (values.tanggal_kembali   || '') === (sOrig.tanggal_kembali   || '')
    && (values.waktu_pelaksanaan_text || '') === (sOrig.waktu_pelaksanaan_text || '')
    && (values.perihal       || '') === (sOrig.perihal       || '')
    && (values.tujuan        || '') === (sOrig.tujuan        || '')
    && (values.menimbang_custom || '') === (sOrig.menimbang_custom || '')
    && (values.alat_angkutan || '') === (sOrig.alat_angkutan || '')
    && (values.pembebanan    || '') === (sOrig.pembebanan    || '')
    && (values.penandatangan_nip  || '') === (sOrig.penandatangan_nip  || '')
    && (values.penandatangan_nama || '') === (sOrig.penandatangan_nama || '')
    && (values.tipe || '') === (sOrig.tipe || '')
    && JSON.stringify(values.pegawai_nip || []) === JSON.stringify(sOrig.pegawai_nip || []);
  try {
    if (isEmpty || isUnchanged) {
      localStorage.removeItem(getAdminDraftKey(suratId));
    } else {
      localStorage.setItem(getAdminDraftKey(suratId), JSON.stringify({
        version: 1,
        savedAt: Date.now(),
        suratId: suratId,
        values: values,
      }));
      // Tampilkan toast "Draft tersimpan" — ringan, auto-hide 1.5 detik.
      showDraftSavedToast();
    }
  } catch (e) {
    console.warn('[Admin Draft] gagal save:', e);
  }
}
function clearAdminDraft(suratId) {
  try { localStorage.removeItem(getAdminDraftKey(suratId)); } catch(_) {}
}
function loadAdminDraft(suratId) {
  try {
    const raw = localStorage.getItem(getAdminDraftKey(suratId));
    if (!raw) return null;
    const obj = JSON.parse(raw);
    if (!obj || !obj.values) return null;
    return obj;
  } catch (_) { return null; }
}
/* Apply draft values ke row DOM. Dipanggil setelah renderTable() selesai. */
function applyAdminDraftToRow(suratId) {
  const draft = loadAdminDraft(suratId);
  if (!draft) return false;
  const row = document.querySelector(`tr[data-surat-id="${suratId}"]`);
  if (!row) return false;
  const v = draft.values;
  // Helper update text field
  const setVal = (field, val) => {
    const el = row.querySelector(`[data-field="${field}"]`);
    if (el && 'value' in el) el.value = val || '';
  };
  setVal('nomor_surat',      v.nomor_surat);
  setVal('perihal',          v.perihal);
  setVal('tujuan',           v.tujuan);
  setVal('menimbang_custom', v.menimbang_custom);
  setVal('alat_angkutan',    v.alat_angkutan);
  setVal('pembebanan',       v.pembebanan);
  // Tanggal
  const tglSurat = row.querySelector('[data-field="tanggal_surat"]');
  if (tglSurat && v.tanggal_surat) {
    tglSurat.dataset.iso = v.tanggal_surat;
    tglSurat.value = fmtTgl(v.tanggal_surat);
  }
  const waktu = row.querySelector('[data-field="waktu"]');
  if (waktu) {
    waktu.dataset.isoMulai = v.tanggal_berangkat || '';
    waktu.dataset.isoSelesai = v.tanggal_kembali || '';
    waktu.value = v.tanggal_berangkat ? fmtWaktu(v.tanggal_berangkat, v.tanggal_kembali) : '';
  }
  // Waktu custom
  const waktuCustom = row.querySelector('input.waktu-custom[data-field="waktu_text"]');
  const waktuToggle = row.querySelector('.waktu-toggle');
  if (waktuCustom && v.waktu_pelaksanaan_text) {
    waktuCustom.value = v.waktu_pelaksanaan_text;
    if (waktuCustom.classList.contains('hidden') && waktuToggle) {
      toggleWaktuCustom(waktuToggle);
      waktuCustom.value = v.waktu_pelaksanaan_text;
    }
  }
  // Pegawai multi
  const pegCell = row.querySelector('[data-field="pegawai_multi"]');
  if (pegCell && Array.isArray(v.pegawai_nip) && v.pegawai_nip.length) {
    pegCell.dataset.nips = JSON.stringify(v.pegawai_nip);
    pegCell.dataset.names = JSON.stringify(v.pegawai_list || []);
    // Hapus tag lama, render ulang
    pegCell.querySelectorAll('.pg-tag').forEach(t => t.remove());
    const inp = pegCell.querySelector('.pg-input');
    const tags = v.pegawai_nip.map((nip, j) => buildPegTag(nip, (v.pegawai_list || [])[j] || nip, false)).join('');
    if (inp) {
      pegCell.insertAdjacentHTML('afterbegin', tags);
      inp.placeholder = '';
    }
  }
  // Penandatangan single
  const ttdCell = row.querySelector('[data-field="penandatangan"]');
  if (ttdCell && v.penandatangan_nip) {
    ttdCell.dataset.nip = v.penandatangan_nip;
    ttdCell.dataset.nama = v.penandatangan_nama || '';
    ttdCell.querySelectorAll('.pg-tag').forEach(t => t.remove());
    const inp = ttdCell.querySelector('.pg-input');
    const tag = buildPegTag(v.penandatangan_nip, v.penandatangan_nama || v.penandatangan_nip, true);
    if (inp) {
      ttdCell.insertAdjacentHTML('afterbegin', tag);
      inp.placeholder = '';
    }
  }
  // Tipe (custom dropdown — pakai data-value pada .tp-cell)
  const tpCell = row.querySelector('.tp-cell[data-field="tipe"]');
  if (tpCell && v.tipe) {
    tpCell.dataset.value = v.tipe;
    const tpDisplay = tpCell.querySelector('.tp-display');
    if (tpDisplay && typeof tipeLabel === 'function') {
      tpDisplay.textContent = tipeLabel(v.tipe);
      tpDisplay.classList.remove('placeholder');
    }
  }
  return true;
}
/* Restore semua draft yang ada untuk surat-surat yang lagi tampil. */
function restoreAllAdminDrafts() {
  let restoredCount = 0;
  Array.from(document.querySelectorAll('tr[data-surat-id]')).forEach(tr => {
    const id = tr.dataset.suratId;
    if (loadAdminDraft(id)) {
      // Untuk surat selesai, draft hanya boleh restore kalau editingRowId
      // sama dengan id (= user sebelumnya sedang edit baris ini).
      // Tapi karena editingRowId tidak persist (cuma in-memory state),
      // kita cek: kalau status selesai dan tidak sedang di-edit, draft
      // tidak applicable — skip.
      const s = suratMap[id];
      if (!s) return;
      if (s.status === 'selesai' && editingRowId !== Number(id)) {
        // Draft ada tapi user belum klik "Edit" — biarkan saja, akan
        // restore otomatis kalau user enable edit.
        return;
      }
      if (applyAdminDraftToRow(id)) restoredCount++;
    }
  });
  if (restoredCount > 0) {
    showPageAlert(`📝 ${restoredCount} draft dimuat dari sesi terakhir.`, 'info');
  }
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
        <th class="col-check col-check-dual" title="☐ atas = centang semua selesai (bulk download) | ☐ bawah = centang semua menunggu (bulk approve)">
          <div class="th-check-inner">
            <input type="checkbox" id="bulk-dl-master" onchange="toggleBulkDownloadAll(this.checked)" title="Centang semua surat selesai untuk bulk download">
            <input type="checkbox" id="bulk-ap-master" class="ap-master ck-success" onchange="toggleBulkApproveAll(this.checked)" title="Centang semua surat menunggu untuk bulk approve">
          </div>
        </th>
        ${sortHeader('no',            'No',                'col-no')}
        ${sortHeader('nomor_surat',   'Nomor Surat',       'col-nomor-surat')}
        ${sortHeader('tanggal_surat', 'Tanggal Surat',     'col-tgl-surat')}
        ${sortHeader('waktu',         'Waktu Pelaksanaan', 'col-waktu')}
        ${sortHeader('perihal',       'Perihal',           'col-perihal')}
        ${sortHeader('tujuan',        'Tempat Tujuan',     'col-tujuan')}
        <th class="col-nama">Nama</th>
        <th class="col-menimbang">Menimbang</th>
        <th class="col-alat">Alat Angkutan</th>
        <th class="col-mak">POK</th>
        <th class="col-ttd">Penandatangan</th>
        <th class="col-tipe">Tipe</th>
        <th class="col-aksi">Aksi</th>
        ${sortHeader('status',        'Status',            'col-status')}
        ${sortHeader('pengaju',       'Diajukan oleh',     'col-pengaju')}
      </tr></thead>
      <tbody>${rows}</tbody>
    </table>`;

  requestAnimationFrame(() => {
    attachEditableListeners();
    setupTopScrollbar();
    // Reset state bulk download (master + counter) setelah re-render
    updateBulkDownloadCounter();
    // Reset state bulk approve
    updateBulkApproveCounter();
    // #8: Restore admin drafts (kalau ada) untuk row-row yang lagi tampil.
    // Dipanggil setelah attachEditableListeners supaya MutationObserver-nya
    // tidak trigger save saat applying draft (bisa kena race condition).
    setTimeout(() => restoreAllAdminDrafts(), 50);
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

function cellDateRangeHTML(id, field, isoMulai, isoSelesai, editable, waktuTextCustom) {
  // Read-only display (status selesai, tidak dalam mode edit):
  // Kalau ada waktu_pelaksanaan_text → tampilkan text-nya (override).
  // Kalau tidak → format default dari range tanggal.
  if (!editable) {
    const hasCustom = !!(waktuTextCustom && waktuTextCustom.trim());
    const displayText = hasCustom ? waktuTextCustom : (isoMulai ? fmtWaktu(isoMulai, isoSelesai) : '');
    const isEmpty = !displayText;
    return `<td class="col-waktu">
      <div class="ro-text${isEmpty ? ' muted' : ''}"
           tabindex="0" data-col-field="${field}">${isEmpty ? '—' : esc(displayText)}</div>
    </td>`;
  }
  // Editable mode: picker tanggal range + toggle expand untuk multi-range.
  // Pattern sama dengan user-side (surat-tugas.html) — toggle .waktu-toggle
  // yang reveal field .waktu-custom (data-field=waktu_text). Picker tanggal
  // tetap di atas untuk struktur DB (tanggal_berangkat / tanggal_kembali).
  const display = isoMulai ? fmtWaktu(isoMulai, isoSelesai) : '';
  const hasCustom = !!(waktuTextCustom && waktuTextCustom.trim());
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
    <button type="button" class="waktu-toggle${hasCustom ? ' active' : ''}"
      onclick="toggleWaktuCustom(this)"
      title="Tambah teks waktu kustom untuk multi-range">
      <span class="waktu-toggle-icon">${hasCustom ? '\u2212' : '+'}</span>
      ${hasCustom ? 'kustom' : 'multi-range'}
    </button>
    <input type="text" class="waktu-custom${hasCustom ? '' : ' hidden'}" data-field="waktu_text"
      value="${escAttr(waktuTextCustom || '')}"
      placeholder="Misal: 25 s.d. 27 Maret &amp; 30 Maret s.d. 1 April 2026">
  </td>`;
}

/* Toggle visibility text input "Waktu Custom" untuk multi-range.
   Pattern sama dengan user-side (surat-tugas.html). Di-call dari
   onclick tombol .waktu-toggle. */
function toggleWaktuCustom(btn){
  const td=btn.closest('td.col-waktu');
  if(!td)return;
  const customInp=td.querySelector('.waktu-custom');
  if(!customInp)return;
  const willShow=customInp.classList.contains('hidden');
  customInp.classList.toggle('hidden',!willShow);
  btn.classList.toggle('active',willShow);
  // Update icon & label tombol
  const iconEl=btn.querySelector('.waktu-toggle-icon');
  if(iconEl)iconEl.textContent=willShow?'\u2212':'+';
  const textNode=Array.from(btn.childNodes).find(n=>n.nodeType===3&&n.textContent.trim());
  if(textNode)textNode.textContent=willShow?'kustom':'multi-range';
  if(willShow){
    setTimeout(()=>customInp.focus(),50);
  } else {
    customInp.value='';
  }
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
      pickPegawai(cellEl, String(pegawaiNip(p)).trim(), pegawaiNama(p) || '');
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

function getRowTanggalSuratForCell(cellEl) {
  const tr = cellEl ? cellEl.closest('tr[data-surat-id]') : null;
  const el = tr ? tr.querySelector('[data-field="tanggal_surat"]') : null;
  return (el && el.dataset && el.dataset.iso) ? el.dataset.iso : todayISO();
}

function acFilter(q) {
  q = (q || '').toLowerCase().trim();
  const cellEl = acState.cellEl;
  const refDate = getRowTanggalSuratForCell(cellEl);
  let selectedNips = [];
  if (cellEl) {
    if (acState.isSingle) {
      const nip = cellEl.dataset.nip;
      if (nip) selectedNips = [nip];
    } else {
      try { selectedNips = JSON.parse(cellEl.dataset.nips || '[]'); } catch(_) {}
    }
  }

  // Build combined pool: pegawai (existing) + mitra (filtered ke tahun aktif).
  // Mitra ditandai _isMitra=true supaya picker bisa render badge "Mitra".
  // NIP placeholder MITRA-{tahun}-{id} di-generate via formatMitraNip().
  //
  // Mitra hanya dari "tahun aktif" (= tahun saat ini) supaya picker bersih.
  // Mitra tahun lama tetap di mitraList & mitraByNip untuk lookup record
  // surat lama, tapi tidak muncul di dropdown form baru.
  //
  // EXCLUDE MITRA UNTUK SINGLE PICKER (penandatangan): mitra tidak boleh
  // jadi penandatangan surat tugas, jadi pool-nya cuma pegawai.
  // acState.isSingle = true → cell adalah penandatangan; false → pegawai_multi.
  let combinedPool;
  if (acState.isSingle) {
    combinedPool = pegawaiList.map(p => ({
      ...p,
      _disabled: isPegawaiPensiunAt(p, refDate),
      _disabledReason: statusPegawaiLabel(p, refDate),
    }));
  } else {
    const tahunAktif = new Date().getFullYear();
    const mitraPool = mitraList
      .filter(m => m.tahun === tahunAktif)
      .map(m => ({
        pegawai_nip: formatMitraNip(m.tahun, m.id),
        nama: m.nama,
        _isMitra: true,
        _mitraTahun: m.tahun,
      }));
    const pegawaiPool = pegawaiList.map(p => ({
      ...p,
      _disabled: isPegawaiPensiunAt(p, refDate),
      _disabledReason: statusPegawaiLabel(p, refDate),
    }));
    combinedPool = pegawaiPool.concat(mitraPool);
  }

  acState.filtered = combinedPool.filter(p => {
    if (!q) return true;
    const nama = (pegawaiNama(p) || '').toLowerCase();
    const nip  = String(pegawaiNip(p) || '').toLowerCase();
    return nama.includes(q) || nip.includes(q);
  }).slice(0, 50);
  acRenderList(selectedNips);
  document.getElementById('ac-count').textContent = `${acState.filtered.length} hasil`;
}

function acRenderList(selectedNips) {
  const list = document.getElementById('ac-list');
  // Cek apakah pool kosong: pegawai DAN mitra dua-duanya kosong
  if (!pegawaiList.length && !mitraList.length) {
    list.innerHTML = `<div class="ac-empty">Data pegawai/mitra belum dimuat.</div>`;
    return;
  }
  if (!acState.filtered.length) {
    list.innerHTML = `<div class="ac-empty">Tidak ada hasil — coba ketik nama lain</div>`;
    return;
  }
  list.innerHTML = acState.filtered.map((p, i) => {
    const nip   = String(pegawaiNip(p) || '').trim();
    const nama  = pegawaiNama(p) || '-';
    const isSel = selectedNips.includes(nip);
    // Badge "Mitra" di samping nama untuk membedakan dari pegawai biasa.
    // Untuk pegawai biasa, tampilkan badge "Pegawai" supaya konsisten —
    // user-side tidak perlu nebak mana mitra mana bukan.
    const badge = p._isMitra
      ? `<span class="ac-badge mitra">MITRA</span>`
      : (p._disabled
          ? `<span class="ac-badge pensiun">PENSIUN</span>`
          : `<span class="ac-badge pegawai">PEGAWAI</span>`);
    // Untuk mitra: tampilkan "Mitra Tahun YYYY" di tempat NIP (NIP placeholder
    // tidak menarik untuk admin — tahunnya lebih informatif).
    const subLine = p._isMitra
      ? `Mitra Tahun ${p._mitraTahun}`
      : (p._disabledReason || `NIP ${nip || '-'}`);
    return `<div class="ac-item${isSel ? ' selected' : ''}${i === acState.focusIdx ? ' focused' : ''}${p._disabled ? ' disabled' : ''}"
      data-nip="${escAttr(nip)}" data-idx="${i}"
      onmousedown="event.preventDefault();acPick(${i})">
      <div class="ac-check">${isSel ? '✓' : ''}</div>
      <div style="flex:1">
        <div class="ac-name">${esc(nama)}${badge}</div>
        <div class="ac-nip">${esc(subLine)}</div>
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
  if (p._disabled) {
    showPageAlert(p._disabledReason || 'Pegawai ini tidak bisa dipilih untuk tanggal surat tersebut.', 'error');
    return;
  }
  pickPegawai(acState.cellEl, String(pegawaiNip(p)).trim(), pegawaiNama(p) || '');
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
         onpointerdown="event.preventDefault(); pickMAK(${jsArg(s.mak)})">
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
    return `<div class="${cls}" data-idx="${i}" onpointerdown="event.preventDefault(); pickTipe(${jsArg(o.value)})">${esc(o.label)}</div>`;
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
     ENTER di textarea: plain Enter = pindah ke sel berikutnya (Tab),
     Alt+Enter = newline. Sebelumnya Enter di textarea = newline default
     browser, sekarang diubah atas permintaan user.
     ENTER di input single-line: no-op (tidak ada newline, tidak navigasi)
     karena field-field seperti nomor_surat, tanggal_surat, MAK
     sudah punya handler Tab sendiri.                                   */
  if (isReadonly) return; // readonly: biarkan Tab default browser

  // Enter di textarea.xls-cell
  if (e.key === 'Enter' && isTextarea && isXlsCell) {
    if (e.altKey) {
      // Alt+Enter: izinkan newline default browser — jangan intercept
      return;
    }
    // Plain Enter (tanpa Alt/Ctrl/Shift): pindah ke sel berikutnya
    e.preventDefault();
    closeAllPopups();
    moveCellFocus(target, 1);
    return;
  }

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
  // #3: waktu_pelaksanaan_text dari toggle expand di kolom waktu.
  // Field .waktu-custom muncul kalau toggle .waktu-toggle ter-expand.
  // Kalau collapsed atau kosong → return string kosong → nanti jadi NULL di payload.
  const waktuTextEl = row.querySelector('input.waktu-custom[data-field="waktu_text"]');
  const waktuTextCustom = waktuTextEl ? (waktuTextEl.value || '').trim() : '';

  return {
    nomor_surat:       get('nomor_surat'),
    tanggal_surat:     getISO('tanggal_surat'),
    tanggal_berangkat: waktu.mulai,
    tanggal_kembali:   waktu.selesai,
    waktu_pelaksanaan_text: waktuTextCustom,
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
    // OPSIONAL (boleh kosong): tujuan, pembebanan, alat_angkutan
    //  - tujuan kosong       → {untuk_text} fallback ke "{perihal} pada tanggal ..."
    //  - pembebanan kosong   → kolom MAK di template render kosong
    //  - alat_angkutan kosong → kolom alat angkutan render kosong
    ['menimbang_custom',   'Menimbang'],
    ['penandatangan_nama', 'Penandatangan'],
    ['tipe',               'Tipe Surat'],
  ];
  const errors = [], errFields = [];
  checks.forEach(([k, label]) => {
    const v = values[k];
    if (!v || (Array.isArray(v) && !v.length)) { errors.push(label); errFields.push(k); }
  });
  if (!values.pegawai_nip.length && !values.pegawai_list.length) {
    errors.push('Nama');
    errFields.push('pegawai_multi');
  }
  // Validasi format MAK Pembebanan — hanya jika field-nya terisi.
  // Kalau kosong, sudah di-tangkap oleh required-check di atas.
  // Format wajib: 054.01.GG.2910.BMA.006.054.A.524119
  if (values.pembebanan && !parseMAK(values.pembebanan)) {
    errors.push('Format POK tidak valid (contoh: 054.01.GG.2910.BMA.006.054.A.524119)');
    if (!errFields.includes('pembebanan')) errFields.push('pembebanan');
  }
  const availability = validatePegawaiAvailability(values);
  availability.errors.forEach(msg => errors.push(msg));
  availability.errFields.forEach(f => {
    if (!errFields.includes(f)) errFields.push(f);
  });
  return { errors, errFields };
}

function validatePegawaiAvailability(values) {
  const errors = [];
  const errFields = [];
  const tglSurat = values.tanggal_surat || '';
  const tugasMulai = values.tanggal_berangkat || '';
  const tugasSelesai = values.tanggal_kembali || values.tanggal_berangkat || '';

  const addField = (f) => {
    if (!errFields.includes(f)) errFields.push(f);
  };
  const checkPegawai = (nip, nama, field) => {
    if (!nip || isMitraNip(nip)) return;
    const p = pegawaiByNIP[String(nip).trim()];
    if (!p) return;
    const tglPensiun = tanggalPensiunPegawai(p);
    if (isPegawaiPensiunAt(p, tglSurat)) {
      errors.push(`${nama || pegawaiNama(p) || nip} sudah pensiun pada tanggal surat${tglPensiun ? ` (${fmtTgl(tglPensiun)})` : ''}.`);
      addField(field);
      return;
    }
    if (normalizeStatusKepegawaian(p) === 'pensiun' && tglPensiun && tugasMulai && tugasMulai >= tglPensiun) {
      errors.push(`${nama || pegawaiNama(p) || nip} tidak bisa ditugaskan karena waktu pelaksanaan dimulai setelah pensiun.`);
      addField(field);
    } else if (normalizeStatusKepegawaian(p) === 'pensiun' && tglPensiun && tugasSelesai && tugasSelesai >= tglPensiun) {
      errors.push(`${nama || pegawaiNama(p) || nip} pensiun per ${fmtTgl(tglPensiun)}, sedangkan waktu pelaksanaan melewati tanggal tersebut.`);
      addField(field);
    }
  };

  (values.pegawai_nip || []).forEach((nip, idx) => checkPegawai(nip, (values.pegawai_list || [])[idx], 'pegawai_multi'));
  checkPegawai(values.penandatangan_nip, values.penandatangan_nama, 'penandatangan');

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
  // File preview dijadwalkan hapus 10 menit sejak dibuat.
  if (id === 'modal-preview' && _previewUploadedPath) {
    schedulePreviewCleanup(_previewUploadedPath);
  }
  // Reset state khusus modal preview
  if (id === 'modal-preview') {
    _previewVisumOpts = null;
    _previewBlob = null;
  }
  // Reset state edit-bertugas saat modal ditutup
  if (id === 'modal-edit-bertugas') {
    _editBertugasSuratId = null;
  }
}

function showPageAlert(msg, type='error') {
  const alert = document.getElementById('page-alert');
  // Map type ke icon. Default ke error icon.
  const iconMap = { success: '✅', error: '⚠️', info: 'ℹ️' };
  document.getElementById('page-alert-icon').textContent = iconMap[type] || iconMap.error;
  document.getElementById('page-alert-text').textContent = msg;
  alert.className = `alert ${type} show`;
  alert.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
  setTimeout(() => { alert.className = 'alert'; }, 5500);
}

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

  const ttdJabatan = lookupJabatan(values.penandatangan_nip, values.tanggal_surat);
  const nomorFull  = buildNomorSuratFull(values.nomor_surat, values.tanggal_surat);

  // Info badge tipe — kalau tipe ada visum, tampilkan reminder agar admin
  // tahu nanti perlu input jumlah responden saat preview/download.
  const tipeBadge = tipeHasVisum(values.tipe)
    ? `${esc(tipeLabel(values.tipe))} <span style="display:inline-block;margin-left:6px;background:rgba(200,168,75,.18);color:#7a5c10;border:1px solid rgba(200,168,75,.4);border-radius:100px;padding:1px 7px;font-size:10px;font-weight:600;letter-spacing:.3px">📋 VISUM</span>`
    : esc(tipeLabel(values.tipe));

  // Helper untuk render row dengan styling "empty" kalau value kosong.
  // Konsisten visual di seluruh preview — UI lebih clean.
  const row = (label, val, opts) => {
    opts = opts || {};
    const isEmpty = !val || (typeof val === 'string' && !val.trim());
    const display = isEmpty ? '—' : val;
    return `<div class="approve-preview-row"><strong>${label}</strong><span class="${isEmpty ? 'empty' : ''}"${opts.style ? ' style="' + opts.style + '"' : ''}>${opts.html ? display : esc(display)}</span></div>`;
  };

  document.getElementById('approve-preview').innerHTML = `
    ${row('Nomor', nomorFull, { style: 'font-family:ui-monospace,monospace;font-size:11.5px' })}
    ${row('Tgl Surat', fmtTgl(values.tanggal_surat))}
    ${row('Waktu', fmtWaktu(values.tanggal_berangkat, values.tanggal_kembali))}
    ${row('Perihal', values.perihal)}
    ${row('Tujuan', values.tujuan)}
    ${row('Pegawai', values.pegawai_list.join(', '))}
    ${row('Menimbang', values.menimbang_custom)}
    ${row('Alat Angkutan', values.alat_angkutan)}
    ${row('POK', values.pembebanan)}
    <div class="approve-preview-row"><strong>Penandatangan</strong><span>
      ${esc(values.penandatangan_nama)}<br>
      <em style="color:var(--muted);font-style:italic;font-size:11px">${ttdJabatan ? esc(ttdJabatan) : '<span style="color:var(--red)">⚠ Jabatan tidak ditemukan di riwayat</span>'}</em><br>
      <span style="color:var(--muted);font-size:11px">NIP. ${esc(values.penandatangan_nip)}</span>
    </span></div>
    ${row('Tipe Surat', tipeBadge, { html: true })}
  `;
  // Catatan: input "Waktu Pelaksanaan Kustom" di modal Approve dihapus
  // (#3). Sekarang admin set multi-range langsung di tabel via toggle
  // "+ multi-range" di kolom Waktu Pelaksanaan. waktu_pelaksanaan_text
  // sudah ter-collect di values via collectRow().

  // ─── Toggle & render section "Bertugas Sebagai" ──────────────────
  // Section muncul HANYA kalau:
  //   1. Tipe surat punya lampiran (T2/T3V), DAN
  //   2. Jumlah pegawai >= 2 (kalau cuma 1 orang, tidak butuh role pembeda)
  // Kondisi 2 disengaja — tabel lampiran 1-baris tidak butuh kolom
  // "bertugas" karena tidak ada apa-apa untuk dibedakan.
  const showBertugas = tipeHasLampiran(values.tipe) && values.pegawai_list.length >= 2;
  const bertugasRow  = document.getElementById('approve-bertugas-row');
  if (bertugasRow) bertugasRow.style.display = showBertugas ? '' : 'none';
  if (showBertugas) {
    // Pre-fill dari nilai yg sudah ada di DB (kalau surat ini sebelumnya
    // pernah disetujui & sudah punya bertugas_sebagai), atau dari draft
    // sebelumnya kalau admin batal lalu buka modal lagi. Default: array kosong.
    const existing = Array.isArray(s.bertugas_sebagai) ? s.bertugas_sebagai : [];
    renderBertugasList(
      'approve-bertugas-list',
      values.pegawai_list,
      existing,
      'inp-bertugas-row'
    );
    // Reset quick-fill input setiap kali modal dibuka
    const quick = document.getElementById('inp-bertugas-quick');
    if (quick) quick.value = '';
  }

  // #2: Toggle section "Jumlah Responden" — hanya muncul untuk tipe visum.
  // Pre-fill dari kolom DB surat_tugas.jumlah_responden (kalau sebelumnya
  // pernah di-set). Kosong = pakai default 7 baris saat render.
  const showResponden = tipeHasVisum(values.tipe);
  const respondenRow  = document.getElementById('approve-responden-row');
  if (respondenRow) respondenRow.style.display = showResponden ? '' : 'none';
  const respondenInp = document.getElementById('inp-jumlah-responden');
  if (respondenInp) {
    respondenInp.value = (s.jumlah_responden != null) ? String(s.jumlah_responden) : '';
  }

  document.getElementById('approve-alert').className = 'alert';
  openModal('modal-approve');
}

/* ════════════════════════════════════════════════════════════════════
   BERTUGAS SEBAGAI — UI helpers
   ─────────────────────────────────────────────────────────────────────
   Render list editor + apply-all/clear-all + collect values.
   Dipakai oleh modal-approve (saat setujui) dan modal-edit-bertugas
   (edit setelah surat selesai).

   Pattern: setiap input punya id `${idPrefix}-${index}` supaya bisa
   di-collect via for-loop. Index sesuai posisi di array pegawai.
═══════════════════════════════════════════════════════════════════════ */

/**
 * Render list pegawai dengan input "Bertugas Sebagai" di sebelahnya.
 *
 * @param {string}    containerId  ID elemen container yg akan diisi
 * @param {string[]}  pegawaiNames Array nama pegawai (urutan = index)
 * @param {string[]}  existing     Array nilai bertugas yg sudah ada
 *                                 (panjang bisa < pegawaiNames; sisanya '')
 * @param {string}    idPrefix     Prefix untuk id input. Mis. 'inp-bertugas-row'
 *                                 → input id = 'inp-bertugas-row-0', '...row-1', dst.
 */
function renderBertugasList(containerId, pegawaiNames, existing, idPrefix) {
  const container = document.getElementById(containerId);
  if (!container) return;
  container.innerHTML = pegawaiNames.map((nm, i) => {
    const val = String(existing[i] || '').replace(/"/g, '&quot;');
    return `
      <div style="display:flex;gap:8px;align-items:center">
        <span style="flex:0 0 28px;text-align:right;font-size:12px;color:var(--muted);font-weight:600">${i + 1}.</span>
        <span style="flex:1;font-size:12.5px;color:var(--navy);font-weight:500;min-width:0;overflow:hidden;text-overflow:ellipsis;white-space:nowrap" title="${esc(nm)}">${esc(nm)}</span>
        <input type="text"
               id="${idPrefix}-${i}"
               value="${val}"
               placeholder="Bertugas sebagai…"
               style="flex:0 0 240px;padding:6px 9px;border:1.5px solid var(--border2);border-radius:6px;font-family:'Plus Jakarta Sans',sans-serif;font-size:12.5px;outline:none">
      </div>
    `;
  }).join('');
}

/**
 * Collect nilai dari list bertugas. Trim setiap value, return array dengan
 * panjang sesuai pegawaiCount.
 *
 * @param {number} pegawaiCount  Jumlah pegawai (= panjang array yg di-return)
 * @param {string} idPrefix      Prefix id input (sama dengan render*)
 * @returns {string[]}  Array string. Element kosong = '' (bukan null).
 */
function collectBertugasValues(pegawaiCount, idPrefix) {
  const out = [];
  for (let i = 0; i < pegawaiCount; i++) {
    const el = document.getElementById(`${idPrefix}-${i}`);
    out.push(el ? String(el.value || '').trim() : '');
  }
  return out;
}

/** Apply nilai dari input quick ke semua field bertugas — modal Approve. */
function bertugasApplyAll() {
  const v = (document.getElementById('inp-bertugas-quick')?.value || '').trim();
  document.querySelectorAll('#approve-bertugas-list input[id^="inp-bertugas-row-"]')
    .forEach(el => { el.value = v; });
}

/** Kosongkan semua field bertugas — modal Approve. */
function bertugasClearAll() {
  document.querySelectorAll('#approve-bertugas-list input[id^="inp-bertugas-row-"]')
    .forEach(el => { el.value = ''; });
}

/** Apply quick value ke semua field — modal Edit Bertugas. */
function editBertugasApplyAll() {
  const v = (document.getElementById('inp-edit-bertugas-quick')?.value || '').trim();
  document.querySelectorAll('#edit-bertugas-list input[id^="inp-edit-bertugas-row-"]')
    .forEach(el => { el.value = v; });
}

/** Kosongkan semua field — modal Edit Bertugas. */
function editBertugasClearAll() {
  document.querySelectorAll('#edit-bertugas-list input[id^="inp-edit-bertugas-row-"]')
    .forEach(el => { el.value = ''; });
}

/* ════════════════════════════════════════════════════════════════════
   EDIT BERTUGAS — open modal & submit handler
   ─────────────────────────────────────────────────────────────────────
   Modal khusus untuk admin mengedit role pegawai pada surat yang sudah
   selesai (status=selesai). Hanya field bertugas_sebagai yang di-PATCH;
   surat lainnya tidak disentuh.

   Diakses dari tombol "✏ Edit Role" di tabel (kolom Aksi). Tombol
   tersebut hanya muncul kalau tipeHasLampiran(s.tipe) && pegawai >= 2.
═══════════════════════════════════════════════════════════════════════ */
let _editBertugasSuratId = null;  // surat ID yang sedang di-edit

/** Buka modal Edit Bertugas untuk surat tertentu. */
function openEditBertugas(suratId) {
  const s = suratMap[suratId];
  if (!s) return;
  // Defensive: kalau dari refresh data tipe sudah berubah jadi non-lampiran
  // atau jumlah pegawai jadi < 2, tampilkan alert dan abort.
  if (!tipeHasLampiran(s.tipe)) {
    showPageAlert('Tipe surat tidak punya halaman lampiran — role tidak relevan.', 'error');
    return;
  }
  const pegawai = Array.isArray(s.pegawai_list) ? s.pegawai_list : [];
  if (pegawai.length < 2) {
    showPageAlert('Surat dengan 1 pegawai tidak butuh role bertugas.', 'error');
    return;
  }

  _editBertugasSuratId = suratId;
  document.getElementById('edit-bertugas-perihal').textContent = s.perihal || '—';
  // Reset quick-fill
  const quick = document.getElementById('inp-edit-bertugas-quick');
  if (quick) quick.value = '';

  // Render list pakai existing data
  const existing = Array.isArray(s.bertugas_sebagai) ? s.bertugas_sebagai : [];
  renderBertugasList(
    'edit-bertugas-list',
    pegawai,
    existing,
    'inp-edit-bertugas-row'
  );

  openModal('modal-edit-bertugas');
}

/** Submit perubahan ke DB via PATCH. Hanya kolom bertugas_sebagai. */
async function submitEditBertugas() {
  const id = _editBertugasSuratId;
  if (!id) return;
  const s = suratMap[id];
  if (!s) return;

  const pegawaiCount = Array.isArray(s.pegawai_list) ? s.pegawai_list.length : 0;
  const arr = collectBertugasValues(pegawaiCount, 'inp-edit-bertugas-row');

  // Sama logic dengan submitApprove: kalau semua kosong, kirim NULL
  // (tidak ngotorin DB dengan array empty-strings).
  const payload = {
    bertugas_sebagai: arr.some(v => v.length > 0) ? arr : null,
  };

  const btn = document.getElementById('btn-edit-bertugas-submit');
  if (btn) { btn.disabled = true; btn.classList.add('loading'); }

  try {
    const res = await fetch(`${SUPABASE_URL}/rest/v1/surat_tugas?id=eq.${id}`, {
      method: 'PATCH',
      headers: { ...H, 'Prefer': 'return=minimal' },
      body: JSON.stringify(payload),
    });
    if (!res.ok) {
      let msg = `HTTP ${res.status}`;
      try { const j = await res.json(); msg = j.message || msg; } catch(_) {}
      throw new Error(msg);
    }

    // Update suratMap in-memory supaya tidak perlu reload semua data
    s.bertugas_sebagai = payload.bertugas_sebagai;

    closeModal('modal-edit-bertugas');
    _editBertugasSuratId = null;
    showPageAlert('✅ Role pegawai berhasil disimpan.', 'success');
  } catch(e) {
    console.error('[9201] submitEditBertugas:', e);
    showPageAlert(`Gagal menyimpan: ${e.message}`, 'error');
  } finally {
    if (btn) { btn.disabled = false; btn.classList.remove('loading'); }
  }
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

  // #3: waktu_pelaksanaan_text sudah ter-collect di values via collectRow()
  // dari toggle expand di kolom waktu (bukan lagi dari modal).
  // Kalau kosong setelah trim, kirim NULL (fallback ke auto format dari
  // tanggal_berangkat & tanggal_kembali saat render docx).
  const waktuPelaksanaanText = values.waktu_pelaksanaan_text && values.waktu_pelaksanaan_text.length
    ? values.waktu_pelaksanaan_text
    : null;

  const payload = {
    status: 'selesai',
    nomor_surat:           values.nomor_surat,
    tanggal_surat:         values.tanggal_surat,
    tanggal_berangkat:     values.tanggal_berangkat,
    tanggal_kembali:       values.tanggal_kembali || null,
    waktu_pelaksanaan_text: waktuPelaksanaanText,
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
    // bertugas_sebagai: hanya disertakan kalau tipe ada lampiran & ≥2 pegawai
    // (kondisi yang sama dengan show-condition di openApprove). Untuk tipe
    // lain, NULL — biarkan engine docxtemplater lookup → string kosong.
    // collectBertugasValues() return string[] — element kosong = '' (string).
    // Kalau SEMUA element kosong, kita kirim NULL juga supaya tidak nyampah
    // di DB. Cek: ada minimal 1 element non-empty?
    bertugas_sebagai:      (tipeHasLampiran(values.tipe) && values.pegawai_list.length >= 2)
                             ? (() => {
                                 const arr = collectBertugasValues(values.pegawai_list.length, 'inp-bertugas-row');
                                 return arr.some(v => v.length > 0) ? arr : null;
                               })()
                             : null,
    // #2: jumlah_responden — hanya disertakan kalau tipe surat ada visum.
    // Untuk non-visum, kirim NULL (kolom unused). Empty string → NULL juga.
    jumlah_responden:      (tipeHasVisum(values.tipe))
                             ? (() => {
                                 const el = document.getElementById('inp-jumlah-responden');
                                 const v = el ? el.value.trim() : '';
                                 if (!v) return null;
                                 const n = parseInt(v, 10);
                                 return (isFinite(n) && n >= 0) ? n : null;
                               })()
                             : null,
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

    // Auto-save defaults (alat_angkutan/POK/penandatangan) tanpa checkbox.
    // Sebelumnya ada checkbox "Simpan default" yang dihapus dari UI —
    // sekarang behavior-nya selalu auto-save kalau ada nilai. Admin
    // tidak perlu mikir lagi, dan halaman pengajuan baru langsung
    // pakai defaults yang sudah pernah dipakai sebelumnya.
    {
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
    // #8: cleanup draft surat ini setelah sukses approve.
    clearAdminDraft(selectedId);
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
  // #8: cleanup draft saat user explicit cancel edit.
  clearAdminDraft(id);
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

  // #3: waktu_pelaksanaan_text sekarang bisa di-edit langsung di tabel
  // via toggle expand di kolom waktu. Ambil dari values (collectRow).
  // Kalau kosong → kirim NULL.
  const waktuTextForEdit = values.waktu_pelaksanaan_text && values.waktu_pelaksanaan_text.length
    ? values.waktu_pelaksanaan_text
    : null;

  // Payload TANPA field 'status' — status tetap 'selesai'.
  const payload = {
    nomor_surat:           values.nomor_surat,
    tanggal_surat:         values.tanggal_surat,
    tanggal_berangkat:     values.tanggal_berangkat,
    tanggal_kembali:       values.tanggal_kembali || null,
    waktu_pelaksanaan_text: waktuTextForEdit,
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

  // Edge case: kalau admin edit pegawai_list (tambah/hapus/ubah urutan),
  // index di bertugas_sebagai tidak lagi valid — reset ke NULL agar admin
  // ulangi via tombol "Edit Role". Kita compare urutan NIP lama vs baru.
  // Kalau sama persis → biarkan (tidak ikutkan di payload). Kalau beda →
  // include bertugas_sebagai: null di payload.
  const sBefore = suratMap[id];
  const oldNips = Array.isArray(sBefore && sBefore.pegawai_nip) ? sBefore.pegawai_nip : [];
  const newNips = values.pegawai_nip;
  const nipsChanged = oldNips.length !== newNips.length
                      || oldNips.some((n, i) => String(n).trim() !== String(newNips[i] || '').trim());
  if (nipsChanged && Array.isArray(sBefore && sBefore.bertugas_sebagai) && sBefore.bertugas_sebagai.length) {
    payload.bertugas_sebagai = null;
    console.log('[9201] saveRowEdit: pegawai_list berubah → bertugas_sebagai di-reset ke NULL');
  }

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
    // #8: cleanup draft surat ini setelah sukses save edit.
    clearAdminDraft(id);
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
  // Early exit untuk mitra — mitra tidak punya riwayat_jabatan (selalu "Mitra").
  // Caller (buildPegawaiRow) akan handle jabatan mitra dari record mitra
  // sendiri. Skip filter array supaya hemat compute.
  if (isMitraNip(nip)) return '';
  // Filter jenis='utama' — jabatan struktural sehari-hari pegawai. Jabatan
  // jenis='lainnya' (Plt., PPK, dll) di-lookup terpisah via lookupJabatanLainnya.
  // Defensive: kalau row tidak punya kolom jenis (kompatibel data lama),
  // anggap 'utama' (cocokkan dengan default DB).
  const candidates = riwayatJabatan
    .filter(r => String(r.pegawai_nip || '').trim() === String(nip).trim())
    .filter(r => (r.jenis || 'utama') === 'utama')
    .filter(r => r.tmt && r.tmt <= tglSuratIso)
    .sort((a, b) => (b.tmt || '').localeCompare(a.tmt || ''));
  if (candidates.length) return candidates[0].jabatan || '';
  const peg = pegawaiByNIP[nip];
  const pegNama = pegawaiNama(peg);
  if (peg && pegNama) {
    const candByName = riwayatJabatan
      .filter(r => (r.nama || '').trim().toLowerCase() === pegNama.trim().toLowerCase())
      .filter(r => (r.jenis || 'utama') === 'utama')
      .filter(r => r.tmt && r.tmt <= tglSuratIso)
      .sort((a, b) => (b.tmt || '').localeCompare(a.tmt || ''));
    if (candByName.length) return candByName[0].jabatan || '';
  }
  return '';
}

/**
 * Lookup jabatan_lainnya (Plt., PPK, dll) by nama jabatan & tanggal surat.
 * Return record terbaru yang TMT-nya <= tanggal_surat.
 *
 * Pattern paralel dengan lookupJabatan tapi:
 *   - Filter jenis='lainnya'
 *   - Filter jabatan = nama yang dicari (mis. 'Pejabat Pembuat Komitmen')
 *
 * Dipakai untuk:
 *   - {nama_ppk} & {nip_ppk} → cari "Pejabat Pembuat Komitmen"
 *   - Detect apakah pegawai sedang Plt. (untuk transformJabatanPenandatangan)
 *
 * @param {string} jabatanLainnya  Nama jabatan persis (case-sensitive)
 * @param {string} tglSuratIso     Tanggal surat (YYYY-MM-DD)
 * @returns {object|null} record riwayat_jabatan, atau null kalau tidak ketemu
 */
function lookupJabatanLainnya(jabatanLainnya, tglSuratIso) {
  if (!jabatanLainnya || !tglSuratIso) return null;
  const candidates = riwayatJabatan
    .filter(r => (r.jenis || '') === 'lainnya')
    .filter(r => (r.jabatan || '') === jabatanLainnya)
    .filter(r => r.tmt && r.tmt <= tglSuratIso)
    .filter(r => !r.tmt_selesai || r.tmt_selesai >= tglSuratIso)
    .sort((a, b) => (b.tmt || '').localeCompare(a.tmt || ''));
  return candidates[0] || null;
}

/**
 * Lookup pangkat & golongan terbaru per pegawai by NIP & tanggal surat.
 * Return record dari riwayat_pangkat_golongan, atau null.
 *
 * @param {string} nip          NIP pegawai (skip kalau mitra)
 * @param {string} tglSuratIso  Tanggal surat (YYYY-MM-DD)
 * @returns {object|null} {pangkat, golongan, tmt, ...} atau null
 */
function lookupPangkatGolongan(nip, tglSuratIso) {
  if (!nip || !tglSuratIso || isMitraNip(nip)) return null;
  const candidates = riwayatPangkatGolongan
    .filter(r => String(r.pegawai_nip || '').trim() === String(nip).trim())
    .filter(r => r.tmt && r.tmt <= tglSuratIso)
    .sort((a, b) => (b.tmt || '').localeCompare(a.tmt || ''));
  return candidates[0] || null;
}

function lookupGelar(nip, tglSuratIso) {
  if (!nip || !tglSuratIso || isMitraNip(nip)) return '';
  const candidates = riwayatGelar
    .filter(r => String(r.pegawai_nip || '').trim() === String(nip).trim())
    .filter(r => r.tmt && r.tmt <= tglSuratIso)
    .sort((a, b) => (b.tmt || '').localeCompare(a.tmt || ''));
  return candidates[0] ? (candidates[0].gelar || '') : '';
}

/**
 * Transform jabatan_penandatangan sesuai rule bisnis.
 * Input value WAJIB salah satu dari 3 nilai canonical (di-validate di UI/DB):
 *   - "Kepala BPS Kabupaten Raja Ampat"      → tampilkan apa adanya
 *   - "Plt. Kepala Badan Pusat Statistik"    → "Plt. Kepala BPS Kabupaten Raja Ampat"
 *   - lainnya                                 → "Plh. Kepala BPS Kabupaten Raja Ampat"
 *
 * Logic ini di-compute saat render, BUKAN saat approve. Konsekuensinya:
 * kalau di masa depan rule diubah (mis. "Plh." jadi sesuatu yang lain),
 * surat lama yang di-regenerate akan ikut rule baru. Untuk audit trail,
 * teks asli yang user pilih tetap tersimpan di kolom penandatangan_jabatan.
 */
function transformJabatanPenandatangan(jab) {
  const j = String(jab || '').trim();
  if (j === 'Kepala BPS Kabupaten Raja Ampat')      return j;
  if (j === 'Plt. Kepala Badan Pusat Statistik')    return 'Plt. Kepala BPS Kabupaten Raja Ampat';
  return 'Plh. Kepala BPS Kabupaten Raja Ampat';
}

/**
 * Compute {awalan} dari perihal surat tugas.
 * Rule: kalau perihal mengandung kata "pelatihan" (case-insensitive,
 *       partial match) → "Daftar Peserta", lainnya → "Daftar Petugas".
 */
function computeAwalan(perihal) {
  return /pelatihan/i.test(String(perihal || ''))
    ? 'Daftar Peserta'
    : 'Daftar Petugas';
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
   File otomatis dijadwalkan hapus 10 menit sejak preview dibuat.
   Orphan files (kalau user crash/close paksa) dibersihkan saat halaman load
   via cleanupOrphanPreviewFiles().
─────────────────────────────────────────────────────────────────── */
const PREVIEW_BUCKET = 'surat-tugas-preview';
const PREVIEW_TTL_MS = 10 * 60 * 1000;
const PREVIEW_SIGNED_URL_TTL_SEC = Math.ceil(PREVIEW_TTL_MS / 1000);
const _previewCleanupTimers = {};
let _previewUploadedPath = null;
let _previewBlob = null;
// Opts visum yg sedang aktif untuk modal preview saat ini. Dipakai oleh
// downloadFromPreview() dan openInWordForPrint() agar tombol-tombol di
// modal preview pakai jumlah responden yang sama dengan preview yg sedang
// ditampilkan, tanpa tanya ulang ke user. Reset saat modal preview ditutup.
let _previewVisumOpts = null;

/* ── Visum prompt state ──────────────────────────────────────────────
   #2 refactor: jumlah_responden sekarang persistent di DB (kolom
   surat_tugas.jumlah_responden — di-input admin saat approve). Tidak ada
   lagi modal prompt on-demand di preview/download. Code lama dihapus:
     - promptVisumResponden() / confirmVisumPrompt() helpers
     - _visumLastInput[] / _visumPromptResolver state
     - modal-visum-prompt HTML
─────────────────────────────────────────────────────────────────── */

/**
 * Wrapper: jalankan callback dengan opts yang sudah di-resolve.
 *
 * #2 (refactor): jumlah_responden sekarang persistent di DB (kolom
 * surat_tugas.jumlah_responden — di-input admin saat approve). Tidak ada
 * lagi prompt modal di preview/download. Function ini sekarang sekedar
 * pull value dari record surat & pass ke callback.
 *
 * @param {object}   surat    - object surat (harus include jumlah_responden field)
 * @param {Function} callback - async (opts) => any
 * @returns {Promise<any>}    - return value callback
 */
async function withVisumOpts(surat, callback) {
  if (!tipeHasVisum(surat && surat.tipe)) {
    return callback({});  // tipe non-visum → opts kosong
  }
  // Baca jumlah_responden dari record DB. NULL = pakai default (7 baris).
  const jr = (surat && surat.jumlah_responden != null) ? surat.jumlah_responden : null;
  return callback({ jumlahResponden: jr });
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
  schedulePreviewCleanup(filename);
  return filename;
}

async function getPreviewSignedUrl(filename, expiresInSec = PREVIEW_SIGNED_URL_TTL_SEC) {
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

function buildPreviewViewerUrl(signedUrl) {
  return `https://view.officeapps.live.com/op/embed.aspx?src=${encodeURIComponent(signedUrl)}`;
}

async function deletePreviewFile(filename) {
  if (!filename) return;
  clearPreviewCleanupTimer(filename);
  try {
    await fetch(`${SUPABASE_URL}/storage/v1/object/${PREVIEW_BUCKET}/${filename}`, {
      method: 'DELETE',
      headers: { apikey: SUPABASE_ANON_KEY, Authorization: `Bearer ${SUPABASE_ANON_KEY}` },
    });
  } catch (e) { console.warn('[9201] Cleanup preview file gagal:', e); }
}

function clearPreviewCleanupTimer(filename) {
  const timer = _previewCleanupTimers[filename];
  if (timer) {
    clearTimeout(timer);
    delete _previewCleanupTimers[filename];
  }
}

function schedulePreviewCleanup(filename, delayMs = PREVIEW_TTL_MS) {
  if (!filename || _previewCleanupTimers[filename]) return;
  const delay = Math.max(0, Number(delayMs) || PREVIEW_TTL_MS);
  _previewCleanupTimers[filename] = setTimeout(() => {
    delete _previewCleanupTimers[filename];
    deletePreviewFile(filename);
  }, delay);
}

/* Defensive cleanup: hapus file preview lama (>10 menit) yang mungkin
   ter-orphan kalau user close paksa. Fire-and-forget, dipanggil di init(). */
async function cleanupOrphanPreviewFiles() {
  try {
    const res = await fetch(`${SUPABASE_URL}/storage/v1/object/list/${PREVIEW_BUCKET}`, {
      method: 'POST',
      headers: { ...H },
      body: JSON.stringify({ limit: 1000, prefix: '' }),
    });
    if (!res.ok) return;
    const files = await res.json();
    const expiredBefore = Date.now() - PREVIEW_TTL_MS;
    files.forEach(f => {
      // Filename pattern: "{suratId}_{timestamp}.docx"
      const m = f.name && f.name.match(/_(\d+)\.docx$/);
      if (f.name && /^pak_/.test(f.name)) return;
      if (m && parseInt(m[1], 10) < expiredBefore) deletePreviewFile(f.name);
    });
  } catch (e) { /* fire-and-forget */ }
}

function ensureLibrariesLoaded() {
  if (typeof saveAs === 'undefined') {
    throw new Error('Library FileSaver gagal dimuat. Refresh halaman.');
  }
}

/* #5: Lazy-load SheetJS untuk export/import Excel.
   SheetJS hanya di-load saat user klik tombol Export/Import — tidak
   membebani initial page load. Pakai pattern Promise + caching. */
let _sheetjsLoadPromise = null;
function ensureSheetJSLoaded() {
  if (typeof window.XLSX !== 'undefined' && window.XLSX.utils) {
    return Promise.resolve();
  }
  if (_sheetjsLoadPromise) return _sheetjsLoadPromise;
  _sheetjsLoadPromise = new Promise((resolve, reject) => {
    const s = document.createElement('script');
    s.src = 'https://cdn.sheetjs.com/xlsx-0.20.0/package/dist/xlsx.full.min.js';
    s.async = true;
    s.onload = () => {
      if (typeof window.XLSX !== 'undefined') resolve();
      else reject(new Error('SheetJS load tidak return XLSX object'));
    };
    s.onerror = () => {
      _sheetjsLoadPromise = null;  // izinkan retry
      reject(new Error('Gagal load library Excel. Periksa koneksi internet.'));
    };
    document.head.appendChild(s);
  });
  return _sheetjsLoadPromise;
}

/* ════════════════════════════════════════════════════════════════════
   #5: EXPORT & IMPORT EXCEL
   ─────────────────────────────────────────────────────────────────────
   Format: XLSX (Excel native) via SheetJS.
   Export: semua kolom utama (lengkap) dari semua surat user.
   Import: parse → match pegawai by NAMA exact → INSERT ke DB sebagai
           surat 'menunggu' (admin lalu approve via tabel).
   ════════════════════════════════════════════════════════════════════ */

// Definisi header kolom Excel — urutan, label, dan field DB-nya.
// Single source of truth untuk export DAN import (consistency).
const EXCEL_COLUMNS = [
  // Field input-able (boleh diisi user di Excel, akan di-import balik)
  { header: 'Nomor Surat',           field: 'nomor_surat',            importable: true },
  { header: 'Tanggal Surat',         field: 'tanggal_surat',          importable: true,  isDate: true },
  { header: 'Tanggal Berangkat',     field: 'tanggal_berangkat',      importable: true,  isDate: true },
  { header: 'Tanggal Kembali',       field: 'tanggal_kembali',        importable: true,  isDate: true },
  { header: 'Waktu Pelaksanaan Custom', field: 'waktu_pelaksanaan_text', importable: true },
  { header: 'Perihal',               field: 'perihal',                importable: true },
  { header: 'Tempat Tujuan',         field: 'tujuan',                 importable: true },
  { header: 'Pegawai',               field: 'pegawai_list',           importable: true,  isList: true },
  { header: 'Bertugas Sebagai',      field: 'bertugas_sebagai',       importable: true,  isList: true },
  { header: 'Menimbang',             field: 'menimbang_custom',       importable: true },
  { header: 'Alat Angkutan',         field: 'alat_angkutan',          importable: true },
  { header: 'POK (Pembebanan)',      field: 'pembebanan',             importable: true },
  { header: 'Penandatangan',         field: 'penandatangan_nama',     importable: true,
    importMatchField: 'penandatangan_nip' /* match by nama, dapat NIP */ },
  { header: 'Tipe Surat',            field: 'tipe',                   importable: true },
  { header: 'Jumlah Responden',      field: 'jumlah_responden',       importable: true,  isNumber: true },
  { header: 'Tempat Terbit',         field: 'tempat_terbit',          importable: true },
  // Field NON-input-able (read-only — hanya tampil di export, di-skip saat import)
  { header: 'Status',                field: 'status',                 importable: false },
  { header: 'Diajukan Oleh',         field: '_pengaju_nama',          importable: false },
  { header: 'Dibuat',                field: 'created_at',             importable: false, isDate: true },
];

/**
 * EXPORT: Generate XLSX file dari semua surat (allSurat) yang sedang tampil
 * di tabel admin (sudah di-filter & di-sort). Download via saveAs.
 */
async function exportSuratToExcel() {
  if (!Array.isArray(allSurat) || !allSurat.length) {
    showPageAlert('Belum ada data surat untuk di-export.', 'error');
    return;
  }
  try {
    showPageAlert('⏳ Memuat library Excel...', 'info');
    await ensureSheetJSLoaded();

    // Build data array — header row + data rows.
    // SheetJS aoa_to_sheet expects 2D array.
    const headers = EXCEL_COLUMNS.map(c => c.header);
    const rows = allSurat.map(s => EXCEL_COLUMNS.map(c => extractFieldForExport(s, c)));
    const aoa = [headers, ...rows];

    const ws = XLSX.utils.aoa_to_sheet(aoa);

    // Auto-width per kolom — hitung max length tiap kolom (sederhana).
    ws['!cols'] = headers.map((h, ci) => {
      let maxLen = h.length;
      for (let r = 1; r < aoa.length; r++) {
        const cell = aoa[r][ci];
        const len = cell != null ? String(cell).length : 0;
        if (len > maxLen) maxLen = len;
      }
      // Cap width supaya tidak terlalu lebar (mis. kolom Pegawai bisa panjang)
      return { wch: Math.min(Math.max(maxLen + 2, 10), 45) };
    });

    // Freeze header row supaya saat scroll, header tetap visible
    ws['!freeze'] = { xSplit: 0, ySplit: 1 };
    ws['!autofilter'] = { ref: XLSX.utils.encode_range({
      s: { c: 0, r: 0 },
      e: { c: headers.length - 1, r: aoa.length - 1 }
    })};

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Surat Tugas');

    // Filename pakai timestamp supaya unik
    const now = new Date();
    const ts = `${now.getFullYear()}${String(now.getMonth()+1).padStart(2,'0')}${String(now.getDate()).padStart(2,'0')}_${String(now.getHours()).padStart(2,'0')}${String(now.getMinutes()).padStart(2,'0')}`;
    const filename = `Surat-Tugas_Export_${ts}.xlsx`;

    XLSX.writeFile(wb, filename);
    showPageAlert(`✅ Berhasil export ${allSurat.length} surat ke ${filename}.`, 'success');
  } catch (e) {
    console.error('[9201] export gagal:', e);
    showPageAlert(`Gagal export: ${e.message}`, 'error');
  }
}

/* Helper: ekstrak nilai dari surat object untuk kolom Excel tertentu.
   Mengembalikan string/number yang siap masuk cell Excel. */
function extractFieldForExport(s, col) {
  if (col.field === '_pengaju_nama') {
    return getPengajuNama(s) || '';
  }
  let v = s[col.field];
  if (v == null) return '';
  if (col.isList && Array.isArray(v)) {
    // pegawai_list, bertugas_sebagai → joined dengan ", "
    return v.filter(x => x != null && x !== '').join(', ');
  }
  if (col.isDate && typeof v === 'string') {
    // ISO format → biarkan seperti itu (Excel akan deteksi sebagai date)
    return v;
  }
  if (col.isNumber) {
    const n = parseInt(v, 10);
    return isFinite(n) ? n : '';
  }
  return String(v);
}

let suratImportState = {
  fileName: '',
  rows: [],
  payloads: [],
  warnings: [],
};

const SURAT_IMPORT_ALIASES = {
  nomor_surat: ['nomor surat', 'nomor surat tugas'],
  tanggal_surat: ['tanggal surat', 'tanggal surat tugas'],
  tanggal_berangkat: ['tanggal berangkat'],
  tanggal_kembali: ['tanggal kembali'],
  waktu_pelaksanaan_text: ['waktu pelaksanaan', 'waktu pelaksanaan custom', 'waktu'],
  perihal: ['perihal', 'tujuan/tugas', 'tujuan / tugas', 'tujuan tugas', 'tugas'],
  tujuan: ['tempat tujuan', 'tujuan'],
  pegawai_list: ['pegawai', 'nama'],
  bertugas_sebagai: ['bertugas sebagai'],
  menimbang_custom: ['menimbang'],
  alat_angkutan: ['alat angkutan'],
  pembebanan: ['pok', 'pok (pembebanan)', 'mak pembebanan', 'pembebanan'],
  _program: ['program'],
  penandatangan_nama: ['penandatangan'],
  tipe: ['tipe', 'tipe surat'],
  jumlah_responden: ['jumlah responden'],
  tempat_terbit: ['tempat terbit'],
};

function normalizeImportHeader(value) {
  return String(value || '')
    .toLowerCase()
    .replace(/\s+/g, ' ')
    .replace(/\s*\/\s*/g, '/')
    .trim();
}

function buildSuratImportHeaderMap(headerRow) {
  const aliasMap = {};
  Object.entries(SURAT_IMPORT_ALIASES).forEach(([field, aliases]) => {
    aliases.forEach(alias => { aliasMap[normalizeImportHeader(alias)] = field; });
  });
  EXCEL_COLUMNS.forEach(col => {
    if (col.importable) aliasMap[normalizeImportHeader(col.header)] = col.field;
  });

  const colIdxMap = {};
  headerRow.forEach((header, index) => {
    const field = aliasMap[normalizeImportHeader(header)];
    if (field && colIdxMap[field] === undefined) colIdxMap[field] = index;
  });
  return colIdxMap;
}

function findSuratImportHeaderRow(aoa) {
  const maxRows = Math.min(10, aoa.length);
  for (let i = 0; i < maxRows; i++) {
    const map = buildSuratImportHeaderMap(aoa[i] || []);
    if (map.perihal !== undefined || map.nomor_surat !== undefined || map.pegawai_list !== undefined) {
      return i;
    }
  }
  return -1;
}

function normalizeImportPersonName(value) {
  return String(value || '')
    .toLowerCase()
    .replace(/\b(s\.?\s*tr\.?\s*stat\.?|a\.?\s*md\.?\s*stat\.?|m\.?\s*ec\.?\s*dev\.?|s\.?\s*e\.?|s\.?\s*s\.?\s*t\.?|sst)\b/g, ' ')
    .replace(/[^a-z0-9]+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function makePegawaiNameMatcher() {
  const entries = [];
  const byKey = {};
  pegawaiList.forEach(p => {
    const nama = String(pegawaiNama(p) || '').trim();
    const nip = String(pegawaiNip(p) || '').trim();
    if (!nama || !nip) return;
    const keys = [
      nama.toLowerCase(),
      normalizeImportPersonName(nama),
    ].filter(Boolean);
    keys.forEach(key => { if (!byKey[key]) byKey[key] = p; });
    entries.push({ key: normalizeImportPersonName(nama), pegawai: p });
  });
  entries.sort((a, b) => b.key.length - a.key.length);
  return { byKey, entries };
}

function matchImportPeople(raw, matcher) {
  const text = String(raw || '').trim();
  if (!text) return { people: [], unmatched: [] };
  const exact = matcher.byKey[text.toLowerCase()] || matcher.byKey[normalizeImportPersonName(text)];
  if (exact) return { people: [exact], unmatched: [] };

  const normalizedRaw = normalizeImportPersonName(text);
  const found = [];
  const seen = new Set();
  matcher.entries.forEach(entry => {
    if (!entry.key || entry.key.length < 4) return;
    if (normalizedRaw.includes(entry.key)) {
      const nip = String(pegawaiNip(entry.pegawai) || '');
      if (nip && !seen.has(nip)) {
        seen.add(nip);
        found.push(entry.pegawai);
      }
    }
  });

  if (found.length) return { people: found, unmatched: [] };

  const parts = text.split(/[;\n]+|,(?=\s*[A-Z])/).map(s => s.trim()).filter(Boolean);
  const partMatches = [];
  const unmatched = [];
  parts.forEach(part => {
    const p = matcher.byKey[part.toLowerCase()] || matcher.byKey[normalizeImportPersonName(part)];
    if (p) partMatches.push(p);
    else unmatched.push(part);
  });
  return { people: partMatches, unmatched: unmatched.length ? unmatched : [text] };
}

function combineProgramMak(program, mak) {
  const cleanProgram = String(program || '').trim().replace(/\.$/, '');
  const cleanMak = String(mak || '').trim().replace(/^\./, '');
  if (!cleanProgram) return cleanMak || '';
  if (!cleanMak) return cleanProgram || '';
  if (cleanMak.toLowerCase().startsWith(`${cleanProgram.toLowerCase()}.`)) return cleanMak;
  return `${cleanProgram}.${cleanMak}`.replace(/\.+/g, '.');
}

function makeIsoDate(year, month, day) {
  year = Number(year); month = Number(month); day = Number(day);
  if (!year || month < 1 || month > 12 || day < 1 || day > 31) return '';
  return `${year}-${String(month).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
}

function extractDatesFromIndonesianText(raw) {
  const text = String(raw || '').toLowerCase();
  const dates = new Set();
  const addDate = (day, monthName, year) => {
    const month = BULAN_KEYS[String(monthName || '').toLowerCase()];
    const iso = makeIsoDate(year, month, day);
    if (iso) dates.add(iso);
  };

  const rangeRe = /(\d{1,2})(?:\s+([a-z]+))?(?:\s+(\d{4}))?\s*(?:s\.?\s*d\.?|sd|s\/d|sampai|-)\s*(\d{1,2})\s+([a-z]+)\s+(\d{4})/gi;
  let match;
  while ((match = rangeRe.exec(text))) {
    const endMonth = match[5];
    const endYear = match[6];
    addDate(match[1], match[2] || endMonth, match[3] || endYear);
    addDate(match[4], endMonth, endYear);
  }

  const fullDateRe = /(\d{1,2})\s+([a-z]+)\s+(\d{4})/gi;
  while ((match = fullDateRe.exec(text))) {
    addDate(match[1], match[2], match[3]);
  }

  const isoRe = /(\d{4})-(\d{1,2})-(\d{1,2})/g;
  while ((match = isoRe.exec(text))) {
    const iso = makeIsoDate(match[1], match[2], match[3]);
    if (iso) dates.add(iso);
  }

  const single = normalizeImportDate(raw);
  if (single) dates.add(single);
  return Array.from(dates).sort();
}

function parseImportWaktu(raw) {
  const dates = extractDatesFromIndonesianText(raw);
  if (!dates.length) return { mulai: '', selesai: '' };
  return {
    mulai: dates[0],
    selesai: dates[dates.length - 1],
  };
}

function renderSuratImportPreview() {
  const rows = suratImportState.rows || [];
  const payloads = suratImportState.payloads || [];
  const warnings = suratImportState.warnings || [];
  const body = document.getElementById('surat-import-preview-body');
  const submit = document.getElementById('btn-surat-import-submit');
  document.getElementById('surat-import-total').textContent = rows.length;
  document.getElementById('surat-import-ready').textContent = payloads.length;
  document.getElementById('surat-import-warn').textContent = warnings.length;
  document.getElementById('surat-import-file').textContent = suratImportState.fileName || '-';

  body.innerHTML = rows.map(row => {
    const hasWarn = row.warnings.length > 0;
    return `<tr>
      <td><span class="import-status-pill ${hasWarn ? 'warn' : 'ready'}">${hasWarn ? 'Cek' : 'Siap'}</span></td>
      <td>${esc(row.rowNum)}</td>
      <td>${esc(row.payload.nomor_surat || '-')}</td>
      <td>${esc(row.payload.tanggal_surat ? fmtTgl(row.payload.tanggal_surat) : '-')}</td>
      <td>${esc(row.waktuLabel || '-')}</td>
      <td>${esc(row.payload.perihal || '-')}</td>
      <td>${row.payload.pegawai_list.length ? esc(row.payload.pegawai_list.join(', ')) : '<span class="import-muted">Belum match</span>'}</td>
      <td>${esc(row.payload.pembebanan || '-')}</td>
      <td>${esc(row.payload.penandatangan_nama || '-')}</td>
      <td>${row.warnings.length ? `<ul class="import-warning-list">${row.warnings.map(w => `<li>${esc(w)}</li>`).join('')}</ul>` : '<span class="import-muted">-</span>'}</td>
    </tr>`;
  }).join('');

  if (submit) submit.disabled = payloads.length === 0;
}

function closeSuratImportModal() {
  suratImportState = { fileName: '', rows: [], payloads: [], warnings: [] };
  const modal = document.getElementById('modal-import-surat');
  if (modal) modal.style.display = 'none';
}

async function importSuratFromExcel(inputEl) {
  const file = inputEl.files && inputEl.files[0];
  inputEl.value = '';
  if (!file) return;

  try {
    showPageAlert('Memuat library Excel...', 'info');
    await ensureSheetJSLoaded();

    const buf = await file.arrayBuffer();
    const wb = XLSX.read(buf, { type: 'array', cellDates: false });
    if (!wb.SheetNames.length) throw new Error('File Excel tidak berisi sheet apapun.');

    const ws = wb.Sheets[wb.SheetNames[0]];
    const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, raw: false, defval: '' });
    if (aoa.length < 2) throw new Error('File Excel kosong atau hanya berisi header.');

    const headerRowIndex = findSuratImportHeaderRow(aoa);
    if (headerRowIndex < 0) throw new Error('Header import tidak ditemukan. Pastikan ada kolom Nomor Surat, Tujuan / Tugas, atau Nama.');

    const colIdxMap = buildSuratImportHeaderMap(aoa[headerRowIndex] || []);
    if (colIdxMap.perihal === undefined) {
      throw new Error('Kolom Perihal atau Tujuan / Tugas tidak ditemukan.');
    }

    const matcher = makePegawaiNameMatcher();
    const rows = [];
    const payloads = [];
    const warnings = [];

    aoa.slice(headerRowIndex + 1).forEach((row, idx) => {
      const rowNum = headerRowIndex + idx + 2;
      const cell = (field) => {
        const ci = colIdxMap[field];
        if (ci === undefined) return '';
        const value = row[ci];
        return value == null ? '' : String(value).trim();
      };

      const nomor = cell('nomor_surat');
      const perihal = cell('perihal');
      const tujuan = cell('tujuan');
      const pegawaiRaw = cell('pegawai_list');
      if (!nomor && !perihal && !tujuan && !pegawaiRaw) return;

      const rowWarnings = [];
      const peopleMatch = matchImportPeople(pegawaiRaw, matcher);
      if (pegawaiRaw && !peopleMatch.people.length) rowWarnings.push(`Pegawai belum match: ${peopleMatch.unmatched.join(', ')}`);

      const ttdRaw = cell('penandatangan_nama');
      const ttdMatch = matchImportPeople(ttdRaw, matcher);
      const ttd = ttdMatch.people[0] || null;
      if (ttdRaw && !ttd) rowWarnings.push(`Penandatangan belum match: ${ttdRaw}`);

      const waktuRaw = cell('waktu_pelaksanaan_text');
      const waktuParsed = parseImportWaktu(waktuRaw);
      const tanggalBerangkat = normalizeImportDate(cell('tanggal_berangkat')) || waktuParsed.mulai || null;
      const tanggalKembali = normalizeImportDate(cell('tanggal_kembali')) || waktuParsed.selesai || tanggalBerangkat || null;
      const tanggalSurat = normalizeImportDate(cell('tanggal_surat')) || null;
      if (!tanggalSurat) rowWarnings.push('Tanggal surat belum terbaca');
      if (waktuRaw && !tanggalBerangkat) rowWarnings.push('Waktu pelaksanaan belum terbaca sebagai tanggal');

      const jrRaw = cell('jumlah_responden');
      const jr = jrRaw ? parseInt(jrRaw, 10) : null;
      const bertugasRaw = cell('bertugas_sebagai');
      const bertugasArr = bertugasRaw ? bertugasRaw.split(',').map(s => s.trim()) : null;
      const matchedNames = peopleMatch.people.map(p => pegawaiNama(p));
      const matchedNips = peopleMatch.people.map(p => String(pegawaiNip(p)));
      const payload = {
        user_id: SESSION.id,
        status: 'menunggu',
        nomor_surat: nomor || null,
        tanggal_surat: tanggalSurat,
        tanggal_berangkat: tanggalBerangkat,
        tanggal_kembali: tanggalKembali,
        waktu_pelaksanaan_text: waktuRaw || null,
        perihal: perihal,
        tujuan: tujuan || null,
        pegawai_nip: matchedNips,
        pegawai_list: matchedNames,
        menimbang_custom: cell('menimbang_custom') || null,
        alat_angkutan: cell('alat_angkutan') || null,
        pembebanan: combineProgramMak(cell('_program'), cell('pembebanan')) || null,
        penandatangan_nip: ttd ? String(pegawaiNip(ttd)) : null,
        penandatangan_nama: ttd ? pegawaiNama(ttd) : null,
        tempat_terbit: cell('tempat_terbit') || 'Waisai',
        tipe: cell('tipe') || null,
        jumlah_responden: (isFinite(jr) && jr >= 0) ? jr : null,
        bertugas_sebagai: (bertugasArr && matchedNames.length >= 2 && bertugasArr.some(Boolean)) ? bertugasArr : null,
        created_at: new Date().toISOString(),
      };

      payloads.push(payload);
      warnings.push(...rowWarnings.map(w => `Baris ${rowNum}: ${w}`));
      rows.push({
        rowNum,
        payload,
        warnings: rowWarnings,
        waktuLabel: tanggalBerangkat ? fmtWaktu(tanggalBerangkat, tanggalKembali) : waktuRaw,
      });
    });

    if (!payloads.length) throw new Error('Tidak ada baris data valid yang bisa di-import.');

    suratImportState = { fileName: file.name, rows, payloads, warnings };
    renderSuratImportPreview();
    openModal('modal-import-surat');
    showPageAlert(`Preview import siap: ${payloads.length} baris dibaca.`, 'success');
  } catch (e) {
    console.error('[9201] preview import gagal:', e);
    showPageAlert(`Gagal membaca import: ${e.message}`, 'error');
  }
}

async function submitSuratImport() {
  const payloads = suratImportState.payloads || [];
  if (!payloads.length) {
    showPageAlert('Tidak ada data import yang siap diterapkan.', 'error');
    return;
  }
  const btn = document.getElementById('btn-surat-import-submit');
  const oldText = btn ? btn.textContent : '';
  if (btn) {
    btn.disabled = true;
    btn.textContent = 'Menyimpan...';
  }
  try {
    const res = await fetch(`${SUPABASE_URL}/rest/v1/surat_tugas`, {
      method: 'POST',
      headers: { ...H, 'Prefer': 'return=minimal' },
      body: JSON.stringify(payloads),
    });
    if (!res.ok) {
      let msg = `HTTP ${res.status}`;
      try { const j = await res.json(); msg = j.message || msg; } catch(_) {}
      throw new Error(msg);
    }
    const count = payloads.length;
    closeSuratImportModal();
    showPageAlert(`Berhasil import ${count} surat tugas dengan status menunggu.`, 'success');
    await loadSurat();
  } catch (e) {
    console.error('[9201] import gagal:', e);
    showPageAlert(`Gagal import: ${e.message}`, 'error');
    if (btn) {
      btn.disabled = false;
      btn.textContent = oldText || 'Terapkan Import';
    }
  }
}

/**
 * IMPORT: Read file Excel, parse, match pegawai by NAMA, INSERT ke DB
 * sebagai surat 'menunggu'. Trigger dari onchange event input file.
 */
async function importSuratFromExcelLegacy(inputEl) {
  const file = inputEl.files && inputEl.files[0];
  // Reset value supaya kalau user pilih file yang sama lagi, onchange tetap fire
  inputEl.value = '';
  if (!file) return;

  // Konfirmasi explicit — operasi ini akan INSERT banyak row ke DB
  if (!confirm(`File "${file.name}" akan di-import sebagai surat tugas baru dengan status "menunggu". Lanjut?`)) {
    return;
  }

  try {
    showPageAlert('⏳ Memuat library Excel...', 'info');
    await ensureSheetJSLoaded();

    showPageAlert('⏳ Membaca file...', 'info');
    const buf = await file.arrayBuffer();
    const wb  = XLSX.read(buf, { type: 'array', cellDates: false });
    if (!wb.SheetNames.length) {
      throw new Error('File Excel tidak berisi sheet apapun.');
    }
    const ws = wb.Sheets[wb.SheetNames[0]];
    // raw:false → date di-coerce ke string ISO; defval:'' → cell kosong jadi ''
    const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, raw: false, defval: '' });
    if (aoa.length < 2) {
      throw new Error('File Excel kosong atau hanya berisi header.');
    }

    // Parse header → mapping ke kolom DB. Header tidak case-sensitive,
    // di-trim, dan tidak harus punya semua kolom (cuma yang importable).
    const headerRow = aoa[0].map(h => String(h || '').trim());
    const headerToCol = {};  // headerLower → EXCEL_COLUMNS entry
    EXCEL_COLUMNS.forEach(c => {
      if (c.importable) headerToCol[c.header.toLowerCase()] = c;
    });
    const colIdxMap = {};  // field → column index di Excel
    headerRow.forEach((h, i) => {
      const norm = h.toLowerCase();
      const col = headerToCol[norm];
      if (col) colIdxMap[col.field] = i;
    });

    // Validasi: minimal "Perihal" harus ada (paling minimum required)
    if (colIdxMap.perihal === undefined) {
      throw new Error('Kolom "Perihal" tidak ditemukan di file Excel. Pastikan baris pertama berisi header.');
    }

    showPageAlert('⏳ Memvalidasi data...', 'info');

    // Build lookup pegawai by NAMA (case-insensitive, trimmed) — untuk match
    // kolom Pegawai dan Penandatangan.
    const pegawaiByNama = {};
    pegawaiList.forEach(p => {
      const k = String(pegawaiNama(p) || '').trim().toLowerCase();
      if (k) pegawaiByNama[k] = p;
    });

    const dataRows = aoa.slice(1);
    const payloads = [];
    const warnings = [];   // warning per row (pegawai tidak match, dll)

    dataRows.forEach((row, idx) => {
      const rowNum = idx + 2;  // baris di Excel (1-indexed, +1 untuk header)
      // Helper: ambil cell value by field name
      const cell = (field) => {
        const ci = colIdxMap[field];
        if (ci === undefined) return '';
        const v = row[ci];
        return v == null ? '' : String(v).trim();
      };

      const perihal = cell('perihal');
      // Skip baris kosong total (mis. trailing empty row)
      if (!perihal && !cell('tujuan') && !cell('pegawai_list') && !cell('nomor_surat')) {
        return;
      }

      // Match pegawai by NAMA — input format: "Nama1, Nama2, Nama3"
      const pegawaiRaw = cell('pegawai_list');
      const namaList = pegawaiRaw ? pegawaiRaw.split(',').map(n => n.trim()).filter(Boolean) : [];
      const matchedNips = [];
      const matchedNames = [];
      const unmatchedNames = [];
      namaList.forEach(nama => {
        const p = pegawaiByNama[nama.toLowerCase()];
        if (p) {
          matchedNips.push(String(pegawaiNip(p)));
          matchedNames.push(pegawaiNama(p));
        } else {
          unmatchedNames.push(nama);
        }
      });
      if (unmatchedNames.length) {
        warnings.push(`Baris ${rowNum}: pegawai tidak match → ${unmatchedNames.join(', ')}`);
      }

      // Match penandatangan by NAMA → ambil NIP-nya
      const ttdNamaRaw = cell('penandatangan_nama');
      let ttdNip = '', ttdNama = '';
      if (ttdNamaRaw) {
        const p = pegawaiByNama[ttdNamaRaw.toLowerCase()];
        if (p) {
          ttdNip = String(pegawaiNip(p));
          ttdNama = pegawaiNama(p);
        } else {
          warnings.push(`Baris ${rowNum}: penandatangan "${ttdNamaRaw}" tidak match ke data pegawai → di-skip`);
        }
      }

      // Bertugas Sebagai (joined string → array)
      const bertugasRaw = cell('bertugas_sebagai');
      const bertugasArr = bertugasRaw
        ? bertugasRaw.split(',').map(s => s.trim())
        : null;
      // Hanya disertakan kalau matched pegawai ≥ 2 (sama logic dengan UI)
      const bertugasFinal = (bertugasArr && matchedNames.length >= 2)
        ? (bertugasArr.some(v => v.length > 0) ? bertugasArr : null)
        : null;

      // Jumlah responden (number)
      const jrRaw = cell('jumlah_responden');
      const jr = jrRaw ? parseInt(jrRaw, 10) : null;

      // Build payload — status WAJIB 'menunggu' (sesuai permintaan user)
      const payload = {
        user_id:           SESSION.id,
        status:            'menunggu',
        nomor_surat:       cell('nomor_surat') || null,
        tanggal_surat:     normalizeImportDate(cell('tanggal_surat')) || null,
        tanggal_berangkat: normalizeImportDate(cell('tanggal_berangkat')) || null,
        tanggal_kembali:   normalizeImportDate(cell('tanggal_kembali')) || null,
        waktu_pelaksanaan_text: cell('waktu_pelaksanaan_text') || null,
        perihal:           perihal,
        tujuan:            cell('tujuan') || null,
        pegawai_nip:       matchedNips,
        pegawai_list:      matchedNames,
        menimbang_custom:  cell('menimbang_custom') || null,
        alat_angkutan:     cell('alat_angkutan') || null,
        pembebanan:        cell('pembebanan') || null,
        penandatangan_nip: ttdNip || null,
        penandatangan_nama: ttdNama || null,
        tempat_terbit:     cell('tempat_terbit') || 'Waisai',
        tipe:              cell('tipe') || null,
        jumlah_responden:  (isFinite(jr) && jr >= 0) ? jr : null,
        bertugas_sebagai:  bertugasFinal,
        created_at:        new Date().toISOString(),
      };
      payloads.push(payload);
    });

    if (!payloads.length) {
      throw new Error('Tidak ada baris data valid yang bisa di-import. Pastikan minimal kolom Perihal terisi.');
    }

    // Pre-flight summary — kalau ada warning, tampilkan ke user untuk konfirmasi
    if (warnings.length) {
      const preview = warnings.slice(0, 8).join('\n');
      const more = warnings.length > 8 ? `\n... +${warnings.length - 8} warning lain` : '';
      const ok = confirm(
        `⚠️ Ditemukan ${warnings.length} warning:\n\n${preview}${more}\n\n` +
        `${payloads.length} baris akan tetap di-import (data warning tetap masuk dengan field kosong).\n\n` +
        `Lanjut import?`
      );
      if (!ok) {
        showPageAlert('Import dibatalkan.', 'info');
        return;
      }
    }

    showPageAlert(`⏳ Mengirim ${payloads.length} baris ke server...`, 'info');

    const res = await fetch(`${SUPABASE_URL}/rest/v1/surat_tugas`, {
      method: 'POST',
      headers: { ...H, 'Prefer': 'return=minimal' },
      body: JSON.stringify(payloads)
    });
    if (!res.ok) {
      let msg = `HTTP ${res.status}`;
      try { const j = await res.json(); msg = j.message || msg; } catch(_) {}
      throw new Error(msg);
    }

    showPageAlert(`✅ Berhasil import ${payloads.length} surat (status menunggu).${warnings.length ? ` ${warnings.length} warning di console.` : ''}`, 'success');
    if (warnings.length) {
      console.warn('[9201 Import] Warnings:\n' + warnings.join('\n'));
    }
    // Reload data untuk tampilkan row baru
    await loadSurat();
  } catch (e) {
    console.error('[9201] import gagal:', e);
    showPageAlert(`Gagal import: ${e.message}`, 'error');
  }
}

/* Helper: normalisasi tanggal dari Excel ke ISO format (YYYY-MM-DD).
   Excel sering kirim format yang inkonsisten (date object, string lokal,
   dst). Kita coba beberapa parser & fallback ke parseFlexDate (existing
   helper di file ini yang juga handle "2/1/26", "2 jan 2026", dst). */
function normalizeImportDate(s) {
  if (!s) return '';
  s = String(s).trim();
  if (!s) return '';
  // Sudah ISO format (yyyy-mm-dd)?
  if (/^\d{4}-\d{2}-\d{2}/.test(s)) return s.slice(0, 10);
  // Excel serial number (kadang sheet_to_json raw:false tetap return number)
  // — di-skip karena raw:false biasanya sudah ke-format. Pakai parseFlexDate.
  if (typeof parseFlexDate === 'function') {
    const iso = parseFlexDate(s);
    if (iso) return iso;
  }
  return '';
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

  // #2 (refactor): jumlah_responden sekarang dari DB (persistent), tidak
  // perlu prompt modal lagi. Untuk tipe visum, baca dari surat.jumlah_responden.
  // Untuk non-visum, opts={}.
  const visumOpts = tipeHasVisum(surat.tipe)
    ? { jumlahResponden: (surat.jumlah_responden != null) ? surat.jumlah_responden : null }
    : {};

  currentPreviewSurat = surat;
  // Simpan opts visum agar tombol "Download .docx" dan "Buka di Word"
  // di modal preview ini bisa pakai jumlah responden yang sama tanpa
  // tanya ulang. _previewVisumOpts hanya hidup selama modal preview
  // terbuka — di-reset saat modal ditutup.
  _previewVisumOpts = visumOpts;
  _previewBlob = null;

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
      Dokumen yang sama dengan tombol download akan dibuka via Office Online.
    </div>
  </div>`;

  try {
    ensureLibrariesLoaded();
    const blob = await buildSuratTugasDoc(surat, visumOpts);
    _previewBlob = blob;
    const filename = await uploadPreviewDocx(blob, surat.id);
    _previewUploadedPath = filename;
    const signedUrl = await getPreviewSignedUrl(filename);
    const viewerUrl = buildPreviewViewerUrl(signedUrl);

    container.innerHTML = `
      <iframe src="${esc(viewerUrl)}"
        style="width:100%;height:80vh;min-height:600px;border:0;display:block;background:#fff"
        title="Preview Surat Tugas"
        allowfullscreen>
      </iframe>`;
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
      btn.textContent = '📥 Download';
      btn.disabled    = true;
    } else {
      btn.textContent = `📥 Download (${checked.length})`;
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

/* ════════════════════════════════════════════════════════════════════
   DRAFT SAVED TOAST
   Toast ringan yang muncul sebentar (1.5 detik) saat draft otomatis
   tersimpan ke localStorage. Memberi konfirmasi visual ke admin.
════════════════════════════════════════════════════════════════════ */
let _draftToastTimer = null;
function showDraftSavedToast() {
  const el = document.getElementById('draft-saved-toast');
  if (!el) return;
  el.classList.add('show');
  clearTimeout(_draftToastTimer);
  _draftToastTimer = setTimeout(() => el.classList.remove('show'), 1500);
}

/* ════════════════════════════════════════════════════════════════════
   BULK APPROVE — setujui banyak surat menunggu sekaligus
   ─────────────────────────────────────────────────────────────────────
   Flow:
     1. Admin centang checkbox di baris 'menunggu' (class bulk-approve-check)
     2. Tombol "Setujui Terpilih (N)" aktif
     3. openBulkApprove() — validasi semua, tampilkan modal konfirmasi
     4. Admin klik "Setujui Semua" → submitBulkApprove() proses sequential
     5. Setiap surat diapprove dengan nilai dari baris tabel (tanpa bertugas_sebagai
        dan jumlah_responden — bisa diedit setelah via Edit Role / approve individual)
════════════════════════════════════════════════════════════════════ */

/** Update label & state tombol bulk approve berdasarkan checkbox yang dipilih. */
function updateBulkApproveCounter() {
  const all     = document.querySelectorAll('.bulk-approve-check');
  const checked = document.querySelectorAll('.bulk-approve-check:checked');
  const btn     = document.getElementById('btn-bulk-approve');
  const master  = document.getElementById('bulk-ap-master');

  if (btn) {
    if (checked.length === 0) {
      btn.textContent = '✅ Setujui';
      btn.disabled    = true;
    } else {
      btn.textContent = `✅ Setujui (${checked.length})`;
      btn.disabled    = false;
    }
  }

  if (master) {
    if (all.length === 0 || checked.length === 0) {
      master.checked = false; master.indeterminate = false;
    } else if (checked.length === all.length) {
      master.checked = true;  master.indeterminate = false;
    } else {
      master.checked = false; master.indeterminate = true;
    }
  }
}

/** Toggle semua checkbox baris menunggu. */
function toggleBulkApproveAll(checked) {
  document.querySelectorAll('.bulk-approve-check').forEach(c => { c.checked = !!checked; });
  updateBulkApproveCounter();
}

/** Buka modal konfirmasi bulk approve — validasi semua surat terpilih terlebih dahulu. */
function openBulkApprove() {
  const checkeds = Array.from(document.querySelectorAll('.bulk-approve-check:checked'));
  const ids = checkeds.map(c => parseInt(c.dataset.suratId, 10)).filter(Number.isFinite);
  if (!ids.length) {
    showPageAlert('Pilih minimal 1 surat menunggu untuk disetujui.', 'error');
    return;
  }

  // Validasi semua — kumpulkan yang valid dan yang tidak
  const valid    = [];
  const invalid  = [];
  ids.forEach(id => {
    const values = collectRowFields(id);
    if (!values) { invalid.push({ id, reason: 'Data tidak ditemukan' }); return; }
    const { errors } = validateApproveFields(values);
    if (errors.length) {
      invalid.push({ id, reason: errors.join(', '), values });
    } else {
      valid.push({ id, values });
    }
  });

  // Render list di modal
  const listEl = document.getElementById('bulk-approve-list');
  if (listEl) {
    listEl.innerHTML = valid.map(({ id, values }) => {
      const s = suratMap[id];
      const nomorFull = buildNomorSuratFull(values.nomor_surat, values.tanggal_surat);
      const tipeLbl   = tipeLabel(values.tipe);
      const hasVisum  = tipeHasVisum(values.tipe);
      const hasLamp   = tipeHasLampiran(values.tipe) && values.pegawai_list.length >= 2;
      const warnings  = [];
      if (hasVisum) warnings.push('visum — isi Jumlah Responden setelah approve via individual');
      if (hasLamp)  warnings.push('lampiran — isi Bertugas Sebagai via Edit Role setelah approve');
      return `
        <div style="background:var(--bg);border:1px solid var(--border);border-radius:8px;padding:10px 14px;display:flex;align-items:flex-start;gap:10px">
          <span style="font-size:14px;margin-top:1px">✅</span>
          <div style="flex:1;min-width:0">
            <div style="font-size:12.5px;font-weight:600;color:var(--navy);white-space:nowrap;overflow:hidden;text-overflow:ellipsis">${esc(nomorFull)}</div>
            <div style="font-size:11.5px;color:var(--muted);margin-top:2px">${esc(values.perihal || '—')} · ${esc(tipeLbl)}</div>
            ${warnings.length ? `<div style="font-size:10.5px;color:#92400e;margin-top:3px;background:#fef3c7;padding:2px 7px;border-radius:5px;display:inline-block">⚠ ${esc(warnings.join(' · '))}</div>` : ''}
          </div>
        </div>`;
    }).join('') + (invalid.length ? `
      <div style="background:#fef2f2;border:1px solid #fecaca;border-radius:8px;padding:10px 14px;margin-top:6px">
        <div style="font-size:12px;font-weight:600;color:#991b1b;margin-bottom:6px">⚠ ${invalid.length} surat dilewati (tidak valid):</div>
        ${invalid.map(({ id, reason, values }) => {
          const v = values || {};
          return `<div style="font-size:11.5px;color:#991b1b;margin-bottom:3px">• ${esc(v.perihal || 'ID #' + id)}: ${esc(reason)}</div>`;
        }).join('')}
      </div>` : '');
  }

  document.getElementById('bulk-approve-count').textContent = valid.length;
  document.getElementById('bulk-approve-progress').style.display = 'none';

  const submitBtn = document.getElementById('btn-bulk-approve-submit');
  if (submitBtn) submitBtn.disabled = valid.length === 0;

  // Simpan IDs valid ke attribute modal supaya bisa diakses submitBulkApprove
  document.getElementById('modal-bulk-approve').dataset.validIds = JSON.stringify(valid.map(v => v.id));

  openModal('modal-bulk-approve');
}

/** Proses bulk approve — sequential, update progress bar. */
async function submitBulkApprove() {
  const modal = document.getElementById('modal-bulk-approve');
  let ids;
  try {
    ids = JSON.parse(modal.dataset.validIds || '[]');
  } catch(_) { ids = []; }
  if (!ids.length) return;

  const submitBtn = document.getElementById('btn-bulk-approve-submit');
  const cancelBtn = document.getElementById('btn-bulk-approve-cancel');
  const progressEl = document.getElementById('bulk-approve-progress');
  const progressText = document.getElementById('bulk-approve-progress-text');
  const progressBar  = document.getElementById('bulk-approve-bar');

  if (submitBtn) { submitBtn.disabled = true; }
  if (cancelBtn) { cancelBtn.disabled = true; }
  if (progressEl) progressEl.style.display = '';

  let success = 0;
  const failures = [];

  for (let i = 0; i < ids.length; i++) {
    const id = ids[i];
    const pct = Math.round(((i) / ids.length) * 100);
    if (progressBar)  progressBar.style.width  = pct + '%';
    if (progressText) progressText.textContent = `Memproses ${i + 1} / ${ids.length}…`;

    const values = collectRowFields(id);
    if (!values) { failures.push(`#${id}: data tidak ditemukan`); continue; }

    const jabatan = lookupJabatan(values.penandatangan_nip, values.tanggal_surat);
    const waktuText = values.waktu_pelaksanaan_text && values.waktu_pelaksanaan_text.trim()
      ? values.waktu_pelaksanaan_text : null;

    const payload = {
      status:                'selesai',
      nomor_surat:           values.nomor_surat,
      tanggal_surat:         values.tanggal_surat,
      tanggal_berangkat:     values.tanggal_berangkat,
      tanggal_kembali:       values.tanggal_kembali || null,
      waktu_pelaksanaan_text: waktuText,
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
      // bertugas_sebagai & jumlah_responden: sengaja NULL di bulk approve.
      // Admin bisa isi setelah via tombol Edit Role / approve individual.
      bertugas_sebagai:      null,
      jumlah_responden:      null,
    };

    try {
      const res = await fetch(`${SUPABASE_URL}/rest/v1/surat_tugas?id=eq.${id}`, {
        method: 'PATCH',
        headers: { ...H, 'Prefer': 'return=minimal' },
        body: JSON.stringify(payload),
      });
      if (!res.ok) {
        let msg = `HTTP ${res.status}`;
        try { const j = await res.json(); msg = j.message || msg; } catch(_) {}
        throw new Error(msg);
      }
      clearAdminDraft(id);
      // Auto-save defaults dari surat pertama yang berhasil
      if (success === 0) {
        const nd = {};
        if (values.alat_angkutan)     nd.alat_angkutan = values.alat_angkutan;
        if (values.pembebanan)        nd.pembebanan    = values.pembebanan;
        if (values.penandatangan_nip) nd.ttd_nip       = values.penandatangan_nip;
        if (values.penandatangan_nama)nd.ttd_nama      = values.penandatangan_nama;
        saveApproveDefaults({ ...loadApproveDefaults(), ...nd });
      }
      success++;
    } catch(e) {
      const s = suratMap[id];
      failures.push(`${(s && s.perihal) || '#' + id}: ${e.message}`);
    }
  }

  // Selesai
  if (progressBar)  progressBar.style.width  = '100%';
  if (progressText) progressText.textContent = `Selesai: ${success} berhasil${failures.length ? `, ${failures.length} gagal` : ''}.`;

  closeModal('modal-bulk-approve');
  if (failures.length === 0) {
    showPageAlert(`✅ ${success} surat berhasil disetujui.`, 'success');
  } else {
    showPageAlert(`⚠️ ${success} berhasil, ${failures.length} gagal: ${failures.slice(0,2).join('; ')}${failures.length > 2 ? '…' : ''}`, 'error');
    console.warn('[9201 Bulk Approve] failures:', failures);
  }

  // Unceklis semua approve checkbox
  document.querySelectorAll('.bulk-approve-check').forEach(c => { c.checked = false; });

  // Reload data
  const dirtySnapshot = captureMenungguDirty(null);
  await loadSurat();
  reapplyMenungguDirty(dirtySnapshot);
  loadMAKSuggestions();
  updateBulkApproveCounter();
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

  // #2 (refactor): jumlah_responden sekarang persistent di DB. Tidak
  // perlu prompt sequentially — langsung ambil dari surat.jumlah_responden.
  // Surat yang tidak punya value (NULL) → render pakai 7 baris default.
  const visumOptsById = {};
  for (let i = 0; i < ids.length; i++) {
    const id = ids[i];
    const surat = suratMap[id];
    if (!surat || surat.status !== 'selesai') continue;
    if (!tipeHasVisum(surat.tipe)) {
      visumOptsById[id] = {};
    } else {
      visumOptsById[id] = {
        jumlahResponden: (surat.jumlah_responden != null) ? surat.jumlah_responden : null
      };
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

    // (#2 refactor) — visumOptsById[id] tidak pernah null lagi setelah
    // jumlah_responden disimpan di DB. Tidak ada lagi skip path "dibatalkan
    // saat input visum" karena tidak ada prompt yang bisa di-cancel.

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
    const blob = _previewBlob || await buildSuratTugasDoc(currentPreviewSurat, opts);
    _previewBlob = blob;
    saveAs(blob, buildFileName(currentPreviewSurat));
    showPageAlert(`📥 Berhasil di-download: ${buildFileName(currentPreviewSurat)}`, 'success');
    closeModal('modal-preview');
  } catch(e) {
    console.error(e);
    showPageAlert(`Gagal download: ${e.message}`, 'error');
  }
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
      signedUrl = await getPreviewSignedUrl(_previewUploadedPath);
    } else {
      ensureLibrariesLoaded();
      const blob = await buildSuratTugasDoc(currentPreviewSurat, _previewVisumOpts || {});
      const filename = await uploadPreviewDocx(blob, currentPreviewSurat.id);
      _previewUploadedPath = filename;
      signedUrl = await getPreviewSignedUrl(filename);
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
                          {pangkat_p}, {golongan_p}, {pangkat_golongan_p},
                          {jabatan_p}, {bertugas_p}, {jabatan_bertugas_p}
       Catatan: {jabatan_bertugas_p} adalah field gabungan untuk kolom
       "Jabatan/Bertugas Sebagai" — sudah di-format jadi
       "{jabatan_p} / {bertugas_p}" kalau bertugas_p terisi, atau
       "{jabatan_p}" saja kalau kosong. Pakai {jabatan_bertugas_p}
       di template (single placeholder), bukan dua tag terpisah.
       Sama untuk {pangkat_golongan_p}: gabungan "pangkat / golongan"
       (atau "-" untuk mitra). Tag {pangkat_p} & {golongan_p} masih
       di-populate untuk backward-compat dengan template lama, tapi
       template baru cukup pakai {pangkat_golongan_p}.
     Header lampiran ikut pakai: {nomor_surat}, {tgl_surat}, {awalan},
                                 {menimbang},
                                 {jabatan_penandatangan}, {penandatangan},
                                 {nip_penandatangan}
*/
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
  // Lookup pegawai DAN mitra. Mitra punya NIP placeholder MITRA-{tahun}-{id};
  // untuk pegawai asli, lookup pegawaiByNIP[nip]. Lihat helper isMitraNip()
  // untuk deteksi.
  const peg              = isMitraNip(firstNip) ? null : pegawaiByNIP[firstNip];
  const mitr             = isMitraNip(firstNip) ? mitraByNip[firstNip]      : null;
  const namaPegawai      = lookupGelar(firstNip, data.tanggal_surat)
                             || pegawaiNama(peg)
                             || (mitr && mitr.nama)
                             || firstNm
                             || '';
  // Untuk mitra: jabatan = "Mitra" (dari record mitra), pangkat = "-"
  // Untuk pegawai: jabatan = lookup riwayat_jabatan, pangkat = '' (belum ada di skema)
  const jabatanPegawai   = mitr
                             ? (mitr.jabatan || 'Mitra')
                             : (lookupJabatan(firstNip, data.tanggal_surat) || '');
  // Pangkat: pegawai = '' (belum di skema), mitra = '-' (memang tidak ada)
  const pangkatPegawai   = mitr ? '-' : '';
  // Satuan kerja: pegawai dari kolom UNIT KERJA, mitra dari kolom instansi
  const skerjaPegawai    = mitr
                             ? (mitr.instansi || '')
                             : pegawaiUnitKerja(peg);

  // ── Halaman 1: nama, jabatan, pangkat → "Terlampir" kalau ≥2 pegawai
  // Sesuai konfirmasi user: untuk pegawai >1 nama, jabatan, pangkat &
  // golongan, serta jabatan/instansi memang menjadi "Terlampir".
  // (Termasuk di SPD — sengaja, supaya konsisten.)
  // ─────────────────────────────────────────────────────────────────────
  const banyakPegawai = nameList.length >= 2;
  const namaHal1      = banyakPegawai ? 'Terlampir' : namaPegawai;
  const jabatanHal1   = banyakPegawai ? 'Terlampir' : jabatanPegawai;
  const pangkatHal1   = banyakPegawai ? 'Terlampir' : pangkatPegawai;

  // Array role tambahan per pegawai (paralel dengan nipList & nameList).
  // Format DB: TEXT[] — element kosong/NULL artinya pegawai tsb tanpa role
  // tambahan. Hanya relevan untuk tipe lampiran (T2/T3V); tipe non-lampiran
  // tidak akan punya field ini di DB → array kosong → semua bertugas_p = ''.
  const bertugasList = Array.isArray(data.bertugas_sebagai) ? data.bertugas_sebagai : [];

  // Helper kecil — bangun row pegawai standar yang dipakai di lampiran
  // maupun di array kendaraan/menginap. DRY: kedua array isi pegawai-nya
  // identik, hanya beda nama variable. Sebelumnya ada duplikasi 100% di
  // dua .map() berbeda — sekarang konsolidasi via helper ini.
  //
  // Aturan format kolom "Jabatan/Bertugas Sebagai":
  //   - bertugas_p kosong  → tampilkan {jabatan_p} saja
  //   - bertugas_p terisi  → tampilkan "{jabatan_p} / {bertugas_p}"
  // Logic ini di-precompute jadi field jabatan_bertugas_p supaya template
  // tinggal pakai 1 placeholder saja (tidak perlu {#bertugas_p}{/bertugas_p}
  // conditional di template, lebih simple di-maintain).
  function buildPegawaiRow(nm, i) {
    const nip  = String(nipList[i] || '').trim();
    // Branch: pegawai biasa vs mitra. Lookup ke pegawaiByNIP atau mitraByNip
    // sesuai prefix NIP. Lihat isMitraNip() helper.
    const isMitra = isMitraNip(nip);
    const p    = isMitra ? null : pegawaiByNIP[nip];
    const m    = isMitra ? mitraByNip[nip] : null;

    // Field-field dasar (nama, jabatan, pangkat, golongan, skerja, nip_p)
    // di-resolve sesuai sumber data:
    //   - Pegawai biasa: lookup pegawaiByNIP + riwayat_jabatan + riwayat_pangkat_golongan
    //   - Mitra: dari record mitra (jabatan='Mitra', pangkat='-', nip='-')
    const nama_p     = lookupGelar(nip, data.tanggal_surat)
                       || pegawaiNama(p) || (m && m.nama) || nm || '';
    const nip_p      = isMitra ? '-' : (nip || '-');
    const jabatan_p  = isMitra
                         ? (m ? (m.jabatan || 'Mitra') : 'Mitra')
                         : (lookupJabatan(nip, data.tanggal_surat) || '');
    // Pangkat & golongan — lookup dari riwayat_pangkat_golongan (TMT-aware).
    // Mitra: keduanya "-" (mitra tidak punya pangkat resmi).
    const pgRecord = isMitra ? null : lookupPangkatGolongan(nip, data.tanggal_surat);
    const pangkat_p  = isMitra ? '-' : (pgRecord ? (pgRecord.pangkat  || '') : '');
    const golongan_p = isMitra ? '-' : (pgRecord ? (pgRecord.golongan || '') : '');
    // Field gabungan untuk template (rule: "pangkat / golongan" untuk pegawai,
    // "-" untuk mitra). Ini yang dipakai di template (single placeholder).
    let pangkat_golongan_p;
    if (isMitra) {
      pangkat_golongan_p = '-';
    } else if (pangkat_p && golongan_p) {
      pangkat_golongan_p = `${pangkat_p} / ${golongan_p}`;
    } else if (pangkat_p) {
      pangkat_golongan_p = pangkat_p;
    } else if (golongan_p) {
      pangkat_golongan_p = golongan_p;
    } else {
      pangkat_golongan_p = '';
    }

    const skerja_p   = isMitra
                         ? (m ? (m.instansi || '') : '')
                         : pegawaiUnitKerja(p);

    const bertugas_p = String(bertugasList[i] || '').trim();
    const jabatan_bertugas_p = bertugas_p
      ? `${jabatan_p} / ${bertugas_p}`
      : jabatan_p;

    return {
      no:                 String(i + 1),
      nama_p,
      nip_p,
      pangkat_p,
      golongan_p,
      pangkat_golongan_p,
      jabatan_p,
      bertugas_p,
      jabatan_bertugas_p,
      skerja_p,
    };
  }

  // Bangun array pegawai untuk loop tabel di halaman lampiran (T2/T3V).
  // Setiap item: { no, nama_p, nip_p, pangkat_p, golongan_p,
  //                pangkat_golongan_p, jabatan_p, bertugas_p,
  //                jabatan_bertugas_p, skerja_p }
  // pangkat_p, golongan_p, pangkat_golongan_p di-lookup dari
  // riwayat_pangkat_golongan (TMT-aware). Mitra → semua "-".
  const pegawaiLampiran = nameList.map(buildPegawaiRow);

  // Array kendaraan & menginap — sama isinya per pegawai. Saat ini cuma
  // ada satu "iterasi" per pegawai (tabel di template di-loop sekali per
  // orang). Field belum diisi karena belum ada kolom DB-nya.
  const pegawaiBaris = nameList.map(buildPegawaiRow);

  // ── Penandatangan ────────────────────────────────────────────────────
  // Rule transform jabatan_penandatangan (lihat transformJabatanPenandatangan):
  //   Input dari kolom penandatangan_jabatan adalah salah satu dari:
  //     - "Kepala BPS Kabupaten Raja Ampat"
  //     - "Plt. Kepala Badan Pusat Statistik"
  //     - lainnya
  //   Output di template:
  //     - "Kepala BPS Kabupaten Raja Ampat"   (apa adanya)
  //     - "Plt. Kepala BPS Kabupaten Raja Ampat"
  //     - "Plh. Kepala BPS Kabupaten Raja Ampat"
  // Fallback: kalau penandatangan_jabatan kosong, lookup via lookupJabatan
  // (jenis='utama') untuk konsistensi dengan behavior lama.
  const ttdNip        = data.penandatangan_nip  || '';
  const ttdNama       = lookupGelar(ttdNip, data.tanggal_surat)
                     || data.penandatangan_nama
                     || '';
  const ttdJabRaw     = data.penandatangan_jabatan
                     || lookupJabatan(ttdNip, data.tanggal_surat)
                     || '';
  const ttdJabatan    = transformJabatanPenandatangan(ttdJabRaw);

  // ── PPK (Pejabat Pembuat Komitmen) ───────────────────────────────────
  // Lookup dari riwayat_jabatan jenis='lainnya' untuk row dengan
  // jabatan='Pejabat Pembuat Komitmen' dan tmt <= tanggal_surat.
  // Ambil yang TMT-nya paling baru.
  const ppkRecord = lookupJabatanLainnya('Pejabat Pembuat Komitmen', data.tanggal_surat);
  const nipPPK    = ppkRecord ? String(ppkRecord.pegawai_nip || '').trim() : '';
  // Resolve nama PPK: pertama dari data_pegawai (NAMA), fallback ke
  // riwayat_jabatan.nama (denormalized), lalu fallback ke default.
  let namaPPK = '';
  if (nipPPK) {
    const pegPPK = pegawaiByNIP[nipPPK];
    namaPPK = lookupGelar(nipPPK, data.tanggal_surat)
            || pegawaiNama(pegPPK)
            || (ppkRecord && ppkRecord.nama)
            || '';
  }
  // Fallback terakhir: kalau lookup tidak return (data belum di-seed),
  // pakai PPK_NAMA_DEFAULT supaya surat tetap render dengan placeholder
  // yang masuk akal — bukan string kosong.
  if (!namaPPK) namaPPK = PPK_NAMA_DEFAULT;

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

  // ── Waktu pelaksanaan ─────────────────────────────────────────────────
  // Prioritas: kolom waktu_pelaksanaan_text (admin override untuk kasus
  // multi-range, mis. "25 s.d. 27 Maret & 30 Maret s.d. 1 April 2026").
  // Kalau kosong, fallback ke format auto dari range tanggal.
  const waktuPelaksanaan = (data.waktu_pelaksanaan_text && data.waktu_pelaksanaan_text.trim())
                        || fmtWaktu(data.tanggal_berangkat, data.tanggal_kembali)
                        || '';

  // ── Awalan untuk header tabel lampiran ───────────────────────────────
  // "Daftar Peserta" untuk surat pelatihan, "Daftar Petugas" untuk lainnya.
  const awalan = computeAwalan(data.perihal);

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
    // {untuk_text} — field gabungan untuk baris "Untuk" di surat tugas:
    //   - Kalau tempat_tujuan terisi → pakai tempat_tujuan
    //   - Kalau tempat_tujuan kosong → "{perihal} pada tanggal {waktu_pelaksanaan}"
    // Pakai tag {untuk_text} di template (single placeholder), bukan
    // docxtemplater conditional yang sulit dengan format multi-paragraf.
    untuk_text:            (data.tujuan && data.tujuan.trim())
                             ? data.tujuan
                             : `${data.perihal || ''} pada tanggal ${waktuPelaksanaan}`,
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
    // {nip} — tetap NIP asli (dipakai di beberapa tempat lain di template).
    // Untuk SPD baris "Nama/NIP", pakai tag TUNGGAL {nama_nip} di template
    // supaya hasilnya "Terlampir" (bukan "Terlampir / Terlampir").
    // → Di template Word: ganti "{nama} / {nip}" menjadi "{nama_nip}"
    nip:                   firstNip,
    // {nama_nip} — tag gabungan untuk sel "Nama / NIP" di SPD.
    //   1 pegawai  : "Nama Lengkap / NIP123"
    //   ≥ 2 pegawai: "Terlampir"
    nama_nip:              banyakPegawai
                             ? 'Terlampir'
                             : (namaPegawai && firstNip
                                  ? `${namaPegawai} / ${firstNip}`
                                  : (namaPegawai || firstNip || '')),
    // {skerja_p} halaman 1: unit kerja pegawai / instansi mitra.
    // "Terlampir" jika ≥2 pegawai (konsisten dengan {nama}/{jabatan}).
    skerja_p:              banyakPegawai ? 'Terlampir' : skerjaPegawai,
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
    //   "Daftar Peserta" untuk surat pelatihan, "Daftar Petugas" lainnya.
    // ─────────────────────────────────────────────────────────────────
    has_lampiran_st:       flags.has_lampiran,
    awalan:                awalan,
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
async function init() {
  SESSION = novaCheckSession({ requireAdmin: true });
  if (!SESSION) return;
  SESSION = await novaVerifyAdminSession(SESSION);
  if (!SESSION) return;
  Topbar9201.setUser(SESSION);
  initRoleSwitcher(SESSION, true);
  Promise.all([loadPegawai(), loadMitra(), loadRiwayatJabatan(), loadRiwayatPangkatGolongan(), loadRiwayatGelar(), loadUsers(), loadSurat()]);

  // Panaskan library saja. Template .docx tidak diambil saat halaman dibuka
  // supaya download manager tidak menangkap request storage sebagai unduhan.
  ensureDocxtemplaterLoaded().catch(() => {});

  // Load history POK (MAK Pembebanan) untuk autocomplete dropdown.
  // Dijalankan setelah loadSurat selesai biar tidak ada race kalau RLS lambat.
  loadMAKSuggestions();

}
init();
