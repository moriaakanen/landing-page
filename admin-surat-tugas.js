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
const APPROVE_DEFAULTS_KEY = 'nova_approve_defaults_v2';
const FACTORY_DEFAULTS = {
  alat_angkutan:  'Kendaraan Darat',
  pembebanan:     'DIPA BPS Kabupaten Raja Ampat Tahun Anggaran 2026',
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
  const m = { menunggu:['menunggu','⏳ Menunggu'], disetujui:['disetujui','✅ Disetujui'], ditolak:['ditolak','❌ Ditolak'] };
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

function updateStats() {
  document.getElementById('st-total').textContent    = allSurat.length;
  document.getElementById('st-menunggu').textContent = allSurat.filter(s => s.status === 'menunggu').length;
  document.getElementById('st-disetujui').textContent= allSurat.filter(s => s.status === 'disetujui').length;
  document.getElementById('st-ditolak').textContent  = allSurat.filter(s => s.status === 'ditolak').length;
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
        // Urutan logis: menunggu → disetujui → ditolak
        const order = { menunggu: 0, disetujui: 1, ditolak: 2 };
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
function renderTable(data) {
  document.getElementById('table-count').textContent = `${data.length} surat`;
  if (!data.length) {
    document.getElementById('table-area').innerHTML = `<div class="table-empty"><div class="table-empty-icon">📭</div><div class="table-empty-text">Tidak ada surat tugas ditemukan.</div></div>`;
    return;
  }

  const todayStr = todayISO();

  const rows = data.map(s => {
    const isMenunggu  = s.status === 'menunggu';
    const isDisetujui = s.status === 'disetujui';
    const isDitolak   = s.status === 'ditolak';

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
        <button class="btn-approve" onclick="openApprove(${s.id})">✅ Setujui</button>
        <button class="btn-reject" onclick="openReject(${s.id})">❌ Tolak</button>`;
    } else if (isDisetujui) {
      aksi = `
        <button class="btn-preview" onclick="openPreview(${s.id})">👁 Preview</button>
        <button class="btn-download" onclick="downloadSuratTugas(${s.id})">📥</button>`;
    } else {
      // Status ditolak — tanpa aksi
      aksi = `<span style="font-size:11px;color:var(--muted);font-style:italic">—</span>`;
    }

    return `
      <tr data-surat-id="${s.id}" data-status="${s.status}">
        <td class="col-no">${urutNo}</td>

        ${cellTextHTML(s.id, 'nomor_surat', nomorSurat, isMenunggu, 'cth: 001 / 013A')}
        ${cellDateHTML(s.id, 'tanggal_surat', tanggalSurat, isMenunggu, 'tgl/bln/thn')}
        ${cellDateRangeHTML(s.id, 'waktu', waktuMulai, waktuSelesai, isMenunggu)}
        ${cellTextareaHTML(s.id, 'perihal', perihal, isMenunggu, 'Perihal surat')}
        ${cellTextHTML(s.id, 'tujuan', tujuan, isMenunggu, 'Kota/instansi')}
        ${cellPegawaiMultiHTML(s.id, pegNips, pegNames, isMenunggu)}
        ${cellTextareaHTML(s.id, 'menimbang_custom', menimbang, isMenunggu, 'cth: pelaksanaan Survei...')}
        ${cellTextareaHTML(s.id, 'alat_angkutan', alat, isMenunggu, 'cth: Kendaraan Darat')}
        ${cellTextareaHTML(s.id, 'pembebanan', mak, isMenunggu, 'cth: DIPA BPS...')}
        ${cellPenandatanganHTML(s.id, ttdNip, ttdNama, isMenunggu)}

        <td class="col-status">${badgeHTML(s.status)}</td>
        <td class="col-aksi"><div class="aksi-wrap">${aksi}</div></td>
        <td class="col-pengaju" title="${esc(getPengajuNama(s))}">${esc(getPengajuNama(s)) || '<span style="color:var(--muted);font-style:italic">—</span>'}</td>
      </tr>`;
  }).join('');

  document.getElementById('table-area').innerHTML = `
    <table class="list-table">
      <thead><tr>
        ${sortHeader('no',            'No',                'col-no')}
        ${sortHeader('nomor_surat',   'Nomor Surat',       'col-nomor-surat')}
        ${sortHeader('tanggal_surat', 'Tgl Surat',         'col-tgl-surat')}
        ${sortHeader('waktu',         'Waktu Pelaksanaan', 'col-waktu')}
        ${sortHeader('perihal',       'Perihal',           'col-perihal')}
        ${sortHeader('tujuan',        'Tempat Tujuan',     'col-tujuan')}
        <th class="col-nama">Nama Pegawai</th>
        <th class="col-menimbang">Menimbang</th>
        <th class="col-alat">Alat Angkutan</th>
        <th class="col-mak">MAK Pembebanan</th>
        <th class="col-ttd">Penandatangan</th>
        ${sortHeader('status',        'Status',            'col-status')}
        <th class="col-aksi">Aksi</th>
        ${sortHeader('pengaju',       'Diajukan oleh',     'col-pengaju')}
      </tr></thead>
      <tbody>${rows}</tbody>
    </table>`;

  requestAnimationFrame(() => {
    document.querySelectorAll('tr[data-status="menunggu"] textarea.xls-cell').forEach(ta => {
      autoGrow(ta);
      ta.addEventListener('input', () => { autoGrow(ta); ta.classList.remove('err'); });
    });
  });
  document.querySelectorAll('tr[data-status="menunggu"] input.xls-cell').forEach(inp => {
    inp.addEventListener('input', () => inp.classList.remove('err'));
  });
}

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
    const ds = d.dataset.date;
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
  };
  document.getElementById('cal-year-grid').onclick = (e) => {
    const c = e.target.closest('.cal-year-cell');
    if (!c) return;
    calState.year = parseInt(c.dataset.year, 10);
    calState.yearMode = false;
    renderCal();
  };
});

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

  /* ── Escape ─────────────────────────────────────────────────────── */
  if (e.key === 'Escape') {
    closeAllPopups();
    ['modal-approve','modal-reject','modal-preview'].forEach(closeModal);
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

  /* ── Tab / Enter: navigasi linear (hanya untuk sel editable) ────── */
  if (isReadonly) return; // readonly: biarkan Tab default browser

  if (isTextarea && e.key === 'Enter' && e.shiftKey) return; // Shift+Enter = newline

  if (e.key === 'Tab') {
    e.preventDefault();
    moveCellFocus(target, e.shiftKey ? -1 : 1);
  } else if (e.key === 'Enter' && !e.shiftKey) {
    e.preventDefault();
    moveCellFocus(target, 1);
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
    ['pembebanan',         'MAK Pembebanan'],
    ['penandatangan_nama', 'Penandatangan'],
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
function closeModal(id) { const el = document.getElementById(id); if (el) el.style.display = 'none'; }

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
    <div class="approve-preview-row"><strong>MAK Pembebanan</strong><span>${esc(values.pembebanan)}</span></div>
    <div class="approve-preview-row"><strong>Penandatangan</strong><span>
      ${esc(values.penandatangan_nama)}<br>
      <em style="color:var(--muted);font-style:italic;font-size:11px">${ttdJabatan ? esc(ttdJabatan) : '<span style="color:var(--red)">⚠ Jabatan tidak ditemukan di riwayat (akan ditampilkan "-" di docx)</span>'}</em><br>
      <span style="color:var(--muted);font-size:11px">NIP. ${esc(values.penandatangan_nip)}</span>
    </span></div>
  `;
  document.getElementById('inp-catatan-approve').value = s.catatan_admin || '';
  document.getElementById('approve-alert').className = 'alert';
  openModal('modal-approve');
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
    status: 'disetujui',
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
    showPageAlert('✅ Surat tugas berhasil disetujui.', 'success');
    await loadSurat();
  } catch(e) {
    document.getElementById('approve-alert-icon').textContent = '⚠️';
    document.getElementById('approve-alert-text').textContent = `Gagal: ${e.message}`;
    document.getElementById('approve-alert').className = 'alert error show';
  } finally {
    btn.disabled = false; btn.classList.remove('loading');
  }
}

/* ════════════════════════════════════════════════════════════════════
   REJECT
═══════════════════════════════════════════════════════════════════════ */
function openReject(id) {
  selectedId = id;
  const s = suratMap[id]; if (!s) return;
  document.getElementById('reject-perihal').textContent = s.perihal || '—';
  document.getElementById('inp-catatan-reject').value = '';
  document.getElementById('err-catatan-reject').classList.remove('show');
  document.getElementById('reject-alert').className = 'alert';
  openModal('modal-reject');
}
async function submitReject() {
  const catatan = document.getElementById('inp-catatan-reject').value.trim();
  if (!catatan) { document.getElementById('err-catatan-reject').classList.add('show'); return; }
  document.getElementById('err-catatan-reject').classList.remove('show');
  const btn = document.getElementById('btn-reject-submit');
  btn.disabled = true; btn.classList.add('loading');
  try {
    const res = await fetch(`${SUPABASE_URL}/rest/v1/surat_tugas?id=eq.${selectedId}`, {
      method: 'PATCH',
      headers: { ...H, 'Prefer': 'return=minimal' },
      body: JSON.stringify({ status: 'ditolak', catatan_admin: catatan })
    });
    if (!res.ok) throw new Error(`HTTP ${res.status}`);
    closeModal('modal-reject');
    showPageAlert('❌ Surat tugas telah ditolak.', 'success');
    await loadSurat();
  } catch(e) {
    document.getElementById('reject-alert-icon').textContent = '⚠️';
    document.getElementById('reject-alert-text').textContent = `Gagal: ${e.message}`;
    document.getElementById('reject-alert').className = 'alert error show';
  } finally {
    btn.disabled = false; btn.classList.remove('loading');
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
  if (surat.status !== 'disetujui') {
    showPageAlert('Surat hanya bisa di-preview jika sudah disetujui.', 'error');
    return;
  }
  currentPreviewSurat = surat;
  openModal('modal-preview');

  const container = document.getElementById('preview-container');
  container.innerHTML = `<div class="preview-loading"><div class="preview-loading-spin"></div><div>Memuat dokumen...</div></div>`;

  try {
    ensureLibrariesLoaded();
    const blob = await buildSuratTugasDoc(surat);
    container.innerHTML = '';
    await window.docxPreview.renderAsync(blob, container, null, {
      className: 'docx', inWrapper: true,
      ignoreWidth: false, ignoreHeight: false, ignoreFonts: false,
      breakPages: true, ignoreLastRenderedPageBreak: true,
      experimental: true, trimXmlDeclaration: true, useBase64URL: false,
      renderHeaders: true, renderFooters: true,
      renderFootnotes: true, renderEndnotes: true,
    });
  } catch(e) {
    console.error(e);
    container.innerHTML = `<div class="preview-error"><div class="preview-error-icon">⚠️</div>
      <div><strong>Gagal memuat preview</strong></div>
      <div style="font-size:12px;margin-top:8px;color:var(--muted)">${esc(e.message)}</div></div>`;
  }
}

async function downloadSuratTugas(suratId) {
  const surat = suratMap[suratId];
  if (!surat) return;
  if (surat.status !== 'disetujui') {
    showPageAlert('Surat hanya bisa di-download jika sudah disetujui.', 'error'); return;
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
  if (!currentPreviewSurat) return;
  const container = document.getElementById('preview-container');
  if (!container || !container.querySelector('.docx')) {
    showPageAlert('Preview belum selesai dimuat. Tunggu sebentar lalu coba lagi.', 'error');
    return;
  }
  window.print();
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
   TEMPLATE-BASED BUILDER — docxtemplater
   
   KONFIGURASI URL TEMPLATE:
   Ubah TEMPLATE_URL di bawah sesuai lokasi file template-surat-tugas.docx.
   
   Pilihan umum:
   - Same-origin (taruh di folder yang sama dengan HTML):
       'template-surat-tugas.docx'
   - Supabase Storage (public bucket) — GANTI <bucket> dan path-nya:
       'https://jsmmtqeoukkgugorrvmg.supabase.co/storage/v1/object/public/templates/template-surat-tugas.docx'
   - CDN eksternal atau URL absolut lainnya:
       'https://example.com/path/template.docx'
════════════════════════════════════════════════════════════════════ */

const TEMPLATE_URL = 'https://jsmmtqeoukkgugorrvmg.supabase.co/storage/v1/object/public/templates/template-surat-tugas.docx';

// Cache template binary supaya tidak di-fetch berulang kali
let _templateBuffer = null;

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

async function loadTemplateBuffer() {
  if (_templateBuffer) return _templateBuffer;
  console.log('[NOVA] Memuat template dari:', TEMPLATE_URL);
  const res = await fetch(TEMPLATE_URL);
  if (!res.ok) throw new Error(`Gagal memuat template (${res.status}). Pastikan ${TEMPLATE_URL} bisa diakses publik.`);
  _templateBuffer = await res.arrayBuffer();
  console.log('[NOVA] Template loaded:', _templateBuffer.byteLength, 'bytes');
  return _templateBuffer;
}

/* Helper format tanggal ISO → "DD Nama-Bulan YYYY" (Indonesia) */
function fmtTglId(isoStr) {
  if (!isoStr) return '';
  const d = parseISODate(isoStr);
  if (!d) return '';
  return `${d.getDate()} ${BULAN[d.getMonth()]} ${d.getFullYear()}`;
}

/* Bangun objek mapping data → placeholder template.
   Satu-satunya tempat di mana field dari DB di-mapping ke nama placeholder. */
function buildTemplateData(data) {
  // Ambil info pegawai pertama (template saat ini hanya support 1 pegawai per surat)
  const nipList  = Array.isArray(data.pegawai_nip)  ? data.pegawai_nip  : [];
  const nameList = Array.isArray(data.pegawai_list) ? data.pegawai_list : [];
  const firstNip = nipList[0] || '';
  const firstNm  = nameList[0] || '';
  const pegInfo  = getPegawaiInfoForDoc(firstNip, firstNm, data.tanggal_surat);

  const tempatBerangkat = data.tempat_berangkat || 'Kabupaten Raja Ampat';
  const tempatTujuan    = data.tujuan            || '';

  const nomorFull = buildNomorSuratFull(data.nomor_surat, data.tanggal_surat);

  return {
    // ── Header ────────────────────────────────────────────────────
    nomor:             nomorFull || '',
    lembar:            data.lembar || '',

    // ── PPK ───────────────────────────────────────────────────────
    ppk_nama:          data.penandatangan_nama || '',
    ppk_nip:           data.penandatangan_nip  || '',

    // ── Data pegawai perjalanan ──────────────────────────────────
    pegawai_nama:      pegInfo.nama || firstNm || '-',
    pangkat_gol:       pegInfo.pangkat_gol || pegInfo.pangkat || '-',
    jabatan:           pegInfo.jabatan || '-',

    // ── Maksud / alat / lokasi ────────────────────────────────────
    maksud:            data.perihal || data.maksud || '',
    alat_angkutan:     data.alat_angkutan || 'Kendaraan Umum',
    tempat_berangkat:  tempatBerangkat,
    tempat_tujuan:     tempatTujuan,

    // ── Durasi ────────────────────────────────────────────────────
    lama_hari:         data.lama_hari || '',
    tgl_berangkat:     fmtTglId(data.tanggal_berangkat),
    tgl_kembali:       fmtTglId(data.tanggal_kembali),

    // ── Pembebanan anggaran ──────────────────────────────────────
    program_kode:      data.program_kode    || 'GG',
    program_desc:      data.program_desc    || '',
    kegiatan_kode:     data.kegiatan_kode   || '',
    kegiatan_desc:     data.kegiatan_desc   || '',
    komponen_kode:     data.komponen_kode   || '',
    komponen_desc:     data.komponen_desc   || '',
    instansi_anggaran: data.instansi_anggaran || 'Badan Pusat Statistik Kabupaten Raja Ampat',
    mata_anggaran:     data.mata_anggaran   || '',

    // ── Footer halaman 2 (Dikeluarkan) ───────────────────────────
    dikeluarkan_di:    data.tempat_terbit || 'Waisai',
    tgl_dikeluarkan:   fmtTglId(data.tanggal_surat),

    // ── Halaman 3: Kepala BPS ────────────────────────────────────
    kepala_nama:       data.kepala_nama || '',
    kepala_nip:        data.kepala_nip  || '',

    // ── Halaman 3: Tanggal-tanggal perjalanan ────────────────────
    tiba_tujuan_tgl:     fmtTglId(data.tiba_tujuan_tgl     || data.tanggal_berangkat),
    berangkat_balik_tgl: fmtTglId(data.berangkat_balik_tgl || data.tanggal_kembali),
    tiba_kembali_tgl:    fmtTglId(data.tiba_kembali_tgl    || data.tanggal_kembali),
  };
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

  const buf = await loadTemplateBuffer();
  const zip = new PizZipCtor(buf);
  const doc = new DocxtemplaterCtor(zip, {
    paragraphLoop: true,
    linebreaks: true,
    delimiters: { start: '{', end: '}' },
  });

  const templateData = buildTemplateData(data);
  try {
    doc.render(templateData);
  } catch (err) {
    // Docxtemplater error biasanya informatif — munculkan ke console
    console.error('Template render error:', err, err.properties);
    throw new Error(`Template error: ${err.message}`);
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
}
init();
