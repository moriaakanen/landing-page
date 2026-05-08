/**
 * 9201 PAK DATA — Penilaian Angka Kredit
 * ─────────────────────────────────────────────────────────────────────
 * Single source of truth untuk semua perhitungan AK di Portal 9201.
 * Diload SETELAH config.js (tidak butuh dependency lain).
 *
 * Berisi:
 *   - PREDIKAT_KINERJA       — 5 tingkat predikat + persentase koefisien
 *   - JENJANG_FUNGSIONAL     — 8 jenjang (Pemula → Ahli Utama)
 *   - GOLONGAN_PNS           — 13 golongan (IIa → IVe)
 *   - Helper lookup          — getPredikat(), getJenjangByName(), dll.
 *   - Helper progresi        — getNextJenjang(), progressGolongan(), dll.
 *   - Calculation functions:
 *       calcAK_annual()        — Case 1: setahun penuh, predikat tahunan
 *       calcAK_periode()       — Case 2: pro-rata bulanan untuk 1 periode
 *       calcAK_tahun()         — Wrapper: auto-pilih Case 1 vs 2 per tahun
 *       buildPeriodsFromTMTs() — Pecah tahun jadi periode berdasarkan TMT
 *
 * Sumber regulasi: PermenPAN-RB 1/2023 dan turunannya. Kalau regulasi
 * berubah (mis. multiplier predikat direvisi), edit konstanta di file
 * ini saja — semua halaman otomatis ikut.
 *
 * Catatan: nilai-nilai di bawah disesuaikan dengan tabel Kriteria.xlsx
 * yang disediakan unit pengelola. JANGAN HARDCODE angka ini di file
 * lain — selalu import lewat konstanta di sini.
 *
 * REV 2026-05 (a) — Sinkronisasi ke Kriteria.xlsx terbaru:
 *   - JENJANG_FUNGSIONAL: ak_awal & ak_naik di-update sesuai xlsx.
 *   - GOLONGAN_PNS: ak_awal & ak_naik di-update sesuai xlsx.
 *
 * REV 2026-05 (b) — Restructure {pangkat_min} & {jenjang_min}:
 *   Sebelumnya: kebutuhan_naik (delta) di JENJANG_FUNGSIONAL & GOLONGAN_PNS.
 *   Sekarang  : pangkat_min & jenjang_min di GOLONGAN_PNS saja, karena
 *               delta untuk naik jenjang tergantung POSISI golongan dalam
 *               jenjang (mis. IIa→Pelaksana butuh 30 AK, IIb→Pelaksana
 *               cuma 15 AK karena naik pangkat IIb→IIc otomatis naik
 *               jenjang). Generator pengajuan PAK pakai field baru ini.
 *               Field kebutuhan_naik dihapus.
 */

// ─── PREDIKAT KINERJA ─────────────────────────────────────────────────
// Multiplier untuk hitung AK per tahun: AK = koefisien × persentase / 100
// Order dari terbaik ke terburuk (rank 5 = terbaik).
window.PREDIKAT_KINERJA = [
  { key: 'sangat_baik',     label: 'Sangat Baik',     persentase: 150, rank: 5, color: '#1a7a4a' },
  { key: 'baik',            label: 'Baik',            persentase: 100, rank: 4, color: '#0d2340' },
  { key: 'butuh_perbaikan', label: 'Butuh Perbaikan', persentase: 75,  rank: 3, color: '#92400e' },
  { key: 'kurang',          label: 'Kurang',          persentase: 50,  rank: 2, color: '#c2410c' },
  { key: 'sangat_kurang',   label: 'Sangat Kurang',   persentase: 25,  rank: 1, color: '#c0392b' },
];

window.PREDIKAT_BY_KEY = window.PREDIKAT_KINERJA.reduce((acc, p) => {
  acc[p.key] = p;
  return acc;
}, {});

/**
 * Normalisasi string predikat ke key kanonis. Toleran terhadap
 * kapitalisasi/spasi (mis. "SANGAT BAIK", "Sangat Baik", "sangat baik",
 * "sangatbaik" → 'sangat_baik'). Return null kalau tidak dikenal.
 *
 * Krusial untuk import data dari Excel — kolom predikat di sumber
 * berbeda kapitalisasi tergantung siapa yang mengetik.
 */
function normalizePredikatKey(s) {
  if (s == null) return null;
  const lower = String(s).toLowerCase().trim();
  if (!lower) return null;
  // Direct key match
  if (window.PREDIKAT_BY_KEY[lower]) return lower;
  // Match dengan/tanpa underscore
  const compact = lower.replace(/[\s_]+/g, '');
  for (const p of window.PREDIKAT_KINERJA) {
    if (p.key.replace(/_/g, '') === compact) return p.key;
    if (p.label.toLowerCase().replace(/\s+/g, '') === compact) return p.key;
  }
  return null;
}
window.normalizePredikatKey = normalizePredikatKey;

/** @returns {object|null} predikat object atau null */
function getPredikat(key_or_label) {
  const k = normalizePredikatKey(key_or_label);
  return k ? window.PREDIKAT_BY_KEY[k] : null;
}
window.getPredikat = getPredikat;

// ─── JENJANG FUNGSIONAL ───────────────────────────────────────────────
// Sumber: Kriteria.xlsx — sheet "Kriteria", section "Angka Kredit
// Jabatan Fungsional".
//
// Field per jenjang:
//   koefisien   = pengali AK per tahun (untuk hitung {ak} di pengajuan PAK)
//   kategori    = 'keterampilan' atau 'keahlian' (untuk getNextJenjang)
//   order       = urutan dalam kategori (untuk getNextJenjang)
//   is_puncak   = true kalau jenjang tertinggi dalam kategori-nya
//                 (Penyelia di keterampilan, Ahli Utama di keahlian)
//   min_pendidikan = persyaratan pendidikan minimum
//
// CATATAN: Threshold AK pegawai (untuk pengajuan PAK & card profil)
// di-track di GOLONGAN_PNS lewat field pangkat_min & jenjang_min,
// BUKAN di sini. Karena threshold tergantung posisi golongan, bukan
// jenjang.
window.JENJANG_FUNGSIONAL = [
  // Kategori Keterampilan
  { key: 'pemula',       nama: 'Pemula',       kategori: 'keterampilan',
    koefisien: 3.75, is_puncak: false, min_pendidikan: 'SMA',     order: 1 },
  { key: 'pelaksana',    nama: 'Pelaksana',    kategori: 'keterampilan',
    koefisien: 5,    is_puncak: false, min_pendidikan: 'SMA',     order: 2 },
  { key: 'mahir',        nama: 'Mahir',        kategori: 'keterampilan',
    koefisien: 12.5, is_puncak: false, min_pendidikan: 'SMA',     order: 3 },
  { key: 'penyelia',     nama: 'Penyelia',     kategori: 'keterampilan',
    koefisien: 25,   is_puncak: true /* puncak Keterampilan, harus pindah ke Keahlian */,
    min_pendidikan: 'S1/DIV', order: 4 },
  // Kategori Keahlian
  { key: 'ahli_pertama', nama: 'Ahli Pertama', kategori: 'keahlian',
    koefisien: 12.5, is_puncak: false, min_pendidikan: 'S1/DIV', order: 1 },
  { key: 'ahli_muda',    nama: 'Ahli Muda',    kategori: 'keahlian',
    koefisien: 25,   is_puncak: false, min_pendidikan: 'S1/DIV', order: 2 },
  { key: 'ahli_madya',   nama: 'Ahli Madya',   kategori: 'keahlian',
    koefisien: 37.5, is_puncak: false, min_pendidikan: 'S2',     order: 3 },
  { key: 'ahli_utama',   nama: 'Ahli Utama',   kategori: 'keahlian',
    koefisien: 50,   is_puncak: true /* MAX */,
    min_pendidikan: 'S2', order: 4 },
];

window.JENJANG_BY_KEY = window.JENJANG_FUNGSIONAL.reduce((acc, j) => {
  acc[j.key] = j;
  return acc;
}, {});

/**
 * Cari jenjang berdasarkan nama (case-insensitive, toleran spasi
 * & tanda baca). Mis. "ahli pertama", "Ahli Pertama", "ahli_pertama"
 * semua return jenjang Ahli Pertama.
 */
function getJenjangByName(name) {
  if (!name) return null;
  const norm = String(name).toLowerCase().trim().replace(/[\s_-]+/g, '_');
  if (window.JENJANG_BY_KEY[norm]) return window.JENJANG_BY_KEY[norm];
  // Fallback: match nama field
  const lower = String(name).toLowerCase().trim();
  return window.JENJANG_FUNGSIONAL.find(j => j.nama.toLowerCase() === lower) || null;
}
window.getJenjangByName = getJenjangByName;

/**
 * Jenjang berikutnya dalam kategori yang sama. Return null kalau
 * sudah di puncak (Penyelia → null karena harus pindah kategori,
 * Ahli Utama → null karena MAX).
 */
function getNextJenjang(jenjang_aktif) {
  if (!jenjang_aktif || jenjang_aktif.is_puncak) return null;
  return window.JENJANG_FUNGSIONAL.find(j =>
    j.kategori === jenjang_aktif.kategori &&
    j.order === jenjang_aktif.order + 1
  ) || null;
}
window.getNextJenjang = getNextJenjang;

// ─── GOLONGAN PNS ─────────────────────────────────────────────────────
// Sumber: Kriteria.xlsx — sheet "Kriteria", section "Angka Kredit
// Golongan".
//
// Field per golongan:
//   pangkat_min       = AK threshold untuk naik pangkat dari golongan ini.
//                       NULL kalau pangkat puncak (IVe).
//                       Dipakai oleh tag {pangkat_min} di template surat
//                       dan untuk kalkulasi kekurangan AK di card profil.
//   jenjang_min       = AK threshold untuk naik jenjang dari golongan ini.
//                       Sama untuk semua golongan dalam satu jenjang
//                       (mis. IVa, IVb, IVc semuanya 450 untuk Ahli Madya).
//                       NULL kalau jenjang puncak (Ahli Utama: IVd, IVe).
//                       Dipakai oleh tag {jenjang_min} di template surat.
//
// Rumus kekurangan AK (konsisten di seluruh aplikasi):
//   sisa_pangkat = ak_terbaru − pangkat_min
//   sisa_jenjang = ak_terbaru − jenjang_min
//
//   Positif (>0)  = AK SUDAH CUKUP/LEBIH untuk eligible naik
//   Nol/Negatif   = AK MASIH KURANG sebesar |nilai|
//
// Catatan logika {keputusan} — kalau pangkat_min === jenjang_min →
// "naik pangkat sekaligus naik jenjang". Hanya pada golongan terakhir
// di jenjang (IIb, IId, IIIb, IIId, IVc).
window.GOLONGAN_PNS = [
  // Pemula
  { kode:1,  golongan:'IIa',  pangkat:'Pengatur Muda',           pangkat_min:15,  jenjang_min:30,  min_pendidikan:'SMA' },
  { kode:2,  golongan:'IIb',  pangkat:'Pengatur Muda Tk. I',     pangkat_min:30,  jenjang_min:30,  min_pendidikan:'SMA' },
  // Pelaksana
  { kode:3,  golongan:'IIc',  pangkat:'Pengatur',                pangkat_min:20,  jenjang_min:40,  min_pendidikan:'SMA' },
  { kode:4,  golongan:'IId',  pangkat:'Pengatur Tk. I',          pangkat_min:40,  jenjang_min:40,  min_pendidikan:'SMA' },
  // Mahir / Ahli Pertama
  { kode:5,  golongan:'IIIa', pangkat:'Penata Muda',             pangkat_min:50,  jenjang_min:100, min_pendidikan:'S1/DIV' },
  { kode:6,  golongan:'IIIb', pangkat:'Penata Muda Tk. I',       pangkat_min:100, jenjang_min:100, min_pendidikan:'S1/DIV' },
  // Penyelia / Ahli Muda
  { kode:7,  golongan:'IIIc', pangkat:'Penata',                  pangkat_min:100, jenjang_min:200, min_pendidikan:'S1/DIV' },
  { kode:8,  golongan:'IIId', pangkat:'Penata Tk. I',            pangkat_min:200, jenjang_min:200, min_pendidikan:'S2' },
  // Ahli Madya (3 golongan: IVa, IVb, IVc — semua jenjang_min 450 ke Ahli Utama)
  { kode:9,  golongan:'IVa',  pangkat:'Pembina',                 pangkat_min:150, jenjang_min:450, min_pendidikan:'S2' },
  { kode:10, golongan:'IVb',  pangkat:'Pembina Tk. I',           pangkat_min:300, jenjang_min:450, min_pendidikan:'S2' },
  { kode:11, golongan:'IVc',  pangkat:'Pembina Utama Muda',      pangkat_min:450, jenjang_min:450, min_pendidikan:'S2' },
  // Ahli Utama (puncak jenjang — jenjang_min null)
  { kode:12, golongan:'IVd',  pangkat:'Pembina Utama Madya',     pangkat_min:200, jenjang_min:null, min_pendidikan:'S2' },
  // IVe = puncak total (semua null)
  { kode:13, golongan:'IVe',  pangkat:'Pembina Utama',           pangkat_min:null, jenjang_min:null, min_pendidikan:'S2' },
];

window.GOLONGAN_BY_KODE = window.GOLONGAN_PNS.reduce((acc, g) => {
  acc[g.kode] = g;
  return acc;
}, {});
window.GOLONGAN_BY_NAMA = window.GOLONGAN_PNS.reduce((acc, g) => {
  acc[g.golongan.toLowerCase()] = g;
  return acc;
}, {});

/**
 * Cari golongan berdasarkan nama Romawi (mis. 'IIIa', 'iiia', 'III/a').
 * Toleran terhadap "/" pemisah dan kapitalisasi.
 */
function getGolonganByName(name) {
  if (!name) return null;
  // Hapus spasi & "/" — "III/a" → "iiia", "III a" → "iiia"
  const norm = String(name).toLowerCase().replace(/[\s/]+/g, '').trim();
  return window.GOLONGAN_BY_NAMA[norm] || null;
}
window.getGolonganByName = getGolonganByName;

function getGolonganByKode(kode) {
  return window.GOLONGAN_BY_KODE[Number(kode)] || null;
}
window.getGolonganByKode = getGolonganByKode;

/** Golongan berikutnya. Return null kalau sudah IVe (puncak). */
function getNextGolongan(gol_aktif) {
  // Puncak = pangkat_min null (artinya tidak ada threshold naik pangkat
  // = sudah di IVe).
  if (!gol_aktif || gol_aktif.pangkat_min === null) return null;
  return window.GOLONGAN_BY_KODE[gol_aktif.kode + 1] || null;
}
window.getNextGolongan = getNextGolongan;

/**
 * Ekstrak jenjang dari teks `jabatan` (mis. "Statistisi Ahli Pertama" →
 * jenjang Ahli Pertama). Cek dari yang paling spesifik dulu agar
 * "Ahli Pertama" tidak ke-claim oleh "Ahli" generic.
 *
 * Return jenjang object dari JENJANG_FUNGSIONAL atau null kalau tidak
 * cocok dengan jenjang manapun.
 *
 * Dipakai oleh halaman Profil Saya (index.html) dan Riwayat Kepegawaian
 * Admin (admin-riwayat.html) — pindah ke sini agar tidak duplikasi.
 */
function extractJenjangFromJabatan(jabatanText) {
  if (!jabatanText || typeof window.JENJANG_BY_KEY === 'undefined') return null;
  const t = String(jabatanText).toLowerCase();
  // Order penting: yang LEBIH SPESIFIK didahulukan (Ahli Pertama sebelum
  // pattern "ahli" lain agar match akurat).
  const matches = [
    ['ahli_pertama', /\bahli\s+pertama\b/],
    ['ahli_madya',   /\bahli\s+madya\b/],
    ['ahli_utama',   /\bahli\s+utama\b/],
    ['ahli_muda',    /\bahli\s+muda\b/],
    ['penyelia',     /\bpenyelia\b/],
    ['mahir',        /\bmahir\b/],
    ['pelaksana',    /\bpelaksana\b/],
    ['pemula',       /\bpemula\b/],
  ];
  for (const [key, re] of matches) {
    if (re.test(t)) return window.JENJANG_BY_KEY[key];
  }
  return null;
}
window.extractJenjangFromJabatan = extractJenjangFromJabatan;

/**
 * Ekstrak NAMA JABATAN tanpa jenjang dari teks lengkap.
 * Mis. "Statistisi Ahli Pertama" → "Statistisi"
 *      "Pranata Komputer Pelaksana" → "Pranata Komputer"
 *      "Statistisi Ahli Madya" → "Statistisi"
 *
 * Strip suffix jenjang (case-insensitive, dengan whitespace toleran).
 * Return string asli kalau tidak ada match.
 *
 * Dipakai oleh generator pengajuan PAK ({keputusan}).
 */
function extractNamaJabatanTanpaJenjang(jabatanText) {
  if (!jabatanText) return '';
  const t = String(jabatanText).trim();
  // Order: yang lebih spesifik dulu agar "Ahli Pertama" di-strip sebelum
  // pattern "Ahli" yang lebih pendek kena match duluan.
  const patterns = [
    /\s+ahli\s+pertama\s*$/i,
    /\s+ahli\s+madya\s*$/i,
    /\s+ahli\s+utama\s*$/i,
    /\s+ahli\s+muda\s*$/i,
    /\s+penyelia\s*$/i,
    /\s+mahir\s*$/i,
    /\s+pelaksana\s*$/i,
    /\s+pemula\s*$/i,
  ];
  for (const re of patterns) {
    if (re.test(t)) return t.replace(re, '').trim();
  }
  return t;
}
window.extractNamaJabatanTanpaJenjang = extractNamaJabatanTanpaJenjang;

// ─── PROGRESSI / GAP ANALYSIS ─────────────────────────────────────────

/**
 * Hitung progress AK pegawai untuk naik PANGKAT dan naik JENJANG.
 *
 * Rumus simple (sesuai aturan unit pengelola):
 *   selisih_p = ak_kumulatif − pangkat_min
 *   selisih_j = ak_kumulatif − jenjang_min
 *
 *   Positif (≥0) = AK SUDAH CUKUP/LEBIH untuk eligible naik (can_promote=true)
 *   Negatif      = AK MASIH KURANG sebesar |selisih| (can_promote=false)
 *
 * Catatan: AK didapat per pengajuan dibatasi koefisien × 150% (Sangat
 * Baik), jadi pegawai TIDAK akan melompati threshold sampai jauh dalam
 * 1 pengajuan. Skenario realistis: pegawai dekat dengan threshold —
 * kurang sedikit, tepat, atau lebih sedikit. Kasus extreme (AK jauh
 * di bawah pangkat_min) hanya muncul kalau data riwayat AK tidak
 * lengkap — tetap di-handle dengan benar (menunjukkan kekurangan besar).
 *
 * Konsisten dengan rumus pengajuan PAK ({kurang_lebih_p} & {kurang_lebih_j}
 * di template surat).
 *
 * @param {number} ak_kumulatif - total AK pegawai sekarang (dari
 *                                 riwayat_angka_kredit kolom angka_kredit
 *                                 untuk tmt terbaru)
 * @param {object} gol_aktif - dari GOLONGAN_PNS
 * @returns {object} {
 *   pangkat: { sisa, can_promote, target_min, blocker } | null kalau puncak
 *   jenjang: { sisa, can_promote, target_min, blocker } | null kalau puncak
 * }
 *
 * Field per sub-object:
 *   sisa         = |selisih| kalau AK kurang. 0 kalau sudah cukup/lebih.
 *                  Selalu ≥ 0 (display absolut).
 *   can_promote  = boolean — apakah AK ≥ threshold (selisih ≥ 0)
 *   target_min   = pangkat_min atau jenjang_min (untuk display)
 *   blocker      = pesan informatif (string atau null)
 */
function progressGolongan(ak_kumulatif, gol_aktif) {
  if (!gol_aktif) {
    return { error: 'Golongan tidak diketahui' };
  }

  const ak = Number(ak_kumulatif) || 0;

  // Sub-object untuk PANGKAT
  let pangkat;
  if (gol_aktif.pangkat_min == null) {
    // Pangkat puncak (IVe)
    pangkat = null;
  } else {
    const selisih = ak - gol_aktif.pangkat_min;
    const cukup   = (selisih >= 0);
    pangkat = {
      sisa: cukup ? 0 : r3(-selisih),  // display absolut
      can_promote: cukup,
      target_min: gol_aktif.pangkat_min,
      blocker: cukup
        ? 'AK sudah cukup. Pastikan masa kerja minimal 4 tahun di golongan '
          + 'sekarang dan pendidikan memenuhi syarat.'
        : null,
    };
  }

  // Sub-object untuk JENJANG
  let jenjang;
  if (gol_aktif.jenjang_min == null) {
    // Jenjang puncak (Ahli Utama: IVd, IVe)
    jenjang = null;
  } else {
    const selisih = ak - gol_aktif.jenjang_min;
    const cukup   = (selisih >= 0);
    jenjang = {
      sisa: cukup ? 0 : r3(-selisih),
      can_promote: cukup,
      target_min: gol_aktif.jenjang_min,
      blocker: cukup
        ? 'AK sudah cukup. Pastikan sudah lulus UKOM dan pendidikan '
          + 'memenuhi syarat jenjang berikutnya.'
        : null,
    };
  }

  return { pangkat, jenjang };
}
window.progressGolongan = progressGolongan;

/** Internal: round ke 3 desimal. Hindari floating-point noise. */
function r3(n) {
  return Number(Number(n).toFixed(3));
}

// ─── KALKULASI AK ──────────────────────────────────────────────────────

/**
 * Case 1 — AK setahun penuh tanpa perubahan jenjang/pangkat di tahun T.
 * Pakai predikat tahunan langsung × koefisien jenjang.
 *
 * @param {number} koefisien     - koefisien jenjang aktif
 * @param {string|object} predikat - 'Baik', 'Sangat Baik', dst, atau object predikat
 * @returns {number} AK yang didapat untuk tahun itu
 *
 * @throws Kalau predikat tidak dikenal.
 */
function calcAK_annual(koefisien, predikat) {
  const p = (typeof predikat === 'string') ? getPredikat(predikat) : predikat;
  if (!p) throw new Error(`Predikat tidak dikenal: ${predikat}`);
  if (typeof koefisien !== 'number' || koefisien <= 0) {
    throw new Error(`Koefisien tidak valid: ${koefisien}`);
  }
  return Number((koefisien * (p.persentase / 100)).toFixed(3));
}
window.calcAK_annual = calcAK_annual;

/**
 * Case 2 — AK pro-rata bulanan untuk SATU periode (mis. Jan–Apr 2026
 * sebelum naik pangkat). Kontribusi bulanan = (koef/12) × pct_predikat.
 *
 * @param {number} koefisien - koefisien jenjang aktif di periode ini
 * @param {Array<{bulan, predikat}>} bulanan
 *        bulan = 1-12; predikat = string atau object
 * @returns {number} AK total periode tersebut
 *
 * Bulan tanpa predikat (null/undefined) di-skip dengan warning ke
 * console — ini supaya pegawai dengan data tidak lengkap tetap bisa
 * dihitung partial, bukan crash.
 */
function calcAK_periode(koefisien, bulanan) {
  if (!Array.isArray(bulanan) || !bulanan.length) return 0;
  if (typeof koefisien !== 'number' || koefisien <= 0) {
    throw new Error(`Koefisien tidak valid: ${koefisien}`);
  }
  let total = 0;
  for (const b of bulanan) {
    const p = (typeof b.predikat === 'string') ? getPredikat(b.predikat) : b.predikat;
    if (!p) {
      console.warn(`[calcAK_periode] Predikat bulan ${b.bulan} tidak dikenal/kosong, skipped`);
      continue;
    }
    total += (koefisien / 12) * (p.persentase / 100);
  }
  return Number(total.toFixed(3));
}
window.calcAK_periode = calcAK_periode;

/**
 * Wrapper tingkat tinggi: hitung AK untuk SATU TAHUN, otomatis pilih
 * Case 1 (annual) atau Case 2 (split) berdasarkan jumlah periode.
 *
 * @param {object} args
 * @param {number} args.tahun
 * @param {Array<{bulan_start, bulan_end, koefisien, label?}>} args.periods
 *        Periode jenjang aktif tahun itu, urut waktu. Untuk tahun tanpa
 *        promosi → 1 elemen {bulan_start:1, bulan_end:12, koefisien}.
 *        Untuk naik pangkat April → 2 elemen [{1,4,k1}, {5,12,k2}].
 * @param {string|null} [args.predikat_tahunan]
 *        Predikat tahunan (untuk Case 1). Kalau null/undefined,
 *        fallback ke proporsional bulanan.
 * @param {object} [args.predikat_bulanan]
 *        Map { 1: 'Baik', 2: 'Sangat Baik', ... 12: ... }. Wajib
 *        kalau periods.length > 1, opsional kalau cuma 1 periode
 *        (tapi recommended sebagai fallback).
 *
 * @returns {object} { total, mode, periods, warnings }
 *   mode: 'annual' (Case 1) | 'split' (Case 2) | 'partial' (1 periode tapi <12 bulan)
 */
function calcAK_tahun({ tahun, periods, predikat_tahunan, predikat_bulanan }) {
  if (!Array.isArray(periods) || !periods.length) {
    throw new Error('calcAK_tahun: minimal 1 periode required');
  }

  const warnings = [];

  // Case 1: 1 periode penuh setahun (Jan-Des)
  if (periods.length === 1) {
    const p = periods[0];
    const isFullYear = (p.bulan_start === 1 && p.bulan_end === 12);

    if (isFullYear && predikat_tahunan) {
      const ak = calcAK_annual(p.koefisien, predikat_tahunan);
      return {
        total: ak,
        mode: 'annual',
        periods: [{ ...p, ak, predikat_dipakai: predikat_tahunan }],
        warnings,
      };
    }

    // Partial year (CPNS baru, pensiun, atau predikat tahunan tidak ada)
    if (isFullYear && !predikat_tahunan) {
      warnings.push(`Tahun ${tahun}: predikat tahunan tidak tersedia, fallback ke perhitungan bulanan.`);
    }
    if (!predikat_bulanan) {
      throw new Error(`calcAK_tahun: predikat_bulanan wajib ada untuk perhitungan partial/fallback (tahun ${tahun})`);
    }
    const bulanan = monthsToBulanan(p.bulan_start, p.bulan_end, predikat_bulanan, warnings, tahun);
    const ak = calcAK_periode(p.koefisien, bulanan);
    return {
      total: ak,
      mode: isFullYear ? 'annual_fallback' : 'partial',
      periods: [{ ...p, ak, bulanan }],
      warnings,
    };
  }

  // Case 2: split — wajib pakai predikat bulanan
  if (!predikat_bulanan) {
    throw new Error(`calcAK_tahun: predikat_bulanan wajib ada untuk split (tahun ${tahun})`);
  }
  let total = 0;
  const detailedPeriods = periods.map((p) => {
    const bulanan = monthsToBulanan(p.bulan_start, p.bulan_end, predikat_bulanan, warnings, tahun);
    const ak = calcAK_periode(p.koefisien, bulanan);
    total += ak;
    return { ...p, ak, bulanan };
  });
  return {
    total: Number(total.toFixed(3)),
    mode: 'split',
    periods: detailedPeriods,
    warnings,
  };
}
window.calcAK_tahun = calcAK_tahun;

/** Internal helper: bangun array bulanan dari range + map predikat. */
function monthsToBulanan(start, end, predikat_bulanan, warnings, tahun) {
  const arr = [];
  for (let m = start; m <= end; m++) {
    const p = predikat_bulanan[m];
    if (!p && warnings) {
      warnings.push(`Tahun ${tahun} bulan ${m}: predikat tidak tersedia, kontribusi 0.`);
    }
    arr.push({ bulan: m, predikat: p || null });
  }
  return arr;
}

/**
 * Berdasarkan tanggal-tanggal TMT promosi di tahun T (urut ascending),
 * bangun list periode dengan bulan_start & bulan_end.
 *
 * Aturan (dikonfirmasi user): bulan TMT masuk ke periode SEBELUM
 * promosi. Periode baru mulai bulan TMT+1.
 *   TMT 15 April → Jan-Apr di periode pertama, Mei-Des di periode kedua.
 *
 * @param {Array<string>} tmt_dates - ISO 'YYYY-MM-DD', urut ascending
 *        DI dalam tahun T. Boleh empty array → return 1 periode penuh.
 * @param {Array<number>} koefisien_per_period
 *        Length harus = tmt_dates.length + 1.
 *        Index 0 = koefisien sebelum TMT pertama, dst.
 * @returns {Array<{bulan_start, bulan_end, koefisien}>}
 *
 * @example
 *   buildPeriodsFromTMTs(['2026-04-15'], [12.5, 25])
 *   // → [{bulan_start:1, bulan_end:4, koefisien:12.5},
 *   //    {bulan_start:5, bulan_end:12, koefisien:25}]
 */
function buildPeriodsFromTMTs(tmt_dates, koefisien_per_period) {
  if (!Array.isArray(tmt_dates)) tmt_dates = [];
  if (!Array.isArray(koefisien_per_period)) {
    throw new Error('koefisien_per_period harus array');
  }
  if (koefisien_per_period.length !== tmt_dates.length + 1) {
    throw new Error(
      `koefisien_per_period harus length = tmt_dates.length + 1 `
      + `(got tmt=${tmt_dates.length}, koef=${koefisien_per_period.length})`
    );
  }
  const periods = [];
  let cursor = 1;
  for (let i = 0; i < tmt_dates.length; i++) {
    const m = parseInt(String(tmt_dates[i]).slice(5, 7), 10);
    if (m < cursor || m > 12) {
      throw new Error(`TMT ke-${i + 1} bulan ${m} tidak valid (cursor di ${cursor})`);
    }
    periods.push({ bulan_start: cursor, bulan_end: m, koefisien: koefisien_per_period[i] });
    cursor = m + 1;
  }
  // Periode terakhir
  if (cursor <= 12) {
    periods.push({
      bulan_start: cursor,
      bulan_end: 12,
      koefisien: koefisien_per_period[koefisien_per_period.length - 1],
    });
  }
  return periods;
}
window.buildPeriodsFromTMTs = buildPeriodsFromTMTs;

// ─── HELPERS UNTUK UI ─────────────────────────────────────────────────

/**
 * Format AK dengan 3 desimal, tapi buang trailing zero.
 * Konsisten dengan fmtAK() di admin-riwayat.html.
 */
function fmtAKNumber(v) {
  if (v == null || v === '') return '—';
  const n = Number(v);
  if (isNaN(n)) return String(v);
  return Number.isInteger(n) ? String(n) : n.toFixed(3).replace(/\.?0+$/, '');
}
window.fmtAKNumber = fmtAKNumber;

/**
 * Format AK untuk OUTPUT DOKUMEN (locale Indonesia: koma sebagai
 * pemisah desimal, titik untuk ribuan). Mis. 1234.567 → "1.234,567".
 *
 * Bedanya dengan fmtAKNumber: yang ini selalu menampilkan minimal 1
 * digit desimal kalau ada koma, dan pakai locale ID. fmtAKNumber pakai
 * dot untuk konteks input/UI English.
 *
 * Dipakai oleh 9201-pak-generator.js untuk semua tag {ak}, {ak_n1},
 * {ak_tot}, {akb}, dll. di template surat.
 */
function fmtAKNumberID(v) {
  if (v == null || v === '') return '';
  const n = Number(v);
  if (isNaN(n)) return String(v);
  if (Number.isInteger(n)) return n.toLocaleString('id-ID');
  // Buang trailing zero, tapi keep decimal separator sebagai koma.
  let s = n.toFixed(3).replace(/\.?0+$/, '');
  // Ubah dot decimal jadi koma, tambah titik ribuan di integer part.
  const [intPart, fracPart] = s.split('.');
  const intFormatted = parseInt(intPart, 10).toLocaleString('id-ID');
  return fracPart ? `${intFormatted},${fracPart}` : intFormatted;
}
window.fmtAKNumberID = fmtAKNumberID;

/**
 * Cek apakah pendidikan pegawai memenuhi minimum jenjang/golongan.
 *
 * Toleran terhadap berbagai cara penulisan: "D-IV", "DIV", "D.IV", "D IV"
 * semua dianggap sama (rank 3 = setara S1). Untuk pendidikan compound
 * mis. "S1/DIV", split berdasarkan "/" lalu ambil rank tertinggi.
 *
 * @param {string} pendidikan_pegawai - mis. 'S1', 'S2', 'D-IV', 'SMA'
 * @param {string} pendidikan_minimum - dari jenjang/golongan, mis. 'S1/DIV', 'S2'
 * @returns {boolean} true kalau memenuhi
 */
function pendidikanMencukupi(pendidikan_pegawai, pendidikan_minimum) {
  if (!pendidikan_minimum) return true;  // tidak ada syarat
  if (!pendidikan_pegawai) return false;

  // Hierarki sederhana — angka makin besar makin tinggi.
  // Setiap diploma punya alias Roman numeral juga (DI=D1, DII=D2, dst).
  const TINGKAT = {
    'sd':0, 'smp':0, 'sltp':0,
    'sma':1, 'smk':1, 'slta':1,
    'd1':1.5, 'di':1.5,
    'd2':1.7, 'dii':1.7,
    'd3':2,   'diii':2,
    'd4':3,   'div':3, 's1':3,
    's2':4,
    's3':5,
  };

  function rank(s) {
    // Pisah berdasarkan "/" — ambil rank tertinggi (mis. "S1/DIV" valid
    // untuk pegawai dengan S1 ATAU DIV, keduanya rank 3 anyway)
    const parts = String(s).split('/');
    let maxRank = 0;
    for (const part of parts) {
      // Strip SEMUA non-alphanumeric → "D-IV", "D.IV", "D IV" → "div"
      const norm = part.toLowerCase().replace(/[^a-z0-9]/g, '');
      const r = TINGKAT[norm] || 0;
      if (r > maxRank) maxRank = r;
    }
    return maxRank;
  }

  return rank(pendidikan_pegawai) >= rank(pendidikan_minimum);
}
window.pendidikanMencukupi = pendidikanMencukupi;
