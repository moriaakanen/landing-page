/**
 * 9201 PAK GENERATOR — Pengajuan Penilaian Angka Kredit
 * ─────────────────────────────────────────────────────────────────────
 * Engine untuk:
 *   1. Fetch semua data dari Supabase yang dibutuhkan tag template
 *   2. Resolve setiap tag jadi nilai final (string / number / raw XML)
 *   3. Render template .docx via docxtemplater + PizZip → Blob
 *
 * Diload SETELAH 9201-shared.js dan 9201-pak-data.js. Diload SEBELUM
 * script init() halaman yang memakainya (index.html, admin-pengajuan-
 * pak.html). docxtemplater + PizZip + FileSaver di-load on-demand
 * dengan fallback CDN — sama pattern dengan admin-surat-tugas.
 *
 * Public API (semua di window.PakGenerator):
 *   - resolveContext(opts) → Promise<context>
 *       Fetch + resolve semua tag dari Supabase. Hasil = object ready
 *       buat di-pass ke docxtemplater. Juga punya field _meta untuk
 *       preview di UI sebelum submit (ak_n1, akb, ak_tot, dst.).
 *   - renderDoc(context) → Promise<Blob>
 *       Load template + render → return Blob siap saveAs.
 *   - generateAndDownload(opts) → Promise<{context, blob, filename}>
 *       Convenience: resolveContext + renderDoc + saveAs sekaligus.
 *   - TEMPLATE_URL                         (pointer ke storage)
 *   - resolvePeriodeText(b1, b2, tahun)    (helper expose untuk UI)
 *   - resolveModeFromPeriode(b1, b2)       ('tahunan' kalau 1-12, else 'bulanan')
 *
 * Tag yang di-handle (sumber kebenaran: task.txt user) — total 30+ tag.
 * Sebagian besar plain string, dua tag pakai raw XML untuk strikethrough
 * (lihat KEP_XML_HOWTO di bawah).
 *
 * KEP_XML_HOWTO:
 *   Tag {kep_pangkat} & {kep_jenjang} butuh inline strikethrough yang
 *   tidak bisa di-render via plain string. User HARUS update template
 *   dari `{kep_pangkat}` → `{@kep_pangkat}` (sama untuk kep_jenjang).
 *   Generator akan inject Word XML <w:r><w:rPr><w:strike/></w:rPr>...
 *   yang menampilkan teks ber-strikethrough.
 */
(function () {
  'use strict';

  // ═══════════════════════════════════════════════════════════════════
  // KONSTANTA & CONFIG
  // ═══════════════════════════════════════════════════════════════════

  // URL template di Supabase Storage (bucket `template`, public).
  // Sama bucket dengan template surat tugas — sesuai keputusan user.
  const TEMPLATE_URL =
    'https://jsmmtqeoukkgugorrvmg.supabase.co/storage/v1/object/public/template/template-konversi-pk-ke-ak.docx';

  // Cache buffer template supaya tidak fetch ulang setiap render.
  let _templateBufferCache = null;

  // ═══════════════════════════════════════════════════════════════════
  // CDN LOADER — docxtemplater + PizZip + FileSaver
  // ═══════════════════════════════════════════════════════════════════
  // Pattern paralel dengan ensureDocxtemplaterLoaded() di
  // admin-surat-tugas.js. Kalau halaman pemanggil sudah include
  // <script> CDN-nya statis di HTML, ini tinggal short-circuit.

  function loadScript(src) {
    return new Promise((resolve, reject) => {
      const s = document.createElement('script');
      s.src = src;
      s.async = false;
      s.onload  = () => resolve();
      s.onerror = () => reject(new Error(`Gagal load ${src}`));
      document.head.appendChild(s);
    });
  }

  async function ensureLibsLoaded() {
    const hasDxt   = window.docxtemplater || window.Docxtemplater;
    const hasPzz   = window.PizZip        || window.pizzip;
    const hasFs    = typeof window.saveAs === 'function';
    if (hasDxt && hasPzz && hasFs) return true;

    const CDNS = [
      { dxt: 'https://unpkg.com/docxtemplater@3.68.5/build/docxtemplater.js',
        pzz: 'https://unpkg.com/pizzip@3.2.0/dist/pizzip.js',
        fs:  'https://cdnjs.cloudflare.com/ajax/libs/FileSaver.js/2.0.5/FileSaver.min.js' },
      { dxt: 'https://cdn.jsdelivr.net/npm/docxtemplater@3.68.5/build/docxtemplater.js',
        pzz: 'https://cdn.jsdelivr.net/npm/pizzip@3.2.0/dist/pizzip.js',
        fs:  'https://cdn.jsdelivr.net/npm/file-saver@2.0.5/dist/FileSaver.min.js' },
      { dxt: 'https://cdnjs.cloudflare.com/ajax/libs/docxtemplater/3.68.5/docxtemplater.js',
        pzz: 'https://cdn.jsdelivr.net/npm/pizzip@3.2.0/dist/pizzip.js',
        fs:  'https://cdnjs.cloudflare.com/ajax/libs/FileSaver.js/2.0.5/FileSaver.min.js' },
    ];

    for (const cdn of CDNS) {
      try {
        if (!window.PizZip && !window.pizzip)        await loadScript(cdn.pzz);
        if (!window.docxtemplater && !window.Docxtemplater) await loadScript(cdn.dxt);
        if (typeof window.saveAs !== 'function')     await loadScript(cdn.fs);
        if ((window.docxtemplater || window.Docxtemplater) &&
            (window.PizZip || window.pizzip) &&
            typeof window.saveAs === 'function') {
          return true;
        }
      } catch (e) {
        console.warn('[PakGen] CDN gagal, coba berikutnya:', e.message);
      }
    }
    return false;
  }

  async function loadTemplateBuffer() {
    if (_templateBufferCache) return _templateBufferCache;
    const res = await fetch(TEMPLATE_URL);
    if (!res.ok) {
      throw new Error(
        `Gagal memuat template PAK (HTTP ${res.status}). ` +
        `Pastikan file template-konversi-pk-ke-ak.docx sudah di-upload ` +
        `ke Supabase Storage bucket "template" dan dapat diakses publik.`
      );
    }
    _templateBufferCache = await res.arrayBuffer();
    return _templateBufferCache;
  }

  // ═══════════════════════════════════════════════════════════════════
  // DATE & FORMAT HELPERS
  // ═══════════════════════════════════════════════════════════════════

  /** ISO 'YYYY-MM-DD' → Date object (timezone-safe untuk format display). */
  function parseISODate(iso) {
    if (!iso) return null;
    const m = String(iso).slice(0, 10).match(/^(\d{4})-(\d{2})-(\d{2})$/);
    if (!m) return null;
    return new Date(parseInt(m[1], 10), parseInt(m[2], 10) - 1, parseInt(m[3], 10));
  }

  /** ISO date → "DD Nama-Bulan YYYY" (Indonesia). Empty kalau invalid. */
  function fmtTglID(iso) {
    if (!iso) return '';
    const d = parseISODate(iso);
    if (!d) return '';
    return `${d.getDate()} ${BULAN[d.getMonth()]} ${d.getFullYear()}`;
  }

  /**
   * Pad nomor urut surat ke 3 digit. Mis. 1 → "001", 12 → "012", 123 → "123".
   * Kalau >999 → tetap kasar string apa adanya (regulasi tidak jelas
   * untuk kasus ini, sangat unlikely tercapai dalam 1 tahun).
   */
  function padNomor(n) {
    return String(n).padStart(3, '0');
  }

  /** Escape XML special chars untuk raw XML output. */
  function xmlEscape(s) {
    return String(s ?? '')
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&apos;');
  }

  // ═══════════════════════════════════════════════════════════════════
  // SUPABASE FETCH HELPERS — pakai SUPABASE_HEADERS dari shared.js
  // ═══════════════════════════════════════════════════════════════════

  async function fetchJson(path) {
    const res = await fetch(`${SUPABASE_URL}/rest/v1/${path}`, { headers: SUPABASE_HEADERS });
    if (!res.ok) {
      throw new Error(`Fetch ${path} gagal (HTTP ${res.status})`);
    }
    return res.json();
  }

  /**
   * Fetch row pegawai dari data_pegawai by NIP.
   * Throws kalau tidak ketemu.
   */
  async function fetchPegawai(nip) {
    const rows = await fetchJson(
      `data_pegawai?NIP=eq.${encodeURIComponent(nip)}&limit=1`
    );
    if (!rows.length) throw new Error(`Pegawai dengan NIP ${nip} tidak ditemukan`);
    return rows[0];
  }

  /**
   * Lookup row terbaru di tabel riwayat_* dengan tmt ≤ tglRef. Generic
   * helper buat riwayat_gelar, riwayat_pangkat_golongan, riwayat_jabatan.
   *
   * @param {string} table     - nama tabel
   * @param {string} nip       - pegawai_nip
   * @param {string} tglRef    - ISO 'YYYY-MM-DD'
   * @param {object} [extraEq] - extra filter mis. {jenis: 'utama'}
   * @returns {object|null}
   */
  async function fetchLatestBeforeOrEq(table, nip, tglRef, extraEq) {
    let q = `${table}?pegawai_nip=eq.${encodeURIComponent(nip)}`
          + `&tmt=lte.${encodeURIComponent(tglRef)}`
          + `&order=tmt.desc&limit=1`;
    if (extraEq) {
      for (const k of Object.keys(extraEq)) {
        q += `&${k}=eq.${encodeURIComponent(extraEq[k])}`;
      }
    }
    const rows = await fetchJson(q);
    return rows[0] || null;
  }

  async function fetchPredikatTahunan(nip, tahun) {
    const rows = await fetchJson(
      `predikat_kinerja_tahunan?pegawai_nip=eq.${encodeURIComponent(nip)}`
      + `&tahun=eq.${tahun}&limit=1`
    );
    return rows[0] || null;
  }

  async function fetchPredikatBulanan(nip, tahun, bulan_start, bulan_end) {
    const rows = await fetchJson(
      `predikat_kinerja_bulanan?pegawai_nip=eq.${encodeURIComponent(nip)}`
      + `&tahun=eq.${tahun}`
      + `&bulan=gte.${bulan_start}&bulan=lte.${bulan_end}`
      + `&order=bulan.asc`
    );
    return rows;
  }

  /**
   * Pengajuan PAK terakhir untuk pegawai (semua status). Dipakai untuk
   * hitung {n1}: kalau pengajuan terakhir bulanan & sama-tahun dengan
   * pengajuan ini, n1 = tahun. Selain itu n1 = tahun - 1.
   */
  async function fetchLastPengajuan(nip) {
    const rows = await fetchJson(
      `pengajuan_pak?pegawai_nip=eq.${encodeURIComponent(nip)}`
      + `&order=created_at.desc&limit=1`
    );
    return rows[0] || null;
  }

  /** Riwayat AK terbaru by tmt ≤ tglRef. Untuk {ak_n1}. */
  async function fetchLastAK(nip, tglRef) {
    const rows = await fetchJson(
      `riwayat_angka_kredit?pegawai_nip=eq.${encodeURIComponent(nip)}`
      + `&tmt=lte.${encodeURIComponent(tglRef)}`
      + `&order=tmt.desc&limit=1`
    );
    return rows[0] || null;
  }

  // ═══════════════════════════════════════════════════════════════════
  // PERIODE FORMATTERS
  // ═══════════════════════════════════════════════════════════════════

  /**
   * {periode}: "Mei 2025" (1 bulan) atau "Januari s.d Maret 2025" (range).
   */
  function resolvePeriodeText(bulan_start, bulan_end, tahun) {
    if (bulan_start === bulan_end) {
      return `${BULAN[bulan_start - 1]} ${tahun}`;
    }
    return `${BULAN[bulan_start - 1]} s.d ${BULAN[bulan_end - 1]} ${tahun}`;
  }

  /**
   * Mode pengajuan dari range bulan: 'tahunan' kalau Jan-Des, else 'bulanan'.
   * Caveat: kalau periode Jan-Des tapi data predikat tahunan tidak ada,
   * caller akan fallback ke 'bulanan' (lihat resolveContext).
   */
  function resolveModeFromPeriode(bulan_start, bulan_end) {
    return (bulan_start === 1 && bulan_end === 12) ? 'tahunan' : 'bulanan';
  }

  /**
   * Format daftar bulan jadi string ringkas:
   *   - 1 bulan:                "Mei"
   *   - 2 bulan, contiguous:    "April - Mei"
   *   - 2 bulan, gap:           "Januari & Maret"
   *   - >2 bulan, all contig:   "April - Juni"
   *   - >2 bulan, mixed:        "Februari, Mei, Agustus - Desember"
   *
   * Aturan dari task.txt user. Empty array → "".
   *
   * @param {Array<number>} months  bulan 1-12, urut asc, unique
   * @returns {string}
   */
  function formatBulanList(months) {
    if (!Array.isArray(months) || !months.length) return '';
    const sorted = [...new Set(months)].sort((a, b) => a - b);

    if (sorted.length === 1) return BULAN[sorted[0] - 1];

    // Pecah jadi runs of consecutive months
    const runs = [];
    let curRun = [sorted[0]];
    for (let i = 1; i < sorted.length; i++) {
      if (sorted[i] === sorted[i - 1] + 1) {
        curRun.push(sorted[i]);
      } else {
        runs.push(curRun);
        curRun = [sorted[i]];
      }
    }
    runs.push(curRun);

    // Kasus 2 bulan (special, output dengan "&" kalau gap)
    if (sorted.length === 2) {
      if (runs.length === 1) {
        // contiguous: "April - Mei"
        return `${BULAN[sorted[0] - 1]} - ${BULAN[sorted[1] - 1]}`;
      }
      // gap: "Januari & Maret"
      return `${BULAN[sorted[0] - 1]} & ${BULAN[sorted[1] - 1]}`;
    }

    // >2 bulan: tiap run jadi "Awal - Akhir" kalau panjang>1, atau "Bulan" kalau singleton.
    // Gabungkan dengan ", "
    return runs.map(run => {
      if (run.length === 1) return BULAN[run[0] - 1];
      return `${BULAN[run[0] - 1]} - ${BULAN[run[run.length - 1] - 1]}`;
    }).join(', ');
  }

  // ═══════════════════════════════════════════════════════════════════
  // PREDIKAT GROUPING & RANKING
  // ═══════════════════════════════════════════════════════════════════

  /**
   * Kelompokkan rows predikat bulanan berdasarkan predikat key.
   *
   * @param {Array<{bulan, predikat}>} rows
   * @returns {Array<{key, label, persentase, rank, months}>}  sorted by rank desc (terbaik dulu)
   */
  function groupPredikatByKey(rows) {
    const map = {};
    for (const r of rows) {
      const p = getPredikat(r.predikat);
      if (!p) continue;
      if (!map[p.key]) {
        map[p.key] = { ...p, months: [] };
      }
      map[p.key].months.push(Number(r.bulan));
    }
    // Sort per group: months ascending
    Object.values(map).forEach(g => g.months.sort((a, b) => a - b));
    // Sort groups: rank descending (terbaik dulu)
    return Object.values(map).sort((a, b) => b.rank - a.rank);
  }

  // ═══════════════════════════════════════════════════════════════════
  // KALKULASI AK
  // ═══════════════════════════════════════════════════════════════════

  /**
   * AK bulanan: bulan_count/12 × persen/100 × koef
   */
  function calcAkBulanan(bulan_count, persentase, koef) {
    return (bulan_count / 12) * (persentase / 100) * koef;
  }

  /** Round ke 3 desimal. */
  function r3(n) {
    return Number(Number(n).toFixed(3));
  }

  // ═══════════════════════════════════════════════════════════════════
  // STRIKETHROUGH RAW XML BUILDER (untuk {@kep_pangkat} & {@kep_jenjang})
  // ═══════════════════════════════════════════════════════════════════

  /**
   * Build Word XML dengan satu kata yg di-strikethrough.
   *
   * Output untuk {@kep_pangkat}: dua run sequence.
   *   - kalau "kelebihan_strike=true": "Kelebihan" strike, "/Kekurangan" normal
   *   - kalau false:                   "Kelebihan/" normal, "Kekurangan" strike
   *
   * XML harus valid SEQUENCE OF RUNS — tidak boleh paragraph/section
   * elements karena akan di-inject INLINE menggantikan placeholder yg
   * berada di dalam paragraph.
   */
  function buildKepXml(strikeKelebihan) {
    if (strikeKelebihan) {
      return (
        '<w:r><w:rPr><w:strike w:val="true"/></w:rPr><w:t xml:space="preserve">Kelebihan</w:t></w:r>'
      + '<w:r><w:t xml:space="preserve">/Kekurangan</w:t></w:r>'
      );
    }
    return (
      '<w:r><w:t xml:space="preserve">Kelebihan/</w:t></w:r>'
    + '<w:r><w:rPr><w:strike w:val="true"/></w:rPr><w:t xml:space="preserve">Kekurangan</w:t></w:r>'
    );
  }

  // ═══════════════════════════════════════════════════════════════════
  // KEPUTUSAN BUILDER (halaman 3)
  // ═══════════════════════════════════════════════════════════════════

  /**
   * Build string {keputusan} berdasar 4 kondisi (rule dari task.txt user):
   *
   *   IF kurang_lebih_p < 0:
   *     IF pangkat_min == jenjang_min: "BELUM... PANGKAT/JENJANG JABATAN ... JENJANG <jenjang+1> ..."
   *     ELSE:                          "BELUM... PANGKAT ... JENJANG <jenjang> ..."
   *   IF kurang_lebih_p >= 0:
   *     (sama, tapi awalnya "DAPAT...")
   *
   * @param {object} args
   *   - kurang_lebih_p: number (bisa negatif)
   *   - pangkat_min, jenjang_min: number atau null
   *   - jabatan_aktif: string (mis. "Statistisi Ahli Pertama")
   *   - jenjang_aktif: object dari JENJANG_FUNGSIONAL (atau null)
   *   - golongan_aktif: object dari GOLONGAN_PNS (atau null)
   * @returns {string}
   */
  function buildKeputusan({ kurang_lebih_p, pangkat_min, jenjang_min,
                            jabatan_aktif, jenjang_aktif, golongan_aktif }) {
    // Defensive: kalau di puncak, pangkat_min/jenjang_min null. Generator
    // tetap output sesuatu — admin bisa edit manual dokumen kalau salah.
    if (pangkat_min == null || jenjang_min == null) {
      return 'TIDAK ADA REKOMENDASI KENAIKAN — pegawai sudah di puncak pangkat/jenjang.';
    }

    const namaJabatan = extractNamaJabatanTanpaJenjang(jabatan_aktif || '');
    const golNext = golongan_aktif ? getNextGolongan(golongan_aktif) : null;
    const jenjangNext = jenjang_aktif ? getNextJenjang(jenjang_aktif) : null;

    const golNextStr = golNext
      ? `${golNext.golongan}/${golNext.pangkat}`
      : '— (puncak)';

    // jenjang yang ditulis di keputusan:
    //   - kalau pangkat_min == jenjang_min → jenjang+1 (naik bareng)
    //   - kalau beda                       → jenjang sekarang (cuma naik pangkat)
    const samaThreshold = pangkat_min === jenjang_min;
    const jenjangStr = samaThreshold
      ? (jenjangNext ? jenjangNext.nama.toUpperCase() : '— (puncak jenjang)')
      : (jenjang_aktif ? jenjang_aktif.nama.toUpperCase() : '—');

    const prefix = kurang_lebih_p < 0
      ? 'BELUM DAPAT DIPERTIMBANGKAN UNTUK KENAIKAN'
      : 'DAPAT DIPERTIMBANGKAN UNTUK KENAIKAN';

    const tipe = samaThreshold
      ? 'PANGKAT/JENJANG JABATAN'
      : 'PANGKAT';

    return `${prefix} ${tipe} SETINGKAT LEBIH TINGGI MENJADI `
         + `${namaJabatan.toUpperCase()} JENJANG ${jenjangStr} `
         + `PANGKAT/GOLONGAN RUANG ${golNextStr.toUpperCase()}`;
  }

  // ═══════════════════════════════════════════════════════════════════
  // RESOLVE CONTEXT — main pipeline
  // ═══════════════════════════════════════════════════════════════════

  /**
   * Build full template context dari Supabase data + user input.
   *
   * @param {object} opts
   *   - pegawai_nip:        string
   *   - tahun_periode:      number
   *   - bulan_start:        number 1-12
   *   - bulan_end:          number 1-12
   *   - tgl_pengajuan:      ISO 'YYYY-MM-DD'
   *   - penandatangan_nip:  string
   *   - nomor_urut:         number  (kalau preview, generator akan output
   *                                  "{no}" placeholder; di flow real,
   *                                  caller passing no urut dari RPC create)
   * @returns {Promise<object>}  Object dengan SEMUA tag + _meta untuk UI
   */
  async function resolveContext(opts) {
    const {
      pegawai_nip, tahun_periode, bulan_start, bulan_end,
      tgl_pengajuan, penandatangan_nip,
      nomor_urut,
    } = opts;

    if (!pegawai_nip) throw new Error('pegawai_nip wajib');
    if (!tahun_periode) throw new Error('tahun_periode wajib');
    if (!bulan_start || !bulan_end) throw new Error('bulan_start & bulan_end wajib');
    if (bulan_start > bulan_end) throw new Error('bulan_start > bulan_end tidak valid');
    if (!tgl_pengajuan) throw new Error('tgl_pengajuan wajib');
    if (!penandatangan_nip) throw new Error('penandatangan_nip wajib');

    // ─── Fetch parallel ────────────────────────────────────────
    const [
      pegawai, gelarRow, pangkatRow, jabatanRow,
      predTahunan, predBulanan,
      lastPengajuan, lastAK,
      penandatanganPegawai, penandatanganJabatan,
    ] = await Promise.all([
      fetchPegawai(pegawai_nip),
      fetchLatestBeforeOrEq('riwayat_gelar', pegawai_nip, tgl_pengajuan),
      fetchLatestBeforeOrEq('riwayat_pangkat_golongan', pegawai_nip, tgl_pengajuan),
      fetchLatestBeforeOrEq('riwayat_jabatan', pegawai_nip, tgl_pengajuan, { jenis: 'utama' }),
      fetchPredikatTahunan(pegawai_nip, tahun_periode),
      fetchPredikatBulanan(pegawai_nip, tahun_periode, bulan_start, bulan_end),
      fetchLastPengajuan(pegawai_nip),
      fetchLastAK(pegawai_nip, tgl_pengajuan),
      fetchPegawai(penandatangan_nip).catch(() => null),
      fetchLatestBeforeOrEq('riwayat_jabatan', penandatangan_nip, tgl_pengajuan, { jenis: 'utama' }),
    ]);

    // ─── Validasi data minimal ────────────────────────────────
    if (!pangkatRow) {
      throw new Error('Pegawai belum punya riwayat pangkat/golongan sebelum tanggal pengajuan');
    }
    if (!jabatanRow) {
      throw new Error('Pegawai belum punya riwayat jabatan (jenis=utama) sebelum tanggal pengajuan');
    }

    // ─── Tentukan mode hitung (tahunan vs bulanan) ────────────
    const isFullYear = (bulan_start === 1 && bulan_end === 12);
    const mode = (isFullYear && predTahunan)
      ? 'tahunan'
      : 'bulanan';

    if (mode === 'bulanan') {
      // Validasi: semua bulan dalam range harus ada predikat bulanan.
      const expected = bulan_end - bulan_start + 1;
      if (predBulanan.length < expected) {
        const adaBulan = new Set(predBulanan.map(r => Number(r.bulan)));
        const missing = [];
        for (let m = bulan_start; m <= bulan_end; m++) {
          if (!adaBulan.has(m)) missing.push(BULAN[m - 1]);
        }
        throw new Error(
          `Predikat bulanan tidak lengkap untuk periode ini. ` +
          `Bulan tanpa predikat: ${missing.join(', ')}. ` +
          `Pastikan semua bulan sudah di-import ke predikat_kinerja_bulanan.`
        );
      }
    }

    // ─── Resolve jenjang & golongan ────────────────────────────
    const jenjangAktif  = extractJenjangFromJabatan(jabatanRow.jabatan);
    const golonganAktif = getGolonganByName(pangkatRow.golongan);

    if (!jenjangAktif) {
      throw new Error(
        `Tidak bisa mengenali jenjang dari jabatan "${jabatanRow.jabatan}". ` +
        `Pastikan jabatan mengandung nama jenjang (mis. "Pelaksana", "Ahli Pertama").`
      );
    }
    if (!golonganAktif) {
      throw new Error(`Tidak bisa mengenali golongan "${pangkatRow.golongan}".`);
    }

    const koef = jenjangAktif.koefisien;

    // ─── Hitung {predikat}, {predikat_2}, {persen}, {persen_2}, {ak}, {ak_2} ─
    let predikat_str = '', predikat_2_str = '';
    let persen_str   = '', persen_2_str   = '';
    let ak_num = 0, ak_2_num = 0;
    let detail_predikat;     // jsonb buat audit di pengajuan_pak
    let monthsTopGroup = []; // bulan-bulan predikat terbaik (untuk {per})
    let months2ndGroup = []; // (untuk {per2})

    if (mode === 'tahunan') {
      const p = getPredikat(predTahunan.predikat);
      if (!p) throw new Error(`Predikat tahunan "${predTahunan.predikat}" tidak dikenal`);
      predikat_str = p.label;
      persen_str   = `${p.persentase}%`;
      ak_num       = r3(koef * p.persentase / 100);
      // predikat_2, persen_2, ak_2 → blank
      detail_predikat = { mode: 'tahunan', predikat: p.label };
      // Untuk halaman 2, mode tahunan: per = "Januari - Desember" (Jan-Des all),
      // per2 = blank. n2 = blank.
      monthsTopGroup = [1,2,3,4,5,6,7,8,9,10,11,12];
    } else {
      // mode bulanan
      const groups = groupPredikatByKey(predBulanan);
      if (!groups.length) throw new Error('Tidak ada predikat bulanan yang valid di periode ini');

      // Predikat terbaik (rank tertinggi)
      const top = groups[0];
      predikat_str = top.label;
      const ak1 = calcAkBulanan(top.months.length, top.persentase, koef);
      ak_num = r3(ak1);
      persen_str = `${top.months.length}/12 X ${top.persentase}%`;
      monthsTopGroup = top.months;

      // Predikat kedua (rank tertinggi setelah top, kalau ada)
      // User kasih tahu kasus 3 predikat sangat jarang → ambil top + secondary saja, predikat ketiga di-skip dengan warning.
      if (groups.length >= 2) {
        const second = groups[1];
        predikat_2_str = second.label;
        const ak2 = calcAkBulanan(second.months.length, second.persentase, koef);
        ak_2_num   = r3(ak2);
        persen_2_str = `${second.months.length}/12 X ${second.persentase}%`;
        months2ndGroup = second.months;
        if (groups.length >= 3) {
          console.warn('[PakGen] Predikat ketiga di-skip — template hanya support 2 tipe predikat:',
                       groups.slice(2).map(g => `${g.label} (${g.months.length} bln)`).join(', '));
        }
      }

      detail_predikat = {
        mode: 'bulanan',
        groups: groups.map(g => ({
          predikat: g.label,
          persentase: g.persentase,
          months: g.months,
        })),
      };
    }

    const akb = r3(ak_num + ak_2_num);

    // ─── Hitung {ak_n1} ──────────────────────────────────────
    // riwayat_angka_kredit terbaru dengan tmt ≤ tgl_pengajuan. Default 0.
    const ak_n1 = lastAK ? r3(Number(lastAK.angka_kredit) || 0) : 0;
    const ak_tot = r3(ak_n1 + akb);

    // ─── Hitung {n1}, {n}, {n2} ──────────────────────────────
    // n  = tahun_periode
    // n1 = kalau pengajuan terakhir bulanan & tahun sama → tahun_periode
    //      kalau tidak (atau belum pernah) → tahun_periode - 1
    // n2 = tahun_periode kalau ada predikat_2, blank kalau hanya 1 tipe
    const n  = tahun_periode;
    let n1;
    if (lastPengajuan
        && lastPengajuan.tahun_periode === tahun_periode
        && lastPengajuan.mode_hitung === 'bulanan') {
      n1 = tahun_periode;
    } else {
      n1 = tahun_periode - 1;
    }
    const n2 = predikat_2_str ? tahun_periode : '';

    // ─── Build {per}, {per2} ─────────────────────────────────
    const per   = formatBulanList(monthsTopGroup);
    const per2  = months2ndGroup.length ? formatBulanList(months2ndGroup) : '';

    // ─── Hitung {pangkat_min}, {jenjang_min}, {kurang_lebih_*} ─
    const pangkat_min  = golonganAktif.kebutuhan_naik;   // dari pak-data.js (Kriteria.xlsx)
    const jenjang_min  = jenjangAktif.kebutuhan_naik;
    const kurang_lebih_p = (pangkat_min  != null) ? r3(ak_tot - pangkat_min)  : 0;
    const kurang_lebih_j = (jenjang_min  != null) ? r3(ak_tot - jenjang_min)  : 0;

    // ─── Build {kep_pangkat}, {kep_jenjang} (raw XML) ────────
    // Kalau kurang_lebih < 0 → strike "Kelebihan"
    // Kalau kurang_lebih ≥ 0 → strike "Kekurangan"
    const kep_pangkat_xml = buildKepXml(kurang_lebih_p < 0);
    const kep_jenjang_xml = buildKepXml(kurang_lebih_j < 0);

    // ─── Build {keputusan} ──────────────────────────────────
    const keputusan = buildKeputusan({
      kurang_lebih_p, pangkat_min, jenjang_min,
      jabatan_aktif: jabatanRow.jabatan,
      jenjang_aktif: jenjangAktif,
      golongan_aktif: golonganAktif,
    });

    // ─── Resolve nama (riwayat_gelar terbaru, fallback NAMA) ──
    const nama = (gelarRow && gelarRow.gelar) || pegawai.NAMA || '';

    // ─── Resolve {pangkat_golongan_tmt} & {jabatan_tmt} ──────
    const pangkat_golongan_tmt =
      `${pangkatRow.pangkat || '—'}/${pangkatRow.golongan || '—'}/${fmtTglID(pangkatRow.tmt) || '—'}`;
    const jabatan_tmt =
      `${jabatanRow.jabatan || '—'}/${fmtTglID(jabatanRow.tmt) || '—'}`;

    // ─── Penandatangan (pakai pattern surat tugas) ───────────
    const ttdNama        = penandatanganPegawai ? (penandatanganPegawai.NAMA || '') : '';
    const ttdJabRaw      = penandatanganJabatan ? (penandatanganJabatan.jabatan || '') : '';
    const ttdJabFinal    = (typeof transformJabatanPenandatangan === 'function')
                            ? transformJabatanPenandatangan(ttdJabRaw)
                            : ttdJabRaw;

    // ─── Nomor surat ────────────────────────────────────────
    // Kalau caller passing nomor_urut → format final.
    // Kalau tidak (preview mode) → output placeholder dengan "..." biar
    // jelas di preview kalau belum di-assign nomor.
    const noStr = (typeof nomor_urut === 'number' && nomor_urut > 0)
      ? padNomor(nomor_urut)
      : '___';
    const yyyy = String(tahun_periode);
    const nomor_surat_1 = `9201.${noStr}/Konv/ST/${yyyy}`;
    const nomor_surat_2 = `9201.${noStr}/Akm/ST/${yyyy}`;
    const nomor_surat_3 = `9201.${noStr}/PAK/ST/${yyyy}`;

    // ─── Format AK semua dengan locale Indonesia (koma decimal) ─
    const fmt = fmtAKNumberID;

    // ─── BUILD CONTEXT FINAL ────────────────────────────────
    const ctx = {
      // Halaman 1
      nomor_surat_1,
      nomor_surat_2,
      nomor_surat_3,
      periode: resolvePeriodeText(bulan_start, bulan_end, tahun_periode),
      nama,
      nip: pegawai.NIP || '',
      karpeg:    pegawai['NOMOR SERI KARPEG'] || '',
      ttl:       pegawai['TEMPAT/TANGGAL LAHIR'] || '',
      jk:        pegawai['JENIS KELAMIN'] || '',
      pangkat_golongan_tmt,
      jabatan_tmt,
      instansi:  pegawai['UNIT KERJA'] || '',
      predikat:    predikat_str,
      predikat_2:  predikat_2_str,
      persen:      persen_str,
      persen_2:    persen_2_str,
      koef:        fmtAKNumberID(koef),
      ak:          fmt(ak_num),
      ak_2:        ak_2_num ? fmt(ak_2_num) : '',
      tgl_pengajuan: fmtTglID(tgl_pengajuan),
      jabatan_penandatangan: ttdJabFinal,
      penandatangan: ttdNama,
      nip_penandatangan: penandatangan_nip,

      // Halaman 2
      n1: String(n1),
      n:  String(n),
      n2: n2 ? String(n2) : '',
      per,
      per2,
      ak_n1:  fmt(ak_n1),
      ak_tot: fmt(ak_tot),

      // Halaman 3
      akb:           fmt(akb),
      pangkat_min:   pangkat_min != null ? fmt(pangkat_min) : '—',
      jenjang_min:   jenjang_min != null ? fmt(jenjang_min) : '—',
      kurang_lebih_p: fmt(kurang_lebih_p),
      kurang_lebih_j: fmt(kurang_lebih_j),
      kep_pangkat:   kep_pangkat_xml,   // template harus pakai {@kep_pangkat}
      kep_jenjang:   kep_jenjang_xml,   // template harus pakai {@kep_jenjang}
      keputusan,

      // ─── _meta: untuk UI preview & untuk dipakai saat call RPC ──
      _meta: {
        mode_hitung:  mode,
        ak_didapat:   ak_num + ak_2_num,    // raw number, nanti DB simpan numeric
        ak_n1_num:    ak_n1,
        akb_num:      akb,
        ak_tot_num:   ak_tot,
        pangkat_min_num: pangkat_min,
        jenjang_min_num: jenjang_min,
        kurang_lebih_p_num: kurang_lebih_p,
        kurang_lebih_j_num: kurang_lebih_j,
        detail_predikat,
        // Snapshot identitas (buat caller display di preview)
        nama_display: nama,
        nip_display:  pegawai.NIP || '',
        pangkat_golongan_tmt,
        jabatan_tmt,
        // Snapshot penandatangan
        penandatangan_nama: ttdNama,
      },
    };

    return ctx;
  }

  // ═══════════════════════════════════════════════════════════════════
  // RENDER DOC
  // ═══════════════════════════════════════════════════════════════════

  /**
   * Render context object → Word .docx Blob via docxtemplater.
   */
  async function renderDoc(context) {
    const ok = await ensureLibsLoaded();
    if (!ok) {
      throw new Error(
        'Library docxtemplater / PizZip / FileSaver gagal dimuat. ' +
        'Periksa koneksi internet/firewall, lalu refresh halaman.'
      );
    }
    const DocxtemplaterCtor =
        (window.docxtemplater && (window.docxtemplater.default || window.docxtemplater))
     || (window.Docxtemplater && (window.Docxtemplater.default || window.Docxtemplater));
    const PizZipCtor =
        (window.PizZip && (window.PizZip.default || window.PizZip))
     || (window.pizzip  && (window.pizzip.default  || window.pizzip));

    const buf = await loadTemplateBuffer();
    const zip = new PizZipCtor(buf);
    const doc = new DocxtemplaterCtor(zip, {
      paragraphLoop: true,
      linebreaks: true,
      delimiters: { start: '{', end: '}' },
    });

    // Buang field _meta supaya tidak dilihat docxtemplater
    const { _meta, ...templateData } = context;

    try {
      doc.render(templateData);
    } catch (err) {
      console.error('[PakGen] Template render error:', err, err && err.properties);
      // Pesan error docxtemplater biasanya informatif, tapi panjang.
      // Ambil pesan utama saja untuk display ke user.
      const msg = (err && err.message) || 'Render gagal';
      throw new Error(`Template error: ${msg}`);
    }

    return doc.getZip().generate({
      type: 'blob',
      mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
      compression: 'DEFLATE',
    });
  }

  /**
   * Convenience: resolveContext + renderDoc + saveAs.
   *
   * @param {object} opts  same as resolveContext + optional `filename`
   * @returns {Promise<{context, blob, filename}>}
   */
  async function generateAndDownload(opts) {
    const context  = await resolveContext(opts);
    const blob     = await renderDoc(context);
    const filename = opts.filename || buildDefaultFilename(context);
    if (typeof window.saveAs === 'function') {
      window.saveAs(blob, filename);
    }
    return { context, blob, filename };
  }

  function buildDefaultFilename(context) {
    // "Konversi PK ke AK — <nama_singkat> — <periode>.docx"
    const nama = (context._meta && context._meta.nama_display) || 'Pegawai';
    const namaSingkat = String(nama).split(/\s+/).slice(0, 3).join(' ');
    const per = context.periode || '';
    return `Konversi PK ke AK — ${namaSingkat} — ${per}.docx`;
  }

  // ═══════════════════════════════════════════════════════════════════
  // PUBLIC API
  // ═══════════════════════════════════════════════════════════════════
  window.PakGenerator = {
    TEMPLATE_URL,
    resolveContext,
    renderDoc,
    generateAndDownload,
    // Helpers expose untuk UI (mis. modal preview):
    resolvePeriodeText,
    resolveModeFromPeriode,
    formatBulanList,
    fmtTglID,
    parseISODate,
    padNomor,
  };

  // ═══════════════════════════════════════════════════════════════════
  // DEPENDENCY: transformJabatanPenandatangan
  // ─────────────────────────────────────────────────────────────────
  // Fungsi ini didefinisikan di admin-surat-tugas.js. Kalau halaman
  // yang load generator ini TIDAK include admin-surat-tugas.js (mis.
  // halaman user index.html), kita provide fallback minimal supaya
  // generator tetap jalan.
  if (typeof window.transformJabatanPenandatangan !== 'function') {
    window.transformJabatanPenandatangan = function (jab) {
      const j = String(jab || '').trim();
      if (j === 'Kepala BPS Kabupaten Raja Ampat')   return j;
      if (j === 'Plt. Kepala Badan Pusat Statistik') return 'Plt. Kepala BPS Kabupaten Raja Ampat';
      return 'Plh. Kepala BPS Kabupaten Raja Ampat';
    };
  }
})();
