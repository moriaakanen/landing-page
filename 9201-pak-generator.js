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

  function loadTemplateBufferViaXhr(url) {
    return new Promise((resolve, reject) => {
      if (typeof XMLHttpRequest === 'undefined') {
        reject(new Error('XMLHttpRequest tidak tersedia'));
        return;
      }
      const xhr = new XMLHttpRequest();
      xhr.open('GET', url, true);
      xhr.responseType = 'arraybuffer';
      xhr.timeout = 25000;
      xhr.onload = () => {
        if (xhr.status >= 200 && xhr.status < 300 && xhr.response) {
          resolve(xhr.response);
          return;
        }
        reject(new Error(`HTTP ${xhr.status || 0}`));
      };
      xhr.onerror = () => reject(new Error('koneksi ke Supabase Storage gagal'));
      xhr.ontimeout = () => reject(new Error('koneksi ke Supabase Storage timeout'));
      xhr.send();
    });
  }

  async function loadTemplateBufferViaFetch(url) {
    const res = await fetch(url, { cache: 'no-store', mode: 'cors' });
    if (!res.ok) throw new Error(`HTTP ${res.status}`);
    return res.arrayBuffer();
  }

  function embeddedTemplateBase64() {
    return (typeof window !== 'undefined' && window.PAK_TEMPLATE_DOCX_BASE64)
      ? String(window.PAK_TEMPLATE_DOCX_BASE64)
      : '';
  }

  function base64ToArrayBuffer(base64) {
    if (typeof atob !== 'function') {
      throw new Error('decoder base64 browser tidak tersedia');
    }
    const binary = atob(String(base64 || '').replace(/\s+/g, ''));
    const bytes = new Uint8Array(binary.length);
    for (let i = 0; i < binary.length; i += 1) {
      bytes[i] = binary.charCodeAt(i);
    }
    return bytes.buffer;
  }

  async function loadTemplateBuffer() {
    if (_templateBufferCache) return _templateBufferCache;

    const errors = [];
    const embedded = embeddedTemplateBase64();
    if (embedded) {
      try {
        _templateBufferCache = base64ToArrayBuffer(embedded);
        return _templateBufferCache;
      } catch (e) {
        errors.push(`embedded: ${e.message || e}`);
        console.warn('[PakGen] Template embedded gagal dibaca, coba Supabase:', e);
      }
    }

    try {
      _templateBufferCache = await loadTemplateBufferViaXhr(TEMPLATE_URL);
      return _templateBufferCache;
    } catch (e) {
      errors.push(`XHR: ${e.message || e}`);
      console.warn('[PakGen] Load template via XHR gagal, coba fetch:', e);
    }

    try {
      _templateBufferCache = await loadTemplateBufferViaFetch(TEMPLATE_URL);
      return _templateBufferCache;
    } catch (e) {
      errors.push(`fetch: ${e.message || e}`);
      console.warn('[PakGen] Load template via fetch gagal:', e);
    }

    throw new Error(
      'Gagal memuat template PAK dari Supabase Storage. ' +
      'Pastikan koneksi internet aktif dan download manager/browser extension tidak memblokir request dokumen .docx. ' +
      `Detail: ${errors.join(' | ')}`
    );
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
    let res;
    try {
      res = await fetch(`${SUPABASE_URL}/rest/v1/${path}`, { headers: SUPABASE_HEADERS });
    } catch (e) {
      throw new Error(`Gagal menghubungi Supabase untuk data PAK. Detail: ${e.message || e}`);
    }
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
      `data_pegawai?pegawai_nip=eq.${encodeURIComponent(nip)}&limit=1`
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
      + `&order=bulan.asc`
    );
    return normalizePredikatBulananRows(rows, bulan_start, bulan_end);
  }

  function quarterMonthsFromLabel(label) {
    const raw = String(label || '').trim().toLowerCase();
    if (!raw) return null;
    const compact = raw.replace(/\./g, '').replace(/\s+/g, ' ');
    const roman = compact.match(/^(?:triwulan|tw|q)?\s*(i|ii|iii|iv)$/i);
    if (!roman) return null;
    const map = { i: [1,2,3], ii: [4,5,6], iii: [7,8,9], iv: [10,11,12] };
    return map[roman[1].toLowerCase()] || null;
  }

  function normalizePredikatBulananRows(rows, bulan_start, bulan_end) {
    const byMonth = {};
    (rows || []).forEach(row => {
      const qMonths = quarterMonthsFromLabel(row.bulan_nama);
      const months = qMonths || [Number(row.bulan)];
      months.forEach(month => {
        if (!month || month < bulan_start || month > bulan_end) return;
        const expanded = {
          ...row,
          bulan: month,
          bulan_nama: BULAN[month - 1],
          _sumber_periode: qMonths ? (row.bulan_nama || `Triwulan ${Math.ceil(month / 3)}`) : (row.bulan_nama || BULAN[month - 1]),
        };
        if (!byMonth[month] || byMonth[month]._sumber_periode !== byMonth[month].bulan_nama) {
          byMonth[month] = expanded;
        }
      });
    });
    return Object.values(byMonth).sort((a, b) => Number(a.bulan) - Number(b.bulan));
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
   * Build Word XML untuk paragraf teks plain (tanpa strikethrough).
   * Dipakai untuk kasus pegawai sudah di puncak — tag {@kep_*} di-output
   * sebagai placeholder "—" agar tetap valid Word XML.
   *
   * Font/spacing match buildKepXml supaya tidak ada visual jump.
   */
  function buildPlainParagraphXml(text) {
    const pPr = '<w:pPr><w:spacing w:after="0"/><w:rPr><w:rFonts w:ascii="Cambria" w:eastAsia="Cambria" w:hAnsi="Cambria" w:cs="Cambria"/><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr></w:pPr>';
    const rPrBase = '<w:rPr><w:rFonts w:ascii="Cambria" w:eastAsia="Cambria" w:hAnsi="Cambria" w:cs="Cambria"/><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr>';
    return `<w:p>${pPr}<w:r>${rPrBase}<w:t xml:space="preserve">${xmlEscape(text)}</w:t></w:r></w:p>`;
  }

  /**
   * Build Word XML untuk tag {@kep_pangkat} / {@kep_jenjang}.
   *
   * PENTING: tag {@...} di docxtemplater menggantikan SELURUH paragraf
   * (`<w:p>...</w:p>`), bukan hanya isi run/text-nya. Jadi output di
   * sini wajib berupa COMPLETE paragraph element. Kalau cuma `<w:r>`
   * saja, hasil akhir-nya jadi malformed Word XML — file ter-download
   * tapi error "trying to open the file" saat dibuka di Word.
   *
   * Format output: 1 paragraf yang isinya 4 kemungkinan kalimat:
   *   - kep_pangkat, AK kurang  : "~Kelebihan~/Kekurangan Angka Kredit yang harus dipenuhi untuk kenaikan pangkat"
   *   - kep_pangkat, AK cukup   : "Kelebihan/~Kekurangan~ Angka Kredit yang harus dipenuhi untuk kenaikan pangkat"
   *   - kep_jenjang, AK kurang  : "~Kelebihan~/Kekurangan Angka Kredit yang harus dipenuhi untuk kenaikan jenjang"
   *   - kep_jenjang, AK cukup   : "Kelebihan/~Kekurangan~ Angka Kredit yang harus dipenuhi untuk kenaikan jenjang"
   *
   * Font/spacing: Cambria 10pt, spacing after=0 — sesuai paragraf
   * asli di template supaya tidak ada visual jump.
   *
   * @param {boolean} strikeKelebihan  true → "Kelebihan" yg di-strike, false → "Kekurangan"
   * @param {string}  jenis            'pangkat' atau 'jenjang' (untuk suffix kalimat)
   */
  function buildKepXml(strikeKelebihan, jenis) {
    // Properti paragraf & run — match template asli
    const pPr = '<w:pPr><w:spacing w:after="0"/><w:rPr><w:rFonts w:ascii="Cambria" w:eastAsia="Cambria" w:hAnsi="Cambria" w:cs="Cambria"/><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr></w:pPr>';
    const rPrBase   = '<w:rPr><w:rFonts w:ascii="Cambria" w:eastAsia="Cambria" w:hAnsi="Cambria" w:cs="Cambria"/><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr>';
    const rPrStrike = '<w:rPr><w:rFonts w:ascii="Cambria" w:eastAsia="Cambria" w:hAnsi="Cambria" w:cs="Cambria"/><w:sz w:val="20"/><w:szCs w:val="20"/><w:strike w:val="true"/></w:rPr>';

    // Suffix kalimat berdasarkan jenis
    const suffix = ` Angka Kredit yang harus dipenuhi untuk kenaikan ${jenis}`;

    // Build runs di dalam 1 paragraf:
    //   run 1: "Kelebihan" atau "Kelebihan/" (strike kalau perlu)
    //   run 2: "/Kekurangan" atau "Kekurangan" (strike kalau perlu)
    //   run 3: " Angka Kredit yang harus dipenuhi untuk kenaikan pangkat/jenjang" (selalu normal)
    let runs;
    if (strikeKelebihan) {
      // "Kelebihan" di-strike, "/Kekurangan" normal, " Angka..." normal
      runs =
        `<w:r>${rPrStrike}<w:t xml:space="preserve">Kelebihan</w:t></w:r>`
      + `<w:r>${rPrBase}<w:t xml:space="preserve">/Kekurangan${suffix}</w:t></w:r>`;
    } else {
      // "Kelebihan/" normal, "Kekurangan" di-strike, " Angka..." normal
      runs =
        `<w:r>${rPrBase}<w:t xml:space="preserve">Kelebihan/</w:t></w:r>`
      + `<w:r>${rPrStrike}<w:t xml:space="preserve">Kekurangan</w:t></w:r>`
      + `<w:r>${rPrBase}<w:t xml:space="preserve">${suffix}</w:t></w:r>`;
    }

    return `<w:p>${pPr}${runs}</w:p>`;
  }

  // ═══════════════════════════════════════════════════════════════════
  // KEPUTUSAN BUILDER (halaman 3)
  // ═══════════════════════════════════════════════════════════════════

  /**
   * Build string {keputusan} berdasar kondisi (rule dari task.txt user):
   *
   * Aturan asli (saat pangkat_min DAN jenjang_min keduanya ada):
   *   IF diff_p > 0 (AK kurang):
   *     IF pangkat_min == jenjang_min: "BELUM... PANGKAT/JENJANG JABATAN ... JENJANG <jenjang+1> ..."
   *     ELSE:                          "BELUM... PANGKAT ... JENJANG <jenjang> ..."
   *   IF diff_p <= 0 (AK cukup/lebih):
   *     (sama, tapi awalnya "DAPAT...")
   *
   * Edge cases:
   *   - Both null (IVe puncak)           : "TIDAK ADA REKOMENDASI..."
   *   - pangkat_min null saja            : tidak terjadi (kalau di IVe, jenjang juga null)
   *   - jenjang_min null saja (IVd)      : pegawai di puncak jenjang (Ahli Utama) tapi
   *                                        masih bisa naik pangkat. Output keputusan "PANGKAT" saja.
   *
   * @param {object} args
   *   - diff_p, diff_j  : SIGNED number (pangkat_min/jenjang_min − ak_total) atau null.
   *                        > 0 = AK kurang, ≤ 0 = cukup/lebih.
   *   - pangkat_min, jenjang_min: number (delta) atau null
   *   - jabatan_aktif: string (mis. "Statistisi Ahli Pertama")
   *   - jenjang_aktif: object dari JENJANG_FUNGSIONAL (atau null)
   *   - golongan_aktif: object dari GOLONGAN_PNS (atau null)
   * @returns {string}
   */
  function buildKeputusan({ diff_p, diff_j, pangkat_min, jenjang_min,
                            jabatan_aktif, jenjang_aktif, golongan_aktif }) {
    // Both puncak — pegawai sudah di pangkat & jenjang tertinggi
    if (pangkat_min == null && jenjang_min == null) {
      return 'TIDAK ADA REKOMENDASI KENAIKAN — pegawai sudah di puncak pangkat dan jenjang.';
    }
    // Pangkat puncak tapi jenjang masih bisa — case ini tidak terjadi
    // dalam data Indonesia PNS sekarang (kalau di IVe, jenjang juga puncak),
    // tapi di-handle defensif.
    if (pangkat_min == null) {
      return 'TIDAK ADA REKOMENDASI KENAIKAN PANGKAT — pegawai sudah di pangkat tertinggi (IVe).';
    }

    const namaJabatan = extractNamaJabatanTanpaJenjang(jabatan_aktif || '');
    const golNext = golongan_aktif ? getNextGolongan(golongan_aktif) : null;
    const jenjangNext = jenjang_aktif ? getNextJenjang(jenjang_aktif) : null;

    const golNextStr = golNext
      ? `${golNext.golongan}/${golNext.pangkat}`
      : '— (puncak)';

    // Special case: Ahli Utama (jenjang_min null) — tidak ada naik
    // jenjang, hanya naik pangkat. {keputusan} cuma bicarakan pangkat.
    if (jenjang_min == null) {
      // diff_p > 0 = AK kurang (BELUM); ≤ 0 = cukup/lebih (DAPAT)
      const prefix = diff_p > 0
        ? 'BELUM DAPAT DIPERTIMBANGKAN UNTUK KENAIKAN'
        : 'DAPAT DIPERTIMBANGKAN UNTUK KENAIKAN';
      const jenjangStr = jenjang_aktif ? jenjang_aktif.nama.toUpperCase() : '—';
      return `${prefix} PANGKAT SETINGKAT LEBIH TINGGI MENJADI `
           + `${namaJabatan.toUpperCase()} JENJANG ${jenjangStr} `
           + `PANGKAT/GOLONGAN RUANG ${golNextStr.toUpperCase()}`;
    }

    // Normal case: keduanya ada nilai
    // jenjang yang ditulis di keputusan:
    //   - kalau pangkat_min == jenjang_min → jenjang+1 (naik bareng)
    //   - kalau beda                       → jenjang sekarang (cuma naik pangkat)
    const samaThreshold = (pangkat_min === jenjang_min);
    const jenjangStr = samaThreshold
      ? (jenjangNext ? jenjangNext.nama.toUpperCase() : '— (puncak jenjang)')
      : (jenjang_aktif ? jenjang_aktif.nama.toUpperCase() : '—');

    // diff_p > 0 = AK kurang (BELUM); ≤ 0 = cukup/lebih (DAPAT)
    const prefix = diff_p > 0
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
          `Pastikan predikat bulanan atau triwulanan sudah di-import ke predikat_kinerja_bulanan.`
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
    // pangkat_min & jenjang_min adalah DELTA AK dari AWAL JENJANG (bukan
    // dari posisi golongan saat ini). Lihat dokumentasi GOLONGAN_PNS di
    // pak-data.js untuk detail.
    //
    // Rumus internal (signed, untuk logic decision):
    //   _diff_p = pangkat_min − ak_total
    //   _diff_j = jenjang_min − ak_total
    //
    //   Positif (>0) = AK MASIH KURANG sebesar nilai itu untuk capai threshold.
    //                  → strikethrough "Kelebihan" (artinya: tinggal "Kekurangan")
    //   Nol/Negatif  = AK sudah cukup/lebih sebesar |nilai| dari threshold.
    //                  → strikethrough "Kekurangan" (artinya: tinggal "Kelebihan")
    //
    // Display di template ({kurang_lebih_p} & {kurang_lebih_j}): nilai
    // ABSOLUTE (tanpa tanda + atau −). Tanda kurang/lebih sudah
    // di-konvey lewat strikethrough Kelebihan/Kekurangan.
    //
    // Kalau pangkat_min/jenjang_min == null → puncak. Tag display "—".
    const pangkat_min  = golonganAktif.pangkat_min;   // delta dari awal jenjang, atau null kalau puncak
    const jenjang_min  = golonganAktif.jenjang_min;   // delta dari awal jenjang, atau null kalau puncak

    const _diff_p = (pangkat_min != null) ? r3(pangkat_min - ak_tot) : null;
    const _diff_j = (jenjang_min != null) ? r3(jenjang_min - ak_tot) : null;

    // Display value: absolute, tanpa tanda
    const kurang_lebih_p = (_diff_p != null) ? r3(Math.abs(_diff_p)) : null;
    const kurang_lebih_j = (_diff_j != null) ? r3(Math.abs(_diff_j)) : null;

    // ─── Build {kep_pangkat}, {kep_jenjang} (raw XML) ────────
    // Pakai _diff (signed) untuk tentukan strikethrough:
    //   _diff > 0 → AK kurang → strike "Kelebihan"
    //   _diff ≤ 0 → AK cukup/lebih → strike "Kekurangan"
    //
    // Kalau pangkat_min/jenjang_min null (puncak) → output paragraf
    // plain "—" (tidak ada strikethrough applicable).
    const kep_pangkat_xml = (pangkat_min == null)
      ? buildPlainParagraphXml('— (sudah di pangkat puncak)')
      : buildKepXml(_diff_p > 0, 'pangkat');
    const kep_jenjang_xml = (jenjang_min == null)
      ? buildPlainParagraphXml('— (sudah di jenjang puncak)')
      : buildKepXml(_diff_j > 0, 'jenjang');

    // ─── Build {keputusan} ──────────────────────────────────
    // Pass signed _diff_p (bukan absolute kurang_lebih_p) supaya logic
    // BELUM/DAPAT tetap benar.
    const keputusan = buildKeputusan({
      diff_p: _diff_p, diff_j: _diff_j,
      pangkat_min, jenjang_min,
      jabatan_aktif: jabatanRow.jabatan,
      jenjang_aktif: jenjangAktif,
      golongan_aktif: golonganAktif,
    });

    // ─── Resolve nama (riwayat_gelar terbaru, fallback NAMA) ──
    const nama = (gelarRow && gelarRow.gelar) || (window.pegawaiNama ? window.pegawaiNama(pegawai) : (pegawai.nama || pegawai.NAMA)) || '';

    // ─── Resolve {pangkat_golongan_tmt} & {jabatan_tmt} ──────
    const pangkat_golongan_tmt =
      `${pangkatRow.pangkat || '—'}/${pangkatRow.golongan || '—'}/${fmtTglID(pangkatRow.tmt) || '—'}`;
    const jabatan_tmt =
      `${jabatanRow.jabatan || '—'}/${fmtTglID(jabatanRow.tmt) || '—'}`;

    // ─── Penandatangan (pakai pattern surat tugas) ───────────
    const ttdNama        = penandatanganPegawai ? ((window.pegawaiNama ? window.pegawaiNama(penandatanganPegawai) : (penandatanganPegawai.nama || penandatanganPegawai.NAMA)) || '') : '';
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
      nip: window.pegawaiNip ? window.pegawaiNip(pegawai) : (pegawai.pegawai_nip || pegawai.NIP || ''),
      karpeg:    window.pegawaiKarpeg ? window.pegawaiKarpeg(pegawai) : (pegawai.karpeg || pegawai['NOMOR SERI KARPEG'] || ''),
      ttl:       window.pegawaiTtl ? window.pegawaiTtl(pegawai) : (pegawai.ttl || pegawai['TEMPAT/TANGGAL LAHIR'] || ''),
      jk:        window.pegawaiJk ? window.pegawaiJk(pegawai) : (pegawai.jk || pegawai['JENIS KELAMIN'] || ''),
      pangkat_golongan_tmt,
      jabatan_tmt,
      instansi:  window.pegawaiUnitKerja ? window.pegawaiUnitKerja(pegawai) : (pegawai.unit_kerja || pegawai['UNIT KERJA'] || ''),
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
      pangkat_min:   pangkat_min   != null ? fmt(pangkat_min)   : '—',
      jenjang_min:   jenjang_min   != null ? fmt(jenjang_min)   : '—',
      kurang_lebih_p: kurang_lebih_p != null ? fmt(kurang_lebih_p) : '—',
      kurang_lebih_j: kurang_lebih_j != null ? fmt(kurang_lebih_j) : '—',
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
        pangkat_min_num: pangkat_min,        // delta atau null
        jenjang_min_num: jenjang_min,        // delta atau null
        kurang_lebih_p_num: kurang_lebih_p,  // signed, atau null
        kurang_lebih_j_num: kurang_lebih_j,  // signed, atau null
        detail_predikat,
        // Snapshot identitas (buat caller display di preview)
        nama_display: nama,
        nip_display:  window.pegawaiNip ? window.pegawaiNip(pegawai) : (pegawai.pegawai_nip || pegawai.NIP || ''),
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
      // docxtemplater "Multi error" wraps individual tag errors di
      // err.properties.errors. Tanpa unpack ini, user cuma lihat
      // pesan generic dan tidak tau tag mana yang salah.
      if (err && err.properties && Array.isArray(err.properties.errors) && err.properties.errors.length) {
        const errs = err.properties.errors;
        // Log SEMUA sub-error ke console untuk debugging
        errs.forEach((e, i) => {
          console.error(`[PakGen] Sub-error ${i + 1}/${errs.length}:`, e, e.properties);
        });
        // Bangun summary 3 error pertama untuk display ke user
        const summary = errs.slice(0, 3).map(e => {
          const props = e.properties || {};
          const tag = props.xtag || props.id || props.tag || '?';
          const explain = props.explanation || e.message || 'unknown';
          return `{${tag}}: ${explain}`;
        }).join(' | ');
        const more = errs.length > 3 ? ` (+${errs.length - 3} error lain — cek console)` : '';
        throw new Error(`Template error (${errs.length}): ${summary}${more}`);
      }
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
    // "Konversi PK ke AK - <nama_singkat> - <periode>.docx"
    const nama = (context._meta && context._meta.nama_display) || 'Pegawai';
    const namaSingkat = String(nama).split(/\s+/).slice(0, 3).join(' ');
    const per = context.periode || '';
    return `Konversi PK ke AK - ${namaSingkat} - ${per}.docx`;
  }

  // ═══════════════════════════════════════════════════════════════════
  // PREVIEW HELPERS — pola sama dengan admin-surat-tugas.js
  // ─────────────────────────────────────────────────────────────────
  // Flow preview:
  //   1. resolveContext(opts)  → context object
  //   2. renderDoc(context)    → Blob .docx
  //   3. uploadPreviewDocx(blob, pakId) → upload ke Supabase Storage
  //                              return filename di bucket
  //   4. getPreviewSignedUrl(filename) → URL temporary (10 menit expiry)
  //   5. buildOfficeViewerUrl(signedUrl) → URL final iframe
  //
  // Bucket: dipakai sama dengan surat-tugas — 'surat-tugas-preview'.
  // Filename pakai prefix 'pak_' supaya tidak bentrok dengan surat tugas.
  // URL request storage meng-encode titik pada ".docx" menjadi "%2E" supaya
  // download manager tidak mudah menangkap request upload/sign/delete sebagai
  // download dokumen. Nama object di Storage tetap berakhiran .docx.
  // Cleanup logic admin-surat-tugas.js (regex /_(\d+)\.docx$/) tetap match
  // file PAK karena timestamp ada di akhir.
  //
  // Convenience function: generatePreviewUrl(opts) — orchestrate semuanya
  // sekaligus, return { signedUrl, viewerUrl, filename, blob, context }.
  // ═══════════════════════════════════════════════════════════════════

  // Bucket di Supabase Storage. Reuse bucket surat-tugas-preview supaya
  // tidak butuh setup tambahan di Supabase Dashboard.
  const PREVIEW_BUCKET = 'surat-tugas-preview';
  const PREVIEW_TTL_MS = 10 * 60 * 1000;
  const PREVIEW_SIGNED_URL_TTL_SEC = Math.ceil(PREVIEW_TTL_MS / 1000);
  const _previewCleanupTimers = {};

  function storageObjectPath(filename) {
    return encodeURIComponent(filename).replace(/\./g, '%2E');
  }

  /**
   * Upload Blob .docx ke Supabase Storage, return filename.
   * @param {Blob}        blob   - hasil dari renderDoc(context)
   * @param {string|number} pakId  - id pengajuan_pak (untuk filename)
   * @returns {Promise<string>}    filename yang dihasilkan
   */
  async function uploadPreviewDocx(blob, pakId) {
    const filename = `pak_${pakId || 'tmp'}_${Date.now()}.docx`;
    const url = `${SUPABASE_URL}/storage/v1/object/${PREVIEW_BUCKET}/${storageObjectPath(filename)}`;
    let res;
    try {
      res = await fetch(url, {
        method: 'POST',
        headers: {
          apikey: SUPABASE_ANON_KEY,
          Authorization: `Bearer ${SUPABASE_ANON_KEY}`,
          'Content-Type': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
          'x-upsert': 'true',
        },
        body: blob,
      });
    } catch (e) {
      throw new Error(
        'Upload preview ke Supabase gagal. Jika download manager aktif, pastikan ia tidak menangkap request storage. ' +
        `Detail: ${e.message || e}`
      );
    }
    if (!res.ok) {
      let msg = `HTTP ${res.status}`;
      try { const err = await res.json(); msg = err.message || err.error || msg; } catch(_) {}
      throw new Error(`Upload preview gagal: ${msg}`);
    }
    schedulePreviewCleanup(filename);
    return filename;
  }

  /**
   * Buat signed URL untuk file di bucket preview.
   * @param {string} filename
   * @param {number} [expiresInSec=600]
   * @returns {Promise<string>}  URL lengkap siap embed di iframe
   */
  async function getPreviewSignedUrl(filename, expiresInSec) {
    const exp = expiresInSec || PREVIEW_SIGNED_URL_TTL_SEC;
    const url = `${SUPABASE_URL}/storage/v1/object/sign/${PREVIEW_BUCKET}/${storageObjectPath(filename)}`;
    let res;
    try {
      res = await fetch(url, {
        method: 'POST',
        headers: {
          apikey: SUPABASE_ANON_KEY,
          Authorization: `Bearer ${SUPABASE_ANON_KEY}`,
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({ expiresIn: exp }),
      });
    } catch (e) {
      throw new Error(`Gagal membuat signed URL preview. Detail: ${e.message || e}`);
    }
    if (!res.ok) throw new Error(`Gagal membuat signed URL (HTTP ${res.status})`);
    const data = await res.json();
    if (!data || !data.signedURL) throw new Error('Supabase tidak mengembalikan signed URL preview.');
    const signedPath = String(data.signedURL).replace(filename, storageObjectPath(filename));
    return `${SUPABASE_URL}/storage/v1${signedPath}`;
  }

  /**
   * Hapus file preview dari bucket. Fire-and-forget — tidak throw error.
   */
  async function deletePreviewFile(filename) {
    if (!filename) return;
    clearPreviewCleanupTimer(filename);
    try {
      await fetch(`${SUPABASE_URL}/storage/v1/object/${PREVIEW_BUCKET}/${storageObjectPath(filename)}`, {
        method: 'DELETE',
        headers: {
          apikey: SUPABASE_ANON_KEY,
          Authorization: `Bearer ${SUPABASE_ANON_KEY}`,
        },
      });
    } catch (e) {
      console.warn('[PakGen] Cleanup preview file gagal:', e);
    }
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

  function getPreviewTimestamp(filename) {
    const m = filename && String(filename).match(/_(\d+)(?:\.docx|_docx)$/);
    return m ? parseInt(m[1], 10) : null;
  }

  async function cleanupExpiredPreviewFiles(maxAgeMs = PREVIEW_TTL_MS) {
    try {
      const res = await fetch(`${SUPABASE_URL}/storage/v1/object/list/${PREVIEW_BUCKET}`, {
        method: 'POST',
        headers: {
          apikey: SUPABASE_ANON_KEY,
          Authorization: `Bearer ${SUPABASE_ANON_KEY}`,
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({ limit: 1000, prefix: '' }),
      });
      if (!res.ok) return;
      const files = await res.json();
      const expiredBefore = Date.now() - maxAgeMs;
      files.forEach(file => {
        const ts = getPreviewTimestamp(file && file.name);
        if (ts && ts < expiredBefore) deletePreviewFile(file.name);
      });
    } catch (_) {
      // Fire-and-forget cleanup; preview should still work if list policy is blocked.
    }
  }

  /**
   * Bangun URL iframe Office Online viewer dari signed URL.
   * Office Online butuh URL public/signed yang accessible dari Microsoft.
   */
  function buildOfficeViewerUrl(signedUrl) {
    return `https://view.officeapps.live.com/op/embed.aspx?src=${encodeURIComponent(signedUrl)}`;
  }

  /**
   * Convenience: generate dokumen + upload + signed URL + viewer URL.
   * Caller tinggal embed `viewerUrl` ke iframe.
   *
   * @param {object} opts  - same as resolveContext (pegawai_nip, bulan_*, dst.)
   *                         + opts.pakId (untuk filename, optional)
   * @returns {Promise<{viewerUrl, signedUrl, filename, blob, context}>}
   */
  async function generatePreviewUrl(opts) {
    const context  = await resolveContext(opts);
    const blob     = await renderDoc(context);
    const pakId    = (opts && opts.pakId) || (opts && opts.id) || 'tmp';
    cleanupExpiredPreviewFiles();
    const filename = await uploadPreviewDocx(blob, pakId);
    const signedUrl = await getPreviewSignedUrl(filename);
    const viewerUrl = buildOfficeViewerUrl(signedUrl);
    return { viewerUrl, signedUrl, filename, blob, context };
  }

  // ═══════════════════════════════════════════════════════════════════
  // PUBLIC API
  // ═══════════════════════════════════════════════════════════════════
  window.PakGenerator = {
    TEMPLATE_URL,
    resolveContext,
    renderDoc,
    generateAndDownload,
    // Preview API (mirror admin-surat-tugas.js):
    uploadPreviewDocx,
    getPreviewSignedUrl,
    deletePreviewFile,
    schedulePreviewCleanup,
    cleanupExpiredPreviewFiles,
    PREVIEW_TTL_MS,
    buildOfficeViewerUrl,
    generatePreviewUrl,
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
