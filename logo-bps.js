// ═══════════════════════════════════════════════════════════════════════
// logo-bps.js — Logo BPS untuk kop surat tugas
// ═══════════════════════════════════════════════════════════════════════
//
// 🖼 CARA MENGGANTI DENGAN LOGO ASLI:
// ─────────────────────────────────────
// 1. Siapkan file logo PNG/JPG (disarankan ukuran persegi, min. 200×200 px)
// 2. Convert ke base64:
//    - Cara mudah: buka https://www.base64-image.de/
//    - Upload file logo → copy hasil base64 (YANG BAGIAN "data:image/png;base64,..." DIBUANG,
//      cukup string base64-nya saja, walaupun kalau di-include prefix-nya juga ok)
// 3. Replace nilai LOGO_BPS_BASE64 di bawah ini dengan base64 logo asli
// 4. Refresh halaman admin-surat-tugas.html — logo akan otomatis muncul di preview
//
// 📏 UKURAN LOGO DI SURAT:
// Diatur di fungsi buildSuratTugasDoc() di admin-surat-tugas.html:
//   transformation: { width: 90, height: 90 }  ← dalam pixel EMU (1px ≈ 9525 EMU)
// Ubah angka ini untuk memperbesar/memperkecil logo di surat.
//
// ⚠️ JIKA LOGO_BPS_BASE64 KOSONG ATAU NULL:
// Logo otomatis di-skip (tidak ditampilkan di surat) — aman.
// ═══════════════════════════════════════════════════════════════════════

// PLACEHOLDER — Logo BPS sementara (ikon sederhana berbentuk lingkaran)
// Ganti dengan logo resmi BPS Kabupaten Raja Ampat.
//
// Logo placeholder di bawah adalah PNG 1×1 transparan (invisible).
// Script akan otomatis skip jika base64 kosong/invalid.
//
const LOGO_BPS_BASE64 = '';

// ─── Contoh format setelah Anda paste base64 asli ───
// const LOGO_BPS_BASE64 = 'iVBORw0KGgoAAAANSUhEUgAAA...';  ← string panjang
//
// Atau dengan prefix data URI (juga valid):
// const LOGO_BPS_BASE64 = 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAA...';
