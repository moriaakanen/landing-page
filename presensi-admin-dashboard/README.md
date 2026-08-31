# 🖥️ Web Admin Dashboard - Presensi & Geofencing HR

Dashboard Web modern untuk HR / Administrator dalam mengelola dan memantau kehadiran pegawai, konfigurasi titik geofencing kantor, dan simulasi pengujian.

## 🌟 Fitur Dashboard
1. **Live Attendance Monitoring**:
   - Kartu statistik kehadiran harian (*Total Pegawai, Hadir Tepat Waktu, Terlambat, Belum Presensi*).
   - Tabel kehadiran realtime lengkap dengan jam masuk, jam pulang, jarak GPS, dan status validitas lokasi.
   - Fitur pencarian & filter berdasarkan status / departemen.
2. **Peta Interaktif Geofencing (Leaflet OSM)**:
   - Drag & drop pin lokasi kantor pada peta.
   - Slider pengaturan radius presensi (10m - 200m, default: 50 meter).
   - Visualisasi lingkaran hijau radius yang langsung menyesuaikan.
3. **Ekspor Data (CSV/Excel)**:
   - Download rekap kehadiran harian/bulanan dalam format CSV untuk pelaporan payroll.
4. **Simulator Pengujian Geofencing & Anti-Fraud**:
   - Simulasi presensi dari berbagai jarak (0m s/d 150m).
   - Simulasi pengujian perangkat yang mengaktifkan Fake GPS / Mock Location.
   - Konsol log realtime untuk melihat respon *geofencing engine*.

## 🚀 Cara Menjalankan
Buka file `index.html` langsung di browser Anda, atau jalankan menggunakan live server:
```bash
# Menggunakan npx serve atau Python SimpleHTTPServer
npx -y serve presensi-admin-dashboard
# atau
python -m http.server 3000
```
Lalu buka `http://localhost:3000` di browser Anda.
