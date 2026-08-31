# 📱 Presensi App - Mobile Android (Flutter)

Aplikasi Presensi/Absensi Karyawan berbasis Android dengan fitur Geofencing radius 50 meter dan deteksi Fake GPS / Mock Location.

## 🚀 Fitur Utama
1. **Autentikasi Pegawai**: Login dengan email & password terintegrasi Firebase Authentication.
2. **Geofencing GPS Engine**:
   - Menghitung jarak realtime dari perangkat ke titik koordinat kantor dengan rumus Haversine (`Geolocator.distanceBetween`).
   - Membatasi tombol absen hanya dapat ditekan jika $\le 50\text{ meter}$.
3. **Anti-Fraud (Fake GPS Detection)**:
   - Mendeteksi apakah perangkat menggunakan aplikasi Mock Location / Fake GPS (`position.isMocked`).
   - Menolak absensi secara otomatis dan memberi peringatan jika terdeteksi manipulasi GPS.
4. **Pencatatan Presensi**:
   - Absen Masuk (*Check-In*) dengan status otomatis (*Tepat Waktu* / *Terlambat*).
   - Absen Pulang (*Check-Out*).
   - Riwayat presensi bulanan dengan rincian jam dan jarak GPS.

---

## 🛠️ Persiapan & Menjalankan Aplikasi

### 1. Prasyarat
- Flutter SDK (>= 3.0.0)
- Android Studio / Android SDK (API Level 21+)
- Firebase Project

### 2. Hubungkan ke Firebase
1. Buat project baru di [Firebase Console](https://console.firebase.google.com/).
2. Daftarkan aplikasi Android dengan package `com.example.presensi_app`.
3. Unduh file `google-services.json` dan letakkan di:
   ```
   presensi-mobile/android/app/google-services.json
   ```
4. Aktifkan **Firebase Authentication** (metode Email/Password) dan **Cloud Firestore**.

### 3. Install Dependensi & Jalankan
```bash
cd presensi-mobile
flutter pub get
flutter run
```

### 4. Build APK Release
```bash
flutter build apk --release
```
File APK akan berada di `build/app/outputs/flutter-apk/app-release.apk`.
