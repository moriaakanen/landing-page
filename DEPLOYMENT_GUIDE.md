# 🚀 Panduan Lengkap Step-by-Step: Setup, Build APK, Deploy & Install

Panduan lengkap ini membimbing Anda dari awal hingga aplikasi presensi berjalan di HP Android dan Web Admin Dashboard dapat diakses online.

---

## 📑 Daftar Isi
1. [Langkah 1: Setup Firebase (Auth & Database)](#1-setup-firebase)
2. [Langkah 2: Konfigurasi Android App dengan Firebase](#2-konfigurasi-android-app)
3. [Langkah 3: Build APK Android](#3-build-apk-android)
4. [Langkah 4: Deploy Web Admin Dashboard (Online)](#4-deploy-web-admin-dashboard)
5. [Langkah 5: Install & Uji Coba APK di HP Android](#5-install--uji-coba-di-hp)

---

<a name="1-setup-firebase"></a>
## 1. Setup Firebase

### A. Buat Project Firebase
1. Buka [Firebase Console](https://console.firebase.google.com/) dan login dengan akun Google.
2. Klik **"Add project"** (Tambah project).
3. Beri nama project, misal: `presensi-geofencing-app` -> Klik **Continue**.
4. (Opsional) Nonaktifkan Google Analytics jika tidak diperlukan -> Klik **Create project**.

### B. Aktifkan Firebase Authentication (Login Email/Password)
1. Di menu sidebar kiri, klik **Build** > **Authentication** -> Klik **Get Started**.
2. Di tab **Sign-in method**, pilih **Email/Password**.
3. Aktifkan opsi **Enable** pertama (Email/Password) -> Klik **Save**.
4. Di tab **Users**, klik **Add user** untuk membuat akun pegawai pertama:
   - **Email**: `pegawai@kantor.com`
   - **Password**: `password123`

### C. Aktifkan Cloud Firestore (Database)
1. Di menu sidebar kiri, klik **Build** > **Firestore Database** -> Klik **Create database**.
2. Pilih lokasi server (disarankan: `asia-southeast2` / Jakarta).
3. Pilih **Start in test mode** (agar aplikasi langsung bisa membaca/menulis data) -> Klik **Create**.
4. Buat dokumen konfigurasi kantor pertama:
   - Klik **Start collection** -> Collection ID: `offices` -> Klik **Next**.
   - Document ID: `office_main`
   - Tambahkan field:
     - `name` (string): `Kantor Pusat`
     - `latitude` (number): `-6.208800` *(ganti sesuai latitude kantor Anda)*
     - `longitude` (number): `106.845600` *(ganti sesuai longitude kantor Anda)*
     - `radius_meters` (number): `50`
     - `work_start_time` (string): `08:00`
     - `work_end_time` (string): `17:00`
   - Klik **Save**.

### D. Daftarkan Aplikasi Android & Download `google-services.json`
1. Di halaman utama Firebase Console, klik icon **Android** untuk menambahkan aplikasi.
2. Isi **Android package name**: `com.example.presensi_app`
3. Klik **Register app**.
4. Klik tombol **Download google-services.json**.
5. Simpan file `google-services.json` tersebut ke dalam folder:
   ```
   presensi-mobile/android/app/google-services.json
   ```

---

<a name="2-konfigurasi-android-app"></a>
## 2. Konfigurasi Android App

Pastikan file `google-services.json` sudah berada di direktori `presensi-mobile/android/app/`.

Pastikan struktur file Gradle Android sudah mengaktifkan plugin Google Services:

1. Di file `android/build.gradle` (Project-level):
   ```groovy
   buildscript {
       dependencies {
           classpath 'com.google.gms:google-services:4.4.1'
       }
   }
   ```

2. Di file `android/app/build.gradle` (App-level):
   ```groovy
   apply plugin: 'com.android.application'
   apply plugin: 'com.google.gms.google-services' // Tambahkan di baris paling bawah
   ```

---

<a name="3-build-apk-android"></a>
## 3. Build APK Android

### A. Prasyarat di Laptop/PC
1. **Flutter SDK**: [Download Flutter](https://docs.flutter.dev/get-started/install) dan pastikan perintah `flutter` dapat dijalankan di terminal/CMD.
2. **Android Studio**: Install Android SDK & Command-line Tools.
3. Jalankan `flutter doctor` di terminal untuk memastikan tidak ada tanda silang merah pada Android toolchain.

### B. Langkah Build File APK
Buka terminal / PowerShell di laptop Anda, lalu jalankan:

```bash
# 1. Masuk ke folder aplikasi mobile
cd presensi-mobile

# 2. Unduh semua pustaka / dependency
flutter pub get

# 3. Build APK mode Release
flutter build apk --release
```

> **💡 Tips Ukuran File Lebih Kecil**:
> Gunakan perintah: `flutter build apk --split-per-abi`
> Ini akan menghasilkan file APK yang dioptimalkan per arsitektur HP (misal: `app-arm64-v8a-release.apk`) dengan ukuran jauh lebih ringan (hanya ~15-20 MB).

### C. Lokasi File APK Jadi
Setelah proses build selesai, file APK Anda berada di folder:
```
presensi-mobile/build/app/outputs/flutter-apk/app-release.apk
```

---

<a name="4-deploy-web-admin-dashboard"></a>
## 4. Deploy Web Admin Dashboard (Online)

Web Admin Dashboard berupa Single Page Web modern (HTML/JS/Leaflet Map) sehingga dapat di-hosting secara gratis dengan cepat melalui beberapa opsi:

### Opsi A: Menggunakan Vercel (Paling Cepat & Mudah)
1. Install Vercel CLI: `npm install -g vercel` (atau upload lewat dashboard [vercel.com](https://vercel.com/)).
2. Jalankan perintah di terminal:
   ```bash
   cd presensi-admin-dashboard
   npx vercel
   ```
3. Tekan Enter untuk menyetujui semua pengaturan default. Dalam 10 detik, Anda akan mendapatkan URL online gratis (misal: `https://presensi-admin-xyz.vercel.app`).

### Opsi B: Menggunakan Firebase Hosting
1. Di folder `presensi-admin-dashboard`, jalankan:
   ```bash
   npm install -g firebase-tools
   firebase login
   firebase init hosting
   ```
2. Pilih project Firebase yang sudah Anda buat di Langkah 1.
3. Set public directory ke `.` (titik = folder saat ini).
4. Configure as single-page app? Ketik `N`.
5. Jalankan deploy:
   ```bash
   firebase deploy --only hosting
   ```
6. Anda akan mendapatkan URL seperti `https://presensi-geofencing-app.web.app`.

---

<a name="5-install--uji-coba-di-hp"></a>
## 5. Install & Uji Coba APK di HP Android

### A. Kirim File APK ke HP
Anda bisa mentransfer file `app-release.apk` ke HP melalui salah satu cara berikut:
- **Kabel USB**: Hubungkan HP ke PC dan copy file APK ke folder *Downloads* di HP.
- **Google Drive / WhatsApp**: Upload file APK ke Google Drive Anda, lalu buka dan unduh dari HP.
- **Download Langsung**: Jika di-host di web server/cloud storage.

### B. Install APK di HP Android
1. Buka aplikasi **File Manager** / **Pengelola File** di HP Anda.
2. Cari dan klik file `app-release.apk`.
3. Jika muncul notifikasi *"Demi keamanan, ponsel Anda tidak diizinkan memasang aplikasi dari sumber tidak dikenal"*:
   - Klik **Setelan / Settings**.
   - Aktifkan toggle **"Izinkan dari sumber ini"** (Allow from this source).
4. Klik **Install** / **Pasang**.
5. Tunggu proses instalasi selesai, lalu klik **Buka / Open**.

### C. Pengujian di HP Pegawai
1. **Izin GPS**: Saat pertama kali dibuka, aplikasi akan meminta izin lokasi. Pilih **"Saat aplikasi digunakan"** (While using the app) dan pastikan **Lokasi Akurat (Precise Location)** aktif.
2. **Login**: Masukkan email dan password yang telah dibuat di Firebase Auth (`pegawai@kantor.com` / `password123`).
3. **Cek Radar Jarak**:
   - Jika Anda berada dalam jarak $\le 50\text{ meter}$ dari koordinat kantor, radar akan berwarna **HIJAU** dan tombol **"ABSEN MASUK SEKARANG"** aktif.
   - Jika Anda berada di luar radius ($> 50\text{ m}$), radar akan berwarna kuning/oranye dan tombol dinonaktifkan.
   - Jika HP mencoba menggunakan Fake GPS / Mock Location, radar berwarna **MERAH** dan absensi otomatis ditolak untuk mencegah kecurangan.
