import 'dart:math';
import 'package:geolocator/geolocator.dart';

class LocationCheckResult {
  final bool isSuccess;
  final String? errorMessage;
  final Position? position;
  final double distanceMeters;
  final bool isWithinRadius;
  final bool isMockLocation;

  LocationCheckResult({
    required this.isSuccess,
    this.errorMessage,
    this.position,
    this.distanceMeters = 0.0,
    this.isWithinRadius = false,
    this.isMockLocation = false,
  });
}

class LocationService {
  /// Memeriksa & meminta izin akses lokasi
  static Future<bool> handleLocationPermission() async {
    bool serviceEnabled = await Geolocator.isLocationServiceEnabled();
    if (!serviceEnabled) {
      return false;
    }

    LocationPermission permission = await Geolocator.checkPermission();
    if (permission == LocationPermission.denied) {
      permission = await Geolocator.requestPermission();
      if (permission == LocationPermission.denied) {
        return false;
      }
    }

    if (permission == LocationPermission.deniedForever) {
      return false;
    }

    return true;
  }

  /// Menghitung jarak antara 2 koordinat (menggunakan Haversine)
  static double calculateDistance({
    required double startLatitude,
    required double startLongitude,
    required double endLatitude,
    required double endLongitude,
  }) {
    return Geolocator.distanceBetween(
      startLatitude,
      startLongitude,
      endLatitude,
      endLongitude,
    );
  }

  /// Mengambil posisi saat ini dan mengecek apakah dalam radius target (misal 50 meter)
  /// serta mendeteksi Fake GPS / Mock Location.
  static Future<LocationCheckResult> verifyPresenceLocation({
    required double officeLat,
    required double officeLng,
    double allowedRadiusMeters = 50.0,
  }) async {
    final hasPermission = await handleLocationPermission();
    if (!hasPermission) {
      return LocationCheckResult(
        isSuccess: false,
        errorMessage: "Izin akses lokasi ditolak atau GPS belum aktif. Silakan aktifkan GPS perangkat Anda.",
      );
    }

    try {
      // Mengambil lokasi presisi tinggi
      Position position = await Geolocator.getCurrentPosition(
        desiredAccuracy: LocationAccuracy.best,
        timeLimit: const Duration(seconds: 15),
      );

      // Deteksi Fake GPS / Mock Location
      bool isMocked = position.isMocked;

      if (isMocked) {
        return LocationCheckResult(
          isSuccess: false,
          errorMessage: "⚠️ Peringatan: Aplikasi Fake GPS / Lokasi Palsu terdeteksi! Presensi dibatalkan.",
          position: position,
          isMockLocation: true,
        );
      }

      // Hitung jarak ke kantor
      double distance = calculateDistance(
        startLatitude: position.latitude,
        startLongitude: position.longitude,
        endLatitude: officeLat,
        endLongitude: officeLng,
      );

      bool withinRadius = distance <= allowedRadiusMeters;

      return LocationCheckResult(
        isSuccess: true,
        position: position,
        distanceMeters: distance,
        isWithinRadius: withinRadius,
        isMockLocation: false,
        errorMessage: withinRadius
            ? null
            : "Anda berada di luar radius kantor (${distance.toStringAsFixed(1)} meter dari kantor. Maksimal $allowedRadiusMeters meter).",
      );
    } catch (e) {
      return LocationCheckResult(
        isSuccess: false,
        errorMessage: "Gagal mendapatkan koordinat GPS: ${e.toString()}",
      );
    }
  }
}
