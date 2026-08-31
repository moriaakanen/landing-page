import 'package:cloud_firestore/cloud_firestore.dart';
import 'package:intl/intl.dart';
import '../../models/attendance_model.dart';
import '../../models/office_model.dart';
import '../../models/user_model.dart';
import 'location_service.dart';

class AttendanceService {
  final FirebaseFirestore _firestore = FirebaseFirestore.instance;

  String get _todayDateString => DateFormat('yyyy-MM-dd').format(DateTime.now());

  /// Mengambil data konfigurasi kantor
  Future<OfficeModel> getOfficeConfig(String officeId) async {
    try {
      final doc = await _firestore.collection('offices').doc(officeId).get();
      if (doc.exists) {
        return OfficeModel.fromMap(doc.data() as Map<String, dynamic>, doc.id);
      }
    } catch (e) {
      // Fallback default jika kantor belum diset di firestore
    }
    // Default kantor Jakarta (contoh)
    return OfficeModel(
      id: 'office_main',
      name: 'Kantor Pusat',
      latitude: -6.200000,
      longitude: 106.816666,
      radiusMeters: 50.0,
      workStartTime: '08:00',
      workEndTime: '17:00',
    );
  }

  /// Mengambil log presensi hari ini untuk user
  Future<AttendanceModel?> getTodayAttendance(String userId) async {
    final docId = 'att_${userId}_$_todayDateString';
    final doc = await _firestore.collection('attendances').doc(docId).get();
    if (doc.exists) {
      return AttendanceModel.fromMap(doc.data() as Map<String, dynamic>, doc.id);
    }
    return null;
  }

  /// Stream log presensi hari ini (real-time)
  Stream<AttendanceModel?> streamTodayAttendance(String userId) {
    final docId = 'att_${userId}_$_todayDateString';
    return _firestore.collection('attendances').doc(docId).snapshots().map((doc) {
      if (doc.exists && doc.data() != null) {
        return AttendanceModel.fromMap(doc.data() as Map<String, dynamic>, doc.id);
      }
      return null;
    });
  }

  /// Stream riwayat presensi user (1 bulan terakhir)
  Stream<List<AttendanceModel>> streamUserAttendanceHistory(String userId) {
    return _firestore
        .collection('attendances')
        .where('user_id', isEqualTo: userId)
        .orderBy('date', descending: true)
        .limit(30)
        .snapshots()
        .map((snapshot) {
      return snapshot.docs
          .map((doc) => AttendanceModel.fromMap(doc.data(), doc.id))
          .toList();
    });
  }

  /// Eksekusi Absen Masuk (Check-In)
  Future<AttendanceModel> checkIn({
    required UserModel user,
    required OfficeModel office,
  }) async {
    // 1. Verifikasi Lokasi GPS & Radius 50m & Deteksi Fake GPS
    final locationResult = await LocationService.verifyPresenceLocation(
      officeLat: office.latitude,
      officeLng: office.longitude,
      allowedRadiusMeters: office.radiusMeters,
    );

    if (!locationResult.isSuccess || !locationResult.isWithinRadius) {
      throw locationResult.errorMessage ?? "Gagal verifikasi lokasi presensi.";
    }

    final now = DateTime.now();
    final docId = 'att_${user.uid}_$_todayDateString';

    // Cek status keterlambatan (misal jam masuk 08:00)
    final startTimeParts = office.workStartTime.split(':');
    final officeStartHour = int.tryParse(startTimeParts[0]) ?? 8;
    final officeStartMinute = int.tryParse(startTimeParts.length > 1 ? startTimeParts[1] : '0') ?? 0;
    
    final thresholdTime = DateTime(now.year, now.month, now.day, officeStartHour, officeStartMinute);
    final status = now.isAfter(thresholdTime) ? 'LATE' : 'ON_TIME';

    final checkInDetail = AttendanceDetail(
      time: now,
      latitude: locationResult.position!.latitude,
      longitude: locationResult.position!.longitude,
      distanceMeters: locationResult.distanceMeters,
      isMockLocation: locationResult.isMockLocation,
      status: status,
    );

    final attendance = AttendanceModel(
      id: docId,
      userId: user.uid,
      userName: user.name,
      date: _todayDateString,
      checkIn: checkInDetail,
    );

    await _firestore.collection('attendances').doc(docId).set(
      attendance.toMap(),
      SetOptions(merge: true),
    );

    return attendance;
  }

  /// Eksekusi Absen Pulang (Check-Out)
  Future<void> checkOut({
    required UserModel user,
    required OfficeModel office,
  }) async {
    // 1. Verifikasi Lokasi GPS & Radius 50m & Deteksi Fake GPS
    final locationResult = await LocationService.verifyPresenceLocation(
      officeLat: office.latitude,
      officeLng: office.longitude,
      allowedRadiusMeters: office.radiusMeters,
    );

    if (!locationResult.isSuccess || !locationResult.isWithinRadius) {
      throw locationResult.errorMessage ?? "Gagal verifikasi lokasi presensi pulang.";
    }

    final now = DateTime.now();
    final docId = 'att_${user.uid}_$_todayDateString';

    final checkOutDetail = AttendanceDetail(
      time: now,
      latitude: locationResult.position!.latitude,
      longitude: locationResult.position!.longitude,
      distanceMeters: locationResult.distanceMeters,
      isMockLocation: locationResult.isMockLocation,
      status: 'NORMAL',
    );

    await _firestore.collection('attendances').doc(docId).update({
      'check_out': checkOutDetail.toMap(),
    });
  }
}
