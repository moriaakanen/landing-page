import 'package:cloud_firestore/cloud_firestore.dart';
import 'package:intl/intl.dart';
import '../../models/permit_model.dart';
import '../../models/office_model.dart';
import '../../models/user_model.dart';
import 'location_service.dart';

class PermitService {
  final FirebaseFirestore _firestore = FirebaseFirestore.instance;

  String get _todayDateString => DateFormat('yyyy-MM-dd').format(DateTime.now());

  /// Mengambil data izin aktif (belum selesai) milik user tertentu
  Future<PermitModel?> getActivePermit(String userId) async {
    try {
      final snapshot = await _firestore
          .collection('permits')
          .where('user_id', isEqualTo: userId)
          .where('status', isEqualTo: 'ACTIVE')
          .limit(1)
          .get();

      if (snapshot.docs.isNotEmpty) {
        final doc = snapshot.docs.first;
        return PermitModel.fromMap(doc.data(), doc.id);
      }
    } catch (e) {
      // Ignore or log error
    }
    return null;
  }

  /// Stream izin aktif milik user secara real-time
  Stream<PermitModel?> streamActivePermit(String userId) {
    return _firestore
        .collection('permits')
        .where('user_id', isEqualTo: userId)
        .where('status', isEqualTo: 'ACTIVE')
        .limit(1)
        .snapshots()
        .map((snapshot) {
      if (snapshot.docs.isNotEmpty) {
        final doc = snapshot.docs.first;
        return PermitModel.fromMap(doc.data(), doc.id);
      }
      return null;
    });
  }

  /// Memulai Izin (Merekam Mulai Izin)
  /// Wajib berada dalam radius kantor & deteksi Fake GPS lolos
  Future<PermitModel> startPermit({
    required UserModel user,
    required OfficeModel office,
    required String purpose,
  }) async {
    if (purpose.trim().isEmpty) {
      throw "Harap isi deskripsi keperluan izin terlebih dahulu.";
    }

    // 1. Verifikasi GPS & Radius Kantor
    final locationResult = await LocationService.verifyPresenceLocation(
      officeLat: office.latitude,
      officeLng: office.longitude,
      allowedRadiusMeters: office.radiusMeters,
    );

    if (!locationResult.isSuccess || !locationResult.isWithinRadius) {
      throw locationResult.errorMessage ??
          "Gagal merekam izin: Anda harus berada di dalam radius kantor (${office.radiusMeters.toInt()}m).";
    }

    // Cek apakah masih ada izin aktif yang belum diselesaikan
    final active = await getActivePermit(user.uid);
    if (active != null) {
      throw "Anda masih memiliki izin aktif yang belum diselesaikan.";
    }

    final now = DateTime.now();
    final dateStr = DateFormat('yyyy-MM-dd').format(now);

    final newPermit = PermitModel(
      id: '',
      userId: user.uid,
      userName: user.name,
      userDepartment: user.department,
      officeId: office.id,
      officeName: office.name,
      purpose: purpose.trim(),
      startTime: now,
      startLatitude: locationResult.position!.latitude,
      startLongitude: locationResult.position!.longitude,
      endTime: null,
      endLatitude: null,
      endLongitude: null,
      status: 'ACTIVE',
      date: dateStr,
      createdAt: now,
    );

    final docRef = await _firestore.collection('permits').add(newPermit.toMap());

    return PermitModel(
      id: docRef.id,
      userId: newPermit.userId,
      userName: newPermit.userName,
      userDepartment: newPermit.userDepartment,
      officeId: newPermit.officeId,
      officeName: newPermit.officeName,
      purpose: newPermit.purpose,
      startTime: newPermit.startTime,
      startLatitude: newPermit.startLatitude,
      startLongitude: newPermit.startLongitude,
      endTime: newPermit.endTime,
      endLatitude: newPermit.endLatitude,
      endLongitude: newPermit.endLongitude,
      status: newPermit.status,
      date: newPermit.date,
      createdAt: newPermit.createdAt,
    );
  }

  /// Menyelesaikan Izin (Merekam Selesai Izin)
  /// Wajib berada dalam radius kantor saat kembali
  Future<void> endPermit({
    required String permitId,
    required OfficeModel office,
  }) async {
    // 1. Verifikasi GPS & Radius Kantor
    final locationResult = await LocationService.verifyPresenceLocation(
      officeLat: office.latitude,
      officeLng: office.longitude,
      allowedRadiusMeters: office.radiusMeters,
    );

    if (!locationResult.isSuccess || !locationResult.isWithinRadius) {
      throw locationResult.errorMessage ??
          "Gagal menyelesaikan izin: Anda harus berada di dalam radius kantor (${office.radiusMeters.toInt()}m) setelah kembali.";
    }

    final now = DateTime.now();

    await _firestore.collection('permits').doc(permitId).update({
      'end_time': Timestamp.fromDate(now),
      'end_latitude': locationResult.position!.latitude,
      'end_longitude': locationResult.position!.longitude,
      'status': 'COMPLETED',
    });
  }

  /// Stream seluruh izin pada tanggal tertentu (Realtime Monitoring Harian)
  Stream<List<PermitModel>> streamDailyPermits(String dateString) {
    return _firestore
        .collection('permits')
        .where('date', isEqualTo: dateString)
        .orderBy('start_time', descending: true)
        .snapshots()
        .map((snapshot) {
      return snapshot.docs
          .map((doc) => PermitModel.fromMap(doc.data(), doc.id))
          .toList();
    });
  }

  /// Mengambil data seluruh izin pada tanggal tertentu
  Future<List<PermitModel>> getDailyPermits(String dateString) async {
    final snapshot = await _firestore
        .collection('permits')
        .where('date', isEqualTo: dateString)
        .orderBy('start_time', descending: true)
        .get();

    return snapshot.docs
        .map((doc) => PermitModel.fromMap(doc.data(), doc.id))
        .toList();
  }
}
