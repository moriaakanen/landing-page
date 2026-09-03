import 'package:cloud_firestore/cloud_firestore.dart';
import 'package:intl/intl.dart';
import '../../models/cepu_model.dart';
import '../../models/office_model.dart';
import '../../models/user_model.dart';
import '../constants/employees_data.dart';
import 'location_service.dart';

class CepuService {
  final FirebaseFirestore _firestore = FirebaseFirestore.instance;

  /// Mengambil daftar seluruh user pegawai untuk pilihan target pelaporan
  Future<List<UserModel>> getEmployeesList({required String excludeUid}) async {
    try {
      final snapshot = await _firestore.collection('users').get();
      if (snapshot.docs.isNotEmpty) {
        final users = snapshot.docs
            .map((doc) => UserModel.fromMap(doc.data(), doc.id))
            .where((u) => u.uid != excludeUid && u.email != 'pegawai@kantor.com')
            .toList();
        if (users.isNotEmpty && users.length >= AppEmployees.list.length) {
          return users;
        }
      }
    } catch (e) {
      // Fallback to master list
    }

    // Default 21 Pegawai Resmi
    return AppEmployees.list
        .map((emp) => UserModel(
              uid: 'user_${emp.username.replaceAll('.', '_')}',
              username: emp.username,
              name: emp.fullName,
              email: '${emp.username}@kantor.go.id',
              department: '',
              role: 'employee',
              officeId: 'office_main',
            ))
        .where((u) => u.uid != excludeUid)
        .toList();
  }

  /// Mengambil daftar UID pegawai yang sedang izin aktif hari ini
  Future<Set<String>> getActivePermitUserIdsToday() async {
    final todayStr = DateFormat('yyyy-MM-dd').format(DateTime.now());
    final set = <String>{};
    try {
      final snapshot = await _firestore
          .collection('permits')
          .where('date', isEqualTo: todayStr)
          .where('status', isEqualTo: 'ACTIVE')
          .where('end_time', isNull: true)
          .get();
      for (final doc in snapshot.docs) {
        final uid = doc.data()['user_id'] as String?;
        if (uid != null && uid.isNotEmpty) {
          set.add(uid);
        }
      }
    } catch (_) {}
    return set;
  }

  /// Mengambil daftar UID pegawai yang sedang aktif dilaporkan Cepu hari ini
  Future<Set<String>> getActiveCepuTargetUserIdsToday() async {
    final todayStr = DateFormat('yyyy-MM-dd').format(DateTime.now());
    final set = <String>{};
    try {
      final snapshot = await _firestore
          .collection('cepu_reports')
          .where('date', isEqualTo: todayStr)
          .where('end_time', isNull: true)
          .get();
      for (final doc in snapshot.docs) {
        final uid = doc.data()['target_uid'] as String?;
        if (uid != null && uid.isNotEmpty) {
          set.add(uid);
        }
      }
    } catch (_) {}
    return set;
  }

  /// Membuat laporan Cepu baru
  Future<CepuModel> createCepuReport({
    required UserModel reporter,
    required UserModel targetUser,
    required String description,
    required DateTime startTime,
    DateTime? endTime,
    String? photoBase64,
  }) async {
    if (description.trim().isEmpty) {
      throw "Keterangan laporan wajib diisi.";
    }

    final now = DateTime.now();
    final dateStr = DateFormat('yyyy-MM-dd').format(startTime);

    final report = CepuModel(
      id: '',
      reporterUid: reporter.uid,
      reporterName: reporter.name,
      reporterDepartment: '',
      targetUid: targetUser.uid,
      targetName: targetUser.name,
      targetDepartment: '',
      description: description.trim(),
      startTime: startTime,
      endTime: endTime,
      photoBase64: photoBase64,
      verifiedByUids: [],
      verifiedByNames: [],
      status: 'PENDING',
      date: dateStr,
      createdAt: now,
    );

    final docRef = await _firestore.collection('cepu_reports').add(report.toMap());

    // Notifikasi sistem ke Firestore untuk broadcast notifikasi baru
    try {
      await _firestore.collection('notifications').add({
        'type': 'CEPU_NEW',
        'title': '🚨 Laporan Cepu Baru',
        'message': '${targetUser.name} dilaporkan tidak berada di kantor. Seluruh rekan dapat memverifikasi.',
        'target_uid': targetUser.uid,
        'reporter_name': reporter.name,
        'cepu_id': docRef.id,
        'created_at': FieldValue.serverTimestamp(),
      });
    } catch (_) {}

    return CepuModel(
      id: docRef.id,
      reporterUid: report.reporterUid,
      reporterName: report.reporterName,
      reporterDepartment: '',
      targetUid: report.targetUid,
      targetName: report.targetName,
      targetDepartment: '',
      description: report.description,
      startTime: report.startTime,
      endTime: report.endTime,
      photoBase64: report.photoBase64,
      verifiedByUids: report.verifiedByUids,
      verifiedByNames: report.verifiedByNames,
      status: report.status,
      date: report.date,
      createdAt: report.createdAt,
    );
  }

  /// Memverifikasi laporan Cepu oleh pegawai lain
  /// Minimal 4 untuk valid, dan setelah 4 pegawai lain tetap dapat memverifikasi
  Future<void> verifyCepuReport({
    required String cepuId,
    required UserModel verifier,
  }) async {
    final docRef = _firestore.collection('cepu_reports').doc(cepuId);
    final doc = await docRef.get();

    if (!doc.exists) {
      throw "Laporan tidak ditemukan.";
    }

    final data = doc.data() as Map<String, dynamic>;
    final List<String> currentVerifiedUids = List<String>.from(data['verified_by_uids'] ?? []);
    final List<String> currentVerifiedNames = List<String>.from(data['verified_by_names'] ?? []);
    final String reporterUid = data['reporter_uid'] ?? '';
    final String targetUid = data['target_uid'] ?? '';

    if (verifier.uid == reporterUid) {
      throw "Pelapor tidak dapat memverifikasi laporannya sendiri.";
    }

    if (verifier.uid == targetUid) {
      throw "Pegawai yang dilaporkan tidak dapat memverifikasi laporan ini.";
    }

    if (currentVerifiedUids.contains(verifier.uid)) {
      throw "Anda sudah memverifikasi laporan ini sebelumnya.";
    }

    currentVerifiedUids.add(verifier.uid);
    currentVerifiedNames.add(verifier.name);

    final bool isNowValid = currentVerifiedUids.length >= 4;

    await docRef.update({
      'verified_by_uids': currentVerifiedUids,
      'verified_by_names': currentVerifiedNames,
      'status': isNowValid ? 'VERIFIED' : 'PENDING',
    });
  }

  /// Mengambil laporan Cepu terverifikasi aktif untuk user yang dilaporkan pada hari ini (belum ada endTime)
  Future<CepuModel?> getActiveCepuForTarget(String targetUid) async {
    try {
      final todayStr = DateFormat('yyyy-MM-dd').format(DateTime.now());
      final snapshot = await _firestore
          .collection('cepu_reports')
          .where('target_uid', isEqualTo: targetUid)
          .where('date', isEqualTo: todayStr)
          .where('end_time', isNull: true)
          .get();

      if (snapshot.docs.isNotEmpty) {
        final list = snapshot.docs
            .map((d) => CepuModel.fromMap(d.data(), d.id))
            .where((c) => c.isValid && c.isActive) // Harus sudah terverifikasi (min 4) & belum kadaluarsa
            .toList();

        if (list.isNotEmpty) {
          list.sort((a, b) => b.createdAt.compareTo(a.createdAt));
          return list.first;
        }
      }
    } catch (_) {}
    return null;
  }

  /// Stream laporan Cepu aktif untuk user terlapor
  Stream<List<CepuModel>> streamActiveCepuForTarget(String targetUid) {
    return _firestore
        .collection('cepu_reports')
        .where('target_uid', isEqualTo: targetUid)
        .where('end_time', isNull: true)
        .snapshots()
        .map((snapshot) {
      return snapshot.docs
          .map((d) => CepuModel.fromMap(d.data(), d.id))
          .where((c) => c.isValid)
          .toList();
    });
  }

  /// Rekam waktu kembali bagi pegawai yang dilaporkan Cepu (harus dalam radius kantor)
  Future<void> recordCepuReturnTime({
    required String cepuId,
    required OfficeModel office,
  }) async {
    final locRes = await LocationService.verifyPresenceLocation(
      officeLat: office.latitude,
      officeLng: office.longitude,
      allowedRadiusMeters: office.radiusMeters,
    );

    if (!locRes.isSuccess || !locRes.isWithinRadius || locRes.isMockLocation) {
      throw locRes.errorMessage ??
          "Anda harus berada di dalam radius kantor (${office.radiusMeters.toInt()}m) untuk merekam waktu kembali.";
    }

    final docRef = _firestore.collection('cepu_reports').doc(cepuId);
    await docRef.update({
      'end_time': FieldValue.serverTimestamp(),
    });
  }

  /// Stream real-time laporan Cepu pada tanggal tertentu
  Stream<List<CepuModel>> streamDailyCepuReports(String dateString) {
    return _firestore
        .collection('cepu_reports')
        .where('date', isEqualTo: dateString)
        .snapshots()
        .map((snapshot) {
      final list = snapshot.docs
          .map((doc) => CepuModel.fromMap(doc.data(), doc.id))
          .toList();
      list.sort((a, b) => b.createdAt.compareTo(a.createdAt));
      return list;
    });
  }

  /// Mengambil seluruh laporan Cepu (valid maupun pending) pada tanggal tertentu untuk target user tertentu
  Future<List<CepuModel>> getAllDailyCepuForTarget(String targetUid, String dateString) async {
    try {
      final snapshot = await _firestore
          .collection('cepu_reports')
          .where('target_uid', isEqualTo: targetUid)
          .where('date', isEqualTo: dateString)
          .get();

      final list = snapshot.docs
          .map((d) => CepuModel.fromMap(d.data(), d.id))
          .toList();
      list.sort((a, b) => b.startTime.compareTo(a.startTime));
      return list;
    } catch (_) {
      return [];
    }
  }

  /// Mengambil laporan Cepu valid pada tanggal tertentu untuk target user tertentu
  Future<List<CepuModel>> getDailyValidCepuForTarget(String targetUid, String dateString) async {
    try {
      final snapshot = await _firestore
          .collection('cepu_reports')
          .where('target_uid', isEqualTo: targetUid)
          .where('date', isEqualTo: dateString)
          .get();

      final list = snapshot.docs
          .map((d) => CepuModel.fromMap(d.data(), d.id))
          .where((c) => c.isValid)
          .toList();
      list.sort((a, b) => b.startTime.compareTo(a.startTime));
      return list;
    } catch (_) {
      return [];
    }
  }

  /// Mengambil laporan Cepu pending hari ini yang belum diverifikasi oleh user tertentu
  Future<List<CepuModel>> getPendingVerificationCepuForUser(String userId, String dateString) async {
    try {
      final snapshot = await _firestore
          .collection('cepu_reports')
          .where('date', isEqualTo: dateString)
          .where('status', isEqualTo: 'PENDING')
          .get();

      final list = snapshot.docs
          .map((d) => CepuModel.fromMap(d.data(), d.id))
          .where((c) =>
              c.reporterUid != userId &&
              c.targetUid != userId &&
              !c.verifiedByUids.contains(userId))
          .toList();
      list.sort((a, b) => b.createdAt.compareTo(a.createdAt));
      return list;
    } catch (_) {
      return [];
    }
  }

  /// Mengambil semua laporan Cepu (valid & non-valid) dalam rentang tanggal untuk target user
  Future<List<CepuModel>> getAllCepuDateRangeForTarget(String targetUid, String startDate, String endDate) async {
    try {
      final snapshot = await _firestore
          .collection('cepu_reports')
          .where('target_uid', isEqualTo: targetUid)
          .where('date', isGreaterThanOrEqualTo: startDate)
          .where('date', isLessThanOrEqualTo: endDate)
          .get();

      final list = snapshot.docs
          .map((d) => CepuModel.fromMap(d.data(), d.id))
          .toList();
      list.sort((a, b) => b.startTime.compareTo(a.startTime));
      return list;
    } catch (_) {
      try {
        final snapshot = await _firestore
            .collection('cepu_reports')
            .where('target_uid', isEqualTo: targetUid)
            .get();
        final list = snapshot.docs
            .map((d) => CepuModel.fromMap(d.data(), d.id))
            .where((c) => c.date.compareTo(startDate) >= 0 && c.date.compareTo(endDate) <= 0)
            .toList();
        list.sort((a, b) => b.startTime.compareTo(a.startTime));
        return list;
      } catch (_) {
        return [];
      }
    }
  }

  /// Mengambil laporan Cepu valid dalam rentang tanggal untuk target user
  Future<List<CepuModel>> getValidCepuDateRangeForTarget(String targetUid, String startDate, String endDate) async {
    final all = await getAllCepuDateRangeForTarget(targetUid, startDate, endDate);
    return all.where((c) => c.isValid).toList();
  }
}

