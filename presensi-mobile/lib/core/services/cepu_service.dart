import 'package:cloud_firestore/cloud_firestore.dart';
import 'package:intl/intl.dart';
import '../../models/cepu_model.dart';
import '../../models/user_model.dart';

import '../constants/employees_data.dart';

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
              department: emp.department,
              role: 'employee',
              officeId: 'office_main',
            ))
        .where((u) => u.uid != excludeUid)
        .toList();
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
      reporterDepartment: reporter.department,
      targetUid: targetUser.uid,
      targetName: targetUser.name,
      targetDepartment: targetUser.department,
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

    return CepuModel(
      id: docRef.id,
      reporterUid: report.reporterUid,
      reporterName: report.reporterName,
      reporterDepartment: report.reporterDepartment,
      targetUid: report.targetUid,
      targetName: report.targetName,
      targetDepartment: report.targetDepartment,
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
  /// Dibutuhkan 4 pegawai untuk mencapai status VALID / VERIFIED
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
}
