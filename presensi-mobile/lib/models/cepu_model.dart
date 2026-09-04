import 'package:cloud_firestore/cloud_firestore.dart';
import 'package:intl/intl.dart';

class CepuModel {
  final String id;
  final String reporterUid;
  final String reporterName;
  final String reporterDepartment;
  final String targetUid;
  final String targetName;
  final String targetDepartment;
  final String description;
  final DateTime startTime;
  final DateTime? endTime;
  final String? returnReason;
  final String? photoBase64;
  final List<String> verifiedByUids;
  final List<String> verifiedByNames;
  final String status; // 'PENDING' | 'VERIFIED'
  final String date; // 'yyyy-MM-dd'
  final DateTime createdAt;

  CepuModel({
    required this.id,
    required this.reporterUid,
    required this.reporterName,
    required this.reporterDepartment,
    required this.targetUid,
    required this.targetName,
    required this.targetDepartment,
    required this.description,
    required this.startTime,
    this.endTime,
    this.returnReason,
    this.photoBase64,
    required this.verifiedByUids,
    required this.verifiedByNames,
    required this.status,
    required this.date,
    required this.createdAt,
  });

  int get verificationCount => verifiedByUids.length;
  bool get isValid => verificationCount >= 4;

  /// Cek apakah laporan masih dalam batas waktu verifikasi:
  /// Senin-Kamis: s.d. 16:00 WIT
  /// Jumat: s.d. 16:30 WIT
  bool get canBeVerified {
    final now = DateTime.now();
    final today = DateFormat('yyyy-MM-dd').format(now);
    if (date != today) return false;

    final isFriday = now.weekday == DateTime.friday;
    final cutoffHour = 16;
    final cutoffMinute = isFriday ? 30 : 0;
    final cutoffTime = DateTime(now.year, now.month, now.day, cutoffHour, cutoffMinute);

    return now.isBefore(cutoffTime);
  }

  /// Laporan kadaluarsa jika:
  /// 1. Belum diverifikasi minimal 4 user hingga batas jam 16:00 (Senin-Kamis) / 16:30 (Jumat) -> Tidak Valid/Kadaluarsa
  /// 2. Atau sudah valid tapi tidak menyelesaikan waktu kembali hingga esok hari
  bool get isExpired {
    final now = DateTime.now();
    final today = DateFormat('yyyy-MM-dd').format(now);
    final isFriday = now.weekday == DateTime.friday;
    final cutoffHour = 16;
    final cutoffMinute = isFriday ? 30 : 0;
    final cutoffTime = DateTime(now.year, now.month, now.day, cutoffHour, cutoffMinute);

    if (!isValid && (date != today || now.isAfter(cutoffTime))) {
      return true;
    }

    return date != today && endTime == null;
  }

  /// Laporan yang tidak mencapai 4 verifikasi hingga batas cutoff
  bool get isFailedVerification => !isValid && !canBeVerified;

  bool get isActive => isValid && endTime == null && !isExpired;

  bool isVerifiedByUser(String uid) => verifiedByUids.contains(uid);

  String get returnStatusDisplay {
    if (isFailedVerification) {
      return "Tidak Valid (Kurang dari 4 Verifikasi)";
    }
    if (endTime != null) {
      return "Kembali: ${DateFormat('HH:mm').format(endTime!)} WIT";
    }
    if (isExpired) {
      return "Tidak Kembali";
    }
    if (!isValid) {
      return "Menunggu Verifikasi ($verificationCount/4)";
    }
    return "Belum Kembali";
  }

  String get durationString {
    final startStr = DateFormat('HH:mm').format(startTime);
    if (isFailedVerification) {
      return '$startStr - Tidak Valid';
    }
    final endStr = endTime != null
        ? DateFormat('HH:mm').format(endTime!)
        : (isExpired ? 'Tidak Kembali' : 'Belum Kembali');
    return '$startStr - $endStr';
  }

  factory CepuModel.fromMap(Map<String, dynamic> map, String id) {
    return CepuModel(
      id: id,
      reporterUid: map['reporter_uid'] ?? '',
      reporterName: map['reporter_name'] ?? 'Pegawai',
      reporterDepartment: map['reporter_department'] ?? '',
      targetUid: map['target_uid'] ?? '',
      targetName: map['target_name'] ?? '',
      targetDepartment: map['target_department'] ?? '',
      description: map['description'] ?? '',
      startTime: map['start_time'] != null
          ? (map['start_time'] is Timestamp
              ? (map['start_time'] as Timestamp).toDate()
              : DateTime.tryParse(map['start_time'].toString()) ?? DateTime.now())
          : DateTime.now(),
      endTime: map['end_time'] != null
          ? (map['end_time'] is Timestamp
              ? (map['end_time'] as Timestamp).toDate()
              : DateTime.tryParse(map['end_time'].toString()))
          : null,
      returnReason: map['return_reason'],
      photoBase64: map['photo_base64'],
      verifiedByUids: List<String>.from(map['verified_by_uids'] ?? []),
      verifiedByNames: List<String>.from(map['verified_by_names'] ?? []),
      status: map['status'] ?? 'PENDING',
      date: map['date'] ?? DateFormat('yyyy-MM-dd').format(DateTime.now()),
      createdAt: map['created_at'] != null
          ? (map['created_at'] is Timestamp
              ? (map['created_at'] as Timestamp).toDate()
              : DateTime.tryParse(map['created_at'].toString()) ?? DateTime.now())
          : DateTime.now(),
    );
  }

  Map<String, dynamic> toMap() {
    return {
      'reporter_uid': reporterUid,
      'reporter_name': reporterName,
      'reporter_department': reporterDepartment,
      'target_uid': targetUid,
      'target_name': targetName,
      'target_department': targetDepartment,
      'description': description,
      'start_time': Timestamp.fromDate(startTime),
      'end_time': endTime != null ? Timestamp.fromDate(endTime!) : null,
      'return_reason': returnReason,
      'photo_base64': photoBase64,
      'verified_by_uids': verifiedByUids,
      'verified_by_names': verifiedByNames,
      'status': status,
      'date': date,
      'created_at': Timestamp.fromDate(createdAt),
    };
  }
}
