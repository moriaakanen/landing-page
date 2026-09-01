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
    this.photoBase64,
    required this.verifiedByUids,
    required this.verifiedByNames,
    required this.status,
    required this.date,
    required this.createdAt,
  });

  int get verificationCount => verifiedByUids.length;
  bool get isValid => verificationCount >= 4;

  /// Kadaluarsa pada esok hari jam 00.00 WIT jika tidak diselesaikan
  bool get isExpired {
    final today = DateFormat('yyyy-MM-dd').format(DateTime.now());
    return date != today && endTime == null;
  }

  bool get isActive => endTime == null && !isExpired;

  bool isVerifiedByUser(String uid) => verifiedByUids.contains(uid);

  String get returnStatusDisplay {
    if (endTime != null) {
      return "Kembali: ${DateFormat('HH:mm').format(endTime!)} WIT";
    }
    if (isExpired) {
      return "Tidak Kembali";
    }
    return "Belum Kembali";
  }

  String get durationString {
    final startStr = DateFormat('HH:mm').format(startTime);
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
      'photo_base64': photoBase64,
      'verified_by_uids': verifiedByUids,
      'verified_by_names': verifiedByNames,
      'status': status,
      'date': date,
      'created_at': Timestamp.fromDate(createdAt),
    };
  }
}
