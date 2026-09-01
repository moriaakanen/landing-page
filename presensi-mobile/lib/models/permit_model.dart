import 'package:cloud_firestore/cloud_firestore.dart';
import 'package:intl/intl.dart';

class PermitModel {
  final String id;
  final String userId;
  final String userName;
  final String userDepartment;
  final String officeId;
  final String officeName;
  final String purpose;
  final DateTime startTime;
  final double startLatitude;
  final double startLongitude;
  final DateTime? endTime;
  final double? endLatitude;
  final double? endLongitude;
  final String status; // 'ACTIVE' | 'COMPLETED'
  final String date; // 'yyyy-MM-dd'
  final DateTime createdAt;

  PermitModel({
    required this.id,
    required this.userId,
    required this.userName,
    required this.userDepartment,
    required this.officeId,
    required this.officeName,
    required this.purpose,
    required this.startTime,
    required this.startLatitude,
    required this.startLongitude,
    this.endTime,
    this.endLatitude,
    this.endLongitude,
    required this.status,
    required this.date,
    required this.createdAt,
  });

  /// Kadaluarsa pada esok hari jam 00.00 WIT jika tidak diselesaikan
  bool get isExpired {
    final today = DateFormat('yyyy-MM-dd').format(DateTime.now());
    return date != today && endTime == null;
  }

  bool get isActive => endTime == null && !isExpired && status == 'ACTIVE';

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
    final end = endTime ?? (isExpired ? startTime : DateTime.now());
    final diff = end.difference(startTime);
    final hours = diff.inHours;
    final minutes = diff.inMinutes % 60;

    if (hours > 0) {
      return '$hours jam $minutes menit';
    } else {
      return '$minutes menit';
    }
  }

  factory PermitModel.fromMap(Map<String, dynamic> map, String id) {
    return PermitModel(
      id: id,
      userId: map['user_id'] ?? '',
      userName: map['user_name'] ?? '',
      userDepartment: map['user_department'] ?? '',
      officeId: map['office_id'] ?? '',
      officeName: map['office_name'] ?? '',
      purpose: map['purpose'] ?? '',
      startTime: map['start_time'] != null
          ? (map['start_time'] is Timestamp
              ? (map['start_time'] as Timestamp).toDate()
              : DateTime.tryParse(map['start_time'].toString()) ?? DateTime.now())
          : DateTime.now(),
      startLatitude: (map['start_latitude'] ?? 0.0).toDouble(),
      startLongitude: (map['start_longitude'] ?? 0.0).toDouble(),
      endTime: map['end_time'] != null
          ? (map['end_time'] is Timestamp
              ? (map['end_time'] as Timestamp).toDate()
              : DateTime.tryParse(map['end_time'].toString()))
          : null,
      endLatitude: map['end_latitude'] != null ? (map['end_latitude'] as num).toDouble() : null,
      endLongitude: map['end_longitude'] != null ? (map['end_longitude'] as num).toDouble() : null,
      status: map['status'] ?? 'ACTIVE',
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
      'user_id': userId,
      'user_name': userName,
      'user_department': userDepartment,
      'office_id': officeId,
      'office_name': officeName,
      'purpose': purpose,
      'start_time': Timestamp.fromDate(startTime),
      'start_latitude': startLatitude,
      'start_longitude': startLongitude,
      'end_time': endTime != null ? Timestamp.fromDate(endTime!) : null,
      'end_latitude': endLatitude,
      'end_longitude': endLongitude,
      'status': status,
      'date': date,
      'created_at': Timestamp.fromDate(createdAt),
    };
  }
}
