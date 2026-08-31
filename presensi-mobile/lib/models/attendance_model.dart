class AttendanceDetail {
  final DateTime time;
  final double latitude;
  final double longitude;
  final double distanceMeters;
  final bool isMockLocation;
  final String status; // 'ON_TIME' | 'LATE' | 'EARLY_LEAVE' | 'NORMAL'

  AttendanceDetail({
    required this.time,
    required this.latitude,
    required this.longitude,
    required this.distanceMeters,
    required this.isMockLocation,
    this.status = 'NORMAL',
  });

  factory AttendanceDetail.fromMap(Map<String, dynamic> map) {
    return AttendanceDetail(
      time: (map['time'] != null)
          ? DateTime.tryParse(map['time'].toString()) ?? DateTime.now()
          : DateTime.now(),
      latitude: (map['latitude'] as num?)?.toDouble() ?? 0.0,
      longitude: (map['longitude'] as num?)?.toDouble() ?? 0.0,
      distanceMeters: (map['distance_meters'] as num?)?.toDouble() ?? 0.0,
      isMockLocation: map['is_mock_location'] ?? false,
      status: map['status'] ?? 'NORMAL',
    );
  }

  Map<String, dynamic> toMap() {
    return {
      'time': time.toIso8601String(),
      'latitude': latitude,
      'longitude': longitude,
      'distance_meters': distanceMeters,
      'is_mock_location': isMockLocation,
      'status': status,
    };
  }
}

class AttendanceModel {
  final String id;
  final String userId;
  final String userName;
  final String date; // YYYY-MM-DD
  final AttendanceDetail? checkIn;
  final AttendanceDetail? checkOut;

  AttendanceModel({
    required this.id,
    required this.userId,
    required this.userName,
    required this.date,
    this.checkIn,
    this.checkOut,
  });

  factory AttendanceModel.fromMap(Map<String, dynamic> map, String id) {
    return AttendanceModel(
      id: id,
      userId: map['user_id'] ?? '',
      userName: map['user_name'] ?? '',
      date: map['date'] ?? '',
      checkIn: map['check_in'] != null
          ? AttendanceDetail.fromMap(Map<String, dynamic>.from(map['check_in']))
          : null,
      checkOut: map['check_out'] != null
          ? AttendanceDetail.fromMap(Map<String, dynamic>.from(map['check_out']))
          : null,
    );
  }

  Map<String, dynamic> toMap() {
    return {
      'user_id': userId,
      'user_name': userName,
      'date': date,
      'check_in': checkIn?.toMap(),
      'check_out': checkOut?.toMap(),
    };
  }
}
