class OfficeModel {
  final String id;
  final String name;
  final double latitude;
  final double longitude;
  final double radiusMeters;
  final String workStartTime; // e.g., "08:00"
  final String workEndTime;   // e.g., "17:00"

  OfficeModel({
    required this.id,
    required this.name,
    required this.latitude,
    required this.longitude,
    this.radiusMeters = 50.0,
    this.workStartTime = "08:00",
    this.workEndTime = "17:00",
  });

  factory OfficeModel.fromMap(Map<String, dynamic> map, String id) {
    return OfficeModel(
      id: id,
      name: map['name'] ?? 'Kantor Pusat',
      latitude: (map['latitude'] as num?)?.toDouble() ?? -6.200000,
      longitude: (map['longitude'] as num?)?.toDouble() ?? 106.816666,
      radiusMeters: (map['radius_meters'] as num?)?.toDouble() ?? 50.0,
      workStartTime: map['work_start_time'] ?? '08:00',
      workEndTime: map['work_end_time'] ?? '17:00',
    );
  }

  Map<String, dynamic> toMap() {
    return {
      'name': name,
      'latitude': latitude,
      'longitude': longitude,
      'radius_meters': radiusMeters,
      'work_start_time': workStartTime,
      'work_end_time': workEndTime,
    };
  }
}
