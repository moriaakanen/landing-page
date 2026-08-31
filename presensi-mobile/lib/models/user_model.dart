class UserModel {
  final String uid;
  final String name;
  final String email;
  final String department;
  final String role; // 'employee' | 'admin'
  final String officeId;

  UserModel({
    required this.uid,
    required this.name,
    required this.email,
    required this.department,
    required this.role,
    required this.officeId,
  });

  factory UserModel.fromMap(Map<String, dynamic> map, String id) {
    return UserModel(
      uid: id,
      name: map['name'] ?? '',
      email: map['email'] ?? '',
      department: map['department'] ?? 'Umum',
      role: map['role'] ?? 'employee',
      officeId: map['office_id'] ?? 'office_main',
    );
  }

  Map<String, dynamic> toMap() {
    return {
      'name': name,
      'email': email,
      'department': department,
      'role': role,
      'office_id': officeId,
    };
  }
}
