class UserModel {
  final String uid;
  final String username;
  final String name;
  final String email;
  final String department;
  final String role; // 'employee' | 'admin'
  final String officeId;
  final bool isPasswordChanged;

  UserModel({
    required this.uid,
    required this.username,
    required this.name,
    required this.email,
    required this.department,
    required this.role,
    required this.officeId,
    this.isPasswordChanged = false,
  });

  bool get mustChangePassword => !isPasswordChanged;

  factory UserModel.fromMap(Map<String, dynamic> map, String id) {
    return UserModel(
      uid: id,
      username: map['username'] ?? map['name']?.toString().toLowerCase().replaceAll(' ', '') ?? id,
      name: map['name'] ?? '',
      email: map['email'] ?? '',
      department: map['department'] ?? 'Umum',
      role: map['role'] ?? 'employee',
      officeId: map['office_id'] ?? 'office_main',
      isPasswordChanged: map['is_password_changed'] ?? false,
    );
  }

  Map<String, dynamic> toMap() {
    return {
      'username': username,
      'name': name,
      'email': email,
      'department': department,
      'role': role,
      'office_id': officeId,
      'is_password_changed': isPasswordChanged,
    };
  }

  UserModel copyWith({
    String? uid,
    String? username,
    String? name,
    String? email,
    String? department,
    String? role,
    String? officeId,
    bool? isPasswordChanged,
  }) {
    return UserModel(
      uid: uid ?? this.uid,
      username: username ?? this.username,
      name: name ?? this.name,
      email: email ?? this.email,
      department: department ?? this.department,
      role: role ?? this.role,
      officeId: officeId ?? this.officeId,
      isPasswordChanged: isPasswordChanged ?? this.isPasswordChanged,
    );
  }
}
