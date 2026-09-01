import 'package:cloud_firestore/cloud_firestore.dart';
import 'package:firebase_auth/firebase_auth.dart';
import 'package:shared_preferences/shared_preferences.dart';
import '../constants/employees_data.dart';
import '../../models/user_model.dart';

class AuthService {
  final FirebaseAuth _auth = FirebaseAuth.instance;
  final FirebaseFirestore _firestore = FirebaseFirestore.instance;

  static const String _prefKeyUsername = 'active_logged_in_username';

  User? get currentUser => _auth.currentUser;

  /// Login dengan username / email & kata sandi
  Future<UserModel> login({
    required String usernameOrEmail,
    required String password,
  }) async {
    final cleanInput = usernameOrEmail.trim().toLowerCase();
    final cleanPassword = password.trim();

    // 1. Validasi akun terlarang: pegawai@kantor.com sudah dihapus / dinonaktifkan
    if (cleanInput == 'pegawai@kantor.com' || cleanInput == 'pegawai') {
      throw "Akun pegawai@kantor.com telah dinonaktifkan dan tidak dapat digunakan lagi.";
    }

    if (cleanInput.isEmpty || cleanPassword.isEmpty) {
      throw "Harap masukkan username dan kata sandi.";
    }

    // 2. Cari data pegawai di Master List
    final normalizedUsername = cleanInput.contains('@')
        ? cleanInput.split('@')[0]
        : cleanInput;

    final employeeMaster = AppEmployees.list.firstWhere(
      (e) => e.username.toLowerCase() == normalizedUsername,
      orElse: () => EmployeeData(
        username: normalizedUsername,
        fullName: normalizedUsername,
        department: 'Pegawai',
      ),
    );

    // Cek apakah username valid ada di daftar pegawai
    final bool isKnownEmployee = AppEmployees.list.any(
      (e) => e.username.toLowerCase() == normalizedUsername,
    );

    if (!isKnownEmployee) {
      throw "Username '$usernameOrEmail' tidak terdaftar dalam sistem pegawai.";
    }

    final docId = 'user_${employeeMaster.username.replaceAll('.', '_')}';
    final userDocRef = _firestore.collection('users').doc(docId);

    try {
      final userDoc = await userDocRef.get();

      if (userDoc.exists && userDoc.data() != null) {
        final data = userDoc.data() as Map<String, dynamic>;
        final storedPassword = data['password']?.toString() ?? '123456';
        final bool isPasswordChanged = data['is_password_changed'] ?? false;

        // Verifikasi kata sandi
        if (cleanPassword != storedPassword) {
          throw "Kata sandi yang Anda masukkan salah.";
        }

        final user = UserModel(
          uid: docId,
          username: employeeMaster.username,
          name: data['name'] ?? employeeMaster.fullName,
          email: '${employeeMaster.username}@kantor.go.id',
          department: data['department'] ?? employeeMaster.department,
          role: data['role'] ?? 'employee',
          officeId: data['office_id'] ?? 'office_main',
          isPasswordChanged: isPasswordChanged,
        );

        if (isPasswordChanged) {
          await _saveSession(employeeMaster.username);
        }

        return user;
      } else {
        // Belum pernah login sebelumnya -> Cek password default: 123456
        if (cleanPassword != '123456') {
          throw "Kata sandi salah. Gunakan kata sandi default (123456) untuk login pertama kali.";
        }

        final newUser = UserModel(
          uid: docId,
          username: employeeMaster.username,
          name: employeeMaster.fullName,
          email: '${employeeMaster.username}@kantor.go.id',
          department: employeeMaster.department,
          role: 'employee',
          officeId: 'office_main',
          isPasswordChanged: false, // Wajib ganti password saat login pertama
        );

        // Inisialisasi dokumen di Firestore
        await userDocRef.set({
          ...newUser.toMap(),
          'password': '123456',
        });

        return newUser;
      }
    } catch (e) {
      if (e is String) rethrow;
      // Fallback offline / network error jika Firestore belum reachable
      if (cleanPassword == '123456') {
        return UserModel(
          uid: docId,
          username: employeeMaster.username,
          name: employeeMaster.fullName,
          email: '${employeeMaster.username}@kantor.go.id',
          department: employeeMaster.department,
          role: 'employee',
          officeId: 'office_main',
          isPasswordChanged: false,
        );
      }
      throw e.toString();
    }
  }

  /// Mengganti kata sandi pegawai (minimal 6 karakter)
  Future<UserModel> changePassword({
    required UserModel user,
    required String newPassword,
  }) async {
    final cleanPass = newPassword.trim();
    if (cleanPass.length < 6) {
      throw "Kata sandi baru minimal harus 6 karakter.";
    }

    final docId = user.uid;
    await _firestore.collection('users').doc(docId).set({
      ...user.toMap(),
      'password': cleanPass,
      'is_password_changed': true,
      'updated_at': FieldValue.serverTimestamp(),
    }, SetOptions(merge: true));

    await _saveSession(user.username);

    return user.copyWith(isPasswordChanged: true);
  }

  /// Ambil profil user saat ini yang tersimpan di sesi lokal
  Future<UserModel?> getCurrentUserProfile() async {
    try {
      final prefs = await SharedPreferences.getInstance();
      final username = prefs.getString(_prefKeyUsername);

      if (username == null || username.isEmpty) {
        return null;
      }

      final docId = 'user_${username.replaceAll('.', '_')}';
      final doc = await _firestore.collection('users').doc(docId).get();

      if (doc.exists && doc.data() != null) {
        return UserModel.fromMap(doc.data() as Map<String, dynamic>, doc.id);
      }

      // Fallback dari list pegawai jika offline
      final emp = AppEmployees.list.firstWhere(
        (e) => e.username.toLowerCase() == username.toLowerCase(),
        orElse: () => EmployeeData(username: username, fullName: username),
      );

      return UserModel(
        uid: docId,
        username: emp.username,
        name: emp.fullName,
        email: '${emp.username}@kantor.go.id',
        department: emp.department,
        role: 'employee',
        officeId: 'office_main',
        isPasswordChanged: true,
      );
    } catch (_) {
      return null;
    }
  }

  Future<void> _saveSession(String username) async {
    try {
      final prefs = await SharedPreferences.getInstance();
      await prefs.setString(_prefKeyUsername, username);
    } catch (_) {}
  }

  /// Logout
  Future<void> logout() async {
    try {
      final prefs = await SharedPreferences.getInstance();
      await prefs.remove(_prefKeyUsername);
      await _auth.signOut();
    } catch (_) {}
  }
}
