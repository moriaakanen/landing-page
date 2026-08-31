import 'package:firebase_auth/firebase_auth.dart';
import 'package:cloud_firestore/cloud_firestore.dart';
import '../../models/user_model.dart';

class AuthService {
  final FirebaseAuth _auth = FirebaseAuth.instance;
  final FirebaseFirestore _firestore = FirebaseFirestore.instance;

  User? get currentUser => _auth.currentUser;

  Stream<User?> get authStateChanges => _auth.authStateChanges();

  /// Login dengan email & password
  Future<UserModel?> login({required String email, required String password}) async {
    try {
      UserCredential userCredential = await _auth.signInWithEmailAndPassword(
        email: email.trim(),
        password: password,
      );

      if (userCredential.user != null) {
        DocumentSnapshot userDoc = await _firestore
            .collection('users')
            .doc(userCredential.user!.uid)
            .get();

        if (userDoc.exists) {
          return UserModel.fromMap(userDoc.data() as Map<String, dynamic>, userDoc.id);
        } else {
          // Buat default user document jika belum ada
          UserModel newUser = UserModel(
            uid: userCredential.user!.uid,
            name: userCredential.user!.displayName ?? email.split('@')[0],
            email: email,
            department: 'Umum',
            role: 'employee',
            officeId: 'office_main',
          );
          await _firestore.collection('users').doc(userCredential.user!.uid).set(newUser.toMap());
          return newUser;
        }
      }
      return null;
    } on FirebaseAuthException catch (e) {
      throw e.message ?? "Gagal login. Periksa email dan password Anda.";
    } catch (e) {
      throw "Terjadi kesalahan saat login: ${e.toString()}";
    }
  }

  /// Ambil profil user saat ini
  Future<UserModel?> getCurrentUserProfile() async {
    final user = _auth.currentUser;
    if (user == null) return null;

    final doc = await _firestore.collection('users').doc(user.uid).get();
    if (doc.exists) {
      return UserModel.fromMap(doc.data() as Map<String, dynamic>, doc.id);
    }
    return null;
  }

  /// Logout
  Future<void> logout() async {
    await _auth.signOut();
  }
}
