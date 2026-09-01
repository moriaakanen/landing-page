import 'package:flutter/material.dart';
import '../models/user_model.dart';
import '../core/services/auth_service.dart';
import 'home_attendance_view.dart';

class ChangePasswordView extends StatefulWidget {
  final UserModel user;

  const ChangePasswordView({Key? key, required this.user}) : super(key: key);

  @override
  State<ChangePasswordView> createState() => _ChangePasswordViewState();
}

class _ChangePasswordViewState extends State<ChangePasswordView> {
  final _newPasswordController = TextEditingController();
  final _confirmPasswordController = TextEditingController();
  final _authService = AuthService();

  bool _obscureNew = true;
  bool _obscureConfirm = true;
  bool _isLoading = false;
  String? _errorMessage;

  @override
  void dispose() {
    _newPasswordController.dispose();
    _confirmPasswordController.dispose();
    super.dispose();
  }

  Future<void> _handleChangePassword() async {
    final newPass = _newPasswordController.text.trim();
    final confirmPass = _confirmPasswordController.text.trim();

    if (newPass.length < 6) {
      setState(() {
        _errorMessage = "Kata sandi baru minimal harus 6 karakter.";
      });
      return;
    }

    if (newPass != confirmPass) {
      setState(() {
        _errorMessage = "Konfirmasi kata sandi tidak cocok.";
      });
      return;
    }

    setState(() {
      _isLoading = true;
      _errorMessage = null;
    });

    try {
      final updatedUser = await _authService.changePassword(
        user: widget.user,
        newPassword: newPass,
      );

      if (mounted) {
        ScaffoldMessenger.of(context).showSnackBar(
          const SnackBar(
            content: Text("✅ Kata sandi baru berhasil disimpan! Selamat datang."),
            backgroundColor: Color(0xFF10B981),
          ),
        );

        Navigator.pushReplacement(
          context,
          MaterialPageRoute(
            builder: (context) => HomeAttendanceView(user: updatedUser),
          ),
        );
      }
    } catch (e) {
      setState(() {
        _errorMessage = e.toString();
      });
    } finally {
      if (mounted) {
        setState(() => _isLoading = false);
      }
    }
  }

  @override
  Widget build(BuildContext context) {
    return Scaffold(
      backgroundColor: const Color(0xFFF8FAFC),
      body: SafeArea(
        child: Center(
          child: SingleChildScrollView(
            padding: const EdgeInsets.symmetric(horizontal: 28, vertical: 20),
            child: Column(
              crossAxisAlignment: CrossAxisAlignment.stretch,
              children: [
                // Key / Security Icon Header
                Center(
                  child: Container(
                    width: 80,
                    height: 80,
                    decoration: BoxDecoration(
                      gradient: const LinearGradient(
                        colors: [Color(0xFF2563EB), Color(0xFF4F46E5)],
                        begin: Alignment.topLeft,
                        end: Alignment.bottomRight,
                      ),
                      borderRadius: BorderRadius.circular(24),
                      boxShadow: [
                        BoxShadow(
                          color: const Color(0xFF2563EB).withOpacity(0.35),
                          blurRadius: 20,
                          offset: const Offset(0, 10),
                        ),
                      ],
                    ),
                    child: const Center(
                      child: Icon(
                        Icons.lock_reset_rounded,
                        color: Colors.white,
                        size: 44,
                      ),
                    ),
                  ),
                ),

                const SizedBox(height: 24),

                const Text(
                  "Ganti Kata Sandi Baru",
                  textAlign: TextAlign.center,
                  style: TextStyle(
                    fontSize: 22,
                    fontWeight: FontWeight.w900,
                    color: Color(0xFF0F172A),
                    letterSpacing: -0.3,
                  ),
                ),

                const SizedBox(height: 8),

                Text(
                  "Halo, ${widget.user.name}!\nIni adalah login pertama Anda. Demi keamanan akun, silakan buat kata sandi baru Anda.",
                  textAlign: TextAlign.center,
                  style: const TextStyle(
                    fontSize: 13,
                    color: Color(0xFF64748B),
                    height: 1.4,
                  ),
                ),

                const SizedBox(height: 24),

                // Notice Box
                Container(
                  padding: const EdgeInsets.symmetric(horizontal: 14, vertical: 10),
                  decoration: BoxDecoration(
                    color: const Color(0xFFEFF6FF),
                    borderRadius: BorderRadius.circular(14),
                    border: Border.all(color: const Color(0xFFBFDBFE)),
                  ),
                  child: Row(
                    children: const [
                      Icon(Icons.shield_outlined, color: Color(0xFF2563EB), size: 20),
                      SizedBox(width: 10),
                      Expanded(
                        child: Text(
                          "Syarat: Minimal 6 karakter",
                          style: TextStyle(
                            fontSize: 12,
                            fontWeight: FontWeight.bold,
                            color: Color(0xFF1E40AF),
                          ),
                        ),
                      ),
                    ],
                  ),
                ),

                const SizedBox(height: 20),

                if (_errorMessage != null) ...[
                  Container(
                    padding: const EdgeInsets.all(12),
                    decoration: BoxDecoration(
                      color: const Color(0xFFFEE2E2),
                      borderRadius: BorderRadius.circular(12),
                      border: Border.all(color: const Color(0xFFF87171)),
                    ),
                    child: Row(
                      children: [
                        const Icon(Icons.error_outline_rounded, color: Color(0xFFDC2626), size: 20),
                        const SizedBox(width: 10),
                        Expanded(
                          child: Text(
                            _errorMessage!,
                            style: const TextStyle(color: Color(0xFFDC2626), fontSize: 12.5),
                          ),
                        ),
                      ],
                    ),
                  ),
                  const SizedBox(height: 16),
                ],

                // Input Password Baru
                TextField(
                  controller: _newPasswordController,
                  obscureText: _obscureNew,
                  decoration: InputDecoration(
                    labelText: "Kata Sandi Baru",
                    hintText: "Minimal 6 karakter",
                    prefixIcon: const Icon(Icons.lock_outline_rounded, color: Color(0xFF2563EB)),
                    suffixIcon: IconButton(
                      icon: Icon(
                        _obscureNew ? Icons.visibility_off_rounded : Icons.visibility_rounded,
                        color: const Color(0xFF64748B),
                      ),
                      onPressed: () => setState(() => _obscureNew = !_obscureNew),
                    ),
                    filled: true,
                    fillColor: Colors.white,
                    contentPadding: const EdgeInsets.symmetric(horizontal: 16, vertical: 14),
                    border: OutlineInputBorder(
                      borderRadius: BorderRadius.circular(16),
                      borderSide: const BorderSide(color: Color(0xFFCBD5E1)),
                    ),
                    enabledBorder: OutlineInputBorder(
                      borderRadius: BorderRadius.circular(16),
                      borderSide: const BorderSide(color: Color(0xFFE2E8F0)),
                    ),
                    focusedBorder: OutlineInputBorder(
                      borderRadius: BorderRadius.circular(16),
                      borderSide: const BorderSide(color: Color(0xFF2563EB), width: 1.5),
                    ),
                  ),
                ),

                const SizedBox(height: 16),

                // Input Konfirmasi Password
                TextField(
                  controller: _confirmPasswordController,
                  obscureText: _obscureConfirm,
                  decoration: InputDecoration(
                    labelText: "Ulangi Kata Sandi Baru",
                    hintText: "Ketik ulang kata sandi",
                    prefixIcon: const Icon(Icons.check_circle_outline_rounded, color: Color(0xFF2563EB)),
                    suffixIcon: IconButton(
                      icon: Icon(
                        _obscureConfirm ? Icons.visibility_off_rounded : Icons.visibility_rounded,
                        color: const Color(0xFF64748B),
                      ),
                      onPressed: () => setState(() => _obscureConfirm = !_obscureConfirm),
                    ),
                    filled: true,
                    fillColor: Colors.white,
                    contentPadding: const EdgeInsets.symmetric(horizontal: 16, vertical: 14),
                    border: OutlineInputBorder(
                      borderRadius: BorderRadius.circular(16),
                      borderSide: const BorderSide(color: Color(0xFFCBD5E1)),
                    ),
                    enabledBorder: OutlineInputBorder(
                      borderRadius: BorderRadius.circular(16),
                      borderSide: const BorderSide(color: Color(0xFFE2E8F0)),
                    ),
                    focusedBorder: OutlineInputBorder(
                      borderRadius: BorderRadius.circular(16),
                      borderSide: const BorderSide(color: Color(0xFF2563EB), width: 1.5),
                    ),
                  ),
                ),

                const SizedBox(height: 28),

                // Submit Button
                SizedBox(
                  height: 50,
                  child: ElevatedButton(
                    onPressed: _isLoading ? null : _handleChangePassword,
                    style: ElevatedButton.styleFrom(
                      backgroundColor: const Color(0xFF2563EB),
                      shape: RoundedRectangleBorder(borderRadius: BorderRadius.circular(16)),
                      elevation: 3,
                    ),
                    child: _isLoading
                        ? const SizedBox(
                            width: 22,
                            height: 22,
                            child: CircularProgressIndicator(color: Colors.white, strokeWidth: 2.5),
                          )
                        : const Text(
                            "SIMPAN KATA SANDI & MASUK",
                            style: TextStyle(
                              color: Colors.white,
                              fontWeight: FontWeight.bold,
                              fontSize: 14,
                              letterSpacing: 0.5,
                            ),
                          ),
                  ),
                ),
              ],
            ),
          ),
        ),
      ),
    );
  }
}
