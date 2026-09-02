import 'package:flutter/material.dart';
import '../models/user_model.dart';
import '../core/services/auth_service.dart';
import '../core/utils/custom_toast.dart';
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

  bool _isLoading = false;
  bool _obscureNewPass = true;
  bool _obscureConfirmPass = true;
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

    if (newPass.isEmpty || confirmPass.isEmpty) {
      AppToast.showWarning(context, "Harap isi kata sandi baru dan konfirmasi kata sandi.", title: "Data Belum Lengkap");
      return;
    }

    if (newPass.length < 6) {
      AppToast.showWarning(context, "Kata sandi minimal harus 6 karakter.", title: "Kata Sandi Terlalu Pendek");
      return;
    }

    if (newPass != confirmPass) {
      AppToast.showWarning(context, "Konfirmasi kata sandi tidak cocok dengan kata sandi baru.", title: "Kata Sandi Berbeda");
      return;
    }

    if (newPass == '123456') {
      AppToast.showWarning(context, "Jangan gunakan kata sandi default 123456.", title: "Kata Sandi Tidak Aman");
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
        AppToast.showSuccess(
          context,
          "Kata sandi baru berhasil disimpan! Selamat datang di Waigama.",
          title: "Kata Sandi Berhasil Diubah",
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
      if (mounted) {
        AppToast.showError(context, e.toString(), title: "Gagal Mengubah Kata Sandi");
      }
    } finally {
      if (mounted) {
        setState(() => _isLoading = false);
      }
    }
  }

  @override
  Widget build(BuildContext context) {
    return Scaffold(
      backgroundColor: const Color(0xFF0F172A),
      body: Stack(
        children: [
          // Ambient Gradient Orbs
          Positioned(
            top: -70,
            right: -50,
            child: Container(
              width: 240,
              height: 240,
              decoration: BoxDecoration(
                shape: BoxShape.circle,
                gradient: RadialGradient(
                  colors: [
                    const Color(0xFF2563EB).withOpacity(0.4),
                    Colors.transparent,
                  ],
                ),
              ),
            ),
          ),
          Positioned(
            bottom: -70,
            left: -50,
            child: Container(
              width: 240,
              height: 240,
              decoration: BoxDecoration(
                shape: BoxShape.circle,
                gradient: RadialGradient(
                  colors: [
                    const Color(0xFF38BDF8).withOpacity(0.3),
                    Colors.transparent,
                  ],
                ),
              ),
            ),
          ),

          SafeArea(
            child: Center(
              child: SingleChildScrollView(
                padding: const EdgeInsets.symmetric(horizontal: 24, vertical: 20),
                child: Column(
                  mainAxisAlignment: MainAxisAlignment.center,
                  crossAxisAlignment: CrossAxisAlignment.stretch,
                  children: [
                    // Manta Ray Logo with Cyan Glow
                    Center(
                      child: Container(
                        width: 96,
                        height: 96,
                        decoration: BoxDecoration(
                          color: Colors.white,
                          borderRadius: BorderRadius.circular(28),
                          border: Border.all(color: const Color(0xFF60A5FA), width: 2.5),
                          boxShadow: [
                            BoxShadow(
                              color: const Color(0xFF38BDF8).withOpacity(0.4),
                              blurRadius: 28,
                              offset: const Offset(0, 10),
                            ),
                          ],
                        ),
                        padding: const EdgeInsets.all(10),
                        child: Image.asset(
                          'assets/images/app_logo.png',
                          fit: BoxFit.contain,
                        ),
                      ),
                    ),

                    const SizedBox(height: 20),

                    const Text(
                      "Ganti Kata Sandi",
                      textAlign: TextAlign.center,
                      style: TextStyle(
                        fontSize: 26,
                        fontWeight: FontWeight.w900,
                        color: Colors.white,
                        letterSpacing: -0.3,
                      ),
                    ),

                    const SizedBox(height: 4),

                    Text(
                      "Halo, ${widget.user.name}!\nIni adalah login pertama Anda. Silakan amankan akun Anda.",
                      textAlign: TextAlign.center,
                      style: TextStyle(
                        fontSize: 12.5,
                        color: Colors.white.withOpacity(0.7),
                        height: 1.35,
                      ),
                    ),

                    const SizedBox(height: 24),

                    // Modern Frosted Glass Card
                    Container(
                      decoration: BoxDecoration(
                        color: const Color(0xFF1E293B).withOpacity(0.9),
                        borderRadius: BorderRadius.circular(28),
                        border: Border.all(color: Colors.white.withOpacity(0.12), width: 1.5),
                        boxShadow: [
                          BoxShadow(
                            color: Colors.black.withOpacity(0.35),
                            blurRadius: 24,
                            offset: const Offset(0, 12),
                          ),
                        ],
                      ),
                      padding: const EdgeInsets.fromLTRB(22, 24, 22, 24),
                      child: Column(
                        crossAxisAlignment: CrossAxisAlignment.stretch,
                        children: [
                          // Requirement Notice Pill
                          Container(
                            padding: const EdgeInsets.symmetric(horizontal: 14, vertical: 10),
                            decoration: BoxDecoration(
                              color: const Color(0xFF0F172A),
                              borderRadius: BorderRadius.circular(14),
                              border: Border.all(color: const Color(0xFF38BDF8).withOpacity(0.4)),
                            ),
                            child: Row(
                              children: const [
                                Icon(Icons.info_outline_rounded, color: Color(0xFF38BDF8), size: 18),
                                SizedBox(width: 8),
                                Expanded(
                                  child: Text(
                                    "Syarat kata sandi: Minimal 6 karakter.",
                                    style: TextStyle(fontSize: 11.5, color: Color(0xFF93C5FD), fontWeight: FontWeight.bold),
                                  ),
                                ),
                              ],
                            ),
                          ),

                          const SizedBox(height: 18),

                          if (_errorMessage != null) ...[
                            Container(
                              padding: const EdgeInsets.all(12),
                              decoration: BoxDecoration(
                                color: const Color(0xFF7F1D1D).withOpacity(0.6),
                                borderRadius: BorderRadius.circular(14),
                                border: Border.all(color: const Color(0xFFEF4444)),
                              ),
                              child: Row(
                                children: [
                                  const Icon(Icons.error_outline_rounded, color: Color(0xFFFCA5A5), size: 18),
                                  const SizedBox(width: 8),
                                  Expanded(
                                    child: Text(
                                      _errorMessage!,
                                      style: const TextStyle(color: Color(0xFFFEE2E2), fontSize: 12),
                                    ),
                                  ),
                                ],
                              ),
                            ),
                            const SizedBox(height: 16),
                          ],

                          // New Password Input
                          TextField(
                            controller: _newPasswordController,
                            obscureText: _obscureNewPass,
                            style: const TextStyle(color: Colors.white, fontSize: 14, fontWeight: FontWeight.w500),
                            decoration: InputDecoration(
                              labelText: "Kata Sandi Baru",
                              labelStyle: const TextStyle(color: Color(0xFF94A3B8), fontSize: 13),
                              hintText: "Minimal 6 karakter",
                              hintStyle: const TextStyle(fontSize: 12, color: Color(0xFF64748B)),
                              prefixIcon: const Icon(Icons.lock_outline_rounded, color: Color(0xFF38BDF8), size: 20),
                              suffixIcon: IconButton(
                                icon: Icon(
                                  _obscureNewPass ? Icons.visibility_off_rounded : Icons.visibility_rounded,
                                  color: const Color(0xFF64748B),
                                  size: 20,
                                ),
                                onPressed: () {
                                  setState(() {
                                    _obscureNewPass = !_obscureNewPass;
                                  });
                                },
                              ),
                              filled: true,
                              fillColor: const Color(0xFF0F172A),
                              contentPadding: const EdgeInsets.symmetric(horizontal: 16, vertical: 14),
                              border: OutlineInputBorder(borderRadius: BorderRadius.circular(16), borderSide: const BorderSide(color: Color(0xFF334155))),
                              enabledBorder: OutlineInputBorder(borderRadius: BorderRadius.circular(16), borderSide: const BorderSide(color: Color(0xFF334155))),
                              focusedBorder: OutlineInputBorder(borderRadius: BorderRadius.circular(16), borderSide: const BorderSide(color: Color(0xFF38BDF8), width: 1.8)),
                            ),
                          ),

                          const SizedBox(height: 14),

                          // Confirm Password Input
                          TextField(
                            controller: _confirmPasswordController,
                            obscureText: _obscureConfirmPass,
                            style: const TextStyle(color: Colors.white, fontSize: 14, fontWeight: FontWeight.w500),
                            decoration: InputDecoration(
                              labelText: "Konfirmasi Kata Sandi Baru",
                              labelStyle: const TextStyle(color: Color(0xFF94A3B8), fontSize: 13),
                              hintText: "Ulangi kata sandi baru",
                              hintStyle: const TextStyle(fontSize: 12, color: Color(0xFF64748B)),
                              prefixIcon: const Icon(Icons.lock_reset_rounded, color: Color(0xFF38BDF8), size: 20),
                              suffixIcon: IconButton(
                                icon: Icon(
                                  _obscureConfirmPass ? Icons.visibility_off_rounded : Icons.visibility_rounded,
                                  color: const Color(0xFF64748B),
                                  size: 20,
                                ),
                                onPressed: () {
                                  setState(() {
                                    _obscureConfirmPass = !_obscureConfirmPass;
                                  });
                                },
                              ),
                              filled: true,
                              fillColor: const Color(0xFF0F172A),
                              contentPadding: const EdgeInsets.symmetric(horizontal: 16, vertical: 14),
                              border: OutlineInputBorder(borderRadius: BorderRadius.circular(16), borderSide: const BorderSide(color: Color(0xFF334155))),
                              enabledBorder: OutlineInputBorder(borderRadius: BorderRadius.circular(16), borderSide: const BorderSide(color: Color(0xFF334155))),
                              focusedBorder: OutlineInputBorder(borderRadius: BorderRadius.circular(16), borderSide: const BorderSide(color: Color(0xFF38BDF8), width: 1.8)),
                            ),
                            onSubmitted: (_) => _handleChangePassword(),
                          ),

                          const SizedBox(height: 22),

                          // Submit Button
                          SizedBox(
                            height: 50,
                            child: Container(
                              decoration: BoxDecoration(
                                gradient: const LinearGradient(
                                  colors: [Color(0xFF2563EB), Color(0xFF38BDF8)],
                                  begin: Alignment.centerLeft,
                                  end: Alignment.centerRight,
                                ),
                                borderRadius: BorderRadius.circular(16),
                                boxShadow: [
                                  BoxShadow(
                                    color: const Color(0xFF2563EB).withOpacity(0.4),
                                    blurRadius: 16,
                                    offset: const Offset(0, 6),
                                  ),
                                ],
                              ),
                              child: ElevatedButton(
                                onPressed: _isLoading ? null : _handleChangePassword,
                                style: ElevatedButton.styleFrom(
                                  backgroundColor: Colors.transparent,
                                  shadowColor: Colors.transparent,
                                  shape: RoundedRectangleBorder(borderRadius: BorderRadius.circular(16)),
                                ),
                                child: _isLoading
                                    ? const SizedBox(
                                        width: 22,
                                        height: 22,
                                        child: CircularProgressIndicator(color: Colors.white, strokeWidth: 2.5),
                                      )
                                    : const Text(
                                        "SIMPAN KATA SANDI",
                                        style: TextStyle(
                                          color: Colors.white,
                                          fontSize: 14,
                                          fontWeight: FontWeight.w900,
                                          letterSpacing: 0.8,
                                        ),
                                      ),
                              ),
                            ),
                          ),
                        ],
                      ),
                    ),
                  ],
                ),
              ),
            ),
          ),
        ],
      ),
    );
  }
}
