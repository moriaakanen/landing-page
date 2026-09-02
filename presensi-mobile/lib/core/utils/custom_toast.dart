import 'package:flutter/material.dart';

class AppToast {
  static void showSuccess(
    BuildContext context,
    String message, {
    String title = "Sukses",
    Duration duration = const Duration(seconds: 4),
  }) {
    _showCustomToast(
      context,
      title: title,
      message: message,
      icon: Icons.check_circle_rounded,
      primaryColor: const Color(0xFF10B981),
      bgGradient: const [Color(0xFF064E3B), Color(0xFF065F46)],
      borderColor: const Color(0xFF34D399),
      duration: duration,
    );
  }

  static void showError(
    BuildContext context,
    String message, {
    String title = "Perhatian",
    Duration duration = const Duration(seconds: 4),
  }) {
    _showCustomToast(
      context,
      title: title,
      message: message,
      icon: Icons.error_rounded,
      primaryColor: const Color(0xFFEF4444),
      bgGradient: const [Color(0xFF7F1D1D), Color(0xFF991B1B)],
      borderColor: const Color(0xFFF87171),
      duration: duration,
    );
  }

  static void showWarning(
    BuildContext context,
    String message, {
    String title = "Peringatan",
    Duration duration = const Duration(seconds: 4),
  }) {
    _showCustomToast(
      context,
      title: title,
      message: message,
      icon: Icons.warning_amber_rounded,
      primaryColor: const Color(0xFFF59E0B),
      bgGradient: const [Color(0xFF78350F), Color(0xFF92400E)],
      borderColor: const Color(0xFFFBBF24),
      duration: duration,
    );
  }

  static void showInfo(
    BuildContext context,
    String message, {
    String title = "Informasi",
    Duration duration = const Duration(seconds: 4),
  }) {
    _showCustomToast(
      context,
      title: title,
      message: message,
      icon: Icons.info_outline_rounded,
      primaryColor: const Color(0xFF3B82F6),
      bgGradient: const [Color(0xFF1E3A8A), Color(0xFF1E40AF)],
      borderColor: const Color(0xFF60A5FA),
      duration: duration,
    );
  }

  static void _showCustomToast(
    BuildContext context, {
    required String title,
    required String message,
    required IconData icon,
    required Color primaryColor,
    required List<Color> bgGradient,
    required Color borderColor,
    required Duration duration,
  }) {
    ScaffoldMessenger.of(context).hideCurrentSnackBar();
    ScaffoldMessenger.of(context).showSnackBar(
      SnackBar(
        elevation: 8,
        behavior: SnackBarBehavior.floating,
        margin: const EdgeInsets.fromLTRB(16, 0, 16, 18),
        padding: EdgeInsets.zero,
        backgroundColor: Colors.transparent,
        duration: duration,
        content: Container(
          padding: const EdgeInsets.symmetric(horizontal: 16, vertical: 14),
          decoration: BoxDecoration(
            gradient: LinearGradient(
              colors: bgGradient,
              begin: Alignment.topLeft,
              end: Alignment.bottomRight,
            ),
            borderRadius: BorderRadius.circular(18),
            border: Border.all(color: borderColor.withOpacity(0.4), width: 1.5),
            boxShadow: [
              BoxShadow(
                color: bgGradient.first.withOpacity(0.4),
                blurRadius: 18,
                offset: const Offset(0, 6),
              ),
            ],
          ),
          child: Row(
            children: [
              Container(
                width: 38,
                height: 38,
                decoration: BoxDecoration(
                  color: primaryColor.withOpacity(0.2),
                  shape: BoxShape.circle,
                  border: Border.all(color: primaryColor.withOpacity(0.5), width: 1.5),
                ),
                child: Icon(icon, color: Colors.white, size: 22),
              ),
              const SizedBox(width: 14),
              Expanded(
                child: Column(
                  mainAxisSize: MainAxisSize.min,
                  crossAxisAlignment: CrossAxisAlignment.start,
                  children: [
                    Text(
                      title,
                      style: const TextStyle(
                        color: Colors.white,
                        fontWeight: FontWeight.w800,
                        fontSize: 13.5,
                        letterSpacing: 0.2,
                      ),
                    ),
                    const SizedBox(height: 2),
                    Text(
                      message,
                      style: TextStyle(
                        color: Colors.white.withOpacity(0.9),
                        fontSize: 12,
                        height: 1.3,
                        fontWeight: FontWeight.w400,
                      ),
                    ),
                  ],
                ),
              ),
            ],
          ),
        ),
      ),
    );
  }
}
