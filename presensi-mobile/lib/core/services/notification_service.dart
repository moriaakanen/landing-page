import 'package:cloud_firestore/cloud_firestore.dart';
import 'package:flutter/material.dart';

class NotificationService {
  static final FirebaseFirestore _firestore = FirebaseFirestore.instance;

  /// Stream notifikasi broadcast untuk laporan Cepu baru
  static Stream<QuerySnapshot> streamRecentCepuNotifications() {
    final tenMinutesAgo = DateTime.now().subtract(const Duration(minutes: 60));
    return _firestore
        .collection('notifications')
        .where('type', isEqualTo: 'CEPU_NEW')
        .where('created_at', isGreaterThan: Timestamp.fromDate(tenMinutesAgo))
        .snapshots();
  }

  /// Menampilkan In-App Top Notification Banner
  static void showInAppAlert(BuildContext context, {
    required String title,
    required String message,
    Color backgroundColor = const Color(0xFF1E293B),
    IconData icon = Icons.notifications_active_rounded,
    VoidCallback? onTap,
  }) {
    ScaffoldMessenger.of(context).showSnackBar(
      SnackBar(
        elevation: 6,
        behavior: SnackBarBehavior.floating,
        margin: const EdgeInsets.fromLTRB(16, 0, 16, 20),
        padding: const EdgeInsets.symmetric(horizontal: 16, vertical: 12),
        backgroundColor: backgroundColor,
        shape: RoundedRectangleBorder(borderRadius: BorderRadius.circular(16)),
        content: Row(
          children: [
            Icon(icon, color: Colors.white, size: 24),
            const SizedBox(width: 12),
            Expanded(
              child: Column(
                mainAxisSize: MainAxisSize.min,
                crossAxisAlignment: CrossAxisAlignment.start,
                children: [
                  Text(
                    title,
                    style: const TextStyle(
                      color: Colors.white,
                      fontWeight: FontWeight.bold,
                      fontSize: 13,
                    ),
                  ),
                  const SizedBox(height: 2),
                  Text(
                    message,
                    style: const TextStyle(
                      color: Color(0xFFE2E8F0),
                      fontSize: 11.5,
                      height: 1.3,
                    ),
                  ),
                ],
              ),
            ),
          ],
        ),
        action: onTap != null
            ? SnackBarAction(
                label: "LIHAT",
                textColor: const Color(0xFFFBBF24),
                onPressed: onTap,
              )
            : null,
        duration: const Duration(seconds: 5),
      ),
    );
  }
}
