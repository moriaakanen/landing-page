import 'package:cloud_firestore/cloud_firestore.dart';
import 'package:flutter/material.dart';
import 'package:flutter_local_notifications/flutter_local_notifications.dart';
import '../utils/custom_toast.dart';

class NotificationService {
  static final FirebaseFirestore _firestore = FirebaseFirestore.instance;
  static final FlutterLocalNotificationsPlugin _localNotifications = FlutterLocalNotificationsPlugin();
  static bool _isInitialized = false;

  /// Inisialisasi Android System Notifications & Request Permission
  static Future<void> initialize() async {
    if (_isInitialized) return;

    const AndroidInitializationSettings androidSettings =
        AndroidInitializationSettings('@drawable/ic_notification');
    const InitializationSettings initSettings = InitializationSettings(
      android: androidSettings,
    );

    await _localNotifications.initialize(
      initSettings,
      onDidReceiveNotificationResponse: (NotificationResponse response) {
        debugPrint("Notification clicked with payload: ${response.payload}");
      },
    );

    // Request Runtime Permission untuk Android 13+ (API 33+)
    final androidImplementation =
        _localNotifications.resolvePlatformSpecificImplementation<AndroidFlutterLocalNotificationsPlugin>();
    if (androidImplementation != null) {
      await androidImplementation.requestNotificationsPermission();
    }

    _isInitialized = true;
  }

  /// Menampilkan Notifikasi Asli Sistem Android (Status Bar, Heads-up banner, Lockscreen)
  /// dengan Small Icon monokrom dan Large Icon (Logo Ikan Pari Biru berwarna)
  static Future<void> showSystemNotification({
    int id = 0,
    required String title,
    required String body,
    String? payload,
  }) async {
    await initialize();

    const AndroidNotificationDetails androidDetails = AndroidNotificationDetails(
      'waigama_channel_id',
      'Notifikasi Waigama',
      channelDescription: 'Pemberitahuan izin kegiatan, laporan Cepu, dan status kehadiran',
      importance: Importance.max,
      priority: Priority.high,
      showWhen: true,
      icon: '@drawable/ic_notification',
      largeIcon: DrawableResourceAndroidBitmap('@mipmap/ic_launcher'),
      color: Color(0xFF1E60F2),
      enableVibration: true,
      playSound: true,
    );

    const NotificationDetails notificationDetails = NotificationDetails(
      android: androidDetails,
    );

    await _localNotifications.show(
      id,
      title,
      body,
      notificationDetails,
      payload: payload,
    );
  }

  /// Stream notifikasi broadcast untuk laporan Cepu baru
  static Stream<QuerySnapshot> streamRecentCepuNotifications() {
    final tenMinutesAgo = DateTime.now().subtract(const Duration(minutes: 60));
    return _firestore
        .collection('notifications')
        .where('type', isEqualTo: 'CEPU_NEW')
        .where('created_at', isGreaterThan: Timestamp.fromDate(tenMinutesAgo))
        .snapshots();
  }

  /// Menampilkan In-App Top Notification Banner / Floating Toast
  static void showInAppAlert(BuildContext context, {
    required String title,
    required String message,
    Color backgroundColor = const Color(0xFF1E293B),
    IconData icon = Icons.notifications_active_rounded,
    VoidCallback? onTap,
  }) {
    // Tampilkan juga system status bar notification secara bersamaan
    showSystemNotification(
      id: DateTime.now().millisecondsSinceEpoch ~/ 1000,
      title: title,
      body: message,
    );

    AppToast.showWarning(context, message, title: title);
  }
}
