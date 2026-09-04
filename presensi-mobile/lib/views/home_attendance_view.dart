import 'dart:async';
import 'dart:math';
import 'package:flutter/material.dart';
import 'package:intl/intl.dart';
import 'package:geolocator/geolocator.dart';
import '../models/user_model.dart';
import '../models/office_model.dart';
import '../models/permit_model.dart';
import '../models/cepu_model.dart';
import '../core/services/attendance_service.dart';
import '../core/services/permit_service.dart';
import '../core/services/cepu_service.dart';
import '../core/services/auth_service.dart';
import '../core/services/location_service.dart';
import '../core/services/notification_service.dart';
import '../core/utils/custom_toast.dart';
import 'login_view.dart';
import 'record_permit_map_view.dart';
import 'cepu_map_report_view.dart';
import 'daily_monitoring_view.dart';
import 'analytics_view.dart';
import 'settings_view.dart';

/// Item aktivitas harian user (Izin Waigama maupun Yang Di-Cepukan Valid)
class UserActivityItem {
  final String id;
  final String type; // 'PERMIT' atau 'CEPU'
  final String title;
  final String purposeOrDesc;
  final DateTime startTime;
  final DateTime? endTime;
  final bool isActive;
  final String? reporterName;
  final CepuModel? cepuModel;
  final PermitModel? permitModel;

  UserActivityItem({
    required this.id,
    required this.type,
    required this.title,
    required this.purposeOrDesc,
    required this.startTime,
    this.endTime,
    required this.isActive,
    this.reporterName,
    this.cepuModel,
    this.permitModel,
  });

  int getDurationMinutes() {
    final now = DateTime.now();
    final isFriday = now.weekday == DateTime.friday;
    final cutoffHour = 16;
    final cutoffMinute = isFriday ? 30 : 0;
    final cutoffTime = DateTime(now.year, now.month, now.day, cutoffHour, cutoffMinute);

    DateTime effectiveEnd;
    if (endTime != null) {
      effectiveEnd = endTime!;
    } else {
      effectiveEnd = now.isAfter(cutoffTime) ? cutoffTime : now;
    }

    if (effectiveEnd.isBefore(startTime)) return 0;
    return effectiveEnd.difference(startTime).inMinutes;
  }
}

class HomeAttendanceView extends StatefulWidget {
  final UserModel user;

  const HomeAttendanceView({Key? key, required this.user}) : super(key: key);

  @override
  State<HomeAttendanceView> createState() => _HomeAttendanceViewState();
}

class _HomeAttendanceViewState extends State<HomeAttendanceView> {
  final _attendanceService = AttendanceService();
  final _permitService = PermitService();
  final _cepuService = CepuService();
  final _authService = AuthService();

  OfficeModel? _office;
  PermitModel? _activePermit;
  CepuModel? _activeCepuForMe;
  List<UserActivityItem> _todayActivities = [];
  List<CepuModel> _pendingVerificationCepu = [];

  int _currentActivityPage = 0;
  static const int _activityPageSize = 3;

  bool _isLoading = true;
  Timer? _clockTimer;
  DateTime _currentTime = DateTime.now();

  // GPS Location Status
  bool _isCheckingLocation = true;
  bool _isInOfficeRadius = false;

  StreamSubscription? _cepuNotifSub;
  final Set<String> _seenNotifIds = {};

  @override
  void initState() {
    super.initState();
    _loadInitialData();
    _listenToNewCepuNotifications();
    _startClockTimer();
  }

  void _startClockTimer() {
    _clockTimer = Timer.periodic(const Duration(seconds: 1), (timer) {
      if (mounted) {
        setState(() {
          _currentTime = DateTime.now();
        });
      }
    });
  }

  @override
  void dispose() {
    _clockTimer?.cancel();
    _cepuNotifSub?.cancel();
    super.dispose();
  }

  Future<void> _loadInitialData() async {
    setState(() {
      _isLoading = true;
      _isCheckingLocation = true;
    });

    final todayStr = DateFormat('yyyy-MM-dd').format(DateTime.now());
    final office = await _attendanceService.getOfficeConfig(widget.user.officeId);
    final activePermit = await _permitService.getActivePermit(widget.user.uid);
    final activeCepu = await _cepuService.getActiveCepuForTarget(widget.user.uid);

    // 1. Ambil seluruh izin user hari ini dari database
    final userPermits = await _permitService.getUserDailyPermits(widget.user.uid, todayStr);

    // 2. Ambil seluruh laporan Cepu valid untuk user hari ini dari database
    final userValidCepu = await _cepuService.getDailyValidCepuForTarget(widget.user.uid, todayStr);

    // 3. Ambil laporan Cepu pending rekan yang belum diverifikasi
    final pendingCepu = await _cepuService.getPendingVerificationCepuForUser(widget.user.uid, todayStr);

    // Gabungkan menjadi list aktivitas harian user
    final List<UserActivityItem> activities = [];
    for (final p in userPermits) {
      activities.add(UserActivityItem(
        id: p.id,
        type: 'PERMIT',
        title: 'Izin Waigama',
        purposeOrDesc: p.purpose,
        startTime: p.startTime,
        endTime: p.endTime,
        isActive: p.isActive,
        permitModel: p,
      ));
    }

    for (final c in userValidCepu) {
      activities.add(UserActivityItem(
        id: c.id,
        type: 'CEPU',
        title: 'Laporan Cepu Terverifikasi',
        purposeOrDesc: c.description,
        startTime: c.startTime,
        endTime: c.endTime,
        isActive: c.isActive,
        reporterName: c.reporterName,
        cepuModel: c,
      ));
    }

    // Urutkan dari aktivitas terbaru
    activities.sort((a, b) => b.startTime.compareTo(a.startTime));

    if (mounted) {
      setState(() {
        _office = office;
        _activePermit = activePermit;
        _activeCepuForMe = activeCepu;
        _todayActivities = activities;
        _currentActivityPage = 0;
        _pendingVerificationCepu = pendingCepu;
        _isLoading = false;
      });

      _checkCurrentLocationStatus(office);

      if (activePermit != null) {
        NotificationService.showSystemNotification(
          id: 101,
          title: "Anda sedang berada di luar kantor",
          body: "Segera rekam waktu kembali setelah tiba di kantor",
        );
      }
      if (activeCepu != null) {
        NotificationService.showSystemNotification(
          id: 102,
          title: "⚠️ Peringatan Laporan Cepu",
          body: "Anda dilaporkan tidak berada di kantor, segera rekam waktu kembali jika anda berada di kantor",
        );
      }
    }
  }

  Future<void> _checkCurrentLocationStatus(OfficeModel? office) async {
    if (office == null) {
      if (mounted) setState(() => _isCheckingLocation = false);
      return;
    }

    try {
      final hasPermission = await LocationService.handleLocationPermission();
      if (!hasPermission) {
        if (mounted) setState(() => _isCheckingLocation = false);
        return;
      }

      final position = await Geolocator.getCurrentPosition(
        desiredAccuracy: LocationAccuracy.medium,
        timeLimit: const Duration(seconds: 10),
      );

      final distance = LocationService.calculateDistance(
        startLatitude: position.latitude,
        startLongitude: position.longitude,
        endLatitude: office.latitude,
        endLongitude: office.longitude,
      );

      if (mounted) {
        setState(() {
          _isInOfficeRadius = distance <= office.radiusMeters;
          _isCheckingLocation = false;
        });
      }
    } catch (_) {
      if (mounted) {
        setState(() => _isCheckingLocation = false);
      }
    }
  }

  void _listenToNewCepuNotifications() {
    bool isFirstFetch = true;
    _cepuNotifSub = NotificationService.streamRecentCepuNotifications().listen((snapshot) {
      if (isFirstFetch) {
        for (final doc in snapshot.docs) {
          _seenNotifIds.add(doc.id);
        }
        isFirstFetch = false;
        return;
      }

      for (final doc in snapshot.docs) {
        if (!_seenNotifIds.contains(doc.id)) {
          _seenNotifIds.add(doc.id);
          final data = doc.data() as Map<String, dynamic>;
          final String title = data['title'] ?? '🚨 Laporan Cepu Baru';
          final String message = data['message'] ?? 'Ada laporan baru yang membutuhkan verifikasi rekan.';
          final String targetUid = data['target_uid'] ?? '';

          if (mounted && targetUid != widget.user.uid) {
            NotificationService.showInAppAlert(
              context,
              title: title,
              message: message,
              backgroundColor: const Color(0xFFEA580C),
              icon: Icons.campaign_rounded,
              onTap: _navigateToDailyMonitoring,
            );
          }
        }
      }
    });
  }

  Future<void> _handleLogout() async {
    final confirm = await showDialog<bool>(
      context: context,
      builder: (ctx) => AlertDialog(
        shape: RoundedRectangleBorder(borderRadius: BorderRadius.circular(22)),
        title: const Text("Konfirmasi Keluar", style: TextStyle(fontWeight: FontWeight.bold, fontSize: 16.5)),
        content: const Text("Apakah Anda yakin ingin keluar dari aplikasi Waigama?"),
        actions: [
          TextButton(
            onPressed: () => Navigator.pop(ctx, false),
            child: const Text("Batal", style: TextStyle(color: Color(0xFF64748B), fontWeight: FontWeight.w600)),
          ),
          ElevatedButton(
            onPressed: () => Navigator.pop(ctx, true),
            style: ElevatedButton.styleFrom(
              backgroundColor: const Color(0xFFEF4444),
              shape: RoundedRectangleBorder(borderRadius: BorderRadius.circular(12)),
              elevation: 0,
            ),
            child: const Text("Keluar", style: TextStyle(color: Colors.white, fontWeight: FontWeight.bold)),
          ),
        ],
      ),
    );

    if (confirm == true) {
      await _authService.logout();
      if (mounted) {
        AppToast.showInfo(context, "Sampai jumpa kembali!", title: "Berhasil Keluar");
        Navigator.pushReplacement(
          context,
          MaterialPageRoute(builder: (context) => const LoginView()),
        );
      }
    }
  }

  void _navigateToRecordPermit() async {
    if (_office == null) return;
    final result = await Navigator.push(
      context,
      MaterialPageRoute(
        builder: (context) => RecordPermitMapView(
          user: widget.user,
          office: _office!,
          initialActivePermit: _activePermit,
        ),
      ),
    );

    if (result == true) {
      _loadInitialData();
    }
  }

  void _navigateToCepuMapReport() async {
    if (_office == null) return;
    final result = await Navigator.push(
      context,
      MaterialPageRoute(
        builder: (context) => CepuMapReportView(
          reporter: widget.user,
          office: _office!,
        ),
      ),
    );

    if (result == true) {
      _loadInitialData();
    }
  }

  void _navigateToDailyMonitoring() {
    Navigator.push(
      context,
      MaterialPageRoute(
        builder: (context) => DailyMonitoringView(
          user: widget.user,
          office: _office,
        ),
      ),
    );
  }

  void _navigateToAnalytics() {
    Navigator.push(
      context,
      MaterialPageRoute(
        builder: (context) => AnalyticsView(
          user: widget.user,
          office: _office,
        ),
      ),
    );
  }

  void _navigateToSettings() {
    Navigator.push(
      context,
      MaterialPageRoute(
        builder: (context) => SettingsView(
          user: widget.user,
          office: _office,
        ),
      ),
    );
  }

  void _showProfileDialog() {
    showDialog(
      context: context,
      builder: (ctx) => AlertDialog(
        shape: RoundedRectangleBorder(borderRadius: BorderRadius.circular(24)),
        contentPadding: const EdgeInsets.all(22),
        content: Column(
          mainAxisSize: MainAxisSize.min,
          children: [
            Container(
              width: 72,
              height: 72,
              decoration: BoxDecoration(
                color: const Color(0xFFEFF6FF),
                shape: BoxShape.circle,
                border: Border.all(color: const Color(0xFF1E60F2), width: 2),
              ),
              padding: const EdgeInsets.all(10),
              child: Image.asset('assets/images/app_logo.png', fit: BoxFit.contain),
            ),
            const SizedBox(height: 14),
            Text(
              widget.user.name,
              style: const TextStyle(fontSize: 18, fontWeight: FontWeight.bold, color: Color(0xFF1E293B)),
              textAlign: TextAlign.center,
            ),
            const SizedBox(height: 4),
            Text(
              "@${widget.user.username}",
              style: const TextStyle(fontSize: 13, color: Color(0xFF64748B), fontWeight: FontWeight.w500),
            ),
            const SizedBox(height: 16),
            const Divider(height: 1, color: Color(0xFFF1F5F9)),
            const SizedBox(height: 14),
            Row(
              mainAxisAlignment: MainAxisAlignment.spaceBetween,
              children: [
                const Text("Status Akun", style: TextStyle(fontSize: 12.5, color: Color(0xFF64748B))),
                Container(
                  padding: const EdgeInsets.symmetric(horizontal: 8, vertical: 3),
                  decoration: BoxDecoration(
                    color: const Color(0xFFEFF6FF),
                    borderRadius: BorderRadius.circular(8),
                  ),
                  child: const Text("Pegawai Aktif", style: TextStyle(fontSize: 11, fontWeight: FontWeight.bold, color: Color(0xFF1E60F2))),
                ),
              ],
            ),
            const SizedBox(height: 20),
            SizedBox(
              width: double.infinity,
              height: 44,
              child: ElevatedButton.icon(
                onPressed: () {
                  Navigator.pop(ctx);
                  _handleLogout();
                },
                icon: const Icon(Icons.logout_rounded, size: 18, color: Colors.white),
                label: const Text("Keluar Akun", style: TextStyle(color: Colors.white, fontWeight: FontWeight.bold)),
                style: ElevatedButton.styleFrom(
                  backgroundColor: const Color(0xFFEF4444),
                  shape: RoundedRectangleBorder(borderRadius: BorderRadius.circular(14)),
                  elevation: 0,
                ),
              ),
            ),
          ],
        ),
      ),
    );
  }

  /// POIN 6: Modal Pop-up Notifikasi & Pengingat
  /// Berisi laporan cepu yang belum diverifikasi dan izin user yang belum kembali
  void _showNotificationsDialog() {
    showModalBottomSheet(
      context: context,
      backgroundColor: Colors.transparent,
      isScrollControlled: true,
      builder: (ctx) {
        final hasActivePermitNotice = _activePermit != null;
        final hasActiveCepuNotice = _activeCepuForMe != null;
        final hasPendingVerifyNotice = _pendingVerificationCepu.isNotEmpty;
        final hasAnyNotification = hasActivePermitNotice || hasActiveCepuNotice || hasPendingVerifyNotice;

        return Container(
          constraints: BoxConstraints(
            maxHeight: MediaQuery.of(context).size.height * 0.75,
          ),
          decoration: const BoxDecoration(
            color: Colors.white,
            borderRadius: BorderRadius.vertical(top: Radius.circular(28)),
          ),
          padding: const EdgeInsets.fromLTRB(20, 16, 20, 24),
          child: Column(
            mainAxisSize: MainAxisSize.min,
            crossAxisAlignment: CrossAxisAlignment.start,
            children: [
              // Drag Indicator Bar
              Center(
                child: Container(
                  width: 40,
                  height: 4,
                  decoration: BoxDecoration(
                    color: const Color(0xFFCBD5E1),
                    borderRadius: BorderRadius.circular(10),
                  ),
                ),
              ),
              const SizedBox(height: 16),

              // Title Row
              Row(
                mainAxisAlignment: MainAxisAlignment.spaceBetween,
                children: [
                  Row(
                    children: const [
                      Icon(Icons.notifications_active_rounded, color: Color(0xFF1E60F2), size: 22),
                      SizedBox(width: 8),
                      Text(
                        "Notifikasi & Pengingat",
                        style: TextStyle(
                          fontSize: 18,
                          fontWeight: FontWeight.bold,
                          color: Color(0xFF0F172A),
                        ),
                      ),
                    ],
                  ),
                  IconButton(
                    icon: const Icon(Icons.close_rounded, color: Color(0xFF64748B)),
                    onPressed: () => Navigator.pop(ctx),
                    padding: EdgeInsets.zero,
                    constraints: const BoxConstraints(),
                  ),
                ],
              ),

              const SizedBox(height: 16),
              const Divider(height: 1, color: Color(0xFFF1F5F9)),
              const SizedBox(height: 16),

              // Content
              if (!hasAnyNotification)
                Container(
                  width: double.infinity,
                  padding: const EdgeInsets.symmetric(vertical: 40),
                  child: Column(
                    children: [
                      Container(
                        width: 56,
                        height: 56,
                        decoration: const BoxDecoration(
                          color: Color(0xFFF1F5F9),
                          shape: BoxShape.circle,
                        ),
                        child: const Icon(Icons.notifications_off_outlined, color: Color(0xFF94A3B8), size: 30),
                      ),
                      const SizedBox(height: 14),
                      const Text(
                        "Tidak ada notifikasi",
                        style: TextStyle(
                          fontSize: 16,
                          fontWeight: FontWeight.bold,
                          color: Color(0xFF334155),
                        ),
                      ),
                      const SizedBox(height: 4),
                      const Text(
                        "Semua laporan dan izin Anda telah diperbarui.",
                        style: TextStyle(fontSize: 12.5, color: Color(0xFF94A3B8)),
                        textAlign: TextAlign.center,
                      ),
                    ],
                  ),
                )
              else
                Flexible(
                  child: ListView(
                    shrinkWrap: true,
                    children: [
                      // 1. Notifikasi Izin User Belum Kembali
                      if (hasActivePermitNotice) ...[
                        _buildNotificationCard(
                          icon: Icons.directions_walk_rounded,
                          iconColor: const Color(0xFFD97706),
                          iconBgColor: const Color(0xFFFEF3C7),
                          badgeText: "IZIN AKTIF",
                          badgeColor: const Color(0xFFFEF3C7),
                          badgeTextColor: const Color(0xFFB45309),
                          title: "Izin Belum Kembali",
                          subtitle: "Anda sedang izin untuk keperluan: ${_activePermit!.purpose}. Segera rekam waktu kembali setelah tiba di kantor.",
                          actionText: "Rekam Waktu Kembali",
                          onTapAction: () {
                            Navigator.pop(ctx);
                            _navigateToRecordPermit();
                          },
                        ),
                        const SizedBox(height: 12),
                      ],

                      // 2. Notifikasi Laporan Cepu Aktif Untuk User Belum Kembali
                      if (hasActiveCepuNotice) ...[
                        _buildNotificationCard(
                          icon: Icons.warning_amber_rounded,
                          iconColor: const Color(0xFFDC2626),
                          iconBgColor: const Color(0xFFFEE2E2),
                          badgeText: "CEPU AKTIF",
                          badgeColor: const Color(0xFFFEE2E2),
                          badgeTextColor: const Color(0xFF991B1B),
                          title: "Laporan Keberadaan (Cepu)",
                          subtitle: "Anda dilaporkan oleh ${_activeCepuForMe!.reporterName} tidak berada di kantor. Segera rekam waktu kembali.",
                          actionText: "Rekam Waktu Kembali",
                          onTapAction: () {
                            Navigator.pop(ctx);
                            _navigateToCepuMapReport();
                          },
                        ),
                        const SizedBox(height: 12),
                      ],

                      // 3. Notifikasi Laporan Cepu Rekan yang Belum Diverifikasi
                      for (final cepu in _pendingVerificationCepu) ...[
                        _buildNotificationCard(
                          icon: Icons.how_to_reg_rounded,
                          iconColor: const Color(0xFF1E60F2),
                          iconBgColor: const Color(0xFFEFF6FF),
                          badgeText: "VERIFIKASI CEPU",
                          badgeColor: const Color(0xFFEFF6FF),
                          badgeTextColor: const Color(0xFF1E60F2),
                          title: "Verifikasi Keberadaan: ${cepu.targetName}",
                          subtitle: "Dilaporkan oleh ${cepu.reporterName}: \"${cepu.description}\". Progres: ${cepu.verifiedByUids.length}/4 rekan telah verifikasi.",
                          actionText: "Buka Monitoring & Verifikasi",
                          onTapAction: () {
                            Navigator.pop(ctx);
                            _navigateToDailyMonitoring();
                          },
                        ),
                        const SizedBox(height: 12),
                      ],
                    ],
                  ),
                ),
            ],
          ),
        );
      },
    );
  }

  Widget _buildNotificationCard({
    required IconData icon,
    required Color iconColor,
    required Color iconBgColor,
    required String badgeText,
    required Color badgeColor,
    required Color badgeTextColor,
    required String title,
    required String subtitle,
    required String actionText,
    required VoidCallback onTapAction,
  }) {
    return Container(
      padding: const EdgeInsets.all(14),
      decoration: BoxDecoration(
        color: const Color(0xFFF8FAFC),
        borderRadius: BorderRadius.circular(18),
        border: Border.all(color: const Color(0xFFE2E8F0)),
      ),
      child: Column(
        crossAxisAlignment: CrossAxisAlignment.start,
        children: [
          Row(
            crossAxisAlignment: CrossAxisAlignment.start,
            children: [
              Container(
                width: 38,
                height: 38,
                decoration: BoxDecoration(
                  color: iconBgColor,
                  borderRadius: BorderRadius.circular(12),
                ),
                child: Icon(icon, color: iconColor, size: 20),
              ),
              const SizedBox(width: 12),
              Expanded(
                child: Column(
                  crossAxisAlignment: CrossAxisAlignment.start,
                  children: [
                    Row(
                      children: [
                        Container(
                          padding: const EdgeInsets.symmetric(horizontal: 6, vertical: 2),
                          decoration: BoxDecoration(
                            color: badgeColor,
                            borderRadius: BorderRadius.circular(5),
                          ),
                          child: Text(
                            badgeText,
                            style: TextStyle(
                              fontSize: 9.5,
                              fontWeight: FontWeight.bold,
                              color: badgeTextColor,
                            ),
                          ),
                        ),
                      ],
                    ),
                    const SizedBox(height: 4),
                    Text(
                      title,
                      style: const TextStyle(
                        fontSize: 13.5,
                        fontWeight: FontWeight.bold,
                        color: Color(0xFF0F172A),
                      ),
                    ),
                    const SizedBox(height: 3),
                    Text(
                      subtitle,
                      style: const TextStyle(
                        fontSize: 12,
                        color: Color(0xFF64748B),
                        height: 1.3,
                      ),
                    ),
                  ],
                ),
              ),
            ],
          ),
          const SizedBox(height: 10),
          Align(
            alignment: Alignment.centerRight,
            child: TextButton.icon(
              onPressed: onTapAction,
              icon: const Icon(Icons.arrow_forward_rounded, size: 14, color: Color(0xFF1E60F2)),
              label: Text(
                actionText,
                style: const TextStyle(
                  fontSize: 12,
                  fontWeight: FontWeight.bold,
                  color: Color(0xFF1E60F2),
                ),
              ),
              style: TextButton.styleFrom(
                padding: const EdgeInsets.symmetric(horizontal: 10, vertical: 4),
                backgroundColor: const Color(0xFFEFF6FF),
                shape: RoundedRectangleBorder(borderRadius: BorderRadius.circular(8)),
              ),
            ),
          ),
        ],
      ),
    );
  }

  String _getGreeting() {
    final hour = DateTime.now().hour;
    if (hour < 11) return "Selamat Pagi,";
    if (hour < 15) return "Selamat Siang,";
    if (hour < 18) return "Selamat Sore,";
    return "Selamat Malam,";
  }

  /// POIN 5: Logic pemecahan nama user
  /// - Jika nama > 1 kata: dipecah jadi 2 baris seimbang
  /// - Jika nama 1 kata: tidak dipecah
  List<String> _splitUserName(String fullName) {
    final trimmed = fullName.trim();
    final words = trimmed.split(RegExp(r'\s+'));
    if (words.length <= 1) {
      return [trimmed, ""];
    }
    final half = (words.length / 2).ceil();
    final line1 = words.sublist(0, half).join(' ');
    final line2 = words.sublist(half).join(' ');
    return [line1, line2];
  }

  @override
  Widget build(BuildContext context) {
    String dayNameFormatted;
    String dateFormatted;

    try {
      dayNameFormatted = DateFormat('EEEE', 'id_ID').format(DateTime.now());
      dateFormatted = DateFormat('dd MMMM yyyy', 'id_ID').format(DateTime.now());
    } catch (_) {
      dayNameFormatted = DateFormat('EEEE').format(DateTime.now());
      dateFormatted = DateFormat('dd MMMM yyyy').format(DateTime.now());
    }

    final hasActivePermit = _activePermit != null;
    final startPermitTime = hasActivePermit
        ? "${DateFormat('HH:mm').format(_activePermit!.startTime)} WIT"
        : "--:--";
    final endPermitTime = (_activePermit != null && _activePermit!.endTime != null)
        ? "${DateFormat('HH:mm').format(_activePermit!.endTime!)} WIT"
        : "--:--";
    final timeFormatted = DateFormat('HH:mm').format(_currentTime);

    final nameParts = _splitUserName(widget.user.name);
    final hasNotificationBadge = hasActivePermit || _activeCepuForMe != null || _pendingVerificationCepu.isNotEmpty;

    return Scaffold(
      backgroundColor: const Color(0xFFF8FAFC),
      body: SafeArea(
        bottom: false,
        child: RefreshIndicator(
          onRefresh: _loadInitialData,
          color: const Color(0xFF1E60F2),
          child: SingleChildScrollView(
            physics: const AlwaysScrollableScrollPhysics(),
            padding: const EdgeInsets.fromLTRB(20, 16, 20, 30),
            child: Column(
              crossAxisAlignment: CrossAxisAlignment.start,
              children: [
                // =========================================================================
                // TOPBAR: Logo + "Waigama" & Icon Notifikasi
                // =========================================================================
                Row(
                  mainAxisAlignment: MainAxisAlignment.spaceBetween,
                  children: [
                    // Pojok Kiri: Logo Ikan Pari + Tulisan "Waigama"
                    Container(
                      padding: const EdgeInsets.symmetric(horizontal: 10, vertical: 6),
                      decoration: BoxDecoration(
                        color: Colors.white,
                        borderRadius: BorderRadius.circular(16),
                        border: Border.all(color: const Color(0xFFE2E8F0), width: 1.2),
                        boxShadow: [
                          BoxShadow(
                            color: const Color(0xFF1E60F2).withOpacity(0.06),
                            blurRadius: 10,
                            offset: const Offset(0, 3),
                          ),
                        ],
                      ),
                      child: Row(
                        mainAxisSize: MainAxisSize.min,
                        children: [
                          Image.asset(
                            'assets/images/app_logo.png',
                            width: 32,
                            height: 32,
                            fit: BoxFit.contain,
                          ),
                          const SizedBox(width: 8),
                          RichText(
                            text: const TextSpan(
                              children: [
                                TextSpan(
                                  text: "Wai",
                                  style: TextStyle(
                                    fontSize: 20,
                                    fontWeight: FontWeight.w900,
                                    color: Color(0xFF0F172A),
                                    letterSpacing: -0.5,
                                  ),
                                ),
                                TextSpan(
                                  text: "gama",
                                  style: TextStyle(
                                    fontSize: 20,
                                    fontWeight: FontWeight.w900,
                                    color: Color(0xFF1E60F2),
                                    letterSpacing: -0.5,
                                  ),
                                ),
                              ],
                            ),
                          ),
                        ],
                      ),
                    ),

                    // Pojok Kanan: Icon Notifikasi (Memicu Pop-up Notifikasi)
                    _buildIconButton(
                      icon: Icons.notifications_none_rounded,
                      onTap: _showNotificationsDialog,
                      hasBadge: hasNotificationBadge,
                    ),
                  ],
                ),

                const SizedBox(height: 18),

                // =========================================================================
                // POIN 5: DIBAWAH TOPBAR DENGAN TATA LETAK SERAGAM (Kiri & Kanan 3 Baris)
                // Kiri:
                //   Baris 1: Selamat Siang,
                //   Baris 2: Nama Part 1
                //   Baris 3: Nama Part 2 (jika > 1 kata)
                // Kanan:
                //   Baris 1: Kamis,
                //   Baris 2: 03 September 2026
                //   Baris 3: [ 📍 Kantor ]
                // =========================================================================
                Row(
                  mainAxisAlignment: MainAxisAlignment.spaceBetween,
                  crossAxisAlignment: CrossAxisAlignment.start,
                  children: [
                    // Sisi Kiri: Sapaan & Nama User
                    Expanded(
                      child: Column(
                        crossAxisAlignment: CrossAxisAlignment.start,
                        children: [
                          Text(
                            _getGreeting(),
                            style: const TextStyle(
                              fontSize: 13,
                              color: Color(0xFF64748B),
                              fontWeight: FontWeight.w500,
                            ),
                          ),
                          const SizedBox(height: 2),
                          Text(
                            nameParts[0],
                            style: const TextStyle(
                              fontSize: 18,
                              fontWeight: FontWeight.bold,
                              color: Color(0xFF0F172A),
                              letterSpacing: -0.3,
                              height: 1.18,
                            ),
                            maxLines: 1,
                            overflow: TextOverflow.ellipsis,
                          ),
                          if (nameParts[1].isNotEmpty) ...[
                            const SizedBox(height: 1),
                            Text(
                              nameParts[1],
                              style: const TextStyle(
                                fontSize: 18,
                                fontWeight: FontWeight.bold,
                                color: Color(0xFF0F172A),
                                letterSpacing: -0.3,
                                height: 1.18,
                              ),
                              maxLines: 1,
                              overflow: TextOverflow.ellipsis,
                            ),
                          ],
                        ],
                      ),
                    ),

                    const SizedBox(width: 12),

                    // Sisi Kanan: Hari, Tanggal, dan Status Lokasi (3 Baris)
                    Column(
                      crossAxisAlignment: CrossAxisAlignment.end,
                      children: [
                        Text(
                          "$dayNameFormatted,",
                          style: const TextStyle(
                            fontSize: 13,
                            fontWeight: FontWeight.w600,
                            color: Color(0xFF334155),
                          ),
                        ),
                        Text(
                          dateFormatted,
                          style: const TextStyle(
                            fontSize: 11.5,
                            fontWeight: FontWeight.w500,
                            color: Color(0xFF64748B),
                          ),
                        ),
                        const SizedBox(height: 6),
                        // Border Persegi Panjang Status Lokasi
                        Container(
                          padding: const EdgeInsets.symmetric(horizontal: 10, vertical: 4.5),
                          decoration: BoxDecoration(
                            color: _isInOfficeRadius
                                ? const Color(0xFFECFDF5)
                                : const Color(0xFFF1F5F9),
                            borderRadius: BorderRadius.circular(10),
                            border: Border.all(
                              color: _isInOfficeRadius
                                  ? const Color(0xFF10B981)
                                  : const Color(0xFFCBD5E1),
                              width: 1.2,
                            ),
                          ),
                          child: Row(
                            mainAxisSize: MainAxisSize.min,
                            children: [
                              Icon(
                                Icons.location_on_rounded,
                                size: 14,
                                color: _isInOfficeRadius
                                    ? const Color(0xFF059669)
                                    : const Color(0xFF64748B),
                              ),
                              const SizedBox(width: 4),
                              Text(
                                _isCheckingLocation
                                    ? "Mengecek..."
                                    : (_isInOfficeRadius ? "Kantor" : "Di Luar Kantor"),
                                style: TextStyle(
                                  fontSize: 11,
                                  fontWeight: FontWeight.bold,
                                  color: _isInOfficeRadius
                                      ? const Color(0xFF059669)
                                      : const Color(0xFF475569),
                                ),
                              ),
                            ],
                          ),
                        ),
                      ],
                    ),
                  ],
                ),

                const SizedBox(height: 20),

                // =========================================================================
                // BORDER BIRU BESAR
                // POIN 3: "Rekam Waktu Pergi"
                // POIN 4: "Rekam Waktu Kembali"
                // =========================================================================
                Container(
                  decoration: BoxDecoration(
                    gradient: LinearGradient(
                      colors: hasActivePermit
                          ? const [Color(0xFFEA580C), Color(0xFFC2410C)]
                          : const [Color(0xFF1E60F2), Color(0xFF0E3A99)],
                      begin: Alignment.topLeft,
                      end: Alignment.bottomRight,
                    ),
                    borderRadius: BorderRadius.circular(26),
                    boxShadow: [
                      BoxShadow(
                        color: (hasActivePermit
                                ? const Color(0xFFEA580C)
                                : const Color(0xFF1E60F2))
                            .withOpacity(0.35),
                        blurRadius: 20,
                        offset: const Offset(0, 8),
                      ),
                    ],
                  ),
                  padding: const EdgeInsets.all(22),
                  child: Column(
                    children: [
                      // Digital Time Display
                      Column(
                        children: [
                          Text(
                            "WAKTU SAAT INI",
                            style: TextStyle(
                              fontSize: 11,
                              fontWeight: FontWeight.w700,
                              color: Colors.white.withOpacity(0.75),
                              letterSpacing: 0.8,
                            ),
                          ),
                          const SizedBox(height: 2),
                          Text(
                            "$timeFormatted WIT",
                            style: const TextStyle(
                              fontSize: 32,
                              fontWeight: FontWeight.w900,
                              color: Colors.white,
                              letterSpacing: -0.5,
                            ),
                          ),

                          // Indikator "Sedang Izin" Di Bawah Jam
                          if (hasActivePermit) ...[
                            const SizedBox(height: 8),
                            Container(
                              padding: const EdgeInsets.symmetric(horizontal: 14, vertical: 5),
                              decoration: BoxDecoration(
                                color: Colors.black.withOpacity(0.18),
                                borderRadius: BorderRadius.circular(20),
                                border: Border.all(
                                  color: const Color(0xFFFDE68A).withOpacity(0.7),
                                  width: 1.2,
                                ),
                              ),
                              child: Row(
                                mainAxisSize: MainAxisSize.min,
                                children: [
                                  Container(
                                    width: 7,
                                    height: 7,
                                    decoration: const BoxDecoration(
                                      color: Color(0xFFFDE68A),
                                      shape: BoxShape.circle,
                                    ),
                                  ),
                                  const SizedBox(width: 6),
                                  const Text(
                                    "Sedang Izin",
                                    style: TextStyle(
                                      color: Color(0xFFFEF3C7),
                                      fontWeight: FontWeight.bold,
                                      fontSize: 12,
                                      letterSpacing: 0.2,
                                    ),
                                  ),
                                ],
                              ),
                            ),
                          ],
                        ],
                      ),

                      const SizedBox(height: 20),

                      // 2 Tombol Aksi: Mulai Izin & Selesai Izin
                      Row(
                        children: [
                          // Tombol 1: Mulai Izin (POIN 3: Rekam Waktu Pergi)
                          Expanded(
                            child: InkWell(
                              onTap: hasActivePermit ? null : _navigateToRecordPermit,
                              borderRadius: BorderRadius.circular(18),
                              child: Container(
                                padding: const EdgeInsets.symmetric(vertical: 14),
                                decoration: BoxDecoration(
                                  color: hasActivePermit
                                      ? Colors.white.withOpacity(0.22)
                                      : Colors.white,
                                  borderRadius: BorderRadius.circular(18),
                                  border: hasActivePermit
                                      ? Border.all(color: Colors.white.withOpacity(0.35), width: 1.2)
                                      : null,
                                  boxShadow: hasActivePermit
                                      ? null
                                      : [
                                          BoxShadow(
                                            color: Colors.black.withOpacity(0.08),
                                            blurRadius: 8,
                                            offset: const Offset(0, 2),
                                          ),
                                        ],
                                ),
                                child: Column(
                                  children: [
                                    Icon(
                                      Icons.check_circle_rounded,
                                      color: hasActivePermit
                                          ? Colors.white
                                          : const Color(0xFF1E60F2),
                                      size: 24,
                                    ),
                                    const SizedBox(height: 6),
                                    Text(
                                      "Mulai Izin",
                                      style: TextStyle(
                                        color: hasActivePermit
                                            ? Colors.white
                                            : const Color(0xFF1E60F2),
                                        fontWeight: FontWeight.bold,
                                        fontSize: 13.5,
                                      ),
                                    ),
                                    const SizedBox(height: 2),
                                    Text(
                                      hasActivePermit ? startPermitTime : "Rekam Waktu Pergi",
                                      style: TextStyle(
                                        fontSize: 11,
                                        color: hasActivePermit
                                            ? Colors.white.withOpacity(0.8)
                                            : const Color(0xFF1E60F2).withOpacity(0.75),
                                        fontWeight: FontWeight.w600,
                                      ),
                                    ),
                                  ],
                                ),
                              ),
                            ),
                          ),

                          const SizedBox(width: 14),

                          // Tombol 2: Selesai Izin (POIN 4: Rekam Waktu Kembali)
                          Expanded(
                            child: InkWell(
                              onTap: hasActivePermit ? _navigateToRecordPermit : null,
                              borderRadius: BorderRadius.circular(18),
                              child: Container(
                                padding: const EdgeInsets.symmetric(vertical: 14),
                                decoration: BoxDecoration(
                                  color: hasActivePermit
                                      ? Colors.white
                                      : Colors.white.withOpacity(0.15),
                                  borderRadius: BorderRadius.circular(18),
                                  border: hasActivePermit
                                      ? null
                                      : Border.all(
                                          color: Colors.white.withOpacity(0.2),
                                          width: 1.2,
                                        ),
                                  boxShadow: hasActivePermit
                                      ? [
                                          BoxShadow(
                                            color: Colors.black.withOpacity(0.14),
                                            blurRadius: 10,
                                            offset: const Offset(0, 3),
                                          ),
                                        ]
                                      : null,
                                ),
                                child: Column(
                                  children: [
                                    Icon(
                                      Icons.access_time_filled_rounded,
                                      color: hasActivePermit
                                          ? const Color(0xFFEA580C)
                                          : Colors.white,
                                      size: 24,
                                    ),
                                    const SizedBox(height: 6),
                                    Text(
                                      "Selesai Izin",
                                      style: TextStyle(
                                        color: hasActivePermit
                                            ? const Color(0xFFEA580C)
                                            : Colors.white,
                                        fontWeight: FontWeight.bold,
                                        fontSize: 13.5,
                                      ),
                                    ),
                                    const SizedBox(height: 2),
                                    Text(
                                      hasActivePermit ? "Rekam Waktu Kembali" : endPermitTime,
                                      style: TextStyle(
                                        fontSize: 11,
                                        color: hasActivePermit
                                            ? const Color(0xFFEA580C).withOpacity(0.85)
                                            : Colors.white.withOpacity(0.75),
                                        fontWeight: FontWeight.w700,
                                      ),
                                    ),
                                  ],
                                ),
                              ),
                            ),
                          ),
                        ],
                      ),
                    ],
                  ),
                ),

                // Card Peringatan Sedang Izin
                if (hasActivePermit) ...[
                  const SizedBox(height: 16),
                  InkWell(
                    onTap: _navigateToRecordPermit,
                    borderRadius: BorderRadius.circular(20),
                    child: Container(
                      padding: const EdgeInsets.all(16),
                      decoration: BoxDecoration(
                        color: Colors.white,
                        borderRadius: BorderRadius.circular(20),
                        border: Border.all(color: const Color(0xFFFDE68A), width: 1.2),
                        boxShadow: [
                          BoxShadow(
                            color: const Color(0xFFD97706).withOpacity(0.08),
                            blurRadius: 14,
                            offset: const Offset(0, 4),
                          ),
                        ],
                      ),
                      child: Row(
                        children: [
                          Container(
                            width: 44,
                            height: 44,
                            decoration: BoxDecoration(
                              color: const Color(0xFFFEF3C7),
                              borderRadius: BorderRadius.circular(14),
                            ),
                            child: const Icon(Icons.directions_walk_rounded, color: Color(0xFFD97706), size: 24),
                          ),
                          const SizedBox(width: 14),
                          Expanded(
                            child: Column(
                              crossAxisAlignment: CrossAxisAlignment.start,
                              children: [
                                Container(
                                  padding: const EdgeInsets.symmetric(horizontal: 7, vertical: 2),
                                  decoration: BoxDecoration(
                                    color: const Color(0xFFFEF3C7),
                                    borderRadius: BorderRadius.circular(6),
                                  ),
                                  child: const Text(
                                    "SEDANG IZIN",
                                    style: TextStyle(
                                      fontSize: 10,
                                      fontWeight: FontWeight.bold,
                                      color: Color(0xFFB45309),
                                    ),
                                  ),
                                ),
                                const SizedBox(height: 5),
                                const Text(
                                  "Segera rekam waktu kembali setelah tiba di kantor",
                                  style: TextStyle(
                                    fontSize: 12.5,
                                    fontWeight: FontWeight.w600,
                                    color: Color(0xFF1E293B),
                                    height: 1.3,
                                  ),
                                ),
                                const SizedBox(height: 5),
                                Row(
                                  crossAxisAlignment: CrossAxisAlignment.start,
                                  children: [
                                    const Padding(
                                      padding: EdgeInsets.only(top: 2),
                                      child: Icon(Icons.edit_note_rounded, size: 14, color: Color(0xFF64748B)),
                                    ),
                                    const SizedBox(width: 4),
                                    Expanded(
                                      child: Text(
                                        "Keperluan: ${_activePermit!.purpose}",
                                        style: const TextStyle(
                                          fontSize: 11.5,
                                          color: Color(0xFF64748B),
                                          fontWeight: FontWeight.w500,
                                          height: 1.3,
                                        ),
                                        maxLines: 3,
                                        overflow: TextOverflow.ellipsis,
                                      ),
                                    ),
                                  ],
                                ),
                              ],
                            ),
                          ),
                        ],
                      ),
                    ),
                  ),
                ],

                // Card Peringatan Laporan Cepu
                if (_activeCepuForMe != null) ...[
                  const SizedBox(height: 16),
                  InkWell(
                    onTap: _navigateToCepuMapReport,
                    borderRadius: BorderRadius.circular(20),
                    child: Container(
                      padding: const EdgeInsets.all(16),
                      decoration: BoxDecoration(
                        color: Colors.white,
                        borderRadius: BorderRadius.circular(20),
                        border: Border.all(color: const Color(0xFFFECACA), width: 1.2),
                        boxShadow: [
                          BoxShadow(
                            color: const Color(0xFFDC2626).withOpacity(0.08),
                            blurRadius: 14,
                            offset: const Offset(0, 4),
                          ),
                        ],
                      ),
                      child: Row(
                        children: [
                          Container(
                            width: 44,
                            height: 44,
                            decoration: BoxDecoration(
                              color: const Color(0xFFFEE2E2),
                              borderRadius: BorderRadius.circular(14),
                            ),
                            child: const Icon(Icons.warning_amber_rounded, color: Color(0xFFDC2626), size: 24),
                          ),
                          const SizedBox(width: 14),
                          Expanded(
                            child: Column(
                              crossAxisAlignment: CrossAxisAlignment.start,
                              children: [
                                Container(
                                  padding: const EdgeInsets.symmetric(horizontal: 7, vertical: 2),
                                  decoration: BoxDecoration(
                                    color: const Color(0xFFFEE2E2),
                                    borderRadius: BorderRadius.circular(6),
                                  ),
                                  child: const Text(
                                    "LAPORAN CEPU",
                                    style: TextStyle(
                                      fontSize: 10,
                                      fontWeight: FontWeight.bold,
                                      color: Color(0xFF991B1B),
                                    ),
                                  ),
                                ),
                                const SizedBox(height: 5),
                                const Text(
                                  "Anda dilaporkan tidak berada di kantor. Segera rekam waktu kembali.",
                                  style: TextStyle(
                                    fontSize: 12.5,
                                    fontWeight: FontWeight.w600,
                                    color: Color(0xFF1E293B),
                                    height: 1.3,
                                  ),
                                ),
                                const SizedBox(height: 5),
                                const Text(
                                  "👉 Ketuk untuk rekam waktu kembali",
                                  style: TextStyle(
                                    fontSize: 11.5,
                                    color: Color(0xFFDC2626),
                                    fontWeight: FontWeight.bold,
                                  ),
                                ),
                              ],
                            ),
                          ),
                        ],
                      ),
                    ),
                  ),
                ],

                const SizedBox(height: 22),

                // =========================================================================
                // POIN 1 & 2: AKTIVITAS HARI INI TERHUBUNG DENGAN DATABASE
                // - Menampilkan seluruh kegiatan user hari ini (Izin & Yang Di-Cepukan Valid)
                // - Jika belum ada aktivitas, menampilkan desain "Belum ada aktivitas"
                // =========================================================================
                Row(
                  mainAxisAlignment: MainAxisAlignment.spaceBetween,
                  children: [
                    const Text(
                      "Aktivitas Hari Ini",
                      style: TextStyle(
                        fontSize: 16.5,
                        fontWeight: FontWeight.bold,
                        color: Color(0xFF0F172A),
                      ),
                    ),
                    InkWell(
                      onTap: _navigateToAnalytics,
                      child: const Text(
                        "Lihat Rekap",
                        style: TextStyle(fontSize: 12, color: Color(0xFF1E60F2), fontWeight: FontWeight.bold),
                      ),
                    ),
                  ],
                ),

                const SizedBox(height: 12),

                // POIN 2: Kondisi Belum Ada Aktivitas
                if (_todayActivities.isEmpty)
                  Container(
                    width: double.infinity,
                    padding: const EdgeInsets.symmetric(vertical: 36, horizontal: 20),
                    decoration: BoxDecoration(
                      color: Colors.white,
                      borderRadius: BorderRadius.circular(22),
                      border: Border.all(color: const Color(0xFFE2E8F0)),
                      boxShadow: [
                        BoxShadow(
                          color: Colors.black.withOpacity(0.02),
                          blurRadius: 10,
                          offset: const Offset(0, 2),
                        ),
                      ],
                    ),
                    child: Column(
                      children: [
                        Container(
                          width: 52,
                          height: 52,
                          decoration: const BoxDecoration(
                            color: Color(0xFFF1F5F9),
                            shape: BoxShape.circle,
                          ),
                          child: const Icon(Icons.event_available_rounded, color: Color(0xFF94A3B8), size: 28),
                        ),
                        const SizedBox(height: 12),
                        const Text(
                          "Belum ada aktivitas",
                          style: TextStyle(
                            fontSize: 15,
                            fontWeight: FontWeight.bold,
                            color: Color(0xFF334155),
                          ),
                        ),
                        const SizedBox(height: 4),
                        const Text(
                          "Anda belum merekam izin atau memiliki riwayat kegiatan hari ini.",
                          textAlign: TextAlign.center,
                          style: TextStyle(
                            fontSize: 12,
                            color: Color(0xFF94A3B8),
                          ),
                        ),
                      ],
                    ),
                  )
                else ...[
                  // POIN 1: List Seluruh Aktivitas Hari Ini dengan Pagination
                  Builder(
                    builder: (context) {
                      final int totalPages = (_todayActivities.length / _activityPageSize).ceil();
                      final int safePage = _currentActivityPage.clamp(0, max(0, totalPages - 1)).toInt();
                      final int startIndex = safePage * _activityPageSize;
                      final int endIndex = min(startIndex + _activityPageSize, _todayActivities.length);
                      final pageItems = _todayActivities.sublist(startIndex, endIndex);

                      return GestureDetector(
                        behavior: HitTestBehavior.opaque,
                        onHorizontalDragEnd: (details) {
                          if (details.primaryVelocity != null && totalPages > 1) {
                            if (details.primaryVelocity! < -150) {
                              // Geser ke kiri -> Halaman berikutnya
                              if (safePage < totalPages - 1) {
                                setState(() => _currentActivityPage = safePage + 1);
                              }
                            } else if (details.primaryVelocity! > 150) {
                              // Geser ke kanan -> Halaman sebelumnya
                              if (safePage > 0) {
                                setState(() => _currentActivityPage = safePage - 1);
                              }
                            }
                          }
                        },
                        child: Column(
                          crossAxisAlignment: CrossAxisAlignment.start,
                          children: [
                            ListView.separated(
                              shrinkWrap: true,
                              physics: const NeverScrollableScrollPhysics(),
                              itemCount: pageItems.length,
                              separatorBuilder: (_, __) => const SizedBox(height: 12),
                              itemBuilder: (context, index) {
                                final item = pageItems[index];
                                if (item.type == 'CEPU') {
                                  return _buildCepuActivityCard(item);
                                }
                                return _buildPermitActivityCard(item);
                              },
                            ),
                            if (totalPages > 1)
                              _buildActivityPagination(totalPages, safePage),
                          ],
                        ),
                      );
                    },
                  ),
                ],
              ],
            ),
          ),
        ),
      ),

      // =========================================================================
      // BOTTOM NAVIGATION BAR (Flush di Dasar Layar)
      // =========================================================================
      bottomNavigationBar: Container(
        height: 72,
        decoration: BoxDecoration(
          color: Colors.white,
          borderRadius: const BorderRadius.vertical(top: Radius.circular(24)),
          boxShadow: [
            BoxShadow(
              color: Colors.black.withOpacity(0.06),
              blurRadius: 16,
              offset: const Offset(0, -4),
            ),
          ],
        ),
        child: Row(
          mainAxisAlignment: MainAxisAlignment.spaceAround,
          children: [
            IconButton(
              icon: const Icon(Icons.home_rounded, color: Color(0xFF1E60F2), size: 26),
              onPressed: () {},
              tooltip: "Home",
            ),
            IconButton(
              icon: const Icon(Icons.bar_chart_rounded, color: Color(0xFF94A3B8), size: 26),
              onPressed: _navigateToAnalytics,
              tooltip: "Analytics",
            ),
            IconButton(
              icon: const Icon(Icons.fact_check_outlined, color: Color(0xFF94A3B8), size: 26),
              onPressed: _navigateToDailyMonitoring,
              tooltip: "Monitoring Harian",
            ),
            IconButton(
              icon: const Icon(Icons.campaign_outlined, color: Color(0xFF94A3B8), size: 26),
              onPressed: _navigateToCepuMapReport,
              tooltip: "Cepu",
            ),
            IconButton(
              icon: const Icon(Icons.settings_outlined, color: Color(0xFF94A3B8), size: 26),
              onPressed: _navigateToSettings,
              tooltip: "Settings",
            ),
          ],
        ),
      ),
    );
  }

  String _formatDuration(int minutes) {
    if (minutes < 60) return "$minutes Menit";
    final hours = minutes ~/ 60;
    final rem = minutes % 60;
    if (rem == 0) return "$hours Jam";
    return "$hours Jam $rem Menit";
  }

  Widget _buildActivityCardDuration(int duration, Color durationColor) {
    if (duration < 60) {
      return Text(
        "$duration Menit",
        style: TextStyle(
          fontSize: 13,
          fontWeight: FontWeight.w900,
          color: durationColor,
        ),
      );
    }
    final hours = duration ~/ 60;
    final rem = duration % 60;
    return Column(
      crossAxisAlignment: CrossAxisAlignment.end,
      mainAxisSize: MainAxisSize.min,
      children: [
        Text(
          "$hours Jam",
          style: TextStyle(
            fontSize: 13,
            fontWeight: FontWeight.w900,
            color: durationColor,
            height: 1.15,
          ),
        ),
        const SizedBox(height: 1),
        Text(
          "$rem Menit",
          style: TextStyle(
            fontSize: 11.5,
            fontWeight: FontWeight.w800,
            color: durationColor.withOpacity(0.92),
            height: 1.15,
          ),
        ),
      ],
    );
  }

  /// Pagination Control untuk Aktivitas Hari Ini
  Widget _buildActivityPagination(int totalPages, int currentPage) {
    return Padding(
      padding: const EdgeInsets.only(top: 12),
      child: Row(
        mainAxisAlignment: MainAxisAlignment.center,
        children: [
          IconButton(
            icon: const Icon(Icons.chevron_left_rounded),
            color: currentPage > 0 ? const Color(0xFF1E60F2) : const Color(0xFFCBD5E1),
            onPressed: currentPage > 0
                ? () {
                    setState(() {
                      _currentActivityPage--;
                    });
                  }
                : null,
            splashRadius: 20,
            tooltip: "Halaman Sebelumnya",
          ),
          Container(
            padding: const EdgeInsets.symmetric(horizontal: 14, vertical: 6),
            decoration: BoxDecoration(
              color: Colors.white,
              borderRadius: BorderRadius.circular(20),
              border: Border.all(color: const Color(0xFFE2E8F0)),
              boxShadow: [
                BoxShadow(
                  color: Colors.black.withOpacity(0.02),
                  blurRadius: 6,
                  offset: const Offset(0, 2),
                ),
              ],
            ),
            child: Text(
              "Halaman ${currentPage + 1} dari $totalPages",
              style: const TextStyle(
                fontSize: 12,
                fontWeight: FontWeight.bold,
                color: Color(0xFF475569),
              ),
            ),
          ),
          IconButton(
            icon: const Icon(Icons.chevron_right_rounded),
            color: currentPage < totalPages - 1 ? const Color(0xFF1E60F2) : const Color(0xFFCBD5E1),
            onPressed: currentPage < totalPages - 1
                ? () {
                    setState(() {
                      _currentActivityPage++;
                    });
                  }
                : null,
            splashRadius: 20,
            tooltip: "Halaman Berikutnya",
          ),
        ],
      ),
    );
  }

  /// Modal Bottom Sheet untuk Melihat Keperluan & Detail Lengkap Aktivitas Selesai
  void _showActivityDetailModal(UserActivityItem item) {
    final startStr = "${DateFormat('HH:mm').format(item.startTime)} WIT";
    final endStr = item.endTime != null ? "${DateFormat('HH:mm').format(item.endTime!)} WIT" : "--:--";
    final duration = item.getDurationMinutes();
    final durationFormatted = _formatDuration(duration);
    final isCepu = item.type == 'CEPU';

    showModalBottomSheet(
      context: context,
      isScrollControlled: true,
      backgroundColor: Colors.transparent,
      builder: (context) {
        return Container(
          decoration: const BoxDecoration(
            color: Colors.white,
            borderRadius: BorderRadius.vertical(top: Radius.circular(28)),
          ),
          padding: EdgeInsets.fromLTRB(24, 16, 24, MediaQuery.of(context).viewInsets.bottom + 28),
          child: Column(
            mainAxisSize: MainAxisSize.min,
            crossAxisAlignment: CrossAxisAlignment.start,
            children: [
              // Drag Handle
              Center(
                child: Container(
                  width: 44,
                  height: 4,
                  decoration: BoxDecoration(
                    color: const Color(0xFFCBD5E1),
                    borderRadius: BorderRadius.circular(2),
                  ),
                ),
              ),
              const SizedBox(height: 18),

              // Keterangan khusus jika izin sedang berlangsung
              if (item.isActive) ...[
                Container(
                  width: double.infinity,
                  padding: const EdgeInsets.symmetric(horizontal: 14, vertical: 12),
                  decoration: BoxDecoration(
                    color: const Color(0xFFFEF3C7),
                    borderRadius: BorderRadius.circular(14),
                    border: Border.all(color: const Color(0xFFFDE68A)),
                  ),
                  child: Row(
                    children: [
                      const Icon(Icons.info_rounded, color: Color(0xFFD97706), size: 20),
                      const SizedBox(width: 10),
                      const Expanded(
                        child: Text(
                          "Izin saat ini masih berlangsung. Segera rekam waktu kembali setelah Anda tiba di kantor.",
                          style: TextStyle(
                            fontSize: 12,
                            color: Color(0xFF92400E),
                            fontWeight: FontWeight.w600,
                            height: 1.35,
                          ),
                        ),
                      ),
                    ],
                  ),
                ),
                const SizedBox(height: 16),
              ],

              // Header Type & Status
              Row(
                mainAxisAlignment: MainAxisAlignment.spaceBetween,
                children: [
                  Row(
                    children: [
                      Container(
                        padding: const EdgeInsets.all(8),
                        decoration: BoxDecoration(
                          color: isCepu ? const Color(0xFFFEE2E2) : const Color(0xFFEFF6FF),
                          borderRadius: BorderRadius.circular(10),
                        ),
                        child: Icon(
                          isCepu ? Icons.campaign_rounded : Icons.assignment_ind_rounded,
                          color: isCepu ? const Color(0xFFDC2626) : const Color(0xFF1E60F2),
                          size: 20,
                        ),
                      ),
                      const SizedBox(width: 10),
                      Text(
                        isCepu ? "Laporan Cepu" : "Izin Waigama",
                        style: const TextStyle(
                          fontSize: 17,
                          fontWeight: FontWeight.bold,
                          color: Color(0xFF0F172A),
                        ),
                      ),
                    ],
                  ),
                  Container(
                    padding: const EdgeInsets.symmetric(horizontal: 10, vertical: 4),
                    decoration: BoxDecoration(
                      color: item.isActive
                          ? (isCepu ? const Color(0xFFFEE2E2) : const Color(0xFFFEF3C7))
                          : const Color(0xFFD1FAE5),
                      borderRadius: BorderRadius.circular(8),
                    ),
                    child: Text(
                      item.isActive ? "Sedang Berjalan" : "Selesai",
                      style: TextStyle(
                        fontSize: 11,
                        fontWeight: FontWeight.bold,
                        color: item.isActive
                            ? (isCepu ? const Color(0xFFDC2626) : const Color(0xFFB45309))
                            : const Color(0xFF059669),
                      ),
                    ),
                  ),
                ],
              ),

              const SizedBox(height: 20),

              // Info Box: Waktu Mulai, Selesai, & Durasi
              Container(
                padding: const EdgeInsets.all(16),
                decoration: BoxDecoration(
                  color: const Color(0xFFF8FAFC),
                  borderRadius: BorderRadius.circular(16),
                  border: Border.all(color: const Color(0xFFE2E8F0)),
                ),
                child: Row(
                  mainAxisAlignment: MainAxisAlignment.spaceAround,
                  children: [
                    Column(
                      crossAxisAlignment: CrossAxisAlignment.start,
                      children: [
                        const Text("Waktu Pergi", style: TextStyle(fontSize: 11, color: Color(0xFF64748B))),
                        const SizedBox(height: 3),
                        Text(startStr, style: const TextStyle(fontSize: 14, fontWeight: FontWeight.bold, color: Color(0xFF0F172A))),
                      ],
                    ),
                    Container(height: 30, width: 1, color: const Color(0xFFE2E8F0)),
                    Column(
                      crossAxisAlignment: CrossAxisAlignment.start,
                      children: [
                        const Text("Waktu Kembali", style: TextStyle(fontSize: 11, color: Color(0xFF64748B))),
                        const SizedBox(height: 3),
                        Text(endStr, style: const TextStyle(fontSize: 14, fontWeight: FontWeight.bold, color: Color(0xFF0F172A))),
                      ],
                    ),
                    Container(height: 30, width: 1, color: const Color(0xFFE2E8F0)),
                    Column(
                      crossAxisAlignment: CrossAxisAlignment.start,
                      children: [
                        const Text("Total Durasi", style: TextStyle(fontSize: 11, color: Color(0xFF64748B))),
                        const SizedBox(height: 3),
                        Text(durationFormatted, style: const TextStyle(fontSize: 13, fontWeight: FontWeight.bold, color: Color(0xFF1E60F2))),
                      ],
                    ),
                  ],
                ),
              ),

              const SizedBox(height: 18),

              // Detail Keperluan / Deskripsi Lengkap
              const Text(
                "Keperluan / Deskripsi Lengkap",
                style: TextStyle(
                  fontSize: 13,
                  fontWeight: FontWeight.bold,
                  color: Color(0xFF334155),
                ),
              ),
              const SizedBox(height: 8),
              Container(
                width: double.infinity,
                padding: const EdgeInsets.all(14),
                decoration: BoxDecoration(
                  color: const Color(0xFFF8FAFC),
                  borderRadius: BorderRadius.circular(14),
                  border: Border.all(color: const Color(0xFFE2E8F0)),
                ),
                child: SelectableText(
                  item.purposeOrDesc.isNotEmpty ? item.purposeOrDesc : "-",
                  style: const TextStyle(
                    fontSize: 13,
                    color: Color(0xFF1E293B),
                    height: 1.5,
                  ),
                ),
              ),

              if (isCepu && item.cepuModel?.returnReason != null && item.cepuModel!.returnReason!.isNotEmpty) ...[
                const SizedBox(height: 14),
                const Text(
                  "Alasan / Keperluan Kembali",
                  style: TextStyle(
                    fontSize: 13,
                    fontWeight: FontWeight.bold,
                    color: Color(0xFF334155),
                  ),
                ),
                const SizedBox(height: 8),
                Container(
                  width: double.infinity,
                  padding: const EdgeInsets.all(14),
                  decoration: BoxDecoration(
                    color: const Color(0xFFF0FDF4),
                    borderRadius: BorderRadius.circular(14),
                    border: Border.all(color: const Color(0xFFBBF7D0)),
                  ),
                  child: SelectableText(
                    item.cepuModel!.returnReason!,
                    style: const TextStyle(
                      fontSize: 13,
                      color: Color(0xFF166534),
                      height: 1.5,
                    ),
                  ),
                ),
              ],

              if (isCepu && item.reporterName != null) ...[
                const SizedBox(height: 14),
                Row(
                  children: [
                    const Icon(Icons.person_pin_rounded, size: 16, color: Color(0xFF64748B)),
                    const SizedBox(width: 6),
                    Text(
                      "Dilaporkan oleh: ${item.reporterName}",
                      style: const TextStyle(fontSize: 12, color: Color(0xFF64748B)),
                    ),
                  ],
                ),
              ],

              const SizedBox(height: 24),

              // Tombol Aksi Tambahan jika Sedang Izin
              if (item.isActive) ...[
                SizedBox(
                  width: double.infinity,
                  child: ElevatedButton.icon(
                    onPressed: () {
                      Navigator.pop(context);
                      if (item.type == 'CEPU') {
                        _navigateToCepuMapReport();
                      } else {
                        _navigateToRecordPermit();
                      }
                    },
                    icon: const Icon(Icons.camera_alt_rounded, size: 18),
                    label: const Text(
                      "Rekam Waktu Kembali",
                      style: TextStyle(fontSize: 14, fontWeight: FontWeight.bold),
                    ),
                    style: ElevatedButton.styleFrom(
                      backgroundColor: const Color(0xFFEA580C),
                      foregroundColor: Colors.white,
                      shape: RoundedRectangleBorder(borderRadius: BorderRadius.circular(14)),
                      padding: const EdgeInsets.symmetric(vertical: 14),
                      elevation: 1,
                    ),
                  ),
                ),
                const SizedBox(height: 10),
              ],

              // Tombol Tutup
              SizedBox(
                width: double.infinity,
                child: ElevatedButton(
                  onPressed: () => Navigator.pop(context),
                  style: ElevatedButton.styleFrom(
                    backgroundColor: item.isActive ? const Color(0xFFF1F5F9) : const Color(0xFF1E60F2),
                    foregroundColor: item.isActive ? const Color(0xFF475569) : Colors.white,
                    shape: RoundedRectangleBorder(borderRadius: BorderRadius.circular(14)),
                    padding: const EdgeInsets.symmetric(vertical: 14),
                    elevation: 0,
                  ),
                  child: const Text("Tutup", style: TextStyle(fontSize: 14, fontWeight: FontWeight.bold)),
                ),
              ),
            ],
          ),
        );
      },
    );
  }

  /// Card Aktivitas: Izin Waigama (Desain Sebelumnya: Gradasi Elegan & Tidak Monoton)
  Widget _buildPermitActivityCard(UserActivityItem item) {
    final startStr = "${DateFormat('HH:mm').format(item.startTime)} WIT";
    final endStr = item.endTime != null ? "${DateFormat('HH:mm').format(item.endTime!)} WIT" : "--:--";
    final duration = item.getDurationMinutes();

    // Warna dinamis: Sedang Izin (Amber-Oranye Hangat), Selesai (Emerald Teal Tenang)
    final gradientColors = item.isActive
        ? const [Color(0xFFF59E0B), Color(0xFFD97706)]
        : const [Color(0xFF0D9488), Color(0xFF0F766E)];
    final shadowColor = item.isActive
        ? const Color(0xFFF59E0B).withOpacity(0.28)
        : const Color(0xFF0D9488).withOpacity(0.25);
    final durationColor = item.isActive
        ? const Color(0xFFFEF3C7)
        : const Color(0xFFCCFBF1);

    return InkWell(
      onTap: () => _showActivityDetailModal(item),
      borderRadius: BorderRadius.circular(22),
      child: Container(
        decoration: BoxDecoration(
          gradient: LinearGradient(
            colors: gradientColors,
            begin: Alignment.topLeft,
            end: Alignment.bottomRight,
          ),
          borderRadius: BorderRadius.circular(22),
          boxShadow: [
            BoxShadow(
              color: shadowColor,
              blurRadius: 14,
              offset: const Offset(0, 5),
            ),
          ],
        ),
        padding: const EdgeInsets.all(16),
        child: Row(
          children: [
            // Kotak Kiri: "Sedang Izin" atau "Selesai"
            Container(
              width: 72,
              padding: const EdgeInsets.symmetric(horizontal: 6, vertical: 12),
              decoration: BoxDecoration(
                color: Colors.white.withOpacity(0.22),
                borderRadius: BorderRadius.circular(16),
                border: Border.all(
                  color: Colors.white.withOpacity(0.35),
                  width: 1.2,
                ),
              ),
              child: Column(
                mainAxisSize: MainAxisSize.min,
                children: [
                  Icon(
                    item.isActive ? Icons.directions_walk_rounded : Icons.check_circle_rounded,
                    color: Colors.white,
                    size: 22,
                  ),
                  const SizedBox(height: 4),
                  Text(
                    item.isActive ? "Sedang\nIzin" : "Selesai",
                    textAlign: TextAlign.center,
                    style: const TextStyle(
                      fontSize: 11,
                      fontWeight: FontWeight.w900,
                      color: Colors.white,
                      height: 1.2,
                    ),
                  ),
                ],
              ),
            ),

            const SizedBox(width: 16),

            // Rincian Waktu & Durasi Menit / Jam
            Expanded(
              child: Column(
                crossAxisAlignment: CrossAxisAlignment.start,
                children: [
                  Row(
                    mainAxisAlignment: MainAxisAlignment.spaceBetween,
                    children: [
                      Column(
                        crossAxisAlignment: CrossAxisAlignment.start,
                        children: [
                          Text(
                            "Mulai",
                            style: TextStyle(fontSize: 11, color: Colors.white.withOpacity(0.75)),
                          ),
                          const SizedBox(height: 2),
                          Text(
                            startStr,
                            style: const TextStyle(fontSize: 13, fontWeight: FontWeight.bold, color: Colors.white),
                          ),
                        ],
                      ),
                      Column(
                        crossAxisAlignment: CrossAxisAlignment.start,
                        children: [
                          Text(
                            "Selesai",
                            style: TextStyle(fontSize: 11, color: Colors.white.withOpacity(0.75)),
                          ),
                          const SizedBox(height: 2),
                          Text(
                            endStr,
                            style: const TextStyle(fontSize: 13, fontWeight: FontWeight.bold, color: Colors.white),
                          ),
                        ],
                      ),
                      Column(
                        crossAxisAlignment: CrossAxisAlignment.end,
                        children: [
                          const SizedBox(height: 4),
                          // POIN 2: Durasi > 60 menit dibuat 2 baris (jam lalu menit)
                          _buildActivityCardDuration(duration, durationColor),
                        ],
                      ),
                    ],
                  ),
                  const SizedBox(height: 10),
                  Row(
                    crossAxisAlignment: CrossAxisAlignment.start,
                    children: [
                      const Padding(
                        padding: EdgeInsets.only(top: 2),
                        child: Icon(Icons.edit_note_rounded, color: Colors.white70, size: 15),
                      ),
                      const SizedBox(width: 4),
                      Expanded(
                        child: Text(
                          "Keperluan: ${item.purposeOrDesc}",
                          style: TextStyle(
                            fontSize: 11.5,
                            color: Colors.white.withOpacity(0.92),
                            fontWeight: FontWeight.w500,
                            height: 1.3,
                          ),
                          // POIN 3: Keperluan maksimal 3 baris
                          maxLines: 3,
                          overflow: TextOverflow.ellipsis,
                        ),
                      ),
                    ],
                  ),
                ],
              ),
            ),
          ],
        ),
      ),
    );
  }

  /// POIN 1: Card Aktivitas Khusus Laporan Cepu Yang Valid
  Widget _buildCepuActivityCard(UserActivityItem item) {
    final startStr = "${DateFormat('HH:mm').format(item.startTime)} WIT";
    final endStr = item.endTime != null ? "${DateFormat('HH:mm').format(item.endTime!)} WIT" : "--:--";
    final duration = item.getDurationMinutes();

    // Warna dinamis: Di-Cepu (Merah Crimson), Cepu Selesai (Deep Indigo)
    final gradientColors = item.isActive
        ? const [Color(0xFFDC2626), Color(0xFF991B1B)]
        : const [Color(0xFF4338CA), Color(0xFF312E81)];
    final shadowColor = item.isActive
        ? const Color(0xFFDC2626).withOpacity(0.28)
        : const Color(0xFF4338CA).withOpacity(0.25);
    final durationColor = item.isActive
        ? const Color(0xFFFDE68A)
        : const Color(0xFFE0E7FF);

    return InkWell(
      onTap: () => _showActivityDetailModal(item),
      borderRadius: BorderRadius.circular(22),
      child: Container(
        decoration: BoxDecoration(
          gradient: LinearGradient(
            colors: gradientColors,
            begin: Alignment.topLeft,
            end: Alignment.bottomRight,
          ),
          borderRadius: BorderRadius.circular(22),
          boxShadow: [
            BoxShadow(
              color: shadowColor,
              blurRadius: 14,
              offset: const Offset(0, 5),
            ),
          ],
        ),
        padding: const EdgeInsets.all(16),
        child: Row(
          children: [
            // Kotak Kiri: "Di-Cepu" atau "Selesai"
            Container(
              width: 72,
              padding: const EdgeInsets.symmetric(horizontal: 6, vertical: 12),
              decoration: BoxDecoration(
                color: Colors.white.withOpacity(0.2),
                borderRadius: BorderRadius.circular(16),
                border: Border.all(
                  color: Colors.white.withOpacity(0.35),
                  width: 1.2,
                ),
              ),
              child: Column(
                mainAxisSize: MainAxisSize.min,
                children: [
                  Icon(
                    item.isActive ? Icons.campaign_rounded : Icons.verified_rounded,
                    color: Colors.white,
                    size: 22,
                  ),
                  const SizedBox(height: 4),
                  Text(
                    item.isActive ? "Di-Cepu" : "Selesai",
                    textAlign: TextAlign.center,
                    style: const TextStyle(
                      fontSize: 11,
                      fontWeight: FontWeight.w900,
                      color: Colors.white,
                      height: 1.2,
                    ),
                  ),
                ],
              ),
            ),

            const SizedBox(width: 16),

            // Rincian Waktu & Durasi Menit / Jam
            Expanded(
              child: Column(
                crossAxisAlignment: CrossAxisAlignment.start,
                children: [
                  Row(
                    mainAxisAlignment: MainAxisAlignment.spaceBetween,
                    children: [
                      Column(
                        crossAxisAlignment: CrossAxisAlignment.start,
                        children: [
                          Text(
                            "Mulai",
                            style: TextStyle(fontSize: 11, color: Colors.white.withOpacity(0.75)),
                          ),
                          const SizedBox(height: 2),
                          Text(
                            startStr,
                            style: const TextStyle(fontSize: 13, fontWeight: FontWeight.bold, color: Colors.white),
                          ),
                        ],
                      ),
                      Column(
                        crossAxisAlignment: CrossAxisAlignment.start,
                        children: [
                          Text(
                            "Selesai",
                            style: TextStyle(fontSize: 11, color: Colors.white.withOpacity(0.75)),
                          ),
                          const SizedBox(height: 2),
                          Text(
                            endStr,
                            style: const TextStyle(fontSize: 13, fontWeight: FontWeight.bold, color: Colors.white),
                          ),
                        ],
                      ),
                      Column(
                        crossAxisAlignment: CrossAxisAlignment.end,
                        children: [
                          const SizedBox(height: 4),
                          // POIN 2: Durasi > 60 menit dibuat 2 baris
                          _buildActivityCardDuration(duration, durationColor),
                        ],
                      ),
                    ],
                  ),
                  const SizedBox(height: 10),
                  Row(
                    crossAxisAlignment: CrossAxisAlignment.start,
                    children: [
                      const Padding(
                        padding: EdgeInsets.only(top: 2),
                        child: Icon(Icons.person_pin_circle_rounded, color: Colors.white70, size: 15),
                      ),
                      const SizedBox(width: 4),
                      Expanded(
                        child: Text(
                          "Pelapor: ${item.reporterName ?? '-'} • Ket: ${item.purposeOrDesc}",
                          style: TextStyle(
                            fontSize: 11.5,
                            color: Colors.white.withOpacity(0.95),
                            fontWeight: FontWeight.w500,
                            height: 1.3,
                          ),
                          // POIN 3: Keperluan maksimal 3 baris
                          maxLines: 3,
                          overflow: TextOverflow.ellipsis,
                        ),
                      ),
                    ],
                  ),
                ],
              ),
            ),
          ],
        ),
      ),
    );
  }


  Widget _buildIconButton({
    required IconData icon,
    required VoidCallback onTap,
    bool hasBadge = false,
  }) {
    return Stack(
      children: [
        Container(
          width: 40,
          height: 40,
          decoration: BoxDecoration(
            color: Colors.white,
            borderRadius: BorderRadius.circular(12),
            border: Border.all(color: const Color(0xFFE2E8F0)),
            boxShadow: [
              BoxShadow(
                color: Colors.black.withOpacity(0.04),
                blurRadius: 6,
                offset: const Offset(0, 2),
              ),
            ],
          ),
          child: IconButton(
            icon: Icon(icon, color: const Color(0xFF475569), size: 20),
            onPressed: onTap,
            padding: EdgeInsets.zero,
          ),
        ),
        if (hasBadge)
          Positioned(
            top: 3,
            right: 3,
            child: Container(
              width: 9,
              height: 9,
              decoration: BoxDecoration(
                color: const Color(0xFFEF4444),
                shape: BoxShape.circle,
                border: Border.all(color: Colors.white, width: 1.5),
              ),
            ),
          ),
      ],
    );
  }
}
