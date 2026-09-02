import 'dart:async';
import 'package:flutter/material.dart';
import 'package:intl/intl.dart';
import '../models/user_model.dart';
import '../models/office_model.dart';
import '../models/permit_model.dart';
import '../models/cepu_model.dart';
import '../core/services/attendance_service.dart';
import '../core/services/permit_service.dart';
import '../core/services/cepu_service.dart';
import '../core/services/auth_service.dart';
import '../core/services/notification_service.dart';
import '../core/utils/custom_toast.dart';
import 'login_view.dart';
import 'record_permit_map_view.dart';
import 'cepu_map_report_view.dart';
import 'daily_monitoring_view.dart';

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
  bool _isLoading = true;
  int _currentNavIndex = 0;

  StreamSubscription? _cepuNotifSub;
  final Set<String> _seenNotifIds = {};

  @override
  void initState() {
    super.initState();
    _loadInitialData();
    _listenToNewCepuNotifications();
  }

  @override
  void dispose() {
    _cepuNotifSub?.cancel();
    super.dispose();
  }

  Future<void> _loadInitialData() async {
    setState(() => _isLoading = true);
    final office = await _attendanceService.getOfficeConfig(widget.user.officeId);
    final activePermit = await _permitService.getActivePermit(widget.user.uid);
    final activeCepu = await _cepuService.getActiveCepuForTarget(widget.user.uid);

    if (mounted) {
      setState(() {
        _office = office;
        _activePermit = activePermit;
        _activeCepuForMe = activeCepu;
        _isLoading = false;
      });

      if (activePermit != null) {
        NotificationService.showSystemNotification(
          id: 101,
          title: "Izin Waigama Aktif",
          body: "Anda sedang izin, segera rekam waktu kembali jika anda sudah berada di kantor",
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
        builder: (context) => DailyMonitoringView(user: widget.user),
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
                color: const Color(0xFFE0F2F1),
                shape: BoxShape.circle,
                border: Border.all(color: const Color(0xFF26A69A), width: 2),
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
                    color: const Color(0xFFE0F2F1),
                    borderRadius: BorderRadius.circular(8),
                  ),
                  child: const Text("Pegawai Aktif", style: TextStyle(fontSize: 11, fontWeight: FontWeight.bold, color: Color(0xFF00796B))),
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

  @override
  Widget build(BuildContext context) {
    String todayFormatted;
    try {
      todayFormatted = DateFormat('EEEE, d MMMM yyyy', 'id_ID').format(DateTime.now());
    } catch (_) {
      todayFormatted = DateFormat('EEEE, d MMMM yyyy').format(DateTime.now());
    }

    final hasActivePermit = _activePermit != null;
    final startPermitTime = hasActivePermit
        ? "${DateFormat('HH:mm').format(_activePermit!.startTime)} WIT"
        : "--:--";

    return Scaffold(
      backgroundColor: const Color(0xFFF4F6F8),
      body: SafeArea(
        child: Stack(
          children: [
            RefreshIndicator(
              onRefresh: _loadInitialData,
              color: const Color(0xFF26A69A),
              child: SingleChildScrollView(
                physics: const AlwaysScrollableScrollPhysics(),
                padding: const EdgeInsets.fromLTRB(20, 14, 20, 100),
                child: Column(
                  crossAxisAlignment: CrossAxisAlignment.stretch,
                  children: [
                    // =========================================================================
                    // 1. TOP BAR (Sesuai Screenshot: Brand + Bell Notifikasi di Kanan)
                    // =========================================================================
                    Row(
                      mainAxisAlignment: MainAxisAlignment.spaceBetween,
                      children: [
                        // Brand Logo & Title
                        Row(
                          children: [
                            Container(
                              width: 38,
                              height: 38,
                              decoration: BoxDecoration(
                                color: Colors.white,
                                borderRadius: BorderRadius.circular(12),
                                border: Border.all(color: const Color(0xFF26A69A).withOpacity(0.3), width: 1.5),
                                boxShadow: [
                                  BoxShadow(
                                    color: const Color(0xFF26A69A).withOpacity(0.12),
                                    blurRadius: 10,
                                    offset: const Offset(0, 4),
                                  ),
                                ],
                              ),
                              padding: const EdgeInsets.all(5),
                              child: Image.asset('assets/images/app_logo.png', fit: BoxFit.contain),
                            ),
                            const SizedBox(width: 10),
                            Column(
                              crossAxisAlignment: CrossAxisAlignment.start,
                              children: [
                                RichText(
                                  text: const TextSpan(
                                    children: [
                                      TextSpan(
                                        text: "Waigama",
                                        style: TextStyle(
                                          fontSize: 20,
                                          fontWeight: FontWeight.w900,
                                          color: Color(0xFF26A69A),
                                          letterSpacing: -0.5,
                                        ),
                                      ),
                                      TextSpan(
                                        text: "App",
                                        style: TextStyle(
                                          fontSize: 20,
                                          fontWeight: FontWeight.w800,
                                          color: Color(0xFFFF7043),
                                          letterSpacing: -0.5,
                                        ),
                                      ),
                                    ],
                                  ),
                                ),
                              ],
                            ),
                          ],
                        ),

                        // Action Buttons: Notification Bell & Refresh
                        Row(
                          children: [
                            _buildTopRoundButton(
                              icon: Icons.refresh_rounded,
                              onTap: _loadInitialData,
                              iconColor: const Color(0xFF546E7A),
                            ),
                            const SizedBox(width: 8),
                            _buildTopRoundButton(
                              icon: Icons.notifications_active_rounded,
                              onTap: _navigateToDailyMonitoring,
                              iconColor: const Color(0xFFFF7043),
                              bgColor: const Color(0xFFFFEBEE),
                              hasBadge: hasActivePermit || _activeCepuForMe != null,
                            ),
                          ],
                        ),
                      ],
                    ),

                    const SizedBox(height: 16),

                    // =========================================================================
                    // 2. LOCATION & USER CARD (Sesuai Screenshot: Icon Bulat Teal + Address/Status)
                    // =========================================================================
                    InkWell(
                      onTap: _showProfileDialog,
                      borderRadius: BorderRadius.circular(20),
                      child: Container(
                        padding: const EdgeInsets.symmetric(horizontal: 14, vertical: 12),
                        decoration: BoxDecoration(
                          color: Colors.white,
                          borderRadius: BorderRadius.circular(20),
                          boxShadow: [
                            BoxShadow(
                              color: Colors.black.withOpacity(0.04),
                              blurRadius: 14,
                              offset: const Offset(0, 4),
                            ),
                          ],
                        ),
                        child: Row(
                          children: [
                            // Teal Circular Location Pin Container
                            Container(
                              width: 44,
                              height: 44,
                              decoration: const BoxDecoration(
                                color: Color(0xFF26A69A),
                                shape: BoxShape.circle,
                              ),
                              child: const Icon(Icons.location_on_rounded, color: Colors.white, size: 22),
                            ),
                            const SizedBox(width: 12),

                            Expanded(
                              child: Column(
                                crossAxisAlignment: CrossAxisAlignment.start,
                                children: [
                                  Text(
                                    widget.user.name,
                                    style: const TextStyle(
                                      fontSize: 14.5,
                                      fontWeight: FontWeight.w800,
                                      color: Color(0xFF1E293B),
                                    ),
                                    maxLines: 1,
                                    overflow: TextOverflow.ellipsis,
                                  ),
                                  const SizedBox(height: 2),
                                  Text(
                                    _office != null
                                        ? "${_office!.name} • Radius ${_office!.radiusMeters.toInt()}m"
                                        : "Kantor Utama • WIT",
                                    style: const TextStyle(
                                      fontSize: 11.5,
                                      color: Color(0xFF64748B),
                                      fontWeight: FontWeight.w500,
                                    ),
                                  ),
                                ],
                              ),
                            ),

                            const Icon(Icons.chevron_right_rounded, color: Color(0xFF94A3B8), size: 24),
                          ],
                        ),
                      ),
                    ),

                    const SizedBox(height: 18),

                    // =========================================================================
                    // 3. HERO CAROUSEL BANNER (Sesuai Screenshot: Card Banner Ilustrasi)
                    // =========================================================================
                    Container(
                      height: 165,
                      decoration: BoxDecoration(
                        gradient: const LinearGradient(
                          colors: [Color(0xFFE0F2F1), Color(0xFFB2DFDB)],
                          begin: Alignment.topLeft,
                          end: Alignment.bottomRight,
                        ),
                        borderRadius: BorderRadius.circular(24),
                        boxShadow: [
                          BoxShadow(
                            color: const Color(0xFF26A69A).withOpacity(0.12),
                            blurRadius: 16,
                            offset: const Offset(0, 6),
                          ),
                        ],
                      ),
                      child: Stack(
                        children: [
                          // Background Organic Circles
                          Positioned(
                            top: -20,
                            right: -20,
                            child: Container(
                              width: 140,
                              height: 140,
                              decoration: BoxDecoration(
                                shape: BoxShape.circle,
                                color: Colors.white.withOpacity(0.35),
                              ),
                            ),
                          ),
                          Positioned(
                            bottom: -30,
                            left: 100,
                            child: Container(
                              width: 100,
                              height: 100,
                              decoration: BoxDecoration(
                                shape: BoxShape.circle,
                                color: const Color(0xFF26A69A).withOpacity(0.15),
                              ),
                            ),
                          ),

                          // Image Illustration on Right
                          Positioned(
                            right: 12,
                            bottom: 10,
                            top: 10,
                            child: Image.asset(
                              'assets/images/app_logo.png',
                              width: 130,
                              fit: BoxFit.contain,
                            ),
                          ),

                          // Text Content on Left
                          Padding(
                            padding: const EdgeInsets.all(20),
                            child: Column(
                              crossAxisAlignment: CrossAxisAlignment.start,
                              mainAxisAlignment: MainAxisAlignment.center,
                              children: [
                                Container(
                                  padding: const EdgeInsets.symmetric(horizontal: 10, vertical: 4),
                                  decoration: BoxDecoration(
                                    color: Colors.white.withOpacity(0.85),
                                    borderRadius: BorderRadius.circular(12),
                                  ),
                                  child: Text(
                                    hasActivePermit ? "IZIN AKTIF" : "SISTEM PRESENSI",
                                    style: TextStyle(
                                      fontSize: 10.5,
                                      fontWeight: FontWeight.w900,
                                      color: hasActivePermit ? const Color(0xFFD97706) : const Color(0xFF00796B),
                                      letterSpacing: 0.5,
                                    ),
                                  ),
                                ),
                                const SizedBox(height: 8),
                                const Text(
                                  "Waigama",
                                  style: TextStyle(
                                    fontSize: 22,
                                    fontWeight: FontWeight.w900,
                                    color: Color(0xFF004D40),
                                    letterSpacing: -0.5,
                                  ),
                                ),
                                const SizedBox(height: 3),
                                SizedBox(
                                  width: 180,
                                  child: Text(
                                    "Wajib Isi Guna Kebaikan Bersama",
                                    style: TextStyle(
                                      fontSize: 12,
                                      fontWeight: FontWeight.w600,
                                      color: const Color(0xFF00695C).withOpacity(0.9),
                                      height: 1.25,
                                    ),
                                  ),
                                ),
                              ],
                            ),
                          ),
                        ],
                      ),
                    ),

                    // POIN 7A: NOTIFIKASI PUSH / BANNER SAAT USER SEDANG IZIN
                    if (hasActivePermit) ...[
                      const SizedBox(height: 16),
                      InkWell(
                        onTap: _navigateToRecordPermit,
                        borderRadius: BorderRadius.circular(20),
                        child: Container(
                          padding: const EdgeInsets.all(14),
                          decoration: BoxDecoration(
                            color: const Color(0xFFFFFBEB),
                            borderRadius: BorderRadius.circular(20),
                            border: Border.all(color: const Color(0xFFFCD34D), width: 1.5),
                            boxShadow: [
                              BoxShadow(
                                color: const Color(0xFFF59E0B).withOpacity(0.12),
                                blurRadius: 10,
                                offset: const Offset(0, 4),
                              ),
                            ],
                          ),
                          child: Row(
                            children: [
                              Container(
                                padding: const EdgeInsets.all(8),
                                decoration: const BoxDecoration(
                                  color: Color(0xFFF59E0B),
                                  shape: BoxShape.circle,
                                ),
                                child: const Icon(Icons.directions_walk_rounded, color: Colors.white, size: 20),
                              ),
                              const SizedBox(width: 12),
                              Expanded(
                                child: Column(
                                  crossAxisAlignment: CrossAxisAlignment.start,
                                  children: [
                                    const Text(
                                      "Anda sedang izin, segera rekam waktu kembali jika anda sudah berada di kantor",
                                      style: TextStyle(
                                        fontSize: 12,
                                        fontWeight: FontWeight.bold,
                                        color: Color(0xFF92400E),
                                        height: 1.3,
                                      ),
                                    ),
                                    const SizedBox(height: 3),
                                    Text(
                                      "Keperluan: ${_activePermit!.purpose}",
                                      style: const TextStyle(fontSize: 11, color: Color(0xFF1E293B)),
                                      maxLines: 1,
                                      overflow: TextOverflow.ellipsis,
                                    ),
                                  ],
                                ),
                              ),
                              const Icon(Icons.chevron_right_rounded, color: Color(0xFF92400E)),
                            ],
                          ),
                        ),
                      ),
                    ],

                    // POIN 7B & 8: NOTIFIKASI PUSH / BANNER SAAT USER DI-CEPU-KAN & SUDAH TERVERIFIKASI
                    if (_activeCepuForMe != null) ...[
                      const SizedBox(height: 16),
                      InkWell(
                        onTap: _navigateToCepuMapReport,
                        borderRadius: BorderRadius.circular(20),
                        child: Container(
                          padding: const EdgeInsets.all(14),
                          decoration: BoxDecoration(
                            color: const Color(0xFFFEF2F2),
                            borderRadius: BorderRadius.circular(20),
                            border: Border.all(color: const Color(0xFFF87171), width: 1.5),
                            boxShadow: [
                              BoxShadow(
                                color: const Color(0xFFDC2626).withOpacity(0.12),
                                blurRadius: 10,
                                offset: const Offset(0, 4),
                              ),
                            ],
                          ),
                          child: Row(
                            children: [
                              Container(
                                padding: const EdgeInsets.all(8),
                                decoration: const BoxDecoration(
                                  color: Color(0xFFDC2626),
                                  shape: BoxShape.circle,
                                ),
                                child: const Icon(Icons.warning_amber_rounded, color: Colors.white, size: 20),
                              ),
                              const SizedBox(width: 12),
                              Expanded(
                                child: Column(
                                  crossAxisAlignment: CrossAxisAlignment.start,
                                  children: const [
                                    Text(
                                      "Anda dilaporkan tidak berada di kantor, segera rekam waktu kembali jika anda berada di kantor",
                                      style: TextStyle(
                                        fontSize: 12,
                                        fontWeight: FontWeight.bold,
                                        color: Color(0xFF991B1B),
                                        height: 1.3,
                                      ),
                                    ),
                                    SizedBox(height: 3),
                                    Text(
                                      "👉 Ketuk untuk rekam waktu kembali pada fitur Cepu",
                                      style: TextStyle(
                                        fontSize: 11,
                                        color: Color(0xFFDC2626),
                                        fontWeight: FontWeight.w600,
                                      ),
                                    ),
                                  ],
                                ),
                              ),
                              const Icon(Icons.chevron_right_rounded, color: Color(0xFF991B1B)),
                            ],
                          ),
                        ),
                      ),
                    ],

                    const SizedBox(height: 22),

                    // =========================================================================
                    // 4. CATEGORIES / FITUR UTAMA (Sesuai Screenshot: Pastel Squircle Grid Cards)
                    // =========================================================================
                    Row(
                      mainAxisAlignment: MainAxisAlignment.spaceBetween,
                      children: [
                        const Text(
                          "Fitur Utama",
                          style: TextStyle(
                            fontSize: 17,
                            fontWeight: FontWeight.w900,
                            color: Color(0xFF1E293B),
                          ),
                        ),
                        const Icon(Icons.arrow_forward_rounded, color: Color(0xFF64748B), size: 20),
                      ],
                    ),

                    const SizedBox(height: 14),

                    // Horizontal / Grid Row of Pastel Squircle Cards
                    Row(
                      mainAxisAlignment: MainAxisAlignment.spaceBetween,
                      children: [
                        // Card 1: Waigama (Teal Squircle)
                        _buildCategorySquircleCard(
                          title: "Waigama",
                          icon: Icons.assignment_ind_rounded,
                          bgColor: const Color(0xFF26A69A),
                          onTap: _navigateToRecordPermit,
                          badgeText: hasActivePermit ? "Aktif" : null,
                        ),

                        // Card 2: Cepu (Sky Blue Squircle)
                        _buildCategorySquircleCard(
                          title: "Cepu",
                          icon: Icons.campaign_rounded,
                          bgColor: const Color(0xFF42A5F5),
                          onTap: _navigateToCepuMapReport,
                        ),

                        // Card 3: Monitoring (Purple/Pink Squircle)
                        _buildCategorySquircleCard(
                          title: "Monitoring",
                          icon: Icons.analytics_rounded,
                          bgColor: const Color(0xFFAB47BC),
                          onTap: _navigateToDailyMonitoring,
                        ),

                        // Card 4: Profil / Akun (Violet Squircle)
                        _buildCategorySquircleCard(
                          title: "Profil",
                          icon: Icons.person_rounded,
                          bgColor: const Color(0xFF7E57C2),
                          onTap: _showProfileDialog,
                        ),
                      ],
                    ),

                    const SizedBox(height: 24),

                    // =========================================================================
                    // 5. PREVIOUS STATUS CARD (Sesuai Screenshot: "Previous Order" Card Style)
                    // =========================================================================
                    Row(
                      mainAxisAlignment: MainAxisAlignment.spaceBetween,
                      children: [
                        const Text(
                          "Status Presensi Hari Ini",
                          style: TextStyle(
                            fontSize: 16,
                            fontWeight: FontWeight.w900,
                            color: Color(0xFF1E293B),
                          ),
                        ),
                        Text(
                          todayFormatted,
                          style: const TextStyle(fontSize: 11, color: Color(0xFF64748B), fontWeight: FontWeight.w600),
                        ),
                      ],
                    ),

                    const SizedBox(height: 12),

                    Container(
                      decoration: BoxDecoration(
                        color: Colors.white,
                        borderRadius: BorderRadius.circular(22),
                        boxShadow: [
                          BoxShadow(
                            color: Colors.black.withOpacity(0.04),
                            blurRadius: 16,
                            offset: const Offset(0, 4),
                          ),
                        ],
                      ),
                      padding: const EdgeInsets.all(18),
                      child: Column(
                        children: [
                          Row(
                            children: [
                              Container(
                                width: 44,
                                height: 44,
                                decoration: BoxDecoration(
                                  color: hasActivePermit ? const Color(0xFFFFF7ED) : const Color(0xFFE0F2F1),
                                  borderRadius: BorderRadius.circular(14),
                                ),
                                child: Icon(
                                  hasActivePermit ? Icons.directions_run_rounded : Icons.check_circle_rounded,
                                  color: hasActivePermit ? const Color(0xFFEA580C) : const Color(0xFF00897B),
                                  size: 24,
                                ),
                              ),
                              const SizedBox(width: 14),
                              Expanded(
                                child: Column(
                                  crossAxisAlignment: CrossAxisAlignment.start,
                                  children: [
                                    Text(
                                      hasActivePermit ? "Status: Sedang di Luar (Izin)" : "Status: Standby / Di Kantor",
                                      style: TextStyle(
                                        fontSize: 13.5,
                                        fontWeight: FontWeight.w800,
                                        color: hasActivePermit ? const Color(0xFFC2410C) : const Color(0xFF00796B),
                                      ),
                                    ),
                                    const SizedBox(height: 3),
                                    Text(
                                      "Mulai Izin: $startPermitTime  •  Selesai: --:--",
                                      style: const TextStyle(fontSize: 12, color: Color(0xFF64748B), fontWeight: FontWeight.w500),
                                    ),
                                  ],
                                ),
                              ),
                              Container(
                                padding: const EdgeInsets.symmetric(horizontal: 10, vertical: 4),
                                decoration: BoxDecoration(
                                  color: hasActivePermit ? const Color(0xFFFFEDD5) : const Color(0xFFE0F2F1),
                                  borderRadius: BorderRadius.circular(10),
                                ),
                                child: Text(
                                  hasActivePermit ? "Aktif" : "Normal",
                                  style: TextStyle(
                                    fontSize: 11,
                                    fontWeight: FontWeight.bold,
                                    color: hasActivePermit ? const Color(0xFFEA580C) : const Color(0xFF00796B),
                                  ),
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

            // =========================================================================
            // 6. BOTTOM NAVIGATION BAR (Sesuai Screenshot: Floating FAB Orange di Tengah)
            // =========================================================================
            Positioned(
              bottom: 0,
              left: 0,
              right: 0,
              child: Container(
                height: 78,
                decoration: BoxDecoration(
                  color: Colors.white,
                  borderRadius: const BorderRadius.vertical(top: Radius.circular(28)),
                  boxShadow: [
                    BoxShadow(
                      color: Colors.black.withOpacity(0.08),
                      blurRadius: 20,
                      offset: const Offset(0, -4),
                    ),
                  ],
                ),
                child: Row(
                  mainAxisAlignment: MainAxisAlignment.spaceAround,
                  children: [
                    // Nav 1: Home (Active Teal)
                    IconButton(
                      icon: const Icon(Icons.home_rounded, color: Color(0xFF26A69A), size: 26),
                      onPressed: () {},
                    ),

                    // Nav 2: Monitoring Grid
                    IconButton(
                      icon: const Icon(Icons.grid_view_rounded, color: Color(0xFF94A3B8), size: 24),
                      onPressed: _navigateToDailyMonitoring,
                    ),

                    // Center Floating Action Button (Orange Circle like screenshot!)
                    Transform.translate(
                      offset: const Offset(0, -18),
                      child: InkWell(
                        onTap: _navigateToRecordPermit,
                        borderRadius: BorderRadius.circular(30),
                        child: Container(
                          width: 58,
                          height: 58,
                          decoration: BoxDecoration(
                            gradient: const LinearGradient(
                              colors: [Color(0xFFFF8A65), Color(0xFFFF7043)],
                              begin: Alignment.topLeft,
                              end: Alignment.bottomRight,
                            ),
                            shape: BoxShape.circle,
                            boxShadow: [
                              BoxShadow(
                                color: const Color(0xFFFF7043).withOpacity(0.4),
                                blurRadius: 14,
                                offset: const Offset(0, 6),
                              ),
                            ],
                          ),
                          child: const Icon(Icons.assignment_ind_rounded, color: Colors.white, size: 28),
                        ),
                      ),
                    ),

                    // Nav 4: Cepu / Campaign
                    IconButton(
                      icon: const Icon(Icons.campaign_outlined, color: Color(0xFF94A3B8), size: 26),
                      onPressed: _navigateToCepuMapReport,
                    ),

                    // Nav 5: Profile / Logout
                    IconButton(
                      icon: const Icon(Icons.person_outline_rounded, color: Color(0xFF94A3B8), size: 26),
                      onPressed: _showProfileDialog,
                    ),
                  ],
                ),
              ),
            ),
          ],
        ),
      ),
    );
  }

  Widget _buildTopRoundButton({
    required IconData icon,
    required VoidCallback onTap,
    Color iconColor = const Color(0xFF475569),
    Color bgColor = Colors.white,
    bool hasBadge = false,
  }) {
    return Stack(
      children: [
        Container(
          width: 42,
          height: 42,
          decoration: BoxDecoration(
            color: bgColor,
            shape: BoxShape.circle,
            boxShadow: [
              BoxShadow(
                color: Colors.black.withOpacity(0.05),
                blurRadius: 10,
                offset: const Offset(0, 2),
              ),
            ],
          ),
          child: IconButton(
            icon: Icon(icon, color: iconColor, size: 20),
            onPressed: onTap,
            padding: EdgeInsets.zero,
          ),
        ),
        if (hasBadge)
          Positioned(
            top: 4,
            right: 4,
            child: Container(
              width: 10,
              height: 10,
              decoration: BoxDecoration(
                color: const Color(0xFFEF4444),
                shape: BoxShape.circle,
                border: Border.all(color: Colors.white, width: 2),
              ),
            ),
          ),
      ],
    );
  }

  Widget _buildCategorySquircleCard({
    required String title,
    required IconData icon,
    required Color bgColor,
    required VoidCallback onTap,
    String? badgeText,
  }) {
    return Material(
      color: Colors.transparent,
      child: InkWell(
        onTap: onTap,
        borderRadius: BorderRadius.circular(22),
        child: Column(
          children: [
            Stack(
              clipBehavior: Clip.none,
              children: [
                Container(
                  width: 70,
                  height: 70,
                  decoration: BoxDecoration(
                    color: bgColor,
                    borderRadius: BorderRadius.circular(22),
                    boxShadow: [
                      BoxShadow(
                        color: bgColor.withOpacity(0.3),
                        blurRadius: 12,
                        offset: const Offset(0, 5),
                      ),
                    ],
                  ),
                  child: Icon(icon, color: Colors.white, size: 30),
                ),
                if (badgeText != null)
                  Positioned(
                    top: -4,
                    right: -4,
                    child: Container(
                      padding: const EdgeInsets.symmetric(horizontal: 6, vertical: 2),
                      decoration: BoxDecoration(
                        color: const Color(0xFFEF4444),
                        borderRadius: BorderRadius.circular(8),
                        border: Border.all(color: Colors.white, width: 1.5),
                      ),
                      child: Text(
                        badgeText,
                        style: const TextStyle(color: Colors.white, fontSize: 9, fontWeight: FontWeight.bold),
                      ),
                    ),
                  ),
              ],
            ),
            const SizedBox(height: 8),
            Text(
              title,
              style: const TextStyle(
                fontSize: 12,
                fontWeight: FontWeight.w700,
                color: Color(0xFF1E293B),
              ),
            ),
          ],
        ),
      ),
    );
  }
}
