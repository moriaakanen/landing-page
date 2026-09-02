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
        title: const Text("Konfirmasi Keluar", style: TextStyle(fontWeight: FontWeight.w800, fontSize: 16.5)),
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

  String _getGreeting() {
    final hour = DateTime.now().hour;
    if (hour < 11) return "Selamat Pagi";
    if (hour < 15) return "Selamat Siang";
    if (hour < 18) return "Selamat Sore";
    return "Selamat Malam";
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
      backgroundColor: const Color(0xFFF8FAFC),
      body: SafeArea(
        child: RefreshIndicator(
          onRefresh: _loadInitialData,
          color: const Color(0xFF2563EB),
          child: SingleChildScrollView(
            physics: const AlwaysScrollableScrollPhysics(),
            padding: const EdgeInsets.symmetric(horizontal: 20, vertical: 16),
            child: Column(
              crossAxisAlignment: CrossAxisAlignment.stretch,
              children: [
                // =========================================================================
                // 1. TOP HEADER: BRAND + PROFILE + ACTIONS (Clean Minimalist Style)
                // =========================================================================
                Row(
                  children: [
                    // Manta Ray App Logo in Soft Glowing Squircle
                    Container(
                      width: 52,
                      height: 52,
                      decoration: BoxDecoration(
                        color: Colors.white,
                        borderRadius: BorderRadius.circular(18),
                        border: Border.all(color: const Color(0xFFE2E8F0), width: 1.5),
                        boxShadow: [
                          BoxShadow(
                            color: const Color(0xFF2563EB).withOpacity(0.12),
                            blurRadius: 16,
                            offset: const Offset(0, 6),
                          ),
                        ],
                      ),
                      padding: const EdgeInsets.all(7),
                      child: Image.asset(
                        'assets/images/app_logo.png',
                        fit: BoxFit.contain,
                      ),
                    ),
                    const SizedBox(width: 14),

                    // Greeting & User Name
                    Expanded(
                      child: Column(
                        crossAxisAlignment: CrossAxisAlignment.start,
                        children: [
                          Row(
                            children: [
                              Text(
                                _getGreeting(),
                                style: const TextStyle(
                                  color: Color(0xFF64748B),
                                  fontSize: 12,
                                  fontWeight: FontWeight.w600,
                                ),
                              ),
                              const SizedBox(width: 4),
                              const Text("👋", style: TextStyle(fontSize: 12)),
                            ],
                          ),
                          const SizedBox(height: 2),
                          Text(
                            widget.user.name,
                            style: const TextStyle(
                              color: Color(0xFF0F172A),
                              fontSize: 17,
                              fontWeight: FontWeight.w900,
                              letterSpacing: -0.3,
                            ),
                            maxLines: 1,
                            overflow: TextOverflow.ellipsis,
                          ),
                        ],
                      ),
                    ),

                    // Refresh Button
                    _buildCircularActionBtn(
                      icon: Icons.refresh_rounded,
                      tooltip: "Perbarui Data",
                      onTap: _loadInitialData,
                    ),
                    const SizedBox(width: 8),

                    // Logout Button
                    _buildCircularActionBtn(
                      icon: Icons.logout_rounded,
                      tooltip: "Keluar Akun",
                      onTap: _handleLogout,
                      isDanger: true,
                    ),
                  ],
                ),

                const SizedBox(height: 18),

                // =========================================================================
                // 2. HERO "SMART PRESENCE HUB" CARD (Dark Modern Glass Bento)
                // =========================================================================
                Container(
                  decoration: BoxDecoration(
                    gradient: const LinearGradient(
                      colors: [
                        Color(0xFF0F172A),
                        Color(0xFF1E293B),
                        Color(0xFF1E3A8A),
                      ],
                      begin: Alignment.topLeft,
                      end: Alignment.bottomRight,
                    ),
                    borderRadius: BorderRadius.circular(28),
                    boxShadow: [
                      BoxShadow(
                        color: const Color(0xFF1E3A8A).withOpacity(0.25),
                        blurRadius: 24,
                        offset: const Offset(0, 10),
                      ),
                    ],
                  ),
                  padding: const EdgeInsets.all(22),
                  child: Column(
                    crossAxisAlignment: CrossAxisAlignment.stretch,
                    children: [
                      // Status Bar & Today's Date
                      Row(
                        mainAxisAlignment: MainAxisAlignment.spaceBetween,
                        children: [
                          // Live Status Pill
                          Container(
                            padding: const EdgeInsets.symmetric(horizontal: 10, vertical: 5),
                            decoration: BoxDecoration(
                              color: hasActivePermit
                                  ? const Color(0xFFD97706).withOpacity(0.3)
                                  : const Color(0xFF10B981).withOpacity(0.2),
                              borderRadius: BorderRadius.circular(20),
                              border: Border.all(
                                color: hasActivePermit
                                    ? const Color(0xFFF59E0B).withOpacity(0.6)
                                    : const Color(0xFF34D399).withOpacity(0.5),
                              ),
                            ),
                            child: Row(
                              mainAxisSize: MainAxisSize.min,
                              children: [
                                Container(
                                  width: 7,
                                  height: 7,
                                  decoration: BoxDecoration(
                                    shape: BoxShape.circle,
                                    color: hasActivePermit ? const Color(0xFFFBBF24) : const Color(0xFF34D399),
                                    boxShadow: [
                                      BoxShadow(
                                        color: hasActivePermit ? const Color(0xFFFBBF24) : const Color(0xFF34D399),
                                        blurRadius: 6,
                                      ),
                                    ],
                                  ),
                                ),
                                const SizedBox(width: 6),
                                Text(
                                  hasActivePermit ? "SEDANG IZIN" : "STANDBY DI KANTOR",
                                  style: TextStyle(
                                    fontSize: 10.5,
                                    fontWeight: FontWeight.w800,
                                    color: hasActivePermit ? const Color(0xFFFDE68A) : const Color(0xFFA7F3D0),
                                    letterSpacing: 0.5,
                                  ),
                                ),
                              ],
                            ),
                          ),

                          // Date Pill
                          Row(
                            children: [
                              const Icon(Icons.calendar_month_rounded, size: 13, color: Color(0xFF93C5FD)),
                              const SizedBox(width: 5),
                              Text(
                                todayFormatted,
                                style: TextStyle(
                                  fontSize: 11,
                                  color: Colors.white.withOpacity(0.8),
                                  fontWeight: FontWeight.w600,
                                ),
                              ),
                            ],
                          ),
                        ],
                      ),

                      const SizedBox(height: 20),

                      // Time Display Metrics
                      Row(
                        children: [
                          // Column 1: Mulai Izin
                          Expanded(
                            child: Column(
                              crossAxisAlignment: CrossAxisAlignment.start,
                              children: [
                                Text(
                                  "MULAI IZIN",
                                  style: TextStyle(
                                    fontSize: 10.5,
                                    fontWeight: FontWeight.w700,
                                    color: Colors.white.withOpacity(0.55),
                                    letterSpacing: 0.8,
                                  ),
                                ),
                                const SizedBox(height: 4),
                                Text(
                                  startPermitTime,
                                  style: TextStyle(
                                    fontSize: hasActivePermit ? 20 : 26,
                                    fontWeight: FontWeight.w900,
                                    color: hasActivePermit ? const Color(0xFFFDBA74) : Colors.white,
                                    letterSpacing: -0.5,
                                  ),
                                ),
                              ],
                            ),
                          ),

                          // Center Divider
                          Container(
                            height: 38,
                            width: 1,
                            color: Colors.white.withOpacity(0.15),
                          ),

                          const SizedBox(width: 18),

                          // Column 2: Selesai Izin
                          Expanded(
                            child: Column(
                              crossAxisAlignment: CrossAxisAlignment.start,
                              children: [
                                Text(
                                  "SELESAI IZIN",
                                  style: TextStyle(
                                    fontSize: 10.5,
                                    fontWeight: FontWeight.w700,
                                    color: Colors.white.withOpacity(0.55),
                                    letterSpacing: 0.8,
                                  ),
                                ),
                                const SizedBox(height: 4),
                                Text(
                                  "--:--",
                                  style: TextStyle(
                                    fontSize: 26,
                                    fontWeight: FontWeight.w900,
                                    color: Colors.white.withOpacity(0.4),
                                    letterSpacing: -0.5,
                                  ),
                                ),
                              ],
                            ),
                          ),
                        ],
                      ),

                      // Active Permit Purpose Footer (if any)
                      if (hasActivePermit) ...[
                        const SizedBox(height: 14),
                        Container(
                          padding: const EdgeInsets.symmetric(horizontal: 14, vertical: 10),
                          decoration: BoxDecoration(
                            color: Colors.white.withOpacity(0.1),
                            borderRadius: BorderRadius.circular(14),
                            border: Border.all(color: Colors.white.withOpacity(0.15)),
                          ),
                          child: Row(
                            children: [
                              const Icon(Icons.info_outline_rounded, color: Color(0xFFFDBA74), size: 16),
                              const SizedBox(width: 8),
                              Expanded(
                                child: Text(
                                  "Keperluan: ${_activePermit!.purpose}",
                                  style: const TextStyle(
                                    fontSize: 11.5,
                                    color: Colors.white,
                                    fontWeight: FontWeight.w500,
                                  ),
                                  maxLines: 1,
                                  overflow: TextOverflow.ellipsis,
                                ),
                              ),
                            ],
                          ),
                        ),
                      ],
                    ],
                  ),
                ),

                // POIN 7A: NOTIFIKASI / BANNER SAAT USER SEDANG IZIN
                if (hasActivePermit) ...[
                  const SizedBox(height: 14),
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
                            blurRadius: 12,
                            offset: const Offset(0, 4),
                          ),
                        ],
                      ),
                      child: Row(
                        children: [
                          Container(
                            padding: const EdgeInsets.all(8),
                            decoration: BoxDecoration(
                              color: const Color(0xFFF59E0B),
                              borderRadius: BorderRadius.circular(12),
                            ),
                            child: const Icon(Icons.directions_walk_rounded, color: Colors.white, size: 22),
                          ),
                          const SizedBox(width: 12),
                          Expanded(
                            child: Column(
                              crossAxisAlignment: CrossAxisAlignment.start,
                              children: const [
                                Text(
                                  "Anda sedang izin, segera rekam waktu kembali jika anda sudah berada di kantor",
                                  style: TextStyle(
                                    fontSize: 12,
                                    fontWeight: FontWeight.w800,
                                    color: Color(0xFF92400E),
                                    height: 1.35,
                                  ),
                                ),
                              ],
                            ),
                          ),
                          const SizedBox(width: 4),
                          const Icon(Icons.chevron_right_rounded, color: Color(0xFF92400E)),
                        ],
                      ),
                    ),
                  ),
                ],

                // POIN 7B & 8: NOTIFIKASI / BANNER SAAT USER DI-CEPU-KAN & SUDAH TERVERIFIKASI
                if (_activeCepuForMe != null) ...[
                  const SizedBox(height: 14),
                  InkWell(
                    onTap: _navigateToCepuMapReport,
                    borderRadius: BorderRadius.circular(20),
                    child: Container(
                      padding: const EdgeInsets.all(14),
                      decoration: BoxDecoration(
                        color: const Color(0xFFFEF2F2),
                        borderRadius: BorderRadius.circular(20),
                        border: Border.all(color: const Color(0xFFFCA5A5), width: 1.5),
                        boxShadow: [
                          BoxShadow(
                            color: const Color(0xFFDC2626).withOpacity(0.12),
                            blurRadius: 12,
                            offset: const Offset(0, 4),
                          ),
                        ],
                      ),
                      child: Row(
                        children: [
                          Container(
                            padding: const EdgeInsets.all(8),
                            decoration: BoxDecoration(
                              color: const Color(0xFFDC2626),
                              borderRadius: BorderRadius.circular(12),
                            ),
                            child: const Icon(Icons.warning_amber_rounded, color: Colors.white, size: 22),
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
                                    fontWeight: FontWeight.w800,
                                    color: Color(0xFF991B1B),
                                    height: 1.35,
                                  ),
                                ),
                                SizedBox(height: 2),
                                Text(
                                  "👉 Ketuk untuk rekam waktu kembali pada fitur Cepu",
                                  style: TextStyle(
                                    fontSize: 11,
                                    color: Color(0xFFDC2626),
                                    fontWeight: FontWeight.w700,
                                  ),
                                ),
                              ],
                            ),
                          ),
                          const SizedBox(width: 4),
                          const Icon(Icons.chevron_right_rounded, color: Color(0xFF991B1B)),
                        ],
                      ),
                    ),
                  ),
                ],

                const SizedBox(height: 24),

                // =========================================================================
                // 3. SECTION HEADER
                // =========================================================================
                Row(
                  children: const [
                    Text(
                      "Menu Utama",
                      style: TextStyle(
                        fontSize: 16,
                        fontWeight: FontWeight.w900,
                        color: Color(0xFF0F172A),
                        letterSpacing: -0.2,
                      ),
                    ),
                  ],
                ),

                const SizedBox(height: 12),

                // =========================================================================
                // 4. ASYMMETRIC BENTO GRID LAYOUT
                // =========================================================================

                // A. HERO BENTO CARD: WAIGAMA (Full Width Premium Hero Card)
                Material(
                  color: Colors.transparent,
                  child: InkWell(
                    onTap: _navigateToRecordPermit,
                    borderRadius: BorderRadius.circular(24),
                    child: Container(
                      decoration: BoxDecoration(
                        color: Colors.white,
                        borderRadius: BorderRadius.circular(24),
                        border: Border.all(color: const Color(0xFFDBEAFE), width: 1.5),
                        boxShadow: [
                          BoxShadow(
                            color: const Color(0xFF2563EB).withOpacity(0.08),
                            blurRadius: 20,
                            offset: const Offset(0, 6),
                          ),
                        ],
                      ),
                      padding: const EdgeInsets.all(18),
                      child: Column(
                        crossAxisAlignment: CrossAxisAlignment.stretch,
                        children: [
                          Row(
                            children: [
                              // Icon Squircle
                              Container(
                                width: 48,
                                height: 48,
                                decoration: BoxDecoration(
                                  gradient: const LinearGradient(
                                    colors: [Color(0xFF2563EB), Color(0xFF38BDF8)],
                                    begin: Alignment.topLeft,
                                    end: Alignment.bottomRight,
                                  ),
                                  borderRadius: BorderRadius.circular(16),
                                  boxShadow: [
                                    BoxShadow(
                                      color: const Color(0xFF2563EB).withOpacity(0.3),
                                      blurRadius: 10,
                                      offset: const Offset(0, 4),
                                    ),
                                  ],
                                ),
                                child: const Icon(Icons.assignment_ind_rounded, color: Colors.white, size: 26),
                              ),
                              const SizedBox(width: 14),

                              Expanded(
                                child: Column(
                                  crossAxisAlignment: CrossAxisAlignment.start,
                                  children: [
                                    Row(
                                      children: [
                                        const Text(
                                          "Waigama",
                                          style: TextStyle(
                                            fontSize: 18,
                                            fontWeight: FontWeight.w900,
                                            color: Color(0xFF0F172A),
                                          ),
                                        ),
                                        if (hasActivePermit) ...[
                                          const SizedBox(width: 8),
                                          Container(
                                            padding: const EdgeInsets.symmetric(horizontal: 7, vertical: 2),
                                            decoration: BoxDecoration(
                                              color: const Color(0xFFFEF3C7),
                                              borderRadius: BorderRadius.circular(6),
                                            ),
                                            child: const Text(
                                              "Sedang Izin",
                                              style: TextStyle(
                                                fontSize: 10,
                                                fontWeight: FontWeight.bold,
                                                color: Color(0xFFD97706),
                                              ),
                                            ),
                                          ),
                                        ],
                                      ],
                                    ),
                                    const SizedBox(height: 2),
                                    const Text(
                                      "Wajib Isi Guna Kebaikan Bersama",
                                      style: TextStyle(
                                        fontSize: 11.5,
                                        fontWeight: FontWeight.w600,
                                        color: Color(0xFF2563EB),
                                      ),
                                    ),
                                  ],
                                ),
                              ),

                              Container(
                                width: 34,
                                height: 34,
                                decoration: const BoxDecoration(
                                  color: Color(0xFFEFF6FF),
                                  shape: BoxShape.circle,
                                ),
                                child: const Icon(Icons.arrow_forward_rounded, color: Color(0xFF2563EB), size: 18),
                              ),
                            ],
                          ),
                          const SizedBox(height: 12),
                          const Divider(height: 1, color: Color(0xFFF1F5F9)),
                          const SizedBox(height: 10),
                          Row(
                            mainAxisAlignment: MainAxisAlignment.spaceBetween,
                            children: [
                              Text(
                                "Rekam izin keluar dinas atau urusan kantor",
                                style: TextStyle(
                                  fontSize: 11.5,
                                  color: Colors.grey.shade600,
                                  fontWeight: FontWeight.w500,
                                ),
                              ),
                              Text(
                                hasActivePermit ? "Selesaikan →" : "Mulai →",
                                style: const TextStyle(
                                  fontSize: 12,
                                  fontWeight: FontWeight.w800,
                                  color: Color(0xFF2563EB),
                                ),
                              ),
                            ],
                          ),
                        ],
                      ),
                    ),
                  ),
                ),

                const SizedBox(height: 14),

                // B. DUAL BENTO GRID: CEPU & MONITORING HARIAN (Side by Side)
                Row(
                  children: [
                    // Card 1: Cepu
                    Expanded(
                      child: Material(
                        color: Colors.transparent,
                        child: InkWell(
                          onTap: _navigateToCepuMapReport,
                          borderRadius: BorderRadius.circular(24),
                          child: Container(
                            height: 175,
                            decoration: BoxDecoration(
                              color: Colors.white,
                              borderRadius: BorderRadius.circular(24),
                              border: Border.all(color: const Color(0xFFFFEDD5), width: 1.5),
                              boxShadow: [
                                BoxShadow(
                                  color: const Color(0xFFEA580C).withOpacity(0.06),
                                  blurRadius: 18,
                                  offset: const Offset(0, 6),
                                ),
                              ],
                            ),
                            padding: const EdgeInsets.all(16),
                            child: Column(
                              crossAxisAlignment: CrossAxisAlignment.start,
                              mainAxisAlignment: MainAxisAlignment.spaceBetween,
                              children: [
                                Row(
                                  mainAxisAlignment: MainAxisAlignment.spaceBetween,
                                  children: [
                                    Container(
                                      width: 44,
                                      height: 44,
                                      decoration: BoxDecoration(
                                        gradient: const LinearGradient(
                                          colors: [Color(0xFFEA580C), Color(0xFFFB923C)],
                                          begin: Alignment.topLeft,
                                          end: Alignment.bottomRight,
                                        ),
                                        borderRadius: BorderRadius.circular(14),
                                        boxShadow: [
                                          BoxShadow(
                                            color: const Color(0xFFEA580C).withOpacity(0.3),
                                            blurRadius: 8,
                                            offset: const Offset(0, 3),
                                          ),
                                        ],
                                      ),
                                      child: const Icon(Icons.campaign_rounded, color: Colors.white, size: 24),
                                    ),
                                    Container(
                                      width: 28,
                                      height: 28,
                                      decoration: const BoxDecoration(
                                        color: Color(0xFFFFF7ED),
                                        shape: BoxShape.circle,
                                      ),
                                      child: const Icon(Icons.arrow_outward_rounded, color: Color(0xFFEA580C), size: 16),
                                    ),
                                  ],
                                ),
                                Column(
                                  crossAxisAlignment: CrossAxisAlignment.start,
                                  children: const [
                                    Text(
                                      "Cepu",
                                      style: TextStyle(
                                        fontSize: 17,
                                        fontWeight: FontWeight.w900,
                                        color: Color(0xFF0F172A),
                                      ),
                                    ),
                                    SizedBox(height: 3),
                                    Text(
                                      "Lapor & verifikasi keberadaan rekan",
                                      style: TextStyle(
                                        fontSize: 11,
                                        color: Color(0xFF64748B),
                                        height: 1.3,
                                      ),
                                      maxLines: 2,
                                    ),
                                  ],
                                ),
                              ],
                            ),
                          ),
                        ),
                      ),
                    ),

                    const SizedBox(width: 14),

                    // Card 2: Monitoring Harian
                    Expanded(
                      child: Material(
                        color: Colors.transparent,
                        child: InkWell(
                          onTap: _navigateToDailyMonitoring,
                          borderRadius: BorderRadius.circular(24),
                          child: Container(
                            height: 175,
                            decoration: BoxDecoration(
                              color: Colors.white,
                              borderRadius: BorderRadius.circular(24),
                              border: Border.all(color: const Color(0xFFCCFBF1), width: 1.5),
                              boxShadow: [
                                BoxShadow(
                                  color: const Color(0xFF0D9488).withOpacity(0.06),
                                  blurRadius: 18,
                                  offset: const Offset(0, 6),
                                ),
                              ],
                            ),
                            padding: const EdgeInsets.all(16),
                            child: Column(
                              crossAxisAlignment: CrossAxisAlignment.start,
                              mainAxisAlignment: MainAxisAlignment.spaceBetween,
                              children: [
                                Row(
                                  mainAxisAlignment: MainAxisAlignment.spaceBetween,
                                  children: [
                                    Container(
                                      width: 44,
                                      height: 44,
                                      decoration: BoxDecoration(
                                        gradient: const LinearGradient(
                                          colors: [Color(0xFF0D9488), Color(0xFF2DD4BF)],
                                          begin: Alignment.topLeft,
                                          end: Alignment.bottomRight,
                                        ),
                                        borderRadius: BorderRadius.circular(14),
                                        boxShadow: [
                                          BoxShadow(
                                            color: const Color(0xFF0D9488).withOpacity(0.3),
                                            blurRadius: 8,
                                            offset: const Offset(0, 3),
                                          ),
                                        ],
                                      ),
                                      child: const Icon(Icons.analytics_rounded, color: Colors.white, size: 24),
                                    ),
                                    Container(
                                      width: 28,
                                      height: 28,
                                      decoration: const BoxDecoration(
                                        color: Color(0xFFF0FDFA),
                                        shape: BoxShape.circle,
                                      ),
                                      child: const Icon(Icons.arrow_outward_rounded, color: Color(0xFF0D9488), size: 16),
                                    ),
                                  ],
                                ),
                                Column(
                                  crossAxisAlignment: CrossAxisAlignment.start,
                                  children: const [
                                    Text(
                                      "Monitoring",
                                      style: TextStyle(
                                        fontSize: 17,
                                        fontWeight: FontWeight.w900,
                                        color: Color(0xFF0F172A),
                                      ),
                                    ),
                                    SizedBox(height: 3),
                                    Text(
                                      "Rekap harian & status seluruh staf",
                                      style: TextStyle(
                                        fontSize: 11,
                                        color: Color(0xFF64748B),
                                        height: 1.3,
                                      ),
                                      maxLines: 2,
                                    ),
                                  ],
                                ),
                              ],
                            ),
                          ),
                        ),
                      ),
                    ),
                  ],
                ),

                const SizedBox(height: 24),

                // Quick Info Note Footer
                Container(
                  padding: const EdgeInsets.symmetric(horizontal: 16, vertical: 12),
                  decoration: BoxDecoration(
                    color: const Color(0xFFF1F5F9),
                    borderRadius: BorderRadius.circular(16),
                  ),
                  child: Row(
                    children: const [
                      Icon(Icons.shield_outlined, size: 18, color: Color(0xFF64748B)),
                      SizedBox(width: 10),
                      Expanded(
                        child: Text(
                          "Presensi Waigama diverifikasi otomatis dengan GPS & Radius Kantor 50 meter.",
                          style: TextStyle(fontSize: 11, color: Color(0xFF64748B), height: 1.35),
                        ),
                      ),
                    ],
                  ),
                ),

                const SizedBox(height: 20),
              ],
            ),
          ),
        ),
      ),
    );
  }

  Widget _buildCircularActionBtn({
    required IconData icon,
    required String tooltip,
    required VoidCallback onTap,
    bool isDanger = false,
  }) {
    return Container(
      decoration: BoxDecoration(
        color: isDanger ? const Color(0xFFFEF2F2) : const Color(0xFFF1F5F9),
        shape: BoxShape.circle,
        border: Border.all(
          color: isDanger ? const Color(0xFFFECACA) : const Color(0xFFE2E8F0),
        ),
      ),
      child: IconButton(
        icon: Icon(
          icon,
          color: isDanger ? const Color(0xFFEF4444) : const Color(0xFF475569),
          size: 19,
        ),
        onPressed: onTap,
        tooltip: tooltip,
        constraints: const BoxConstraints(minWidth: 40, minHeight: 40),
        padding: EdgeInsets.zero,
      ),
    );
  }
}
