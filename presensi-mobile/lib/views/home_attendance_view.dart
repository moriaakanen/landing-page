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
    }
  }

  /// Poin 11: Real-time broadcast notification untuk seluruh user ketika ada laporan Cepu baru
  void _listenToNewCepuNotifications() {
    _cepuNotifSub = NotificationService.streamRecentCepuNotifications().listen((snapshot) {
      for (final doc in snapshot.docs) {
        if (!_seenNotifIds.contains(doc.id)) {
          _seenNotifIds.add(doc.id);
          final data = doc.data() as Map<String, dynamic>;
          final String title = data['title'] ?? '🚨 Laporan Cepu Baru';
          final String message = data['message'] ?? 'Ada laporan baru yang membutuhkan verifikasi rekan.';
          final String targetUid = data['target_uid'] ?? '';

          // Jangan munculkan popup banner jika ini adalah event awal saat buka app pertama kali
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
    await _authService.logout();
    if (mounted) {
      Navigator.pushReplacement(
        context,
        MaterialPageRoute(builder: (context) => const LoginView()),
      );
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
    const endPermitTime = "--:--";

    return Scaffold(
      backgroundColor: const Color(0xFFF1F5F9),
      body: Stack(
        children: [
          // Background Header
          Positioned(
            top: 0,
            left: 0,
            right: 0,
            height: 250,
            child: Container(
              decoration: const BoxDecoration(
                gradient: LinearGradient(
                  colors: [Color(0xFF0F172A), Color(0xFF1E293B), Color(0xFF1E3A8A)],
                  begin: Alignment.topLeft,
                  end: Alignment.bottomRight,
                ),
                borderRadius: BorderRadius.vertical(bottom: Radius.circular(32)),
              ),
              child: Stack(
                children: [
                  Positioned(
                    top: -40,
                    right: -40,
                    child: Container(
                      width: 180,
                      height: 180,
                      decoration: BoxDecoration(
                        shape: BoxShape.circle,
                        color: const Color(0xFF3B82F6).withOpacity(0.18),
                      ),
                    ),
                  ),
                  Positioned(
                    bottom: 20,
                    left: -30,
                    child: Container(
                      width: 140,
                      height: 140,
                      decoration: BoxDecoration(
                        shape: BoxShape.circle,
                        color: const Color(0xFF6366F1).withOpacity(0.15),
                      ),
                    ),
                  ),
                ],
              ),
            ),
          ),

          SafeArea(
            child: RefreshIndicator(
              onRefresh: _loadInitialData,
              child: SingleChildScrollView(
                physics: const AlwaysScrollableScrollPhysics(),
                padding: const EdgeInsets.symmetric(horizontal: 20, vertical: 12),
                child: Column(
                  crossAxisAlignment: CrossAxisAlignment.stretch,
                  children: [
                    // Top Profile Bar (Poin 1: Hapus Department)
                    Row(
                      children: [
                        Container(
                          padding: const EdgeInsets.all(2),
                          decoration: BoxDecoration(
                            shape: BoxShape.circle,
                            border: Border.all(color: const Color(0xFF60A5FA), width: 2),
                          ),
                          child: CircleAvatar(
                            radius: 20,
                            backgroundColor: const Color(0xFF334155),
                            child: Text(
                              widget.user.name.isNotEmpty ? widget.user.name[0].toUpperCase() : 'U',
                              style: const TextStyle(
                                color: Colors.white,
                                fontWeight: FontWeight.bold,
                                fontSize: 16,
                              ),
                            ),
                          ),
                        ),
                        const SizedBox(width: 12),
                        Expanded(
                          child: Column(
                            crossAxisAlignment: CrossAxisAlignment.start,
                            children: [
                              Text(
                                widget.user.name,
                                style: const TextStyle(
                                  color: Colors.white,
                                  fontSize: 16,
                                  fontWeight: FontWeight.w800,
                                  letterSpacing: 0.2,
                                ),
                                maxLines: 1,
                                overflow: TextOverflow.ellipsis,
                              ),
                              const SizedBox(height: 2),
                              Text(
                                "@${widget.user.username}",
                                style: TextStyle(
                                  color: Colors.white.withOpacity(0.75),
                                  fontSize: 12,
                                  fontWeight: FontWeight.w500,
                                ),
                              ),
                            ],
                          ),
                        ),
                        Container(
                          decoration: BoxDecoration(
                            color: Colors.white.withOpacity(0.12),
                            borderRadius: BorderRadius.circular(12),
                          ),
                          child: IconButton(
                            icon: const Icon(Icons.refresh_rounded, color: Colors.white, size: 20),
                            onPressed: _loadInitialData,
                            tooltip: "Perbarui Data",
                            constraints: const BoxConstraints(minWidth: 38, minHeight: 38),
                            padding: EdgeInsets.zero,
                          ),
                        ),
                        const SizedBox(width: 8),
                        Container(
                          decoration: BoxDecoration(
                            color: Colors.white.withOpacity(0.12),
                            borderRadius: BorderRadius.circular(12),
                          ),
                          child: IconButton(
                            icon: const Icon(Icons.logout_rounded, color: Colors.white, size: 20),
                            onPressed: _handleLogout,
                            tooltip: "Keluar",
                            constraints: const BoxConstraints(minWidth: 38, minHeight: 38),
                            padding: EdgeInsets.zero,
                          ),
                        ),
                      ],
                    ),

                    const SizedBox(height: 16),

                    // Date Pill Header
                    Center(
                      child: Container(
                        padding: const EdgeInsets.symmetric(horizontal: 14, vertical: 5),
                        decoration: BoxDecoration(
                          color: Colors.white.withOpacity(0.12),
                          borderRadius: BorderRadius.circular(20),
                          border: Border.all(color: Colors.white.withOpacity(0.2)),
                        ),
                        child: Row(
                          mainAxisSize: MainAxisSize.min,
                          children: [
                            const Icon(Icons.calendar_today_rounded, size: 13, color: Color(0xFF93C5FD)),
                            const SizedBox(width: 6),
                            Text(
                              todayFormatted,
                              style: const TextStyle(
                                color: Colors.white,
                                fontSize: 12,
                                fontWeight: FontWeight.w600,
                              ),
                            ),
                          ],
                        ),
                      ),
                    ),

                    const SizedBox(height: 18),

                    // Floating Time Card: "Mulai Izin" & "Selesai Izin"
                    Container(
                      padding: const EdgeInsets.symmetric(vertical: 18, horizontal: 18),
                      decoration: BoxDecoration(
                        color: Colors.white,
                        borderRadius: BorderRadius.circular(22),
                        boxShadow: [
                          BoxShadow(
                            color: const Color(0xFF0F172A).withOpacity(0.06),
                            blurRadius: 20,
                            offset: const Offset(0, 8),
                          ),
                        ],
                      ),
                      child: Row(
                        children: [
                          Expanded(
                            child: Column(
                              children: [
                                Row(
                                  mainAxisAlignment: MainAxisAlignment.center,
                                  children: const [
                                    Icon(Icons.play_circle_fill_rounded, size: 15, color: Color(0xFF2563EB)),
                                    SizedBox(width: 5),
                                    Text(
                                      "Mulai Izin",
                                      style: TextStyle(
                                        fontSize: 12,
                                        color: Color(0xFF64748B),
                                        fontWeight: FontWeight.w600,
                                      ),
                                    ),
                                  ],
                                ),
                                const SizedBox(height: 6),
                                Text(
                                  startPermitTime,
                                  style: TextStyle(
                                    fontSize: hasActivePermit ? 19 : 24,
                                    fontWeight: FontWeight.bold,
                                    color: hasActivePermit ? const Color(0xFFEA580C) : const Color(0xFF1E293B),
                                    letterSpacing: 0.5,
                                  ),
                                ),
                              ],
                            ),
                          ),
                          Container(
                            height: 40,
                            width: 1,
                            color: const Color(0xFFE2E8F0),
                          ),
                          Expanded(
                            child: Column(
                              children: [
                                Row(
                                  mainAxisAlignment: MainAxisAlignment.center,
                                  children: const [
                                    Icon(Icons.stop_circle_rounded, size: 15, color: Color(0xFF10B981)),
                                    SizedBox(width: 5),
                                    Text(
                                      "Selesai Izin",
                                      style: TextStyle(
                                        fontSize: 12,
                                        color: Color(0xFF64748B),
                                        fontWeight: FontWeight.w600,
                                      ),
                                    ),
                                  ],
                                ),
                                const SizedBox(height: 6),
                                const Text(
                                  endPermitTime,
                                  style: TextStyle(
                                    fontSize: 24,
                                    fontWeight: FontWeight.bold,
                                    color: Color(0xFF1E293B),
                                    letterSpacing: 0.5,
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
                      const SizedBox(height: 14),
                      InkWell(
                        onTap: _navigateToRecordPermit,
                        borderRadius: BorderRadius.circular(16),
                        child: Container(
                          padding: const EdgeInsets.all(14),
                          decoration: BoxDecoration(
                            gradient: const LinearGradient(
                              colors: [Color(0xFFFFFBEB), Color(0xFFFEF3C7)],
                              begin: Alignment.topLeft,
                              end: Alignment.bottomRight,
                            ),
                            borderRadius: BorderRadius.circular(16),
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
                                child: const Icon(Icons.directions_walk_rounded, color: Colors.white, size: 22),
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
                                    const SizedBox(height: 4),
                                    Text(
                                      "Keperluan: ${_activePermit!.purpose}",
                                      style: const TextStyle(
                                        fontSize: 11.5,
                                        color: Color(0xFF1E293B),
                                        fontWeight: FontWeight.w500,
                                      ),
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
                      const SizedBox(height: 14),
                      InkWell(
                        onTap: _navigateToCepuMapReport,
                        borderRadius: BorderRadius.circular(16),
                        child: Container(
                          padding: const EdgeInsets.all(14),
                          decoration: BoxDecoration(
                            color: const Color(0xFFFEF2F2),
                            borderRadius: BorderRadius.circular(16),
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

                    const SizedBox(height: 20),

                    // =========================================================================
                    // 3 ACTION CARDS (POIN 2: HILANGKAN SUB JUDUL CEPU & MONITORING HARIAN)
                    // =========================================================================

                    // 1. CARD 1: WAIGAMA (Sub judul: "Wajib Isi Guna Kebaikan Bersama")
                    _buildHorizontalMenuCard(
                      title: "Waigama",
                      subtitle: "Wajib Isi Guna Kebaikan Bersama",
                      icon: Icons.assignment_ind_rounded,
                      gradientColors: hasActivePermit
                          ? const [Color(0xFFF59E0B), Color(0xFFEA580C)]
                          : const [Color(0xFF2563EB), Color(0xFF3B82F6), Color(0xFF60A5FA)],
                      showActiveBadge: hasActivePermit,
                      onTap: _navigateToRecordPermit,
                    ),

                    const SizedBox(height: 14),

                    // 2. CARD 2: CEPU (Poin 2: Sub judul dihilangkan)
                    _buildHorizontalMenuCard(
                      title: "Cepu",
                      subtitle: null,
                      icon: Icons.campaign_rounded,
                      gradientColors: const [Color(0xFFEA580C), Color(0xFFF97316), Color(0xFFFB923C)],
                      showActiveBadge: false,
                      onTap: _navigateToCepuMapReport,
                    ),

                    const SizedBox(height: 14),

                    // 3. CARD 3: MONITORING HARIAN (Poin 2: Sub judul dihilangkan)
                    _buildHorizontalMenuCard(
                      title: "Monitoring Harian",
                      subtitle: null,
                      icon: Icons.analytics_rounded,
                      gradientColors: const [Color(0xFF0F766E), Color(0xFF0D9488), Color(0xFF14B8A6)],
                      showActiveBadge: false,
                      onTap: _navigateToDailyMonitoring,
                    ),

                    const SizedBox(height: 24),
                  ],
                ),
              ),
            ),
          ),
        ],
      ),
    );
  }

  Widget _buildHorizontalMenuCard({
    required String title,
    String? subtitle,
    required IconData icon,
    required List<Color> gradientColors,
    bool showActiveBadge = false,
    required VoidCallback onTap,
  }) {
    return Material(
      color: Colors.transparent,
      child: InkWell(
        onTap: onTap,
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
                color: gradientColors.first.withOpacity(0.3),
                blurRadius: 16,
                offset: const Offset(0, 6),
              ),
            ],
          ),
          padding: EdgeInsets.symmetric(
            horizontal: 20,
            vertical: subtitle != null ? 18 : 20,
          ),
          child: Row(
            children: [
              Container(
                width: 48,
                height: 48,
                decoration: BoxDecoration(
                  color: Colors.white.withOpacity(0.2),
                  shape: BoxShape.circle,
                  border: Border.all(color: Colors.white.withOpacity(0.35), width: 1.5),
                ),
                child: Icon(icon, color: Colors.white, size: 25),
              ),

              const SizedBox(width: 16),

              Expanded(
                child: Column(
                  crossAxisAlignment: CrossAxisAlignment.start,
                  mainAxisAlignment: MainAxisAlignment.center,
                  children: [
                    Row(
                      children: [
                        Text(
                          title,
                          style: const TextStyle(
                            color: Colors.white,
                            fontSize: 18,
                            fontWeight: FontWeight.w900,
                            letterSpacing: 0.2,
                          ),
                        ),
                        if (showActiveBadge) ...[
                          const SizedBox(width: 8),
                          Container(
                            padding: const EdgeInsets.symmetric(horizontal: 7, vertical: 2),
                            decoration: BoxDecoration(
                              color: Colors.black.withOpacity(0.25),
                              borderRadius: BorderRadius.circular(6),
                            ),
                            child: const Text(
                              "Sedang Izin",
                              style: TextStyle(
                                color: Colors.white,
                                fontSize: 9.5,
                                fontWeight: FontWeight.bold,
                              ),
                            ),
                          ),
                        ],
                      ],
                    ),
                    if (subtitle != null && subtitle.isNotEmpty) ...[
                      const SizedBox(height: 3),
                      Text(
                        subtitle,
                        style: TextStyle(
                          color: Colors.white.withOpacity(0.9),
                          fontSize: 11.5,
                          height: 1.25,
                        ),
                        maxLines: 2,
                        overflow: TextOverflow.ellipsis,
                      ),
                    ],
                  ],
                ),
              ),

              const SizedBox(width: 8),

              Container(
                padding: const EdgeInsets.all(7),
                decoration: BoxDecoration(
                  color: Colors.white.withOpacity(0.2),
                  shape: BoxShape.circle,
                ),
                child: const Icon(Icons.arrow_forward_ios_rounded, color: Colors.white, size: 14),
              ),
            ],
          ),
        ),
      ),
    );
  }
}
