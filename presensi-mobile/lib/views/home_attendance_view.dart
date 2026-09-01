import 'dart:async';
import 'package:flutter/material.dart';
import 'package:intl/intl.dart';
import '../models/user_model.dart';
import '../models/office_model.dart';
import '../models/permit_model.dart';
import '../core/services/attendance_service.dart';
import '../core/services/permit_service.dart';
import '../core/services/auth_service.dart';
import 'login_view.dart';
import 'record_permit_map_view.dart';
import 'cepu_report_view.dart';
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
  final _authService = AuthService();

  OfficeModel? _office;
  PermitModel? _activePermit;
  bool _isLoading = true;

  @override
  void initState() {
    super.initState();
    _loadInitialData();
  }

  Future<void> _loadInitialData() async {
    setState(() => _isLoading = true);
    final office = await _attendanceService.getOfficeConfig(widget.user.officeId);
    final activePermit = await _permitService.getActivePermit(widget.user.uid);

    if (mounted) {
      setState(() {
        _office = office;
        _activePermit = activePermit;
        _isLoading = false;
      });
    }
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

  void _navigateToCepuReport() async {
    final result = await Navigator.push(
      context,
      MaterialPageRoute(
        builder: (context) => CepuReportView(reporter: widget.user),
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
        ? DateFormat('HH:mm').format(_activePermit!.startTime)
        : "--:--";
    const endPermitTime = "--:--"; // Selalu reset ke --:-- ketika tidak ada izin atau selesai izin

    return Scaffold(
      backgroundColor: const Color(0xFFF8FAFC),
      body: Stack(
        children: [
          // Cheerful Blue-Indigo Gradient Header Background
          Container(
            height: 290,
            decoration: const BoxDecoration(
              gradient: LinearGradient(
                colors: [Color(0xFF1E60F2), Color(0xFF2563EB), Color(0xFF4F46E5)],
                begin: Alignment.topLeft,
                end: Alignment.bottomRight,
              ),
            ),
          ),

          SafeArea(
            child: RefreshIndicator(
              onRefresh: _loadInitialData,
              child: SingleChildScrollView(
                physics: const AlwaysScrollableScrollPhysics(),
                padding: const EdgeInsets.symmetric(horizontal: 18, vertical: 12),
                child: Column(
                  crossAxisAlignment: CrossAxisAlignment.stretch,
                  children: [
                    // Top App Bar: Profile, Name, Departemen & Logout
                    Row(
                      children: [
                        // Avatar with Active Green Dot
                        Stack(
                          children: [
                            CircleAvatar(
                              radius: 22,
                              backgroundColor: Colors.white24,
                              child: Text(
                                widget.user.name.isNotEmpty ? widget.user.name[0].toUpperCase() : 'U',
                                style: const TextStyle(
                                  color: Colors.white,
                                  fontWeight: FontWeight.bold,
                                  fontSize: 18,
                                ),
                              ),
                            ),
                            Positioned(
                              top: 0,
                              right: 0,
                              child: Container(
                                width: 11,
                                height: 11,
                                decoration: BoxDecoration(
                                  color: const Color(0xFF10B981),
                                  shape: BoxShape.circle,
                                  border: Border.all(color: Colors.white, width: 2),
                                ),
                              ),
                            ),
                          ],
                        ),
                        const SizedBox(width: 12),
                        // Name & Department
                        Expanded(
                          child: Column(
                            crossAxisAlignment: CrossAxisAlignment.start,
                            children: [
                              Text(
                                widget.user.name,
                                style: const TextStyle(
                                  color: Colors.white,
                                  fontSize: 16,
                                  fontWeight: FontWeight.bold,
                                ),
                                maxLines: 1,
                                overflow: TextOverflow.ellipsis,
                              ),
                              Text(
                                widget.user.department,
                                style: TextStyle(
                                  color: Colors.white.withOpacity(0.85),
                                  fontSize: 12,
                                  fontWeight: FontWeight.w500,
                                ),
                                maxLines: 1,
                                overflow: TextOverflow.ellipsis,
                              ),
                            ],
                          ),
                        ),
                        // Action Icons
                        IconButton(
                          icon: const Icon(Icons.refresh_rounded, color: Colors.white, size: 22),
                          onPressed: _loadInitialData,
                          tooltip: "Perbarui Data",
                        ),
                        IconButton(
                          icon: const Icon(Icons.logout_rounded, color: Colors.white, size: 22),
                          onPressed: _handleLogout,
                          tooltip: "Keluar",
                        ),
                      ],
                    ),

                    const SizedBox(height: 16),

                    // Date Pill Badge (WFOL telah dihapus)
                    Center(
                      child: Container(
                        padding: const EdgeInsets.symmetric(horizontal: 16, vertical: 6),
                        decoration: BoxDecoration(
                          color: Colors.white.withOpacity(0.2),
                          borderRadius: BorderRadius.circular(20),
                          border: Border.all(color: Colors.white.withOpacity(0.3)),
                        ),
                        child: Row(
                          mainAxisSize: MainAxisSize.min,
                          children: [
                            const Icon(Icons.calendar_today_rounded, size: 14, color: Colors.white),
                            const SizedBox(width: 8),
                            Text(
                              todayFormatted,
                              style: const TextStyle(
                                color: Colors.white,
                                fontSize: 13,
                                fontWeight: FontWeight.w600,
                              ),
                            ),
                          ],
                        ),
                      ),
                    ),

                    const SizedBox(height: 18),

                    // Card Status: "Mulai Izin" & "Selesai Izin"
                    Container(
                      padding: const EdgeInsets.symmetric(vertical: 20, horizontal: 20),
                      decoration: BoxDecoration(
                        color: Colors.white,
                        borderRadius: BorderRadius.circular(22),
                        boxShadow: [
                          BoxShadow(
                            color: Colors.black.withOpacity(0.08),
                            blurRadius: 16,
                            offset: const Offset(0, 6),
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
                                    Icon(Icons.play_circle_fill_rounded, size: 16, color: Color(0xFF3B82F6)),
                                    SizedBox(width: 6),
                                    Text(
                                      "Mulai Izin",
                                      style: TextStyle(
                                        fontSize: 13,
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
                                    fontSize: 28,
                                    fontWeight: FontWeight.bold,
                                    color: hasActivePermit ? const Color(0xFFEA580C) : const Color(0xFF1E40AF),
                                    letterSpacing: 0.5,
                                  ),
                                ),
                              ],
                            ),
                          ),
                          Container(
                            height: 48,
                            width: 1,
                            color: const Color(0xFFE2E8F0),
                          ),
                          Expanded(
                            child: Column(
                              children: [
                                Row(
                                  mainAxisAlignment: MainAxisAlignment.center,
                                  children: const [
                                    Icon(Icons.stop_circle_rounded, size: 16, color: Color(0xFF10B981)),
                                    SizedBox(width: 6),
                                    Text(
                                      "Selesai Izin",
                                      style: TextStyle(
                                        fontSize: 13,
                                        color: Color(0xFF64748B),
                                        fontWeight: FontWeight.w600,
                                      ),
                                    ),
                                  ],
                                ),
                                const SizedBox(height: 6),
                                Text(
                                  endPermitTime,
                                  style: const TextStyle(
                                    fontSize: 28,
                                    fontWeight: FontWeight.bold,
                                    color: Color(0xFF1E40AF),
                                    letterSpacing: 0.5,
                                  ),
                                ),
                              ],
                            ),
                          ),
                        ],
                      ),
                    ),

                    // Active Permit Alert Banner
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
                                    Row(
                                      children: [
                                        const Text(
                                          "Sedang Izin Keluar Kantor",
                                          style: TextStyle(
                                            fontSize: 12,
                                            fontWeight: FontWeight.bold,
                                            color: Color(0xFF92400E),
                                          ),
                                        ),
                                        const SizedBox(width: 6),
                                        Text(
                                          "(${DateFormat('HH:mm').format(_activePermit!.startTime)})",
                                          style: const TextStyle(
                                            fontSize: 11,
                                            color: Color(0xFFB45309),
                                            fontWeight: FontWeight.w600,
                                          ),
                                        ),
                                      ],
                                    ),
                                    const SizedBox(height: 2),
                                    Text(
                                      _activePermit!.purpose,
                                      style: const TextStyle(
                                        fontSize: 12,
                                        color: Color(0xFF1E293B),
                                        fontWeight: FontWeight.w500,
                                      ),
                                      maxLines: 1,
                                      overflow: TextOverflow.ellipsis,
                                    ),
                                    const SizedBox(height: 2),
                                    const Text(
                                      "👉 Klik di sini untuk selesaikan izin saat kembali ke kantor",
                                      style: TextStyle(
                                        fontSize: 11,
                                        color: Color(0xFF2563EB),
                                        fontWeight: FontWeight.bold,
                                      ),
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

                    const SizedBox(height: 24),

                    // Section Title
                    const Text(
                      "Menu Utama",
                      style: TextStyle(
                        fontSize: 16,
                        fontWeight: FontWeight.bold,
                        color: Color(0xFF1E293B),
                      ),
                    ),

                    const SizedBox(height: 12),

                    // =========================================================================
                    // 3 MODERN HORIZONTAL WIDE ACTION CARDS (COLORFUL, CHEERFUL & MODERN)
                    // =========================================================================

                    // 1. CARD 1: IZIN KEGIATAN KANTOR (Ada kegiatan di jam kantor? / Izin dulu!)
                    _buildHorizontalMenuCard(
                      tagText: "Ada kegiatan di jam kantor?",
                      title: hasActivePermit ? "Selesaikan Izin" : "Izin dulu!",
                      subtitle: "Rekam waktu mulai & selesai izin kegiatan dinas keluar kantor",
                      icon: Icons.assignment_ind_rounded,
                      gradientColors: hasActivePermit
                          ? const [Color(0xFFF59E0B), Color(0xFFEA580C)]
                          : const [Color(0xFFEF4444), Color(0xFFF43F5E), Color(0xFFFB7185)],
                      badgeText: hasActivePermit ? "Sedang Izin Aktif" : "GPS Geofencing",
                      badgeColor: hasActivePermit ? const Color(0xFFB45309) : const Color(0xFF9F1239),
                      onTap: _navigateToRecordPermit,
                    ),

                    const SizedBox(height: 14),

                    // 2. CARD 2: FITUR CEPU (Lapor Pegawai Tanpa Izin)
                    _buildHorizontalMenuCard(
                      tagText: "Disiplin & Transparansi",
                      title: "Fitur Cepu",
                      subtitle: "Laporkan rekan yang tidak ada di kantor tanpa izin",
                      icon: Icons.campaign_rounded,
                      gradientColors: const [Color(0xFFF97316), Color(0xFFEA580C), Color(0xFFDC2626)],
                      badgeText: "Verifikasi 4 Rekan",
                      badgeColor: const Color(0xFF7C2D12),
                      onTap: _navigateToCepuReport,
                    ),

                    const SizedBox(height: 14),

                    // 3. CARD 3: MONITORING HARIAN (Pantau Izin & Laporan Cepu)
                    _buildHorizontalMenuCard(
                      tagText: "Rekap & Verifikasi",
                      title: "Monitoring Harian",
                      subtitle: "Pantau seluruh izin pegawai & lakukan verifikasi laporan Cepu",
                      icon: Icons.analytics_rounded,
                      gradientColors: const [Color(0xFF0D9488), Color(0xFF059669), Color(0xFF10B981)],
                      badgeText: "Real-time Rekap",
                      badgeColor: const Color(0xFF064E3B),
                      onTap: _navigateToDailyMonitoring,
                    ),

                    const SizedBox(height: 24),

                    // Informasi Singkat Aturan Jam Kantor
                    Container(
                      padding: const EdgeInsets.all(16),
                      decoration: BoxDecoration(
                        color: Colors.white,
                        borderRadius: BorderRadius.circular(18),
                        border: Border.all(color: const Color(0xFFE2E8F0)),
                      ),
                      child: Column(
                        crossAxisAlignment: CrossAxisAlignment.start,
                        children: [
                          Row(
                            children: const [
                              Icon(Icons.info_outline_rounded, size: 18, color: Color(0xFF2563EB)),
                              SizedBox(width: 8),
                              Text(
                                "Panduan Penggunaan",
                                style: TextStyle(
                                  fontSize: 13,
                                  fontWeight: FontWeight.bold,
                                  color: Color(0xFF2563EB),
                                ),
                              ),
                            ],
                          ),
                          const SizedBox(height: 8),
                          Text(
                            "• Izin Kantor: Wajib rekam Mulai Izin di radius kantor dan Selesai Izin setelah kembali.\n• Fitur Cepu: Gunakan untuk melaporkan rekan kerja yang meninggalkan kantor tanpa izin. Laporan akan dinyatakan sah setelah diverifikasi oleh 4 rekan kerja lain.\n• Monitoring Harian: Pantau seluruh aktivitas dan berikan verifikasi objektif.",
                            style: TextStyle(
                              fontSize: 11.5,
                              color: Colors.grey[700],
                              height: 1.45,
                            ),
                          ),
                        ],
                      ),
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

  /// Widget Card Menu Horizontal yang Colorful, Cheerful, Clean, dan Modern
  Widget _buildHorizontalMenuCard({
    required String tagText,
    required String title,
    required String subtitle,
    required IconData icon,
    required List<Color> gradientColors,
    required String badgeText,
    required Color badgeColor,
    required VoidCallback onTap,
  }) {
    return Material(
      color: Colors.transparent,
      child: InkWell(
        onTap: onTap,
        borderRadius: BorderRadius.circular(20),
        child: Container(
          decoration: BoxDecoration(
            gradient: LinearGradient(
              colors: gradientColors,
              begin: Alignment.topLeft,
              end: Alignment.bottomRight,
            ),
            borderRadius: BorderRadius.circular(20),
            boxShadow: [
              BoxShadow(
                color: gradientColors.first.withOpacity(0.32),
                blurRadius: 14,
                offset: const Offset(0, 6),
              ),
            ],
          ),
          padding: const EdgeInsets.symmetric(horizontal: 18, vertical: 16),
          child: Row(
            children: [
              // Icon Circle Badge
              Container(
                width: 52,
                height: 52,
                decoration: BoxDecoration(
                  color: Colors.white.withOpacity(0.22),
                  shape: BoxShape.circle,
                  border: Border.all(color: Colors.white.withOpacity(0.35), width: 1.5),
                ),
                child: Icon(icon, color: Colors.white, size: 28),
              ),

              const SizedBox(width: 14),

              // Title, Tag & Subtitle
              Expanded(
                child: Column(
                  crossAxisAlignment: CrossAxisAlignment.start,
                  children: [
                    // Top Tag & Badge Row
                    Row(
                      children: [
                        Flexible(
                          child: Text(
                            tagText,
                            style: TextStyle(
                              color: Colors.white.withOpacity(0.92),
                              fontSize: 10.5,
                              fontWeight: FontWeight.w600,
                              letterSpacing: 0.2,
                            ),
                            maxLines: 1,
                            overflow: TextOverflow.ellipsis,
                          ),
                        ),
                        const SizedBox(width: 6),
                        Container(
                          padding: const EdgeInsets.symmetric(horizontal: 6, vertical: 2),
                          decoration: BoxDecoration(
                            color: Colors.black.withOpacity(0.18),
                            borderRadius: BorderRadius.circular(6),
                          ),
                          child: Text(
                            badgeText,
                            style: const TextStyle(
                              color: Colors.white,
                              fontSize: 9,
                              fontWeight: FontWeight.bold,
                            ),
                          ),
                        ),
                      ],
                    ),

                    const SizedBox(height: 4),

                    // Main Big Title
                    Text(
                      title,
                      style: const TextStyle(
                        color: Colors.white,
                        fontSize: 17,
                        fontWeight: FontWeight.w900,
                        letterSpacing: 0.3,
                      ),
                    ),

                    const SizedBox(height: 2),

                    // Subtitle
                    Text(
                      subtitle,
                      style: TextStyle(
                        color: Colors.white.withOpacity(0.88),
                        fontSize: 11,
                        height: 1.25,
                      ),
                      maxLines: 2,
                      overflow: TextOverflow.ellipsis,
                    ),
                  ],
                ),
              ),

              const SizedBox(width: 8),

              // Action Arrow Circle
              Container(
                padding: const EdgeInsets.all(8),
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
