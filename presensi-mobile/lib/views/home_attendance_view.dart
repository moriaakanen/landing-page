import 'dart:async';
import 'package:flutter/material.dart';
import 'package:intl/intl.dart';
import '../models/user_model.dart';
import '../models/office_model.dart';
import '../models/attendance_model.dart';
import '../models/permit_model.dart';
import '../core/services/attendance_service.dart';
import '../core/services/permit_service.dart';
import '../core/services/auth_service.dart';
import 'history_view.dart';
import 'login_view.dart';
import 'record_attendance_map_view.dart';
import 'record_permit_map_view.dart';
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
  AttendanceModel? _todayAttendance;
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
    final attendance = await _attendanceService.getTodayAttendance(widget.user.uid);
    final activePermit = await _permitService.getActivePermit(widget.user.uid);

    if (mounted) {
      setState(() {
        _office = office;
        _todayAttendance = attendance;
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

  void _navigateToRecordAttendance() async {
    if (_office == null) return;
    final result = await Navigator.push(
      context,
      MaterialPageRoute(
        builder: (context) => RecordAttendanceMapView(
          user: widget.user,
          office: _office!,
          todayAttendance: _todayAttendance,
        ),
      ),
    );

    if (result == true) {
      _loadInitialData();
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
      todayFormatted = DateFormat('EEE, d MMMM yyyy', 'id_ID').format(DateTime.now());
    } catch (_) {
      todayFormatted = DateFormat('EEE, d MMMM yyyy').format(DateTime.now());
    }

    final hasCheckIn = _todayAttendance?.checkIn != null;
    final hasCheckOut = _todayAttendance?.checkOut != null;
    final checkInTime = hasCheckIn
        ? DateFormat('HH:mm').format(_todayAttendance!.checkIn!.time)
        : "--:--";
    final checkOutTime = hasCheckOut
        ? DateFormat('HH:mm').format(_todayAttendance!.checkOut!.time)
        : "--:--";

    return Scaffold(
      backgroundColor: const Color(0xFFF3F4F6),
      body: Stack(
        children: [
          // Blue Gradient Background Header
          Container(
            height: 280,
            decoration: const BoxDecoration(
              gradient: LinearGradient(
                colors: [Color(0xFF1E60F2), Color(0xFF2563EB), Color(0xFF3B82F6)],
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
                    // Top App Bar: Profile, Name, Icons
                    Row(
                      children: [
                        // Avatar with Green Active Dot
                        Stack(
                          children: [
                            CircleAvatar(
                              radius: 20,
                              backgroundColor: Colors.white24,
                              child: Text(
                                widget.user.name.isNotEmpty ? widget.user.name[0].toUpperCase() : 'U',
                                style: const TextStyle(
                                  color: Colors.white,
                                  fontWeight: FontWeight.bold,
                                  fontSize: 16,
                                ),
                              ),
                            ),
                            Positioned(
                              top: 0,
                              right: 0,
                              child: Container(
                                width: 10,
                                height: 10,
                                decoration: BoxDecoration(
                                  color: const Color(0xFF10B981),
                                  shape: BoxShape.circle,
                                  border: Border.all(color: Colors.white, width: 1.5),
                                ),
                              ),
                            ),
                          ],
                        ),
                        const SizedBox(width: 10),
                        // Name
                        Expanded(
                          child: Text(
                            widget.user.name,
                            style: const TextStyle(
                              color: Colors.white,
                              fontSize: 16,
                              fontWeight: FontWeight.bold,
                            ),
                            maxLines: 1,
                            overflow: TextOverflow.ellipsis,
                          ),
                        ),
                        // Action Icons
                        IconButton(
                          icon: const Icon(Icons.refresh, color: Colors.white, size: 22),
                          onPressed: _loadInitialData,
                          tooltip: "Perbarui Data",
                        ),
                        IconButton(
                          icon: const Icon(Icons.share_outlined, color: Colors.white, size: 20),
                          onPressed: () {},
                          tooltip: "Bagikan",
                        ),
                        IconButton(
                          icon: const Icon(Icons.logout, color: Colors.white, size: 20),
                          onPressed: _handleLogout,
                          tooltip: "Keluar",
                        ),
                      ],
                    ),

                    const SizedBox(height: 18),

                    // Date & WFOL Pill Badge
                    Center(
                      child: Container(
                        padding: const EdgeInsets.symmetric(horizontal: 16, vertical: 6),
                        decoration: BoxDecoration(
                          color: const Color(0xFFF97316).withOpacity(0.85),
                          borderRadius: BorderRadius.circular(20),
                        ),
                        child: Row(
                          mainAxisSize: MainAxisSize.min,
                          children: [
                            Text(
                              todayFormatted,
                              style: const TextStyle(
                                color: Colors.white,
                                fontSize: 13,
                                fontWeight: FontWeight.w600,
                              ),
                            ),
                            const SizedBox(width: 8),
                            Container(
                              padding: const EdgeInsets.symmetric(horizontal: 6, vertical: 2),
                              decoration: BoxDecoration(
                                color: const Color(0xFFDC2626),
                                borderRadius: BorderRadius.circular(10),
                              ),
                              child: const Text(
                                "WFOL",
                                style: TextStyle(
                                  color: Colors.white,
                                  fontSize: 10,
                                  fontWeight: FontWeight.bold,
                                ),
                              ),
                            ),
                          ],
                        ),
                      ),
                    ),

                    const SizedBox(height: 18),

                    // White Card: Jam Datang & Jam Pulang
                    Container(
                      padding: const EdgeInsets.symmetric(vertical: 22, horizontal: 20),
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
                                Text(
                                  "Jam Datang",
                                  style: TextStyle(
                                    fontSize: 13,
                                    color: Colors.grey[600],
                                    fontWeight: FontWeight.w500,
                                  ),
                                ),
                                const SizedBox(height: 6),
                                Text(
                                  checkInTime,
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
                          Container(
                            height: 48,
                            width: 1,
                            color: const Color(0xFFE2E8F0),
                          ),
                          Expanded(
                            child: Column(
                              children: [
                                Text(
                                  "Jam Pulang",
                                  style: TextStyle(
                                    fontSize: 13,
                                    color: Colors.grey[600],
                                    fontWeight: FontWeight.w500,
                                  ),
                                ),
                                const SizedBox(height: 6),
                                Text(
                                  checkOutTime,
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

                    // Active Permit Banner Alert if user currently on permit
                    if (_activePermit != null) ...[
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
                                          "Status: Sedang Izin Keluar",
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
                                      "Klik di sini untuk selesaikan izin saat kembali",
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

                    const SizedBox(height: 22),

                    // 6 Grid Action Tiles
                    GridView.count(
                      crossAxisCount: 3,
                      crossAxisSpacing: 12,
                      mainAxisSpacing: 14,
                      childAspectRatio: 0.88,
                      shrinkWrap: true,
                      physics: const NeverScrollableScrollPhysics(),
                      children: [
                        // 1. Presensi Bulanan
                        _buildMenuCard(
                          icon: Icons.calendar_month_rounded,
                          iconColor: const Color(0xFFF59E0B),
                          title: "Presensi\nBulanan",
                          bgColor: const Color(0xFF60A5FA),
                          onTap: () {
                            Navigator.push(
                              context,
                              MaterialPageRoute(
                                builder: (context) => HistoryView(user: widget.user),
                              ),
                            );
                          },
                        ),

                        // 2. Izin Kegiatan Kantor (CUSTOM PERMIT BUTTON WITH SPECIFIED TEXTS)
                        _buildPermitMenuCard(
                          topText: "Ada kegiatan di jam kantor?",
                          icon: _activePermit != null
                              ? Icons.assignment_turned_in_rounded
                              : Icons.assignment_ind_rounded,
                          bottomText: _activePermit != null ? "Selesai Izin" : "Izin dulu!",
                          isHighlighted: true,
                          isActive: _activePermit != null,
                          onTap: _navigateToRecordPermit,
                        ),

                        // 3. Presensi Manual
                        _buildMenuCard(
                          icon: Icons.touch_app_rounded,
                          iconColor: const Color(0xFFF97316),
                          title: "Presensi\nManual",
                          bgColor: const Color(0xFF60A5FA),
                          onTap: () {
                            ScaffoldMessenger.of(context).showSnackBar(
                              const SnackBar(content: Text("Fitur Presensi Manual khusus jika ada kendala GPS")),
                            );
                          },
                        ),

                        // 4. Lapor Aktivitas
                        _buildMenuCard(
                          icon: Icons.alt_route_rounded,
                          iconColor: const Color(0xFFEF4444),
                          title: "Lapor\nAktivitas",
                          bgColor: const Color(0xFF60A5FA),
                          onTap: () {},
                        ),

                        // 5. Monitoring Harian (OPEN DAILY MONITORING VIEW)
                        _buildMenuCard(
                          icon: Icons.search_rounded,
                          iconColor: const Color(0xFFFCD34D),
                          title: "Monitoring\nHarian",
                          bgColor: const Color(0xFF60A5FA),
                          onTap: _navigateToDailyMonitoring,
                        ),

                        // 6. Menu Lainnya
                        _buildMenuCard(
                          icon: Icons.grid_view_rounded,
                          iconColor: const Color(0xFFFBBF24),
                          title: "Menu\nLainnya",
                          bgColor: const Color(0xFF60A5FA),
                          onTap: () {},
                        ),
                      ],
                    ),

                    const SizedBox(height: 24),

                    // Information / Changelog Card at bottom
                    Row(
                      mainAxisAlignment: MainAxisAlignment.spaceBetween,
                      children: [
                        const Text(
                          "Informasi:",
                          style: TextStyle(
                            fontSize: 14,
                            fontWeight: FontWeight.bold,
                            color: Color(0xFF1E293B),
                          ),
                        ),
                        TextButton.icon(
                          onPressed: _loadInitialData,
                          icon: const Icon(Icons.refresh, size: 16, color: Color(0xFF2563EB)),
                          label: const Text("Refresh", style: TextStyle(fontSize: 12, color: Color(0xFF2563EB))),
                          style: TextButton.styleFrom(padding: EdgeInsets.zero),
                        ),
                      ],
                    ),

                    const SizedBox(height: 6),

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
                          const Text(
                            "Fitur Izin Kegiatan Jam Kantor",
                            style: TextStyle(
                              fontSize: 13,
                              fontWeight: FontWeight.bold,
                              color: Color(0xFF2563EB),
                            ),
                          ),
                          const SizedBox(height: 6),
                          Text(
                            "1. Mulai Izin: Rekam waktu mulai & isi keperluan saat berada di kantor.\n2. Selesai Izin: Rekam waktu selesai setelah kembali ke radius kantor.\n3. Monitoring Harian: Pantau seluruh izin pegawai hari ini dan rekap riwayat.",
                            style: TextStyle(
                              fontSize: 12,
                              color: Colors.grey[700],
                              height: 1.4,
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

  /// Custom Menu Card Khusus Tombol Izin Sesuai Permintaan User:
  /// - "Ada kegiatan di jam kantor?" di atas logo menu
  /// - Logo menu di tengah
  /// - "Izin dulu!" di bawah logo menu
  Widget _buildPermitMenuCard({
    required String topText,
    required IconData icon,
    required String bottomText,
    bool isHighlighted = true,
    bool isActive = false,
    required VoidCallback onTap,
  }) {
    return Material(
      color: Colors.transparent,
      child: InkWell(
        onTap: onTap,
        borderRadius: BorderRadius.circular(18),
        child: Container(
          decoration: BoxDecoration(
            gradient: isActive
                ? const LinearGradient(
                    colors: [Color(0xFFF59E0B), Color(0xFFD97706)],
                    begin: Alignment.topLeft,
                    end: Alignment.bottomRight,
                  )
                : const LinearGradient(
                    colors: [Color(0xFFEF4444), Color(0xFFF43F5E), Color(0xFFFB7185)],
                    begin: Alignment.topLeft,
                    end: Alignment.bottomRight,
                  ),
            borderRadius: BorderRadius.circular(18),
            boxShadow: [
              BoxShadow(
                color: isActive
                    ? const Color(0xFFF59E0B).withOpacity(0.35)
                    : const Color(0xFFEF4444).withOpacity(0.3),
                blurRadius: 10,
                offset: const Offset(0, 4),
              ),
            ],
          ),
          padding: const EdgeInsets.symmetric(vertical: 8, horizontal: 6),
          child: Column(
            mainAxisAlignment: MainAxisAlignment.spaceBetween,
            children: [
              // Top Text
              Text(
                topText,
                textAlign: TextAlign.center,
                style: const TextStyle(
                  color: Colors.white,
                  fontSize: 9.5,
                  fontWeight: FontWeight.bold,
                  height: 1.1,
                ),
                maxLines: 2,
                overflow: TextOverflow.ellipsis,
              ),

              // Middle Icon
              Container(
                width: 38,
                height: 38,
                decoration: BoxDecoration(
                  color: Colors.white.withOpacity(0.25),
                  shape: BoxShape.circle,
                ),
                child: Icon(icon, color: Colors.white, size: 22),
              ),

              // Bottom Text
              Container(
                padding: const EdgeInsets.symmetric(horizontal: 8, vertical: 2),
                decoration: BoxDecoration(
                  color: Colors.black.withOpacity(0.18),
                  borderRadius: BorderRadius.circular(8),
                ),
                child: Text(
                  bottomText,
                  textAlign: TextAlign.center,
                  style: const TextStyle(
                    color: Colors.white,
                    fontSize: 11,
                    fontWeight: FontWeight.w900,
                  ),
                ),
              ),
            ],
          ),
        ),
      ),
    );
  }

  Widget _buildMenuCard({
    required IconData icon,
    required Color iconColor,
    required String title,
    Color bgColor = const Color(0xFF3B82F6),
    bool isHighlighted = false,
    required VoidCallback onTap,
  }) {
    return Material(
      color: Colors.transparent,
      child: InkWell(
        onTap: onTap,
        borderRadius: BorderRadius.circular(18),
        child: Container(
          decoration: BoxDecoration(
            gradient: isHighlighted
                ? const LinearGradient(
                    colors: [Color(0xFFEF4444), Color(0xFFF43F5E), Color(0xFFFB7185)],
                    begin: Alignment.topLeft,
                    end: Alignment.bottomRight,
                  )
                : const LinearGradient(
                    colors: [Color(0xFF60A5FA), Color(0xFF3B82F6)],
                    begin: Alignment.topLeft,
                    end: Alignment.bottomRight,
                  ),
            borderRadius: BorderRadius.circular(18),
            boxShadow: [
              BoxShadow(
                color: isHighlighted
                    ? const Color(0xFFEF4444).withOpacity(0.3)
                    : const Color(0xFF3B82F6).withOpacity(0.25),
                blurRadius: 10,
                offset: const Offset(0, 4),
              ),
            ],
          ),
          padding: const EdgeInsets.symmetric(vertical: 12, horizontal: 8),
          child: Column(
            mainAxisAlignment: MainAxisAlignment.center,
            children: [
              Container(
                width: 40,
                height: 40,
                decoration: BoxDecoration(
                  color: isHighlighted
                      ? Colors.white.withOpacity(0.25)
                      : Colors.white.withOpacity(0.2),
                  shape: BoxShape.circle,
                ),
                child: Icon(icon, color: iconColor, size: 22),
              ),
              const SizedBox(height: 8),
              Text(
                title,
                textAlign: TextAlign.center,
                style: const TextStyle(
                  color: Colors.white,
                  fontSize: 11,
                  fontWeight: FontWeight.bold,
                  height: 1.2,
                ),
              ),
            ],
          ),
        ),
      ),
    );
  }
}
