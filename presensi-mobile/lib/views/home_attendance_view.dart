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
  Timer? _clockTimer;
  DateTime _currentTime = DateTime.now();

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

  String _getGreeting() {
    final hour = DateTime.now().hour;
    if (hour < 11) return "Selamat Pagi,";
    if (hour < 15) return "Selamat Siang,";
    if (hour < 18) return "Selamat Sore,";
    return "Selamat Malam,";
  }

  @override
  Widget build(BuildContext context) {
    String todayFormatted;
    String dayNumber;
    String dayName;
    try {
      todayFormatted = DateFormat('EEEE, d MMMM yyyy', 'id_ID').format(DateTime.now());
      dayNumber = DateFormat('dd').format(DateTime.now());
      dayName = DateFormat('E', 'id_ID').format(DateTime.now());
    } catch (_) {
      todayFormatted = DateFormat('EEEE, d MMMM yyyy').format(DateTime.now());
      dayNumber = DateFormat('dd').format(DateTime.now());
      dayName = DateFormat('E').format(DateTime.now());
    }

    final hasActivePermit = _activePermit != null;
    final startPermitTime = hasActivePermit
        ? "${DateFormat('HH:mm').format(_activePermit!.startTime)} WIT"
        : "--:--";
    const endPermitTime = "--:--";
    final timeFormatted = DateFormat('HH:mm').format(_currentTime);

    return Scaffold(
      backgroundColor: const Color(0xFFF8FAFC),
      body: SafeArea(
        child: Stack(
          children: [
            RefreshIndicator(
              onRefresh: _loadInitialData,
              color: const Color(0xFF1E60F2),
              child: SingleChildScrollView(
                physics: const AlwaysScrollableScrollPhysics(),
                padding: const EdgeInsets.fromLTRB(20, 16, 20, 95),
                child: Column(
                  crossAxisAlignment: CrossAxisAlignment.start,
                  children: [
                    // =========================================================================
                    // 1. HEADER (Sesuai Home.tsx: Sapaan, Nama Pegawai, & Lonceng Notifikasi)
                    // =========================================================================
                    Row(
                      mainAxisAlignment: MainAxisAlignment.spaceBetween,
                      crossAxisAlignment: CrossAxisAlignment.start,
                      children: [
                        Row(
                          children: [
                            // Manta Ray App Logo in Clean Rounded Container
                            Container(
                              width: 44,
                              height: 44,
                              decoration: BoxDecoration(
                                color: Colors.white,
                                borderRadius: BorderRadius.circular(14),
                                border: Border.all(color: const Color(0xFFBFDBFE), width: 1.5),
                                boxShadow: [
                                  BoxShadow(
                                    color: const Color(0xFF1E60F2).withOpacity(0.08),
                                    blurRadius: 10,
                                    offset: const Offset(0, 4),
                                  ),
                                ],
                              ),
                              padding: const EdgeInsets.all(6),
                              child: Image.asset('assets/images/app_logo.png', fit: BoxFit.contain),
                            ),
                            const SizedBox(width: 12),
                            Column(
                              crossAxisAlignment: CrossAxisAlignment.start,
                              children: [
                                Text(
                                  _getGreeting(),
                                  style: const TextStyle(
                                    fontSize: 12.5,
                                    color: Color(0xFF64748B),
                                    fontWeight: FontWeight.w500,
                                  ),
                                ),
                                const SizedBox(height: 2),
                                Text(
                                  widget.user.name,
                                  style: const TextStyle(
                                    fontSize: 18,
                                    fontWeight: FontWeight.bold,
                                    color: Color(0xFF0F172A),
                                    letterSpacing: -0.3,
                                  ),
                                ),
                              ],
                            ),
                          ],
                        ),

                        // Action Buttons: Notification Bell & Refresh
                        Row(
                          children: [
                            _buildIconButton(
                              icon: Icons.refresh_rounded,
                              onTap: _loadInitialData,
                            ),
                            const SizedBox(width: 8),
                            _buildIconButton(
                              icon: Icons.notifications_none_rounded,
                              onTap: _navigateToDailyMonitoring,
                              hasBadge: hasActivePermit || _activeCepuForMe != null,
                            ),
                          ],
                        ),
                      ],
                    ),

                    const SizedBox(height: 14),

                    // =========================================================================
                    // 2. LOCATION BADGE & DATE (Sesuai Home.tsx: Pill Badge Biru Ocean)
                    // =========================================================================
                    InkWell(
                      onTap: _showProfileDialog,
                      borderRadius: BorderRadius.circular(20),
                      child: Container(
                        padding: const EdgeInsets.symmetric(horizontal: 14, vertical: 7),
                        decoration: BoxDecoration(
                          color: const Color(0xFF1E60F2),
                          borderRadius: BorderRadius.circular(20),
                          boxShadow: [
                            BoxShadow(
                              color: const Color(0xFF1E60F2).withOpacity(0.25),
                              blurRadius: 10,
                              offset: const Offset(0, 3),
                            ),
                          ],
                        ),
                        child: Row(
                          mainAxisSize: MainAxisSize.min,
                          children: [
                            const Icon(Icons.location_on_rounded, color: Colors.white, size: 15),
                            const SizedBox(width: 6),
                            Text(
                              _office != null
                                  ? "${_office!.name} (Radius ${_office!.radiusMeters.toInt()}m)"
                                  : "Kantor Utama (Radius 50m)",
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

                    const SizedBox(height: 4),

                    Text(
                      todayFormatted,
                      style: const TextStyle(
                        fontSize: 11.5,
                        color: Color(0xFF94A3B8),
                        fontWeight: FontWeight.w500,
                      ),
                    ),

                    const SizedBox(height: 16),

                    // =========================================================================
                    // 3. HERO CHECK IN/OUT CARD (Sesuai Home.tsx: Kartu Gradasi Besar & 2 Tombol)
                    // =========================================================================
                    Container(
                      decoration: BoxDecoration(
                        gradient: const LinearGradient(
                          colors: [Color(0xFF1E60F2), Color(0xFF0E3A99)],
                          begin: Alignment.topLeft,
                          end: Alignment.bottomRight,
                        ),
                        borderRadius: BorderRadius.circular(28),
                        boxShadow: [
                          BoxShadow(
                            color: const Color(0xFF1E60F2).withOpacity(0.3),
                            blurRadius: 20,
                            offset: const Offset(0, 8),
                          ),
                        ],
                      ),
                      padding: const EdgeInsets.all(22),
                      child: Column(
                        children: [
                          // Top: Digital Time Display
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
                            ],
                          ),

                          const SizedBox(height: 20),

                          // 2 Prominent Action Buttons: Check In (Mulai) & Check Out (Selesai)
                          Row(
                            children: [
                              // Button 1: Mulai Izin (Solid White Button Sesuai Home.tsx)
                              Expanded(
                                child: InkWell(
                                  onTap: hasActivePermit ? null : _navigateToRecordPermit,
                                  borderRadius: BorderRadius.circular(18),
                                  child: Container(
                                    padding: const EdgeInsets.symmetric(vertical: 14),
                                    decoration: BoxDecoration(
                                      color: hasActivePermit
                                          ? Colors.white.withOpacity(0.5)
                                          : Colors.white,
                                      borderRadius: BorderRadius.circular(18),
                                      boxShadow: [
                                        BoxShadow(
                                          color: Colors.black.withOpacity(0.08),
                                          blurRadius: 8,
                                          offset: const Offset(0, 2),
                                        ),
                                      ],
                                    ),
                                    child: Column(
                                      children: [
                                        const Icon(
                                          Icons.check_circle_rounded,
                                          color: Color(0xFF1E60F2),
                                          size: 24,
                                        ),
                                        const SizedBox(height: 6),
                                        const Text(
                                          "Mulai Izin",
                                          style: TextStyle(
                                            color: Color(0xFF1E60F2),
                                            fontWeight: FontWeight.bold,
                                            fontSize: 13.5,
                                          ),
                                        ),
                                        const SizedBox(height: 2),
                                        Text(
                                          hasActivePermit ? startPermitTime : "Rekam Mulai",
                                          style: TextStyle(
                                            fontSize: 11,
                                            color: const Color(0xFF1E60F2).withOpacity(0.75),
                                            fontWeight: FontWeight.w600,
                                          ),
                                        ),
                                      ],
                                    ),
                                  ),
                                ),
                              ),

                              const SizedBox(width: 14),

                              // Button 2: Selesai Izin (Frosted Glass Button Sesuai Home.tsx)
                              Expanded(
                                child: InkWell(
                                  onTap: hasActivePermit ? _navigateToRecordPermit : null,
                                  borderRadius: BorderRadius.circular(18),
                                  child: Container(
                                    padding: const EdgeInsets.symmetric(vertical: 14),
                                    decoration: BoxDecoration(
                                      color: Colors.white.withOpacity(hasActivePermit ? 0.28 : 0.15),
                                      borderRadius: BorderRadius.circular(18),
                                      border: Border.all(
                                        color: Colors.white.withOpacity(hasActivePermit ? 0.4 : 0.2),
                                        width: 1.2,
                                      ),
                                    ),
                                    child: Column(
                                      children: [
                                        const Icon(
                                          Icons.access_time_filled_rounded,
                                          color: Colors.white,
                                          size: 24,
                                        ),
                                        const SizedBox(height: 6),
                                        const Text(
                                          "Selesai Izin",
                                          style: TextStyle(
                                            color: Colors.white,
                                            fontWeight: FontWeight.bold,
                                            fontSize: 13.5,
                                          ),
                                        ),
                                        const SizedBox(height: 2),
                                        Text(
                                          hasActivePermit ? "Rekam Kembali" : "--:--",
                                          style: TextStyle(
                                            fontSize: 11,
                                            color: Colors.white.withOpacity(0.75),
                                            fontWeight: FontWeight.w600,
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

                    const SizedBox(height: 16),

                    // =========================================================================
                    // 4. CHECK IN/OUT STATUS CARDS (Sesuai Home.tsx: 2 Grid Kotak Status Putih)
                    // =========================================================================
                    Row(
                      children: [
                        // Status Card 1: Check In
                        Expanded(
                          child: Container(
                            padding: const EdgeInsets.all(16),
                            decoration: BoxDecoration(
                              color: Colors.white,
                              borderRadius: BorderRadius.circular(20),
                              border: Border.all(color: const Color(0xFFE2E8F0)),
                              boxShadow: [
                                BoxShadow(
                                  color: Colors.black.withOpacity(0.02),
                                  blurRadius: 8,
                                  offset: const Offset(0, 2),
                                ),
                              ],
                            ),
                            child: Column(
                              crossAxisAlignment: CrossAxisAlignment.start,
                              children: [
                                Row(
                                  children: const [
                                    Icon(Icons.check_circle_rounded, color: Color(0xFF1E60F2), size: 16),
                                    SizedBox(width: 6),
                                    Text(
                                      "Mulai Izin",
                                      style: TextStyle(fontSize: 12, color: Color(0xFF64748B), fontWeight: FontWeight.w600),
                                    ),
                                  ],
                                ),
                                const SizedBox(height: 8),
                                Text(
                                  startPermitTime,
                                  style: const TextStyle(
                                    fontSize: 20,
                                    fontWeight: FontWeight.bold,
                                    color: Color(0xFF0F172A),
                                  ),
                                ),
                                const SizedBox(height: 4),
                                Text(
                                  hasActivePermit ? "Sedang Berlangsung" : "Belum Mulai",
                                  style: TextStyle(
                                    fontSize: 11,
                                    color: hasActivePermit ? const Color(0xFFEA580C) : const Color(0xFF94A3B8),
                                    fontWeight: FontWeight.w500,
                                  ),
                                ),
                              ],
                            ),
                          ),
                        ),

                        const SizedBox(width: 12),

                        // Status Card 2: Check Out
                        Expanded(
                          child: Container(
                            padding: const EdgeInsets.all(16),
                            decoration: BoxDecoration(
                              color: Colors.white,
                              borderRadius: BorderRadius.circular(20),
                              border: Border.all(color: const Color(0xFFE2E8F0)),
                              boxShadow: [
                                BoxShadow(
                                  color: Colors.black.withOpacity(0.02),
                                  blurRadius: 8,
                                  offset: const Offset(0, 2),
                                ),
                              ],
                            ),
                            child: Column(
                              crossAxisAlignment: CrossAxisAlignment.start,
                              children: [
                                Row(
                                  children: const [
                                    Icon(Icons.access_time_rounded, color: Color(0xFF64748B), size: 16),
                                    SizedBox(width: 6),
                                    Text(
                                      "Selesai Izin",
                                      style: TextStyle(fontSize: 12, color: Color(0xFF64748B), fontWeight: FontWeight.w600),
                                    ),
                                  ],
                                ),
                                const SizedBox(height: 8),
                                const Text(
                                  endPermitTime,
                                  style: TextStyle(
                                    fontSize: 20,
                                    fontWeight: FontWeight.bold,
                                    color: Color(0xFF94A3B8),
                                  ),
                                ),
                                const SizedBox(height: 4),
                                Text(
                                  hasActivePermit ? "Belum Kembali" : "Standby",
                                  style: const TextStyle(
                                    fontSize: 11,
                                    color: Color(0xFF94A3B8),
                                    fontWeight: FontWeight.w500,
                                  ),
                                ),
                              ],
                            ),
                          ),
                        ),
                      ],
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

                    const SizedBox(height: 24),

                    // =========================================================================
                    // 5. FITUR UTAMA (4 Squircle Cards: Waigama, Cepu, Monitoring, Profil)
                    // =========================================================================
                    Row(
                      mainAxisAlignment: MainAxisAlignment.spaceBetween,
                      children: const [
                        Text(
                          "Fitur Utama",
                          style: TextStyle(
                            fontSize: 16.5,
                            fontWeight: FontWeight.bold,
                            color: Color(0xFF0F172A),
                          ),
                        ),
                        Text(
                          "Pilih Aksi",
                          style: TextStyle(fontSize: 12, color: Color(0xFF1E60F2), fontWeight: FontWeight.bold),
                        ),
                      ],
                    ),

                    const SizedBox(height: 14),

                    Row(
                      mainAxisAlignment: MainAxisAlignment.spaceBetween,
                      children: [
                        // 1. Waigama (Ocean Azure Blue)
                        _buildCategorySquircleCard(
                          title: "Waigama",
                          icon: Icons.assignment_ind_rounded,
                          bgColor: const Color(0xFF1E60F2),
                          onTap: _navigateToRecordPermit,
                          badgeText: hasActivePermit ? "Aktif" : null,
                        ),

                        // 2. Cepu (Sunset Orange)
                        _buildCategorySquircleCard(
                          title: "Cepu",
                          icon: Icons.campaign_rounded,
                          bgColor: const Color(0xFFEA580C),
                          onTap: _navigateToCepuMapReport,
                        ),

                        // 3. Monitoring (Teal Emerald)
                        _buildCategorySquircleCard(
                          title: "Monitoring",
                          icon: Icons.bar_chart_rounded,
                          bgColor: const Color(0xFF0D9488),
                          onTap: _navigateToDailyMonitoring,
                        ),

                        // 4. Profil (Royal Indigo)
                        _buildCategorySquircleCard(
                          title: "Profil",
                          icon: Icons.person_rounded,
                          bgColor: const Color(0xFF4F46E5),
                          onTap: _showProfileDialog,
                        ),
                      ],
                    ),

                    const SizedBox(height: 24),

                    // =========================================================================
                    // 6. ATTENDANCE ACTIVITY CARD (Sesuai Home.tsx: Kartu Aktivitas Hari Ini)
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
                          onTap: _navigateToDailyMonitoring,
                          child: const Text(
                            "Lihat Rekap",
                            style: TextStyle(fontSize: 12, color: Color(0xFF1E60F2), fontWeight: FontWeight.bold),
                          ),
                        ),
                      ],
                    ),

                    const SizedBox(height: 12),

                    // Activity Card Sesuai Home.tsx
                    Container(
                      decoration: BoxDecoration(
                        gradient: const LinearGradient(
                          colors: [Color(0xFF1E60F2), Color(0xFF1544B3)],
                          begin: Alignment.topLeft,
                          end: Alignment.bottomRight,
                        ),
                        borderRadius: BorderRadius.circular(22),
                        boxShadow: [
                          BoxShadow(
                            color: const Color(0xFF1E60F2).withOpacity(0.2),
                            blurRadius: 14,
                            offset: const Offset(0, 5),
                          ),
                        ],
                      ),
                      padding: const EdgeInsets.all(16),
                      child: Row(
                        children: [
                          // Date Square Box
                          Container(
                            padding: const EdgeInsets.symmetric(horizontal: 14, vertical: 10),
                            decoration: BoxDecoration(
                              color: Colors.white.withOpacity(0.2),
                              borderRadius: BorderRadius.circular(14),
                            ),
                            child: Column(
                              mainAxisSize: MainAxisSize.min,
                              children: [
                                Text(
                                  dayNumber,
                                  style: const TextStyle(
                                    fontSize: 22,
                                    fontWeight: FontWeight.w900,
                                    color: Colors.white,
                                  ),
                                ),
                                const SizedBox(height: 2),
                                Text(
                                  dayName,
                                  style: TextStyle(
                                    fontSize: 12,
                                    fontWeight: FontWeight.w600,
                                    color: Colors.white.withOpacity(0.85),
                                  ),
                                ),
                              ],
                            ),
                          ),

                          const SizedBox(width: 16),

                          // Attendance Time Details
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
                                          style: TextStyle(fontSize: 11, color: Colors.white.withOpacity(0.7)),
                                        ),
                                        const SizedBox(height: 2),
                                        Text(
                                          startPermitTime,
                                          style: const TextStyle(fontSize: 13.5, fontWeight: FontWeight.bold, color: Colors.white),
                                        ),
                                      ],
                                    ),
                                    Column(
                                      crossAxisAlignment: CrossAxisAlignment.start,
                                      children: [
                                        Text(
                                          "Selesai",
                                          style: TextStyle(fontSize: 11, color: Colors.white.withOpacity(0.7)),
                                        ),
                                        const SizedBox(height: 2),
                                        Text(
                                          endPermitTime,
                                          style: const TextStyle(fontSize: 13.5, fontWeight: FontWeight.bold, color: Colors.white),
                                        ),
                                      ],
                                    ),
                                    Column(
                                      crossAxisAlignment: CrossAxisAlignment.start,
                                      children: [
                                        Text(
                                          "Status",
                                          style: TextStyle(fontSize: 11, color: Colors.white.withOpacity(0.7)),
                                        ),
                                        const SizedBox(height: 2),
                                        Text(
                                          hasActivePermit ? "Izin" : "Hadir",
                                          style: const TextStyle(fontSize: 13.5, fontWeight: FontWeight.bold, color: Colors.white),
                                        ),
                                      ],
                                    ),
                                  ],
                                ),
                                const SizedBox(height: 10),
                                Row(
                                  children: [
                                    const Icon(Icons.location_on_rounded, color: Colors.white70, size: 13),
                                    const SizedBox(width: 4),
                                    Expanded(
                                      child: Text(
                                        _office != null
                                            ? "${_office!.name} (Area Kantor)"
                                            : "Kantor Utama",
                                        style: TextStyle(
                                          fontSize: 11.5,
                                          color: Colors.white.withOpacity(0.85),
                                        ),
                                        maxLines: 1,
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
                  ],
                ),
              ),
            ),

            // =========================================================================
            // 7. BOTTOM NAVIGATION BAR (Sesuai Navigation.tsx: Floating Bar dengan Center FAB)
            // =========================================================================
            Positioned(
              bottom: 0,
              left: 0,
              right: 0,
              child: Container(
                height: 74,
                decoration: BoxDecoration(
                  color: Colors.white,
                  borderRadius: const BorderRadius.vertical(top: Radius.circular(24)),
                  boxShadow: [
                    BoxShadow(
                      color: Colors.black.withOpacity(0.08),
                      blurRadius: 18,
                      offset: const Offset(0, -4),
                    ),
                  ],
                ),
                child: Row(
                  mainAxisAlignment: MainAxisAlignment.spaceAround,
                  children: [
                    // Item 1: Home
                    IconButton(
                      icon: const Icon(Icons.home_rounded, color: Color(0xFF1E60F2), size: 26),
                      onPressed: () {},
                      tooltip: "Home",
                    ),

                    // Item 2: Monitoring
                    IconButton(
                      icon: const Icon(Icons.bar_chart_rounded, color: Color(0xFF94A3B8), size: 26),
                      onPressed: _navigateToDailyMonitoring,
                      tooltip: "Monitoring",
                    ),

                    // Center Elevated FAB (Ocean Azure Circle Button)
                    Transform.translate(
                      offset: const Offset(0, -14),
                      child: InkWell(
                        onTap: _navigateToRecordPermit,
                        borderRadius: BorderRadius.circular(30),
                        child: Container(
                          width: 54,
                          height: 54,
                          decoration: BoxDecoration(
                            gradient: const LinearGradient(
                              colors: [Color(0xFF1E60F2), Color(0xFF0E3A99)],
                              begin: Alignment.topLeft,
                              end: Alignment.bottomRight,
                            ),
                            shape: BoxShape.circle,
                            boxShadow: [
                              BoxShadow(
                                color: const Color(0xFF1E60F2).withOpacity(0.4),
                                blurRadius: 12,
                                offset: const Offset(0, 6),
                              ),
                            ],
                          ),
                          child: const Icon(Icons.assignment_ind_rounded, color: Colors.white, size: 26),
                        ),
                      ),
                    ),

                    // Item 4: Cepu
                    IconButton(
                      icon: const Icon(Icons.campaign_outlined, color: Color(0xFF94A3B8), size: 26),
                      onPressed: _navigateToCepuMapReport,
                      tooltip: "Cepu",
                    ),

                    // Item 5: Profile
                    IconButton(
                      icon: const Icon(Icons.person_outline_rounded, color: Color(0xFF94A3B8), size: 26),
                      onPressed: _showProfileDialog,
                      tooltip: "Profil",
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
                  width: 68,
                  height: 68,
                  decoration: BoxDecoration(
                    color: bgColor,
                    borderRadius: BorderRadius.circular(22),
                    boxShadow: [
                      BoxShadow(
                        color: bgColor.withOpacity(0.3),
                        blurRadius: 10,
                        offset: const Offset(0, 4),
                      ),
                    ],
                  ),
                  child: Icon(icon, color: Colors.white, size: 28),
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
