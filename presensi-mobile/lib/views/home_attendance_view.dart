import 'dart:async';
import 'package:flutter/material.dart';
import 'package:intl/intl.dart';
import '../models/user_model.dart';
import '../models/office_model.dart';
import '../models/attendance_model.dart';
import '../core/services/attendance_service.dart';
import '../core/services/location_service.dart';
import '../core/services/auth_service.dart';
import '../widgets/radar_location_widget.dart';
import 'history_view.dart';
import 'login_view.dart';

class HomeAttendanceView extends StatefulWidget {
  final UserModel user;

  const HomeAttendanceView({Key? key, required this.user}) : super(key: key);

  @override
  State<HomeAttendanceView> createState() => _HomeAttendanceViewState();
}

class _HomeAttendanceViewState extends State<HomeAttendanceView> {
  final _attendanceService = AttendanceService();
  final _authService = AuthService();

  OfficeModel? _office;
  double _currentDistance = 0.0;
  bool _isCheckingLocation = false;
  bool _isMocked = false;
  bool _isLoadingAction = false;
  String _currentTimeString = "";
  Timer? _clockTimer;
  Timer? _locationPeriodicTimer;

  @override
  void initState() {
    super.initState();
    _updateTime();
    _clockTimer = Timer.periodic(const Duration(seconds: 1), (timer) => _updateTime());
    _loadInitialData();
    // Auto refresh location every 15 seconds
    _locationPeriodicTimer = Timer.periodic(const Duration(seconds: 15), (timer) {
      if (mounted && _office != null) {
        _checkDistance();
      }
    });
  }

  @override
  void dispose() {
    _clockTimer?.cancel();
    _locationPeriodicTimer?.cancel();
    super.dispose();
  }

  void _updateTime() {
    if (mounted) {
      setState(() {
        _currentTimeString = DateFormat('HH:mm:ss').format(DateTime.now());
      });
    }
  }

  Future<void> _loadInitialData() async {
    final office = await _attendanceService.getOfficeConfig(widget.user.officeId);
    if (mounted) {
      setState(() {
        _office = office;
      });
      _checkDistance();
    }
  }

  Future<void> _checkDistance() async {
    if (_office == null) return;
    setState(() {
      _isCheckingLocation = true;
    });

    final res = await LocationService.verifyPresenceLocation(
      officeLat: _office!.latitude,
      officeLng: _office!.longitude,
      allowedRadiusMeters: _office!.radiusMeters,
    );

    if (mounted) {
      setState(() {
        _isCheckingLocation = false;
        _currentDistance = res.distanceMeters;
        _isMocked = res.isMockLocation;
      });
    }
  }

  Future<void> _handleCheckIn() async {
    if (_office == null) return;
    setState(() => _isLoadingAction = true);

    try {
      await _attendanceService.checkIn(user: widget.user, office: _office!);
      if (mounted) {
        ScaffoldMessenger.of(context).showSnackBar(
          const SnackBar(
            content: Text("✅ Berhasil melakukan Absen Masuk!"),
            backgroundColor: Color(0xFF10B981),
          ),
        );
      }
    } catch (e) {
      if (mounted) {
        ScaffoldMessenger.of(context).showSnackBar(
          SnackBar(
            content: Text(e.toString()),
            backgroundColor: const Color(0xFFEF4444),
          ),
        );
      }
    } finally {
      if (mounted) setState(() => _isLoadingAction = false);
    }
  }

  Future<void> _handleCheckOut() async {
    if (_office == null) return;
    setState(() => _isLoadingAction = true);

    try {
      await _attendanceService.checkOut(user: widget.user, office: _office!);
      if (mounted) {
        ScaffoldMessenger.of(context).showSnackBar(
          const SnackBar(
            content: Text("✅ Berhasil melakukan Absen Pulang!"),
            backgroundColor: Color(0xFF10B981),
          ),
        );
      }
    } catch (e) {
      if (mounted) {
        ScaffoldMessenger.of(context).showSnackBar(
          SnackBar(
            content: Text(e.toString()),
            backgroundColor: const Color(0xFFEF4444),
          ),
        );
      }
    } finally {
      if (mounted) setState(() => _isLoadingAction = false);
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

  @override
  Widget build(BuildContext context) {
    String todayFormatted;
    try {
      todayFormatted = DateFormat('EEEE, dd MMMM yyyy', 'id_ID').format(DateTime.now());
    } catch (_) {
      todayFormatted = DateFormat('EEEE, dd MMMM yyyy').format(DateTime.now());
    }
    final bool isInRadius = _currentDistance <= (_office?.radiusMeters ?? 50.0) && !_isMocked;

    return Scaffold(
      backgroundColor: const Color(0xFFF1F5F9),
      appBar: AppBar(
        backgroundColor: Colors.white,
        elevation: 0,
        title: Row(
          children: [
            CircleAvatar(
              backgroundColor: const Color(0xFF3B82F6).withOpacity(0.15),
              child: Text(
                widget.user.name.isNotEmpty ? widget.user.name[0].toUpperCase() : 'U',
                style: const TextStyle(fontWeight: FontWeight.bold, color: Color(0xFF2563EB)),
              ),
            ),
            const SizedBox(width: 12),
            Column(
              crossAxisAlignment: CrossAxisAlignment.start,
              children: [
                Text(
                  widget.user.name,
                  style: const TextStyle(fontSize: 16, fontWeight: FontWeight.bold, color: Color(0xFF0F172A)),
                ),
                Text(
                  "${widget.user.department} • ${_office?.name ?? 'Kantor'}",
                  style: TextStyle(fontSize: 12, color: Colors.grey[600]),
                ),
              ],
            ),
          ],
        ),
        actions: [
          IconButton(
            icon: const Icon(Icons.history_rounded, color: Color(0xFF475569)),
            tooltip: "Riwayat Presensi",
            onPressed: () {
              Navigator.push(
                context,
                MaterialPageRoute(
                  builder: (context) => HistoryView(user: widget.user),
                ),
              );
            },
          ),
          IconButton(
            icon: const Icon(Icons.logout_rounded, color: Color(0xFFEF4444)),
            tooltip: "Keluar",
            onPressed: _handleLogout,
          ),
        ],
      ),
      body: StreamBuilder<AttendanceModel?>(
        stream: _attendanceService.streamTodayAttendance(widget.user.uid),
        builder: (context, snapshot) {
          final todayAttendance = snapshot.data;
          final bool hasCheckedIn = todayAttendance?.checkIn != null;
          final bool hasCheckedOut = todayAttendance?.checkOut != null;

          return SingleChildScrollView(
            padding: const EdgeInsets.symmetric(horizontal: 20, vertical: 18),
            child: Column(
              crossAxisAlignment: CrossAxisAlignment.stretch,
              children: [
                // Digital Clock Card
                Container(
                  padding: const EdgeInsets.symmetric(vertical: 20, horizontal: 24),
                  decoration: BoxDecoration(
                    gradient: const LinearGradient(
                      colors: [Color(0xFF1E293B), Color(0xFF0F172A)],
                      begin: Alignment.topLeft,
                      end: Alignment.bottomRight,
                    ),
                    borderRadius: BorderRadius.circular(24),
                    boxShadow: [
                      BoxShadow(
                        color: Colors.black.withOpacity(0.15),
                        blurRadius: 15,
                        offset: const Offset(0, 6),
                      ),
                    ],
                  ),
                  child: Column(
                    children: [
                      Text(
                        todayFormatted,
                        style: TextStyle(color: Colors.grey[400], fontSize: 13),
                      ),
                      const SizedBox(height: 6),
                      Text(
                        _currentTimeString,
                        style: const TextStyle(
                          fontSize: 36,
                          fontWeight: FontWeight.w800,
                          color: Colors.white,
                          letterSpacing: 2,
                        ),
                      ),
                      const SizedBox(height: 8),
                      Container(
                        padding: const EdgeInsets.symmetric(horizontal: 12, vertical: 4),
                        decoration: BoxDecoration(
                          color: Colors.white.withOpacity(0.1),
                          borderRadius: BorderRadius.circular(20),
                        ),
                        child: Text(
                          "Jam Kerja: ${_office?.workStartTime ?? '08:00'} - ${_office?.workEndTime ?? '17:00'}",
                          style: const TextStyle(color: Colors.white70, fontSize: 12),
                        ),
                      )
                    ],
                  ),
                ),
                const SizedBox(height: 20),

                // Radar Location Widget
                RadarLocationWidget(
                  distanceMeters: _currentDistance,
                  radiusMeters: _office?.radiusMeters ?? 50.0,
                  isChecking: _isCheckingLocation,
                  isMocked: _isMocked,
                  onRefresh: _checkDistance,
                ),
                const SizedBox(height: 20),

                // Attendance Status Summary (Masuk & Pulang)
                Row(
                  children: [
                    Expanded(
                      child: _buildTimeBox(
                        title: "Jam Masuk",
                        time: hasCheckedIn
                            ? DateFormat('HH:mm:ss').format(todayAttendance!.checkIn!.time)
                            : "--:--:--",
                        status: hasCheckedIn
                            ? (todayAttendance!.checkIn!.status == 'LATE' ? 'Terlambat' : 'Tepat Waktu')
                            : 'Belum Absen',
                        statusColor: hasCheckedIn
                            ? (todayAttendance!.checkIn!.status == 'LATE' ? Colors.orange : Colors.green)
                            : Colors.grey,
                        icon: Icons.login_rounded,
                      ),
                    ),
                    const SizedBox(width: 14),
                    Expanded(
                      child: _buildTimeBox(
                        title: "Jam Pulang",
                        time: hasCheckedOut
                            ? DateFormat('HH:mm:ss').format(todayAttendance!.checkOut!.time)
                            : "--:--:--",
                        status: hasCheckedOut ? 'Selesai' : (hasCheckedIn ? 'Belum Pulang' : '-'),
                        statusColor: hasCheckedOut ? Colors.blue : Colors.grey,
                        icon: Icons.logout_rounded,
                      ),
                    ),
                  ],
                ),
                const SizedBox(height: 24),

                // Main Action Button
                if (!hasCheckedIn) ...[
                  ElevatedButton.icon(
                    icon: _isLoadingAction
                        ? const SizedBox(
                            width: 20,
                            height: 20,
                            child: CircularProgressIndicator(color: Colors.white, strokeWidth: 2),
                          )
                        : const Icon(Icons.touch_app_rounded, color: Colors.white, size: 24),
                    label: Text(
                      _isLoadingAction
                          ? "Memproses Presensi..."
                          : (isInRadius ? "ABSEN MASUK SEKARANG" : "LOKASI DILUAR RADIUS (<= 50M)"),
                      style: const TextStyle(
                        fontSize: 15,
                        fontWeight: FontWeight.bold,
                        color: Colors.white,
                        letterSpacing: 0.5,
                      ),
                    ),
                    onPressed: (isInRadius && !_isLoadingAction) ? _handleCheckIn : null,
                    style: ElevatedButton.styleFrom(
                      padding: const EdgeInsets.symmetric(vertical: 18),
                      backgroundColor: const Color(0xFF10B981),
                      disabledBackgroundColor: Colors.grey[400],
                      shape: RoundedRectangleBorder(borderRadius: BorderRadius.circular(18)),
                      elevation: isInRadius ? 4 : 0,
                    ),
                  ),
                ] else if (!hasCheckedOut) ...[
                  ElevatedButton.icon(
                    icon: _isLoadingAction
                        ? const SizedBox(
                            width: 20,
                            height: 20,
                            child: CircularProgressIndicator(color: Colors.white, strokeWidth: 2),
                          )
                        : const Icon(Icons.exit_to_app_rounded, color: Colors.white, size: 24),
                    label: Text(
                      _isLoadingAction
                          ? "Memproses Presensi..."
                          : (isInRadius ? "ABSEN PULANG SEKARANG" : "LOKASI DILUAR RADIUS (<= 50M)"),
                      style: const TextStyle(
                        fontSize: 15,
                        fontWeight: FontWeight.bold,
                        color: Colors.white,
                        letterSpacing: 0.5,
                      ),
                    ),
                    onPressed: (isInRadius && !_isLoadingAction) ? _handleCheckOut : null,
                    style: ElevatedButton.styleFrom(
                      padding: const EdgeInsets.symmetric(vertical: 18),
                      backgroundColor: const Color(0xFF3B82F6),
                      disabledBackgroundColor: Colors.grey[400],
                      shape: RoundedRectangleBorder(borderRadius: BorderRadius.circular(18)),
                      elevation: isInRadius ? 4 : 0,
                    ),
                  ),
                ] else ...[
                  Container(
                    padding: const EdgeInsets.all(16),
                    decoration: BoxDecoration(
                      color: const Color(0xFFECFDF5),
                      borderRadius: BorderRadius.circular(18),
                      border: Border.all(color: const Color(0xFF10B981)),
                    ),
                    child: const Row(
                      mainAxisAlignment: MainAxisAlignment.center,
                      children: [
                        Icon(Icons.check_circle, color: Color(0xFF10B981), size: 24),
                        SizedBox(width: 10),
                        Text(
                          "Presensi Hari Ini Sudah Lengkap",
                          style: TextStyle(
                            fontSize: 15,
                            fontWeight: FontWeight.bold,
                            color: Color(0xFF065F46),
                          ),
                        ),
                      ],
                    ),
                  ),
                ],
              ],
            ),
          );
        },
      ),
    );
  }

  Widget _buildTimeBox({
    required String title,
    required String time,
    required String status,
    required Color statusColor,
    required IconData icon,
  }) {
    return Container(
      padding: const EdgeInsets.all(16),
      decoration: BoxDecoration(
        color: Colors.white,
        borderRadius: BorderRadius.circular(20),
        border: Border.all(color: const Color(0xFFE2E8F0)),
      ),
      child: Column(
        crossAxisAlignment: CrossAxisAlignment.start,
        children: [
          Row(
            children: [
              Icon(icon, size: 18, color: const Color(0xFF64748B)),
              const SizedBox(width: 6),
              Text(
                title,
                style: const TextStyle(fontSize: 13, fontWeight: FontWeight.w600, color: Color(0xFF64748B)),
              ),
            ],
          ),
          const SizedBox(height: 10),
          Text(
            time,
            style: const TextStyle(fontSize: 18, fontWeight: FontWeight.w800, color: Color(0xFF0F172A)),
          ),
          const SizedBox(height: 6),
          Container(
            padding: const EdgeInsets.symmetric(horizontal: 8, vertical: 3),
            decoration: BoxDecoration(
              color: statusColor.withOpacity(0.12),
              borderRadius: BorderRadius.circular(8),
            ),
            child: Text(
              status,
              style: TextStyle(fontSize: 11, fontWeight: FontWeight.bold, color: statusColor),
            ),
          ),
        ],
      ),
    );
  }
}
