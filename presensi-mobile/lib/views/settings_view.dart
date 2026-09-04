import 'package:flutter/material.dart';
import '../models/user_model.dart';
import '../models/office_model.dart';
import '../core/services/auth_service.dart';
import '../core/services/attendance_service.dart';
import '../core/utils/custom_toast.dart';
import 'login_view.dart';
import 'analytics_view.dart';
import 'daily_monitoring_view.dart';
import 'cepu_map_report_view.dart';
import 'home_attendance_view.dart';

class SettingsView extends StatefulWidget {
  final UserModel user;
  final OfficeModel? office;

  const SettingsView({
    Key? key,
    required this.user,
    this.office,
  }) : super(key: key);

  @override
  State<SettingsView> createState() => _SettingsViewState();
}

class _SettingsViewState extends State<SettingsView> {
  final AuthService _authService = AuthService();
  final AttendanceService _attendanceService = AttendanceService();

  OfficeModel? _office;
  bool _pushNotifications = true;
  bool _darkMode = false;

  @override
  void initState() {
    super.initState();
    _office = widget.office;
    if (_office == null) {
      _loadOfficeData();
    }
  }

  Future<void> _loadOfficeData() async {
    try {
      final office = await _attendanceService.getOfficeConfig(widget.user.officeId);
      if (mounted) setState(() => _office = office);
    } catch (_) {}
  }

  void _handleLogout() async {
    final confirmed = await showDialog<bool>(
      context: context,
      builder: (ctx) => AlertDialog(
        shape: RoundedRectangleBorder(borderRadius: BorderRadius.circular(20)),
        title: const Text("Konfirmasi Keluar", style: TextStyle(fontWeight: FontWeight.bold, fontSize: 16)),
        content: const Text("Apakah Anda yakin ingin keluar dari akun Presensi Mobile?", style: TextStyle(fontSize: 13.5, color: Color(0xFF475569))),
        actions: [
          TextButton(
            onPressed: () => Navigator.pop(ctx, false),
            child: const Text("Batal", style: TextStyle(color: Color(0xFF64748B), fontWeight: FontWeight.bold)),
          ),
          ElevatedButton(
            onPressed: () => Navigator.pop(ctx, true),
            style: ElevatedButton.styleFrom(
              backgroundColor: const Color(0xFFEF4444),
              elevation: 0,
              shape: RoundedRectangleBorder(borderRadius: BorderRadius.circular(10)),
            ),
            child: const Text("Ya, Keluar", style: TextStyle(color: Colors.white, fontWeight: FontWeight.bold)),
          ),
        ],
      ),
    );

    if (confirmed == true) {
      await _authService.logout();
      if (!mounted) return;
      Navigator.pushAndRemoveUntil(
        context,
        MaterialPageRoute(builder: (context) => const LoginView()),
        (route) => false,
      );
    }
  }

  void _showProfileDetailDialog() {
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
              child: ElevatedButton(
                onPressed: () => Navigator.pop(ctx),
                style: ElevatedButton.styleFrom(
                  backgroundColor: const Color(0xFF1E60F2),
                  shape: RoundedRectangleBorder(borderRadius: BorderRadius.circular(14)),
                  elevation: 0,
                ),
                child: const Text("Tutup", style: TextStyle(color: Colors.white, fontWeight: FontWeight.bold)),
              ),
            ),
          ],
        ),
      ),
    );
  }

  void _showHelpCenterDialog() {
    showDialog(
      context: context,
      builder: (ctx) => AlertDialog(
        shape: RoundedRectangleBorder(borderRadius: BorderRadius.circular(20)),
        title: Row(
          children: const [
            Icon(Icons.help_center_rounded, color: Color(0xFF1E60F2), size: 22),
            SizedBox(width: 8),
            Text("Pusat Bantuan", style: TextStyle(fontSize: 16, fontWeight: FontWeight.bold)),
          ],
        ),
        content: Column(
          mainAxisSize: MainAxisSize.min,
          crossAxisAlignment: CrossAxisAlignment.start,
          children: const [
            Text(
              "• Izin Waigama:\nGunakan tombol Mulai Izin saat hendak meninggalkan kantor, dan segera Rekam Waktu Kembali setelah tiba di kantor.",
              style: TextStyle(fontSize: 12.5, height: 1.4, color: Color(0xFF334155)),
            ),
            SizedBox(height: 10),
            Text(
              "• Fitur Cepu:\nLaporkan rekan yang tidak berada di kantor tanpa izin. Laporan akan sah jika diverifikasi minimal oleh 4 rekan kerja sebelum jam 16:00 WIT (Senin-Kamis) atau 16:30 WIT (Jumat).",
              style: TextStyle(fontSize: 12.5, height: 1.4, color: Color(0xFF334155)),
            ),
          ],
        ),
        actions: [
          TextButton(
            onPressed: () => Navigator.pop(ctx),
            child: const Text("Mengerti", style: TextStyle(color: Color(0xFF1E60F2), fontWeight: FontWeight.bold)),
          ),
        ],
      ),
    );
  }

  void _showAboutDialog() {
    showDialog(
      context: context,
      builder: (ctx) => AlertDialog(
        shape: RoundedRectangleBorder(borderRadius: BorderRadius.circular(20)),
        title: Row(
          children: const [
            Icon(Icons.info_rounded, color: Color(0xFF1E60F2), size: 22),
            SizedBox(width: 8),
            Text("Tentang Aplikasi", style: TextStyle(fontSize: 16, fontWeight: FontWeight.bold)),
          ],
        ),
        content: Column(
          mainAxisSize: MainAxisSize.min,
          crossAxisAlignment: CrossAxisAlignment.start,
          children: const [
            Text("Presensi Mobile - Waigama & Cepu", style: TextStyle(fontWeight: FontWeight.bold, fontSize: 13.5, color: Color(0xFF0F172A))),
            SizedBox(height: 4),
            Text("Versi: 1.2.0 (Build Release)", style: TextStyle(fontSize: 12, color: Color(0xFF64748B))),
            SizedBox(height: 10),
            Text(
              "Aplikasi presensi dan pemantauan aktivitas kehadiran pegawai berbasis GPS dan verifikasi rekan.",
              style: TextStyle(fontSize: 12.5, height: 1.4, color: Color(0xFF334155)),
            ),
            SizedBox(height: 8),
            Text("BPS Kabupaten Raja Ampat", style: TextStyle(fontSize: 11.5, fontWeight: FontWeight.bold, color: Color(0xFF1E60F2))),
          ],
        ),
        actions: [
          TextButton(
            onPressed: () => Navigator.pop(ctx),
            child: const Text("Tutup", style: TextStyle(color: Color(0xFF1E60F2), fontWeight: FontWeight.bold)),
          ),
        ],
      ),
    );
  }

  // Navigasi Bottom Bar
  void _navigateToHome() {
    Navigator.pushAndRemoveUntil(
      context,
      MaterialPageRoute(builder: (context) => HomeAttendanceView(user: widget.user)),
      (route) => false,
    );
  }

  void _navigateToAnalytics() {
    Navigator.push(
      context,
      MaterialPageRoute(
        builder: (context) => AnalyticsView(user: widget.user, office: _office),
      ),
    );
  }

  void _navigateToDailyMonitoring() {
    Navigator.push(
      context,
      MaterialPageRoute(
        builder: (context) => DailyMonitoringView(user: widget.user, office: _office),
      ),
    );
  }

  void _navigateToCepuReport() async {
    final office = _office ?? await _attendanceService.getOfficeConfig(widget.user.officeId);
    if (office == null || !mounted) return;
    Navigator.push(
      context,
      MaterialPageRoute(
        builder: (context) => CepuMapReportView(
          reporter: widget.user,
          office: office,
        ),
      ),
    );
  }

  @override
  Widget build(BuildContext context) {
    return Scaffold(
      backgroundColor: const Color(0xFFF8FAFC),
      appBar: AppBar(
        elevation: 0,
        backgroundColor: Colors.white,
        leading: IconButton(
          icon: const Icon(Icons.arrow_back_ios_new_rounded, color: Color(0xFF0F172A), size: 18),
          onPressed: () {
            if (Navigator.canPop(context)) {
              Navigator.pop(context);
            } else {
              _navigateToHome();
            }
          },
        ),
        title: const Text(
          "Settings",
          style: TextStyle(
            color: Color(0xFF0F172A),
            fontSize: 18,
            fontWeight: FontWeight.bold,
          ),
        ),
        centerTitle: true,
      ),
      body: SafeArea(
        child: SingleChildScrollView(
          padding: const EdgeInsets.fromLTRB(20, 16, 20, 24),
          child: Column(
            crossAxisAlignment: CrossAxisAlignment.start,
            children: [
              // =========================================================
              // PROFILE CARD
              // Menampilkan Avatar, Nama Pegawai, dan Username Saja
              // (Sesuai Permintaan: Hilangkan jabatan & ID, ganti username)
              // =========================================================
              Container(
                width: double.infinity,
                padding: const EdgeInsets.all(20),
                decoration: BoxDecoration(
                  gradient: const LinearGradient(
                    colors: [Color(0xFF1E60F2), Color(0xFF0E3A99)],
                    begin: Alignment.topLeft,
                    end: Alignment.bottomRight,
                  ),
                  borderRadius: BorderRadius.circular(22),
                  boxShadow: [
                    BoxShadow(
                      color: const Color(0xFF1E60F2).withOpacity(0.32),
                      blurRadius: 16,
                      offset: const Offset(0, 6),
                    ),
                  ],
                ),
                child: Row(
                  children: [
                    // Avatar Lingkaran Elegan
                    Container(
                      width: 58,
                      height: 58,
                      decoration: BoxDecoration(
                        color: Colors.white.withOpacity(0.2),
                        shape: BoxShape.circle,
                        border: Border.all(color: Colors.white.withOpacity(0.4), width: 1.5),
                      ),
                      alignment: Alignment.center,
                      child: Text(
                        widget.user.name.isNotEmpty ? widget.user.name[0].toUpperCase() : 'U',
                        style: const TextStyle(
                          fontSize: 24,
                          fontWeight: FontWeight.w900,
                          color: Colors.white,
                        ),
                      ),
                    ),
                    const SizedBox(width: 16),
                    // Nama & Username Saja
                    Expanded(
                      child: Column(
                        crossAxisAlignment: CrossAxisAlignment.start,
                        children: [
                          Text(
                            widget.user.name,
                            style: const TextStyle(
                              fontSize: 17,
                              fontWeight: FontWeight.bold,
                              color: Colors.white,
                              height: 1.2,
                            ),
                          ),
                          const SizedBox(height: 5),
                          Text(
                            "@${widget.user.username}",
                            style: TextStyle(
                              fontSize: 13.5,
                              color: Colors.white.withOpacity(0.85),
                              fontWeight: FontWeight.w500,
                            ),
                          ),
                        ],
                      ),
                    ),
                    IconButton(
                      icon: const Icon(Icons.arrow_forward_ios_rounded, color: Colors.white70, size: 16),
                      onPressed: _showProfileDetailDialog,
                      tooltip: "Rincian Profil",
                    ),
                  ],
                ),
              ),

              const SizedBox(height: 24),

              // =========================================================
              // SECTION 1: ACCOUNT (AKUN & KEAMANAN)
              // =========================================================
              _buildSectionTitle("Akun & Keamanan"),
              const SizedBox(height: 8),
              _buildSectionCard([
                _buildSettingItem(
                  icon: Icons.person_outline_rounded,
                  label: "Profile Settings",
                  subtitle: "Informasi profil dan akun pegawai",
                  onTap: _showProfileDetailDialog,
                ),
                _buildDivider(),
                _buildSettingItem(
                  icon: Icons.lock_outline_rounded,
                  label: "Privacy & Security",
                  subtitle: "Kata sandi dan keamanan akun",
                  onTap: () {
                    AppToast.showInfo(context, "Fitur keamanan kata sandi aktif dan terenkripsi.", title: "Keamanan Akun");
                  },
                ),
                _buildDivider(),
                _buildSettingItem(
                  icon: Icons.notifications_none_rounded,
                  label: "Notifications",
                  subtitle: "Notifikasi pengingat kembali & cepu",
                  isSwitch: true,
                  switchValue: _pushNotifications,
                  onSwitchChanged: (val) {
                    setState(() => _pushNotifications = val);
                    AppToast.showSuccess(
                      context,
                      val ? "Notifikasi diaktifkan" : "Notifikasi dinonaktifkan",
                    );
                  },
                ),
              ]),

              const SizedBox(height: 24),

              // =========================================================
              // SECTION 2: PREFERENCES (PREFERENSI)
              // =========================================================
              _buildSectionTitle("Preferensi"),
              const SizedBox(height: 8),
              _buildSectionCard([
                _buildSettingItem(
                  icon: Icons.language_rounded,
                  label: "Language",
                  subtitle: "Bahasa Indonesia",
                  onTap: () {
                    AppToast.showInfo(context, "Aplikasi menggunakan Bahasa Indonesia secara bawaan.", title: "Bahasa");
                  },
                ),
                _buildDivider(),
                _buildSettingItem(
                  icon: Icons.dark_mode_outlined,
                  label: "Dark Mode",
                  subtitle: _darkMode ? "Aktif" : "Nonaktif (Mode Terang)",
                  isSwitch: true,
                  switchValue: _darkMode,
                  onSwitchChanged: (val) {
                    setState(() => _darkMode = val);
                    AppToast.showInfo(context, "Tema gelap akan didukung pada pembaruan mendatang.", title: "Preferensi Tema");
                  },
                ),
              ]),

              const SizedBox(height: 24),

              // =========================================================
              // SECTION 3: SUPPORT (DUKUNGAN & TENTANG)
              // =========================================================
              _buildSectionTitle("Dukungan & Informasi"),
              const SizedBox(height: 8),
              _buildSectionCard([
                _buildSettingItem(
                  icon: Icons.help_outline_rounded,
                  label: "Help Center",
                  subtitle: "Panduan presensi Waigama & Cepu",
                  onTap: _showHelpCenterDialog,
                ),
                _buildDivider(),
                _buildSettingItem(
                  icon: Icons.info_outline_rounded,
                  label: "About App",
                  subtitle: "Versi 1.2.0 • BPS Raja Ampat",
                  onTap: _showAboutDialog,
                ),
              ]),

              const SizedBox(height: 28),

              // =========================================================
              // LOGOUT BUTTON (Tombol Keluar Akun)
              // =========================================================
              InkWell(
                onTap: _handleLogout,
                borderRadius: BorderRadius.circular(18),
                child: Container(
                  width: double.infinity,
                  padding: const EdgeInsets.symmetric(vertical: 14),
                  decoration: BoxDecoration(
                    color: const Color(0xFFFEF2F2),
                    borderRadius: BorderRadius.circular(18),
                    border: Border.all(color: const Color(0xFFFECACA)),
                  ),
                  child: Row(
                    mainAxisAlignment: MainAxisAlignment.center,
                    children: const [
                      Icon(Icons.logout_rounded, color: Color(0xFFDC2626), size: 20),
                      SizedBox(width: 8),
                      Text(
                        "Keluar Akun",
                        style: TextStyle(
                          color: Color(0xFFDC2626),
                          fontWeight: FontWeight.bold,
                          fontSize: 14.5,
                        ),
                      ),
                    ],
                  ),
                ),
              ),
            ],
          ),
        ),
      ),

      // BOTTOM NAVIGATION BAR
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
              icon: const Icon(Icons.home_rounded, color: Color(0xFF94A3B8), size: 26),
              onPressed: _navigateToHome,
              tooltip: "Home",
            ),
            IconButton(
              icon: const Icon(Icons.bar_chart_rounded, color: Color(0xFF94A3B8), size: 26),
              onPressed: _navigateToAnalytics,
              tooltip: "Insight",
            ),
            IconButton(
              icon: const Icon(Icons.fact_check_outlined, color: Color(0xFF94A3B8), size: 26),
              onPressed: _navigateToDailyMonitoring,
              tooltip: "Monitoring Harian",
            ),
            IconButton(
              icon: const Icon(Icons.campaign_outlined, color: Color(0xFF94A3B8), size: 26),
              onPressed: _navigateToCepuReport,
              tooltip: "Cepu",
            ),
            IconButton(
              icon: const Icon(Icons.settings_rounded, color: Color(0xFF1E60F2), size: 26),
              onPressed: () {},
              tooltip: "Settings",
            ),
          ],
        ),
      ),
    );
  }

  Widget _buildSectionTitle(String title) {
    return Padding(
      padding: const EdgeInsets.only(left: 4),
      child: Text(
        title.toUpperCase(),
        style: const TextStyle(
          fontSize: 11.5,
          fontWeight: FontWeight.w800,
          color: Color(0xFF64748B),
          letterSpacing: 0.8,
        ),
      ),
    );
  }

  Widget _buildSectionCard(List<Widget> children) {
    return Container(
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
        children: children,
      ),
    );
  }

  Widget _buildDivider() {
    return const Divider(height: 1, indent: 60, color: Color(0xFFF1F5F9));
  }

  Widget _buildSettingItem({
    required IconData icon,
    required String label,
    String? subtitle,
    VoidCallback? onTap,
    bool isSwitch = false,
    bool switchValue = false,
    ValueChanged<bool>? onSwitchChanged,
  }) {
    return InkWell(
      onTap: isSwitch ? () => onSwitchChanged?.call(!switchValue) : onTap,
      borderRadius: BorderRadius.circular(20),
      child: Padding(
        padding: const EdgeInsets.symmetric(horizontal: 16, vertical: 14),
        child: Row(
          children: [
            Container(
              width: 40,
              height: 40,
              decoration: BoxDecoration(
                color: const Color(0xFFEFF6FF),
                borderRadius: BorderRadius.circular(12),
              ),
              child: Icon(icon, color: const Color(0xFF1E60F2), size: 20),
            ),
            const SizedBox(width: 14),
            Expanded(
              child: Column(
                crossAxisAlignment: CrossAxisAlignment.start,
                children: [
                  Text(
                    label,
                    style: const TextStyle(
                      fontSize: 14,
                      fontWeight: FontWeight.bold,
                      color: Color(0xFF0F172A),
                    ),
                  ),
                  if (subtitle != null) ...[
                    const SizedBox(height: 2),
                    Text(
                      subtitle,
                      style: const TextStyle(
                        fontSize: 12,
                        color: Color(0xFF64748B),
                      ),
                    ),
                  ],
                ],
              ),
            ),
            if (isSwitch)
              Switch(
                value: switchValue,
                onChanged: onSwitchChanged,
                activeColor: const Color(0xFF1E60F2),
              )
            else
              const Icon(Icons.arrow_forward_ios_rounded, color: Color(0xFFCBD5E1), size: 14),
          ],
        ),
      ),
    );
  }
}
