import 'dart:convert';
import 'package:flutter/material.dart';
import 'package:intl/intl.dart';
import '../models/permit_model.dart';
import '../models/cepu_model.dart';
import '../models/user_model.dart';
import '../core/services/permit_service.dart';
import '../core/services/cepu_service.dart';

class DailyMonitoringView extends StatefulWidget {
  final UserModel user;

  const DailyMonitoringView({Key? key, required this.user}) : super(key: key);

  @override
  State<DailyMonitoringView> createState() => _DailyMonitoringViewState();
}

class _DailyMonitoringViewState extends State<DailyMonitoringView> with SingleTickerProviderStateMixin {
  late TabController _tabController;
  final PermitService _permitService = PermitService();
  final CepuService _cepuService = CepuService();
  final TextEditingController _searchController = TextEditingController();

  DateTime _selectedDate = DateTime.now();
  String _permitFilter = 'ALL'; // 'ALL', 'ACTIVE', 'COMPLETED'
  String _cepuFilter = 'ALL'; // 'ALL', 'PENDING', 'VERIFIED'
  String _searchQuery = '';

  String get _selectedDateString => DateFormat('yyyy-MM-dd').format(_selectedDate);

  @override
  void initState() {
    super.initState();
    _tabController = TabController(length: 2, vsync: this);
    _tabController.addListener(() {
      setState(() {});
    });
  }

  @override
  void dispose() {
    _tabController.dispose();
    _searchController.dispose();
    super.dispose();
  }

  void _previousDay() {
    setState(() => _selectedDate = _selectedDate.subtract(const Duration(days: 1)));
  }

  void _nextDay() {
    setState(() => _selectedDate = _selectedDate.add(const Duration(days: 1)));
  }

  void _goToToday() {
    setState(() => _selectedDate = DateTime.now());
  }

  Future<void> _pickDate() async {
    final picked = await showDatePicker(
      context: context,
      initialDate: _selectedDate,
      firstDate: DateTime(2020),
      lastDate: DateTime(2035),
      builder: (context, child) {
        return Theme(
          data: Theme.of(context).copyWith(
            colorScheme: const ColorScheme.light(
              primary: Color(0xFF1E60F2),
              onPrimary: Colors.white,
              onSurface: Color(0xFF1E293B),
            ),
          ),
          child: child!,
        );
      },
    );

    if (picked != null && picked != _selectedDate) {
      setState(() => _selectedDate = picked);
    }
  }

  bool _isToday(DateTime date) {
    final now = DateTime.now();
    return date.year == now.year && date.month == now.month && date.day == now.day;
  }

  /// Dialog Verifikasi Cepu Sesuai Permintaan User:
  /// - Pop-up dengan keterangan "Apakah laporan ini valid?"
  /// - Tombol "Ya"
  /// - Tombol Close (X) di kanan atas
  /// - Bisa ditutup dengan menekan area luar (barrierDismissible)
  void _showVerificationDialog(CepuModel cepu) {
    showDialog(
      context: context,
      barrierDismissible: true,
      builder: (dialogContext) {
        return Dialog(
          shape: RoundedRectangleBorder(borderRadius: BorderRadius.circular(22)),
          elevation: 10,
          child: Padding(
            padding: const EdgeInsets.all(20),
            child: Column(
              mainAxisSize: MainAxisSize.min,
              children: [
                // Header with Close Button at Top Right
                Row(
                  mainAxisAlignment: MainAxisAlignment.spaceBetween,
                  children: [
                    Container(
                      padding: const EdgeInsets.symmetric(horizontal: 10, vertical: 4),
                      decoration: BoxDecoration(
                        color: const Color(0xFFFFEDD5),
                        borderRadius: BorderRadius.circular(10),
                      ),
                      child: Row(
                        children: const [
                          Icon(Icons.verified_user_rounded, color: Color(0xFFEA580C), size: 16),
                          SizedBox(width: 4),
                          Text(
                            "Verifikasi Rekan",
                            style: TextStyle(
                              fontSize: 12,
                              fontWeight: FontWeight.bold,
                              color: Color(0xFFC2410C),
                            ),
                          ),
                        ],
                      ),
                    ),
                    IconButton(
                      icon: const Icon(Icons.close_rounded, color: Color(0xFF64748B), size: 22),
                      padding: EdgeInsets.zero,
                      constraints: const BoxConstraints(),
                      onPressed: () => Navigator.pop(dialogContext),
                    ),
                  ],
                ),

                const SizedBox(height: 16),

                // Center Icon / Illustration
                Container(
                  width: 56,
                  height: 56,
                  decoration: BoxDecoration(
                    color: const Color(0xFFEFF6FF),
                    shape: BoxShape.circle,
                    border: Border.all(color: const Color(0xFFBFDBFE), width: 2),
                  ),
                  child: const Icon(
                    Icons.help_outline_rounded,
                    color: Color(0xFF2563EB),
                    size: 32,
                  ),
                ),

                const SizedBox(height: 14),

                // Question Text
                const Text(
                  "Apakah laporan ini valid?",
                  style: TextStyle(
                    fontSize: 17,
                    fontWeight: FontWeight.bold,
                    color: Color(0xFF1E293B),
                  ),
                  textAlign: TextAlign.center,
                ),

                const SizedBox(height: 8),

                Text(
                  "Laporan ketidakhadiran pegawai atas nama ${cepu.targetName} (${cepu.durationString}).",
                  style: const TextStyle(
                    fontSize: 12,
                    color: Color(0xFF64748B),
                    height: 1.4,
                  ),
                  textAlign: TextAlign.center,
                ),

                const SizedBox(height: 20),

                // "Ya" Button
                SizedBox(
                  width: double.infinity,
                  height: 46,
                  child: ElevatedButton(
                    onPressed: () async {
                      Navigator.pop(dialogContext);
                      try {
                        await _cepuService.verifyCepuReport(
                          cepuId: cepu.id,
                          verifier: widget.user,
                        );
                        if (mounted) {
                          ScaffoldMessenger.of(context).showSnackBar(
                            const SnackBar(
                              content: Text("✅ Sukses memverifikasi laporan Cepu!"),
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
                      }
                    },
                    style: ElevatedButton.styleFrom(
                      backgroundColor: const Color(0xFF1E60F2),
                      shape: RoundedRectangleBorder(borderRadius: BorderRadius.circular(14)),
                      elevation: 2,
                    ),
                    child: const Text(
                      "Ya, Laporan Valid",
                      style: TextStyle(
                        color: Colors.white,
                        fontWeight: FontWeight.bold,
                        fontSize: 14,
                      ),
                    ),
                  ),
                ),
              ],
            ),
          ),
        );
      },
    );
  }

  void _showVerificatorsModal(CepuModel cepu) {
    showModalBottomSheet(
      context: context,
      backgroundColor: Colors.transparent,
      builder: (ctx) {
        return Container(
          decoration: const BoxDecoration(
            color: Colors.white,
            borderRadius: BorderRadius.vertical(top: Radius.circular(24)),
          ),
          padding: const EdgeInsets.fromLTRB(20, 14, 20, 24),
          child: Column(
            mainAxisSize: MainAxisSize.min,
            crossAxisAlignment: CrossAxisAlignment.stretch,
            children: [
              Center(
                child: Container(
                  width: 36,
                  height: 4,
                  decoration: BoxDecoration(
                    color: const Color(0xFFE2E8F0),
                    borderRadius: BorderRadius.circular(2),
                  ),
                ),
              ),
              const SizedBox(height: 14),
              Row(
                mainAxisAlignment: MainAxisAlignment.spaceBetween,
                children: [
                  Row(
                    children: const [
                      Icon(Icons.verified_user_rounded, color: Color(0xFF10B981), size: 22),
                      SizedBox(width: 8),
                      Text(
                        "Daftar Verifikator",
                        style: TextStyle(
                          fontSize: 16,
                          fontWeight: FontWeight.bold,
                          color: Color(0xFF1E293B),
                        ),
                      ),
                    ],
                  ),
                  Container(
                    padding: const EdgeInsets.symmetric(horizontal: 10, vertical: 4),
                    decoration: BoxDecoration(
                      color: cepu.isValid ? const Color(0xFFECFDF5) : const Color(0xFFFFF7ED),
                      borderRadius: BorderRadius.circular(8),
                    ),
                    child: Text(
                      cepu.isValid ? "VALID (${cepu.verificationCount} Rekan)" : "${cepu.verificationCount}/4 Rekan",
                      style: TextStyle(
                        fontSize: 11,
                        fontWeight: FontWeight.bold,
                        color: cepu.isValid ? const Color(0xFF047857) : const Color(0xFFEA580C),
                      ),
                    ),
                  ),
                ],
              ),
              const SizedBox(height: 6),
              Text(
                "Laporan atas nama ${cepu.targetName}",
                style: const TextStyle(fontSize: 12, color: Color(0xFF64748B)),
              ),
              const SizedBox(height: 14),
              const Divider(height: 1, color: Color(0xFFF1F5F9)),
              const SizedBox(height: 10),
              if (cepu.verifiedByNames.isEmpty)
                Container(
                  padding: const EdgeInsets.symmetric(vertical: 24),
                  child: Column(
                    children: const [
                      Icon(Icons.info_outline_rounded, size: 36, color: Color(0xFF94A3B8)),
                      SizedBox(height: 8),
                      Text(
                        "Belum ada verifikator",
                        style: TextStyle(
                          fontSize: 13,
                          fontWeight: FontWeight.w600,
                          color: Color(0xFF64748B),
                        ),
                      ),
                    ],
                  ),
                )
              else
                Flexible(
                  child: ListView.separated(
                    shrinkWrap: true,
                    itemCount: cepu.verifiedByNames.length,
                    separatorBuilder: (_, __) => const Divider(height: 1, color: Color(0xFFF1F5F9)),
                    itemBuilder: (context, idx) {
                      final name = cepu.verifiedByNames[idx];
                      return ListTile(
                        contentPadding: const EdgeInsets.symmetric(horizontal: 4, vertical: 2),
                        leading: CircleAvatar(
                          radius: 18,
                          backgroundColor: const Color(0xFFEFF6FF),
                          child: Text(
                            name.isNotEmpty ? name[0].toUpperCase() : 'U',
                            style: const TextStyle(
                              color: Color(0xFF2563EB),
                              fontWeight: FontWeight.bold,
                              fontSize: 13,
                            ),
                          ),
                        ),
                        title: Text(
                          name,
                          style: const TextStyle(
                            fontSize: 13.5,
                            fontWeight: FontWeight.bold,
                            color: Color(0xFF1E293B),
                          ),
                        ),
                        trailing: const Icon(Icons.check_circle_rounded, color: Color(0xFF10B981), size: 20),
                      );
                    },
                  ),
                ),
            ],
          ),
        );
      },
    );
  }

  void _showImagePreview(String base64String) {
    showDialog(
      context: context,
      barrierDismissible: true,
      builder: (context) {
        return Dialog(
          backgroundColor: Colors.transparent,
          insetPadding: const EdgeInsets.all(16),
          child: Stack(
            alignment: Alignment.topRight,
            children: [
              ClipRRect(
                borderRadius: BorderRadius.circular(16),
                child: Image.memory(
                  base64Decode(base64String),
                  fit: BoxFit.contain,
                ),
              ),
              Positioned(
                top: 8,
                right: 8,
                child: IconButton(
                  icon: Container(
                    padding: const EdgeInsets.all(6),
                    decoration: const BoxDecoration(color: Colors.black54, shape: BoxShape.circle),
                    child: const Icon(Icons.close, color: Colors.white, size: 18),
                  ),
                  onPressed: () => Navigator.pop(context),
                ),
              ),
            ],
          ),
        );
      },
    );
  }

  @override
  Widget build(BuildContext context) {
    String formattedDate;
    try {
      formattedDate = DateFormat('EEEE, d MMMM yyyy', 'id_ID').format(_selectedDate);
    } catch (_) {
      formattedDate = DateFormat('EEEE, d MMMM yyyy').format(_selectedDate);
    }

    final isCurrentDay = _isToday(_selectedDate);

    return Scaffold(
      backgroundColor: const Color(0xFFF8FAFC),
      appBar: AppBar(
        backgroundColor: const Color(0xFF1E60F2),
        elevation: 0,
        leading: IconButton(
          icon: const Icon(Icons.arrow_back_ios_new, color: Colors.white, size: 20),
          onPressed: () => Navigator.pop(context),
        ),
        title: const Text(
          "Monitoring Harian",
          style: TextStyle(color: Colors.white, fontWeight: FontWeight.bold, fontSize: 18),
        ),
        centerTitle: true,
        actions: [
          if (!isCurrentDay)
            TextButton(
              onPressed: _goToToday,
              child: const Text(
                "Hari Ini",
                style: TextStyle(color: Colors.white, fontWeight: FontWeight.bold, fontSize: 13),
              ),
            ),
        ],
        bottom: TabBar(
          controller: _tabController,
          indicatorColor: Colors.white,
          indicatorWeight: 3.5,
          labelColor: Colors.white,
          unselectedLabelColor: Colors.white70,
          labelStyle: const TextStyle(fontWeight: FontWeight.bold, fontSize: 13),
          tabs: const [
            Tab(icon: Icon(Icons.badge_rounded, size: 20), text: "Waigama"),
            Tab(icon: Icon(Icons.campaign_rounded, size: 20), text: "Cepu"),
          ],
        ),
      ),
      body: Column(
        children: [
          // Date Selector Header Bar
          Container(
            padding: const EdgeInsets.symmetric(horizontal: 14, vertical: 8),
            decoration: BoxDecoration(
              color: Colors.white,
              boxShadow: [
                BoxShadow(
                  color: Colors.black.withOpacity(0.04),
                  blurRadius: 10,
                  offset: const Offset(0, 4),
                ),
              ],
            ),
            child: Row(
              children: [
                IconButton(
                  icon: const Icon(Icons.chevron_left_rounded, color: Color(0xFF1E60F2), size: 28),
                  onPressed: _previousDay,
                  tooltip: "Hari Sebelumnya",
                ),
                Expanded(
                  child: InkWell(
                    onTap: _pickDate,
                    borderRadius: BorderRadius.circular(12),
                    child: Container(
                      padding: const EdgeInsets.symmetric(vertical: 8, horizontal: 12),
                      decoration: BoxDecoration(
                        color: const Color(0xFFEFF6FF),
                        borderRadius: BorderRadius.circular(12),
                        border: Border.all(color: const Color(0xFFDBEAFE)),
                      ),
                      child: Row(
                        mainAxisAlignment: MainAxisAlignment.center,
                        children: [
                          const Icon(Icons.calendar_today_rounded, size: 15, color: Color(0xFF1E60F2)),
                          const SizedBox(width: 8),
                          Flexible(
                            child: Text(
                              formattedDate,
                              style: const TextStyle(
                                fontSize: 12.5,
                                fontWeight: FontWeight.bold,
                                color: Color(0xFF1E40AF),
                              ),
                              overflow: TextOverflow.ellipsis,
                            ),
                          ),
                          const SizedBox(width: 4),
                          const Icon(Icons.arrow_drop_down_rounded, color: Color(0xFF1E60F2)),
                        ],
                      ),
                    ),
                  ),
                ),
                IconButton(
                  icon: const Icon(Icons.chevron_right_rounded, color: Color(0xFF1E60F2), size: 28),
                  onPressed: _nextDay,
                  tooltip: "Hari Berikutnya",
                ),
              ],
            ),
          ),

          // Tab Bar Views
          Expanded(
            child: TabBarView(
              controller: _tabController,
              children: [
                // TAB 1: Izin Kantor
                _buildPermitsTab(),

                // TAB 2: Laporan Cepu
                _buildCepuTab(),
              ],
            ),
          ),
        ],
      ),
    );
  }

  // ==========================================
  // TAB 1: PERMITS (IZIN KEGIATAN KANTOR)
  // ==========================================
  Widget _buildPermitsTab() {
    return StreamBuilder<List<PermitModel>>(
      stream: _permitService.streamDailyPermits(_selectedDateString),
      builder: (context, snapshot) {
        if (snapshot.connectionState == ConnectionState.waiting) {
          return const Center(child: CircularProgressIndicator(color: Color(0xFF1E60F2)));
        }

        final allPermits = snapshot.data ?? [];
        final filteredPermits = allPermits.where((p) {
          final matchesSearch = _searchQuery.isEmpty ||
              p.userName.toLowerCase().contains(_searchQuery.toLowerCase()) ||
              p.purpose.toLowerCase().contains(_searchQuery.toLowerCase());
          if (!matchesSearch) return false;

          if (_permitFilter == 'ACTIVE') return p.isActive;
          if (_permitFilter == 'COMPLETED') return !p.isActive;
          return true;
        }).toList();

        final activeCount = allPermits.where((p) => p.isActive).length;
        final completedCount = allPermits.where((p) => !p.isActive).length;

        return Column(
          children: [
            // Summary Filter Cards (Poin 7 & 8: Total, Sedang di Luar, Sudah Kembali)
            Padding(
              padding: const EdgeInsets.fromLTRB(16, 12, 16, 6),
              child: Row(
                children: [
                  _buildSummaryCard(
                    label: "Total",
                    count: allPermits.length.toString(),
                    color: const Color(0xFF3B82F6),
                    icon: Icons.list_alt_rounded,
                    isSelected: _permitFilter == 'ALL',
                    onTap: () => setState(() => _permitFilter = 'ALL'),
                  ),
                  const SizedBox(width: 8),
                  _buildSummaryCard(
                    label: "Sedang di Luar",
                    count: activeCount.toString(),
                    color: const Color(0xFFF59E0B),
                    icon: Icons.directions_walk_rounded,
                    isSelected: _permitFilter == 'ACTIVE',
                    onTap: () => setState(() => _permitFilter = 'ACTIVE'),
                  ),
                  const SizedBox(width: 8),
                  _buildSummaryCard(
                    label: "Sudah Kembali",
                    count: completedCount.toString(),
                    color: const Color(0xFF10B981),
                    icon: Icons.check_circle_outline_rounded,
                    isSelected: _permitFilter == 'COMPLETED',
                    onTap: () => setState(() => _permitFilter = 'COMPLETED'),
                  ),
                ],
              ),
            ),

            // List
            Expanded(
              child: filteredPermits.isEmpty
                  ? _buildEmptyState("Belum ada catatan Waigama pada tanggal ini.")
                  : ListView.separated(
                      padding: const EdgeInsets.fromLTRB(16, 6, 16, 24),
                      itemCount: filteredPermits.length,
                      separatorBuilder: (context, index) => const SizedBox(height: 12),
                      itemBuilder: (context, index) => _buildPermitCard(filteredPermits[index]),
                    ),
            ),
          ],
        );
      },
    );
  }

  // ==========================================
  // TAB 2: LAPORAN CEPU
  // ==========================================
  Widget _buildCepuTab() {
    return StreamBuilder<List<CepuModel>>(
      stream: _cepuService.streamDailyCepuReports(_selectedDateString),
      builder: (context, snapshot) {
        if (snapshot.connectionState == ConnectionState.waiting) {
          return const Center(child: CircularProgressIndicator(color: Color(0xFFEA580C)));
        }

        final allCepu = snapshot.data ?? [];
        final filteredCepu = allCepu.where((c) {
          final matchesSearch = _searchQuery.isEmpty ||
              c.targetName.toLowerCase().contains(_searchQuery.toLowerCase()) ||
              c.reporterName.toLowerCase().contains(_searchQuery.toLowerCase()) ||
              c.description.toLowerCase().contains(_searchQuery.toLowerCase());
          if (!matchesSearch) return false;

          if (_cepuFilter == 'PENDING') return !c.isValid;
          if (_cepuFilter == 'VERIFIED') return c.isValid;
          return true;
        }).toList();

        final pendingCount = allCepu.where((c) => !c.isValid).length;
        final verifiedCount = allCepu.where((c) => c.isValid).length;

        return Column(
          children: [
            // Summary Filter Cards (Poin 8: Total Laporan, Butuh Verifikasi, Valid)
            Padding(
              padding: const EdgeInsets.fromLTRB(16, 12, 16, 6),
              child: Row(
                children: [
                  _buildSummaryCard(
                    label: "Total Laporan",
                    count: allCepu.length.toString(),
                    color: const Color(0xFFEA580C),
                    icon: Icons.campaign_rounded,
                    isSelected: _cepuFilter == 'ALL',
                    onTap: () => setState(() => _cepuFilter = 'ALL'),
                  ),
                  const SizedBox(width: 8),
                  _buildSummaryCard(
                    label: "Butuh Verifikasi",
                    count: pendingCount.toString(),
                    color: const Color(0xFFF59E0B),
                    icon: Icons.pending_actions_rounded,
                    isSelected: _cepuFilter == 'PENDING',
                    onTap: () => setState(() => _cepuFilter = 'PENDING'),
                  ),
                  const SizedBox(width: 8),
                  _buildSummaryCard(
                    label: "Valid",
                    count: verifiedCount.toString(),
                    color: const Color(0xFF10B981),
                    icon: Icons.verified_rounded,
                    isSelected: _cepuFilter == 'VERIFIED',
                    onTap: () => setState(() => _cepuFilter = 'VERIFIED'),
                  ),
                ],
              ),
            ),

            // List
            Expanded(
              child: filteredCepu.isEmpty
                  ? _buildEmptyState("Tidak ada laporan Cepu pada tanggal ini.")
                  : ListView.separated(
                      padding: const EdgeInsets.fromLTRB(16, 6, 16, 24),
                      itemCount: filteredCepu.length,
                      separatorBuilder: (context, index) => const SizedBox(height: 14),
                      itemBuilder: (context, index) => _buildCepuCard(filteredCepu[index]),
                    ),
            ),
          ],
        );
      },
    );
  }

  Widget _buildSummaryCard({
    required String label,
    required String count,
    required Color color,
    required IconData icon,
    required bool isSelected,
    required VoidCallback onTap,
  }) {
    return Expanded(
      child: Material(
        color: Colors.transparent,
        child: InkWell(
          onTap: onTap,
          borderRadius: BorderRadius.circular(14),
          child: AnimatedContainer(
            duration: const Duration(milliseconds: 200),
            padding: const EdgeInsets.symmetric(vertical: 10, horizontal: 6),
            decoration: BoxDecoration(
              color: isSelected ? color.withOpacity(0.12) : Colors.white,
              borderRadius: BorderRadius.circular(14),
              border: Border.all(
                color: isSelected ? color : color.withOpacity(0.25),
                width: isSelected ? 2.0 : 1.0,
              ),
              boxShadow: [
                BoxShadow(
                  color: isSelected ? color.withOpacity(0.15) : color.withOpacity(0.04),
                  blurRadius: isSelected ? 10 : 6,
                  offset: const Offset(0, 3),
                ),
              ],
            ),
            child: Column(
              children: [
                Row(
                  mainAxisAlignment: MainAxisAlignment.center,
                  children: [
                    Icon(icon, color: color, size: 16),
                    const SizedBox(width: 4),
                    Text(
                      count,
                      style: TextStyle(
                        fontSize: 17,
                        fontWeight: isSelected ? FontWeight.w900 : FontWeight.bold,
                        color: color,
                      ),
                    ),
                  ],
                ),
                const SizedBox(height: 3),
                Text(
                  label,
                  style: TextStyle(
                    fontSize: 9.5,
                    color: isSelected ? color : const Color(0xFF64748B),
                    fontWeight: isSelected ? FontWeight.bold : FontWeight.w600,
                  ),
                  textAlign: TextAlign.center,
                  maxLines: 1,
                  overflow: TextOverflow.ellipsis,
                ),
              ],
            ),
          ),
        ),
      ),
    );
  }

  Widget _buildEmptyState(String message) {
    return Center(
      child: Column(
        mainAxisAlignment: MainAxisAlignment.center,
        children: [
          Container(
            padding: const EdgeInsets.all(18),
            decoration: const BoxDecoration(color: Color(0xFFEFF6FF), shape: BoxShape.circle),
            child: const Icon(Icons.check_circle_outline_rounded, size: 42, color: Color(0xFF3B82F6)),
          ),
          const SizedBox(height: 12),
          Text(
            message,
            style: const TextStyle(fontSize: 13, fontWeight: FontWeight.w600, color: Color(0xFF475569)),
            textAlign: TextAlign.center,
          ),
        ],
      ),
    );
  }

  Widget _buildPermitCard(PermitModel permit) {
    final isActive = permit.isActive;
    final startTimeStr = DateFormat('HH:mm').format(permit.startTime);
    final endTimeStr = permit.endTime != null ? DateFormat('HH:mm').format(permit.endTime!) : '--:--';

    return Container(
      decoration: BoxDecoration(
        color: Colors.white,
        borderRadius: BorderRadius.circular(16),
        border: Border.all(
          color: isActive ? const Color(0xFFFCD34D) : const Color(0xFFE2E8F0),
          width: isActive ? 1.5 : 1.0,
        ),
        boxShadow: [
          BoxShadow(
            color: isActive ? const Color(0xFFF59E0B).withOpacity(0.08) : Colors.black.withOpacity(0.04),
            blurRadius: 10,
            offset: const Offset(0, 4),
          ),
        ],
      ),
      padding: const EdgeInsets.all(14),
      child: Column(
        crossAxisAlignment: CrossAxisAlignment.start,
        children: [
          Row(
            children: [
              CircleAvatar(
                radius: 17,
                backgroundColor: isActive ? const Color(0xFFFEF3C7) : const Color(0xFFEFF6FF),
                child: Text(
                  permit.userName.isNotEmpty ? permit.userName[0].toUpperCase() : 'U',
                  style: TextStyle(
                    color: isActive ? const Color(0xFFB45309) : const Color(0xFF1E60F2),
                    fontWeight: FontWeight.bold,
                    fontSize: 13,
                  ),
                ),
              ),
              const SizedBox(width: 10),
              Expanded(
                child: Column(
                  crossAxisAlignment: CrossAxisAlignment.start,
                  children: [
                    Text(
                      permit.userName,
                      style: const TextStyle(fontSize: 14, fontWeight: FontWeight.bold, color: Color(0xFF1E293B)),
                      maxLines: 1,
                      overflow: TextOverflow.ellipsis,
                    ),
                  ],
                ),
              ),
              Container(
                padding: const EdgeInsets.symmetric(horizontal: 8, vertical: 3),
                decoration: BoxDecoration(
                  color: permit.isExpired
                      ? const Color(0xFFFEF2F2)
                      : (isActive ? const Color(0xFFFFFBEB) : const Color(0xFFECFDF5)),
                  borderRadius: BorderRadius.circular(8),
                  border: Border.all(
                    color: permit.isExpired
                        ? const Color(0xFFFCA5A5)
                        : (isActive ? const Color(0xFFFDE68A) : const Color(0xFFA7F3D0)),
                  ),
                ),
                child: Text(
                  permit.isExpired ? "Kadaluarsa" : (isActive ? "Sedang Izin" : "Sudah Kembali"),
                  style: TextStyle(
                    fontSize: 10.5,
                    fontWeight: FontWeight.bold,
                    color: permit.isExpired
                        ? const Color(0xFF991B1B)
                        : (isActive ? const Color(0xFFB45309) : const Color(0xFF047857)),
                  ),
                ),
              ),
            ],
          ),

          const SizedBox(height: 10),
          const Divider(height: 1, color: Color(0xFFF1F5F9)),
          const SizedBox(height: 8),

          Text(
            "Keperluan: ${permit.purpose}",
            style: const TextStyle(fontSize: 12, color: Color(0xFF334155), fontWeight: FontWeight.w500),
          ),

          const SizedBox(height: 8),

          Container(
            padding: const EdgeInsets.symmetric(horizontal: 10, vertical: 6),
            decoration: BoxDecoration(color: const Color(0xFFF8FAFC), borderRadius: BorderRadius.circular(8)),
            child: Row(
              mainAxisAlignment: MainAxisAlignment.spaceBetween,
              children: [
                Text("Mulai: $startTimeStr WIT", style: const TextStyle(fontSize: 11, fontWeight: FontWeight.w600, color: Color(0xFF1E293B))),
                Text(
                  permit.returnStatusDisplay,
                  style: TextStyle(
                    fontSize: 11,
                    fontWeight: FontWeight.w600,
                    color: permit.isExpired
                        ? const Color(0xFFDC2626)
                        : (isActive ? const Color(0xFFB45309) : const Color(0xFF1E293B)),
                  ),
                ),
                Text(permit.durationString, style: const TextStyle(fontSize: 11, fontWeight: FontWeight.bold, color: Color(0xFF475569))),
              ],
            ),
          ),
        ],
      ),
    );
  }

  Widget _buildCepuCard(CepuModel cepu) {
    final isValid = cepu.isValid;
    final isVerifiedByMe = cepu.isVerifiedByUser(widget.user.uid);
    final isReporterOrTarget = widget.user.uid == cepu.reporterUid || widget.user.uid == cepu.targetUid;
    final startTimeStr = DateFormat('HH:mm').format(cepu.startTime);
    final endTimeStr = cepu.endTime != null ? DateFormat('HH:mm').format(cepu.endTime!) : '';

    return Material(
      color: Colors.transparent,
      child: InkWell(
        onTap: () => _showVerificatorsModal(cepu),
        borderRadius: BorderRadius.circular(18),
        child: Container(
          decoration: BoxDecoration(
            color: Colors.white,
            borderRadius: BorderRadius.circular(18),
            border: Border.all(
              color: isValid ? const Color(0xFF10B981) : const Color(0xFFFDBA74),
              width: 1.5,
            ),
            boxShadow: [
              BoxShadow(
                color: (isValid ? const Color(0xFF10B981) : const Color(0xFFEA580C)).withOpacity(0.08),
                blurRadius: 10,
                offset: const Offset(0, 4),
              ),
            ],
          ),
          padding: const EdgeInsets.all(15),
          child: Column(
            crossAxisAlignment: CrossAxisAlignment.start,
            children: [
              // Header: Terlapor & Status Badge
              Row(
                children: [
                  Container(
                    padding: const EdgeInsets.all(8),
                    decoration: BoxDecoration(
                      color: isValid ? const Color(0xFFECFDF5) : const Color(0xFFFFF7ED),
                      shape: BoxShape.circle,
                    ),
                    child: Icon(
                      isValid ? Icons.verified_user_rounded : Icons.warning_amber_rounded,
                      color: isValid ? const Color(0xFF10B981) : const Color(0xFFEA580C),
                      size: 20,
                    ),
                  ),
                  const SizedBox(width: 10),
                  Expanded(
                    child: Column(
                      crossAxisAlignment: CrossAxisAlignment.start,
                      children: [
                        Row(
                          children: [
                            const Text(
                              "Terlapor: ",
                              style: TextStyle(fontSize: 11, color: Color(0xFF64748B), fontWeight: FontWeight.w500),
                            ),
                            Flexible(
                              child: Text(
                                cepu.targetName,
                                style: const TextStyle(fontSize: 13.5, fontWeight: FontWeight.bold, color: Color(0xFF1E293B)),
                                overflow: TextOverflow.ellipsis,
                              ),
                            ),
                          ],
                        ),
                        Text(
                          "Pelapor: ${cepu.reporterName}",
                          style: const TextStyle(fontSize: 10.5, color: Color(0xFF94A3B8)),
                          maxLines: 1,
                          overflow: TextOverflow.ellipsis,
                        ),
                      ],
                    ),
                  ),
                  // Badge Validitas (Klik untuk lihat verifikator)
                  InkWell(
                    onTap: () => _showVerificatorsModal(cepu),
                    borderRadius: BorderRadius.circular(10),
                    child: Container(
                      padding: const EdgeInsets.symmetric(horizontal: 8, vertical: 4),
                      decoration: BoxDecoration(
                        color: isValid ? const Color(0xFF10B981) : const Color(0xFFF59E0B),
                        borderRadius: BorderRadius.circular(10),
                      ),
                      child: Row(
                        mainAxisSize: MainAxisSize.min,
                        children: [
                          Text(
                            isValid ? "VALID (${cepu.verificationCount} Rekan)" : "${cepu.verificationCount}/4 Verifikasi",
                            style: const TextStyle(color: Colors.white, fontSize: 10, fontWeight: FontWeight.bold),
                          ),
                          const SizedBox(width: 3),
                          const Icon(Icons.info_outline_rounded, color: Colors.white, size: 12),
                        ],
                      ),
                    ),
                  ),
                ],
              ),

              const SizedBox(height: 12),
              const Divider(height: 1, color: Color(0xFFF1F5F9)),
              const SizedBox(height: 10),

              // Keterangan
              Row(
                crossAxisAlignment: CrossAxisAlignment.start,
                children: [
                  const Icon(Icons.chat_bubble_outline_rounded, size: 14, color: Color(0xFF64748B)),
                  const SizedBox(width: 6),
                  Expanded(
                    child: RichText(
                      text: TextSpan(
                        style: const TextStyle(fontSize: 12, color: Color(0xFF334155)),
                        children: [
                          const TextSpan(text: "Keterangan: ", style: TextStyle(fontWeight: FontWeight.bold)),
                          TextSpan(text: cepu.description),
                        ],
                      ),
                    ),
                  ),
                ],
              ),

              const SizedBox(height: 8),

              // Waktu Mulai & Kembali (Poin 9: jika belum kembali tulis "Belum Kembali")
              Container(
                padding: const EdgeInsets.symmetric(horizontal: 10, vertical: 6),
                decoration: BoxDecoration(color: const Color(0xFFF8FAFC), borderRadius: BorderRadius.circular(8)),
                child: Row(
                  mainAxisAlignment: MainAxisAlignment.spaceBetween,
                  children: [
                    Text("Mulai: $startTimeStr WIT", style: const TextStyle(fontSize: 11, fontWeight: FontWeight.w600, color: Color(0xFF1E293B))),
                    Text(
                      cepu.returnStatusDisplay,
                      style: TextStyle(
                        fontSize: 11,
                        fontWeight: FontWeight.w600,
                        color: cepu.isExpired
                            ? const Color(0xFFDC2626)
                            : (cepu.isActive ? const Color(0xFFEA580C) : const Color(0xFF1E293B)),
                      ),
                    ),
                  ],
                ),
              ),

              // Dokumen Pendukung / Foto Bukti jika ada
              if (cepu.photoBase64 != null) ...[
                const SizedBox(height: 10),
                InkWell(
                  onTap: () => _showImagePreview(cepu.photoBase64!),
                  borderRadius: BorderRadius.circular(10),
                  child: Container(
                    padding: const EdgeInsets.symmetric(horizontal: 10, vertical: 6),
                    decoration: BoxDecoration(
                      color: const Color(0xFFEFF6FF),
                      borderRadius: BorderRadius.circular(10),
                      border: Border.all(color: const Color(0xFFBFDBFE)),
                    ),
                    child: Row(
                      mainAxisSize: MainAxisSize.min,
                      children: const [
                        Icon(Icons.image_rounded, size: 16, color: Color(0xFF2563EB)),
                        SizedBox(width: 6),
                        Text(
                          "Lihat Foto Bukti Pendukung",
                          style: TextStyle(fontSize: 11, fontWeight: FontWeight.bold, color: Color(0xFF1E40AF)),
                        ),
                        SizedBox(width: 4),
                        Icon(Icons.zoom_in_rounded, size: 14, color: Color(0xFF2563EB)),
                      ],
                    ),
                  ),
                ),
              ],

              const SizedBox(height: 12),

              // Progress Verifikasi Bar & Tap to see Verifikators (Poin 10)
              InkWell(
                onTap: () => _showVerificatorsModal(cepu),
                borderRadius: BorderRadius.circular(8),
                child: Padding(
                  padding: const EdgeInsets.symmetric(vertical: 2),
                  child: Column(
                    crossAxisAlignment: CrossAxisAlignment.start,
                    children: [
                      Row(
                        mainAxisAlignment: MainAxisAlignment.spaceBetween,
                        children: [
                          Text(
                            "Verifikator (${cepu.verificationCount}) • Ketuk untuk detail",
                            style: const TextStyle(fontSize: 11, fontWeight: FontWeight.bold, color: Color(0xFF2563EB)),
                          ),
                          if (cepu.verifiedByNames.isNotEmpty)
                            Text(
                              cepu.verifiedByNames.join(", "),
                              style: const TextStyle(fontSize: 10, color: Color(0xFF64748B), fontStyle: FontStyle.italic),
                              maxLines: 1,
                              overflow: TextOverflow.ellipsis,
                            )
                          else
                            const Text(
                              "Belum ada verifikator",
                              style: TextStyle(fontSize: 10, color: Color(0xFF94A3B8), fontStyle: FontStyle.italic),
                            ),
                        ],
                      ),
                      const SizedBox(height: 4),
                      ClipRRect(
                        borderRadius: BorderRadius.circular(6),
                        child: LinearProgressIndicator(
                          value: (cepu.verificationCount / 4.0).clamp(0.0, 1.0),
                          backgroundColor: const Color(0xFFE2E8F0),
                          valueColor: AlwaysStoppedAnimation<Color>(
                            isValid ? const Color(0xFF10B981) : const Color(0xFFF59E0B),
                          ),
                          minHeight: 6,
                        ),
                      ),
                    ],
                  ),
                ),
              ),

              const SizedBox(height: 12),

              // Tombol Aksi Verifikasi (Poin 6: Bisa verifikasi bahkan setelah 4 verifikator)
              if (isVerifiedByMe)
                Container(
                  width: double.infinity,
                  padding: const EdgeInsets.symmetric(vertical: 8),
                  decoration: BoxDecoration(
                    color: const Color(0xFFECFDF5),
                    borderRadius: BorderRadius.circular(10),
                    border: Border.all(color: const Color(0xFFA7F3D0)),
                  ),
                  child: Row(
                    mainAxisAlignment: MainAxisAlignment.center,
                    children: const [
                      Icon(Icons.check_circle_rounded, color: Color(0xFF059669), size: 16),
                      SizedBox(width: 6),
                      Text(
                        "Anda sudah memverifikasi laporan ini",
                        style: TextStyle(fontSize: 11.5, fontWeight: FontWeight.bold, color: Color(0xFF065F46)),
                      ),
                    ],
                  ),
                )
              else if (isReporterOrTarget)
                const SizedBox.shrink()
              else
                SizedBox(
                  width: double.infinity,
                  height: 38,
                  child: ElevatedButton.icon(
                    onPressed: () => _showVerificationDialog(cepu),
                    icon: const Icon(Icons.how_to_reg_rounded, size: 16, color: Colors.white),
                    label: const Text(
                      "Verifikasi Laporan",
                      style: TextStyle(fontSize: 12, fontWeight: FontWeight.bold, color: Colors.white),
                    ),
                    style: ElevatedButton.styleFrom(
                      backgroundColor: const Color(0xFFEA580C),
                      elevation: 1,
                      shape: RoundedRectangleBorder(borderRadius: BorderRadius.circular(10)),
                    ),
                  ),
                ),
            ],
          ),
        ),
      ),
    );
  }
}
