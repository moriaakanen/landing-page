import 'package:flutter/material.dart';
import 'package:intl/intl.dart';
import '../models/permit_model.dart';
import '../models/user_model.dart';
import '../core/services/permit_service.dart';

class DailyMonitoringView extends StatefulWidget {
  final UserModel user;

  const DailyMonitoringView({Key? key, required this.user}) : super(key: key);

  @override
  State<DailyMonitoringView> createState() => _DailyMonitoringViewState();
}

class _DailyMonitoringViewState extends State<DailyMonitoringView> {
  final PermitService _permitService = PermitService();
  final TextEditingController _searchController = TextEditingController();

  DateTime _selectedDate = DateTime.now();
  String _selectedFilter = 'ALL'; // 'ALL', 'ACTIVE', 'COMPLETED'
  String _searchQuery = '';

  String get _selectedDateString => DateFormat('yyyy-MM-dd').format(_selectedDate);

  @override
  void dispose() {
    _searchController.dispose();
    super.dispose();
  }

  void _previousDay() {
    setState(() {
      _selectedDate = _selectedDate.subtract(const Duration(days: 1));
    });
  }

  void _nextDay() {
    setState(() {
      _selectedDate = _selectedDate.add(const Duration(days: 1));
    });
  }

  void _goToToday() {
    setState(() {
      _selectedDate = DateTime.now();
    });
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
      setState(() {
        _selectedDate = picked;
      });
    }
  }

  bool _isToday(DateTime date) {
    final now = DateTime.now();
    return date.year == now.year && date.month == now.month && date.day == now.day;
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
          "Monitoring Izin Harian",
          style: TextStyle(
            color: Colors.white,
            fontWeight: FontWeight.bold,
            fontSize: 18,
          ),
        ),
        centerTitle: true,
        actions: [
          if (!isCurrentDay)
            TextButton(
              onPressed: _goToToday,
              child: const Text(
                "Hari Ini",
                style: TextStyle(
                  color: Colors.white,
                  fontWeight: FontWeight.bold,
                  fontSize: 13,
                ),
              ),
            ),
        ],
      ),
      body: Column(
        children: [
          // Date Selector Header Bar
          Container(
            padding: const EdgeInsets.symmetric(horizontal: 14, vertical: 10),
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
                          const Icon(Icons.calendar_today_rounded, size: 16, color: Color(0xFF1E60F2)),
                          const SizedBox(width: 8),
                          Flexible(
                            child: Text(
                              formattedDate,
                              style: const TextStyle(
                                fontSize: 13,
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

          // Main Stream Area
          Expanded(
            child: StreamBuilder<List<PermitModel>>(
              stream: _permitService.streamDailyPermits(_selectedDateString),
              builder: (context, snapshot) {
                if (snapshot.connectionState == ConnectionState.waiting) {
                  return const Center(
                    child: CircularProgressIndicator(color: Color(0xFF1E60F2)),
                  );
                }

                if (snapshot.hasError) {
                  return Center(
                    child: Padding(
                      padding: const EdgeInsets.all(24.0),
                      child: Column(
                        mainAxisAlignment: MainAxisAlignment.center,
                        children: [
                          const Icon(Icons.error_outline_rounded, color: Color(0xFFEF4444), size: 48),
                          const SizedBox(height: 12),
                          Text(
                            "Gagal memuat data izin:\n${snapshot.error}",
                            textAlign: TextAlign.center,
                            style: const TextStyle(color: Color(0xFF64748B), fontSize: 13),
                          ),
                        ],
                      ),
                    ),
                  );
                }

                final allPermits = snapshot.data ?? [];

                // Filter Search & Tab
                final filteredPermits = allPermits.where((permit) {
                  final matchesSearch = _searchQuery.isEmpty ||
                      permit.userName.toLowerCase().contains(_searchQuery.toLowerCase()) ||
                      permit.userDepartment.toLowerCase().contains(_searchQuery.toLowerCase()) ||
                      permit.purpose.toLowerCase().contains(_searchQuery.toLowerCase());

                  if (!matchesSearch) return false;

                  if (_selectedFilter == 'ACTIVE') {
                    return permit.isActive;
                  } else if (_selectedFilter == 'COMPLETED') {
                    return !permit.isActive;
                  }
                  return true;
                }).toList();

                final activeCount = allPermits.where((p) => p.isActive).length;
                final completedCount = allPermits.where((p) => !p.isActive).length;

                return Column(
                  children: [
                    // Summary Metric Cards
                    Padding(
                      padding: const EdgeInsets.fromLTRB(16, 14, 16, 8),
                      child: Row(
                        children: [
                          _buildSummaryCard(
                            label: "Total Izin",
                            count: allPermits.length.toString(),
                            color: const Color(0xFF3B82F6),
                            icon: Icons.list_alt_rounded,
                          ),
                          const SizedBox(width: 10),
                          _buildSummaryCard(
                            label: "Sedang di Luar",
                            count: activeCount.toString(),
                            color: const Color(0xFFF59E0B),
                            icon: Icons.directions_walk_rounded,
                          ),
                          const SizedBox(width: 10),
                          _buildSummaryCard(
                            label: "Sudah Kembali",
                            count: completedCount.toString(),
                            color: const Color(0xFF10B981),
                            icon: Icons.check_circle_outline_rounded,
                          ),
                        ],
                      ),
                    ),

                    // Search & Filter Pills
                    Padding(
                      padding: const EdgeInsets.symmetric(horizontal: 16, vertical: 8),
                      child: Column(
                        children: [
                          // Search Box
                          Container(
                            height: 42,
                            decoration: BoxDecoration(
                              color: Colors.white,
                              borderRadius: BorderRadius.circular(12),
                              border: Border.all(color: const Color(0xFFE2E8F0)),
                            ),
                            child: TextField(
                              controller: _searchController,
                              onChanged: (val) {
                                setState(() {
                                  _searchQuery = val.trim();
                                });
                              },
                              decoration: InputDecoration(
                                hintText: "Cari nama pegawai atau keperluan...",
                                hintStyle: TextStyle(fontSize: 12, color: Colors.grey[500]),
                                prefixIcon: const Icon(Icons.search, size: 20, color: Color(0xFF64748B)),
                                suffixIcon: _searchQuery.isNotEmpty
                                    ? IconButton(
                                        icon: const Icon(Icons.clear, size: 18),
                                        onPressed: () {
                                          _searchController.clear();
                                          setState(() => _searchQuery = '');
                                        },
                                      )
                                    : null,
                                border: InputBorder.none,
                                contentPadding: const EdgeInsets.symmetric(vertical: 10),
                              ),
                            ),
                          ),

                          const SizedBox(height: 10),

                          // Filter Pills
                          Row(
                            children: [
                              _buildFilterPill("Semua (${allPermits.length})", 'ALL'),
                              const SizedBox(width: 8),
                              _buildFilterPill("Sedang Izin ($activeCount)", 'ACTIVE'),
                              const SizedBox(width: 8),
                              _buildFilterPill("Kembali ($completedCount)", 'COMPLETED'),
                            ],
                          ),
                        ],
                      ),
                    ),

                    // List of Permits
                    Expanded(
                      child: filteredPermits.isEmpty
                          ? Center(
                              child: Column(
                                mainAxisAlignment: MainAxisAlignment.center,
                                children: [
                                  Container(
                                    padding: const EdgeInsets.all(20),
                                    decoration: BoxDecoration(
                                      color: const Color(0xFFEFF6FF),
                                      shape: BoxShape.circle,
                                    ),
                                    child: const Icon(
                                      Icons.assignment_turned_in_outlined,
                                      size: 48,
                                      color: Color(0xFF3B82F6),
                                    ),
                                  ),
                                  const SizedBox(height: 14),
                                  Text(
                                    allPermits.isEmpty
                                        ? "Tidak ada catatan izin pada tanggal ini."
                                        : "Tidak ada data izin yang cocok dengan filter.",
                                    style: const TextStyle(
                                      fontSize: 14,
                                      fontWeight: FontWeight.w600,
                                      color: Color(0xFF475569),
                                    ),
                                  ),
                                  const SizedBox(height: 6),
                                  Text(
                                    allPermits.isEmpty
                                        ? "Pegawai yang memiliki kegiatan di jam kantor akan tercatat di sini."
                                        : "Coba ubah kata kunci pencarian atau filter status.",
                                    style: TextStyle(
                                      fontSize: 12,
                                      color: Colors.grey[500],
                                    ),
                                    textAlign: TextAlign.center,
                                  ),
                                ],
                              ),
                            )
                          : ListView.separated(
                              padding: const EdgeInsets.fromLTRB(16, 6, 16, 24),
                              itemCount: filteredPermits.length,
                              separatorBuilder: (context, index) => const SizedBox(height: 12),
                              itemBuilder: (context, index) {
                                final permit = filteredPermits[index];
                                return _buildPermitCard(permit);
                              },
                            ),
                    ),
                  ],
                );
              },
            ),
          ),
        ],
      ),
    );
  }

  Widget _buildSummaryCard({
    required String label,
    required String count,
    required Color color,
    required IconData icon,
  }) {
    return Expanded(
      child: Container(
        padding: const EdgeInsets.symmetric(vertical: 12, horizontal: 10),
        decoration: BoxDecoration(
          color: Colors.white,
          borderRadius: BorderRadius.circular(14),
          border: Border.all(color: color.withOpacity(0.25)),
          boxShadow: [
            BoxShadow(
              color: color.withOpacity(0.06),
              blurRadius: 8,
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
                    fontSize: 18,
                    fontWeight: FontWeight.bold,
                    color: color,
                  ),
                ),
              ],
            ),
            const SizedBox(height: 4),
            Text(
              label,
              style: const TextStyle(
                fontSize: 10,
                color: Color(0xFF64748B),
                fontWeight: FontWeight.w600,
              ),
              textAlign: TextAlign.center,
              maxLines: 1,
              overflow: TextOverflow.ellipsis,
            ),
          ],
        ),
      ),
    );
  }

  Widget _buildFilterPill(String title, String filterKey) {
    final isSelected = _selectedFilter == filterKey;
    return InkWell(
      onTap: () {
        setState(() => _selectedFilter = filterKey);
      },
      borderRadius: BorderRadius.circular(20),
      child: Container(
        padding: const EdgeInsets.symmetric(horizontal: 12, vertical: 6),
        decoration: BoxDecoration(
          color: isSelected ? const Color(0xFF1E60F2) : Colors.white,
          borderRadius: BorderRadius.circular(20),
          border: Border.all(
            color: isSelected ? const Color(0xFF1E60F2) : const Color(0xFFCBD5E1),
          ),
        ),
        child: Text(
          title,
          style: TextStyle(
            fontSize: 11,
            fontWeight: isSelected ? FontWeight.bold : FontWeight.w500,
            color: isSelected ? Colors.white : const Color(0xFF475569),
          ),
        ),
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
            color: isActive
                ? const Color(0xFFF59E0B).withOpacity(0.08)
                : Colors.black.withOpacity(0.04),
            blurRadius: 10,
            offset: const Offset(0, 4),
          ),
        ],
      ),
      padding: const EdgeInsets.all(14),
      child: Column(
        crossAxisAlignment: CrossAxisAlignment.start,
        children: [
          // Row Header: User Name, Dept & Status Badge
          Row(
            crossAxisAlignment: CrossAxisAlignment.center,
            children: [
              CircleAvatar(
                radius: 18,
                backgroundColor: isActive ? const Color(0xFFFEF3C7) : const Color(0xFFEFF6FF),
                child: Text(
                  permit.userName.isNotEmpty ? permit.userName[0].toUpperCase() : 'U',
                  style: TextStyle(
                    color: isActive ? const Color(0xFFB45309) : const Color(0xFF1E60F2),
                    fontWeight: FontWeight.bold,
                    fontSize: 14,
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
                      style: const TextStyle(
                        fontSize: 14,
                        fontWeight: FontWeight.bold,
                        color: Color(0xFF1E293B),
                      ),
                      maxLines: 1,
                      overflow: TextOverflow.ellipsis,
                    ),
                    Text(
                      permit.userDepartment,
                      style: const TextStyle(
                        fontSize: 11,
                        color: Color(0xFF64748B),
                        fontWeight: FontWeight.w500,
                      ),
                    ),
                  ],
                ),
              ),
              // Status Badge
              Container(
                padding: const EdgeInsets.symmetric(horizontal: 10, vertical: 4),
                decoration: BoxDecoration(
                  color: isActive ? const Color(0xFFFFFBEB) : const Color(0xFFECFDF5),
                  borderRadius: BorderRadius.circular(10),
                  border: Border.all(
                    color: isActive ? const Color(0xFFFDE68A) : const Color(0xFFA7F3D0),
                  ),
                ),
                child: Row(
                  mainAxisSize: MainAxisSize.min,
                  children: [
                    Container(
                      width: 7,
                      height: 7,
                      decoration: BoxDecoration(
                        color: isActive ? const Color(0xFFF59E0B) : const Color(0xFF10B981),
                        shape: BoxShape.circle,
                      ),
                    ),
                    const SizedBox(width: 6),
                    Text(
                      isActive ? "Sedang Izin" : "Sudah Kembali",
                      style: TextStyle(
                        fontSize: 11,
                        fontWeight: FontWeight.bold,
                        color: isActive ? const Color(0xFFB45309) : const Color(0xFF047857),
                      ),
                    ),
                  ],
                ),
              ),
            ],
          ),

          const SizedBox(height: 12),
          const Divider(height: 1, color: Color(0xFFF1F5F9)),
          const SizedBox(height: 10),

          // Purpose / Keperluan
          Row(
            crossAxisAlignment: CrossAxisAlignment.start,
            children: [
              const Icon(Icons.assignment_outlined, size: 16, color: Color(0xFF64748B)),
              const SizedBox(width: 6),
              Expanded(
                child: RichText(
                  text: TextSpan(
                    style: const TextStyle(fontSize: 12, color: Color(0xFF334155)),
                    children: [
                      const TextSpan(
                        text: "Keperluan: ",
                        style: TextStyle(fontWeight: FontWeight.bold),
                      ),
                      TextSpan(text: permit.purpose),
                    ],
                  ),
                ),
              ),
            ],
          ),

          const SizedBox(height: 8),

          // Time Grid: Mulai, Selesai, Total Durasi
          Container(
            padding: const EdgeInsets.symmetric(horizontal: 12, vertical: 8),
            decoration: BoxDecoration(
              color: const Color(0xFFF8FAFC),
              borderRadius: BorderRadius.circular(10),
            ),
            child: Row(
              mainAxisAlignment: MainAxisAlignment.spaceBetween,
              children: [
                Row(
                  children: [
                    const Icon(Icons.play_circle_fill_rounded, size: 14, color: Color(0xFF3B82F6)),
                    const SizedBox(width: 4),
                    Text(
                      "Mulai: $startTimeStr",
                      style: const TextStyle(fontSize: 11, fontWeight: FontWeight.w600, color: Color(0xFF1E293B)),
                    ),
                  ],
                ),
                Row(
                  children: [
                    Icon(
                      isActive ? Icons.pending_rounded : Icons.stop_circle_rounded,
                      size: 14,
                      color: isActive ? const Color(0xFFF59E0B) : const Color(0xFF10B981),
                    ),
                    const SizedBox(width: 4),
                    Text(
                      isActive ? "Belum Kembali" : "Kembali: $endTimeStr",
                      style: TextStyle(
                        fontSize: 11,
                        fontWeight: FontWeight.w600,
                        color: isActive ? const Color(0xFFB45309) : const Color(0xFF1E293B),
                      ),
                    ),
                  ],
                ),
                Row(
                  children: [
                    const Icon(Icons.timelapse_rounded, size: 14, color: Color(0xFF64748B)),
                    const SizedBox(width: 4),
                    Text(
                      permit.durationString,
                      style: const TextStyle(fontSize: 11, fontWeight: FontWeight.bold, color: Color(0xFF475569)),
                    ),
                  ],
                ),
              ],
            ),
          ),
        ],
      ),
    );
  }
}
