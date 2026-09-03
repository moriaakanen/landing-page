import 'dart:math';
import 'package:flutter/material.dart';
import 'package:intl/intl.dart';
import '../models/user_model.dart';
import '../core/services/permit_service.dart';
import '../core/services/cepu_service.dart';

/// Item aktivitas selesai untuk visualisasi chart "Ngapain Aja?"
class AnalyticsActivityItem {
  final String id;
  final String type; // 'PERMIT' atau 'CEPU'
  final String title;
  final String description;
  final DateTime startTime;
  final DateTime endTime;
  final int durationMinutes;
  final Color color;

  AnalyticsActivityItem({
    required this.id,
    required this.type,
    required this.title,
    required this.description,
    required this.startTime,
    required this.endTime,
    required this.durationMinutes,
    required this.color,
  });
}

class AnalyticsView extends StatefulWidget {
  final UserModel user;

  const AnalyticsView({Key? key, required this.user}) : super(key: key);

  @override
  State<AnalyticsView> createState() => _AnalyticsViewState();
}

class _AnalyticsViewState extends State<AnalyticsView> {
  final PermitService _permitService = PermitService();
  final CepuService _cepuService = CepuService();

  // Period: 'HARIAN', 'BULANAN', 'TAHUNAN'
  String _selectedPeriod = 'HARIAN';

  DateTime _selectedDate = DateTime.now();
  bool _isLoading = true;

  // Metrics Harian
  int _todayOutOfOfficeMinutes = 0;
  int _yesterdayOutOfOfficeMinutes = 0;
  int _validCepuCount = 0;
  int _totalCepuCount = 0;
  List<AnalyticsActivityItem> _activities = [];

  // Curated Palette Colors untuk Stacked Bar Chart
  final List<Color> _chartColors = const [
    Color(0xFF1E60F2), // Royal Azure
    Color(0xFFEA580C), // Sunset Orange
    Color(0xFF0D9488), // Teal Emerald
    Color(0xFF6366F1), // Royal Indigo
    Color(0xFFE11D48), // Rose
    Color(0xFFD97706), // Amber
    Color(0xFF8B5CF6), // Violet
    Color(0xFF06B6D4), // Cyan
  ];

  @override
  void initState() {
    super.initState();
    _loadAnalyticsData();
  }

  String _formatDateString(DateTime dt) => DateFormat('yyyy-MM-dd').format(dt);

  Future<void> _loadAnalyticsData() async {
    setState(() => _isLoading = true);

    final selectedDateStr = _formatDateString(_selectedDate);
    final yesterdayDate = _selectedDate.subtract(const Duration(days: 1));
    final yesterdayDateStr = _formatDateString(yesterdayDate);

    // 1. Data Izin Hari Terpilih
    final todayPermits = await _permitService.getUserDailyPermits(widget.user.uid, selectedDateStr);
    final completedPermits = todayPermits.where((p) => p.endTime != null).toList();

    // 2. Data Cepu Valid Hari Terpilih
    final todayValidCepu = await _cepuService.getDailyValidCepuForTarget(widget.user.uid, selectedDateStr);
    final completedValidCepu = todayValidCepu.where((c) => c.endTime != null).toList();

    // 3. Data Seluruh Laporan Cepu Hari Terpilih (untuk rasio valid / total)
    final allTodayCepu = await _cepuService.getAllDailyCepuForTarget(widget.user.uid, selectedDateStr);

    // 4. Data Kemarin untuk Komparasi
    final yesterdayPermits = await _permitService.getUserDailyPermits(widget.user.uid, yesterdayDateStr);
    final completedYesterdayPermits = yesterdayPermits.where((p) => p.endTime != null).toList();
    final yesterdayValidCepu = await _cepuService.getDailyValidCepuForTarget(widget.user.uid, yesterdayDateStr);
    final completedYesterdayCepu = yesterdayValidCepu.where((c) => c.endTime != null).toList();

    // Hitung total menit hari ini
    int todayMinutes = 0;
    final List<AnalyticsActivityItem> items = [];
    int colorIdx = 0;

    for (final p in completedPermits) {
      final mins = max(0, p.endTime!.difference(p.startTime).inMinutes);
      todayMinutes += mins;
      items.add(AnalyticsActivityItem(
        id: p.id,
        type: 'PERMIT',
        title: 'Izin Waigama',
        description: p.purpose,
        startTime: p.startTime,
        endTime: p.endTime!,
        durationMinutes: mins,
        color: _chartColors[colorIdx % _chartColors.length],
      ));
      colorIdx++;
    }

    for (final c in completedValidCepu) {
      final mins = max(0, c.endTime!.difference(c.startTime).inMinutes);
      todayMinutes += mins;
      items.add(AnalyticsActivityItem(
        id: c.id,
        type: 'CEPU',
        title: 'Laporan Cepu',
        description: c.description,
        startTime: c.startTime,
        endTime: c.endTime!,
        durationMinutes: mins,
        color: _chartColors[colorIdx % _chartColors.length],
      ));
      colorIdx++;
    }

    // Urutkan berdasarkan yang terlama (durasi terbesar di atas)
    items.sort((a, b) => b.durationMinutes.compareTo(a.durationMinutes));

    // Hitung total menit kemarin
    int yesterdayMinutes = 0;
    for (final p in completedYesterdayPermits) {
      yesterdayMinutes += max(0, p.endTime!.difference(p.startTime).inMinutes);
    }
    for (final c in completedYesterdayCepu) {
      yesterdayMinutes += max(0, c.endTime!.difference(c.startTime).inMinutes);
    }

    if (mounted) {
      setState(() {
        _todayOutOfOfficeMinutes = todayMinutes;
        _yesterdayOutOfOfficeMinutes = yesterdayMinutes;
        _totalCepuCount = allTodayCepu.length;
        _validCepuCount = allTodayCepu.where((c) => c.isValid).length;
        _activities = items;
        _isLoading = false;
      });
    }
  }

  void _previousDay() {
    setState(() => _selectedDate = _selectedDate.subtract(const Duration(days: 1)));
    _loadAnalyticsData();
  }

  void _nextDay() {
    setState(() => _selectedDate = _selectedDate.add(const Duration(days: 1)));
    _loadAnalyticsData();
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
              onSurface: Color(0xFF0F172A),
            ),
          ),
          child: child!,
        );
      },
    );

    if (picked != null && picked != _selectedDate) {
      setState(() => _selectedDate = picked);
      _loadAnalyticsData();
    }
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
          onPressed: () => Navigator.pop(context),
        ),
        title: const Text(
          "Analytics",
          style: TextStyle(
            color: Color(0xFF0F172A),
            fontSize: 18,
            fontWeight: FontWeight.bold,
          ),
        ),
        centerTitle: true,
        actions: [
          IconButton(
            icon: const Icon(Icons.refresh_rounded, color: Color(0xFF64748B)),
            onPressed: _loadAnalyticsData,
          ),
        ],
      ),
      body: SafeArea(
        child: RefreshIndicator(
          onRefresh: _loadAnalyticsData,
          color: const Color(0xFF1E60F2),
          child: SingleChildScrollView(
            physics: const AlwaysScrollableScrollPhysics(),
            padding: const EdgeInsets.fromLTRB(20, 16, 20, 40),
            child: Column(
              crossAxisAlignment: CrossAxisAlignment.start,
              children: [
                // =========================================================================
                // 1. PERIOD SELECTOR (Harian, Bulanan, Tahunan)
                // =========================================================================
                Container(
                  padding: const EdgeInsets.all(4),
                  decoration: BoxDecoration(
                    color: Colors.white,
                    borderRadius: BorderRadius.circular(16),
                    border: Border.all(color: const Color(0xFFE2E8F0)),
                  ),
                  child: Row(
                    children: [
                      _buildPeriodTab("Harian", "HARIAN"),
                      _buildPeriodTab("Bulanan", "BULANAN"),
                      _buildPeriodTab("Tahunan", "TAHUNAN"),
                    ],
                  ),
                ),

                const SizedBox(height: 16),

                // =========================================================================
                // 2. DATE / MONTH / YEAR SELECTOR BAR
                // =========================================================================
                _buildDatePickerBar(),

                const SizedBox(height: 20),

                // =========================================================================
                // 3. KONTEN PERIODE (Harian vs Bulanan/Tahunan)
                // =========================================================================
                if (_selectedPeriod != 'HARIAN')
                  _buildComingSoonPlaceholder()
                else if (_isLoading)
                  const Padding(
                    padding: EdgeInsets.symmetric(vertical: 60),
                    child: Center(child: CircularProgressIndicator(color: Color(0xFF1E60F2))),
                  )
                else ...[
                  // 2 Summary Cards: "Lama tidak dikantor" & "Berapa kali dilaporkan"
                  _buildSummaryCards(),

                  const SizedBox(height: 20),

                  // Stacked Horizontal Bar Chart Card: "Ngapain Aja?"
                  _buildNgapainAjaCard(),
                ],
              ],
            ),
          ),
        ),
      ),
    );
  }

  Widget _buildPeriodTab(String label, String value) {
    final isSelected = _selectedPeriod == value;
    return Expanded(
      child: InkWell(
        onTap: () {
          setState(() => _selectedPeriod = value);
        },
        borderRadius: BorderRadius.circular(12),
        child: AnimatedContainer(
          duration: const Duration(milliseconds: 200),
          padding: const EdgeInsets.symmetric(vertical: 10),
          decoration: BoxDecoration(
            color: isSelected ? const Color(0xFF1E60F2) : Colors.transparent,
            borderRadius: BorderRadius.circular(12),
            boxShadow: isSelected
                ? [
                    BoxShadow(
                      color: const Color(0xFF1E60F2).withOpacity(0.25),
                      blurRadius: 8,
                      offset: const Offset(0, 2),
                    ),
                  ]
                : null,
          ),
          child: Text(
            label,
            textAlign: TextAlign.center,
            style: TextStyle(
              fontSize: 13,
              fontWeight: isSelected ? FontWeight.bold : FontWeight.w600,
              color: isSelected ? Colors.white : const Color(0xFF64748B),
            ),
          ),
        ),
      ),
    );
  }

  Widget _buildDatePickerBar() {
    String dateLabel;
    if (_selectedPeriod == 'HARIAN') {
      try {
        dateLabel = DateFormat('EEEE, d MMMM yyyy', 'id_ID').format(_selectedDate);
      } catch (_) {
        dateLabel = DateFormat('EEEE, d MMMM yyyy').format(_selectedDate);
      }
    } else if (_selectedPeriod == 'BULANAN') {
      try {
        dateLabel = DateFormat('MMMM yyyy', 'id_ID').format(_selectedDate);
      } catch (_) {
        dateLabel = DateFormat('MMMM yyyy').format(_selectedDate);
      }
    } else {
      dateLabel = DateFormat('yyyy').format(_selectedDate);
    }

    return Container(
      padding: const EdgeInsets.symmetric(horizontal: 8, vertical: 6),
      decoration: BoxDecoration(
        color: Colors.white,
        borderRadius: BorderRadius.circular(16),
        border: Border.all(color: const Color(0xFFE2E8F0)),
        boxShadow: [
          BoxShadow(
            color: Colors.black.withOpacity(0.02),
            blurRadius: 8,
            offset: const Offset(0, 2),
          ),
        ],
      ),
      child: Row(
        mainAxisAlignment: MainAxisAlignment.spaceBetween,
        children: [
          IconButton(
            icon: const Icon(Icons.chevron_left_rounded, color: Color(0xFF475569)),
            onPressed: _selectedPeriod == 'HARIAN' ? _previousDay : null,
          ),
          InkWell(
            onTap: _selectedPeriod == 'HARIAN' ? _pickDate : null,
            borderRadius: BorderRadius.circular(10),
            child: Padding(
              padding: const EdgeInsets.symmetric(horizontal: 10, vertical: 6),
              child: Row(
                children: [
                  const Icon(Icons.calendar_month_rounded, color: Color(0xFF1E60F2), size: 18),
                  const SizedBox(width: 8),
                  Text(
                    dateLabel,
                    style: const TextStyle(
                      fontSize: 13.5,
                      fontWeight: FontWeight.bold,
                      color: Color(0xFF0F172A),
                    ),
                  ),
                ],
              ),
            ),
          ),
          IconButton(
            icon: const Icon(Icons.chevron_right_rounded, color: Color(0xFF475569)),
            onPressed: _selectedPeriod == 'HARIAN' ? _nextDay : null,
          ),
        ],
      ),
    );
  }

  Widget _buildSummaryCards() {
    // Perbandingan dengan kemarin
    final diff = _todayOutOfOfficeMinutes - _yesterdayOutOfOfficeMinutes;
    final bool isBetter = diff <= 0; // Lebih sedikit meninggalkan kantor dianggap lebih baik
    final bool isSame = diff == 0;

    return Row(
      children: [
        // Card 1: "Lama tidak dikantor"
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
                  blurRadius: 10,
                  offset: const Offset(0, 2),
                ),
              ],
            ),
            child: Column(
              crossAxisAlignment: CrossAxisAlignment.start,
              children: [
                Row(
                  mainAxisAlignment: MainAxisAlignment.spaceBetween,
                  children: [
                    const Expanded(
                      child: Text(
                        "Lama tidak dikantor",
                        style: TextStyle(
                          fontSize: 12,
                          fontWeight: FontWeight.w600,
                          color: Color(0xFF64748B),
                        ),
                        maxLines: 1,
                        overflow: TextOverflow.ellipsis,
                      ),
                    ),
                    Container(
                      width: 28,
                      height: 28,
                      decoration: BoxDecoration(
                        color: const Color(0xFFEFF6FF),
                        borderRadius: BorderRadius.circular(8),
                      ),
                      child: const Icon(Icons.schedule_rounded, color: Color(0xFF1E60F2), size: 16),
                    ),
                  ],
                ),
                const SizedBox(height: 10),
                Text(
                  "$_todayOutOfOfficeMinutes",
                  style: const TextStyle(
                    fontSize: 26,
                    fontWeight: FontWeight.w900,
                    color: Color(0xFF0F172A),
                    letterSpacing: -0.5,
                  ),
                ),
                const Text(
                  "Menit",
                  style: TextStyle(fontSize: 11.5, color: Color(0xFF64748B), fontWeight: FontWeight.w500),
                ),
                const SizedBox(height: 8),

                // Insight Hijau / Merah Berdasarkan Perbandingan Kemarin
                Row(
                  children: [
                    Icon(
                      isSame
                          ? Icons.remove_rounded
                          : (isBetter ? Icons.trending_down_rounded : Icons.trending_up_rounded),
                      size: 14,
                      color: isSame
                          ? const Color(0xFF64748B)
                          : (isBetter ? const Color(0xFF10B981) : const Color(0xFFEF4444)),
                    ),
                    const SizedBox(width: 3),
                    Expanded(
                      child: Text(
                        isSame
                            ? "Sama seperti kemarin"
                            : (isBetter
                                ? "-${diff.abs()} mnt vs kemarin"
                                : "+${diff.abs()} mnt vs kemarin"),
                        style: TextStyle(
                          fontSize: 10.5,
                          fontWeight: FontWeight.bold,
                          color: isSame
                              ? const Color(0xFF64748B)
                              : (isBetter ? const Color(0xFF10B981) : const Color(0xFFEF4444)),
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
        ),

        const SizedBox(width: 12),

        // Card 2: "Berapa kali dilaporkan"
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
                  blurRadius: 10,
                  offset: const Offset(0, 2),
                ),
              ],
            ),
            child: Column(
              crossAxisAlignment: CrossAxisAlignment.start,
              children: [
                Row(
                  mainAxisAlignment: MainAxisAlignment.spaceBetween,
                  children: [
                    const Expanded(
                      child: Text(
                        "Berapa kali dilaporkan",
                        style: TextStyle(
                          fontSize: 12,
                          fontWeight: FontWeight.w600,
                          color: Color(0xFF64748B),
                        ),
                        maxLines: 1,
                        overflow: TextOverflow.ellipsis,
                      ),
                    ),
                    Container(
                      width: 28,
                      height: 28,
                      decoration: BoxDecoration(
                        color: const Color(0xFFFFF7ED),
                        borderRadius: BorderRadius.circular(8),
                      ),
                      child: const Icon(Icons.campaign_rounded, color: Color(0xFFEA580C), size: 16),
                    ),
                  ],
                ),
                const SizedBox(height: 10),
                RichText(
                  text: TextSpan(
                    children: [
                      TextSpan(
                        text: "$_validCepuCount ",
                        style: const TextStyle(
                          fontSize: 26,
                          fontWeight: FontWeight.w900,
                          color: Color(0xFF0F172A),
                          letterSpacing: -0.5,
                        ),
                      ),
                      TextSpan(
                        text: "/ $_totalCepuCount",
                        style: const TextStyle(
                          fontSize: 18,
                          fontWeight: FontWeight.bold,
                          color: Color(0xFF94A3B8),
                        ),
                      ),
                    ],
                  ),
                ),
                const Text(
                  "Laporan Valid / Total",
                  style: TextStyle(fontSize: 11.5, color: Color(0xFF64748B), fontWeight: FontWeight.w500),
                ),
                const SizedBox(height: 8),

                // Status verifikasi
                Row(
                  children: [
                    Icon(
                      _validCepuCount == 0 ? Icons.check_circle_outline_rounded : Icons.info_outline_rounded,
                      size: 14,
                      color: _validCepuCount == 0 ? const Color(0xFF10B981) : const Color(0xFFEA580C),
                    ),
                    const SizedBox(width: 4),
                    Expanded(
                      child: Text(
                        _validCepuCount == 0 ? "Tidak ada pelanggaran" : "$_validCepuCount laporan terverifikasi",
                        style: TextStyle(
                          fontSize: 10.5,
                          fontWeight: FontWeight.bold,
                          color: _validCepuCount == 0 ? const Color(0xFF10B981) : const Color(0xFFEA580C),
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
        ),
      ],
    );
  }

  /// Card "Ngapain Aja?" Berisi Stacked Horizontal Bar Chart & List Diurutkan Terlama
  Widget _buildNgapainAjaCard() {
    return Container(
      padding: const EdgeInsets.all(18),
      decoration: BoxDecoration(
        color: Colors.white,
        borderRadius: BorderRadius.circular(22),
        border: Border.all(color: const Color(0xFFE2E8F0)),
        boxShadow: [
          BoxShadow(
            color: Colors.black.withOpacity(0.02),
            blurRadius: 10,
            offset: const Offset(0, 2),
          ),
        ],
      ),
      child: Column(
        crossAxisAlignment: CrossAxisAlignment.start,
        children: [
          Row(
            mainAxisAlignment: MainAxisAlignment.spaceBetween,
            children: [
              Column(
                crossAxisAlignment: CrossAxisAlignment.start,
                children: [
                  const Text(
                    "Ngapain Aja?",
                    style: TextStyle(
                      fontSize: 17,
                      fontWeight: FontWeight.bold,
                      color: Color(0xFF0F172A),
                    ),
                  ),
                  const SizedBox(height: 2),
                  Text(
                    "Total di luar kantor: $_todayOutOfOfficeMinutes menit",
                    style: const TextStyle(fontSize: 12, color: Color(0xFF64748B)),
                  ),
                ],
              ),
              Container(
                padding: const EdgeInsets.symmetric(horizontal: 10, vertical: 5),
                decoration: BoxDecoration(
                  color: const Color(0xFFEFF6FF),
                  borderRadius: BorderRadius.circular(10),
                ),
                child: Row(
                  children: [
                    const Icon(Icons.sort_rounded, size: 14, color: Color(0xFF1E60F2)),
                    const SizedBox(width: 4),
                    const Text(
                      "Terlama",
                      style: TextStyle(fontSize: 11, fontWeight: FontWeight.bold, color: Color(0xFF1E60F2)),
                    ),
                  ],
                ),
              ),
            ],
          ),

          const SizedBox(height: 18),

          if (_activities.isEmpty)
            Container(
              padding: const EdgeInsets.symmetric(vertical: 32),
              alignment: Alignment.center,
              child: Column(
                children: const [
                  Icon(Icons.thumb_up_alt_outlined, color: Color(0xFF10B981), size: 36),
                  SizedBox(height: 10),
                  Text(
                    "Tidak ada izin / laporan selesai",
                    style: TextStyle(fontSize: 14, fontWeight: FontWeight.bold, color: Color(0xFF334155)),
                  ),
                  SizedBox(height: 4),
                  Text(
                    "Seluruh waktu kerja dihabiskan di kantor pada hari ini.",
                    style: TextStyle(fontSize: 12, color: Color(0xFF94A3B8)),
                  ),
                ],
              ),
            )
          else ...[
            // 1. Stacked Horizontal Bar Chart
            ClipRRect(
              borderRadius: BorderRadius.circular(8),
              child: SizedBox(
                height: 18,
                child: Row(
                  children: _activities.map((item) {
                    final double flexWeight = _todayOutOfOfficeMinutes > 0
                        ? item.durationMinutes / _todayOutOfOfficeMinutes
                        : 1.0;
                    return Expanded(
                      flex: max(1, (flexWeight * 1000).toInt()),
                      child: Container(
                        color: item.color,
                        margin: const EdgeInsets.symmetric(horizontal: 0.5),
                      ),
                    );
                  }).toList(),
                ),
              ),
            ),

            const SizedBox(height: 18),

            // 2. Daftar Kegiatan Diurutkan Berdasarkan Yang Terlama
            ListView.separated(
              shrinkWrap: true,
              physics: const NeverScrollableScrollPhysics(),
              itemCount: _activities.length,
              separatorBuilder: (_, __) => const Divider(height: 16, color: Color(0xFFF1F5F9)),
              itemBuilder: (context, index) {
                final item = _activities[index];
                final percent = _todayOutOfOfficeMinutes > 0
                    ? ((item.durationMinutes / _todayOutOfOfficeMinutes) * 100).toStringAsFixed(0)
                    : "0";
                final timeRange =
                    "${DateFormat('HH:mm').format(item.startTime)} - ${DateFormat('HH:mm').format(item.endTime)} WIT";

                return Row(
                  children: [
                    // Dot Penanda Warna Chart
                    Container(
                      width: 10,
                      height: 10,
                      decoration: BoxDecoration(
                        color: item.color,
                        shape: BoxShape.circle,
                      ),
                    ),
                    const SizedBox(width: 10),

                    // Keterangan / Keperluan
                    Expanded(
                      child: Column(
                        crossAxisAlignment: CrossAxisAlignment.start,
                        children: [
                          Text(
                            item.description.isNotEmpty ? item.description : item.title,
                            style: const TextStyle(
                              fontSize: 13,
                              fontWeight: FontWeight.bold,
                              color: Color(0xFF0F172A),
                            ),
                            maxLines: 1,
                            overflow: TextOverflow.ellipsis,
                          ),
                          const SizedBox(height: 2),
                          Row(
                            children: [
                              Text(
                                item.type == 'CEPU' ? "🚨 Di-Cepukan" : "📋 Izin Waigama",
                                style: TextStyle(
                                  fontSize: 11,
                                  fontWeight: FontWeight.w600,
                                  color: item.type == 'CEPU' ? const Color(0xFFDC2626) : const Color(0xFF1E60F2),
                                ),
                              ),
                              const SizedBox(width: 6),
                              Text(
                                "• $timeRange",
                                style: const TextStyle(fontSize: 11, color: Color(0xFF94A3B8)),
                              ),
                            ],
                          ),
                        ],
                      ),
                    ),

                    const SizedBox(width: 10),

                    // Durasi & Persentase
                    Column(
                      crossAxisAlignment: CrossAxisAlignment.end,
                      children: [
                        Text(
                          "${item.durationMinutes} mnt",
                          style: const TextStyle(
                            fontSize: 13.5,
                            fontWeight: FontWeight.bold,
                            color: Color(0xFF0F172A),
                          ),
                        ),
                        Text(
                          "$percent%",
                          style: const TextStyle(fontSize: 10.5, color: Color(0xFF64748B), fontWeight: FontWeight.w500),
                        ),
                      ],
                    ),
                  ],
                );
              },
            ),
          ],
        ],
      ),
    );
  }

  Widget _buildComingSoonPlaceholder() {
    return Container(
      width: double.infinity,
      padding: const EdgeInsets.symmetric(vertical: 60, horizontal: 24),
      decoration: BoxDecoration(
        color: Colors.white,
        borderRadius: BorderRadius.circular(22),
        border: Border.all(color: const Color(0xFFE2E8F0)),
      ),
      child: Column(
        children: [
          Container(
            width: 64,
            height: 64,
            decoration: const BoxDecoration(
              color: Color(0xFFEFF6FF),
              shape: BoxShape.circle,
            ),
            child: const Icon(Icons.analytics_outlined, color: Color(0xFF1E60F2), size: 32),
          ),
          const SizedBox(height: 16),
          Text(
            "Analisis $_selectedPeriod",
            style: const TextStyle(
              fontSize: 17,
              fontWeight: FontWeight.bold,
              color: Color(0xFF0F172A),
            ),
          ),
          const SizedBox(height: 6),
          const Text(
            "Fitur analisis bulanan dan tahunan sedang disiapkan untuk update berikutnya.",
            textAlign: TextAlign.center,
            style: TextStyle(fontSize: 12.5, color: Color(0xFF64748B)),
          ),
        ],
      ),
    );
  }
}
