import 'dart:math';
import 'package:flutter/material.dart';
import 'package:intl/intl.dart';
import '../models/user_model.dart';
import '../models/office_model.dart';
import '../models/permit_model.dart';
import '../models/cepu_model.dart';
import '../core/services/permit_service.dart';
import '../core/services/cepu_service.dart';
import '../core/services/attendance_service.dart';
import 'record_permit_map_view.dart';
import 'cepu_map_report_view.dart';

/// Item aktivitas selesai untuk visualisasi chart "Ngapain Aja?"
class AnalyticsActivityItem {
  final String id;
  final String type; // 'PERMIT', 'CEPU', atau 'OTHERS'
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
  final OfficeModel? office;

  const AnalyticsView({
    Key? key,
    required this.user,
    this.office,
  }) : super(key: key);

  @override
  State<AnalyticsView> createState() => _AnalyticsViewState();
}

class _AnalyticsViewState extends State<AnalyticsView> {
  final PermitService _permitService = PermitService();
  final CepuService _cepuService = CepuService();
  final AttendanceService _attendanceService = AttendanceService();
  OfficeModel? _office;

  // Period: 'HARIAN', 'BULANAN', 'TAHUNAN'
  String _selectedPeriod = 'HARIAN';

  DateTime _selectedDate = DateTime.now();
  bool _isLoading = true;

  // Metrics Terpilih
  int _outOfOfficeMinutes = 0;
  int _comparisonOutOfOfficeMinutes = 0; // Kemarin / Bulan Lalu / Tahun Lalu
  int _validCepuCount = 0;
  int _totalCepuCount = 0;

  // Data Kegiatan untuk "Ngapain Aja?"
  List<AnalyticsActivityItem> _activities = [];

  // Data Rekap Vertikal (Bulanan: tanggal 1..31 -> menit; Tahunan: bulan 1..12 -> menit)
  Map<int, int> _verticalChartData = {};
  int? _selectedVerticalIndex; // Tanggal / Bulan yang sedang diketuk

  // Curated Palette Colors untuk Stacked Bar Chart
  final List<Color> _chartColors = const [
    Color(0xFFEA580C), // Sunset Orange
    Color(0xFF6366F1), // Royal Indigo
    Color(0xFFE11D48), // Rose
    Color(0xFF1E60F2), // Royal Azure
    Color(0xFF0D9488), // Teal Emerald
    Color(0xFFD97706), // Amber
    Color(0xFF8B5CF6), // Violet
    Color(0xFF06B6D4), // Cyan
  ];

  @override
  void initState() {
    super.initState();
    _office = widget.office;
    if (_office == null) {
      _loadOfficeData();
    }
    _loadAnalyticsData();
  }

  Future<void> _loadOfficeData() async {
    try {
      final office = await _attendanceService.getOfficeConfig(widget.user.officeId);
      if (mounted) setState(() => _office = office);
    } catch (_) {}
  }

  String _formatDateString(DateTime dt) => DateFormat('yyyy-MM-dd').format(dt);

  int _getDaysInMonth(int year, int month) => DateTime(year, month + 1, 0).day;

  /// Format menit ke "X Jam Y Menit" jika > 60 menit
  String _formatDurationFull(int minutes) {
    if (minutes < 60) {
      return "$minutes Menit";
    }
    final hours = minutes ~/ 60;
    final remainingMinutes = minutes % 60;
    if (remainingMinutes == 0) {
      return "$hours Jam";
    }
    return "$hours Jam $remainingMinutes Menit";
  }

  String _formatDurationShort(int minutes) {
    if (minutes < 60) {
      return "$minutes mnt";
    }
    final hours = minutes ~/ 60;
    final remainingMinutes = minutes % 60;
    if (remainingMinutes == 0) {
      return "$hours jam";
    }
    return "$hours jam $remainingMinutes mnt";
  }

  Future<void> _loadAnalyticsData() async {
    setState(() {
      _isLoading = true;
      _selectedVerticalIndex = null;
    });

    if (_selectedPeriod == 'HARIAN') {
      await _loadHarianData();
    } else if (_selectedPeriod == 'BULANAN') {
      await _loadBulananData();
    } else {
      await _loadTahunanData();
    }
  }

  // ===========================================================================
  // 1. LOAD DATA HARIAN
  // ===========================================================================
  Future<void> _loadHarianData() async {
    final selectedDateStr = _formatDateString(_selectedDate);
    final yesterdayDate = _selectedDate.subtract(const Duration(days: 1));
    final yesterdayDateStr = _formatDateString(yesterdayDate);

    // Izin & Cepu Hari Terpilih
    final todayPermits = await _permitService.getUserDailyPermits(widget.user.uid, selectedDateStr);
    final completedPermits = todayPermits.where((p) => p.endTime != null).toList();

    final todayValidCepu = await _cepuService.getDailyValidCepuForTarget(widget.user.uid, selectedDateStr);
    final completedValidCepu = todayValidCepu.where((c) => c.endTime != null).toList();

    final allTodayCepu = await _cepuService.getAllDailyCepuForTarget(widget.user.uid, selectedDateStr);

    // Kemarin
    final yesterdayPermits = await _permitService.getUserDailyPermits(widget.user.uid, yesterdayDateStr);
    final completedYesterdayPermits = yesterdayPermits.where((p) => p.endTime != null).toList();
    final yesterdayValidCepu = await _cepuService.getDailyValidCepuForTarget(widget.user.uid, yesterdayDateStr);
    final completedYesterdayCepu = yesterdayValidCepu.where((c) => c.endTime != null).toList();

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

    items.sort((a, b) => b.durationMinutes.compareTo(a.durationMinutes));

    int yesterdayMinutes = 0;
    for (final p in completedYesterdayPermits) {
      yesterdayMinutes += max(0, p.endTime!.difference(p.startTime).inMinutes);
    }
    for (final c in completedYesterdayCepu) {
      yesterdayMinutes += max(0, c.endTime!.difference(c.startTime).inMinutes);
    }

    if (mounted) {
      setState(() {
        _outOfOfficeMinutes = todayMinutes;
        _comparisonOutOfOfficeMinutes = yesterdayMinutes;
        _totalCepuCount = allTodayCepu.length;
        _validCepuCount = allTodayCepu.where((c) => c.isValid).length;
        _activities = items;
        _verticalChartData = {};
        _isLoading = false;
      });
    }
  }

  // ===========================================================================
  // 2. LOAD DATA BULANAN
  // ===========================================================================
  Future<void> _loadBulananData() async {
    final year = _selectedDate.year;
    final month = _selectedDate.month;
    final daysInMonth = _getDaysInMonth(year, month);

    final startStr = "$year-${month.toString().padLeft(2, '0')}-01";
    final endStr = "$year-${month.toString().padLeft(2, '0')}-${daysInMonth.toString().padLeft(2, '0')}";

    // Bulan lalu untuk perbandingan
    final prevMonthDt = DateTime(year, month - 1, 1);
    final prevDaysInMonth = _getDaysInMonth(prevMonthDt.year, prevMonthDt.month);
    final prevStartStr = "${prevMonthDt.year}-${prevMonthDt.month.toString().padLeft(2, '0')}-01";
    final prevEndStr = "${prevMonthDt.year}-${prevMonthDt.month.toString().padLeft(2, '0')}-${prevDaysInMonth.toString().padLeft(2, '0')}";

    final monthPermits = await _permitService.getUserDateRangePermits(widget.user.uid, startStr, endStr);
    final completedPermits = monthPermits.where((p) => p.endTime != null).toList();

    final monthValidCepu = await _cepuService.getValidCepuDateRangeForTarget(widget.user.uid, startStr, endStr);
    final completedValidCepu = monthValidCepu.where((c) => c.endTime != null).toList();

    final allMonthCepu = await _cepuService.getAllCepuDateRangeForTarget(widget.user.uid, startStr, endStr);

    // Bulan Lalu
    final prevPermits = await _permitService.getUserDateRangePermits(widget.user.uid, prevStartStr, prevEndStr);
    final completedPrevPermits = prevPermits.where((p) => p.endTime != null).toList();
    final prevValidCepu = await _cepuService.getValidCepuDateRangeForTarget(widget.user.uid, prevStartStr, prevEndStr);
    final completedPrevCepu = prevValidCepu.where((c) => c.endTime != null).toList();

    int totalMonthMinutes = 0;
    final List<AnalyticsActivityItem> rawItems = [];
    int colorIdx = 0;

    // Rekap harian: tanggal 1..daysInMonth
    final Map<int, int> dailyMap = {};
    for (int d = 1; d <= daysInMonth; d++) {
      dailyMap[d] = 0;
    }

    for (final p in completedPermits) {
      final mins = max(0, p.endTime!.difference(p.startTime).inMinutes);
      totalMonthMinutes += mins;
      final dt = DateTime.tryParse(p.date);
      if (dt != null && dt.month == month && dt.year == year) {
        dailyMap[dt.day] = (dailyMap[dt.day] ?? 0) + mins;
      }
      rawItems.add(AnalyticsActivityItem(
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
      totalMonthMinutes += mins;
      final dt = DateTime.tryParse(c.date);
      if (dt != null && dt.month == month && dt.year == year) {
        dailyMap[dt.day] = (dailyMap[dt.day] ?? 0) + mins;
      }
      rawItems.add(AnalyticsActivityItem(
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

    rawItems.sort((a, b) => b.durationMinutes.compareTo(a.durationMinutes));

    // POIN 4: "Ngapain Aja" hanya menampilkan 5 izin terlama, sisanya ke kategori "Lainnya"
    final List<AnalyticsActivityItem> processedItems = [];
    if (rawItems.length <= 5) {
      processedItems.addAll(rawItems);
    } else {
      processedItems.addAll(rawItems.take(5));
      final others = rawItems.skip(5).toList();
      int othersDuration = 0;
      for (final o in others) {
        othersDuration += o.durationMinutes;
      }
      if (othersDuration > 0) {
        processedItems.add(AnalyticsActivityItem(
          id: 'others',
          type: 'OTHERS',
          title: 'Lainnya',
          description: '${others.length} izin & laporan lainnya',
          startTime: others.first.startTime,
          endTime: others.first.endTime,
          durationMinutes: othersDuration,
          color: const Color(0xFF94A3B8), // Slate Gray
        ));
      }
    }

    int prevMonthMinutes = 0;
    for (final p in completedPrevPermits) {
      prevMonthMinutes += max(0, p.endTime!.difference(p.startTime).inMinutes);
    }
    for (final c in completedPrevCepu) {
      prevMonthMinutes += max(0, c.endTime!.difference(c.startTime).inMinutes);
    }

    if (mounted) {
      setState(() {
        _outOfOfficeMinutes = totalMonthMinutes;
        _comparisonOutOfOfficeMinutes = prevMonthMinutes;
        _totalCepuCount = allMonthCepu.length;
        _validCepuCount = allMonthCepu.where((c) => c.isValid).length;
        _activities = processedItems;
        _verticalChartData = dailyMap;
        _isLoading = false;
      });
    }
  }

  // ===========================================================================
  // 3. LOAD DATA TAHUNAN
  // ===========================================================================
  Future<void> _loadTahunanData() async {
    final year = _selectedDate.year;
    final startStr = "$year-01-01";
    final endStr = "$year-12-31";

    final prevYear = year - 1;
    final prevStartStr = "$prevYear-01-01";
    final prevEndStr = "$prevYear-12-31";

    final yearPermits = await _permitService.getUserDateRangePermits(widget.user.uid, startStr, endStr);
    final completedPermits = yearPermits.where((p) => p.endTime != null).toList();

    final yearValidCepu = await _cepuService.getValidCepuDateRangeForTarget(widget.user.uid, startStr, endStr);
    final completedValidCepu = yearValidCepu.where((c) => c.endTime != null).toList();

    final allYearCepu = await _cepuService.getAllCepuDateRangeForTarget(widget.user.uid, startStr, endStr);

    // Tahun Lalu
    final prevPermits = await _permitService.getUserDateRangePermits(widget.user.uid, prevStartStr, prevEndStr);
    final completedPrevPermits = prevPermits.where((p) => p.endTime != null).toList();
    final prevValidCepu = await _cepuService.getValidCepuDateRangeForTarget(widget.user.uid, prevStartStr, prevEndStr);
    final completedPrevCepu = prevValidCepu.where((c) => c.endTime != null).toList();

    int totalYearMinutes = 0;
    final List<AnalyticsActivityItem> rawItems = [];
    int colorIdx = 0;

    // Rekap bulanan: bulan 1..12
    final Map<int, int> monthlyMap = {};
    for (int m = 1; m <= 12; m++) {
      monthlyMap[m] = 0;
    }

    for (final p in completedPermits) {
      final mins = max(0, p.endTime!.difference(p.startTime).inMinutes);
      totalYearMinutes += mins;
      final dt = DateTime.tryParse(p.date);
      if (dt != null && dt.year == year) {
        monthlyMap[dt.month] = (monthlyMap[dt.month] ?? 0) + mins;
      }
      rawItems.add(AnalyticsActivityItem(
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
      totalYearMinutes += mins;
      final dt = DateTime.tryParse(c.date);
      if (dt != null && dt.year == year) {
        monthlyMap[dt.month] = (monthlyMap[dt.month] ?? 0) + mins;
      }
      rawItems.add(AnalyticsActivityItem(
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

    rawItems.sort((a, b) => b.durationMinutes.compareTo(a.durationMinutes));

    // POIN 5: Top 5 terlama + "Lainnya" pada tahunan
    final List<AnalyticsActivityItem> processedItems = [];
    if (rawItems.length <= 5) {
      processedItems.addAll(rawItems);
    } else {
      processedItems.addAll(rawItems.take(5));
      final others = rawItems.skip(5).toList();
      int othersDuration = 0;
      for (final o in others) {
        othersDuration += o.durationMinutes;
      }
      if (othersDuration > 0) {
        processedItems.add(AnalyticsActivityItem(
          id: 'others',
          type: 'OTHERS',
          title: 'Lainnya',
          description: '${others.length} izin & laporan lainnya',
          startTime: others.first.startTime,
          endTime: others.first.endTime,
          durationMinutes: othersDuration,
          color: const Color(0xFF94A3B8),
        ));
      }
    }

    int prevYearMinutes = 0;
    for (final p in completedPrevPermits) {
      prevYearMinutes += max(0, p.endTime!.difference(p.startTime).inMinutes);
    }
    for (final c in completedPrevCepu) {
      prevYearMinutes += max(0, c.endTime!.difference(c.startTime).inMinutes);
    }

    if (mounted) {
      setState(() {
        _outOfOfficeMinutes = totalYearMinutes;
        _comparisonOutOfOfficeMinutes = prevYearMinutes;
        _totalCepuCount = allYearCepu.length;
        _validCepuCount = allYearCepu.where((c) => c.isValid).length;
        _activities = processedItems;
        _verticalChartData = monthlyMap;
        _isLoading = false;
      });
    }
  }

  // ===========================================================================
  // DATE NAVIGATION CONTROLS
  // ===========================================================================
  void _previousDatePeriod() {
    if (_selectedPeriod == 'HARIAN') {
      setState(() => _selectedDate = _selectedDate.subtract(const Duration(days: 1)));
    } else if (_selectedPeriod == 'BULANAN') {
      setState(() => _selectedDate = DateTime(_selectedDate.year, _selectedDate.month - 1, 1));
    } else {
      setState(() => _selectedDate = DateTime(_selectedDate.year - 1, 1, 1));
    }
    _loadAnalyticsData();
  }

  void _nextDatePeriod() {
    if (_selectedPeriod == 'HARIAN') {
      setState(() => _selectedDate = _selectedDate.add(const Duration(days: 1)));
    } else if (_selectedPeriod == 'BULANAN') {
      setState(() => _selectedDate = DateTime(_selectedDate.year, _selectedDate.month + 1, 1));
    } else {
      setState(() => _selectedDate = DateTime(_selectedDate.year + 1, 1, 1));
    }
    _loadAnalyticsData();
  }

  Future<void> _pickDateOrPeriod() async {
    if (_selectedPeriod == 'HARIAN') {
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
    } else if (_selectedPeriod == 'BULANAN') {
      // Month Picker Dialog
      showDialog(
        context: context,
        builder: (context) {
          return AlertDialog(
            title: const Text("Pilih Bulan", style: TextStyle(fontWeight: FontWeight.bold, fontSize: 16)),
            content: SizedBox(
              width: double.maxFinite,
              child: Wrap(
                spacing: 8,
                runSpacing: 8,
                children: List.generate(12, (index) {
                  final m = index + 1;
                  final isCurrent = m == _selectedDate.month;
                  final monthName = DateFormat('MMM', 'id_ID').format(DateTime(2026, m, 1));
                  return ChoiceChip(
                    label: Text(monthName),
                    selected: isCurrent,
                    selectedColor: const Color(0xFF1E60F2),
                    labelStyle: TextStyle(
                      color: isCurrent ? Colors.white : const Color(0xFF334155),
                      fontWeight: FontWeight.bold,
                    ),
                    onSelected: (selected) {
                      Navigator.pop(context);
                      if (selected) {
                        setState(() => _selectedDate = DateTime(_selectedDate.year, m, 1));
                        _loadAnalyticsData();
                      }
                    },
                  );
                }),
              ),
            ),
          );
        },
      );
    } else {
      // Year Picker Dialog
      showDialog(
        context: context,
        builder: (context) {
          return AlertDialog(
            title: const Text("Pilih Tahun", style: TextStyle(fontWeight: FontWeight.bold, fontSize: 16)),
            content: SizedBox(
              width: double.maxFinite,
              height: 250,
              child: YearPicker(
                firstDate: DateTime(2020),
                lastDate: DateTime(2035),
                selectedDate: _selectedDate,
                onChanged: (DateTime dt) {
                  Navigator.pop(context);
                  setState(() => _selectedDate = DateTime(dt.year, 1, 1));
                  _loadAnalyticsData();
                },
              ),
            ),
          );
        },
      );
    }
  }

  void _navigateToRecordPermit() async {
    final office = _office ?? await _attendanceService.getOfficeConfig(widget.user.officeId);
    if (office == null || !mounted) return;
    final res = await Navigator.push(
      context,
      MaterialPageRoute(
        builder: (context) => RecordPermitMapView(
          user: widget.user,
          office: office,
        ),
      ),
    );
    if (res == true) _loadAnalyticsData();
  }

  void _navigateToCepuReport() async {
    final office = _office ?? await _attendanceService.getOfficeConfig(widget.user.officeId);
    if (office == null || !mounted) return;
    final res = await Navigator.push(
      context,
      MaterialPageRoute(
        builder: (context) => CepuMapReportView(
          reporter: widget.user,
          office: office,
        ),
      ),
    );
    if (res == true) _loadAnalyticsData();
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
          "Insight",
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
            padding: const EdgeInsets.fromLTRB(20, 16, 20, 24),
            child: Column(
              crossAxisAlignment: CrossAxisAlignment.start,
              children: [
                // 1. PERIOD SELECTOR (Harian, Bulanan, Tahunan)
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

                // 2. DATE / MONTH / YEAR SELECTOR BAR
                _buildDatePickerBar(),

                const SizedBox(height: 20),

                // 3. KONTEN INSIGHT
                if (_isLoading)
                  const Padding(
                    padding: EdgeInsets.symmetric(vertical: 60),
                    child: Center(child: CircularProgressIndicator(color: Color(0xFF1E60F2))),
                  )
                else ...[
                  // 2 Summary Cards: "Lama Tidak di Kantor" & "Berapa Kali Dilaporkan"
                  _buildSummaryCards(),

                  // Vertikal Chart (Rekap Harian untuk Bulanan / Rekap Bulanan untuk Tahunan)
                  if (_selectedPeriod != 'HARIAN') ...[
                    const SizedBox(height: 20),
                    _buildVerticalChartCard(),
                  ],

                  const SizedBox(height: 20),

                  // Stacked Horizontal Bar Chart Card: "Ngapain Aja?"
                  _buildNgapainAjaCard(),

                  // POIN 5: Extra Thoughtful Insights
                  const SizedBox(height: 20),
                  _buildExtraInsightsCard(),
                ],
              ],
            ),
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
              onPressed: () => Navigator.pop(context),
              tooltip: "Home",
            ),
            IconButton(
              icon: const Icon(Icons.bar_chart_rounded, color: Color(0xFF1E60F2), size: 26),
              onPressed: () {},
              tooltip: "Insight",
            ),
            Transform.translate(
              offset: const Offset(0, -12),
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
            IconButton(
              icon: const Icon(Icons.campaign_outlined, color: Color(0xFF94A3B8), size: 26),
              onPressed: _navigateToCepuReport,
              tooltip: "Cepu",
            ),
            IconButton(
              icon: const Icon(Icons.person_outline_rounded, color: Color(0xFF94A3B8), size: 26),
              onPressed: () => Navigator.pop(context),
              tooltip: "Profil",
            ),
          ],
        ),
      ),
    );
  }

  Widget _buildPeriodTab(String label, String value) {
    final isSelected = _selectedPeriod == value;
    return Expanded(
      child: InkWell(
        onTap: () {
          if (_selectedPeriod != value) {
            setState(() => _selectedPeriod = value);
            _loadAnalyticsData();
          }
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
      dateLabel = "Tahun ${_selectedDate.year}";
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
            onPressed: _previousDatePeriod,
          ),
          Expanded(
            child: InkWell(
              onTap: _pickDateOrPeriod,
              borderRadius: BorderRadius.circular(10),
              child: Padding(
                padding: const EdgeInsets.symmetric(horizontal: 4, vertical: 6),
                child: Row(
                  mainAxisAlignment: MainAxisAlignment.center,
                  children: [
                    const Icon(Icons.calendar_month_rounded, color: Color(0xFF1E60F2), size: 18),
                    const SizedBox(width: 6),
                    Flexible(
                      child: Text(
                        dateLabel,
                        style: const TextStyle(
                          fontSize: 13,
                          fontWeight: FontWeight.bold,
                          color: Color(0xFF0F172A),
                        ),
                        maxLines: 1,
                        overflow: TextOverflow.ellipsis,
                      ),
                    ),
                  ],
                ),
              ),
            ),
          ),
          IconButton(
            icon: const Icon(Icons.chevron_right_rounded, color: Color(0xFF475569)),
            onPressed: _nextDatePeriod,
          ),
        ],
      ),
    );
  }

  /// Summary Cards: "Lama Tidak di Kantor" & "Berapa Kali Dilaporkan"
  Widget _buildSummaryCards() {
    final diff = _outOfOfficeMinutes - _comparisonOutOfOfficeMinutes;
    final bool isBetter = diff <= 0;
    final bool isSame = diff == 0;

    final isOverAnHour = _outOfOfficeMinutes >= 60;
    final displayHours = _outOfOfficeMinutes ~/ 60;
    final displayRemainingMins = _outOfOfficeMinutes % 60;

    String comparisonTarget = "kemarin";
    if (_selectedPeriod == 'BULANAN') comparisonTarget = "bln lalu";
    if (_selectedPeriod == 'TAHUNAN') comparisonTarget = "thn lalu";

    return Row(
      crossAxisAlignment: CrossAxisAlignment.start,
      children: [
        // Card 1: "Lama Tidak di Kantor"
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
                  crossAxisAlignment: CrossAxisAlignment.start,
                  children: [
                    const Expanded(
                      child: Text(
                        "Lama Tidak di Kantor",
                        style: TextStyle(
                          fontSize: 12.5,
                          fontWeight: FontWeight.w600,
                          color: Color(0xFF64748B),
                          height: 1.25,
                        ),
                        maxLines: 2,
                        softWrap: true,
                      ),
                    ),
                    const SizedBox(width: 4),
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

                // Format Jam & Menit
                if (isOverAnHour)
                  FittedBox(
                    fit: BoxFit.scaleDown,
                    alignment: Alignment.centerLeft,
                    child: RichText(
                      text: TextSpan(
                        children: [
                          TextSpan(
                            text: "$displayHours ",
                            style: const TextStyle(
                              fontSize: 26,
                              fontWeight: FontWeight.w900,
                              color: Color(0xFF0F172A),
                              letterSpacing: -0.5,
                            ),
                          ),
                          const TextSpan(
                            text: "Jam ",
                            style: TextStyle(fontSize: 14, fontWeight: FontWeight.bold, color: Color(0xFF64748B)),
                          ),
                          if (displayRemainingMins > 0) ...[
                            TextSpan(
                              text: "$displayRemainingMins ",
                              style: const TextStyle(
                                fontSize: 22,
                                fontWeight: FontWeight.w900,
                                color: Color(0xFF0F172A),
                              ),
                            ),
                            const TextSpan(
                              text: "Mnt",
                              style: TextStyle(fontSize: 14, fontWeight: FontWeight.bold, color: Color(0xFF64748B)),
                            ),
                          ],
                        ],
                      ),
                    ),
                  )
                else
                  Text(
                    "$_outOfOfficeMinutes",
                    style: const TextStyle(
                      fontSize: 28,
                      fontWeight: FontWeight.w900,
                      color: Color(0xFF0F172A),
                      letterSpacing: -0.5,
                    ),
                  ),

                Text(
                  isOverAnHour ? "Total Durasi" : "Menit",
                  style: const TextStyle(fontSize: 11.5, color: Color(0xFF64748B), fontWeight: FontWeight.w500),
                ),
                const SizedBox(height: 10),

                // Insight Hijau / Merah Berdasarkan Perbandingan
                Row(
                  crossAxisAlignment: CrossAxisAlignment.center,
                  children: [
                    Icon(
                      isSame
                          ? Icons.remove_rounded
                          : (isBetter ? Icons.trending_down_rounded : Icons.trending_up_rounded),
                      size: 15,
                      color: isSame
                          ? const Color(0xFF64748B)
                          : (isBetter ? const Color(0xFF10B981) : const Color(0xFFEF4444)),
                    ),
                    const SizedBox(width: 4),
                    Expanded(
                      child: Text(
                        isSame
                            ? "Sama vs $comparisonTarget"
                            : (isBetter
                                ? "-${diff.abs()} mnt vs $comparisonTarget"
                                : "+${diff.abs()} mnt vs $comparisonTarget"),
                        style: TextStyle(
                          fontSize: 11,
                          fontWeight: FontWeight.bold,
                          color: isSame
                              ? const Color(0xFF64748B)
                              : (isBetter ? const Color(0xFF10B981) : const Color(0xFFEF4444)),
                        ),
                        maxLines: 2,
                        softWrap: true,
                      ),
                    ),
                  ],
                ),
              ],
            ),
          ),
        ),

        const SizedBox(width: 12),

        // Card 2: "Berapa Kali Dilaporkan"
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
                  crossAxisAlignment: CrossAxisAlignment.start,
                  children: [
                    const Expanded(
                      child: Text(
                        "Berapa Kali Dilaporkan",
                        style: TextStyle(
                          fontSize: 12.5,
                          fontWeight: FontWeight.w600,
                          color: Color(0xFF64748B),
                          height: 1.25,
                        ),
                        maxLines: 2,
                        softWrap: true,
                      ),
                    ),
                    const SizedBox(width: 4),
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

                // Rasio Valid / Total
                RichText(
                  text: TextSpan(
                    children: [
                      TextSpan(
                        text: "$_validCepuCount ",
                        style: const TextStyle(
                          fontSize: 28,
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
                const SizedBox(height: 10),

                // Status verifikasi
                Row(
                  crossAxisAlignment: CrossAxisAlignment.center,
                  children: [
                    Icon(
                      _validCepuCount == 0 ? Icons.check_circle_outline_rounded : Icons.info_outline_rounded,
                      size: 15,
                      color: _validCepuCount == 0 ? const Color(0xFF10B981) : const Color(0xFFEA580C),
                    ),
                    const SizedBox(width: 4),
                    Expanded(
                      child: Text(
                        _validCepuCount == 0 ? "Nihil Pelanggaran" : "$_validCepuCount Laporan Valid",
                        style: TextStyle(
                          fontSize: 11,
                          fontWeight: FontWeight.bold,
                          color: _validCepuCount == 0 ? const Color(0xFF10B981) : const Color(0xFFEA580C),
                        ),
                        maxLines: 2,
                        softWrap: true,
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

  // ===========================================================================
  // POIN 4 & 5: VERTICAL BAR CHART (Rekap Harian / Rekap Bulanan)
  // ===========================================================================
  Widget _buildVerticalChartCard() {
    final isBulanan = _selectedPeriod == 'BULANAN';
    final title = isBulanan ? "Rekap Harian" : "Rekap Bulanan";
    final subtitle = isBulanan
        ? "Durasi di luar kantor per tanggal"
        : "Durasi di luar kantor per bulan";

    final maxVal = _verticalChartData.values.isEmpty
        ? 1
        : max(1, _verticalChartData.values.reduce(max));

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
                  Text(
                    title,
                    style: const TextStyle(
                      fontSize: 17,
                      fontWeight: FontWeight.bold,
                      color: Color(0xFF0F172A),
                    ),
                  ),
                  const SizedBox(height: 2),
                  Text(
                    subtitle,
                    style: const TextStyle(fontSize: 12, color: Color(0xFF64748B)),
                  ),
                ],
              ),
              Container(
                padding: const EdgeInsets.symmetric(horizontal: 8, vertical: 4),
                decoration: BoxDecoration(
                  color: const Color(0xFFEFF6FF),
                  borderRadius: BorderRadius.circular(8),
                ),
                child: Text(
                  "Puncak: ${_formatDurationShort(maxVal)}",
                  style: const TextStyle(fontSize: 11, fontWeight: FontWeight.bold, color: Color(0xFF1E60F2)),
                ),
              ),
            ],
          ),

          const SizedBox(height: 20),

          // Chart Area
          if (isBulanan)
            // Rekap Harian (Scrollable horizontal bar chart for 28..31 days)
            SizedBox(
              height: 150,
              child: ListView.separated(
                scrollDirection: Axis.horizontal,
                physics: const BouncingScrollPhysics(),
                itemCount: _verticalChartData.length,
                separatorBuilder: (_, __) => const SizedBox(width: 8),
                itemBuilder: (context, index) {
                  final day = index + 1;
                  final minutes = _verticalChartData[day] ?? 0;
                  final isSelected = _selectedVerticalIndex == day;
                  final ratio = maxVal > 0 ? (minutes / maxVal).clamp(0.0, 1.0) : 0.0;
                  final barHeight = max(4.0, ratio * 95);

                  return InkWell(
                    onTap: () {
                      setState(() {
                        _selectedVerticalIndex = isSelected ? null : day;
                      });
                    },
                    borderRadius: BorderRadius.circular(8),
                    child: Container(
                      width: 24,
                      padding: const EdgeInsets.symmetric(vertical: 4),
                      decoration: isSelected
                          ? BoxDecoration(
                              color: const Color(0xFFEFF6FF),
                              borderRadius: BorderRadius.circular(8),
                              border: Border.all(color: const Color(0xFF1E60F2), width: 1.2),
                            )
                          : null,
                      child: Column(
                        mainAxisAlignment: MainAxisAlignment.end,
                        children: [
                          // Value indicator on top if selected or > 0
                          if (isSelected)
                            FittedBox(
                              fit: BoxFit.scaleDown,
                              child: Text(
                                "$minutes m",
                                style: const TextStyle(
                                  fontSize: 9,
                                  fontWeight: FontWeight.w900,
                                  color: Color(0xFF1E60F2),
                                ),
                              ),
                            )
                          else
                            const SizedBox(height: 12),
                          const Spacer(),
                          // The Vertical Bar
                          Container(
                            width: 10,
                            height: barHeight,
                            decoration: BoxDecoration(
                              gradient: LinearGradient(
                                colors: minutes > 0
                                    ? [const Color(0xFF1E60F2), const Color(0xFF38BDF8)]
                                    : [const Color(0xFFE2E8F0), const Color(0xFFCBD5E1)],
                                begin: Alignment.bottomCenter,
                                end: Alignment.topCenter,
                              ),
                              borderRadius: BorderRadius.circular(5),
                            ),
                          ),
                          const SizedBox(height: 6),
                          // X-axis label (Day number)
                          Text(
                            "$day",
                            style: TextStyle(
                              fontSize: 10,
                              fontWeight: isSelected ? FontWeight.bold : FontWeight.w500,
                              color: isSelected ? const Color(0xFF1E60F2) : const Color(0xFF64748B),
                            ),
                          ),
                        ],
                      ),
                    ),
                  );
                },
              ),
            )
          else
            // Rekap Bulanan (12 bars across width for Tahunan)
            SizedBox(
              height: 150,
              child: Row(
                mainAxisAlignment: MainAxisAlignment.spaceBetween,
                crossAxisAlignment: CrossAxisAlignment.end,
                children: List.generate(12, (index) {
                  final month = index + 1;
                  final minutes = _verticalChartData[month] ?? 0;
                  final isSelected = _selectedVerticalIndex == month;
                  final ratio = maxVal > 0 ? (minutes / maxVal).clamp(0.0, 1.0) : 0.0;
                  final barHeight = max(4.0, ratio * 95);
                  final monthLabel = DateFormat('MMM', 'id_ID').format(DateTime(2026, month, 1));

                  return Expanded(
                    child: InkWell(
                      onTap: () {
                        setState(() {
                          _selectedVerticalIndex = isSelected ? null : month;
                        });
                      },
                      borderRadius: BorderRadius.circular(6),
                      child: Container(
                        padding: const EdgeInsets.symmetric(vertical: 4),
                        decoration: isSelected
                            ? BoxDecoration(
                                color: const Color(0xFFEFF6FF),
                                borderRadius: BorderRadius.circular(6),
                                border: Border.all(color: const Color(0xFF1E60F2), width: 1.2),
                              )
                            : null,
                        child: Column(
                          mainAxisAlignment: MainAxisAlignment.end,
                          children: [
                            if (isSelected)
                              FittedBox(
                                fit: BoxFit.scaleDown,
                                child: Text(
                                  _formatDurationShort(minutes),
                                  style: const TextStyle(
                                    fontSize: 8.5,
                                    fontWeight: FontWeight.w900,
                                    color: Color(0xFF1E60F2),
                                  ),
                                ),
                              )
                            else
                              const SizedBox(height: 12),
                            const Spacer(),
                            Container(
                              width: 12,
                              height: barHeight,
                              decoration: BoxDecoration(
                                gradient: LinearGradient(
                                  colors: minutes > 0
                                      ? [const Color(0xFF1E60F2), const Color(0xFF38BDF8)]
                                      : [const Color(0xFFE2E8F0), const Color(0xFFCBD5E1)],
                                  begin: Alignment.bottomCenter,
                                  end: Alignment.topCenter,
                                ),
                                borderRadius: BorderRadius.circular(6),
                              ),
                            ),
                            const SizedBox(height: 6),
                            Text(
                              monthLabel,
                              style: TextStyle(
                                fontSize: 9.5,
                                fontWeight: isSelected ? FontWeight.bold : FontWeight.w500,
                                color: isSelected ? const Color(0xFF1E60F2) : const Color(0xFF64748B),
                              ),
                            ),
                          ],
                        ),
                      ),
                    ),
                  );
                }),
              ),
            ),

          // Detail box on tap
          if (_selectedVerticalIndex != null) ...[
            const SizedBox(height: 14),
            Container(
              padding: const EdgeInsets.symmetric(horizontal: 14, vertical: 10),
              decoration: BoxDecoration(
                color: const Color(0xFFF8FAFC),
                borderRadius: BorderRadius.circular(12),
                border: Border.all(color: const Color(0xFFE2E8F0)),
              ),
              child: Row(
                children: [
                  const Icon(Icons.info_outline_rounded, color: Color(0xFF1E60F2), size: 18),
                  const SizedBox(width: 8),
                  Expanded(
                    child: Text(
                      isBulanan
                          ? "Tanggal $_selectedVerticalIndex: ${_formatDurationFull(_verticalChartData[_selectedVerticalIndex!] ?? 0)}"
                          : "Bulan ${DateFormat('MMMM', 'id_ID').format(DateTime(2026, _selectedVerticalIndex!, 1))}: ${_formatDurationFull(_verticalChartData[_selectedVerticalIndex!] ?? 0)}",
                      style: const TextStyle(
                        fontSize: 12.5,
                        fontWeight: FontWeight.bold,
                        color: Color(0xFF0F172A),
                      ),
                    ),
                  ),
                ],
              ),
            ),
          ],
        ],
      ),
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
              Expanded(
                child: Column(
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
                      "Total di luar kantor: ${_formatDurationFull(_outOfOfficeMinutes)}",
                      style: const TextStyle(fontSize: 12, color: Color(0xFF64748B)),
                      maxLines: 2,
                    ),
                  ],
                ),
              ),
              Container(
                padding: const EdgeInsets.symmetric(horizontal: 10, vertical: 5),
                decoration: BoxDecoration(
                  color: const Color(0xFFEFF6FF),
                  borderRadius: BorderRadius.circular(10),
                ),
                child: Row(
                  children: const [
                    Icon(Icons.sort_rounded, size: 14, color: Color(0xFF1E60F2)),
                    SizedBox(width: 4),
                    Text(
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
                    "Seluruh waktu kerja dihabiskan di kantor pada periode ini.",
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
                    final double flexWeight = _outOfOfficeMinutes > 0
                        ? item.durationMinutes / _outOfOfficeMinutes
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
                final percent = _outOfOfficeMinutes > 0
                    ? ((item.durationMinutes / _outOfOfficeMinutes) * 100).toStringAsFixed(0)
                    : "0";
                final timeRange =
                    "${DateFormat('HH:mm').format(item.startTime)} - ${DateFormat('HH:mm').format(item.endTime)} WIT";

                final durationLabel = _formatDurationShort(item.durationMinutes);

                String badgeText = "📋 Izin Waigama";
                Color badgeColor = const Color(0xFF1E60F2);
                if (item.type == 'CEPU') {
                  badgeText = "🚨 Di-Cepukan";
                  badgeColor = const Color(0xFFDC2626);
                } else if (item.type == 'OTHERS') {
                  badgeText = "📦 Gabungan";
                  badgeColor = const Color(0xFF64748B);
                }

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
                                badgeText,
                                style: TextStyle(
                                  fontSize: 11,
                                  fontWeight: FontWeight.w600,
                                  color: badgeColor,
                                ),
                              ),
                              if (item.type != 'OTHERS') ...[
                                const SizedBox(width: 6),
                                Text(
                                  "• $timeRange",
                                  style: const TextStyle(fontSize: 11, color: Color(0xFF94A3B8)),
                                ),
                              ],
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
                          durationLabel,
                          style: const TextStyle(
                            fontSize: 13,
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

  // ===========================================================================
  // POIN 5: EXTRA THOUGHTFUL INSIGHTS
  // ===========================================================================
  Widget _buildExtraInsightsCard() {
    String peakTitle = "-";
    String peakSubtitle = "-";

    if (_selectedPeriod == 'HARIAN') {
      peakTitle = _outOfOfficeMinutes > 0 ? "Hari Aktif Izin" : "Hari Nihil Izin";
      peakSubtitle = _outOfOfficeMinutes > 0
          ? "Tercatat ${_activities.length} aktivitas di luar kantor"
          : "Disiplin kerja 100% berada di kantor";
    } else if (_selectedPeriod == 'BULANAN') {
      int maxDay = 1;
      int maxVal = 0;
      _verticalChartData.forEach((day, val) {
        if (val > maxVal) {
          maxVal = val;
          maxDay = day;
        }
      });
      peakTitle = maxVal > 0 ? "Paling Lama di Luar: Tgl $maxDay" : "Nihil Izin";
      peakSubtitle = maxVal > 0
          ? "Durasi puncak: ${_formatDurationFull(maxVal)}"
          : "Tidak ada riwayat izin selesai bulan ini";
    } else {
      int maxMonth = 1;
      int maxVal = 0;
      _verticalChartData.forEach((m, val) {
        if (val > maxVal) {
          maxVal = val;
          maxMonth = m;
        }
      });
      final monthName = DateFormat('MMMM', 'id_ID').format(DateTime(2026, maxMonth, 1));
      peakTitle = maxVal > 0 ? "Bulan Terbanyak Izin: $monthName" : "Nihil Izin";
      peakSubtitle = maxVal > 0
          ? "Total durasi puncak: ${_formatDurationFull(maxVal)}"
          : "Tidak ada riwayat izin selesai tahun ini";
    }

    return Container(
      padding: const EdgeInsets.all(16),
      decoration: BoxDecoration(
        color: const Color(0xFFF1F5F9),
        borderRadius: BorderRadius.circular(18),
      ),
      child: Row(
        children: [
          Container(
            width: 40,
            height: 40,
            decoration: BoxDecoration(
              color: Colors.white,
              borderRadius: BorderRadius.circular(12),
            ),
            child: const Icon(Icons.insights_rounded, color: Color(0xFF1E60F2), size: 22),
          ),
          const SizedBox(width: 14),
          Expanded(
            child: Column(
              crossAxisAlignment: CrossAxisAlignment.start,
              children: [
                Text(
                  peakTitle,
                  style: const TextStyle(
                    fontSize: 13.5,
                    fontWeight: FontWeight.bold,
                    color: Color(0xFF0F172A),
                  ),
                ),
                const SizedBox(height: 2),
                Text(
                  peakSubtitle,
                  style: const TextStyle(fontSize: 11.5, color: Color(0xFF64748B)),
                ),
              ],
            ),
          ),
        ],
      ),
    );
  }
}
