import 'dart:convert';
import 'dart:io';
import 'package:flutter/material.dart';
import 'package:flutter_map/flutter_map.dart';
import 'package:latlong2/latlong.dart';
import 'package:geolocator/geolocator.dart';
import 'package:image_picker/image_picker.dart';
import 'package:intl/intl.dart';
import '../models/user_model.dart';
import '../models/office_model.dart';
import '../models/cepu_model.dart';
import '../core/services/cepu_service.dart';
import '../core/services/location_service.dart';

class CepuMapReportView extends StatefulWidget {
  final UserModel reporter;
  final OfficeModel office;

  const CepuMapReportView({
    Key? key,
    required this.reporter,
    required this.office,
  }) : super(key: key);

  @override
  State<CepuMapReportView> createState() => _CepuMapReportViewState();
}

class _CepuMapReportViewState extends State<CepuMapReportView> {
  final MapController _mapController = MapController();
  final CepuService _cepuService = CepuService();
  final ImagePicker _picker = ImagePicker();
  final TextEditingController _descController = TextEditingController();

  List<UserModel> _employees = [];
  UserModel? _selectedTarget;
  bool _isLoadingEmployees = true;
  bool _isSubmitting = false;

  // Laporan Cepu aktif untuk user yang login (jika dilaporkan oleh orang lain dan sudah valid)
  CepuModel? _activeCepuForMe;
  bool _isRecordingReturn = false;

  Position? _currentPosition;
  double _distanceMeters = 999.0;
  double _accuracyMeters = 0.0;
  bool _isInRadius = false;
  bool _isMocked = false;
  bool _isLoadingLocation = true;
  String? _errorMessage;

  DateTime _startTime = DateTime.now();
  String _searchFilter = '';

  File? _imageFile;
  String? _imageBase64;

  @override
  void initState() {
    super.initState();
    _loadEmployees();
    _fetchCurrentLocation();
    _checkActiveCepuForMe();
  }

  @override
  void dispose() {
    _descController.dispose();
    super.dispose();
  }

  Future<void> _loadEmployees() async {
    setState(() => _isLoadingEmployees = true);
    final list = await _cepuService.getEmployeesList(excludeUid: widget.reporter.uid);
    if (mounted) {
      setState(() {
        _employees = list;
        _isLoadingEmployees = false;
      });
    }
  }

  Future<void> _checkActiveCepuForMe() async {
    final cepu = await _cepuService.getActiveCepuForTarget(widget.reporter.uid);
    if (mounted) {
      setState(() {
        _activeCepuForMe = cepu;
      });
    }
  }

  Future<void> _fetchCurrentLocation() async {
    setState(() {
      _isLoadingLocation = true;
      _errorMessage = null;
    });

    try {
      final res = await LocationService.verifyPresenceLocation(
        officeLat: widget.office.latitude,
        officeLng: widget.office.longitude,
        allowedRadiusMeters: widget.office.radiusMeters,
      );

      if (res.isSuccess && res.position != null) {
        setState(() {
          _currentPosition = res.position;
          _distanceMeters = res.distanceMeters;
          _accuracyMeters = res.position!.accuracy;
          _isInRadius = res.isWithinRadius;
          _isMocked = res.isMockLocation;
          _isLoadingLocation = false;
        });

        _mapController.move(
          LatLng(res.position!.latitude, res.position!.longitude),
          17.5,
        );
      } else {
        setState(() {
          _isLoadingLocation = false;
          _errorMessage = res.errorMessage ?? "Gagal mengambil lokasi GPS.";
          _isMocked = res.isMockLocation;
        });
      }
    } catch (e) {
      setState(() {
        _isLoadingLocation = false;
        _errorMessage = e.toString();
      });
    }
  }

  Future<void> _handleRecordReturnForMe() async {
    if (_activeCepuForMe == null) return;

    if (!_isInRadius || _isMocked) {
      ScaffoldMessenger.of(context).showSnackBar(
        SnackBar(
          content: Text(_isMocked
              ? "⚠️ Deteksi Fake GPS aktif!"
              : "⚠️ Anda harus sudah berada di dalam radius kantor untuk merekam waktu kembali!"),
          backgroundColor: const Color(0xFFEF4444),
        ),
      );
      return;
    }

    setState(() => _isRecordingReturn = true);

    try {
      await _cepuService.recordCepuReturnTime(
        cepuId: _activeCepuForMe!.id,
        office: widget.office,
      );

      if (mounted) {
        ScaffoldMessenger.of(context).showSnackBar(
          const SnackBar(
            content: Text("✅ Waktu kembali Cepu berhasil direkam! Status Anda telah diperbarui."),
            backgroundColor: Color(0xFF10B981),
          ),
        );
        setState(() {
          _activeCepuForMe = null;
        });
        Navigator.pop(context, true);
      }
    } catch (e) {
      if (mounted) {
        ScaffoldMessenger.of(context).showSnackBar(
          SnackBar(content: Text(e.toString()), backgroundColor: const Color(0xFFEF4444)),
        );
      }
    } finally {
      if (mounted) {
        setState(() => _isRecordingReturn = false);
      }
    }
  }

  Future<void> _pickImage(ImageSource source) async {
    try {
      final pickedFile = await _picker.pickImage(
        source: source,
        maxWidth: 800,
        maxHeight: 800,
        imageQuality: 70,
      );

      if (pickedFile != null) {
        final bytes = await File(pickedFile.path).readAsBytes();
        final base64String = base64Encode(bytes);

        setState(() {
          _imageFile = File(pickedFile.path);
          _imageBase64 = base64String;
        });
      }
    } catch (e) {
      if (mounted) {
        ScaffoldMessenger.of(context).showSnackBar(
          SnackBar(content: Text("Gagal mengambil foto: $e"), backgroundColor: const Color(0xFFEF4444)),
        );
      }
    }
  }

  void _openSearchableEmployeePicker() {
    showModalBottomSheet(
      context: context,
      isScrollControlled: true,
      backgroundColor: Colors.transparent,
      builder: (ctx) {
        return StatefulBuilder(
          builder: (context, setModalState) {
            return DraggableScrollableSheet(
              initialChildSize: 0.75,
              minChildSize: 0.4,
              maxChildSize: 0.95,
              builder: (_, scrollController) {
                return Container(
                  decoration: const BoxDecoration(
                    color: Colors.white,
                    borderRadius: BorderRadius.vertical(top: Radius.circular(24)),
                  ),
                  child: Column(
                    children: [
                      Container(
                        margin: const EdgeInsets.only(top: 12, bottom: 8),
                        width: 40,
                        height: 4,
                        decoration: BoxDecoration(
                          color: const Color(0xFFCBD5E1),
                          borderRadius: BorderRadius.circular(2),
                        ),
                      ),
                      Padding(
                        padding: const EdgeInsets.symmetric(horizontal: 20, vertical: 8),
                        child: Row(
                          mainAxisAlignment: MainAxisAlignment.spaceBetween,
                          children: [
                            const Text(
                              "Pilih Pegawai Terlapor",
                              style: TextStyle(
                                fontSize: 16,
                                fontWeight: FontWeight.bold,
                                color: Color(0xFF1E293B),
                              ),
                            ),
                            IconButton(
                              icon: const Icon(Icons.close_rounded, color: Color(0xFF64748B)),
                              onPressed: () => Navigator.pop(ctx),
                            ),
                          ],
                        ),
                      ),
                      Padding(
                        padding: const EdgeInsets.symmetric(horizontal: 18, vertical: 4),
                        child: TextField(
                          autofocus: false,
                          decoration: InputDecoration(
                            hintText: "Cari nama pegawai...",
                            hintStyle: const TextStyle(fontSize: 13, color: Color(0xFF94A3B8)),
                            prefixIcon: const Icon(Icons.search_rounded, color: Color(0xFFEA580C)),
                            filled: true,
                            fillColor: const Color(0xFFF8FAFC),
                            contentPadding: const EdgeInsets.symmetric(vertical: 10),
                            border: OutlineInputBorder(
                              borderRadius: BorderRadius.circular(14),
                              borderSide: const BorderSide(color: Color(0xFFE2E8F0)),
                            ),
                            enabledBorder: OutlineInputBorder(
                              borderRadius: BorderRadius.circular(14),
                              borderSide: const BorderSide(color: Color(0xFFE2E8F0)),
                            ),
                          ),
                          onChanged: (query) {
                            setModalState(() {
                              _searchFilter = query.toLowerCase();
                            });
                          },
                        ),
                      ),
                      const SizedBox(height: 8),
                      Expanded(
                        child: _employees.isEmpty
                            ? const Center(child: CircularProgressIndicator())
                            : Builder(builder: (context) {
                                final filtered = _employees.where((e) {
                                  if (_searchFilter.isEmpty) return true;
                                  return e.name.toLowerCase().contains(_searchFilter);
                                }).toList();

                                if (filtered.isEmpty) {
                                  return const Center(
                                    child: Text(
                                      "Tidak ada pegawai yang cocok.",
                                      style: TextStyle(color: Color(0xFF64748B), fontSize: 13),
                                    ),
                                  );
                                }

                                return ListView.separated(
                                  controller: scrollController,
                                  padding: const EdgeInsets.symmetric(horizontal: 18, vertical: 8),
                                  itemCount: filtered.length,
                                  separatorBuilder: (_, __) => const Divider(height: 1, color: Color(0xFFF1F5F9)),
                                  itemBuilder: (context, idx) {
                                    final emp = filtered[idx];
                                    final isSelected = _selectedTarget?.uid == emp.uid;
                                    return ListTile(
                                      contentPadding: const EdgeInsets.symmetric(horizontal: 10, vertical: 4),
                                      leading: CircleAvatar(
                                        radius: 20,
                                        backgroundColor: isSelected ? const Color(0xFFEA580C) : const Color(0xFFFFEDD5),
                                        child: Text(
                                          emp.name.isNotEmpty ? emp.name[0].toUpperCase() : 'U',
                                          style: TextStyle(
                                            color: isSelected ? Colors.white : const Color(0xFFEA580C),
                                            fontWeight: FontWeight.bold,
                                            fontSize: 14,
                                          ),
                                        ),
                                      ),
                                      title: Text(
                                        emp.name,
                                        style: TextStyle(
                                          fontSize: 14,
                                          fontWeight: isSelected ? FontWeight.bold : FontWeight.w600,
                                          color: isSelected ? const Color(0xFFEA580C) : const Color(0xFF1E293B),
                                        ),
                                      ),
                                      trailing: isSelected
                                          ? const Icon(Icons.check_circle_rounded, color: Color(0xFFEA580C))
                                          : null,
                                      onTap: () {
                                        setState(() {
                                          _selectedTarget = emp;
                                        });
                                        Navigator.pop(ctx);
                                      },
                                    );
                                  },
                                );
                              }),
                      ),
                    ],
                  ),
                );
              },
            );
          },
        );
      },
    );
  }

  Future<void> _handleSubmit() async {
    if (_selectedTarget == null) {
      ScaffoldMessenger.of(context).showSnackBar(
        const SnackBar(content: Text("⚠️ Harap pilih nama pegawai yang dilaporkan!"), backgroundColor: Color(0xFFEF4444)),
      );
      return;
    }

    if (_descController.text.trim().isEmpty) {
      ScaffoldMessenger.of(context).showSnackBar(
        const SnackBar(content: Text("⚠️ Harap tuliskan keterangan laporan!"), backgroundColor: Color(0xFFEF4444)),
      );
      return;
    }

    if (_imageBase64 == null) {
      ScaffoldMessenger.of(context).showSnackBar(
        const SnackBar(
          content: Text("⚠️ Wajib melampirkan Dokumen Pendukung (Foto Bukti) sebelum mengirim laporan!"),
          backgroundColor: Color(0xFFEF4444),
        ),
      );
      return;
    }

    setState(() => _isSubmitting = true);

    try {
      await _cepuService.createCepuReport(
        reporter: widget.reporter,
        targetUser: _selectedTarget!,
        description: _descController.text.trim(),
        startTime: _startTime,
        endTime: null,
        photoBase64: _imageBase64,
      );

      if (mounted) {
        ScaffoldMessenger.of(context).showSnackBar(
          const SnackBar(
            content: Text("✅ Laporan Cepu berhasil dikirim! Menunggu minimal 4 verifikasi rekan."),
            backgroundColor: Color(0xFF10B981),
          ),
        );
        Navigator.pop(context, true);
      }
    } catch (e) {
      if (mounted) {
        ScaffoldMessenger.of(context).showSnackBar(
          SnackBar(content: Text(e.toString()), backgroundColor: const Color(0xFFEF4444)),
        );
      }
    } finally {
      if (mounted) {
        setState(() => _isSubmitting = false);
      }
    }
  }

  @override
  Widget build(BuildContext context) {
    final officeLatLng = LatLng(widget.office.latitude, widget.office.longitude);
    final userLatLng = _currentPosition != null
        ? LatLng(_currentPosition!.latitude, _currentPosition!.longitude)
        : officeLatLng;

    final startTimeStr = DateFormat('HH:mm').format(_startTime);

    return Scaffold(
      backgroundColor: const Color(0xFFF8FAFC),
      appBar: AppBar(
        backgroundColor: const Color(0xFFEA580C),
        elevation: 0,
        leading: IconButton(
          icon: const Icon(Icons.arrow_back_ios_new, color: Colors.white, size: 20),
          onPressed: () => Navigator.pop(context),
        ),
        title: const Text(
          "Cepu",
          style: TextStyle(color: Colors.white, fontWeight: FontWeight.bold, fontSize: 18),
        ),
        centerTitle: true,
        actions: [
          IconButton(
            icon: _isLoadingLocation
                ? const SizedBox(
                    width: 18,
                    height: 18,
                    child: CircularProgressIndicator(color: Colors.white, strokeWidth: 2),
                  )
                : const Icon(Icons.refresh_rounded, color: Colors.white),
            tooltip: "Perbarui Lokasi GPS",
            onPressed: _isLoadingLocation ? null : _fetchCurrentLocation,
          ),
        ],
      ),
      body: Column(
        children: [
          // 1. PETA INTERAKTIF + RADIUS KANTOR 50M
          Expanded(
            flex: 4,
            child: Stack(
              children: [
                FlutterMap(
                  mapController: _mapController,
                  options: MapOptions(
                    initialCenter: officeLatLng,
                    initialZoom: 17.5,
                    minZoom: 14.0,
                    maxZoom: 19.0,
                  ),
                  children: [
                    TileLayer(
                      urlTemplate: 'https://tile.openstreetmap.org/{z}/{x}/{y}.png',
                      userAgentPackageName: 'com.example.presensi_app',
                    ),
                    CircleLayer(
                      circles: [
                        CircleMarker(
                          point: officeLatLng,
                          color: const Color(0xFFEA580C).withOpacity(0.15),
                          borderColor: const Color(0xFFEA580C),
                          borderStrokeWidth: 2,
                          useRadiusInMeter: true,
                          radius: widget.office.radiusMeters,
                        ),
                      ],
                    ),
                    MarkerLayer(
                      markers: [
                        Marker(
                          point: officeLatLng,
                          width: 42,
                          height: 42,
                          alignment: Alignment.center,
                          child: Container(
                            decoration: BoxDecoration(
                              color: const Color(0xFF0F172A),
                              shape: BoxShape.circle,
                              boxShadow: [
                                BoxShadow(
                                  color: Colors.black.withOpacity(0.25),
                                  blurRadius: 8,
                                  offset: const Offset(0, 4),
                                ),
                              ],
                            ),
                            child: const Icon(Icons.business_rounded, color: Colors.white, size: 22),
                          ),
                        ),
                        if (_currentPosition != null)
                          Marker(
                            point: userLatLng,
                            width: 38,
                            height: 38,
                            alignment: Alignment.center,
                            child: Stack(
                              alignment: Alignment.center,
                              children: [
                                Container(
                                  width: 38,
                                  height: 38,
                                  decoration: BoxDecoration(
                                    color: const Color(0xFFEA580C).withOpacity(0.25),
                                    shape: BoxShape.circle,
                                  ),
                                ),
                                Container(
                                  width: 26,
                                  height: 26,
                                  decoration: BoxDecoration(
                                    color: const Color(0xFFEA580C),
                                    shape: BoxShape.circle,
                                    border: Border.all(color: Colors.white, width: 2.5),
                                    boxShadow: [
                                      BoxShadow(
                                        color: Colors.black.withOpacity(0.2),
                                        blurRadius: 6,
                                        offset: const Offset(0, 2),
                                      ),
                                    ],
                                  ),
                                  child: const Center(
                                    child: Icon(Icons.person_rounded, color: Colors.white, size: 15),
                                  ),
                                ),
                              ],
                            ),
                          ),
                      ],
                    ),
                  ],
                ),

                Positioned(
                  top: 12,
                  right: 12,
                  child: Column(
                    children: [
                      _buildFloatingMapButton(
                        icon: Icons.apartment_rounded,
                        tooltip: "Pusatkan ke Kantor",
                        onTap: () => _mapController.move(officeLatLng, 17.5),
                      ),
                      const SizedBox(height: 6),
                      _buildFloatingMapButton(
                        icon: Icons.my_location_rounded,
                        tooltip: "Pusatkan ke Lokasi Saya",
                        onTap: () {
                          if (_currentPosition != null) {
                            _mapController.move(userLatLng, 18.0);
                          }
                        },
                      ),
                    ],
                  ),
                ),
              ],
            ),
          ),

          // 2. FORMULIR CEPU & REKAM WAKTU KEMBALI
          Expanded(
            flex: 6,
            child: Container(
              decoration: const BoxDecoration(
                color: Colors.white,
                borderRadius: BorderRadius.vertical(top: Radius.circular(24)),
                boxShadow: [
                  BoxShadow(
                    color: Color(0x14000000),
                    blurRadius: 16,
                    offset: Offset(0, -4),
                  ),
                ],
              ),
              child: SingleChildScrollView(
                padding: const EdgeInsets.symmetric(horizontal: 20, vertical: 16),
                child: Column(
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

                    const SizedBox(height: 12),

                    // Jika user yang login sedang dilaporkan Cepu & sudah terverifikasi:
                    if (_activeCepuForMe != null) ...[
                      Container(
                        padding: const EdgeInsets.all(14),
                        decoration: BoxDecoration(
                          color: const Color(0xFFFEF2F2),
                          borderRadius: BorderRadius.circular(16),
                          border: Border.all(color: const Color(0xFFFCA5A5), width: 1.5),
                        ),
                        child: Column(
                          crossAxisAlignment: CrossAxisAlignment.start,
                          children: [
                            Row(
                              children: const [
                                Icon(Icons.warning_amber_rounded, color: Color(0xFFDC2626), size: 20),
                                SizedBox(width: 8),
                                Expanded(
                                  child: Text(
                                    "Anda Dilaporkan Tidak di Kantor",
                                    style: TextStyle(
                                      color: Color(0xFF991B1B),
                                      fontWeight: FontWeight.bold,
                                      fontSize: 13,
                                    ),
                                  ),
                                ),
                              ],
                            ),
                            const SizedBox(height: 6),
                            Text(
                              "Laporan telah diverifikasi oleh ${_activeCepuForMe!.verificationCount} rekan kerja. Segera rekam waktu kembali setelah Anda berada di kantor.",
                              style: const TextStyle(fontSize: 11.5, color: Color(0xFF7F1D1D), height: 1.3),
                            ),
                            const SizedBox(height: 12),
                            SizedBox(
                              width: double.infinity,
                              height: 44,
                              child: ElevatedButton.icon(
                                onPressed: (_isInRadius && !_isMocked && !_isRecordingReturn)
                                    ? _handleRecordReturnForMe
                                    : null,
                                icon: const Icon(Icons.check_circle_rounded, color: Colors.white, size: 18),
                                label: _isRecordingReturn
                                    ? const SizedBox(width: 18, height: 18, child: CircularProgressIndicator(color: Colors.white, strokeWidth: 2))
                                    : const Text("REKAM WAKTU KEMBALI SAYA", style: TextStyle(color: Colors.white, fontWeight: FontWeight.bold, fontSize: 12.5)),
                                style: ElevatedButton.styleFrom(
                                  backgroundColor: const Color(0xFF10B981),
                                  disabledBackgroundColor: const Color(0xFFCBD5E1),
                                  shape: RoundedRectangleBorder(borderRadius: BorderRadius.circular(12)),
                                ),
                              ),
                            ),
                          ],
                        ),
                      ),
                      const SizedBox(height: 16),
                    ],

                    // Status Bar Ringkas
                    Container(
                      padding: const EdgeInsets.symmetric(horizontal: 14, vertical: 8),
                      decoration: BoxDecoration(
                        color: const Color(0xFFFFF7ED),
                        borderRadius: BorderRadius.circular(12),
                        border: Border.all(color: const Color(0xFFFED7AA)),
                      ),
                      child: Row(
                        children: [
                          const Icon(Icons.location_pin, color: Color(0xFFEA580C), size: 16),
                          const SizedBox(width: 6),
                          Expanded(
                            child: Text(
                              "Radius Kantor: ${widget.office.radiusMeters.toInt()}m  •  Jarak Anda: ${_distanceMeters.toStringAsFixed(1)}m",
                              style: const TextStyle(fontSize: 11.5, fontWeight: FontWeight.w600, color: Color(0xFF9A3412)),
                            ),
                          ),
                        ],
                      ),
                    ),

                    const SizedBox(height: 14),

                    // A. PILIH PEGAWAI
                    const Text(
                      "Pegawai yang Dilaporkan *",
                      style: TextStyle(fontSize: 12.5, fontWeight: FontWeight.bold, color: Color(0xFF1E293B)),
                    ),
                    const SizedBox(height: 6),
                    InkWell(
                      onTap: _openSearchableEmployeePicker,
                      borderRadius: BorderRadius.circular(14),
                      child: Container(
                        padding: const EdgeInsets.symmetric(horizontal: 14, vertical: 12),
                        decoration: BoxDecoration(
                          color: const Color(0xFFF8FAFC),
                          borderRadius: BorderRadius.circular(14),
                          border: Border.all(
                            color: _selectedTarget != null ? const Color(0xFFEA580C) : const Color(0xFFCBD5E1),
                            width: _selectedTarget != null ? 1.5 : 1.0,
                          ),
                        ),
                        child: Row(
                          children: [
                            if (_selectedTarget != null)
                              CircleAvatar(
                                radius: 14,
                                backgroundColor: const Color(0xFFFFEDD5),
                                child: Text(
                                  _selectedTarget!.name[0].toUpperCase(),
                                  style: const TextStyle(color: Color(0xFFEA580C), fontWeight: FontWeight.bold, fontSize: 12),
                                ),
                              )
                            else
                              const Icon(Icons.person_search_rounded, color: Color(0xFF94A3B8), size: 22),
                            const SizedBox(width: 10),
                            Expanded(
                              child: Text(
                                _selectedTarget != null
                                    ? _selectedTarget!.name
                                    : "Pilih dan cari nama pegawai...",
                                style: TextStyle(
                                  fontSize: 13,
                                  fontWeight: _selectedTarget != null ? FontWeight.bold : FontWeight.normal,
                                  color: _selectedTarget != null ? const Color(0xFF1E293B) : const Color(0xFF94A3B8),
                                ),
                                maxLines: 1,
                                overflow: TextOverflow.ellipsis,
                              ),
                            ),
                            const Icon(Icons.keyboard_arrow_down_rounded, color: Color(0xFF64748B)),
                          ],
                        ),
                      ),
                    ),

                    const SizedBox(height: 14),

                    // B. WAKTU TERPANTAU TIDAK DI KANTOR
                    Row(
                      children: [
                        Expanded(
                          child: Column(
                            crossAxisAlignment: CrossAxisAlignment.start,
                            children: [
                              const Text("Waktu Mulai *", style: TextStyle(fontSize: 12, fontWeight: FontWeight.bold, color: Color(0xFF1E293B))),
                              const SizedBox(height: 6),
                              Container(
                                padding: const EdgeInsets.symmetric(horizontal: 12, vertical: 10),
                                decoration: BoxDecoration(
                                  color: const Color(0xFFF8FAFC),
                                  borderRadius: BorderRadius.circular(12),
                                  border: Border.all(color: const Color(0xFFE2E8F0)),
                                ),
                                child: Row(
                                  children: [
                                    const Icon(Icons.access_time_filled_rounded, size: 16, color: Color(0xFFEA580C)),
                                    const SizedBox(width: 6),
                                    Text("$startTimeStr WIT", style: const TextStyle(fontSize: 13, fontWeight: FontWeight.bold, color: Color(0xFF1E293B))),
                                  ],
                                ),
                              ),
                            ],
                          ),
                        ),
                        const SizedBox(width: 10),
                        Expanded(
                          child: Column(
                            crossAxisAlignment: CrossAxisAlignment.start,
                            children: [
                              const Text("Waktu Kembali", style: TextStyle(fontSize: 12, fontWeight: FontWeight.bold, color: Color(0xFF1E293B))),
                              const SizedBox(height: 6),
                              Container(
                                padding: const EdgeInsets.symmetric(horizontal: 12, vertical: 10),
                                decoration: BoxDecoration(
                                  color: const Color(0xFFF8FAFC),
                                  borderRadius: BorderRadius.circular(12),
                                  border: Border.all(color: const Color(0xFFE2E8F0)),
                                ),
                                child: Row(
                                  children: const [
                                    Icon(Icons.check_circle_rounded, size: 16, color: Color(0xFF94A3B8)),
                                    SizedBox(width: 6),
                                    Text("Belum Kembali", style: TextStyle(fontSize: 12, fontWeight: FontWeight.bold, color: Color(0xFF64748B))),
                                  ],
                                ),
                              ),
                            ],
                          ),
                        ),
                      ],
                    ),

                    const SizedBox(height: 14),

                    // C. KETERANGAN
                    const Text(
                      "Keterangan / Alasan Laporan *",
                      style: TextStyle(fontSize: 12.5, fontWeight: FontWeight.bold, color: Color(0xFF1E293B)),
                    ),
                    const SizedBox(height: 6),
                    TextField(
                      controller: _descController,
                      maxLines: 2,
                      decoration: InputDecoration(
                        hintText: "Tuliskan keterangan detail di sini...",
                        hintStyle: const TextStyle(fontSize: 12, color: Color(0xFF94A3B8)),
                        filled: true,
                        fillColor: const Color(0xFFF8FAFC),
                        contentPadding: const EdgeInsets.all(12),
                        border: OutlineInputBorder(borderRadius: BorderRadius.circular(12), borderSide: const BorderSide(color: Color(0xFFE2E8F0))),
                        enabledBorder: OutlineInputBorder(borderRadius: BorderRadius.circular(12), borderSide: const BorderSide(color: Color(0xFFE2E8F0))),
                        focusedBorder: OutlineInputBorder(borderRadius: BorderRadius.circular(12), borderSide: const BorderSide(color: Color(0xFFEA580C), width: 1.5)),
                      ),
                    ),

                    const SizedBox(height: 14),

                    // D. BUKTI PENDUKUNG (FOTO) - WAJIB
                    Row(
                      mainAxisAlignment: MainAxisAlignment.spaceBetween,
                      children: const [
                        Text("Bukti Pendukung (Foto) *", style: TextStyle(fontSize: 12.5, fontWeight: FontWeight.bold, color: Color(0xFF1E293B))),
                        Text("(Wajib terlampir)", style: TextStyle(fontSize: 11, fontWeight: FontWeight.w600, color: Color(0xFFEA580C))),
                      ],
                    ),
                    const SizedBox(height: 6),

                    if (_imageFile != null) ...[
                      Stack(
                        children: [
                          ClipRRect(
                            borderRadius: BorderRadius.circular(14),
                            child: Image.file(_imageFile!, height: 130, width: double.infinity, fit: BoxFit.cover),
                          ),
                          Positioned(
                            top: 6,
                            right: 6,
                            child: InkWell(
                              onTap: () => setState(() {
                                _imageFile = null;
                                _imageBase64 = null;
                              }),
                              child: Container(
                                padding: const EdgeInsets.all(4),
                                decoration: const BoxDecoration(color: Colors.black54, shape: BoxShape.circle),
                                child: const Icon(Icons.close, color: Colors.white, size: 16),
                              ),
                            ),
                          ),
                          Positioned(
                            bottom: 6,
                            left: 6,
                            child: Container(
                              padding: const EdgeInsets.symmetric(horizontal: 8, vertical: 3),
                              decoration: BoxDecoration(color: Colors.black.withOpacity(0.65), borderRadius: BorderRadius.circular(6)),
                              child: const Text("✓ Foto Terlampir & Terkompresi", style: TextStyle(color: Colors.white, fontSize: 10, fontWeight: FontWeight.bold)),
                            ),
                          ),
                        ],
                      ),
                    ] else ...[
                      Row(
                        children: [
                          Expanded(
                            child: OutlinedButton.icon(
                              onPressed: () => _pickImage(ImageSource.camera),
                              icon: const Icon(Icons.camera_alt_rounded, size: 17, color: Color(0xFFEA580C)),
                              label: const Text("Ambil Kamera", style: TextStyle(fontSize: 12, color: Color(0xFFEA580C), fontWeight: FontWeight.bold)),
                              style: OutlinedButton.styleFrom(
                                padding: const EdgeInsets.symmetric(vertical: 10),
                                side: const BorderSide(color: Color(0xFFFDBA74)),
                                shape: RoundedRectangleBorder(borderRadius: BorderRadius.circular(12)),
                              ),
                            ),
                          ),
                          const SizedBox(width: 8),
                          Expanded(
                            child: OutlinedButton.icon(
                              onPressed: () => _pickImage(ImageSource.gallery),
                              icon: const Icon(Icons.photo_library_rounded, size: 17, color: Color(0xFF475569)),
                              label: const Text("Pilih Galeri", style: TextStyle(fontSize: 12, color: Color(0xFF475569), fontWeight: FontWeight.bold)),
                              style: OutlinedButton.styleFrom(
                                padding: const EdgeInsets.symmetric(vertical: 10),
                                side: const BorderSide(color: Color(0xFFCBD5E1)),
                                shape: RoundedRectangleBorder(borderRadius: BorderRadius.circular(12)),
                              ),
                            ),
                          ),
                        ],
                      ),
                    ],

                    const SizedBox(height: 20),

                    // TOMBOL KIRIM LAPORAN (Requirement 5)
                    SizedBox(
                      height: 48,
                      child: ElevatedButton(
                        onPressed: _isSubmitting ? null : _handleSubmit,
                        style: ElevatedButton.styleFrom(
                          backgroundColor: const Color(0xFFEA580C),
                          elevation: 2,
                          shape: RoundedRectangleBorder(borderRadius: BorderRadius.circular(14)),
                        ),
                        child: _isSubmitting
                            ? const SizedBox(
                                width: 20,
                                height: 20,
                                child: CircularProgressIndicator(color: Colors.white, strokeWidth: 2.5),
                              )
                            : const Text(
                                "KIRIM LAPORAN",
                                style: TextStyle(color: Colors.white, fontWeight: FontWeight.bold, fontSize: 13, letterSpacing: 0.5),
                              ),
                      ),
                    ),
                  ],
                ),
              ),
            ),
          ),
        ],
      ),
    );
  }

  Widget _buildFloatingMapButton({
    required IconData icon,
    required String tooltip,
    required VoidCallback onTap,
  }) {
    return Container(
      decoration: BoxDecoration(
        color: Colors.white,
        borderRadius: BorderRadius.circular(10),
        boxShadow: [
          BoxShadow(
            color: Colors.black.withOpacity(0.12),
            blurRadius: 6,
            offset: const Offset(0, 2),
          ),
        ],
      ),
      child: IconButton(
        icon: Icon(icon, color: const Color(0xFF334155), size: 18),
        tooltip: tooltip,
        onPressed: onTap,
        constraints: const BoxConstraints(minWidth: 36, minHeight: 36),
        padding: EdgeInsets.zero,
      ),
    );
  }
}
