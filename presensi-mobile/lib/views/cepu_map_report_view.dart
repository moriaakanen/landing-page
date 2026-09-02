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
import '../core/utils/custom_toast.dart';

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
  Set<String> _activePermitUserIds = {};
  Set<String> _activeCepuUserIds = {};
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
    final permitIds = await _cepuService.getActivePermitUserIdsToday();
    final cepuIds = await _cepuService.getActiveCepuTargetUserIdsToday();

    if (mounted) {
      setState(() {
        _employees = list;
        _activePermitUserIds = permitIds;
        _activeCepuUserIds = cepuIds;
        _isLoadingEmployees = false;
      });
    }
  }

  bool _isEmpPermitted(UserModel emp) {
    return _activePermitUserIds.contains(emp.uid) ||
        _activePermitUserIds.contains('user_${emp.username.replaceAll('.', '_')}') ||
        _activePermitUserIds.contains(emp.username);
  }

  bool _isEmpReported(UserModel emp) {
    return _activeCepuUserIds.contains(emp.uid) ||
        _activeCepuUserIds.contains('user_${emp.username.replaceAll('.', '_')}') ||
        _activeCepuUserIds.contains(emp.username);
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
      AppToast.showError(
        context,
        _isMocked
            ? "Terdeteksi Fake GPS! Harap matikan aplikasi lokasi palsu."
            : "Anda berada di luar radius kantor (${_distanceMeters.toStringAsFixed(1)}m). Waktu kembali hanya bisa direkam di dalam area kantor (≤${widget.office.radiusMeters.toInt()}m).",
        title: "Gagal Merekam Kembali",
      );
      return;
    }

    setState(() => _isRecordingReturn = true);

    try {
      await _cepuService.recordCepuReturnTime(
        cepuId: _activeCepuForMe!.id,
        user: widget.reporter,
        office: widget.office,
      );

      if (mounted) {
        AppToast.showSuccess(
          context,
          "Waktu kembali Anda telah berhasil terekam di sistem!",
          title: "Berhasil Kembali",
        );
        setState(() {
          _activeCepuForMe = null;
        });
        Navigator.pop(context, true);
      }
    } catch (e) {
      if (mounted) {
        AppToast.showError(context, e.toString(), title: "Gagal Merekam");
      }
    } finally {
      if (mounted) {
        setState(() => _isRecordingReturn = false);
      }
    }
  }

  Future<void> _pickImage(ImageSource source) async {
    try {
      final picked = await _picker.pickImage(
        source: source,
        maxWidth: 1024,
        maxHeight: 1024,
        imageQuality: 70,
      );

      if (picked != null) {
        final bytes = await File(picked.path).readAsBytes();
        setState(() {
          _imageFile = File(picked.path);
          _imageBase64 = base64Encode(bytes);
        });
      }
    } catch (e) {
      if (mounted) {
        AppToast.showError(context, "Gagal mengambil foto: $e", title: "Foto Gagal");
      }
    }
  }

  void _showImageSourceDialog() {
    showModalBottomSheet(
      context: context,
      backgroundColor: Colors.transparent,
      builder: (ctx) => Container(
        decoration: const BoxDecoration(
          color: Colors.white,
          borderRadius: BorderRadius.vertical(top: Radius.circular(24)),
        ),
        padding: const EdgeInsets.symmetric(vertical: 20, horizontal: 16),
        child: Column(
          mainAxisSize: MainAxisSize.min,
          children: [
            Container(
              width: 40,
              height: 4,
              decoration: BoxDecoration(
                color: const Color(0xFFCBD5E1),
                borderRadius: BorderRadius.circular(2),
              ),
            ),
            const SizedBox(height: 16),
            const Text(
              "Ambil Dokumen Pendukung (Foto)",
              style: TextStyle(
                fontSize: 16,
                fontWeight: FontWeight.bold,
                color: Color(0xFF1E293B),
              ),
            ),
            const SizedBox(height: 16),
            ListTile(
              leading: Container(
                padding: const EdgeInsets.all(10),
                decoration: const BoxDecoration(
                  color: Color(0xFFEFF6FF),
                  shape: BoxShape.circle,
                ),
                child: const Icon(Icons.camera_alt_rounded, color: Color(0xFF2563EB)),
              ),
              title: const Text("Gunakan Kamera", style: TextStyle(fontWeight: FontWeight.w600)),
              onTap: () {
                Navigator.pop(ctx);
                _pickImage(ImageSource.camera);
              },
            ),
            ListTile(
              leading: Container(
                padding: const EdgeInsets.all(10),
                decoration: const BoxDecoration(
                  color: Color(0xFFFFF7ED),
                  shape: BoxShape.circle,
                ),
                child: const Icon(Icons.photo_library_rounded, color: Color(0xFFEA580C)),
              ),
              title: const Text("Pilih dari Galeri", style: TextStyle(fontWeight: FontWeight.w600)),
              onTap: () {
                Navigator.pop(ctx);
                _pickImage(ImageSource.gallery);
              },
            ),
          ],
        ),
      ),
    );
  }

  Future<void> _pickStartTime() async {
    final now = TimeOfDay.fromDateTime(_startTime);
    final picked = await showTimePicker(
      context: context,
      initialTime: now,
      helpText: "PILIH WAKTU MULAI TIDAK BERADA DI KANTOR",
    );

    if (picked != null) {
      final nowDt = DateTime.now();
      setState(() {
        _startTime = DateTime(
          nowDt.year,
          nowDt.month,
          nowDt.day,
          picked.hour,
          picked.minute,
        );
      });
    }
  }

  void _showEmployeeSelector() {
    showModalBottomSheet(
      context: context,
      isScrollControlled: true,
      backgroundColor: Colors.transparent,
      builder: (ctx) {
        return StatefulBuilder(
          builder: (context, setModalState) {
            return DraggableScrollableSheet(
              initialChildSize: 0.75,
              maxChildSize: 0.92,
              minChildSize: 0.4,
              builder: (_, scrollController) {
                return Container(
                  decoration: const BoxDecoration(
                    color: Colors.white,
                    borderRadius: BorderRadius.vertical(top: Radius.circular(28)),
                    boxShadow: [
                      BoxShadow(color: Colors.black26, blurRadius: 20, offset: Offset(0, -4)),
                    ],
                  ),
                  child: Column(
                    children: [
                      Container(
                        margin: const EdgeInsets.only(top: 12, bottom: 8),
                        width: 44,
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
                                fontSize: 16.5,
                                fontWeight: FontWeight.w800,
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
                            ? const Center(child: CircularProgressIndicator(color: Color(0xFFEA580C)))
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
                                    final isPermitted = _isEmpPermitted(emp);
                                    final isReported = _isEmpReported(emp);
                                    final isDisabled = isPermitted || isReported;

                                    return ListTile(
                                      contentPadding: const EdgeInsets.symmetric(horizontal: 10, vertical: 4),
                                      enabled: !isDisabled,
                                      leading: CircleAvatar(
                                        radius: 20,
                                        backgroundColor: isDisabled
                                            ? const Color(0xFFF1F5F9)
                                            : (isSelected ? const Color(0xFFEA580C) : const Color(0xFFFFEDD5)),
                                        child: Text(
                                          emp.name.isNotEmpty ? emp.name[0].toUpperCase() : 'U',
                                          style: TextStyle(
                                            color: isDisabled
                                                ? const Color(0xFF94A3B8)
                                                : (isSelected ? Colors.white : const Color(0xFFEA580C)),
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
                                          color: isDisabled
                                              ? const Color(0xFF94A3B8)
                                              : (isSelected ? const Color(0xFFEA580C) : const Color(0xFF1E293B)),
                                        ),
                                      ),
                                      subtitle: Text(
                                        "@${emp.username}",
                                        style: TextStyle(
                                          fontSize: 11.5,
                                          color: isDisabled ? const Color(0xFFCBD5E1) : const Color(0xFF64748B),
                                        ),
                                      ),
                                      trailing: isPermitted
                                          ? Container(
                                              padding: const EdgeInsets.symmetric(horizontal: 8, vertical: 4),
                                              decoration: BoxDecoration(
                                                color: const Color(0xFFEFF6FF),
                                                borderRadius: BorderRadius.circular(8),
                                                border: Border.all(color: const Color(0xFFBFDBFE)),
                                              ),
                                              child: const Text(
                                                "Sedang Izin",
                                                style: TextStyle(
                                                  fontSize: 10.5,
                                                  fontWeight: FontWeight.bold,
                                                  color: Color(0xFF2563EB),
                                                ),
                                              ),
                                            )
                                          : (isReported
                                              ? Container(
                                                  padding: const EdgeInsets.symmetric(horizontal: 8, vertical: 4),
                                                  decoration: BoxDecoration(
                                                    color: const Color(0xFFFEF2F2),
                                                    borderRadius: BorderRadius.circular(8),
                                                    border: Border.all(color: const Color(0xFFFECACA)),
                                                  ),
                                                  child: const Text(
                                                    "Sudah Dilaporkan",
                                                    style: TextStyle(
                                                      fontSize: 10.5,
                                                      fontWeight: FontWeight.bold,
                                                      color: Color(0xFFDC2626),
                                                    ),
                                                  ),
                                                )
                                              : (isSelected
                                                  ? const Icon(Icons.check_circle_rounded, color: Color(0xFFEA580C))
                                                  : null)),
                                      onTap: isDisabled
                                          ? null
                                          : () {
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
      AppToast.showWarning(context, "Harap pilih nama pegawai yang dilaporkan terlebih dahulu!", title: "Pegawai Belum Dipilih");
      return;
    }

    if (_descController.text.trim().isEmpty) {
      AppToast.showWarning(context, "Harap tuliskan keterangan atau alasan laporan!", title: "Keterangan Kosong");
      return;
    }

    if (_imageBase64 == null) {
      AppToast.showWarning(context, "Wajib melampirkan Dokumen Pendukung (Foto Bukti) sebelum mengirim laporan!", title: "Foto Bukti Wajib");
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
        AppToast.showSuccess(
          context,
          "Laporan Cepu berhasil dikirim! Menunggu minimal 4 verifikasi rekan.",
          title: "Laporan Terkirim",
        );
        Navigator.pop(context, true);
      }
    } catch (e) {
      if (mounted) {
        AppToast.showError(context, e.toString(), title: "Gagal Mengirim");
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
                          radius: widget.office.radiusMeters,
                          useRadiusInMeter: true,
                        ),
                      ],
                    ),
                    MarkerLayer(
                      markers: [
                        Marker(
                          point: officeLatLng,
                          width: 44,
                          height: 44,
                          alignment: Alignment.center,
                          child: Container(
                            decoration: BoxDecoration(
                              color: const Color(0xFFEA580C),
                              shape: BoxShape.circle,
                              boxShadow: [
                                BoxShadow(
                                  color: const Color(0xFFEA580C).withOpacity(0.4),
                                  blurRadius: 10,
                                  offset: const Offset(0, 4),
                                ),
                              ],
                            ),
                            child: const Icon(Icons.business_rounded, color: Colors.white, size: 24),
                          ),
                        ),
                        if (_currentPosition != null)
                          Marker(
                            point: userLatLng,
                            width: 36,
                            height: 36,
                            alignment: Alignment.center,
                            child: Container(
                              decoration: BoxDecoration(
                                color: _isInRadius ? const Color(0xFF10B981) : const Color(0xFFEF4444),
                                shape: BoxShape.circle,
                                border: Border.all(color: Colors.white, width: 3),
                                boxShadow: [
                                  BoxShadow(
                                    color: Colors.black.withOpacity(0.25),
                                    blurRadius: 8,
                                    offset: const Offset(0, 3),
                                  ),
                                ],
                              ),
                              child: const Icon(Icons.my_location_rounded, color: Colors.white, size: 18),
                            ),
                          ),
                      ],
                    ),
                  ],
                ),

                // Floating Map Controls
                Positioned(
                  top: 14,
                  right: 14,
                  child: Column(
                    children: [
                      _buildFloatingMapButton(
                        icon: Icons.add_rounded,
                        tooltip: "Zoom In",
                        onTap: () {
                          _mapController.move(
                            _mapController.camera.center,
                            _mapController.camera.zoom + 1,
                          );
                        },
                      ),
                      const SizedBox(height: 8),
                      _buildFloatingMapButton(
                        icon: Icons.remove_rounded,
                        tooltip: "Zoom Out",
                        onTap: () {
                          _mapController.move(
                            _mapController.camera.center,
                            _mapController.camera.zoom - 1,
                          );
                        },
                      ),
                      const SizedBox(height: 8),
                      _buildFloatingMapButton(
                        icon: Icons.my_location_rounded,
                        tooltip: "Posisi Saya",
                        onTap: () {
                          if (_currentPosition != null) {
                            _mapController.move(userLatLng, 18.0);
                          }
                        },
                      ),
                      const SizedBox(height: 8),
                      _buildFloatingMapButton(
                        icon: Icons.business_rounded,
                        tooltip: "Pusat Kantor",
                        onTap: () {
                          _mapController.move(officeLatLng, 17.5);
                        },
                      ),
                    ],
                  ),
                ),

                // Top GPS Status Pill
                Positioned(
                  top: 14,
                  left: 14,
                  child: Container(
                    padding: const EdgeInsets.symmetric(horizontal: 12, vertical: 7),
                    decoration: BoxDecoration(
                      color: Colors.white,
                      borderRadius: BorderRadius.circular(20),
                      boxShadow: [
                        BoxShadow(
                          color: Colors.black.withOpacity(0.12),
                          blurRadius: 8,
                          offset: const Offset(0, 3),
                        ),
                      ],
                    ),
                    child: Row(
                      mainAxisSize: MainAxisSize.min,
                      children: [
                        Container(
                          width: 9,
                          height: 9,
                          decoration: BoxDecoration(
                            shape: BoxShape.circle,
                            color: _isInRadius ? const Color(0xFF10B981) : const Color(0xFFEF4444),
                          ),
                        ),
                        const SizedBox(width: 7),
                        Text(
                          _isLoadingLocation
                              ? "Mencari GPS..."
                              : (_isInRadius
                                  ? "Di Kantor (${_distanceMeters.toStringAsFixed(0)}m)"
                                  : "Di Luar Kantor (${_distanceMeters.toStringAsFixed(0)}m)"),
                          style: TextStyle(
                            fontSize: 11.5,
                            fontWeight: FontWeight.bold,
                            color: _isInRadius ? const Color(0xFF047857) : const Color(0xFFB91C1C),
                          ),
                        ),
                      ],
                    ),
                  ),
                ),
              ],
            ),
          ),

          // 2. FORMULIR PELAPORAN CEPU
          Expanded(
            flex: 6,
            child: Container(
              width: double.infinity,
              decoration: const BoxDecoration(
                color: Colors.white,
                borderRadius: BorderRadius.vertical(top: Radius.circular(28)),
                boxShadow: [
                  BoxShadow(color: Colors.black12, blurRadius: 16, offset: Offset(0, -4)),
                ],
              ),
              child: SingleChildScrollView(
                padding: const EdgeInsets.fromLTRB(20, 16, 20, 24),
                child: Column(
                  crossAxisAlignment: CrossAxisAlignment.stretch,
                  children: [
                    // Banner Khusus jika user yang login dilaporkan di Cepu terverifikasi
                    if (_activeCepuForMe != null) ...[
                      Container(
                        padding: const EdgeInsets.all(14),
                        decoration: BoxDecoration(
                          color: const Color(0xFFFEF2F2),
                          borderRadius: BorderRadius.circular(16),
                          border: Border.all(color: const Color(0xFFFCA5A5)),
                        ),
                        child: Column(
                          crossAxisAlignment: CrossAxisAlignment.stretch,
                          children: [
                            Row(
                              children: const [
                                Icon(Icons.warning_amber_rounded, color: Color(0xFFDC2626), size: 22),
                                SizedBox(width: 8),
                                Expanded(
                                  child: Text(
                                    "Anda Dilaporkan Tidak di Kantor",
                                    style: TextStyle(fontSize: 13, fontWeight: FontWeight.bold, color: Color(0xFF991B1B)),
                                  ),
                                ),
                              ],
                            ),
                            const SizedBox(height: 6),
                            const Text(
                              "Laporan Cepu terhadap Anda telah terverifikasi. Segera rekam waktu kembali jika Anda sudah berada di kantor.",
                              style: TextStyle(fontSize: 11.5, color: Color(0xFF7F1D1D), height: 1.3),
                            ),
                            const SizedBox(height: 10),
                            SizedBox(
                              height: 38,
                              child: ElevatedButton.icon(
                                onPressed: _isRecordingReturn ? null : _handleRecordReturnForMe,
                                icon: const Icon(Icons.check_circle_rounded, size: 16, color: Colors.white),
                                label: _isRecordingReturn
                                    ? const SizedBox(
                                        width: 16,
                                        height: 16,
                                        child: CircularProgressIndicator(color: Colors.white, strokeWidth: 2),
                                      )
                                    : const Text(
                                        "REKAM WAKTU KEMBALI SAYA",
                                        style: TextStyle(fontSize: 12, fontWeight: FontWeight.bold, color: Colors.white),
                                      ),
                                style: ElevatedButton.styleFrom(
                                  backgroundColor: const Color(0xFFDC2626),
                                  shape: RoundedRectangleBorder(borderRadius: BorderRadius.circular(10)),
                                ),
                              ),
                            ),
                          ],
                        ),
                      ),
                      const SizedBox(height: 16),
                      const Divider(height: 1, color: Color(0xFFF1F5F9)),
                      const SizedBox(height: 14),
                    ],

                    // A. Pilih Nama Pegawai yang Dilaporkan
                    Row(
                      mainAxisAlignment: MainAxisAlignment.spaceBetween,
                      children: [
                        const Text(
                          "Pegawai yang Dilaporkan",
                          style: TextStyle(fontSize: 13, fontWeight: FontWeight.w800, color: Color(0xFF0F172A)),
                        ),
                        Container(
                          padding: const EdgeInsets.symmetric(horizontal: 8, vertical: 3),
                          decoration: BoxDecoration(
                            color: const Color(0xFFFFF7ED),
                            borderRadius: BorderRadius.circular(6),
                            border: Border.all(color: const Color(0xFFFFEDD5)),
                          ),
                          child: const Text(
                            "* WAJIB DIPILIH",
                            style: TextStyle(fontSize: 10, fontWeight: FontWeight.bold, color: Color(0xFFEA580C)),
                          ),
                        ),
                      ],
                    ),
                    const SizedBox(height: 8),

                    InkWell(
                      onTap: _showEmployeeSelector,
                      borderRadius: BorderRadius.circular(14),
                      child: Container(
                        padding: const EdgeInsets.symmetric(horizontal: 14, vertical: 12),
                        decoration: BoxDecoration(
                          color: const Color(0xFFF8FAFC),
                          borderRadius: BorderRadius.circular(14),
                          border: Border.all(
                            color: _selectedTarget != null ? const Color(0xFFEA580C) : const Color(0xFFCBD5E1),
                            width: 1.5,
                          ),
                        ),
                        child: Row(
                          children: [
                            Container(
                              padding: const EdgeInsets.all(8),
                              decoration: BoxDecoration(
                                color: _selectedTarget != null ? const Color(0xFFFFEDD5) : const Color(0xFFE2E8F0),
                                shape: BoxShape.circle,
                              ),
                              child: Icon(
                                Icons.person_rounded,
                                color: _selectedTarget != null ? const Color(0xFFEA580C) : const Color(0xFF64748B),
                                size: 20,
                              ),
                            ),
                            const SizedBox(width: 12),
                            Expanded(
                              child: Text(
                                _selectedTarget?.name ?? "Pilih Pegawai...",
                                style: TextStyle(
                                  fontSize: 13.5,
                                  fontWeight: _selectedTarget != null ? FontWeight.bold : FontWeight.normal,
                                  color: _selectedTarget != null ? const Color(0xFF1E293B) : const Color(0xFF94A3B8),
                                ),
                              ),
                            ),
                            const Icon(Icons.arrow_drop_down_rounded, color: Color(0xFF64748B), size: 28),
                          ],
                        ),
                      ),
                    ),

                    const SizedBox(height: 14),

                    // B. Waktu Mulai Tidak di Kantor
                    Row(
                      mainAxisAlignment: MainAxisAlignment.spaceBetween,
                      children: [
                        const Text(
                          "Waktu Mulai Tidak di Kantor",
                          style: TextStyle(fontSize: 13, fontWeight: FontWeight.w800, color: Color(0xFF0F172A)),
                        ),
                        InkWell(
                          onTap: _pickStartTime,
                          borderRadius: BorderRadius.circular(6),
                          child: Padding(
                            padding: const EdgeInsets.symmetric(horizontal: 6, vertical: 2),
                            child: Row(
                              children: const [
                                Icon(Icons.edit_rounded, size: 12, color: Color(0xFFEA580C)),
                                SizedBox(width: 3),
                                Text(
                                  "Ubah Jam",
                                  style: TextStyle(fontSize: 11, fontWeight: FontWeight.bold, color: Color(0xFFEA580C)),
                                ),
                              ],
                            ),
                          ),
                        ),
                      ],
                    ),
                    const SizedBox(height: 8),

                    Container(
                      padding: const EdgeInsets.symmetric(horizontal: 14, vertical: 10),
                      decoration: BoxDecoration(
                        color: const Color(0xFFFFF7ED),
                        borderRadius: BorderRadius.circular(12),
                        border: Border.all(color: const Color(0xFFFFEDD5)),
                      ),
                      child: Row(
                        children: [
                          const Icon(Icons.access_time_rounded, color: Color(0xFFEA580C), size: 20),
                          const SizedBox(width: 10),
                          Text(
                            "$startTimeStr WIT (Hari Ini)",
                            style: const TextStyle(fontSize: 13, fontWeight: FontWeight.bold, color: Color(0xFFC2410C)),
                          ),
                        ],
                      ),
                    ),

                    const SizedBox(height: 14),

                    // C. Keterangan Laporan
                    Row(
                      mainAxisAlignment: MainAxisAlignment.spaceBetween,
                      children: [
                        const Text(
                          "Keterangan",
                          style: TextStyle(fontSize: 13, fontWeight: FontWeight.w800, color: Color(0xFF0F172A)),
                        ),
                        Container(
                          padding: const EdgeInsets.symmetric(horizontal: 8, vertical: 3),
                          decoration: BoxDecoration(
                            color: const Color(0xFFFFF7ED),
                            borderRadius: BorderRadius.circular(6),
                            border: Border.all(color: const Color(0xFFFFEDD5)),
                          ),
                          child: const Text(
                            "* WAJIB DIISI",
                            style: TextStyle(fontSize: 10, fontWeight: FontWeight.bold, color: Color(0xFFEA580C)),
                          ),
                        ),
                      ],
                    ),
                    const SizedBox(height: 8),

                    TextField(
                      controller: _descController,
                      maxLines: 2,
                      style: const TextStyle(fontSize: 13, fontWeight: FontWeight.w500, color: Color(0xFF1E293B)),
                      decoration: InputDecoration(
                        hintText: "Contoh: Tidak ada di meja sejak jam 10 pagi tanpa izin...",
                        hintStyle: const TextStyle(fontSize: 12, color: Color(0xFF94A3B8)),
                        filled: true,
                        fillColor: const Color(0xFFF8FAFC),
                        contentPadding: const EdgeInsets.symmetric(horizontal: 14, vertical: 12),
                        border: OutlineInputBorder(borderRadius: BorderRadius.circular(14), borderSide: const BorderSide(color: Color(0xFFE2E8F0))),
                        enabledBorder: OutlineInputBorder(borderRadius: BorderRadius.circular(14), borderSide: const BorderSide(color: Color(0xFFE2E8F0))),
                        focusedBorder: OutlineInputBorder(borderRadius: BorderRadius.circular(14), borderSide: const BorderSide(color: Color(0xFFEA580C), width: 1.8)),
                      ),
                    ),

                    const SizedBox(height: 14),

                    // D. Dokumen Pendukung (Foto Bukti)
                    Row(
                      mainAxisAlignment: MainAxisAlignment.spaceBetween,
                      children: [
                        const Text(
                          "Dokumen Pendukung (Foto Bukti)",
                          style: TextStyle(fontSize: 13, fontWeight: FontWeight.w800, color: Color(0xFF0F172A)),
                        ),
                        Container(
                          padding: const EdgeInsets.symmetric(horizontal: 8, vertical: 3),
                          decoration: BoxDecoration(
                            color: const Color(0xFFFEF2F2),
                            borderRadius: BorderRadius.circular(6),
                            border: Border.all(color: const Color(0xFFFECACA)),
                          ),
                          child: const Text(
                            "* WAJIB DILAMPIRKAN",
                            style: TextStyle(fontSize: 10, fontWeight: FontWeight.bold, color: Color(0xFFDC2626)),
                          ),
                        ),
                      ],
                    ),
                    const SizedBox(height: 8),

                    if (_imageFile != null) ...[
                      Stack(
                        children: [
                          ClipRRect(
                            borderRadius: BorderRadius.circular(14),
                            child: Image.file(
                              _imageFile!,
                              height: 130,
                              width: double.infinity,
                              fit: BoxFit.cover,
                            ),
                          ),
                          Positioned(
                            top: 8,
                            right: 8,
                            child: InkWell(
                              onTap: () {
                                setState(() {
                                  _imageFile = null;
                                  _imageBase64 = null;
                                });
                              },
                              child: Container(
                                padding: const EdgeInsets.all(6),
                                decoration: const BoxDecoration(color: Colors.black87, shape: BoxShape.circle),
                                child: const Icon(Icons.close_rounded, color: Colors.white, size: 16),
                              ),
                            ),
                          ),
                        ],
                      ),
                    ] else ...[
                      InkWell(
                        onTap: _showImageSourceDialog,
                        borderRadius: BorderRadius.circular(14),
                        child: Container(
                          padding: const EdgeInsets.symmetric(vertical: 16),
                          decoration: BoxDecoration(
                            color: const Color(0xFFFFF7ED),
                            borderRadius: BorderRadius.circular(14),
                            border: Border.all(color: const Color(0xFFFDBA74), style: BorderStyle.solid),
                          ),
                          child: Row(
                            mainAxisAlignment: MainAxisAlignment.center,
                            children: const [
                              Icon(Icons.add_a_photo_rounded, color: Color(0xFFEA580C), size: 20),
                              SizedBox(width: 8),
                              Text(
                                "Ambil / Unggah Foto Bukti",
                                style: TextStyle(fontSize: 12.5, fontWeight: FontWeight.bold, color: Color(0xFFC2410C)),
                              ),
                            ],
                          ),
                        ),
                      ),
                    ],

                    const SizedBox(height: 20),

                    // Tombol Kirim Laporan
                    SizedBox(
                      height: 48,
                      child: ElevatedButton(
                        onPressed: _isSubmitting ? null : _handleSubmit,
                        style: ElevatedButton.styleFrom(
                          backgroundColor: const Color(0xFFEA580C),
                          shape: RoundedRectangleBorder(borderRadius: BorderRadius.circular(14)),
                          elevation: 2,
                        ),
                        child: _isSubmitting
                            ? const SizedBox(
                                width: 22,
                                height: 22,
                                child: CircularProgressIndicator(color: Colors.white, strokeWidth: 2.5),
                              )
                            : const Text(
                                "KIRIM LAPORAN",
                                style: TextStyle(
                                  color: Colors.white,
                                  fontWeight: FontWeight.bold,
                                  fontSize: 14,
                                  letterSpacing: 0.8,
                                ),
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
        borderRadius: BorderRadius.circular(12),
        boxShadow: [
          BoxShadow(
            color: Colors.black.withOpacity(0.12),
            blurRadius: 8,
            offset: const Offset(0, 3),
          ),
        ],
      ),
      child: IconButton(
        icon: Icon(icon, color: const Color(0xFF334155), size: 20),
        tooltip: tooltip,
        onPressed: onTap,
        constraints: const BoxConstraints(minWidth: 40, minHeight: 40),
        padding: EdgeInsets.zero,
      ),
    );
  }
}
