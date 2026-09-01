import 'package:flutter/material.dart';
import 'package:flutter_map/flutter_map.dart';
import 'package:latlong2/latlong.dart';
import 'package:geolocator/geolocator.dart';
import 'package:intl/intl.dart';
import '../models/user_model.dart';
import '../models/office_model.dart';
import '../models/permit_model.dart';
import '../core/services/permit_service.dart';
import '../core/services/location_service.dart';

class RecordPermitMapView extends StatefulWidget {
  final UserModel user;
  final OfficeModel office;
  final PermitModel? initialActivePermit;

  const RecordPermitMapView({
    Key? key,
    required this.user,
    required this.office,
    this.initialActivePermit,
  }) : super(key: key);

  @override
  State<RecordPermitMapView> createState() => _RecordPermitMapViewState();
}

class _RecordPermitMapViewState extends State<RecordPermitMapView> {
  final MapController _mapController = MapController();
  final PermitService _permitService = PermitService();
  final TextEditingController _purposeController = TextEditingController();

  PermitModel? _activePermit;
  Position? _currentPosition;
  double _distanceMeters = 999.0;
  double _accuracyMeters = 0.0;
  bool _isInRadius = false;
  bool _isMocked = false;
  bool _isLoadingLocation = true;
  bool _isSubmitting = false;
  String? _errorMessage;

  @override
  void initState() {
    super.initState();
    _activePermit = widget.initialActivePermit;
    _fetchActivePermitAndLocation();
  }

  @override
  void dispose() {
    _purposeController.dispose();
    super.dispose();
  }

  Future<void> _fetchActivePermitAndLocation() async {
    setState(() {
      _isLoadingLocation = true;
      _errorMessage = null;
    });

    try {
      // 1. Cek izin aktif terbaru
      final active = await _permitService.getActivePermit(widget.user.uid);
      if (mounted) {
        setState(() {
          _activePermit = active;
        });
      }

      // 2. Verifikasi lokasi GPS
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

        // Focus map ke lokasi user
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

  Future<void> _handleStartPermit() async {
    final purpose = _purposeController.text.trim();
    if (purpose.isEmpty) {
      ScaffoldMessenger.of(context).showSnackBar(
        const SnackBar(
          content: Text("⚠️ Harap isi keperluan izin terlebih dahulu!"),
          backgroundColor: Color(0xFFEF4444),
        ),
      );
      return;
    }

    if (!_isInRadius || _isMocked) {
      ScaffoldMessenger.of(context).showSnackBar(
        SnackBar(
          content: Text(_isMocked
              ? "⚠️ Deteksi Fake GPS aktif!"
              : "⚠️ Anda harus berada di radius kantor untuk mulai izin!"),
          backgroundColor: const Color(0xFFEF4444),
        ),
      );
      return;
    }

    setState(() => _isSubmitting = true);

    try {
      final permit = await _permitService.startPermit(
        user: widget.user,
        office: widget.office,
        purpose: purpose,
      );

      if (mounted) {
        ScaffoldMessenger.of(context).showSnackBar(
          const SnackBar(
            content: Text("✅ Berhasil merekam Mulai Izin! Hati-hati di jalan."),
            backgroundColor: Color(0xFF10B981),
          ),
        );
        setState(() {
          _activePermit = permit;
          _purposeController.clear();
        });
        Navigator.pop(context, true);
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
      if (mounted) {
        setState(() => _isSubmitting = false);
      }
    }
  }

  Future<void> _handleEndPermit() async {
    if (_activePermit == null) return;

    if (!_isInRadius || _isMocked) {
      ScaffoldMessenger.of(context).showSnackBar(
        SnackBar(
          content: Text(_isMocked
              ? "⚠️ Deteksi Fake GPS aktif!"
              : "⚠️ Anda harus sudah kembali di dalam radius kantor untuk menyelesaikan izin!"),
          backgroundColor: const Color(0xFFEF4444),
        ),
      );
      return;
    }

    setState(() => _isSubmitting = true);

    try {
      await _permitService.endPermit(
        permitId: _activePermit!.id,
        office: widget.office,
      );

      if (mounted) {
        ScaffoldMessenger.of(context).showSnackBar(
          const SnackBar(
            content: Text("✅ Selesai Izin! Terima kasih sudah kembali tepat waktu."),
            backgroundColor: Color(0xFF10B981),
          ),
        );
        setState(() {
          _activePermit = null;
        });
        Navigator.pop(context, true);
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

    final isPermitActive = _activePermit != null;
    final canSubmit = _isInRadius && !_isMocked && !_isLoadingLocation && !_isSubmitting;

    return Scaffold(
      backgroundColor: const Color(0xFFF8FAFC),
      appBar: AppBar(
        backgroundColor: const Color(0xFF1E60F2),
        elevation: 0,
        leading: IconButton(
          icon: const Icon(Icons.arrow_back_ios_new, color: Colors.white, size: 20),
          onPressed: () => Navigator.pop(context),
        ),
        title: Text(
          isPermitActive ? "Selesai Waigama" : "Waigama",
          style: const TextStyle(
            color: Colors.white,
            fontWeight: FontWeight.w700,
            fontSize: 18,
          ),
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
            onPressed: _isLoadingLocation ? null : _fetchActivePermitAndLocation,
          ),
        ],
      ),
      body: Column(
        children: [
          // Map Area
          Expanded(
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
                    // Tile Layer
                    TileLayer(
                      urlTemplate: 'https://tile.openstreetmap.org/{z}/{x}/{y}.png',
                      userAgentPackageName: 'com.example.presensi_app',
                    ),

                    // Geofence Radius Circle (50m)
                    CircleLayer(
                      circles: [
                        CircleMarker(
                          point: officeLatLng,
                          color: (_isInRadius
                                  ? const Color(0xFF3B82F6)
                                  : const Color(0xFFEF4444))
                              .withOpacity(0.15),
                          borderColor: _isInRadius
                              ? const Color(0xFF2563EB)
                              : const Color(0xFFDC2626),
                          borderStrokeWidth: 2,
                          useRadiusInMeter: true,
                          radius: widget.office.radiusMeters,
                        ),
                      ],
                    ),

                    // Markers
                    MarkerLayer(
                      markers: [
                        // Office Marker (Center aligned)
                        Marker(
                          point: officeLatLng,
                          width: 44,
                          height: 44,
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
                            child: const Icon(
                              Icons.business_rounded,
                              color: Colors.white,
                              size: 24,
                            ),
                          ),
                        ),

                        // User Live Position Marker (Precision Center-Aligned Blue Indicator)
                        // Posisi simetris dengan Alignment.center agar terkunci 100% tepat di koordinat GPS tanpa bergeser pada level zoom berapa pun
                        if (_currentPosition != null)
                          Marker(
                            point: userLatLng,
                            width: 38,
                            height: 38,
                            alignment: Alignment.center,
                            child: Stack(
                              alignment: Alignment.center,
                              children: [
                                // Outer pulsating halo
                                Container(
                                  width: 38,
                                  height: 38,
                                  decoration: BoxDecoration(
                                    color: const Color(0xFF2563EB).withOpacity(0.25),
                                    shape: BoxShape.circle,
                                  ),
                                ),
                                // Inner blue circle with person icon & white border
                                Container(
                                  width: 26,
                                  height: 26,
                                  decoration: BoxDecoration(
                                    color: const Color(0xFF2563EB),
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
                                    child: Icon(
                                      Icons.person_rounded,
                                      color: Colors.white,
                                      size: 15,
                                    ),
                                  ),
                                ),
                              ],
                            ),
                          ),
                      ],
                    ),
                  ],
                ),

                // Floating Map Controls (Right Side)
                Positioned(
                  top: 16,
                  right: 16,
                  child: Column(
                    children: [
                      _buildFloatingMapButton(
                        icon: Icons.apartment_rounded,
                        tooltip: "Pusatkan ke Kantor",
                        onTap: () => _mapController.move(officeLatLng, 17.5),
                      ),
                      const SizedBox(height: 8),
                      _buildFloatingMapButton(
                        icon: Icons.my_location_rounded,
                        tooltip: "Pusatkan ke Lokasi Saya",
                        onTap: () {
                          if (_currentPosition != null) {
                            _mapController.move(userLatLng, 18.0);
                          }
                        },
                      ),
                      const SizedBox(height: 8),
                      _buildFloatingMapButton(
                        icon: Icons.add,
                        tooltip: "Perbesar",
                        onTap: () {
                          _mapController.move(
                            _mapController.camera.center,
                            _mapController.camera.zoom + 1,
                          );
                        },
                      ),
                      const SizedBox(height: 8),
                      _buildFloatingMapButton(
                        icon: Icons.remove,
                        tooltip: "Perkecil",
                        onTap: () {
                          _mapController.move(
                            _mapController.camera.center,
                            _mapController.camera.zoom - 1,
                          );
                        },
                      ),
                    ],
                  ),
                ),
              ],
            ),
          ),

          // Bottom Control Panel
          Container(
            padding: const EdgeInsets.symmetric(horizontal: 20, vertical: 18),
            decoration: const BoxDecoration(
              color: Colors.white,
              borderRadius: BorderRadius.vertical(top: Radius.circular(24)),
              boxShadow: [
                BoxShadow(
                  color: Color(0x0F000000),
                  blurRadius: 15,
                  offset: Offset(0, -4),
                ),
              ],
            ),
            child: SafeArea(
              top: false,
              child: Column(
                mainAxisSize: MainAxisSize.min,
                crossAxisAlignment: CrossAxisAlignment.stretch,
                children: [
                  // Status Distance Badge
                  Container(
                    padding: const EdgeInsets.symmetric(vertical: 12, horizontal: 16),
                    decoration: BoxDecoration(
                      color: _isInRadius
                          ? const Color(0xFFEFF6FF)
                          : const Color(0xFFFEF2F2),
                      borderRadius: BorderRadius.circular(16),
                      border: Border.all(
                        color: _isInRadius
                            ? const Color(0xFFBFDBFE)
                            : const Color(0xFFFECACA),
                      ),
                    ),
                    child: Column(
                      crossAxisAlignment: CrossAxisAlignment.center,
                      children: [
                        Row(
                          mainAxisAlignment: MainAxisAlignment.center,
                          children: [
                            Icon(
                              _isInRadius
                                  ? Icons.check_circle_rounded
                                  : Icons.location_off_rounded,
                              color: _isInRadius
                                  ? const Color(0xFF2563EB)
                                  : const Color(0xFFDC2626),
                              size: 20,
                            ),
                            const SizedBox(width: 8),
                            Text(
                              _isInRadius
                                  ? "Dalam Jangkauan Kantor"
                                  : "Di Luar Jangkauan Kantor",
                              style: TextStyle(
                                fontWeight: FontWeight.bold,
                                fontSize: 14,
                                color: _isInRadius
                                    ? const Color(0xFF1E40AF)
                                    : const Color(0xFF991B1B),
                              ),
                            ),
                          ],
                        ),
                        const SizedBox(height: 6),
                        Text(
                          "Akurasi GPS: ${_accuracyMeters.toStringAsFixed(1)} m  •  Jarak ke Kantor: ${_distanceMeters.toStringAsFixed(1)} m",
                          style: TextStyle(
                            fontSize: 11,
                            color: Colors.grey[700],
                            fontWeight: FontWeight.w500,
                          ),
                          textAlign: TextAlign.center,
                        ),
                      ],
                    ),
                  ),

                  const SizedBox(height: 14),

                  // IF ACTIVE PERMIT: Display active permit info
                  if (isPermitActive) ...[
                    Container(
                      padding: const EdgeInsets.all(14),
                      decoration: BoxDecoration(
                        color: const Color(0xFFFFFBEB),
                        borderRadius: BorderRadius.circular(14),
                        border: Border.all(color: const Color(0xFFFDE68A)),
                      ),
                      child: Column(
                        crossAxisAlignment: CrossAxisAlignment.start,
                        children: [
                          Row(
                            children: [
                              Container(
                                padding: const EdgeInsets.symmetric(horizontal: 8, vertical: 3),
                                decoration: BoxDecoration(
                                  color: const Color(0xFFF59E0B),
                                  borderRadius: BorderRadius.circular(8),
                                ),
                                child: const Text(
                                  "SEDANG IZIN",
                                  style: TextStyle(
                                    color: Colors.white,
                                    fontSize: 10,
                                    fontWeight: FontWeight.bold,
                                  ),
                                ),
                              ),
                              const SizedBox(width: 8),
                              Text(
                                "Mulai: ${DateFormat('HH:mm').format(_activePermit!.startTime)} WIT",
                                style: const TextStyle(
                                  fontSize: 12,
                                  fontWeight: FontWeight.bold,
                                  color: Color(0xFF92400E),
                                ),
                              ),
                            ],
                          ),
                          const SizedBox(height: 8),
                          RichText(
                            text: TextSpan(
                              style: const TextStyle(fontSize: 13, color: Color(0xFF1E293B)),
                              children: [
                                const TextSpan(
                                  text: "Keperluan: ",
                                  style: TextStyle(fontWeight: FontWeight.bold),
                                ),
                                TextSpan(text: _activePermit!.purpose),
                              ],
                            ),
                          ),
                          const SizedBox(height: 4),
                          Text(
                            "Durasi berjalan: ${_activePermit!.durationString}",
                            style: TextStyle(
                              fontSize: 11,
                              color: Colors.grey[600],
                              fontStyle: FontStyle.italic,
                            ),
                          ),
                        ],
                      ),
                    ),
                    const SizedBox(height: 14),
                  ] else ...[
                    // IF NO ACTIVE PERMIT: Input Purpose Field
                    TextField(
                      controller: _purposeController,
                      decoration: InputDecoration(
                        labelText: "Keperluan Waigama",
                        hintText: "Contoh: Koordinasi dinas ke Bappeda",
                        prefixIcon: const Icon(Icons.edit_note_rounded, color: Color(0xFF2563EB)),
                        filled: true,
                        fillColor: const Color(0xFFF8FAFC),
                        contentPadding: const EdgeInsets.symmetric(horizontal: 14, vertical: 12),
                        border: OutlineInputBorder(
                          borderRadius: BorderRadius.circular(14),
                          borderSide: const BorderSide(color: Color(0xFFCBD5E1)),
                        ),
                        enabledBorder: OutlineInputBorder(
                          borderRadius: BorderRadius.circular(14),
                          borderSide: const BorderSide(color: Color(0xFFE2E8F0)),
                        ),
                        focusedBorder: OutlineInputBorder(
                          borderRadius: BorderRadius.circular(14),
                          borderSide: const BorderSide(color: Color(0xFF2563EB), width: 1.5),
                        ),
                      ),
                    ),
                    const SizedBox(height: 14),
                  ],

                  // Action Button (Mulai Waigama / Selesai Waigama)
                  SizedBox(
                    height: 48,
                    child: ElevatedButton(
                      onPressed: canSubmit
                          ? (isPermitActive ? _handleEndPermit : _handleStartPermit)
                          : null,
                      style: ElevatedButton.styleFrom(
                        backgroundColor: isPermitActive
                            ? const Color(0xFF10B981) // Green for finish
                            : const Color(0xFF1E60F2), // Blue for start
                        disabledBackgroundColor: const Color(0xFFCBD5E1),
                        elevation: canSubmit ? 2 : 0,
                        shape: RoundedRectangleBorder(
                          borderRadius: BorderRadius.circular(14),
                        ),
                      ),
                      child: _isSubmitting
                          ? const SizedBox(
                              width: 20,
                              height: 20,
                              child: CircularProgressIndicator(color: Colors.white, strokeWidth: 2.5),
                            )
                          : Text(
                              isPermitActive
                                  ? "REKAM SELESAI WAIGAMA (KEMBALI KE KANTOR)"
                                  : "REKAM MULAI WAIGAMA",
                              style: const TextStyle(
                                color: Colors.white,
                                fontWeight: FontWeight.bold,
                                fontSize: 13,
                                letterSpacing: 0.5,
                              ),
                            ),
                    ),
                  ),
                ],
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
