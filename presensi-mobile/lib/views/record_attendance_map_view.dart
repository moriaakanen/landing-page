import 'package:flutter/material.dart';
import 'package:flutter_map/flutter_map.dart';
import 'package:latlong2/latlong.dart';
import 'package:geolocator/geolocator.dart';
import '../models/user_model.dart';
import '../models/office_model.dart';
import '../models/attendance_model.dart';
import '../core/services/attendance_service.dart';
import '../core/services/location_service.dart';

class RecordAttendanceMapView extends StatefulWidget {
  final UserModel user;
  final OfficeModel office;
  final AttendanceModel? todayAttendance;

  const RecordAttendanceMapView({
    Key? key,
    required this.user,
    required this.office,
    this.todayAttendance,
  }) : super(key: key);

  @override
  State<RecordAttendanceMapView> createState() => _RecordAttendanceMapViewState();
}

class _RecordAttendanceMapViewState extends State<RecordAttendanceMapView> {
  final MapController _mapController = MapController();
  final AttendanceService _attendanceService = AttendanceService();

  Position? _currentPosition;
  double _distanceMeters = 999.0;
  double _accuracyMeters = 0.0;
  bool _isInRadius = false;
  bool _isMocked = false;
  bool _isLoadingLocation = true;
  bool _isSavingAttendance = false;
  String? _errorMessage;

  @override
  void initState() {
    super.initState();
    _fetchCurrentLocation();
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

        // Center map to user location
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

  Future<void> _handleSaveAttendance() async {
    if (!_isInRadius || _isMocked) return;

    setState(() => _isSavingAttendance = true);

    try {
      final bool hasCheckedIn = widget.todayAttendance?.checkIn != null;
      if (!hasCheckedIn) {
        await _attendanceService.checkIn(user: widget.user, office: widget.office);
      } else {
        await _attendanceService.checkOut(user: widget.user, office: widget.office);
      }

      if (mounted) {
        ScaffoldMessenger.of(context).showSnackBar(
          SnackBar(
            content: Text(!hasCheckedIn
                ? "✅ Berhasil merekam Kehadiran Masuk!"
                : "✅ Berhasil merekam Kehadiran Pulang!"),
            backgroundColor: const Color(0xFF10B981),
          ),
        );
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
        setState(() => _isSavingAttendance = false);
      }
    }
  }

  @override
  Widget build(BuildContext context) {
    final officeLatLng = LatLng(widget.office.latitude, widget.office.longitude);
    final userLatLng = _currentPosition != null
        ? LatLng(_currentPosition!.latitude, _currentPosition!.longitude)
        : officeLatLng;

    final bool isActionDisabled = !_isInRadius || _isMocked || _isLoadingLocation || _isSavingAttendance;

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
          "Buat Kehadiran",
          style: TextStyle(color: Colors.white, fontWeight: FontWeight.w700, fontSize: 18),
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
          // Interactive Map Area
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
                    // OpenStreetMap Tile Layer
                    TileLayer(
                      urlTemplate: 'https://tile.openstreetmap.org/{z}/{x}/{y}.png',
                      userAgentPackageName: 'com.example.presensi_app',
                    ),

                    // Geofence Radius Circle (50m around office)
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
                          radius: widget.office.radiusMeters, // 50m
                        ),
                      ],
                    ),

                    // Markers
                    MarkerLayer(
                      markers: [
                        // 1. Office Marker
                        Marker(
                          point: officeLatLng,
                          width: 44,
                          height: 44,
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

                        // 2. User Live Position Marker with Callout Bubble
                        // Anchored at bottomCenter to ensure pin tip stays exactly on GPS coordinates regardless of zoom level
                        if (_currentPosition != null)
                          Marker(
                            point: userLatLng,
                            width: 140,
                            height: 64,
                            alignment: Alignment.bottomCenter,
                            child: Column(
                              mainAxisSize: MainAxisSize.min,
                              crossAxisAlignment: CrossAxisAlignment.center,
                              children: [
                                // Callout Bubble "Anda di sini"
                                Container(
                                  padding: const EdgeInsets.symmetric(horizontal: 9, vertical: 3),
                                  decoration: BoxDecoration(
                                    color: Colors.white,
                                    borderRadius: BorderRadius.circular(10),
                                    boxShadow: [
                                      BoxShadow(
                                        color: Colors.black.withOpacity(0.18),
                                        blurRadius: 6,
                                        offset: const Offset(0, 2),
                                      ),
                                    ],
                                  ),
                                  child: const Text(
                                    "Anda di sini",
                                    style: TextStyle(
                                      fontSize: 11,
                                      fontWeight: FontWeight.bold,
                                      color: Color(0xFF1E293B),
                                    ),
                                  ),
                                ),
                                const SizedBox(height: 2),
                                // Blue Pin Icon with tip at the bottom
                                const Icon(
                                  Icons.location_on,
                                  color: Color(0xFF2563EB),
                                  size: 36,
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
                      // Focus Office Button
                      _buildFloatingMapButton(
                        icon: Icons.apartment_rounded,
                        tooltip: "Pusatkan ke Kantor",
                        onTap: () => _mapController.move(officeLatLng, 17.5),
                      ),
                      const SizedBox(height: 8),
                      // Focus My Location Button
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
                      // Zoom In
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
                      // Zoom Out
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

          // Bottom Control & Status Panel
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
            child: Column(
              mainAxisSize: MainAxisSize.min,
              children: [
                // Status Box Card
                Container(
                  width: double.infinity,
                  padding: const EdgeInsets.symmetric(vertical: 16, horizontal: 16),
                  decoration: BoxDecoration(
                    color: _isMocked
                        ? const Color(0xFFFEF2F2)
                        : (_isInRadius ? const Color(0xFFEFF6FF) : const Color(0xFFFFFBEB)),
                    borderRadius: BorderRadius.circular(16),
                    border: Border.all(
                      color: _isMocked
                          ? const Color(0xFFF87171)
                          : (_isInRadius ? const Color(0xFFBFDBFE) : const Color(0xFFFDE68A)),
                    ),
                  ),
                  child: Column(
                    children: [
                      Text(
                        _isMocked
                            ? "Fake GPS / Lokasi Palsu Terdeteksi!"
                            : (_isInRadius ? "Dalam Jangkauan Kantor" : "Di Luar Jangkauan Kantor"),
                        style: TextStyle(
                          fontSize: 16,
                          fontWeight: FontWeight.bold,
                          color: _isMocked
                              ? const Color(0xFFDC2626)
                              : (_isInRadius ? const Color(0xFF1E40AF) : const Color(0xFFD97706)),
                        ),
                      ),
                      const SizedBox(height: 6),
                      Text(
                        _isMocked
                            ? "Matikan aplikasi Fake GPS untuk melanjutkan presensi."
                            : (_currentPosition != null
                                ? "Lokasi Saat Ini (Akurat sampai ${_accuracyMeters.toStringAsFixed(3)} meter) • Jarak: ${_distanceMeters.toStringAsFixed(1)} m"
                                : (_errorMessage ?? "Sedang mendeteksi sinyal GPS...")),
                        textAlign: TextAlign.center,
                        style: TextStyle(
                          fontSize: 12,
                          color: Colors.grey[700],
                        ),
                      ),
                    ],
                  ),
                ),

                const SizedBox(height: 16),

                // Action Button: SIMPAN KEHADIRAN
                SizedBox(
                  width: double.infinity,
                  height: 52,
                  child: ElevatedButton(
                    onPressed: isActionDisabled ? null : _handleSaveAttendance,
                    style: ElevatedButton.styleFrom(
                      backgroundColor: const Color(0xFF1E60F2),
                      disabledBackgroundColor: const Color(0xFFCBD5E1),
                      shape: RoundedRectangleBorder(
                        borderRadius: BorderRadius.circular(14),
                      ),
                      elevation: _isInRadius ? 2 : 0,
                    ),
                    child: _isSavingAttendance
                        ? const SizedBox(
                            width: 22,
                            height: 22,
                            child: CircularProgressIndicator(color: Colors.white, strokeWidth: 2.5),
                          )
                        : Text(
                            "SIMPAN KEHADIRAN",
                            style: TextStyle(
                              fontSize: 15,
                              fontWeight: FontWeight.bold,
                              color: isActionDisabled ? const Color(0xFF64748B) : Colors.white,
                              letterSpacing: 0.5,
                            ),
                          ),
                  ),
                ),
              ],
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
    return Material(
      color: Colors.white,
      shape: RoundedRectangleBorder(borderRadius: BorderRadius.circular(12)),
      elevation: 3,
      shadowColor: Colors.black26,
      child: InkWell(
        borderRadius: BorderRadius.circular(12),
        onTap: onTap,
        child: SizedBox(
          width: 44,
          height: 44,
          child: Icon(icon, color: const Color(0xFF334155), size: 20),
        ),
      ),
    );
  }
}
