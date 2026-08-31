import 'package:flutter/material.dart';

class RadarLocationWidget extends StatelessWidget {
  final double distanceMeters;
  final double radiusMeters;
  final bool isChecking;
  final bool isMocked;
  final VoidCallback? onRefresh;

  const RadarLocationWidget({
    Key? key,
    required this.distanceMeters,
    this.radiusMeters = 50.0,
    this.isChecking = false,
    this.isMocked = false,
    this.onRefresh,
  }) : super(key: key);

  @override
  Widget build(BuildContext context) {
    final bool isInRadius = distanceMeters <= radiusMeters && !isMocked;
    final Color primaryColor = isMocked
        ? Colors.redAccent
        : (isInRadius ? const Color(0xFF10B981) : const Color(0xFFF59E0B));

    return Container(
      padding: const EdgeInsets.all(20),
      decoration: BoxDecoration(
        color: Colors.white,
        borderRadius: BorderRadius.circular(24),
        boxShadow: [
          BoxShadow(
            color: primaryColor.withOpacity(0.12),
            blurRadius: 20,
            offset: const Offset(0, 8),
          )
        ],
        border: Border.all(
          color: primaryColor.withOpacity(0.3),
          width: 1.5,
        ),
      ),
      child: Column(
        children: [
          Row(
            mainAxisAlignment: MainAxisAlignment.spaceBetween,
            children: [
              Row(
                children: [
                  Container(
                    padding: const EdgeInsets.all(8),
                    decoration: BoxDecoration(
                      color: primaryColor.withOpacity(0.1),
                      shape: BoxShape.circle,
                    ),
                    child: Icon(
                      isMocked
                          ? Icons.warning_amber_rounded
                          : (isInRadius ? Icons.verified_user : Icons.location_off_outlined),
                      color: primaryColor,
                      size: 22,
                    ),
                  ),
                  const SizedBox(width: 10),
                  Text(
                    isMocked
                        ? "Fake GPS Terdeteksi!"
                        : (isInRadius ? "Lokasi Valid (Dalam Radius)" : "Di Luar Radius Kantor"),
                    style: TextStyle(
                      fontWeight: FontWeight.bold,
                      fontSize: 15,
                      color: primaryColor,
                    ),
                  ),
                ],
              ),
              if (onRefresh != null)
                IconButton(
                  icon: isChecking
                      ? const SizedBox(
                          width: 18,
                          height: 18,
                          child: CircularProgressIndicator(strokeWidth: 2),
                        )
                      : const Icon(Icons.refresh, color: Colors.grey),
                  onPressed: isChecking ? null : onRefresh,
                  tooltip: "Perbarui Lokasi GPS",
                ),
            ],
          ),
          const SizedBox(height: 16),
          // Visual Meter
          Stack(
            alignment: Alignment.center,
            children: [
              Container(
                width: 110,
                height: 110,
                decoration: BoxDecoration(
                  shape: BoxShape.circle,
                  color: primaryColor.withOpacity(0.08),
                  border: Border.all(
                    color: primaryColor.withOpacity(0.25),
                    width: 2,
                  ),
                ),
              ),
              Container(
                width: 75,
                height: 75,
                decoration: BoxDecoration(
                  shape: BoxShape.circle,
                  color: primaryColor.withOpacity(0.15),
                ),
              ),
              Icon(
                Icons.my_location,
                color: primaryColor,
                size: 32,
              ),
            ],
          ),
          const SizedBox(height: 14),
          Text(
            isMocked
                ? "Harap matikan aplikasi Fake GPS"
                : "${distanceMeters.toStringAsFixed(1)} m dari Titik Kantor",
            style: const TextStyle(
              fontSize: 18,
              fontWeight: FontWeight.w700,
              color: Color(0xFF1E293B),
            ),
          ),
          const SizedBox(height: 4),
          Text(
            "Batas radius presensi: ${radiusMeters.toStringAsFixed(0)} meter",
            style: TextStyle(
              fontSize: 13,
              color: Colors.grey[600],
            ),
          ),
        ],
      ),
    );
  }
}
