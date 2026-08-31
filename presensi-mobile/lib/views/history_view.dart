import 'package:flutter/material.dart';
import 'package:intl/intl.dart';
import '../models/user_model.dart';
import '../models/attendance_model.dart';
import '../core/services/attendance_service.dart';

class HistoryView extends StatelessWidget {
  final UserModel user;

  const HistoryView({Key? key, required this.user}) : super(key: key);

  @override
  Widget build(BuildContext context) {
    final attendanceService = AttendanceService();

    return Scaffold(
      backgroundColor: const Color(0xFFF8FAFC),
      appBar: AppBar(
        title: const Text(
          "Riwayat Presensi",
          style: TextStyle(fontWeight: FontWeight.bold, color: Color(0xFF0F172A), fontSize: 18),
        ),
        backgroundColor: Colors.white,
        elevation: 0,
        leading: IconButton(
          icon: const Icon(Icons.arrow_back_ios_new, color: Color(0xFF0F172A), size: 20),
          onPressed: () => Navigator.pop(context),
        ),
      ),
      body: StreamBuilder<List<AttendanceModel>>(
        stream: attendanceService.streamUserAttendanceHistory(user.uid),
        builder: (context, snapshot) {
          if (snapshot.connectionState == ConnectionState.waiting) {
            return const Center(child: CircularProgressIndicator());
          }

          final list = snapshot.data ?? [];
          if (list.isEmpty) {
            return Center(
              child: Column(
                mainAxisAlignment: MainAxisAlignment.center,
                children: [
                  Icon(Icons.event_busy_rounded, size: 64, color: Colors.grey[400]),
                  const SizedBox(height: 12),
                  Text(
                    "Belum ada data presensi",
                    style: TextStyle(fontSize: 16, color: Colors.grey[600], fontWeight: FontWeight.w600),
                  ),
                ],
              ),
            );
          }

          return ListView.separated(
            padding: const EdgeInsets.symmetric(horizontal: 20, vertical: 16),
            itemCount: list.length,
            separatorBuilder: (context, index) => const SizedBox(height: 12),
            itemBuilder: (context, index) {
              final item = list[index];
              final dateFormatted = DateFormat('EEEE, dd MMM yyyy', 'id_ID').format(
                DateTime.tryParse(item.date) ?? DateTime.now(),
              );

              final bool isLate = item.checkIn?.status == 'LATE';

              return Container(
                padding: const EdgeInsets.all(16),
                decoration: BoxDecoration(
                  color: Colors.white,
                  borderRadius: BorderRadius.circular(18),
                  border: Border.all(color: const Color(0xFFE2E8F0)),
                  boxShadow: [
                    BoxShadow(
                      color: Colors.black.withOpacity(0.02),
                      blurRadius: 10,
                      offset: const Offset(0, 4),
                    ),
                  ],
                ),
                child: Column(
                  crossAxisAlignment: CrossAxisAlignment.start,
                  children: [
                    Row(
                      mainAxisAlignment: MainAxisAlignment.spaceBetween,
                      children: [
                        Text(
                          dateFormatted,
                          style: const TextStyle(fontWeight: FontWeight.bold, fontSize: 14, color: Color(0xFF1E293B)),
                        ),
                        Container(
                          padding: const EdgeInsets.symmetric(horizontal: 10, vertical: 3),
                          decoration: BoxDecoration(
                            color: (isLate ? Colors.orange : Colors.green).withOpacity(0.12),
                            borderRadius: BorderRadius.circular(8),
                          ),
                          child: Text(
                            isLate ? 'Terlambat' : 'Tepat Waktu',
                            style: TextStyle(
                              fontSize: 11,
                              fontWeight: FontWeight.bold,
                              color: isLate ? Colors.orange[800] : Colors.green[800],
                            ),
                          ),
                        ),
                      ],
                    ),
                    const Divider(height: 20, color: Color(0xFFF1F5F9)),
                    Row(
                      children: [
                        Expanded(
                          child: Row(
                            children: [
                              const Icon(Icons.login, size: 16, color: Color(0xFF10B981)),
                              const SizedBox(width: 6),
                              Column(
                                crossAxisAlignment: CrossAxisAlignment.start,
                                children: [
                                  Text("Masuk", style: TextStyle(fontSize: 11, color: Colors.grey[500])),
                                  Text(
                                    item.checkIn != null
                                        ? DateFormat('HH:mm').format(item.checkIn!.time)
                                        : '-',
                                    style: const TextStyle(fontWeight: FontWeight.bold, fontSize: 14),
                                  ),
                                ],
                              ),
                            ],
                          ),
                        ),
                        Expanded(
                          child: Row(
                            children: [
                              const Icon(Icons.logout, size: 16, color: Color(0xFF3B82F6)),
                              const SizedBox(width: 6),
                              Column(
                                crossAxisAlignment: CrossAxisAlignment.start,
                                children: [
                                  Text("Pulang", style: TextStyle(fontSize: 11, color: Colors.grey[500])),
                                  Text(
                                    item.checkOut != null
                                        ? DateFormat('HH:mm').format(item.checkOut!.time)
                                        : '-',
                                    style: const TextStyle(fontWeight: FontWeight.bold, fontSize: 14),
                                  ),
                                ],
                              ),
                            ],
                          ),
                        ),
                        Expanded(
                          child: Row(
                            children: [
                              const Icon(Icons.near_me, size: 16, color: Color(0xFF64748B)),
                              const SizedBox(width: 6),
                              Column(
                                crossAxisAlignment: CrossAxisAlignment.start,
                                children: [
                                  Text("Jarak GPS", style: TextStyle(fontSize: 11, color: Colors.grey[500])),
                                  Text(
                                    "${item.checkIn?.distanceMeters.toStringAsFixed(1) ?? '0'} m",
                                    style: const TextStyle(fontWeight: FontWeight.bold, fontSize: 13),
                                  ),
                                ],
                              ),
                            ],
                          ),
                        ),
                      ],
                    ),
                  ],
                ),
              );
            },
          );
        },
      ),
    );
  }
}
