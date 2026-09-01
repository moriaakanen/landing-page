import 'dart:convert';
import 'dart:io';
import 'package:flutter/material.dart';
import 'package:image_picker/image_picker.dart';
import 'package:intl/intl.dart';
import '../models/user_model.dart';
import '../core/services/cepu_service.dart';

class CepuReportView extends StatefulWidget {
  final UserModel reporter;

  const CepuReportView({Key? key, required this.reporter}) : super(key: key);

  @override
  State<CepuReportView> createState() => _CepuReportViewState();
}

class _CepuReportViewState extends State<CepuReportView> {
  final _cepuService = CepuService();
  final _picker = ImagePicker();
  final _descController = TextEditingController();

  List<UserModel> _employees = [];
  UserModel? _selectedTarget;
  bool _isLoadingEmployees = true;
  bool _isSubmitting = false;

  DateTime _startTime = DateTime.now();
  DateTime? _endTime;
  bool _hasReturned = false;

  File? _imageFile;
  String? _imageBase64;

  @override
  void initState() {
    super.initState();
    _loadEmployees();
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

  Future<void> _pickImage(ImageSource source) async {
    try {
      // Optimasi ukuran gambar: Max width/height 800px dan imageQuality 70 (hemat penyimpanan, ukuran ~50KB)
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

  Future<void> _selectStartTime() async {
    final pickedTime = await showTimePicker(
      context: context,
      initialTime: TimeOfDay.fromDateTime(_startTime),
    );
    if (pickedTime != null) {
      final now = DateTime.now();
      setState(() {
        _startTime = DateTime(now.year, now.month, now.day, pickedTime.hour, pickedTime.minute);
      });
    }
  }

  Future<void> _selectEndTime() async {
    final initial = _endTime != null ? TimeOfDay.fromDateTime(_endTime!) : TimeOfDay.now();
    final pickedTime = await showTimePicker(
      context: context,
      initialTime: initial,
    );
    if (pickedTime != null) {
      final now = DateTime.now();
      setState(() {
        _endTime = DateTime(now.year, now.month, now.day, pickedTime.hour, pickedTime.minute);
      });
    }
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
        const SnackBar(content: Text("⚠️ Harap tulis keterangan laporan!"), backgroundColor: Color(0xFFEF4444)),
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
        endTime: _hasReturned ? _endTime : null,
        photoBase64: _imageBase64,
      );

      if (mounted) {
        ScaffoldMessenger.of(context).showSnackBar(
          const SnackBar(
            content: Text("✅ Laporan Cepu berhasil dikirim! Menunggu 4 verifikasi rekan."),
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
    final startTimeStr = DateFormat('HH:mm').format(_startTime);
    final endTimeStr = _endTime != null ? DateFormat('HH:mm').format(_endTime!) : '--:--';

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
          "Lapor Pegawai (Fitur Cepu)",
          style: TextStyle(color: Colors.white, fontWeight: FontWeight.bold, fontSize: 18),
        ),
        centerTitle: true,
      ),
      body: SingleChildScrollView(
        padding: const EdgeInsets.symmetric(horizontal: 18, vertical: 16),
        child: Column(
          crossAxisAlignment: CrossAxisAlignment.stretch,
          children: [
            // Cheerful & Informative Header Banner
            Container(
              padding: const EdgeInsets.all(16),
              decoration: BoxDecoration(
                gradient: const LinearGradient(
                  colors: [Color(0xFFFFF7ED), Color(0xFFFFEDD5)],
                  begin: Alignment.topLeft,
                  end: Alignment.bottomRight,
                ),
                borderRadius: BorderRadius.circular(18),
                border: Border.all(color: const Color(0xFFFDBA74)),
                boxShadow: [
                  BoxShadow(
                    color: const Color(0xFFEA580C).withOpacity(0.08),
                    blurRadius: 10,
                    offset: const Offset(0, 4),
                  ),
                ],
              ),
              child: Row(
                children: [
                  Container(
                    width: 44,
                    height: 44,
                    decoration: BoxDecoration(
                      color: const Color(0xFFEA580C),
                      borderRadius: BorderRadius.circular(14),
                    ),
                    child: const Icon(Icons.campaign_rounded, color: Colors.white, size: 26),
                  ),
                  const SizedBox(width: 14),
                  Expanded(
                    child: Column(
                      crossAxisAlignment: CrossAxisAlignment.start,
                      children: const [
                        Text(
                          "Lapor Tanpa Izin Kantor",
                          style: TextStyle(
                            fontSize: 14,
                            fontWeight: FontWeight.bold,
                            color: Color(0xFF9A3412),
                          ),
                        ),
                        SizedBox(height: 2),
                        Text(
                          "Laporkan rekan yang tidak berada di kantor tanpa mengajukan izin. Laporan membutuhkan 4 verifikasi rekan kerja agar valid.",
                          style: TextStyle(fontSize: 11, color: Color(0xFFC2410C), height: 1.3),
                        ),
                      ],
                    ),
                  ),
                ],
              ),
            ),

            const SizedBox(height: 20),

            // Card Form Container
            Container(
              padding: const EdgeInsets.all(18),
              decoration: BoxDecoration(
                color: Colors.white,
                borderRadius: BorderRadius.circular(20),
                border: Border.all(color: const Color(0xFFE2E8F0)),
                boxShadow: [
                  BoxShadow(
                    color: Colors.black.withOpacity(0.04),
                    blurRadius: 12,
                    offset: const Offset(0, 4),
                  ),
                ],
              ),
              child: Column(
                crossAxisAlignment: CrossAxisAlignment.start,
                children: [
                  // 1. Pilih Pegawai
                  const Text(
                    "Pilih Pegawai yang Dilaporkan *",
                    style: TextStyle(fontSize: 13, fontWeight: FontWeight.bold, color: Color(0xFF1E293B)),
                  ),
                  const SizedBox(height: 8),

                  _isLoadingEmployees
                      ? const Center(child: Padding(padding: EdgeInsets.all(8.0), child: CircularProgressIndicator()))
                      : Container(
                          padding: const EdgeInsets.symmetric(horizontal: 14, vertical: 4),
                          decoration: BoxDecoration(
                            color: const Color(0xFFF8FAFC),
                            borderRadius: BorderRadius.circular(14),
                            border: Border.all(color: const Color(0xFFCBD5E1)),
                          ),
                          child: DropdownButtonHideUnderline(
                            child: DropdownButton<UserModel>(
                              isExpanded: true,
                              value: _selectedTarget,
                              hint: const Text(
                                "Pilih nama pegawai...",
                                style: TextStyle(fontSize: 13, color: Color(0xFF94A3B8)),
                              ),
                              icon: const Icon(Icons.keyboard_arrow_down_rounded, color: Color(0xFFEA580C)),
                              items: _employees.map((user) {
                                return DropdownMenuItem<UserModel>(
                                  value: user,
                                  child: Row(
                                    children: [
                                      CircleAvatar(
                                        radius: 13,
                                        backgroundColor: const Color(0xFFFFEDD5),
                                        child: Text(
                                          user.name.isNotEmpty ? user.name[0].toUpperCase() : 'U',
                                          style: const TextStyle(fontSize: 11, fontWeight: FontWeight.bold, color: Color(0xFFEA580C)),
                                        ),
                                      ),
                                      const SizedBox(width: 10),
                                      Expanded(
                                        child: Column(
                                          crossAxisAlignment: CrossAxisAlignment.start,
                                          mainAxisAlignment: MainAxisAlignment.center,
                                          children: [
                                            Text(
                                              user.name,
                                              style: const TextStyle(fontSize: 13, fontWeight: FontWeight.bold, color: Color(0xFF1E293B)),
                                              maxLines: 1,
                                              overflow: TextOverflow.ellipsis,
                                            ),
                                            Text(
                                              user.department,
                                              style: const TextStyle(fontSize: 11, color: Color(0xFF64748B)),
                                            ),
                                          ],
                                        ),
                                      ),
                                    ],
                                  ),
                                );
                              }).toList(),
                              onChanged: (val) {
                                setState(() => _selectedTarget = val);
                              },
                            ),
                          ),
                        ),

                  const SizedBox(height: 18),

                  // 2. Waktu Mulai & Kembali
                  const Text(
                    "Waktu Terpantau Tidak di Kantor *",
                    style: TextStyle(fontSize: 13, fontWeight: FontWeight.bold, color: Color(0xFF1E293B)),
                  ),
                  const SizedBox(height: 8),

                  Row(
                    children: [
                      // Waktu Mulai
                      Expanded(
                        child: InkWell(
                          onTap: _selectStartTime,
                          borderRadius: BorderRadius.circular(14),
                          child: Container(
                            padding: const EdgeInsets.symmetric(horizontal: 14, vertical: 12),
                            decoration: BoxDecoration(
                              color: const Color(0xFFF8FAFC),
                              borderRadius: BorderRadius.circular(14),
                              border: Border.all(color: const Color(0xFFCBD5E1)),
                            ),
                            child: Column(
                              crossAxisAlignment: CrossAxisAlignment.start,
                              children: [
                                const Text("Waktu Mulai", style: TextStyle(fontSize: 11, color: Color(0xFF64748B))),
                                const SizedBox(height: 4),
                                Row(
                                  children: [
                                    const Icon(Icons.access_time_filled_rounded, size: 16, color: Color(0xFFEA580C)),
                                    const SizedBox(width: 6),
                                    Text(
                                      "$startTimeStr WIB",
                                      style: const TextStyle(fontSize: 14, fontWeight: FontWeight.bold, color: Color(0xFF1E293B)),
                                    ),
                                  ],
                                ),
                              ],
                            ),
                          ),
                        ),
                      ),

                      const SizedBox(width: 10),

                      // Waktu Selesai (Jika sudah kembali)
                      Expanded(
                        child: InkWell(
                          onTap: _hasReturned ? _selectEndTime : null,
                          borderRadius: BorderRadius.circular(14),
                          child: Container(
                            padding: const EdgeInsets.symmetric(horizontal: 14, vertical: 12),
                            decoration: BoxDecoration(
                              color: _hasReturned ? const Color(0xFFF8FAFC) : const Color(0xFFF1F5F9),
                              borderRadius: BorderRadius.circular(14),
                              border: Border.all(
                                color: _hasReturned ? const Color(0xFFCBD5E1) : const Color(0xFFE2E8F0),
                              ),
                            ),
                            child: Column(
                              crossAxisAlignment: CrossAxisAlignment.start,
                              children: [
                                const Text("Waktu Kembali", style: TextStyle(fontSize: 11, color: Color(0xFF64748B))),
                                const SizedBox(height: 4),
                                Row(
                                  children: [
                                    Icon(
                                      Icons.check_circle_rounded,
                                      size: 16,
                                      color: _hasReturned ? const Color(0xFF10B981) : Colors.grey[400],
                                    ),
                                    const SizedBox(width: 6),
                                    Text(
                                      _hasReturned ? "$endTimeStr WIB" : "Belum Kembali",
                                      style: TextStyle(
                                        fontSize: 13,
                                        fontWeight: FontWeight.bold,
                                        color: _hasReturned ? const Color(0xFF1E293B) : Colors.grey[500],
                                      ),
                                    ),
                                  ],
                                ),
                              ],
                            ),
                          ),
                        ),
                      ),
                    ],
                  ),

                  const SizedBox(height: 8),

                  // Checkbox Switch "Pegawai sudah kembali ke kantor?"
                  Row(
                    children: [
                      Checkbox(
                        value: _hasReturned,
                        activeColor: const Color(0xFFEA580C),
                        onChanged: (val) {
                          setState(() {
                            _hasReturned = val ?? false;
                            if (_hasReturned && _endTime == null) {
                              _endTime = DateTime.now();
                            }
                          });
                        },
                      ),
                      const Text(
                        "Pegawai sudah terpantau kembali ke kantor",
                        style: TextStyle(fontSize: 12, fontWeight: FontWeight.w500, color: Color(0xFF475569)),
                      ),
                    ],
                  ),

                  const SizedBox(height: 14),

                  // 3. Keterangan / Kronologi
                  const Text(
                    "Keterangan / Kronologi *",
                    style: TextStyle(fontSize: 13, fontWeight: FontWeight.bold, color: Color(0xFF1E293B)),
                  ),
                  const SizedBox(height: 8),
                  TextField(
                    controller: _descController,
                    maxLines: 3,
                    decoration: InputDecoration(
                      hintText: "Tuliskan keterangan detail (misal: Tidak berada di meja/ruangan kerja sejak pagi)...",
                      hintStyle: const TextStyle(fontSize: 12, color: Color(0xFF94A3B8)),
                      filled: true,
                      fillColor: const Color(0xFFF8FAFC),
                      contentPadding: const EdgeInsets.all(14),
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
                        borderSide: const BorderSide(color: Color(0xFFEA580C), width: 1.5),
                      ),
                    ),
                  ),

                  const SizedBox(height: 18),

                  // 4. Dokumen Pendukung / Foto Bukti
                  const Text(
                    "Foto Bukti Pendukung (Opsional)",
                    style: TextStyle(fontSize: 13, fontWeight: FontWeight.bold, color: Color(0xFF1E293B)),
                  ),
                  const SizedBox(height: 8),

                  if (_imageFile != null) ...[
                    Stack(
                      children: [
                        ClipRRect(
                          borderRadius: BorderRadius.circular(14),
                          child: Image.file(
                            _imageFile!,
                            height: 170,
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
                              decoration: const BoxDecoration(
                                color: Colors.black54,
                                shape: BoxShape.circle,
                              ),
                              child: const Icon(Icons.close, color: Colors.white, size: 18),
                            ),
                          ),
                        ),
                        Positioned(
                          bottom: 8,
                          left: 8,
                          child: Container(
                            padding: const EdgeInsets.symmetric(horizontal: 8, vertical: 4),
                            decoration: BoxDecoration(
                              color: Colors.black.withOpacity(0.6),
                              borderRadius: BorderRadius.circular(8),
                            ),
                            child: const Text(
                              "✓ Foto Terkompresi (<100KB)",
                              style: TextStyle(color: Colors.white, fontSize: 10, fontWeight: FontWeight.bold),
                            ),
                          ),
                        ),
                      ],
                    ),
                    const SizedBox(height: 10),
                  ] else ...[
                    Row(
                      children: [
                        Expanded(
                          child: OutlinedButton.icon(
                            onPressed: () => _pickImage(ImageSource.camera),
                            icon: const Icon(Icons.camera_alt_rounded, size: 18, color: Color(0xFFEA580C)),
                            label: const Text("Ambil Foto", style: TextStyle(fontSize: 12, color: Color(0xFFEA580C))),
                            style: OutlinedButton.styleFrom(
                              padding: const EdgeInsets.symmetric(vertical: 12),
                              side: const BorderSide(color: Color(0xFFFDBA74)),
                              shape: RoundedRectangleBorder(borderRadius: BorderRadius.circular(12)),
                            ),
                          ),
                        ),
                        const SizedBox(width: 10),
                        Expanded(
                          child: OutlinedButton.icon(
                            onPressed: () => _pickImage(ImageSource.gallery),
                            icon: const Icon(Icons.photo_library_rounded, size: 18, color: Color(0xFF475569)),
                            label: const Text("Pilih Galeri", style: TextStyle(fontSize: 12, color: Color(0xFF475569))),
                            style: OutlinedButton.styleFrom(
                              padding: const EdgeInsets.symmetric(vertical: 12),
                              side: const BorderSide(color: Color(0xFFCBD5E1)),
                              shape: RoundedRectangleBorder(borderRadius: BorderRadius.circular(12)),
                            ),
                          ),
                        ),
                      ],
                    ),
                  ],
                ],
              ),
            ),

            const SizedBox(height: 24),

            // Submit Button
            SizedBox(
              height: 50,
              child: ElevatedButton(
                onPressed: _isSubmitting ? null : _handleSubmit,
                style: ElevatedButton.styleFrom(
                  backgroundColor: const Color(0xFFEA580C),
                  elevation: 3,
                  shape: RoundedRectangleBorder(borderRadius: BorderRadius.circular(16)),
                ),
                child: _isSubmitting
                    ? const SizedBox(
                        width: 22,
                        height: 22,
                        child: CircularProgressIndicator(color: Colors.white, strokeWidth: 2.5),
                      )
                    : Row(
                        mainAxisAlignment: MainAxisAlignment.center,
                        children: const [
                          Icon(Icons.send_rounded, color: Colors.white, size: 20),
                          SizedBox(width: 8),
                          Text(
                            "KIRIM LAPORAN CEPU",
                            style: TextStyle(
                              color: Colors.white,
                              fontWeight: FontWeight.bold,
                              fontSize: 14,
                              letterSpacing: 0.5,
                            ),
                          ),
                        ],
                      ),
              ),
            ),

            const SizedBox(height: 24),
          ],
        ),
      ),
    );
  }
}
