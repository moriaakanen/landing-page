class EmployeeData {
  final String username;
  final String fullName;
  final String department;

  const EmployeeData({
    required this.username,
    required this.fullName,
    this.department = 'Pegawai',
  });
}

class AppEmployees {
  static const List<EmployeeData> list = [
    EmployeeData(username: 'frida', fullName: 'Frida Irian S. Ompusunggu', department: 'Umum & Kepegawaian'),
    EmployeeData(username: 'andriew', fullName: 'Andrie Wardani', department: 'Teknologi Informasi & Data'),
    EmployeeData(username: 'derek', fullName: 'Derek Mandowen', department: 'Umum & Kepegawaian'),
    EmployeeData(username: 'frida.buratehi', fullName: 'Frida Buratehi', department: 'Keuangan & Perbendaharaan'),
    EmployeeData(username: 'maulana.tahir', fullName: 'Maulana Tahir', department: 'Teknologi Informasi & Data'),
    EmployeeData(username: 'ikbal.ism', fullName: 'Ikbal Ismail', department: 'Statistik & Analisis'),
    EmployeeData(username: 'lestari.amir', fullName: 'Lestari Irfandi Amir', department: 'Statistik & Analisis'),
    EmployeeData(username: 'bagas.sakti', fullName: 'Bagas Indra Sakti', department: 'Teknologi Informasi & Data'),
    EmployeeData(username: 'novalin', fullName: 'Novalin P. Wapai', department: 'Umum & Kepegawaian'),
    EmployeeData(username: 'gerda', fullName: 'Gerda Y. Kalasuat', department: 'Keuangan & Perbendaharaan'),
    EmployeeData(username: 'rizal.akbar', fullName: 'Rizal Akbar Komarudin', department: 'Statistik & Analisis'),
    EmployeeData(username: 'jandriana.ramandei', fullName: 'Jandriana Ramandei', department: 'Umum & Kepegawaian'),
    EmployeeData(username: 'elok', fullName: 'Elok Agustina', department: 'Statistik & Analisis'),
    EmployeeData(username: 'zidna.inayatika', fullName: 'Zidna Inayatika', department: 'Keuangan & Perbendaharaan'),
    EmployeeData(username: 'haidar.nabil', fullName: 'Haidar Nabil', department: 'Teknologi Informasi & Data'),
    EmployeeData(username: 'syayu.hanana', fullName: 'Syayu Hanana Yohana Stefani', department: 'Statistik & Analisis'),
    EmployeeData(username: 'aris.munandar', fullName: 'M. Aris Munandar', department: 'Statistik & Analisis'),
    EmployeeData(username: 'dwiayua', fullName: 'Dwi Ayu Andriani', department: 'Keuangan & Perbendaharaan'),
    EmployeeData(username: 'figri.alrasjid', fullName: 'Figri Al - Rasjid Abdullah', department: 'Teknologi Informasi & Data'),
    EmployeeData(username: 'arikhzasaputri', fullName: 'Arikhza Saputri', department: 'Statistik & Analisis'),
    EmployeeData(username: 'khusenali', fullName: 'M. Khusen Ali Al Anjabi', department: 'Statistik & Analisis'),
  ];
}
