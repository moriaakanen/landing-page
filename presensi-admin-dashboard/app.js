// State Manajemen Dashboard Admin & Geofencing
let officeConfig = {
  name: "Kantor Pusat - Jakarta",
  lat: -6.208800,
  lng: 106.845600,
  radius: 50,
  startTime: "08:00",
  endTime: "17:00"
};

let employees = [
  { id: "user_1", name: "Budi Santoso", dept: "IT & Engineering", email: "budi@perusahaan.com" },
  { id: "user_2", name: "Siti Rahma", dept: "HR & People", email: "siti@perusahaan.com" },
  { id: "user_3", name: "Ahmad Fauzi", dept: "Finance & Accounting", email: "ahmad@perusahaan.com" },
  { id: "user_4", name: "Dewi Lestari", dept: "Marketing & Growth", email: "dewi@perusahaan.com" },
  { id: "user_5", name: "Rizky Pratama", dept: "Operations", email: "rizky@perusahaan.com" },
  { id: "user_6", name: "Nadia Utami", dept: "Product Design", email: "nadia@perusahaan.com" },
  { id: "user_7", name: "Eko Prasetyo", dept: "IT & Engineering", email: "eko@perusahaan.com" },
  { id: "user_8", name: "Maya Indah", dept: "Customer Support", email: "maya@perusahaan.com" },
  { id: "user_9", name: "Hendra Wijaya", dept: "Operations", email: "hendra@perusahaan.com" },
  { id: "user_10", name: "Putri Anggraini", dept: "Marketing & Growth", email: "putri@perusahaan.com" },
  { id: "user_11", name: "Dimas Saputra", dept: "Finance & Accounting", email: "dimas@perusahaan.com" },
  { id: "user_12", name: "Anisa Fitri", dept: "HR & People", email: "anisa@perusahaan.com" }
];

let attendanceLogs = [
  {
    userId: "user_1",
    name: "Budi Santoso",
    dept: "IT & Engineering",
    checkIn: "07:52:10",
    checkOut: "17:05:30",
    distanceMeters: 12.4,
    isMock: false,
    status: "ON_TIME"
  },
  {
    userId: "user_2",
    name: "Siti Rahma",
    dept: "HR & People",
    checkIn: "07:45:00",
    checkOut: "17:02:15",
    distanceMeters: 8.1,
    isMock: false,
    status: "ON_TIME"
  },
  {
    userId: "user_3",
    name: "Ahmad Fauzi",
    dept: "Finance & Accounting",
    checkIn: "08:14:22",
    checkOut: "-",
    distanceMeters: 28.5,
    isMock: false,
    status: "LATE"
  },
  {
    userId: "user_4",
    name: "Dewi Lestari",
    dept: "Marketing & Growth",
    checkIn: "07:58:40",
    checkOut: "-",
    distanceMeters: 35.0,
    isMock: false,
    status: "ON_TIME"
  },
  {
    userId: "user_5",
    name: "Rizky Pratama",
    dept: "Operations",
    checkIn: "08:08:12",
    checkOut: "-",
    distanceMeters: 18.2,
    isMock: false,
    status: "LATE"
  },
  {
    userId: "user_6",
    name: "Nadia Utami",
    dept: "Product Design",
    checkIn: "07:30:19",
    checkOut: "-",
    distanceMeters: 6.7,
    isMock: false,
    status: "ON_TIME"
  },
  {
    userId: "user_7",
    name: "Eko Prasetyo",
    dept: "IT & Engineering",
    checkIn: "07:48:55",
    checkOut: "-",
    distanceMeters: 15.3,
    isMock: false,
    status: "ON_TIME"
  },
  {
    userId: "user_8",
    name: "Maya Indah",
    dept: "Customer Support",
    checkIn: "07:50:02",
    checkOut: "-",
    distanceMeters: 22.0,
    isMock: false,
    status: "ON_TIME"
  },
  {
    userId: "user_9",
    name: "Hendra Wijaya",
    dept: "Operations",
    checkIn: "07:55:40",
    checkOut: "-",
    distanceMeters: 41.5,
    isMock: false,
    status: "ON_TIME"
  },
  {
    userId: "user_10",
    name: "Putri Anggraini",
    dept: "Marketing & Growth",
    checkIn: "07:40:11",
    checkOut: "-",
    distanceMeters: 9.8,
    isMock: false,
    status: "ON_TIME"
  },
  {
    userId: "user_11",
    name: "Dimas Saputra",
    dept: "Finance & Accounting",
    checkIn: "07:59:05",
    checkOut: "-",
    distanceMeters: 14.2,
    isMock: false,
    status: "ON_TIME"
  },
  {
    userId: "user_12",
    name: "Anisa Fitri",
    dept: "HR & People",
    checkIn: "-",
    checkOut: "-",
    distanceMeters: 0,
    isMock: false,
    status: "ABSENT"
  }
];

// Leaflet Map Variables
let map;
let officeMarker;
let geofenceCircle;

document.addEventListener("DOMContentLoaded", () => {
  initDateHeader();
  initLeafletMap();
  renderAttendanceTable();
  updateStats();
});

function initDateHeader() {
  const options = { weekday: 'long', year: 'numeric', month: 'long', day: 'numeric' };
  const todayStr = new Date().toLocaleDateString('id-ID', options);
  document.getElementById('current-date-header').innerText = todayStr;
}

function initLeafletMap() {
  map = L.map('map').setView([officeConfig.lat, officeConfig.lng], 17);

  L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png', {
    maxZoom: 19,
    attribution: '© OpenStreetMap contributors'
  }).addTo(map);

  // Office Marker (Draggable)
  officeMarker = L.marker([officeConfig.lat, officeConfig.lng], {
    draggable: true
  }).addTo(map);

  officeMarker.bindPopup(`<b>${officeConfig.name}</b><br>Titik Utama Presensi`).openPopup();

  // Geofence Radius Circle
  geofenceCircle = L.circle([officeConfig.lat, officeConfig.lng], {
    color: '#10B981',
    fillColor: '#10B981',
    fillOpacity: 0.18,
    radius: officeConfig.radius
  }).addTo(map);

  // When marker is dragged
  officeMarker.on('dragend', function (event) {
    const position = officeMarker.getLatLng();
    officeConfig.lat = parseFloat(position.lat.toFixed(6));
    officeConfig.lng = parseFloat(position.lng.toFixed(6));

    document.getElementById('input-office-lat').value = officeConfig.lat;
    document.getElementById('input-office-lng').value = officeConfig.lng;
    geofenceCircle.setLatLng(position);
  });

  // When map is clicked
  map.on('click', function (e) {
    officeMarker.setLatLng(e.latlng);
    geofenceCircle.setLatLng(e.latlng);
    officeConfig.lat = parseFloat(e.latlng.lat.toFixed(6));
    officeConfig.lng = parseFloat(e.latlng.lng.toFixed(6));
    document.getElementById('input-office-lat').value = officeConfig.lat;
    document.getElementById('input-office-lng').value = officeConfig.lng;
  });
}

function handleRadiusSlider(val) {
  officeConfig.radius = parseInt(val);
  document.getElementById('radius-slider-val').innerText = `${val} Meter`;
  document.getElementById('header-radius-display').innerText = `${val} meter`;
  if (geofenceCircle) {
    geofenceCircle.setRadius(officeConfig.radius);
  }
}

function updateMapFromInputs() {
  const lat = parseFloat(document.getElementById('input-office-lat').value);
  const lng = parseFloat(document.getElementById('input-office-lng').value);
  if (!isNaN(lat) && !isNaN(lng)) {
    officeConfig.lat = lat;
    officeConfig.lng = lng;
    const newLatLng = new L.LatLng(lat, lng);
    officeMarker.setLatLng(newLatLng);
    geofenceCircle.setLatLng(newLatLng);
    map.panTo(newLatLng);
  }
}

function saveOfficeLocation() {
  officeConfig.name = document.getElementById('input-office-name').value;
  officeConfig.startTime = document.getElementById('input-office-start').value;
  officeConfig.endTime = document.getElementById('input-office-end').value;
  document.getElementById('office-name-badge').innerText = officeConfig.name;

  alert(`✅ Konfigurasi lokasi & radius kantor berhasil disimpan!\n\nNama: ${officeConfig.name}\nKoordinat: ${officeConfig.lat}, ${officeConfig.lng}\nRadius: ${officeConfig.radius} meter\nJam Kerja: ${officeConfig.startTime} - ${officeConfig.endTime}`);
}

function renderAttendanceTable(data = attendanceLogs) {
  const tbody = document.getElementById('attendance-table-body');
  tbody.innerHTML = '';

  if (data.length === 0) {
    tbody.innerHTML = `<tr><td colspan="7" class="py-8 text-center text-slate-400">Tidak ada data kehadiran yang cocok dengan pencarian.</td></tr>`;
    return;
  }

  data.forEach(item => {
    let statusBadge = '';
    if (item.status === 'ON_TIME') {
      statusBadge = `<span class="inline-flex items-center px-2.5 py-1 rounded-full text-[11px] font-semibold bg-emerald-50 text-emerald-700 border border-emerald-200">Tepat Waktu</span>`;
    } else if (item.status === 'LATE') {
      statusBadge = `<span class="inline-flex items-center px-2.5 py-1 rounded-full text-[11px] font-semibold bg-amber-50 text-amber-700 border border-amber-200">Terlambat</span>`;
    } else {
      statusBadge = `<span class="inline-flex items-center px-2.5 py-1 rounded-full text-[11px] font-semibold bg-slate-100 text-slate-600">Belum Presensi</span>`;
    }

    let mockBadge = item.isMock
      ? `<span class="inline-flex items-center text-rose-600 font-bold gap-1"><i data-lucide="shield-alert" class="w-3.5 h-3.5"></i> Fake GPS</span>`
      : `<span class="inline-flex items-center text-emerald-600 font-medium gap-1"><i data-lucide="shield-check" class="w-3.5 h-3.5"></i> Asli (Aman)</span>`;

    const tr = document.createElement('tr');
    tr.className = "hover:bg-slate-50/80 transition";
    tr.innerHTML = `
      <td class="py-3 px-4">
        <div class="font-bold text-slate-900">${item.name}</div>
        <div class="text-[11px] text-slate-400">ID: ${item.userId}</div>
      </td>
      <td class="py-3 px-4">${item.dept}</td>
      <td class="py-3 px-4 font-mono font-semibold ${item.checkIn !== '-' ? 'text-slate-900' : 'text-slate-400'}">${item.checkIn}</td>
      <td class="py-3 px-4 font-mono font-semibold ${item.checkOut !== '-' ? 'text-slate-900' : 'text-slate-400'}">${item.checkOut}</td>
      <td class="py-3 px-4">
        ${item.checkIn !== '-' ? `<span class="font-bold text-blue-600">${item.distanceMeters.toFixed(1)} m</span>` : '-'}
      </td>
      <td class="py-3 px-4">${item.checkIn !== '-' ? mockBadge : '-'}</td>
      <td class="py-3 px-4">${statusBadge}</td>
    `;
    tbody.appendChild(tr);
  });

  if (window.lucide) {
    lucide.createIcons();
  }
}

function filterAttendanceTable() {
  const query = document.getElementById('table-search').value.toLowerCase();
  const status = document.getElementById('status-filter').value;

  const filtered = attendanceLogs.filter(item => {
    const matchQuery = item.name.toLowerCase().includes(query) || item.dept.toLowerCase().includes(query);
    const matchStatus = status === 'ALL' || item.status === status;
    return matchQuery && matchStatus;
  });

  renderAttendanceTable(filtered);
}

function updateStats() {
  const total = employees.length;
  const onTime = attendanceLogs.filter(a => a.status === 'ON_TIME').length;
  const late = attendanceLogs.filter(a => a.status === 'LATE').length;
  const absent = attendanceLogs.filter(a => a.status === 'ABSENT').length;

  document.getElementById('stat-total-employees').innerText = total;
  document.getElementById('stat-on-time').innerText = onTime;
  document.getElementById('stat-late').innerText = late;
  document.getElementById('stat-absent').innerText = absent;
}

function exportToCSV() {
  let csvContent = "data:text/csv;charset=utf-8,";
  csvContent += "User ID,Nama Pegawai,Departemen,Jam Masuk,Jam Pulang,Jarak GPS (Meter),Deteksi Fake GPS,Status\n";

  attendanceLogs.forEach(row => {
    csvContent += `"${row.userId}","${row.name}","${row.dept}","${row.checkIn}","${row.checkOut}","${row.distanceMeters}","${row.isMock ? 'YA' : 'TIDAK'}","${row.status}"\n`;
  });

  const encodedUri = encodeURI(csvContent);
  const link = document.createElement("a");
  link.setAttribute("href", encodedUri);
  link.setAttribute("download", `Rekap_Presensi_${new Date().toISOString().slice(0,10)}.csv`);
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
}

// SIMULATOR FUNCTIONS
function handleSimDistanceSlider(val) {
  document.getElementById('sim-distance-val').innerText = `${val} Meter`;
}

function appendSimConsole(message, type = 'normal') {
  const consoleEl = document.getElementById('simulator-console');
  const div = document.createElement('div');
  const time = new Date().toLocaleTimeString('id-ID');
  
  if (type === 'error') {
    div.className = 'text-rose-400 font-semibold';
  } else if (type === 'success') {
    div.className = 'text-emerald-400 font-semibold';
  } else if (type === 'warning') {
    div.className = 'text-amber-300 font-semibold';
  } else {
    div.className = 'text-slate-300';
  }

  div.innerText = `[${time}] ${message}`;
  consoleEl.appendChild(div);
  consoleEl.scrollTop = consoleEl.scrollHeight;
}

function clearSimulatorLog() {
  document.getElementById('simulator-console').innerHTML = '<div class="text-slate-500">> Log dibersihkan. Siap untuk simulasi baru...</div>';
  document.getElementById('sim-status-banner').classList.add('hidden');
}

function runSimulatedCheckIn() {
  const empId = document.getElementById('sim-employee-select').value;
  const distance = parseFloat(document.getElementById('sim-distance-slider').value);
  const isMock = document.getElementById('sim-mock-gps').checked;
  const employee = employees.find(e => e.id === empId);

  appendSimConsole(`> Memulai request Check-in untuk: ${employee.name} (${employee.dept})...`);
  appendSimConsole(`> Koordinat GPS diterima. Jarak terhitung ke kantor: ${distance.toFixed(1)} meter.`);

  if (isMock) {
    appendSimConsole(`[SECURITY ALERT] Position.isMocked == true terdeteksi! Manipulasi GPS terdeteksi.`, 'error');
    appendSimConsole(`❌ Absen DITOLAK: Pegawai terdeteksi menggunakan Fake GPS / Mock Location!`, 'error');
    showSimBanner('Presensi DITOLAK: Deteksi Fake GPS!', 'error');
    return;
  }

  if (distance > officeConfig.radius) {
    appendSimConsole(`[VALIDATION FAILED] Jarak (${distance.toFixed(1)}m) melebihi batas radius kantor (${officeConfig.radius}m).`, 'warning');
    appendSimConsole(`❌ Absen DITOLAK: Pegawai berada di luar area kantor!`, 'warning');
    showSimBanner(`Presensi DITOLAK: Di luar radius (${distance.toFixed(1)}m > ${officeConfig.radius}m)`, 'warning');
    return;
  }

  // Success Check-in
  const now = new Date();
  const timeStr = now.toTimeString().split(' ')[0];
  const isLate = now.getHours() >= 8 && now.getMinutes() > 0;
  const status = isLate ? "LATE" : "ON_TIME";

  appendSimConsole(`[VALIDATION PASSED] Jarak ${distance.toFixed(1)}m <= ${officeConfig.radius}m & Lokasi Asli.`, 'success');
  appendSimConsole(`✅ Absen Masuk BERHASIL dicatat pada ${timeStr} WIB (Status: ${status}).`, 'success');
  showSimBanner(`Presensi Masuk BERHASIL (${timeStr} WIB)`, 'success');

  // Update in table data
  let target = attendanceLogs.find(a => a.userId === empId);
  if (target) {
    target.checkIn = timeStr;
    target.distanceMeters = distance;
    target.isMock = false;
    target.status = status;
  } else {
    attendanceLogs.unshift({
      userId: employee.id,
      name: employee.name,
      dept: employee.dept,
      checkIn: timeStr,
      checkOut: "-",
      distanceMeters: distance,
      isMock: false,
      status: status
    });
  }

  renderAttendanceTable();
  updateStats();
}

function runSimulatedCheckOut() {
  const empId = document.getElementById('sim-employee-select').value;
  const distance = parseFloat(document.getElementById('sim-distance-slider').value);
  const isMock = document.getElementById('sim-mock-gps').checked;
  const employee = employees.find(e => e.id === empId);

  appendSimConsole(`> Memulai request Check-out (Pulang) untuk: ${employee.name}...`);

  if (isMock) {
    appendSimConsole(`[SECURITY ALERT] Fake GPS terdeteksi saat check-out!`, 'error');
    showSimBanner('Absen Pulang DITOLAK: Fake GPS Terdeteksi!', 'error');
    return;
  }

  if (distance > officeConfig.radius) {
    appendSimConsole(`[VALIDATION FAILED] Check-out gagal: Di luar radius (${distance.toFixed(1)}m > ${officeConfig.radius}m).`, 'warning');
    showSimBanner(`Absen Pulang DITOLAK: Di luar radius (${distance.toFixed(1)}m)`, 'warning');
    return;
  }

  const now = new Date();
  const timeStr = now.toTimeString().split(' ')[0];

  let target = attendanceLogs.find(a => a.userId === empId);
  if (target) {
    target.checkOut = timeStr;
    appendSimConsole(`✅ Absen Pulang BERHASIL dicatat pada ${timeStr} WIB.`, 'success');
    showSimBanner(`Absen Pulang BERHASIL (${timeStr} WIB)`, 'success');
    renderAttendanceTable();
  } else {
    appendSimConsole(`❌ Gagal: Pegawai belum melakukan absen masuk hari ini.`, 'warning');
  }
}

function showSimBanner(text, type) {
  const banner = document.getElementById('sim-status-banner');
  banner.classList.remove('hidden', 'bg-emerald-900/60', 'text-emerald-300', 'border-emerald-700', 'bg-rose-900/60', 'text-rose-300', 'border-rose-700', 'bg-amber-900/60', 'text-amber-300', 'border-amber-700');
  banner.classList.add('border');

  if (type === 'success') {
    banner.classList.add('bg-emerald-900/60', 'text-emerald-300', 'border-emerald-700');
  } else if (type === 'error') {
    banner.classList.add('bg-rose-900/60', 'text-rose-300', 'border-rose-700');
  } else {
    banner.classList.add('bg-amber-900/60', 'text-amber-300', 'border-amber-700');
  }

  banner.innerText = text;
}
