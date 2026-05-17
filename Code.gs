// ============================================================
//  SISTEM ABSENSI SISWA ONLINE — Google Apps Script
//  Versi: 1.0 | Siap Pakai
// ============================================================

// ===== KONFIGURASI UTAMA =====
// Ganti dengan koordinat GPS sekolah kamu
// Cara cari: buka Google Maps → klik kanan lokasi sekolah → "Apa yang ada di sini?"
var SEKOLAH_LAT  = -0.5477;   // <-- GANTI dengan latitude sekolah
var SEKOLAH_LNG  = 123.0596;  // <-- GANTI dengan longitude sekolah
var RADIUS_METER = 100;        // Radius absen (meter)

// Jam absen (format 24 jam)
var JAM_DATANG_MULAI = 6;   // 06.00
var JAM_DATANG_BATAS = 7;   // 07.30 (jam 7, menit 30)
var MENIT_DATANG_BATAS = 30;
var JAM_PULANG_MULAI = 15;  // 15.00

// Nama sheet
var SHEET_SISWA   = "DATA_SISWA";
var SHEET_ABSENSI = "ABSENSI";
var SHEET_LOG     = "LOG";

// ============================================================
//  INISIALISASI SHEET (jalankan sekali saat pertama kali setup)
// ============================================================
function initSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // Sheet DATA_SISWA
  var sheetSiswa = ss.getSheetByName(SHEET_SISWA);
  if (!sheetSiswa) {
    sheetSiswa = ss.insertSheet(SHEET_SISWA);
    sheetSiswa.appendRow(["NIS", "Nama Lengkap", "Tanggal Daftar"]);
    sheetSiswa.getRange(1, 1, 1, 3).setFontWeight("bold").setBackground("#1565C0").setFontColor("#FFFFFF");
    sheetSiswa.setColumnWidth(1, 120);
    sheetSiswa.setColumnWidth(2, 220);
    sheetSiswa.setColumnWidth(3, 180);
  }

  // Sheet ABSENSI
  var sheetAbsen = ss.getSheetByName(SHEET_ABSENSI);
  if (!sheetAbsen) {
    sheetAbsen = ss.insertSheet(SHEET_ABSENSI);
    sheetAbsen.appendRow(["Hari/Tanggal", "NIS", "Nama Siswa", "Waktu Absen Datang", "Waktu Absen Pulang", "Keterangan"]);
    sheetAbsen.getRange(1, 1, 1, 6).setFontWeight("bold").setBackground("#1565C0").setFontColor("#FFFFFF");
    sheetAbsen.setColumnWidth(1, 160);
    sheetAbsen.setColumnWidth(2, 120);
    sheetAbsen.setColumnWidth(3, 220);
    sheetAbsen.setColumnWidth(4, 160);
    sheetAbsen.setColumnWidth(5, 160);
    sheetAbsen.setColumnWidth(6, 220);
  }

  // Sheet LOG
  var sheetLog = ss.getSheetByName(SHEET_LOG);
  if (!sheetLog) {
    sheetLog = ss.insertSheet(SHEET_LOG);
    sheetLog.appendRow(["Waktu", "Aksi", "NIS", "Detail"]);
    sheetLog.getRange(1, 1, 1, 4).setFontWeight("bold").setBackground("#37474F").setFontColor("#FFFFFF");
  }

  return "Inisialisasi sheet selesai!";
}

// ============================================================
//  ENTRY POINT — Handle HTTP Request
// ============================================================
function doPost(e) {
  try {
    var params = JSON.parse(e.postData.contents);
    var aksi   = params.aksi;
    var hasil;

    if      (aksi === "daftar")       hasil = registerSiswa(params);
    else if (aksi === "login")        hasil = loginSiswa(params);
    else if (aksi === "absenDatang")  hasil = absenDatang(params);
    else if (aksi === "absenPulang")  hasil = absenPulang(params);
    else if (aksi === "riwayat")      hasil = getRiwayatAbsen(params);
    else if (aksi === "statusHari")   hasil = getStatusHariIni(params);
    else hasil = resError("Aksi tidak dikenal: " + aksi);

    return resJSON(hasil);
  } catch (err) {
    tulisBaris(SHEET_LOG, [waktuSekarang(), "ERROR", "-", err.toString()]);
    return resJSON(resError("Terjadi kesalahan sistem: " + err.toString()));
  }
}

function doGet(e) {
  return ContentService.createTextOutput("Sistem Absensi Siswa Online - API Aktif").setMimeType(ContentService.MimeType.TEXT);
}

// ============================================================
//  1. DAFTAR AKUN SISWA
// ============================================================
function registerSiswa(p) {
  var nis  = String(p.nis  || "").trim();
  var nama = String(p.nama || "").trim();

  if (!nis || !nama) return resError("NIS dan Nama wajib diisi.");

  var sheet = getSheet(SHEET_SISWA);
  var data  = sheet.getDataRange().getValues();

  // Cek apakah NIS sudah terdaftar (mulai baris ke-2, lewati header)
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === nis) {
      return resError("Anda sudah terdaftar! NIS " + nis + " sudah ada dalam sistem.");
    }
  }

  // Simpan data siswa baru
  sheet.appendRow([nis, nama, waktuSekarang()]);
  tulisBaris(SHEET_LOG, [waktuSekarang(), "DAFTAR", nis, "Siswa baru: " + nama]);

  return resOK("Pendaftaran berhasil! Selamat datang, " + nama + ".", { nis: nis, nama: nama });
}

// ============================================================
//  2. LOGIN SISWA
// ============================================================
function loginSiswa(p) {
  var nis = String(p.nis || "").trim();
  if (!nis) return resError("NIS wajib diisi.");

  var siswa = cariSiswa(nis);
  if (!siswa) return resError("NIS tidak ditemukan. Silakan daftar terlebih dahulu.");

  tulisBaris(SHEET_LOG, [waktuSekarang(), "LOGIN", nis, "Login berhasil"]);
  return resOK("Login berhasil.", { nis: siswa[0], nama: siswa[1] });
}

// ============================================================
//  3. ABSEN DATANG
// ============================================================
function absenDatang(p) {
  var nis = String(p.nis || "").trim();
  var lat = parseFloat(p.latitude);
  var lng = parseFloat(p.longitude);

  // Validasi input
  if (!nis) return resError("NIS wajib diisi.");
  if (isNaN(lat) || isNaN(lng)) return resError("Data lokasi tidak valid.");

  // Validasi siswa terdaftar
  var siswa = cariSiswa(nis);
  if (!siswa) return resError("NIS tidak ditemukan. Silakan daftar terlebih dahulu.");

  // Validasi lokasi
  var jarak = hitungJarak(lat, lng, SEKOLAH_LAT, SEKOLAH_LNG);
  if (jarak > RADIUS_METER) {
    return resError("Anda berada di luar area sekolah. Jarak Anda: " + Math.round(jarak) + " meter dari sekolah. Maksimal: " + RADIUS_METER + " meter.");
  }

  // Validasi jam
  var now    = new Date();
  var jam    = now.getHours();
  var menit  = now.getMinutes();

  if (jam < JAM_DATANG_MULAI) {
    return resError("Absen datang belum bisa dilakukan. Mulai pukul 06.00 WIB.");
  }

  // Cek apakah sudah absen datang hari ini
  var tanggalHariIni = formatTanggal(now);
  var barisAbsen = cariBarisSiswaHariIni(nis, tanggalHariIni);
  if (barisAbsen && barisAbsen.waktuDatang) {
    return resError("Anda sudah melakukan absen datang hari ini pada pukul " + barisAbsen.waktuDatang + ".");
  }

  // Tentukan keterangan
  var tepat = (jam < JAM_DATANG_BATAS) || (jam === JAM_DATANG_BATAS && menit <= MENIT_DATANG_BATAS);
  var keterangan = tepat ? "Datang Tepat Waktu" : "Datang Terlambat";
  var waktuAbsen = formatJam(now);
  var hariTanggal = formatHariTanggal(now);

  // Simpan ke sheet ABSENSI
  var sheetAbsen = getSheet(SHEET_ABSENSI);
  sheetAbsen.appendRow([hariTanggal, nis, siswa[1], waktuAbsen, "", keterangan]);

  // Warna baris sesuai keterangan
  var barisKe = sheetAbsen.getLastRow();
  var warna = tepat ? "#E8F5E9" : "#FFF8E1";
  sheetAbsen.getRange(barisKe, 1, 1, 6).setBackground(warna);

  tulisBaris(SHEET_LOG, [waktuSekarang(), "ABSEN_DATANG", nis, keterangan + " | Jarak: " + Math.round(jarak) + "m"]);

  return resOK("Absen datang berhasil! " + keterangan, {
    nama: siswa[1],
    nis: nis,
    waktu: waktuAbsen,
    keterangan: keterangan,
    jarak: Math.round(jarak)
  });
}

// ============================================================
//  4. ABSEN PULANG
// ============================================================
function absenPulang(p) {
  var nis = String(p.nis || "").trim();
  var lat = parseFloat(p.latitude);
  var lng = parseFloat(p.longitude);

  if (!nis) return resError("NIS wajib diisi.");
  if (isNaN(lat) || isNaN(lng)) return resError("Data lokasi tidak valid.");

  var siswa = cariSiswa(nis);
  if (!siswa) return resError("NIS tidak ditemukan.");

  // Validasi lokasi
  var jarak = hitungJarak(lat, lng, SEKOLAH_LAT, SEKOLAH_LNG);
  if (jarak > RADIUS_METER) {
    return resError("Anda berada di luar area sekolah. Jarak Anda: " + Math.round(jarak) + " meter dari sekolah.");
  }

  // Validasi jam >= 15.00
  var now  = new Date();
  var jam  = now.getHours();
  if (jam < JAM_PULANG_MULAI) {
    return resError("Absen pulang belum bisa dilakukan. Baru bisa mulai pukul 15.00 WIB.");
  }

  // Cari baris absen datang hari ini
  var tanggalHariIni = formatTanggal(now);
  var info = cariBarisSiswaHariIni(nis, tanggalHariIni);

  if (!info) {
    return resError("Anda belum melakukan absen datang hari ini.");
  }
  if (info.waktuPulang) {
    return resError("Anda sudah melakukan absen pulang hari ini pada pukul " + info.waktuPulang + ".");
  }

  // Update kolom Waktu Pulang dan Keterangan
  var waktuAbsen = formatJam(now);
  var sheetAbsen = getSheet(SHEET_ABSENSI);
  sheetAbsen.getRange(info.baris, 5).setValue(waktuAbsen);

  // Tentukan keterangan akhir
  var ketDatang = info.keteranganDatang;
  var ketAkhir  = (ketDatang === "Datang Terlambat") ? "Datang Terlambat" : "Datang Tepat Waktu";
  sheetAbsen.getRange(info.baris, 6).setValue(ketAkhir);
  sheetAbsen.getRange(info.baris, 1, 1, 6).setBackground("#E8F5E9");

  tulisBaris(SHEET_LOG, [waktuSekarang(), "ABSEN_PULANG", nis, "Pulang pukul " + waktuAbsen]);

  return resOK("Absen pulang berhasil! Selamat pulang, " + siswa[1] + ".", {
    nama: siswa[1],
    nis: nis,
    waktuPulang: waktuAbsen,
    keterangan: ketAkhir
  });
}

// ============================================================
//  5. STATUS ABSEN HARI INI
// ============================================================
function getStatusHariIni(p) {
  var nis = String(p.nis || "").trim();
  if (!nis) return resError("NIS wajib diisi.");

  var siswa = cariSiswa(nis);
  if (!siswa) return resError("NIS tidak ditemukan.");

  var tanggalHariIni = formatTanggal(new Date());
  var info = cariBarisSiswaHariIni(nis, tanggalHariIni);

  if (!info) {
    return resOK("Belum ada absen hari ini.", {
      sudahDatang: false,
      sudahPulang: false,
      waktuDatang: null,
      waktuPulang: null,
      keterangan: null
    });
  }

  return resOK("Status absen hari ini.", {
    sudahDatang: !!info.waktuDatang,
    sudahPulang: !!info.waktuPulang,
    waktuDatang: info.waktuDatang || null,
    waktuPulang: info.waktuPulang || null,
    keterangan: info.keteranganDatang
  });
}

// ============================================================
//  6. RIWAYAT ABSEN SISWA
// ============================================================
function getRiwayatAbsen(p) {
  var nis   = String(p.nis || "").trim();
  var limit = parseInt(p.limit) || 10;
  if (!nis) return resError("NIS wajib diisi.");

  var sheet  = getSheet(SHEET_ABSENSI);
  var data   = sheet.getDataRange().getValues();
  var hasil  = [];

  for (var i = data.length - 1; i >= 1; i--) {
    if (String(data[i][1]).trim() === nis) {
      hasil.push({
        tanggal: data[i][0],
        nis: data[i][1],
        nama: data[i][2],
        waktuDatang: data[i][3],
        waktuPulang: data[i][4],
        keterangan: data[i][5]
      });
      if (hasil.length >= limit) break;
    }
  }

  return resOK("Riwayat absen ditemukan.", { total: hasil.length, data: hasil });
}

// ============================================================
//  7. TRIGGER OTOMATIS MALAM (daftarkan sebagai Time Trigger)
// ============================================================
function triggerMalam() {
  var ss          = SpreadsheetApp.getActiveSpreadsheet();
  var sheetSiswa  = ss.getSheetByName(SHEET_SISWA);
  var sheetAbsen  = ss.getSheetByName(SHEET_ABSENSI);
  var now         = new Date();
  var hariTanggal = formatHariTanggal(now);
  var tanggal     = formatTanggal(now);

  var dataSiswa   = sheetSiswa.getDataRange().getValues();
  var dataAbsen   = sheetAbsen.getDataRange().getValues();

  // Kumpulkan NIS yang sudah absen hari ini
  var nisSudahAbsen = {};
  for (var i = 1; i < dataAbsen.length; i++) {
    var tgl = String(dataAbsen[i][0]);
    var nis = String(dataAbsen[i][1]).trim();
    if (tgl.indexOf(tanggal) !== -1 || tgl === hariTanggal) {
      nisSudahAbsen[nis] = { baris: i + 1, pulang: dataAbsen[i][4] };
    }
  }

  // Loop semua siswa
  for (var j = 1; j < dataSiswa.length; j++) {
    var nisSiswa = String(dataSiswa[j][0]).trim();
    var namaSiswa = String(dataSiswa[j][1]).trim();

    if (!nisSiswa) continue;

    if (nisSudahAbsen[nisSiswa]) {
      // Sudah absen datang, cek apakah sudah pulang
      if (!nisSudahAbsen[nisSiswa].pulang) {
        sheetAbsen.getRange(nisSudahAbsen[nisSiswa].baris, 5).setValue("Tidak absen");
        sheetAbsen.getRange(nisSudahAbsen[nisSiswa].baris, 6).setValue("Cepat Pulang/Bolos");
        sheetAbsen.getRange(nisSudahAbsen[nisSiswa].baris, 1, 1, 6).setBackground("#FFEBEE");
      }
    } else {
      // Sama sekali belum absen hari ini
      sheetAbsen.appendRow([hariTanggal, nisSiswa, namaSiswa, "", "", "Tidak Hadir"]);
      var barisKe = sheetAbsen.getLastRow();
      sheetAbsen.getRange(barisKe, 1, 1, 6).setBackground("#FFEBEE");
    }
  }

  tulisBaris(SHEET_LOG, [waktuSekarang(), "TRIGGER_MALAM", "-", "Rekap harian selesai diproses."]);
}

// ============================================================
//  HELPER FUNCTIONS
// ============================================================

function hitungJarak(lat1, lon1, lat2, lon2) {
  var R     = 6371000;
  var dLat  = (lat2 - lat1) * Math.PI / 180;
  var dLon  = (lon2 - lon1) * Math.PI / 180;
  var a     = Math.sin(dLat / 2) * Math.sin(dLat / 2) +
              Math.cos(lat1 * Math.PI / 180) * Math.cos(lat2 * Math.PI / 180) *
              Math.sin(dLon / 2) * Math.sin(dLon / 2);
  var c = 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a));
  return R * c;
}

function cariSiswa(nis) {
  var sheet = getSheet(SHEET_SISWA);
  var data  = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === nis) return data[i];
  }
  return null;
}

function cariBarisSiswaHariIni(nis, tanggal) {
  var sheet = getSheet(SHEET_ABSENSI);
  var data  = sheet.getDataRange().getValues();
  for (var i = data.length - 1; i >= 1; i--) {
    var tgl = String(data[i][0]);
    if (String(data[i][1]).trim() === nis && (tgl.indexOf(tanggal) !== -1 || tgl === formatHariTanggalDariStr(tanggal))) {
      return {
        baris: i + 1,
        waktuDatang: data[i][3],
        waktuPulang: data[i][4],
        keteranganDatang: data[i][5]
      };
    }
  }
  return null;
}

function getSheet(nama) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(nama);
  if (!sheet) {
    sheet = ss.insertSheet(nama);
    if (nama === SHEET_SISWA)   sheet.appendRow(["NIS", "Nama Lengkap", "Tanggal Daftar"]);
    if (nama === SHEET_ABSENSI) sheet.appendRow(["Hari/Tanggal", "NIS", "Nama Siswa", "Waktu Absen Datang", "Waktu Absen Pulang", "Keterangan"]);
    if (nama === SHEET_LOG)     sheet.appendRow(["Waktu", "Aksi", "NIS", "Detail"]);
  }
  return sheet;
}

function tulisBaris(sheetNama, baris) {
  try { getSheet(sheetNama).appendRow(baris); } catch(e) {}
}

function waktuSekarang() {
  return Utilities.formatDate(new Date(), "Asia/Makassar", "dd/MM/yyyy HH:mm:ss");
}

function formatTanggal(date) {
  return Utilities.formatDate(date, "Asia/Makassar", "dd/MM/yyyy");
}

function formatJam(date) {
  return Utilities.formatDate(date, "Asia/Makassar", "HH:mm:ss");
}

function formatHariTanggal(date) {
  var hari = ["Minggu","Senin","Selasa","Rabu","Kamis","Jumat","Sabtu"];
  var d    = new Date(date.toLocaleString("en-US", {timeZone: "Asia/Makassar"}));
  return hari[d.getDay()] + ", " + formatTanggal(date);
}

function formatHariTanggalDariStr(tanggal) {
  return tanggal; // fallback
}

function resOK(pesan, data) {
  return { status: "success", message: pesan, data: data || {} };
}

function resError(pesan) {
  return { status: "error", message: pesan, data: {} };
}

function resJSON(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
