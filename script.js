const SPREADSHEET_ID = "1ds_AjC8Ri5JsjxyjFxtgpRatouGIsD3NdXdpa48a5oE";


/* =====================================================
   WEB APP
===================================================== */

function doGet(e) {

  const action =
    e && e.parameter
      ? e.parameter.action
      : "";


  // =========================
  // ABSENSI QR
  // =========================

  if (action === "absen") {

    return prosesAbsensi(
      e.parameter.id
    );

  }


  // =========================
  // LOGIN
  // =========================

  if (action === "login") {

    return prosesLogin(
      e.parameter.username,
      e.parameter.password
    );

  }


  // =========================
  // DASHBOARD ADMIN
  // =========================

  if (action === "dashboard") {

    return getDashboard();

  }


  // =========================
  // DATA GURU
  // =========================

  if (action === "guru") {

    return getDataGuru();

  }


  // =========================
  // HALAMAN UTAMA APPS SCRIPT
  // =========================

  return HtmlService
    .createHtmlOutputFromFile("Index")
    .setTitle("Absensi QR Siswa")
    .setXFrameOptionsMode(
      HtmlService.XFrameOptionsMode.ALLOWALL
    );

}


/* =====================================================
   PROSES ABSENSI QR
===================================================== */

function prosesAbsensi(id) {

  if (!id) {

    return jsonResponse({
      success: false,
      message: "QR Code tidak terbaca."
    });

  }


  const ss =
    SpreadsheetApp.openById(
      SPREADSHEET_ID
    );


  const sheetSiswa =
    ss.getSheetByName("Siswa");

  const sheetAbsensi =
    ss.getSheetByName("Absensi");


  if (!sheetSiswa || !sheetAbsensi) {

    return jsonResponse({
      success: false,
      message:
        "Sheet Siswa atau Absensi tidak ditemukan."
    });

  }


  /* =========================
     CARI DATA SISWA
  ========================= */

  const dataSiswa =
    sheetSiswa
      .getDataRange()
      .getValues();


  let siswa = null;


  for (
    let i = 1;
    i < dataSiswa.length;
    i++
  ) {

    const idSiswa =
      String(dataSiswa[i][0]).trim();


    if (
      idSiswa ===
      String(id).trim()
    ) {

      siswa = {

        id: idSiswa,

        nama: dataSiswa[i][1],

        noAbsen: dataSiswa[i][2]

      };


      break;

    }

  }


  /* =========================
     QR TIDAK TERDAFTAR
  ========================= */

  if (!siswa) {

    return jsonResponse({

      success: false,

      message:
        "QR Code tidak terdaftar."

    });

  }


  /* =========================
     CEK ABSEN HARI INI
  ========================= */

  const dataAbsensi =
    sheetAbsensi
      .getDataRange()
      .getValues();


  const sekarang =
    new Date();


  const tanggalHariIni =
    Utilities.formatDate(
      sekarang,
      Session.getScriptTimeZone(),
      "yyyy-MM-dd"
    );


  for (
    let i = 1;
    i < dataAbsensi.length;
    i++
  ) {

    const idAbsensi =
      String(dataAbsensi[i][0]).trim();


    if (
      idAbsensi !== siswa.id
    ) {

      continue;

    }


    const waktu =
      dataAbsensi[i][3];


    if (!waktu) {

      continue;

    }


    const tanggalAbsensi =
      Utilities.formatDate(
        new Date(waktu),
        Session.getScriptTimeZone(),
        "yyyy-MM-dd"
      );


    if (
      tanggalAbsensi ===
      tanggalHariIni
    ) {

      return jsonResponse({

        success: false,

        message:
          "Anda sudah melakukan absensi hari ini.",

        siswa: siswa,

        waktu:
          Utilities.formatDate(
            new Date(waktu),
            Session.getScriptTimeZone(),
            "dd/MM/yyyy HH:mm:ss"
          )

      });

    }

  }


  /* =========================
     SIMPAN ABSENSI
  ========================= */

  sheetAbsensi.appendRow([

    siswa.id,

    siswa.nama,

    siswa.noAbsen,

    sekarang

  ]);


  /* =========================
     RESPONSE
  ========================= */

  return jsonResponse({

    success: true,

    message:
      "Absensi berhasil.",

    siswa: siswa,

    waktu:
      Utilities.formatDate(
        sekarang,
        Session.getScriptTimeZone(),
        "dd/MM/yyyy HH:mm:ss"
      )

  });

}


/* =====================================================
   LOGIN ADMIN / GURU
===================================================== */

function prosesLogin(
  username,
  password
) {

  if (
    !username ||
    !password
  ) {

    return jsonResponse({

      success: false,

      message:
        "Username dan password wajib diisi."

    });

  }


  const ss =
    SpreadsheetApp.openById(
      SPREADSHEET_ID
    );


  const sheetAkun =
    ss.getSheetByName("Akun");


  if (!sheetAkun) {

    return jsonResponse({

      success: false,

      message:
        "Sheet Akun tidak ditemukan."

    });

  }


  const data =
    sheetAkun
      .getDataRange()
      .getValues();


  for (
    let i = 1;
    i < data.length;
    i++
  ) {

    const user =
      String(data[i][0]).trim();

    const pass =
      String(data[i][1]).trim();

    const role =
      String(data[i][2])
        .trim()
        .toUpperCase();


    if (
      user ===
        String(username).trim() &&
      pass ===
        String(password).trim()
    ) {

      return jsonResponse({

        success: true,

        message:
          "Login berhasil.",

        username: user,

        role: role

      });

    }

  }


  return jsonResponse({

    success: false,

    message:
      "Username atau password salah."

  });

}


/* =====================================================
   DASHBOARD ADMIN
===================================================== */

function getDashboard() {

  const ss =
    SpreadsheetApp.openById(
      SPREADSHEET_ID
    );


  const sheetSiswa =
    ss.getSheetByName("Siswa");

  const sheetAbsensi =
    ss.getSheetByName("Absensi");


  if (
    !sheetSiswa ||
    !sheetAbsensi
  ) {

    return jsonResponse({

      success: false,

      message:
        "Sheet tidak ditemukan."

    });

  }


  /* =========================
     TOTAL SISWA
  ========================= */

  const totalSiswa =
    Math.max(
      sheetSiswa.getLastRow() - 1,
      0
    );


  /* =========================
     DATA ABSENSI
  ========================= */

  const dataAbsensi =
    sheetAbsensi
      .getDataRange()
      .getValues();


  const sekarang =
    new Date();


  const tanggalHariIni =
    Utilities.formatDate(
      sekarang,
      Session.getScriptTimeZone(),
      "yyyy-MM-dd"
    );


  let hadir = 0;


  for (
    let i = 1;
    i < dataAbsensi.length;
    i++
  ) {

    const waktu =
      dataAbsensi[i][3];


    if (!waktu) {

      continue;

    }


    const tanggalAbsensi =
      Utilities.formatDate(
        new Date(waktu),
        Session.getScriptTimeZone(),
        "yyyy-MM-dd"
      );


    if (
      tanggalAbsensi ===
      tanggalHariIni
    ) {

      hadir++;

    }

  }


  /* =========================
     TIDAK HADIR
  ========================= */

  const tidakHadir =
    Math.max(
      totalSiswa - hadir,
      0
    );


  /* =========================
     PERSENTASE
  ========================= */

  let persentase = 0;


  if (
    totalSiswa > 0
  ) {

    persentase =
      (hadir / totalSiswa) *
      100;

  }


  return jsonResponse({

    success: true,

    total: totalSiswa,

    hadir: hadir,

    tidakHadir:
      tidakHadir,

    persentase:
      persentase.toFixed(1)

  });

}


/* =====================================================
   DATA GURU
===================================================== */

function getDataGuru() {

  const ss =
    SpreadsheetApp.openById(
      SPREADSHEET_ID
    );


  const sheetAbsensi =
    ss.getSheetByName(
      "Absensi"
    );


  if (!sheetAbsensi) {

    return jsonResponse({

      success: false,

      message:
        "Sheet Absensi tidak ditemukan."

    });

  }


  const lastRow =
    sheetAbsensi.getLastRow();


  if (
    lastRow <= 1
  ) {

    return jsonResponse({

      success: true,

      data: []

    });

  }


  const data =
    sheetAbsensi
      .getRange(
        2,
        1,
        lastRow - 1,
        4
      )
      .getValues();


  const hasil = [];


  /* =========================
     DATA TERBARU DI ATAS
  ========================= */

  for (
    let i = data.length - 1;
    i >= 0;
    i--
  ) {

    if (!data[i][0]) {

      continue;

    }


    let waktu = "";


    if (data[i][3]) {

      waktu =
        Utilities.formatDate(
          new Date(data[i][3]),
          Session.getScriptTimeZone(),
          "dd/MM/yyyy HH:mm:ss"
        );

    }


    hasil.push({

      id:
        data[i][0],

      nama:
        data[i][1],

      noAbsen:
        data[i][2],

      waktu:
        waktu

    });

  }


  return jsonResponse({

    success: true,

    data:
      hasil

  });

}


/* =====================================================
   RESPONSE JSON
===================================================== */

function jsonResponse(data) {

  return ContentService

    .createTextOutput(
      JSON.stringify(data)
    )

    .setMimeType(
      ContentService.MimeType.JSON
    );

}