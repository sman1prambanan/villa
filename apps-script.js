const SHEET_NAME = 'Data';
const SETTINGS_SHEET_NAME = 'Settings';

const DEFAULT_SETTINGS = {
  badge: 'Data ini di awasi Pengurus kelas 12G',
  title: 'Villa Exsociothree',
  description: 'Silahkan lakukan pendataan keikutsertaan acara villa. Pastikan data yang dikirim sudah benar sebelum dikirim ke pusat.',
  paymentTitle: 'Pembayaran Villa',
  paymentDescription: 'Akses pembayaran hanya untuk peserta yang memilih mengikuti. Pilih nama, buka barcode pembayaran, lalu konfirmasi ke pengurus setelah transfer.',
  announcementEnabled: 'true',
  announcementTitle: 'Pengumuman',
  announcementText: 'Selamat datang di web pendataan Villa Exsociothree. Silakan baca informasi terbaru sebelum mengisi data.',
  footer: '(c) exsociothree 2026. All Right Reserved.',
  barcodeUrl: 'https://i.ibb.co.com/YTFTh0j2/Whats-App-Image-2026-06-01-at-12-12-10.jpg',
  formDeadline: '2026-07-05T23:59:59+07:00',
  paymentDeadline: '2026-07-12T23:59:59+07:00',
  scheduleTitle: 'Jadwal Acara',
  scheduleSubtitle: 'Rangkaian kegiatan yang siap dijalankan',
  scheduleNote: 'Keputusan ini bersifat opsional dan dapat diubah.',
  scheduleRows: '15.00 - 17.00 | Registrasi Kedatangan | Peserta tiba untuk check-in.\n17.00 - 19.00 | Free Time & Ibadah | Peserta dibebaskan untuk beraktivitas dan ibadah.\n19.00 - 20.00 | Makan Malam Bersama | Makan malam dan keakraban bersama.\n20.00 - 23.00 | Malam Bahagia (MABA) | Pesta malam dan aktivitas menarik.\n23.00 - 04.30 | Free Time | Istirahat.\n04.00 - 05.30 | Ibadah | Peserta wajib menjalankan ibadah sholat subuh.\n05.30 - 07.00 | Giat Pagi | Peserta dibebaskan untuk beraktivitas pagi.\n07.00 - 08.00 | Sarapan Pagi | Peserta melakukan makan bersama.\n08.00 - 10.30 | Games | Peserta melakukan aktivitas permainan.\n10.30 - 11.30 | Bersih-Bersih | Peserta melakukan kegiatan kebersihan.\n11.30 - 15.00 | Penutupan | Sayonara dan perpisahan.'
};

function jsonOutput(payload) {
  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}

function requestDrivePermission() {
  const rootFolder = DriveApp.getRootFolder();
  Logger.log(`Drive access granted. Root folder: ${rootFolder.getName()}`);
}

function getDataSheet() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName(SHEET_NAME);

  if (!sheet) {
    throw new Error(`Sheet '${SHEET_NAME}' tidak ditemukan. Pastikan ada sheet bernama '${SHEET_NAME}'.`);
  }

  return sheet;
}

function getSettingsSheet() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = spreadsheet.getSheetByName(SETTINGS_SHEET_NAME);

  if (!sheet) {
    sheet = spreadsheet.insertSheet(SETTINGS_SHEET_NAME);
    sheet.getRange(1, 1, 1, 2).setValues([['key', 'value']]);
    saveSettings(DEFAULT_SETTINGS);
  }

  return sheet;
}

function getRows() {
  try {
    const sheet = getDataSheet();
    const values = sheet.getDataRange().getValues();
    const rows = [];

    for (let i = 1; i < values.length; i++) {
      rows.push({
        no: values[i][0],
        nama: values[i][1],
        status: values[i][2] || '',
        tglIsian: values[i][3] || '',
        pembayaran: values[i][4] || '',
        bukti: values[i][5] || '',
        tglProof: values[i][6] || ''
      });
    }

    return rows;
  }
  catch (error) {
    Logger.log('getRows error: ' + error);
    return [];
  }
}

function saveSettings(settings) {
  const sheet = getSettingsSheet();
  const mergedSettings = Object.assign({}, DEFAULT_SETTINGS, settings || {});
  const rows = Object.keys(mergedSettings).map(function(key) {
    return [key, mergedSettings[key]];
  });

  sheet.clearContents();
  sheet.getRange(1, 1, 1, 2).setValues([['key', 'value']]);

  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, 2).setValues(rows);
  }

  return mergedSettings;
}

function getSettings() {
  const sheet = getSettingsSheet();
  const values = sheet.getDataRange().getValues();
  const settings = Object.assign({}, DEFAULT_SETTINGS);

  for (let i = 1; i < values.length; i++) {
    const key = values[i][0];
    const value = values[i][1];
    if (key) {
      settings[key] = value;
    }
  }

  return settings;
}

function findRowByName(sheet, nama) {
  const values = sheet.getDataRange().getValues();

  for (let i = 1; i < values.length; i++) {
    if (String(values[i][1]).trim() === String(nama).trim()) {
      return {
        rowIndex: i + 1,
        values: values[i]
      };
    }
  }

  return null;
}

function submitPublicData(data) {
  const nama = data.nama;
  const status = data.status;
  const sheet = getDataSheet();
  const found = findRowByName(sheet, nama);

  if (!found) {
    return {
      success: false,
      message: 'Nama tidak ditemukan.'
    };
  }

  if (found.values[2] !== '') {
    return {
      success: false,
      message: 'Data sudah pernah diisi.'
    };
  }

  sheet.getRange(found.rowIndex, 3).setValue(status);
  sheet.getRange(found.rowIndex, 4).setValue(new Date());

  return {
    success: true,
    message: 'Data berhasil dikirim ke pusat.'
  };
}

function updateRow(data) {
  const sheet = getDataSheet();
  const found = findRowByName(sheet, data.nama);

  if (!found) {
    return {
      success: false,
      message: 'Nama tidak ditemukan.'
    };
  }

  sheet.getRange(found.rowIndex, 3).setValue(data.status || '');
  sheet.getRange(found.rowIndex, 4).setValue(data.status ? new Date() : '');
  sheet.getRange(found.rowIndex, 5).setValue(data.pembayaran || data.statusPembayaran || '');

  return {
    success: true,
    message: 'Data berhasil disimpan ke server.'
  };
}

function uploadPaymentProof(data) {
  const sheet = getDataSheet();
  const found = findRowByName(sheet, data.nama);

  if (!found) {
    return {
      success: false,
      message: 'Nama tidak ditemukan.'
    };
  }

  const fileBase64 = typeof data.fileBase64 === 'string' ? data.fileBase64.trim() : '';

  if (!fileBase64) {
    return {
      success: false,
      message: 'Bukti pembayaran belum diterima.'
    };
  }

  try {
    const mimeType = data.mimeType || 'image/png';
    const fileName = String(data.fileName || `${data.nama}-bukti-pembayaran`).replace(/[^\w.\- ]+/g, '_');
    const blob = Utilities.newBlob(Utilities.base64Decode(fileBase64), mimeType, fileName);
    const file = DriveApp.createFile(blob);

    sheet.getRange(found.rowIndex, 5).setValue('Menunggu Konfirmasi');
    sheet.getRange(found.rowIndex, 6).setValue(file.getUrl());
    sheet.getRange(found.rowIndex, 7).setValue(new Date());

    return {
      success: true,
      message: 'Bukti pembayaran berhasil dikirim.'
    };
  }
  catch (error) {
    Logger.log(`uploadPaymentProof error: ${error}`);
    return {
      success: false,
      message: 'Gagal menyimpan bukti pembayaran. ' + (error.message || error.toString()),
      error: error.toString()
    };
  }
}

function confirmPayment(data) {
  const sheet = getDataSheet();
  const found = findRowByName(sheet, data.nama);

  if (!found) {
    return {
      success: false,
      message: 'Nama tidak ditemukan.'
    };
  }

  sheet.getRange(found.rowIndex, 5).setValue('Lunas');

  return {
    success: true,
    message: 'Pembayaran berhasil dikonfirmasi.'
  };
}

function updatePaymentStatus(data) {
  const sheet = getDataSheet();
  const found = findRowByName(sheet, data.nama);

  if (!found) {
    return {
      success: false,
      message: 'Nama tidak ditemukan.'
    };
  }

  const status = data.status || 'Menunggu Konfirmasi';
  sheet.getRange(found.rowIndex, 5).setValue(status);
  sheet.getRange(found.rowIndex, 7).setValue(new Date());

  return {
    success: true,
    message: 'Status pembayaran berhasil diperbarui.'
  };
}

function deleteRow(data) {
  const sheet = getDataSheet();
  const found = findRowByName(sheet, data.nama);

  if (!found) {
    return {
      success: false,
      message: 'Nama tidak ditemukan.'
    };
  }

  sheet.deleteRow(found.rowIndex);

  return {
    success: true,
    message: 'Baris berhasil dihapus.'
  };
}

function doGet(e) {
  const action = e && e.parameter ? e.parameter.action : '';

  if (action === 'getSettings') {
    return jsonOutput({
      success: true,
      settings: getSettings()
    });
  }

  return jsonOutput(getRows());
}

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents || '{}');
    const action = data.action || 'submit';

    if (action === 'getSettings') {
      return jsonOutput({
        success: true,
        settings: getSettings()
      });
    }

    if (action === 'saveSettings') {
      return jsonOutput({
        success: true,
        message: 'Pengaturan berhasil disimpan ke server.',
        settings: saveSettings(data.settings)
      });
    }

    if (action === 'updateRow') {
      return jsonOutput(updateRow(data));
    }

    if (action === 'uploadPaymentProof') {
      return jsonOutput(uploadPaymentProof(data));
    }

    if (action === 'confirmPayment') {
      return jsonOutput(confirmPayment(data));
    }

    if (action === 'updatePaymentStatus') {
      return jsonOutput(updatePaymentStatus(data));
    }

    if (action === 'delete') {
      return jsonOutput(deleteRow(data));
    }

    return jsonOutput(submitPublicData(data));
  }
  catch (error) {
    return jsonOutput({
      success: false,
      message: 'Terjadi kesalahan sistem.',
      error: error.toString()
    });
  }
}
