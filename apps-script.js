const SHEET_NAME = 'Data';
const SETTINGS_SHEET_NAME = 'Settings';

const DEFAULT_SETTINGS = {
  badge: 'Data ini di awasi Pengurus kelas 12G',
  title: 'Villa Exsociothree',
  description: 'Silahkan lakukan pendataan keikutsertaan acara villa. Pastikan data yang dikirim sudah benar sebelum dikirim ke pusat.',
  paymentTitle: 'Pembayaran Villa',
  paymentDescription: 'Akses pembayaran hanya untuk peserta yang memilih mengikuti. Pilih nama, buka barcode pembayaran, lalu konfirmasi ke pengurus setelah transfer.',
  footer: '(c) exsociothree 2026. All Right Reserved.',
  barcodeUrl: 'https://i.ibb.co.com/YTFTh0j2/Whats-App-Image-2026-06-01-at-12-12-10.jpg',
  formDeadline: '2026-07-05T23:59:59+07:00',
  paymentDeadline: '2026-07-12T23:59:59+07:00'
};

function jsonOutput(payload) {
  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}

function getDataSheet() {
  return SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName(SHEET_NAME);
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
  const sheet = getDataSheet();
  const values = sheet.getDataRange().getValues();
  const data = [];

  for (let i = 1; i < values.length; i++) {
    data.push({
      no: values[i][0],
      nama: values[i][1],
      status: values[i][2],
      timestamp: values[i][3],
      pembayaran: values[i][4]
    });
  }

  return data;
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
