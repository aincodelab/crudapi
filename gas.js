// ============================================================
//  KONFIGURASI — ubah sesuai kebutuhan
// https://script.google.com/home/projects/15j8WRjMzzlUmRipt8pYaFqAxePZgGcBnGq5b7YyTtsxixEFbCd0BCFaf/edit
// ============================================================
const CONFIG = {
  SPREADSHEET_ID: '1RkvSr0SNx8D_mu3L9AcpetIKy9Hkgv92dVH47NQZcHQ', // Ganti dengan ID spreadsheet Anda
  SHEET_NAME: 'DATA',          // Nama sheet utama
  PAGE_SIZE: 10,               // Jumlah baris per halaman
  UNIQUE_COLUMN_INDEX: 0       // Indeks kolom unik (0-based)
};

// ============================================================
//  DO GET
// ============================================================
function doGet(e) {
  const action = e.parameter.action || 'read';
  
  if (action === 'config') {
    return respondJson({
      PAGE_SIZE: CONFIG.PAGE_SIZE,
      UNIQUE_COLUMN_INDEX: CONFIG.UNIQUE_COLUMN_INDEX
    });
  }
  
  if (action === 'read') {
    const page = parseInt(e.parameter.page) || 1;
    const pageSize = parseInt(e.parameter.pageSize) || CONFIG.PAGE_SIZE;
    const search = e.parameter.search || '';
    const row = e.parameter.row ? parseInt(e.parameter.row) : null;
    return handleRead(page, pageSize, search, row);
  }
  
  if (action === 'dropdown') {
    const field = e.parameter.field || '';
    return handleDropdown(field);
  }
  
  return respondError('Aksi tidak dikenal');
}

// ============================================================
//  DO POST
// ============================================================
function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const action = body.action || '';
    
    switch (action) {
      case 'create': return handleCreate(body);
      case 'update': return handleUpdate(body);
      case 'delete': return handleDelete(body);
      default: return respondError('Aksi tidak dikenal');
    }
  } catch (err) {
    return respondError('Format data tidak valid: ' + err.message);
  }
}

// ============================================================
//  HANDLE READ — dengan row asli (fisik) dari spreadsheet
// ============================================================
function handleRead(page, pageSize, search, rowNumber) {
  try {
    const sheet = getSheet();
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) {
      return respondJson({ headers: [], data: [], totalRecords: 0, page: 1, totalPages: 1 });
    }

    const headers = data[0];
    const rows = data.slice(1); // indeks 0 → baris 2 di sheet

    // Jika meminta satu baris spesifik berdasarkan rowNumber fisik
    if (rowNumber !== null) {
      const idx = rowNumber - 2;
      if (idx < 0 || idx >= rows.length) {
        return respondError('Baris tidak ditemukan');
      }
      const rowData = rows[idx];
      const obj = buildRowObject(headers, rowData, rowNumber);
      return respondJson({ headers, data: [obj], totalRecords: 1, page: 1, totalPages: 1 });
    }

    // Filter dengan menyimpan indeks asli (posisi dalam rows)
    let filtered = rows.map((row, index) => ({ index, row }));
    if (search.trim()) {
      const q = search.trim().toLowerCase();
      filtered = filtered.filter(item =>
        item.row.some(cell => String(cell).toLowerCase().includes(q))
      );
    }

    const total = filtered.length;
    const totalPages = Math.ceil(total / pageSize);
    const start = (page - 1) * pageSize;
    const end = Math.min(start + pageSize, total);
    const pageData = filtered.slice(start, end);

    // Bangun hasil dengan row fisik yang benar (index + 2)
    const resultData = pageData.map(({ index, row }) => {
      const sheetRow = index + 2; // baris fisik di spreadsheet
      return buildRowObject(headers, row, sheetRow);
    });

    return respondJson({
      headers: headers,
      data: resultData,
      totalRecords: total,
      page: page,
      totalPages: totalPages
    });
  } catch (err) {
    return respondError(err.message);
  }
}

// ============================================================
//  HANDLE DROPDOWN
// ============================================================
function handleDropdown(field) {
  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName(field);
    if (!sheet) {
      return respondJson({ options: [] });
    }
    
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) {
      return respondJson({ options: [] });
    }
    
    const headers = data[0];
    let idIdx = -1, uraianIdx = -1;
    headers.forEach((h, i) => {
      const lower = String(h).toLowerCase().trim();
      if (lower === 'id') idIdx = i;
      if (lower === 'uraian') uraianIdx = i;
    });
    
    if (idIdx === -1 || uraianIdx === -1) {
      return respondJson({ options: [] });
    }
    
    const options = data.slice(1).map(row => ({
      id: row[idIdx],
      uraian: row[uraianIdx]
    })).filter(opt => opt.id !== undefined && opt.id !== '');
    
    return respondJson({ options });
  } catch (err) {
    return respondError(err.message);
  }
}

// ============================================================
//  HANDLE CREATE — dengan pengecekan unik
// ============================================================
function handleCreate(body) {
  try {
    const sheet = getSheet();
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const newRow = headers.map(h => body[h] !== undefined ? body[h] : '');
    
    const uniqueIdx = CONFIG.UNIQUE_COLUMN_INDEX;
    if (uniqueIdx < newRow.length && newRow[uniqueIdx] !== '') {
      const existing = sheet.getDataRange().getValues();
      for (let i = 1; i < existing.length; i++) {
        if (String(existing[i][uniqueIdx]) === String(newRow[uniqueIdx])) {
          return respondError(`Nilai "${newRow[uniqueIdx]}" sudah ada`);
        }
      }
    }
    
    sheet.appendRow(newRow);
    return respondJson({ success: true, message: 'Data berhasil ditambahkan' });
  } catch (err) {
    return respondError(err.message);
  }
}

// ============================================================
//  HANDLE UPDATE — dengan dukungan row fisik atau uniqueValue (tidak menambah data)
// ============================================================
function handleUpdate(body) {
  try {
    const sheet = getSheet();
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) {
      return respondError('Tidak ada data untuk diperbarui');
    }
    const headers = data[0];
    const uniqueIdx = CONFIG.UNIQUE_COLUMN_INDEX;

    // 1. Tentukan nomor baris tujuan (fisik)
    let rowNumber = body.row;
    if (rowNumber && !isNaN(rowNumber) && rowNumber >= 2 && rowNumber <= data.length) {
      // row valid, gunakan langsung
    } else {
      // Coba cari berdasarkan uniqueValue (bisa dari properti 'uniqueValue' atau 'id')
      const uniqueVal = body.uniqueValue !== undefined ? body.uniqueValue : body.id;
      if (uniqueVal === undefined || uniqueVal === '') {
        return respondError('Parameter row atau uniqueValue harus disertakan');
      }
      rowNumber = findRowByUniqueValue(sheet, uniqueIdx, uniqueVal);
      if (rowNumber === -1) {
        return respondError(`Data dengan nilai unik "${uniqueVal}" tidak ditemukan`);
      }
    }

    const rowIndex = rowNumber - 1;
    const rowData = data[rowIndex];
    const oldRow = [...rowData];

    // Update nilai dari body
    headers.forEach((h, i) => {
      if (body[h] !== undefined) {
        rowData[i] = body[h];
      }
    });

    // Cek duplikat kolom unik jika nilai berubah
    if (uniqueIdx < rowData.length) {
      const newVal = rowData[uniqueIdx];
      const oldVal = oldRow[uniqueIdx];
      if (String(newVal) !== String(oldVal)) {
        for (let i = 1; i < data.length; i++) {
          if (i === rowIndex) continue;
          if (String(data[i][uniqueIdx]) === String(newVal)) {
            return respondError(`Nilai "${newVal}" sudah ada di baris lain`);
          }
        }
      }
    }

    // Tulis perubahan (tidak pernah menambah baris)
    sheet.getRange(rowNumber, 1, 1, rowData.length).setValues([rowData]);
    return respondJson({ success: true, message: 'Data berhasil diperbarui' });
  } catch (err) {
    return respondError(err.message);
  }
}

// ============================================================
//  HANDLE DELETE — dengan dukungan row fisik atau uniqueValue
// ============================================================
function handleDelete(body) {
  try {
    const sheet = getSheet();
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) {
      return respondError('Tidak ada data untuk dihapus');
    }
    const uniqueIdx = CONFIG.UNIQUE_COLUMN_INDEX;

    let rowNumber = body.row;
    if (rowNumber && !isNaN(rowNumber) && rowNumber >= 2 && rowNumber <= data.length) {
      // row valid
    } else {
      const uniqueVal = body.uniqueValue !== undefined ? body.uniqueValue : body.id;
      if (uniqueVal === undefined || uniqueVal === '') {
        return respondError('Parameter row atau uniqueValue harus disertakan');
      }
      rowNumber = findRowByUniqueValue(sheet, uniqueIdx, uniqueVal);
      if (rowNumber === -1) {
        return respondError(`Data dengan nilai unik "${uniqueVal}" tidak ditemukan`);
      }
    }

    sheet.deleteRow(rowNumber);
    return respondJson({ success: true, message: 'Data berhasil dihapus' });
  } catch (err) {
    return respondError(err.message);
  }
}

// ============================================================
//  FUNGSI PEMBANTU
// ============================================================

// Mencari nomor baris fisik (1-based) berdasarkan nilai di kolom unik
function findRowByUniqueValue(sheet, colIndex, value) {
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][colIndex]) === String(value)) {
      return i + 1; // nomor baris fisik
    }
  }
  return -1; // tidak ditemukan
}

function getSpreadsheet() {
  return SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
}

function getSheet() {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) throw new Error('Sheet utama tidak ditemukan');
  return sheet;
}

function buildRowObject(headers, rowData, rowNumber) {
  const obj = { row: rowNumber };
  headers.forEach((h, i) => {
    obj[h] = rowData[i];
  });
  return obj;
}

function respondJson(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function respondError(message) {
  return respondJson({ success: false, error: message });
}
