const SHEET_ID = '1U1JwfNE2u1EcUipf5fPIpHtU1t03Wl3oLrGWzB73zIw';
const SHEET_NAME = 'Transaksi';

function doGet(e) {
  const action = e.parameter.action;
  let result;

  if (action === 'get') {
    result = getTransaksi();
  } else if (action === 'add') {
    result = addTransaksi(e.parameter);
  } else {
    result = { ok: false, error: 'Unknown action' };
  }

  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

function getTransaksi() {
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);
    const rows  = sheet.getDataRange().getValues();
    if (rows.length < 2) return { ok: true, data: [] };

    const headers = rows[0].map(h => String(h).toLowerCase().trim());
    const data = rows.slice(1)
      .filter(r => r[headers.indexOf('id')])
      .map(r => {
        const obj = {};
        headers.forEach((h, i) => { obj[h] = r[i]; });

        // Format tanggal jadi string YYYY-MM-DD
        if (obj.tanggal instanceof Date) {
          const d = obj.tanggal;
          obj.tanggal = d.getFullYear() + '-' +
            String(d.getMonth()+1).padStart(2,'0') + '-' +
            String(d.getDate()).padStart(2,'0');
        }
        return obj;
      });

    return { ok: true, data };
  } catch (err) {
    return { ok: false, error: err.message };
  }
}

function addTransaksi(p) {
  try {
    const sheet   = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);
    const lastRow = sheet.getLastRow();

    // Ambil ID terakhir
    let lastId = 0;
    if (lastRow > 1) {
      const ids = sheet.getRange(2, 2, lastRow - 1, 1).getValues();
      lastId = Math.max(...ids.map(r => parseInt(r[0]) || 0));
    }

    const newId   = lastId + 1;
    const tanggal = new Date(p.tanggal);
    const bulan   = tanggal.toLocaleDateString('id-ID', { month: 'long', year: 'numeric' });

    sheet.appendRow([
      '',                  // kolom _RowNumber (kosong, AppSheet isi sendiri)
      newId,               // ID
      tanggal,             // Tanggal
      p.siapa,
      p.jenis,
      p.kategori,
      p.keterangan,
      parseInt(p.jumlah),
      p.metode,
      bulan                // Bulan
    ]);

    return { ok: true, id: newId };
  } catch (err) {
    return { ok: false, error: err.message };
  }
}
