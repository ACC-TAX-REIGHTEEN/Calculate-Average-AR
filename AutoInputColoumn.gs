function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("⚙️ Otomatisasi")
    .addItem("Perbarui / Isi Data AVG yang Kosong", "isiSemuaDataKosong")
    .addSeparator()
    .addItem("▶️ Aktifkan Sinkronisasi Latar Belakang", "pasangPemicuOtomatis")
    .addItem("⏹️ Nonaktifkan Sinkronisasi Latar Belakang", "hapusPemicuOtomatis")
    .addToUi();
}

function formatNoPelangganToUpper(rawKey) {
  if (!rawKey) return "";
  var parts = rawKey.toString().split(/[&,;/]+/).map(function(c) {
    return c.trim().toUpperCase().replace(/\s*-\s*/g, "-");
  }).filter(Boolean);

  return parts.join(" & ");
}

function getLookupMap(ss) {
  var lookupSheet = ss.getSheetByName("AVG AR");
  if (!lookupSheet) {
    Logger.log("Sheet 'AVG AR' tidak ditemukan.");
    return null;
  }

  var lastRow = lookupSheet.getLastRow();
  if (lastRow < 3) return null;

  var rawData = lookupSheet.getRange(3, 1, lastRow - 2, 5).getValues();
  var map = {};

  for (var i = 0; i < rawData.length; i++) {
    var rawCode = rawData[i][0] ? rawData[i][0].toString() : "";
    if (rawCode.trim() !== "") {
      var subCodes = rawCode.split(/[&,;/]+/).map(function(c) {
        return c.trim().toUpperCase().replace(/\s*-\s*/g, "-");
      }).filter(Boolean);

      for (var j = 0; j < subCodes.length; j++) {
        var key = subCodes[j];
        map[key] = {
          rowIndex: i + 3,
          values: [
            rawData[i][2],
            rawData[i][3],
            rawData[i][4]
          ]
        };
      }
    }
  }
  return map;
}

function cariDataTerbaru(rawKey, map) {
  if (!rawKey) return null;

  var codes = rawKey.toString().split(/[&,;/]+/).map(function(c) {
    return c.trim().toUpperCase().replace(/\s*-\s*/g, "-");
  }).filter(Boolean);

  var bestMatch = null;
  var maxRowIndex = -1;

  for (var k = 0; k < codes.length; k++) {
    var code = codes[k];
    if (map[code]) {
      if (map[code].rowIndex > maxRowIndex) {
        maxRowIndex = map[code].rowIndex;
        bestMatch = map[code].values;
      }
    }
  }

  return bestMatch;
}

function onEdit(e) {
  if (!e) return;
  var range = e.range;
  var sheet = range.getSheet();
  
  if (sheet.getName() !== "Form Responses 1") return;

  var colStart = range.getColumn();
  var colEnd = colStart + range.getNumColumns() - 1;

  if (colStart > 7 || colEnd < 7) return;

  var rowStart = Math.max(2, range.getRow());
  var numRows = range.getLastRow() - rowStart + 1;
  if (numRows <= 0) return;

  var ss = e.source;
  var map = getLookupMap(ss);
  if (!map) return;

  var targetRange = sheet.getRange(rowStart, 7, numRows, 4);
  var values = targetRange.getValues();

  for (var i = 0; i < values.length; i++) {
    var rawNoPelanggan = values[i][0] ? values[i][0].toString() : "";
    
    if (rawNoPelanggan.trim() === "") {
      values[i][0] = "";
      values[i][1] = "";
      values[i][2] = "";
      values[i][3] = "";
    } else {
      var formattedCode = formatNoPelangganToUpper(rawNoPelanggan);
      values[i][0] = formattedCode;

      var match = cariDataTerbaru(formattedCode, map);
      if (match) {
        values[i][1] = match[0];
        values[i][2] = match[1];
        values[i][3] = match[2];
      }
    }
  }

  targetRange.setValues(values);
}

function onFormSubmit(e) {
  var sheet;
  var row;

  if (e && e.range) {
    sheet = e.range.getSheet();
    row = e.range.getRow();
  } else {
    return;
  }

  if (sheet.getName() !== "Form Responses 1") return;

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var map = getLookupMap(ss);
  if (!map) return;

  var targetRange = sheet.getRange(row, 7, 1, 4);
  var values = targetRange.getValues();

  var rawNoPelanggan = values[0][0] ? values[0][0].toString() : "";
  if (rawNoPelanggan.trim() !== "") {
    var formattedCode = formatNoPelangganToUpper(rawNoPelanggan);
    values[0][0] = formattedCode;

    var match = cariDataTerbaru(formattedCode, map);
    if (match) {
      values[0][1] = match[0];
      values[0][2] = match[1];
      values[0][3] = match[2];
    }
    targetRange.setValues(values);
    Logger.log("onFormSubmit: Baris " + row + " berhasil diproses di latar belakang.");
  }
}

function isiSemuaDataKosong() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Form Responses 1");
  if (!sheet) {
    Logger.log("Sheet 'Form Responses 1' tidak ditemukan.");
    return;
  }

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  var map = getLookupMap(ss);
  if (!map) {
    Logger.log("Data pada sheet 'AVG AR' tidak ditemukan atau kosong.");
    return;
  }

  var totalRows = lastRow - 1;
  var dataRange = sheet.getRange(2, 7, totalRows, 4);
  var values = dataRange.getValues();
  var totalDiubah = 0;

  for (var i = 0; i < values.length; i++) {
    var rawNoPelanggan = values[i][0] ? values[i][0].toString() : "";
    if (rawNoPelanggan.trim() === "") continue;

    var formattedCode = formatNoPelangganToUpper(rawNoPelanggan);
    var barisBerubah = false;

    if (values[i][0] !== formattedCode) {
      values[i][0] = formattedCode;
      barisBerubah = true;
    }

    var avgUmur = values[i][1];
    var avgFaktur = values[i][2];
    var jmlInv = values[i][3];

    var isKosong = (avgUmur === "" || avgFaktur === "" || jmlInv === "" || 
                    avgUmur === null || avgFaktur === null || jmlInv === null);

    if (isKosong) {
      var match = cariDataTerbaru(formattedCode, map);
      if (match) {
        values[i][1] = match[0];
        values[i][2] = match[1];
        values[i][3] = match[2];
        barisBerubah = true;
      }
    }

    if (barisBerubah) {
      totalDiubah++;
    }
  }

  if (totalDiubah > 0) {
    dataRange.setValues(values);
    Logger.log("isiSemuaDataKosong: Berhasil memproses " + totalDiubah + " baris data.");
    try {
      ss.toast("Berhasil memproses " + totalDiubah + " baris data!", "Sukses", 4);
    } catch (e) {}
  } else {
    Logger.log("isiSemuaDataKosong: Tidak ada baris kosong atau perubahan baru.");
  }
}

function pasangPemicuOtomatis() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  hapusPemicuOtomatis();

  ScriptApp.newTrigger("onFormSubmit")
    .forSpreadsheet(ss)
    .onFormSubmit()
    .create();

  ScriptApp.newTrigger("isiSemuaDataKosong")
    .timeBased()
    .everyMinutes(10)
    .create();

  Logger.log("Pemicu latar belakang berhasil dipasang.");
  try {
    ss.toast("Pemicu aktif! Data akan otomatis terisi 24/7 tanpa perlu membuka sheet.", "Pemicu Terpasang", 5);
  } catch (e) {}
}

function hapusPemicuOtomatis() {
  var triggers = ScriptApp.getProjectTriggers();
  var count = 0;

  for (var i = 0; i < triggers.length; i++) {
    var fn = triggers[i].getHandlerFunction();
    if (fn === "onFormSubmit" || fn === "isiSemuaDataKosong") {
      ScriptApp.deleteTrigger(triggers[i]);
      count++;
    }
  }

  Logger.log("Sebanyak " + count + " pemicu berhasil dihapus.");
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  try {
    ss.toast("Pemicu latar belakang telah dinonaktifkan.", "Pemicu Dihapus", 4);
  } catch (e) {}
}
