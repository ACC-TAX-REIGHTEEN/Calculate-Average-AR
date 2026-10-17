function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("⚙️ Otomatisasi")
    .addItem("Perbarui / Isi Data AVG yang Kosong", "isiSemuaDataKosong")
    .addToUi();
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
    var key = rawData[i][0] ? rawData[i][0].toString().trim() : "";
    if (key !== "") {
      map[key] = [
        rawData[i][2],
        rawData[i][3],
        rawData[i][4] 
      ];
    }
  }
  return map;
}

function onEdit(e) {
  if (!e) return;
  var range = e.range;
  var sheet = range.getSheet();
  
  if (sheet.getName() !== "Form Responses 1") return;

  var colStart = range.getColumn();
  var colEnd = colStart + range.getNumColumns() - 1;

  if (colStart > 9 || colEnd < 9) return;

  var rowStart = Math.max(2, range.getRow());
  var numRows = range.getLastRow() - rowStart + 1;
  if (numRows <= 0) return;

  var ss = e.source;
  var map = getLookupMap(ss);
  if (!map) return;

  var targetRange = sheet.getRange(rowStart, 9, numRows, 4);
  var values = targetRange.getValues();

  for (var i = 0; i < values.length; i++) {
    var noPelanggan = values[i][0] ? values[i][0].toString().trim() : "";
    
    if (noPelanggan === "") {
      values[i][1] = "";
      values[i][2] = "";
      values[i][3] = "";
    } else if (map[noPelanggan]) {
      values[i][1] = map[noPelanggan][0];
      values[i][2] = map[noPelanggan][1];
      values[i][3] = map[noPelanggan][2];
    }
  }

  targetRange.setValues(values);
}

function isiSemuaDataKosong() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Form Responses 1");
  if (!sheet) {
    ss.toast("Sheet 'Form Responses 1' tidak ditemukan.", "Peringatan", 3);
    return;
  }

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  var map = getLookupMap(ss);
  if (!map) {
    ss.toast("Data pada sheet 'AVG AR' tidak ditemukan atau kosong.", "Peringatan", 3);
    return;
  }

  var totalRows = lastRow - 1;
  var dataRange = sheet.getRange(2, 9, totalRows, 4);
  var values = dataRange.getValues();
  var totalDiisi = 0;

  for (var i = 0; i < values.length; i++) {
    var noPelanggan = values[i][0] ? values[i][0].toString().trim() : "";
    var avgUmur = values[i][1];
    var avgFaktur = values[i][2];
    var jmlInv = values[i][3];

    var isKosong = (avgUmur === "" || avgFaktur === "" || jmlInv === "");

    if (noPelanggan !== "" && isKosong) {
      if (map[noPelanggan]) {
        values[i][1] = map[noPelanggan][0];
        values[i][2] = map[noPelanggan][1];
        values[i][3] = map[noPelanggan][2];
        totalDiisi++;
      }
    }
  }

  if (totalDiisi > 0) {
    dataRange.setValues(values);
    ss.toast("Berhasil mengisi " + totalDiisi + " baris data yang kosong!", "Sukses", 4);
  } else {
    ss.toast("Semua baris sudah terisi atau kode pelanggan tidak ditemukan di AVG AR.", "Informasi", 4);
  }
}
