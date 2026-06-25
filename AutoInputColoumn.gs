function onEdit(e) {
  var range = e.range;
  var sheet = range.getSheet();
  var sheetName = sheet.getName();
  
  // Validasi data untuk memastikan edit terjadi di sheet "Form Responses 1", di kolom H
  if (sheetName === "Form Responses 1" && range.getColumn() === 8 && range.getRow() > 1) {
    var row = range.getRow();
    var lookupValue = range.getValue();
    // Target tempel bisa diubah di sini
    var targetRange = sheet.getRange(row, 9, 1, 3);

    if (lookupValue === "") {
      targetRange.clearContent();
      return;
    }

    var ss = e.source;
    var lookupSheet = ss.getSheetByName("AVG AR");
    if (!lookupSheet) {
      Logger.log("Sheet 'AVG AR' tidak ditemukan.");
      return;
    }
    
    var lastRow = lookupSheet.getLastRow();
    if (lastRow < 2) return;
    
    var lookupData = lookupSheet.getRange(2, 1, lastRow - 1, 5).getValues();
    var found = false;
    for (var i = 0; i < lookupData.length; i++) {
      if (lookupData[i][0].toString().trim() === lookupValue.toString().trim()) {
        
        var avgUmurPiutang = lookupData[i][2];
        var avgNilaiFaktur = lookupData[i][3];
        var jumlahInvoice = lookupData[i][4];
        
        targetRange.setValues([[avgUmurPiutang, avgNilaiFaktur, jumlahInvoice]]);
        found = true;
        break;
      }
    }
    
    if (!found) {
      targetRange.clearContent();
    }
  }
}
