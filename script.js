function getSheet() {
  return SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
}

// Fungsi saat website meminta data (Load)
function doGet(e) {
  var sheet = getSheet();
  var rows = sheet.getDataRange().getValues();
  var headers = rows[0];
  var requestedId = e && e.parameter && e.parameter.id ? e.parameter.id : "1";

  for (var i = 1; i < rows.length; i++) {
    if (rows[i][0].toString() === requestedId.toString()) {
      var obj = {};
      for (var j = 0; j < headers.length; j++) {
        obj[headers[j]] = rows[i][j];
      }
      return ContentService.createTextOutput(JSON.stringify(obj))
        .setMimeType(ContentService.MimeType.JSON);
    }
  }
  return ContentService.createTextOutput(JSON.stringify({status: "not found"}))
    .setMimeType(ContentService.MimeType.JSON);
}

// Fungsi saat website mengirim data untuk disimpan (Save/Edit)
function doPost(e) {
  try {
    var sheet = getSheet();
    var data = JSON.parse(e.postData.contents);
    var rows = sheet.getDataRange().getValues();
    var targetId = data.id.toString();
    
    // Cari baris berdasarkan ID, lalu update datanya
    for (var i = 1; i < rows.length; i++) {
      if (rows[i][0].toString() === targetId) {
        sheet.getRange(i + 1, 2).setValue(data.nama);       // Kolom B
        sheet.getRange(i + 1, 3).setValue(data.jabatan);    // Kolom C
        sheet.getRange(i + 1, 5).setValue(data.whatsapp);   // Kolom E
        sheet.getRange(i + 1, 6).setValue(data.email);      // Kolom F
        
        return ContentService.createTextOutput(JSON.stringify({status: "success", message: "Data berhasil disimpan!"}))
          .setMimeType(ContentService.MimeType.JSON);
      }
    }
    
    return ContentService.createTextOutput(JSON.stringify({status: "error", message: "ID tidak ditemukan"}))
      .setMimeType(ContentService.MimeType.JSON);
      
  } catch(error) {
    return ContentService.createTextOutput(JSON.stringify({status: "error", message: error.toString()}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
