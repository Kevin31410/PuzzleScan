function doGet(e) {
  // 1. SI ON DEMANDE UNE RECHERCHE DE CODE-BARRES (Pont API pour éviter le CORS)
  if (e.parameter.barcode) {
    var url = "https://api.upcitemdb.com/prod/trial/lookup?upc=" + e.parameter.barcode;
    var response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    return ContentService.createTextOutput(response.getContentText()).setMimeType(ContentService.MimeType.JSON);
  } 
  
  // 2. SINON, ON RENVOIE TOUTE LA LUDOTHÈQUE
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var data = sheet.getDataRange().getValues();
  var puzzles = [];
  
  for (var i = 1; i < data.length; i++) {
    if(data[i][0] !== "") {
      puzzles.push({
        id: data[i][0].toString(), barcode: data[i][1], brand: data[i][2], title: data[i][3],
        pieces: data[i][4], price: data[i][5], notes: data[i][6], imgUrl: data[i][7],
        emprunteur: data[i][8], datePret: data[i][9]
      });
    }
  }
  return ContentService.createTextOutput(JSON.stringify(puzzles)).setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  // 3. SAUVEGARDE DES DONNÉES ENVOYÉES PAR L'APP
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var puzzles = JSON.parse(e.postData.contents);
  
  sheet.clearContents();
  sheet.appendRow(['ID', 'Code-barres', 'Marque', 'Titre', 'Pièces', 'Prix', 'Notes', 'Image', 'Emprunteur', 'Date_Pret']);
  
  if (puzzles.length > 0) {
    var rows = puzzles.map(function(p) {
      return [p.id, p.barcode, p.brand, p.title, p.pieces, p.price, p.notes, p.imgUrl, p.emprunteur, p.datePret];
    });
    sheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
  }
  
  return ContentService.createTextOutput(JSON.stringify({"status": "ok"})).setMimeType(ContentService.MimeType.JSON);
}
