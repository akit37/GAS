function myFunction() {
  var url = 'http://www.pref.tochigi.lg.jp/e04/welfare/hoken-eisei/kansen/hp/documents/kanjahasseiichiran.pdf';
  var ssname = "9_tochigi";
  var theSS = SpreadsheetApp.create(ssname);
  
  var theSheet = theSS.getActiveSheet();
  var theRange = theSheet.getActiveRange();
  
  var response = UrlFetchApp.fetch(url);
  var blob = response.getBlob();
  var type = blob.getContentType();
  var name = blob.getName();
  //Logger.log(type + ", " + name);
  
  var resource = {
    title: name,
    mimeType: type
    // mimeType: 'application/pdf'
    // mimeType: 'image/png'
    // mimeType: 'image/jpeg'
  };

  // OCR
  var optionalArgs = {
    ocr: true,
    ocrLanguage: 'ja'
    // ocrLanguage: 'en'
  };

  var file = Drive.Files.insert(resource, blob, optionalArgs);

  var doc = DocumentApp.openById(file.id);

  var textArray = [];
  doc.getBody().getParagraphs().forEach( function(value, i) {
    var rowArray = [];
    rowArray.push(value.getText());
    //theSheet.appendRow(rowArray);
    textArray.push(rowArray);
  } )

  var rows = textArray.length;
  var cols = textArray[0].length;
  theRange = theSheet.getRange(1,1,rows,cols);
  theRange.setValues(textArray);
  theRange.splitTextToColumns(SpreadsheetApp.TextToColumnsDelimiter.SPACE);
}
