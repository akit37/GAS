function convertEntrySheets() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert('Confirm', '配下のフォルダ内のExcelをすべて、SpreadSheetとPDFに変換します。実行しますか?', ui.ButtonSet.YES_NO);
  if (response != ui.Button.YES) {
    return;
  }
  // MimeType.GOOGLE_SHEETS
  // MimeType.MICROSOFT_EXCEL
  // MimeType.PDF
  const token = ScriptApp.getOAuthToken();

  // get Root Folder
  var theSS = SpreadsheetApp.getActiveSpreadsheet();
  var theSheet = theSS.getActiveSheet();
  const rootFolderId = theSheet.getRange("B1").getValue();
  const outputSSFolderId = theSheet.getRange("B2").getValue();
  const outputPDFFolderId = theSheet.getRange("B3").getValue();

  const rootFolder = DriveApp.getFolderById(rootFolderId);

  // get Folder List
  const teamFolders = rootFolder.getFolders();
  const teamFoldersArray=[];
  while(teamFolders.hasNext()) {
    var a_folder = teamFolders.next();
    a_folderID = a_folder.getId();
    if(a_folderID != outputSSFolderId && a_folderID != outputPDFFolderId) {
      teamFoldersArray.push([a_folderID, a_folder.getName()]);
    }
  }

  // Excel to SpreadSheet変換
  for(var i=0; i < teamFoldersArray.length; i++) {
    teamFolderId = teamFoldersArray[i][0];
    teamFolderName = teamFoldersArray[i][1];

    var teamFolder = DriveApp.getFolderById(teamFolderId);
    var excelFiles = teamFolder.getFilesByType(MimeType.MICROSOFT_EXCEL);

    while (excelFiles.hasNext()) {
      var excelFile = excelFiles.next();
      var fileid = excelFile.getId();

      var option =  {
        mimeType:MimeType.GOOGLE_SHEETS,  //Google sheets
        parents:[{id:outputSSFolderId}],  //出力先フォルダーを指定
        title:teamFolderName              //出力ファイル名
      }

      // 既存なら差し替え
      var existFiles = DriveApp.getFolderById(outputSSFolderId).getFilesByName(teamFolderName);
      if(existFiles.hasNext()) {
        var existFile = existFiles.next();
        if(excelFile.getLastUpdated() > existFile.getLastUpdated()) {
          Drive.Files.update(option, existFile.getId(), excelFile);
          Logger.log("update:" + teamFolderName);
          Utilities.sleep(6000);
        }
      }
      else{
        Drive.Files.insert(option, excelFile);
        Logger.log("insert:" + teamFolderName);
        Utilities.sleep(6000);
      }
    }
  }

  // SpreadSheet to PDF変換
  const ssFolder = DriveApp.getFolderById(outputSSFolderId);
  const pdfFilder = DriveApp.getFolderById(outputPDFFolderId);
  var sslFiles = ssFolder.getFilesByType(MimeType.GOOGLE_SHEETS);

  while (sslFiles.hasNext()) {
    var ssFile = sslFiles.next();
    var fileid = ssFile.getId();
    var filename = ssFile.getName();

    var ss = SpreadsheetApp.openById(fileid);
    var sheets = ss.getSheets();
    var sheet = ss.getSheets()[0];
    var sheetid = sheet.getSheetId();

    var pdfFilename = filename + ".pdf"

    var option =  {
      mimeType:MimeType.GOOGLE_PDF,
      parents:[{id:outputPDFFolderId}],  //出力先フォルダーを指定
      title:pdfFilename              //出力ファイル名
    }

    // 既存なら差し替え
    var existFiles = DriveApp.getFolderById(outputPDFFolderId).getFilesByName(pdfFilename);
    if(existFiles.hasNext()) {
      var existFile = existFiles.next();
      if(ssFile.getLastUpdated() > existFile.getLastUpdated()) {
        Drive.Files.update(option, existFile.getId(), ssFile);
        Logger.log("update:" + pdfFilename);
        Utilities.sleep(6000);
      }
    }
    else{
      Drive.Files.insert(option, ssFile);
      Logger.log("insert:" + pdfFilename);
      Utilities.sleep(6000); //429エラー回避のためsleep
    }
  }
}
