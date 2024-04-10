// カンガルーマジック配達状況
function generateQuery() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert('Confirm', 'クエリを更新して抽出を実行しますか?', ui.ButtonSet.YES_NO);
  if (response != ui.Button.YES) {
    return;
  }
 
  var isTESTMODE = false;
  var theSS;
  var thetheSheetArray = [];
  var theSheet;
  
  try {
    theSS = SpreadsheetApp.openById('your_sheet_id'); // カンガルーマジック配達状況
    theSheetArray = theSS.getSheets();
  }
  catch(e) {
    Browser.msgBox(e);
    return;
  }

  var theSheet = theSS.getSheetByName("keyword");
  var keywordArrayNotLike = theSheet.getRange(2,2,theSheet.getLastRow()-1,1).getValues();
  var keywordArrayLike = theSheet.getRange(2,3,theSheet.getLastRow()-1,1).getValues();
  var keywordArrayNotLike2 = theSheet.getRange(2,4,theSheet.getLastRow()-1,1).getValues();

  var sqlwhere = "select * where (NOT (";
  for(var i=0; i< keywordArrayNotLike.length; i++) {
    if(keywordArrayNotLike[i] != "") {
  　　　　  if(i!=0) sqlwhere += " or "
      sqlwhere += "E LIKE '%" + keywordArrayNotLike[i] + "%'";
    }
  }
  sqlwhere += ")) and ( ";

  for(var i=0; i< keywordArrayLike.length; i++) {
    if(keywordArrayLike[i] != "") {
  　　　　  if(i!=0) sqlwhere += " or "
      sqlwhere += "G LIKE '%" + keywordArrayLike[i] + "%'";
      sqlwhere += " or H LIKE '%" + keywordArrayLike[i] + "%'";
    }
  }
  sqlwhere += ") and ( ";

  sqlwhere += "NOT (";
  for(var i=0; i< keywordArrayNotLike2.length; i++) {
    if(keywordArrayNotLike2[i] != "") {
  　　　　  if(i!=0) sqlwhere += " or "
      sqlwhere += "AA LIKE '%" + keywordArrayNotLike2[i] + "%'";
      sqlwhere += " or AB LIKE '%" + keywordArrayNotLike2[i] + "%'";
      sqlwhere += " or AC LIKE '%" + keywordArrayNotLike2[i] + "%'";
      sqlwhere += " or AD LIKE '%" + keywordArrayNotLike2[i] + "%'";
      sqlwhere += " or AE LIKE '%" + keywordArrayNotLike2[i] + "%'";
    }
  }

  sqlwhere += "))";
  
// Logger.log(sqlwhere);
  var outputCell = theSheet.getRange("E2");
  outputCell.setValue(sqlwhere);
}

function autoFiltering() {

/*
=arrayformula(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(iferror(REGEXEXTRACT(G2:G,"[０-９][０-９]月[０-９][０-９]日"), iferror(REGEXEXTRACT(H2:H,"[０-９][０-９]月[０-９][０-９]日"), "")),"０","0"), "１","1"),"２","2"),"３","3"),"４","4"),"５","5"),"６","6"),"７","7"),"８","8"),"９","9"))
*/

  var ui = SpreadsheetApp.getUi();
  var response = ui.alert('Confirm', '自動フィルタリングを実行しますか?', ui.ButtonSet.YES_NO);
  if (response != ui.Button.YES) {
    return;
  }
 
　 　var theSS;
  var theSourceSheet;
  var theDestSheet;
  var sourceSheetName = "queryTemp";
  var destSheetName = "カンガルーマジック配達状況_完成";
  
  try {
    theSS = SpreadsheetApp.getActiveSpreadsheet();
    theSourceSheet = theSS.getSheetByName(sourceSheetName);
    theDestSheet = theSS.getSheetByName(destSheetName);
    theDestSheet.clear();
  }
  catch(e) {
    Browser.msgBox(e);
    return;
  }

  var theSheet = theSS.getSheetByName("keyword");
  var todayValue = theSheet.getRange("A2").getValue();

  // 最終行(空白行)を得る
  var cells = theSourceSheet.getRange("A:A").getValues();
  var maxrow = 0;
  for(i=0; i<cells.length; i++) {
    if(cells[i] == "") {
      maxrow = i;
      break;
    }  
  }
  
  var targetCells = theSourceSheet.getRange("AG:AG").getValues(); // AG: dateValue
  var AtoE = 5;
  var GtoAE = 25;
  theSourceSheet.getRange(1,1,1,AtoE).copyTo(theDestSheet.getRange(1,1,1,AtoE), {contentsOnly:true}); //title行 A-E
  theSourceSheet.getRange(1,7,1,GtoAE).copyTo(theDestSheet.getRange(1,6,1,GtoAE), {contentsOnly:true}); //title行 G-AE
  var row = 2;
  for(i=1; i　<　maxrow; i++) {
    if(targetCells[i] - todayValue <= 0) {
      theSourceSheet.getRange(i+1,1,1,AtoE).copyTo(theDestSheet.getRange(row,1,1,AtoE), {contentsOnly:true}); // A-E
      theSourceSheet.getRange(i+1,7,1,GtoAE).copyTo(theDestSheet.getRange(row,6,1,GtoAE), {contentsOnly:true}); // G-AE
      row++;
    }
  }
}

function copyTodaysList() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert('Confirm', '本日分のファイルをコピーしますか?', ui.ButtonSet.YES_NO);
  if (response != ui.Button.YES) {
    return;
  }

　 　var theSS;
  var sheetFrom1 = "カンガルーマジック配達状況";
  var targetSheetFrom1;
  
  var sheetTemp = "copyTemp";
  var targetSheetTemp;

  var today = Utilities.formatDate(new Date(), "GMT+9", "yyyyMMdd");
//  var newSS = SpreadsheetApp.create("過去ログ_カンガルーマジック配達状況/配達状況_"+today);
  var newSS = SpreadsheetApp.create("配達状況_"+today);
  
  try {
    theSS = SpreadsheetApp.getActiveSpreadsheet();
    targetSheetFrom1 = theSS.getSheetByName(sheetFrom1);
    targetSheetTemp = theSS.getSheetByName(sheetTemp);
    targetSheetTemp.clear();
  }
  catch(e) {
    Browser.msgBox(e);
    return;
  }

  // 最終行(空白行)を得る
  var cells = targetSheetFrom1.getRange("A:A").getValues();
  var maxrow = 0;
  for(i=0; i<cells.length; i++) {
    if(cells[i] == "") {
      maxrow = i;
      break;
    }  
  }
 
//  ui.alert(maxrow);
  var targetRange1 = targetSheetFrom1.getRange(1,1, maxrow, targetSheetFrom1.getLastColumn());
  targetRange1.copyTo(targetSheetTemp.getRange(1,1), {contentsOnly:true});
  targetSheetTemp.copyTo(newSS);
  newSS.deleteSheet(newSS.getSheetByName("シート1"));
  
  // バックアップフォルダへ移動
  var backupFolder = DriveApp.getFolderById('your_folder_id'); // 過去ログ_カンガルーマジック配達状況
  var newFile = DriveApp.getFileById(newSS.getId());
  backupFolder.addFile(newFile);
  DriveApp.getRootFolder().removeFile(newFile);
}

