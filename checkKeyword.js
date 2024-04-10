function checkKeyword() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert('Confirm', '抽出を実行しますか?', ui.ButtonSet.YES_NO);
  if (response != ui.Button.YES) {
    return;
  }
 
  var isTESTMODE = false;
  var theSS;
  var theSheetArray = [];
  var theSheet;
  
  try {
    theSS = SpreadsheetApp.openById('your_sheet_id'); // PDA日別個人情報案件チェック
    theSheetArray = theSS.getSheets();
//    theSheet = theSS.getActiveSheet();
  }
  catch(e) {
    Browser.msgBox(e);
    return;
  }

  var keywordArray = [];
  for(var s=0; s < theSheetArray.length; s++) {1
    theSheet = theSheetArray[s];
    if(theSheet.getName() == "keyword") {
      var rows = theSheet.getLastRow();
      keywordArray = theSheet.getRange(2,1,rows-1,1).getValues();
      break;
    }
  }

  var sqlwhere = "select A,B,C,D,E,F,G,H,I,J,K,O where (";
  for(var i=0; i< keywordArray.length; i++) {
    if(keywordArray[i] != "") {
  　　　　  if(i!=0) sqlwhere += " or ";
      sqlwhere += "I LIKE '%" + keywordArray[i] + "%'";
    }
  }
  sqlwhere += ") and (O='該当なし')";

// Logger.log(sqlwhere);
  var outputCell = theSheet.getRange("B2");
  outputCell.setValue(sqlwhere);
}

function copyTodaysList() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert('Confirm', '本日分の一覧へのコピーを実行しますか?', ui.ButtonSet.YES_NO);
  if (response != ui.Button.YES) {
    return;
  }

  var sheetFrom1 = "「該当あり」リスト";
  var sheetFrom2 = "「該当なし」要確認候補リスト";
  var sheetTo1 = "2017「該当あり」一覧";
  var sheetTo2 = "2017「該当なし」確認一覧";

  var theSS;
  var targetSheetFrom1;
  var targetSheetFrom2;
  var targetSheetTo1;
  var targetSheetTo2;
  
  try {
    theSS = SpreadsheetApp.getActiveSpreadsheet();
    targetSheetFrom1 = theSS.getSheetByName(sheetFrom1);
    targetSheetFrom2 = theSS.getSheetByName(sheetFrom2);
    targetSheetTo1 = theSS.getSheetByName(sheetTo1);
    targetSheetTo2 = theSS.getSheetByName(sheetTo2);
  }
  catch(e) {
    Browser.msgBox(e);
    return;
  }

  var targetRange1 = targetSheetFrom1.getRange(2,1, targetSheetFrom1.getLastRow()-1, targetSheetFrom1.getLastColumn());
  var targetRows1 = targetSheetTo1.getLastRow() + 1;
  targetRange1.copyTo(targetSheetTo1.getRange(targetRows1,1), {contentsOnly:true});
  targetRange1.copyTo(targetSheetTo1.getRange(targetRows1,1), {formatOnly:true});

  var targetRange2 = targetSheetFrom2.getRange(2,1, targetSheetFrom2.getLastRow()-1, targetSheetFrom2.getLastColumn());
  var targetRow2 = targetSheetTo2.getLastRow()+1;
  targetRange2.copyTo(targetSheetTo2.getRange(targetRow2,1), {contentsOnly:true});
  targetRange2.copyTo(targetSheetTo2.getRange(targetRow2,1), {formatOnly:true});
}
