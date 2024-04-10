var COL_A = 1;
var COL_B = 2;
var COL_C = 3;
var COL_D = 4;
var COL_E = 5;
var COL_F = 6;
var COL_G = 7;
var COL_H = 8;
var COL_I = 9;
var COL_J = 10;
var COL_K = 11;
var COL_L = 12;
var COL_M = 13;
var COL_N = 14;
var COL_O = 15;
var COL_P = 16;
var ROW_2 = 2;

var ui = SpreadsheetApp.getUi();

function drawGanttchartByMembersTask() {
  var response = ui.alert('Confirm', '集計を実行しますか?', ui.ButtonSet.YES_NO);
  if (response != ui.Button.YES) {
    return;
  }

  var theSS = SpreadsheetApp.getActiveSpreadsheet();
  
  var resultSheetName = "GanttChartByProject";
  var secondResultSheetName = "GanttChartByMember";
  var theConfigSheet = theSS.getSheetByName("MASTER_config");
  var startDatetime = theConfigSheet.getRange("B1").getValue();　// 集計期間開始日時
  var endDatetime = theConfigSheet.getRange("B2").getValue(); // 集計期間終了日時
  endDatetime.setDate(endDatetime.getDate()+1);
  var theResultSheet = theSS.getSheetByName(resultSheetName);

  // 行がロックされていると消せないので表示状態をクリア　(その他の戻す処理は必須ではない)
  theResultSheet.unhideColumn(theResultSheet.getRange("B1:D1"));
  theResultSheet.setFrozenRows(0);
  theResultSheet.setFrozenColumns(0);
  
  // フィルタ削除
  var aFilter = theResultSheet.getFilter();
  if(aFilter) {
    aFilter.remove();
  }
  
  // データをクリア
  if(theResultSheet.getLastRow() > ROW_2) {　// 2行目以降を消す
    theResultSheet.deleteRows(ROW_2+1, theResultSheet.getLastRow()-ROW_2);
  }
  if(theResultSheet.getLastColumn() > COL_L) { // L列より右を消す
    theResultSheet.deleteColumns(COL_L+1, theResultSheet.getLastColumn()-COL_L);
  }

  // 全員のタスクを取得して集計シートにまとめる
  var ret = collectMembersWorktime(theSS, startDatetime, endDatetime, theResultSheet);
  
  // ガントチャートを作る。日にちごと、30分ごとの枠を作り、式をコピー
  setGanttChartLabel(theSS, startDatetime, endDatetime, theResultSheet);
  
  // 条件付き書式設定
  setConditionFormat(theSS, theResultSheet);
  
  // フィルタ再設定
  var filterrange = theResultSheet.getRange(ROW_2,1,theResultSheet.getLastRow()-1,theResultSheet.getLastColumn());
  filterrange.createFilter();
  
  // 表示状態を戻す
  theResultSheet.setFrozenRows(ROW_2); // 2行目で固定
  theResultSheet.setFrozenColumns(COL_F);　// F列で固定
  theResultSheet.hideColumns(COL_B, 3); // B-C列を非表示
  
  // ソート
  var pattern = 1;
  sortSummeryData(theResultSheet);
  
  // 複製して別パターンでソートした表を作る
  var anotherSheet = theSS.getSheetByName(secondResultSheetName);
  if(anotherSheet) {
    theSS.deleteSheet(anotherSheet);
  }
  theSS.setActiveSheet(theResultSheet);
  var the2ndResutSheet = theSS.duplicateActiveSheet();
  the2ndResutSheet.setName(secondResultSheetName);
  var pattern = 2;
  sortSummeryData(the2ndResutSheet, pattern);
  ui.alert('集計が終わりました。');
}

function collectMembersWorktime(theSS, startdate, enddate, theResultSheet) {
  var theProjectrSheet = theSS.getSheetByName("MASTER_project");
  var range = theProjectrSheet.getDataRange();
  var values = range.getValues();
  
  var tempSheet = theSS.getSheetByName("temp_for_gantt");
  tempSheet.clearContents();
　
  for(var i=1; i < values.length; i++) { // i=0は、header
    var aStatus = values[i][0];
    var aNumber = values[i][1];
    var aSubNumber = values[i][2];
    var aClient = values[i][3];
    var aProject = values[i][4];
    var aTask = values[i][5];
    var aStart = values[i][6];
    var aEnd = values[i][7];
    var aWorkDays = values[i][8];
    var aInCharge = values[i][9];
    var aPowerRatio = values[i][10];
    var aManDays = values[i][11];

    var aEndPlusOne = aEnd;
    aEndPlusOne.setDate(aEndPlusOne.getDate()+1);
    
    if(startdate.getTime() <= aStart.getTime() || aEndPlusOne.getTime() <= enddate.getTime()) {
      tempSheet.getRange(i, 1).setValue(aStatus);
      tempSheet.getRange(i, 2).setValue(aNumber);
      tempSheet.getRange(i, 3).setValue(aSubNumber);
      tempSheet.getRange(i, 4).setValue(aClient);
      tempSheet.getRange(i, 5).setValue(aProject);
      tempSheet.getRange(i, 6).setValue(aTask);
      tempSheet.getRange(i, 7).setValue(aStart);
      tempSheet.getRange(i, 8).setValue(aEnd);
      tempSheet.getRange(i, 9).setValue(aWorkDays);
      tempSheet.getRange(i, 10).setValue(aInCharge);
      tempSheet.getRange(i, 11).setValue(aPowerRatio);
    }
  }
  copyToSummerySheet(tempSheet, theResultSheet);
}

function copyToSummerySheet(FromSheet, ToSheet) {
  if(FromSheet.getLastRow() == 0) {
    return;
  }
  var fromRange = FromSheet.getDataRange();
  var rangewidth = fromRange.getWidth();
  var rangeheight = fromRange.getHeight();
  
  var bottomOfToSheet = ToSheet.getLastRow() + 1;
  var toRange = ToSheet.getRange(bottomOfToSheet, 1, rangeheight, rangewidth);
  toRange.setValues(fromRange.getValues());
}

function sortSummeryData(resultSheet, pattern) {
  if(pattern === undefined) {
      pattern = 1;
  }
  var sortRange = resultSheet.getRange(ROW_2+1,1,resultSheet.getLastRow()-ROW_2, resultSheet.getLastColumn());
  // 並び替え
  // COL_A:Status
  // COL_B:No.
  // COL_C:Sub-No.
  // COL_D:Client
  // COL_E:ProjectName
  // COL_F:Task
  // COL_G:Start
  // COL_H:End
  // COL_I:WorkDays
  // COL_J:InCharge
  // COL_K:PowerRatio
  if(pattern == 1) {
      sortRange.sort([COL_B, COL_C, COL_G, COL_H, COL_J]);
  }
  else {
      sortRange.sort([COL_J, COL_B, COL_C, COL_G, COL_H]);
  }
}

function setGanttChartLabel(theSS, startdate, enddate, theResultSheet) {
  var daterange = theResultSheet.getRange("L1");
  var nowdate = new Date(startdate);
  daterange.setValue(nowdate);
  
  var lastdate = new Date(enddate);
  
  var start_column = COL_L; // L列
  var last_column = start_column;
  while(nowdate < lastdate) {
    aRange = theResultSheet.getRange(1,last_column);
    aRange.setNumberFormat('m/d');
    aRange.setFontSize(10);
    aRange.setFontWeight('normal');
    aRange.setValue(nowdate);
    // 週ごとのセル
    nowdate.setDate(nowdate.getDate()+7);

    last_column++;
  }
  
  var formularange = theResultSheet.getRange("L2");

  //曜日を無視してチャートを描く
  formularange.setFormula('=ArrayFormula(if(NOT((L$1>=$G2:$G)*(L$1<=$H2:$H)),"",if(($F2:$F="-"),$E2:$E,$K2:$K)))');
  //土日はガントチャートを引かない場合はこれを使う(週ごと表示の場合に土日スタートを選ぶとチャートが描けない)
  //formularange.setFormula('=ArrayFormula(if(NOT((L$1>=$G2:$G)*(L$1<=$H2:$H)*(WEEKDAY(L$1,2)<6)),"",if(($F2:$F="-"),$E2:$E,$K2:$K)))');
  
  var destination = theResultSheet.getRange(ROW_2,start_column,1,last_column-start_column);
  formularange.autoFill(destination, SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
}

function setConditionFormat(theSS, theResultSheet) {
  // 条件付き書式でガントチャートを色付け(設定済みの書式を更新する)
  var theMemberSheet = theSS.getSheetByName("MASTER_member");
  var range = theMemberSheet.getDataRange();
  var values = range.getValues();
  
  var conditionalFormatRules = theResultSheet.getConditionalFormatRules();
  
  // Project全体の帯
  /*
  conditionalFormatRules.splice(0, conditionalFormatRules.length, SpreadsheetApp.newConditionalFormatRule()
    .setRanges([theResultSheet.getRange(ROW_2,COL_L,theResultSheet.getLastRow()-1,theResultSheet.getLastColumn()-(COL_L-1))])
    .whenFormulaSatisfied('=and(L2<>"", $F2="-")')
    .setBackground('#808080')
    .build());
  theResultSheet.setConditionalFormatRules(conditionalFormatRules);
  */

  for(var i=1; i < values.length; i++) { // i=0は、header
    var aNumber = values[i][0];
    var aName = values[i][1];
    var aColor = values[i][2];
    var condition = '=and(L2<>"", $F2<>"-", $J2="' + aName + '")';
    var condition_member = '=($J2="' + aName + '")';
    
    // memberカラムに色付け
    conditionalFormatRules.push(SpreadsheetApp.newConditionalFormatRule()
      .setRanges([theResultSheet.getRange(ROW_2,COL_J,theResultSheet.getLastRow()-1,1)])
      .whenFormulaSatisfied(condition_member)
      .setBackground(aColor)
      .build());
    theResultSheet.setConditionalFormatRules(conditionalFormatRules);

    // memberごとのバーに色付け
    conditionalFormatRules.push(SpreadsheetApp.newConditionalFormatRule()
      .setRanges([theResultSheet.getRange(ROW_2,COL_L,theResultSheet.getLastRow()-1,theResultSheet.getLastColumn()-(COL_L-1))])
      .whenFormulaSatisfied(condition)
      .setBackground(aColor)
      .build());
    theResultSheet.setConditionalFormatRules(conditionalFormatRules);
  }
}
