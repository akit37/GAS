var COL_A = 1;
var COL_B = 2;
var COL_C = 3;
var COL_D = 4;
var COL_H = 8;
var COL_I = 9;
var COL_J = 10;
var COL_K = 11;
var COL_L = 12;
var ROW_2 = 2;

var ui = SpreadsheetApp.getUi();

function summeryMembersWorktime() {
  var response = ui.alert('Confirm', '集計を実行しますか?', ui.ButtonSet.YES_NO);
  if (response != ui.Button.YES) {
    return;
  }

  var theSS = SpreadsheetApp.getActiveSpreadsheet();
  
  var resultSheetName = "集計_案件順";
  var SecondResultSheetName = "集計_メンバー順";
  var theConfigSheet = theSS.getSheetByName("MASTER_config");
  var startDatetime = theConfigSheet.getRange("B1").getValue();　// 集計期間開始日時
  var endDatetime = theConfigSheet.getRange("B2").getValue(); // 集計期間終了日時
  var theResultSheet = theSS.getSheetByName(resultSheetName);

  // 行がロックされていると消せないので表示状態をクリア　(その他の戻す処理は必須ではない)
  theResultSheet.unhideColumn(theResultSheet.getRange("A1:B1"));
  theResultSheet.unhideColumn(theResultSheet.getRange("D1:H1"));
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

  // 全員のカレンダーからイベントを取得して集計シートにまとめる
  var ret = getEventFromMembersCalendar(theSS, startDatetime, endDatetime, theResultSheet);
  
  // ガントチャートを作る。日にちごと、30分ごとの枠を作り、式をコピー
  setGanttChartLabel(theSS, startDatetime, endDatetime, theResultSheet);
  
  // 条件付き書式設定
  setConditionFormat(theSS, theResultSheet);
  
  // フィルタ再設定
  var filterrange = theResultSheet.getRange(ROW_2,1,theResultSheet.getLastRow()-1,theResultSheet.getLastColumn());
  filterrange.createFilter();
  
  // 表示状態を戻す
  theResultSheet.setFrozenRows(ROW_2); // 2行目で固定
  theResultSheet.setFrozenColumns(COL_L-1);　// L列で固定
  theResultSheet.hideColumns(COL_A, 2); // A,B列を非表示
  theResultSheet.hideColumns(COL_D, 5); // D,E,F,G,H列を非表示
  
  // ソート
  var pattern = 1;
  sortSummeryData(theResultSheet);
  
  // 複製して別パターンでソートした表を作る
  var anotherSheet = theSS.getSheetByName(SecondResultSheetName);
  if(anotherSheet) {
    theSS.deleteSheet(anotherSheet);
  }
  theSS.setActiveSheet(theResultSheet);
  var the2ndResutSheet = theSS.duplicateActiveSheet();
  the2ndResutSheet.setName(SecondResultSheetName);
  var pattern = 2;
  sortSummeryData(the2ndResutSheet, pattern);
  
  ui.alert('集計が終わりました。');
}

/* 指定月の特定カレンダーからイベントすべてを取得してスプレッドシートに書き出す */
function getEventFromMembersCalendar(theSS, startdate, enddate, theResultSheet) {
  var theMemberSheet = theSS.getSheetByName("MASTER_member");
  var range = theMemberSheet.getDataRange();
  var values = range.getValues();
  
　　　　for(var i=1; i < values.length; i++) { // i=0は、header
    var aNumber = values[i][0];
    var aEmail = values[i][1]; // メールアドレス
    var aGroup = values[i][2]; // 所属
    var aName = values[i][3]; // 名前
    
    var aSheet = theSS.getSheetByName(aName);
    aSheet.clearContents();
    
    var aCal=CalendarApp.getCalendarById(aEmail); //特定IDのカレンダーを取得
    if(aCal == null) { // 取得し損ねたらリトライ
      Utilities.sleep(100);
      aCal=CalendarApp.getCalendarById(aEmail);  // Retry once
      if(aCal == null) {
        ui.alert('カレンダーの取得に失敗しました。もう一度試してください。[User: ' + aEmail + ' ]');
        return(false);
      }
    }

    getCalendarContents(aNumber, aSheet, aCal, startdate, enddate);
    
    copyToSummerySheet(aSheet, theResultSheet);
  }
}

function getCalendarContents(number, theSheet, theCal, startdate, enddate) {
  var RANGE = 1;  // スプレッドシート：開始位置
  var FORMAT_TIME = "mm/dd hh:mm";  // スプレッドシート
  var FORMAT_DAY = "m/d";

  var all_schedules = theCal.getEvents(startdate, enddate);  //予定オブジェクトの生成
  
  // 終日の予定を取り除く
  var schedules = [];
  var j = 0;
  for(var i=0; i < all_schedules.length; i++) {
    if(all_schedules[i].isAllDayEvent()) {
      continue;
    }
    schedules[j] = all_schedules[i];
    j++;
  }

  // 予定を繰り返し出力する
  for(var i=0; i < schedules.length; i++) {
    var range = RANGE + i;
    // IDを出力
    theSheet.getRange(range, 1).setValue(number);
    // カレンダー名を出力
    theSheet.getRange(range, 2).setValue(theCal.getName());
    // 予定名を出力
    theSheet.getRange(range, 3).setValue(schedules[i].getTitle());
    // 開始日
    theSheet.getRange(range, 4).setValue(Utilities.formatDate(schedules[i].getStartTime(), "JST", "M/d"));
    // 開始時間を出力
    theSheet.getRange(range, 5).setValue(schedules[i].getStartTime()).setNumberFormat(FORMAT_TIME);
    // 終了時間を出力
    theSheet.getRange(range, 6).setValue(schedules[i].getEndTime()).setNumberFormat(FORMAT_TIME);
    // 稼働時間を出力
    //theSheet.getRange(range, 7).setValue("=INDIRECT(\"RC[-1]\",FALSE)-INDIRECT(\"RC[-2]\",FALSE)");
    // イベント内容を出力
    //theSheet.getRange(range, 8).setValue(schedules[i].getDescription()).setNumberFormat(FORMAT_TIME);
  }  
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
  // COLA:No.
  // COL_C:案件名
  // COL_D:開始時間
  // COL_H:案件タイトル
  // COL_I:業務
  // COL_J:部署
  // COL_K:名前
  if(pattern == 1) {
      sortRange.sort([COL_H, COL_D, COL_J, COL_K, COL_A]);
  }
  else {
      sortRange.sort([COL_J, COL_K, COL_D, COL_H, COL_A]);
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
    for(var i=0; i<24*2; i++) { // 24時間、30分きざみ
      aRange = theResultSheet.getRange(1,last_column);
      if(i==0) {
    　　　　  aRange.setNumberFormat('m/d');
        aRange.setFontSize(10);
        aRange.setFontWeight('bold');
      }
      else {
        aRange.setNumberFormat('h:mm');
        aRange.setFontSize(9);
        aRange.setFontWeight('normal');
      }
      aRange.setValue(nowdate);
      nowdate.setMinutes(nowdate.getMinutes()+30);
      last_column++;
    }
  }
  
  var formularange = theResultSheet.getRange("L2");
  formularange.setFormula('=ArrayFormula(if((L$1>=$E2:$E)*(L$1<$F2:$F),$I2:$I,""))');
  
  var destination = theResultSheet.getRange(ROW_2,start_column,1,last_column-start_column);
  formularange.autoFill(destination, SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
}

function setConditionFormat(theSS, theResultSheet) {
  // 条件付き書式でガントチャートを色付け(設定済みの書式を更新する)
  
  var conditionalFormatRules = theResultSheet.getConditionalFormatRules();
  
  conditionalFormatRules.splice(0, conditionalFormatRules.length, SpreadsheetApp.newConditionalFormatRule()
    .setRanges([theResultSheet.getRange(ROW_2,COL_L,theResultSheet.getLastRow()-1,theResultSheet.getLastColumn()-(COL_L-1))])
    .whenFormulaSatisfied('=and(L2<>"", $J2="大阪")')
    .setBackground('#F4C7C3')
    .build());
  theResultSheet.setConditionalFormatRules(conditionalFormatRules);
  
  conditionalFormatRules.push(SpreadsheetApp.newConditionalFormatRule()
    .setRanges([theResultSheet.getRange(ROW_2,COL_L,theResultSheet.getLastRow()-1,theResultSheet.getLastColumn()-(COL_L-1))])
    .whenFormulaSatisfied('=and(L2<>"", $J2="沖縄")')
    .setBackground('#B7E1CD')
    .build());
  theResultSheet.setConditionalFormatRules(conditionalFormatRules);
}
