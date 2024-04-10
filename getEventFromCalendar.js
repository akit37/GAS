function getCalendar() {
  var mySheet = SpreadsheetApp.getActiveSheet();
  var no = 1;
  
  var myCal = CalendarApp.getCalendarById('yours@gmail.com');
  var date = '2019/03/01 00:00:00';　// 日付の初期値
  var startDate = new Date(date);
  var endDate = new Date(date);
  endDate.setMonth(endDate.getMonth()+1);
  
  var myEvents = myCal.getEvents(startDate, endDate);
  
  // イベントを一つずつ表示する
  var maxRow = 1;
  for each(var evt in myEvents) {
  /*
    Logger.log(
      maxRow + "|" + evt.getTitle() + "|" +
      evt.getStartTime() + "|" + evt.getEndTime()
    );
    maxRow++;
  */
    mySheet.appendRow(
      [
        no,
        evt.getTitle(),
        evt.getStartTime(),
        evt.getEndTime(),
        "=INDIRECT(\"RC[-1]\",FALSE) - INDIRECT(\"RC[-2]\",FALSE)"
      ]
    );
    no++;
  }
}
