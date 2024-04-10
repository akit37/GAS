function makeTemplate() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert('Confirm', 'テンプレートを作り直しますか?', ui.ButtonSet.YES_NO);
  if (response != ui.Button.YES) {
    return;
  }
  
  const ONE_PAGE_COL = 4; // 1ページのラベル列数
  const ONE_PAGE_ROW = 11; // 1ページのラベル行数
  
  try {
    thePresentation = SlidesApp.openById('your_slide_id');
  }
  catch(e) {
    Browser.msgBox(e);
    return;
  }
  
  var theSlides = thePresentation.getSlides();
  if(theSlides.length < 1) {
    ui.alert('テンプレートマスターSlideがありません');
    return;
  }
  var theLabelMaster = theSlides[0];
  var theSlideTemplate = theLabelMaster.duplicate();
  var theGroups = theSlideTemplate.getGroups();
  if(theGroups.length != 1) {
    ui.alert('テンプレートマスターGroupがありません');
    return;
  }
  var aGroup = theGroups[0];
  var width = aGroup.getWidth();
  var height = aGroup.getHeight();
  var left = aGroup.getLeft();
  var top = aGroup.getTop();
  Logger.log(width + "," + height + "," + left + ","+ top);
  
  for(var row=0; row < ONE_PAGE_ROW; row++) {
    for(var col=0; col < ONE_PAGE_COL; col++) {
      if(row==0 && col==0) { continue; }
      var thisBox = aGroup.duplicate();
      thisBox.setLeft(left + width * col);
      thisBox.setTop(top + height * row);
    }
  }
}

function insertToLabel() {
  const ONE_PAGE_MAX = 44; // 1ページのラベル枚数
  const OVER_MAX_LABEL = 10000; // 接待ない大きな件数
  var theSheet;
  var thePresentation;
  
  try {
    theSheet = SpreadsheetApp.getActiveSheet();
    thePresentation = SlidesApp.openById('your_slide_id');
  }
  catch(e) {
    Browser.msgBox(e);
    return;
  }
  
  var theSlides = thePresentation.getSlides();
  if(theSlides.length < 2) {
    ui.alert('テンプレートSlideがありません');
    return;
  }
  var theSlideTemplate = theSlides[1];
  
  var rows = theSheet.getLastRow();
  var dataArray = theSheet.getRange(2,1,rows-1,8).getDisplayValues();
  var datalen = dataArray.length; // ラベルデータの件数
  if(datalen < 1) {
    return;
  }
  var startPosition = dataArray[0][7]; // ラベル配置開始位置
  
  var row = 0; // Spreadsheet上の行
  var nowCount = OVER_MAX_LABEL;
  
  while(row < datalen) {
    var theLabelSlide = theSlideTemplate.duplicate();
    var theGroups = theLabelSlide.getGroups();
    for(var i=0; i<theGroups.length; i++) {
      var thisNumber = "";
      var thisTitle = "";
      var thisColor = "";
      var thisPrice = "";
      var thisJancode = "";
      var thisBarcode = "";
      var thisCount = "";

      startPosition--;
    
      if(startPosition < 1 && row < datalen) {
        thisNumber = dataArray[row][0];
        thisTitle = dataArray[row][1];
        thisColor = dataArray[row][2];
        thisPrice = dataArray[row][3] + "円";
        thisJancode = dataArray[row][4];
        //thisBarcode = dataArray[row][5];
        thisCount = dataArray[row][6];
    
        if(nowCount > thisCount) {
          nowCount = thisCount;
        }
        nowCount--;
        if(nowCount == 0) {
          nowCount = OVER_MAX_LABEL;
          row++;
        }
      }
    
      var elements = theGroups[i].getChildren();
      elements.forEach(function(element) {
        var aTitle = trimString(element.getTitle());
        //Logger.log("i="+i + " title=" + aTitle);

        switch(aTitle) {
        case "NUM":
          element.asShape().getText().setText(thisNumber);
          break;
        case "TITLE":
          element.asShape().getText().setText(thisTitle);
          break;
        case "COLOR":
          element.asShape().getText().setText(thisColor);
          break;
        case "PRICE":
          element.asShape().getText().setText(thisPrice);
          break;
        case "JANCODE":
          element.asShape().getText().setText(thisJancode);
          break;
        case "BARCODE":
          element.asShape().getText().setText(thisJancode);
          if(thisJancode != "") {
            element.asShape().getText().getTextStyle().setFontFamily("Libre Barcode 39");
            element.asShape().getText().getTextStyle().setFontSize(18);
            //element.asShape().getText().getTextStyle().setFontFamily("Libre Barcode 128");
            //element.asShape().getText().getTextStyle().setFontSize(28);
          }
          break;
        }
      });
    };
  }
}

function trimString(src){
  if (src == null || src == undefined){
    return "";
  }
  return src.replace(/(^\s+)|(\s+$)/g, "");
}
