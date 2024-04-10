function insertImageToCell() {
  var theSS = SpreadsheetApp.getActiveSpreadsheet();
  var theSheet = theSS.getActiveSheet();
 
  // テスト画像という名前で png 形式の Blob オブジェクトを作成
  var theImage = DriveApp.getFileById("1bM_qMnfpuP11lX-sNnIfE3mNO-MXtH4r").getBlob();
  //var aBlob = Utilities.newBlob(theImage, 'image/jpg', 'TestImage');
  // そのシートの (1,1) セルにその画像を挿入
  theSheet.insertImage(theImage, 1, 1);
}
