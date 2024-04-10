var FOLDER_ID = ''; //保存するフォルダ
var SEARCH_TERM = 'from:(xxx@gmail.com) smaller:299K';

function getMail(){
  var myFolder = DriveApp.getFolderById(FOLDER_ID); //フォルダを取得
  var myThreads = GmailApp.search(SEARCH_TERM, 442, 500); //条件にマッチしたスレッドを検索して取得 / 最大500 
  var myMessages = GmailApp.getMessagesForThreads(myThreads); //スレッドからメールを取得し二次元配列で格納

  Logger.log(myMessages.length);

  for(var i in myMessages){
    for(var j in myMessages[i]){

      var attachments = myMessages[i][j].getAttachments(); //添付ファイルを取得
      var from = myMessages[i][j].getFrom(); //添付ファイルを取得
      for(var k in attachments){
          attachments[k].setName("vf_"+attachments[k].getName()); // 添付ファイルの保存名変更
          //Logger.log(attachments[k]);
          myFolder.createFile(attachments[k]); //ドライブに添付ファイルを保存
      }
      //myMessages[i][j].markRead(); //既読（処理が終わったら既読にする）
    }
  }
}

function fileCount() {
  var folder = DriveApp.getFolderById(FOLDER_ID);
  var contents = folder.getFiles();
  var i = 0;
  while(contents.hasNext()) {
    file = contents.next();
    i++
  }
  Logger.log("このフォルダ内のファイル数は"+i+"です。");
}
