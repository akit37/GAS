var SheetName = "getDriveUser";
activeSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SheetName)

function getDriveUser() {
  firstStep();
  adminListAllTeamDrives();
}

function firstStep(sheetName) {
  activeSheet.clear();

  activeSheet.getRange(1, 1).setValue("ドライブ名")
  activeSheet.getRange(1, 1).setBackground("#7169e5");
  activeSheet.getRange(1, 1).setFontColor("#ffffff");
}  

function adminListAllTeamDrives(){
  var pageTokenDrive;
  var pageTokenMember;
  var teamDrives;
  var permissions;

  //ドライブ名の一覧を取得
  do{
    teamDrives = Drive.Drives.list({pageToken:pageTokenDrive,maxResults:100,useDomainAdminAccess:true})
    if(teamDrives.items && teamDrives.items.length > 0){

      for (var i = 0; i < teamDrives.items.length; i++) {

        var teamDrive = teamDrives.items[i];
        //ドライブ名の一覧情報を転記
        activeSheet.getRange(i+2, 1).setValue(teamDrive.name)

        skipflg = false;
        switch(teamDrive.id) {
          case "0AMem6pWTHuxgUk9PVA":
          case "0APTmeHHf4OXAUk9PVA":
          case "0AHY8rqJSUI6gUk9PVA":
          case "0ALeWRWWb79IpUk9PVA":
          case "0AIN6MAq8TsbpUk9PVA":
          case "0AIvajs7b_jrvUk9PVA":
          case "0ANE66Gff16MVUk9PVA":
          case "0AM5uhFQebSiUUk9PVA":
          case "0AKluYmu7yikaUk9PVA":
          case "0AEzLzzemgz2jUk9PVA":
          case "0AKBZ56b7PRrsUk9PVA":
          case "0AOrHos3ABu5yUk9PVA":
          case "0AOlBcKDRynpUUk9PVA":
          case "0AJcx0bSI_YYKUk9PVA":
          case "0ALkQBB1McwbcUk9PVA":
          case "0ADI4rvrxg8SzUk9PVA":
          case "0AHYJ55C9FUmQUk9PVA":
          case "0AG2K9YsNE6-oUk9PVA":
          case "0AARMA6tdhTu-Uk9PVA":
          case "0AOsno6027a9vUk9PVA":
          case "0AOTQCrfvowovUk9PVA":
          case "0AJyF0uDZj3srUk9PVA":
          case "0ANIxrgwvNye8Uk9PVA":
          case "0APOn88fVHyxMUk9PVA":
          case "0AJElt7CzLSAkUk9PVA":
          case "0AD7vNXHJ0wobUk9PVA":
          case "0AAHnxH1o7BcMUk9PVA":
          case "0ALTY50yHK4lHUk9PVA":
          case "0AGeQAr_cMFtMUk9PVA":
            skipflg = true;
            break;
        }

        if(!skipflg) {
          //ドライブごとのメンバーの権限を取得
          do{
            permissions = Drive.Permissions.list(teamDrive.id, {maxResults:40,pageToken:pageTokenMember,supportsAllDrives:true}) ;
            if(permissions.items && permissions.items.length > 0){
              for (var j = 0,k = 2; j < permissions.items.length; j++,k=k+2) {

              activeSheet.getRange(1, k).setValue("メンバー")
              activeSheet.getRange(1, k+1).setValue("権限")
              activeSheet.getRange(1, k, 1,k).setBackground("#7169e5");
              activeSheet.getRange(1, k, 1,k).setFontColor("#ffffff");


              //権限情報を取得して変数に格納
              var permission = permissions.items[j];
              activeSheet.getRange(i+2, k).setValue(permission.emailAddress)

              switch(permission.role){
              case "organizer":
                activeSheet.getRange(i+2, k+1).setValue("管理者")
                break;
              case "fileOrganizer":
                activeSheet.getRange(i+2, k+1).setValue("コンテンツ管理者")
                break;
              case "writer":
                activeSheet.getRange(i+2, k+1).setValue("投稿者")
                break;
              case "commenter":
                activeSheet.getRange(i+2, k+1).setValue("閲覧者(コメント可)")
                break;
              case "reader":
                activeSheet.getRange(i+2, k+1).setValue("閲覧者")
                break;
              }
            }
            }else{
              Logger.log("メンバー/権限が見つかりませんでした。");
            }

          //次のページのpageTokenを取得する
          pageTokenMember = permissions.nextPageTokens
          }while(pageTokenMember)
        }
    }

    }else{
      Logger.log("共有ドライブが見つかりませんでした。");
    }

    //次のページのpageTokenを取得する
    pageTokenDrive = teamDrives.nextPageToken
    }while(pageTokenDrive)
}
