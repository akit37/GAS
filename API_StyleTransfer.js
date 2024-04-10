function styleTransfer() {
  var theSS;
  var theSheet;
  
  try {
    theSS = SpreadsheetApp.getActiveSpreadsheet();
    theSheet = theSS.getActiveSheet();
  }
  catch(e) {
    Browser.msgBox(e);
    return;
  }

  var sourceFolderName = theSheet.getRange("B1").getValue();
  var convertedFolderName = theSheet.getRange("B2").getValue();
  var style = theSheet.getRange("B3").getValue();
  Logger.log(style);

  var sourceFolder;
  var convertedFolder;

  var parentFolder = DriveApp.getFileById(theSS.getId()).getParents();
  var currentFolder = parentFolder.next();
  var folders = currentFolder.getFolders();
  while (folders.hasNext()) {
    var folder = folders.next();
    switch(folder.getName()) {
      case sourceFolderName:
        sourceFolder = folder;
        break;
      case convertedFolderName:
        convertedFolder = folder;
        break;
    }
  }

  var files = sourceFolder.getFiles();
  var mimetype = "image/jpeg";

  while (files.hasNext()) {
    var file = files.next();
    if(file.getMimeType() != mimetype) {
      continue;
    }
    var outfilename = style + "-" + file.getName();

    var b64image = Utilities.base64Encode(file.getBlob().getBytes());
  
    jsonbody = post2metabase(b64image, style);

    var result = JSON.parse(jsonbody);
    if(result.result) {
      var b64image = result.result_data;
      var decoded = Utilities.base64Decode(b64image);
      var blob_result = Utilities.newBlob(decoded, mimetype, outfilename);
      convertedFolder.createFile(blob_result).setName(outfilename);
      Logger.log(outfilename);
    }
    else {
      Logger.log(result.result_data);
    }
  }
}

function post2metabase(b64image, style) { 
  var url = 'https://your_api_server';
  var headers = {
    'Content-Type' : 'application/json',
    'Authorization': 'token YOURTOKEN'
  };

  var post_data = {
    'b64image': b64image,
    'style':style
  };

  var options = {
    "method" : "post",
    "headers" : headers,
    "payload" : JSON.stringify(post_data)
  };

  return UrlFetchApp.fetch(url, options); 
}
