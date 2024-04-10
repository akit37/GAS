function printHelloWorld() {
  var message = "Hello World!2";
  var ui = SpreadsheetApp.getUi();
  ui.alert(message);
  
  Logger.log(message);
}

function nabeatu() {
  var ui = SpreadsheetApp.getUi();
  var num = 6;
  var result = num % 3;
  if(result == 0) {
    ui.alert(num + ":AHO");
  }
  else {
    ui.alert(num);
  }
  
  for(var i=0; i<10; i++) {
    ui.alert(i);
  }
  
}
