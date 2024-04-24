function onOpen() {
  var ui = SpreadsheetApp.getUi();
  // Or DocumentApp, SlidesApp or FormApp.
  ui.createMenu('メール可能確認')
      .addItem('メール可能確認', 'checkStatus')
      .addToUi();
}

function checkStatus() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var lastRow = sheet.getLastRow();
  var mailClass1Range = sheet.getRange("H3");
  var mailClass2Range = sheet.getRange("H4:H5");
  for (var i = 0; i <= lastRow-3; i+=1) {
    var statusDropDown = sheet.getRange(i+3,1);
    if(sheet.getRange(i+3,4).getValue()=="" && 
    sheet.getRange(i+3,5).getValue()=="B" && 
    sheet.getRange(i+3,1).getValue()!="送信済み" && 
    sheet.getRange(i+3,1).getValue()!="送信"){
      statusDropDown.setValue("可能");
      var rule = SpreadsheetApp.newDataValidation().requireValueInRange(mailClass2Range, true).build();
      statusDropDown.setDataValidation(rule);
    } else {
      statusDropDown.setValue("確認必要");
      var rule = SpreadsheetApp.newDataValidation().requireValueInRange(mailClass1Range, true).setAllowInvalid(false).build();
      statusDropDown.setDataValidation(rule);
    }
  }
}

function onEdit(e){
  var sheet = SpreadsheetApp.getActiveSheet();
  var edittedRange = e.range;
  var edittedValue = edittedRange.getValue();
  var mailClass3Range = sheet.getRange("H6");
  if (edittedRange.getColumn() == 1 && edittedValue == "送信"){
    edittedRange.setValue("送信済み");
    var rule = SpreadsheetApp.newDataValidation().requireValueInRange(mailClass3Range, true).setAllowInvalid(false).build();
    edittedRange.setDataValidation(rule);
  }
}
