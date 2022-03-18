function onEdit(e) {
  if (SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getName() === 'Invoices tracker' && e.range.getColumn() == 5) {
    var cell = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange(e.range.getRow(), 7)
    var rule = SpreadsheetApp.newDataValidation().requireValueInRange(SpreadsheetApp.getActiveSpreadsheet().getSheetByName('LE list').getRange("B:B")).build()
    //var rule = SpreadsheetApp.newDataValidation().requireValueInList(SpreadsheetApp.getActiveSpreadsheet().getSheetByName('LE list').getRange("B:B").getValues(), true).build()
    cell.clearContent().clearDataValidations().setDataValidation(rule)
    cell.setValue("Loading...")
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName('LE list').getRange("A1").setValue(e.range.getValue())
    //Utilities.sleep(1000)
    //cell.clearContent()
    //SpreadsheetApp.getActiveSpreadsheet().getSheetByName('LE list').getRange("A1").setValues(findCellByvalue(e.range.getValue()).getValues())
    //SpreadsheetApp.getActiveSpreadsheet().getSheetByName('LE list').getRange("B1").setFormula("=QUERY('Projects List'!B2:D, " + '"SELECT D WHERE B = ' + "'" + '"' + '&A1&' + '"' + "'" + '"' + ')')
  }
  if (SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getName() === 'Invoices tracker' && e.range.getColumn() == 7) {
    var cell = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange(e.range.getRow(), 7)
    cell.clearDataValidations()
  }
}

/*function tryOut() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Projects List')
  var lastRow = sheet.getLastRow()
  console.log(sheet.getRange(2, 2, lastRow - 1, 1).getValues())
  var finder = sheet.getRange(2, 2, lastRow - 1, 1).createTextFinder("Старики закупки").matchEntireCell(true)
  ranges = finder.findAll()
  Logger.log('ranges.length = ' + ranges.length)
  for (var i = 0; i < ranges.length; i++) {
    ranges[i] = ranges[i].getA1Notation().toString().replace("B", "D")
    var value = console.log(SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Projects List').getRange(ranges[i]).getValue())
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName('LE list').getRange("B" + (i+1)).setValue(value)
  }
}

function findCellByvalue(value) {
  var sheet = SpreadsheetApp.getActive().getSheetByName('Inventory')
  var lastRow = sheet.getLastRow()
  var searchRange = sheet.getRange(3, 1, lastRow - 1, 1)

  // returns a range (of the cell)
  var finder = searchRange.createTextFinder(value).matchEntireCell(true)
  ranges = finder.findAll()
  Logger.log('ranges.length = ' + ranges.length)
  for (var i = 0; i < ranges.length; i++) {
    Logger.log('A1Notation = ' + ranges[i].getA1Notation())
  }
}

function getFirstEmptyRowWholeRow() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var range = sheet.getDataRange();
  var values = range.getValues();
  var row = 0;
  for (var row = 0; row < values.length; row++) {
    if (!values[row].join("")) break;
  }
  return (row + 1);
}

function getFirstEmptyRowByColumnArray() {
  var spr = SpreadsheetApp.getActiveSpreadsheet();
  var column = spr.getRange('U:U');
  var values = column.getValues(); // get all data in one call
  var ct = 0;
  while (values[ct] && values[ct][0] != "") {
    ct++;
  }
  return (ct + 1);
}

function setupTrigger() {
  ScriptApp.newTrigger('onChange')
    .forSpreadsheet(SpreadsheetApp.getActive())
    .onChange()
    .create()
}

function onChange(e) {
  if (e.changeType == "INSERT_ROW") {
    var emptyRow = getFirstEmptyRowByColumnArray()
    var rules = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Invoices tracker").getRange("A" + emptyRow + ":U" + emptyRow).getDataValidations()
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Invoices tracker").getRange("A" + emptyRow + ":U" + emptyRow).clearDataValidations()
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Invoices tracker").getRange("D" + emptyRow).setFormulaR1C1("=VLOOKUP(R[0]C[1],'Projects List'!C[-2]:C[1],4,FALSE)")
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Invoices tracker").getRange("F" + emptyRow).setFormulaR1C1("=VLOOKUP(R[0]C[-1],'Projects List'!C[-4]:C[-3],2,FALSE)")
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Invoices tracker").getRange("O" + emptyRow).setFormulaR1C1("=R[0]C[-6]+R[0]C[-1]")
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Invoices tracker").getRange("P" + emptyRow).setFormulaR1C1('=IF(R[0]C[1]<>"","YES",IF(AND((R[0]C[-1]-TODAY())>0,(R[0]C[-1]-TODAY())<=R[0]C[-2]*30%),"PRE-DELAY",IF(AND((R[0]C[-1]-TODAY())>R[0]C[-2]*30%,(R[0]C[-1]-TODAY())<=R[0]C[-2]),"NO",IF(AND((R[0]C[-1]-TODAY())<-1,(R[0]C[-1]-TODAY())>-15),"DELAY","DELAY > 15"))))')
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Invoices tracker").getRange("S" + emptyRow).setFormulaR1C1("=IF(R[0]C[-1], R[0]C[-7]-R[0]C[-1], 0)")
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Invoices tracker").getRange("T" + emptyRow).setFormulaR1C1("=VLOOKUP(R[0]C[-15],'Projects List'!R2C2:C9,7,FALSE)")
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Invoices tracker").getRange("U" + emptyRow).setFormulaR1C1("=VLOOKUP(R[0]C[-16],'Projects List'!R2C2:C9,8,FALSE)")
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Invoices tracker").getRange("A" + emptyRow + ":U" + emptyRow).setDataValidations(rules)
  }
}
*/