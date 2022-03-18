function addRow() {
    var lastRow = getFirstEmptyRowByColumnArray()
    console.log(lastRow)
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Депозит").insertRowAfter(lastRow)
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Депозит").getRange("E" + (lastRow + 1)).setValue(Utilities.formatDate(new Date(), "GMT+3", "dd/MM/yyyy"))
    var formulas = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Депозит").getRange("F" + lastRow + ":G" + lastRow).getFormulasR1C1()
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Депозит").getRange("F" + (lastRow + 1) + ":G" + (lastRow + 1)).setFormulasR1C1(formulas)
  }
  
  function getFirstEmptyRowByColumnArray() {
    var spr = SpreadsheetApp.getActiveSpreadsheet()
    var column = spr.getRange('E:E')
    var values = column.getValues()
    var ct = 0
    while (values[ct] && values[ct][0] != "") {
      ct++
    }
    return ct
  }
  
  function setupTrigger() {
    ScriptApp.newTrigger('addRow')
      .timeBased()
      .atHour(0)
      .everyDays(1)
      .create()
  }