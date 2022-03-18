function addRow() {
    var lastRow = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Статистика расходов").getLastRow()
    console.log(lastRow)
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Статистика расходов").insertRowAfter(lastRow)
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Статистика расходов").getRange("E" + (lastRow+1)).setValue(Utilities.formatDate(new Date(), "GMT+3", "dd/MM/yyyy"))
    var formulas = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Статистика расходов").getRange("F" + lastRow + ":G" + lastRow).getFormulasR1C1()
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Статистика расходов").getRange("F" + (lastRow+1) + ":G" + (lastRow+1)).setFormulasR1C1(formulas)
  }
  
  function setupTrigger() {
    ScriptApp.newTrigger('addRow')
      .timeBased()
      .atHour(0)
      .everyDays(1)
      .create()
  }