function getFirstEmptyRowByColumnArray(ssID, sheetName, letter) {
    var spr = SpreadsheetApp.openById(ssID).getSheetByName(sheetName)
    var column = spr.getRange(letter + ':' + letter)
    var values = column.getValues(); // get all data in one call
    var ct = 0;
    while (values[ct] && values[ct][0] != "") {
      ct++;
    }
    return (ct + 1)
  }
  
  function getNewAnswersAndCheck() {
    var answers = SpreadsheetApp.openById("1CB4MGdVNitZZ-v89luCu4JkDsJ4q8CkXGNcTfwjYXP8").getSheetByName("Ответы на форму").getRange("A2:EG" + getFirstEmptyRowByColumnArray("1CB4MGdVNitZZ-v89luCu4JkDsJ4q8CkXGNcTfwjYXP8", "Ответы на форму", "A")).getValues()
    var recruitmentDataInSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Реестр счетов по рекрутменту").getRange("A2:V" + getFirstEmptyRowByColumnArray("1RaKEkcdozQmz-GYAw4nW9otYYqfgF3DhrpGJHv_MNw4", "Реестр счетов по рекрутменту", "A")).getValues()
    //var massDataInSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Реестр счетов по рекрутменту").getRange("A2:V" + getFirstEmptyRowByColumnArray("1RaKEkcdozQmz-GYAw4nW9otYYqfgF3DhrpGJHv_MNw4", "Реестр счетов по рекрутменту", "A"))
  
    for (var i = 0; i < answers.length; i++) {
      if (answers[i][2] == "Выставление счета") {
        var recruitmentInvoices = []
        var invoices = []
        var massInvoices = []
        if (answers[i][3] == "IT Recruitment") {
          recruitmentInvoices.push([answers[i][0], answers[i][1], answers[i][3], answers[i][4], answers[i][5], answers[i][6], answers[i][7], answers[i][8], answers[i][9], answers[i][10], answers[i][11], answers[i][12], answers[i][13], answers[i][14], answers[i][15], answers[i][16], answers[i][17], answers[i][18], answers[i][19], "Сформирован"])
        } else if (answers[i][3] == "Non-IT Recruitment") {
          recruitmentInvoices.push([answers[i][0], answers[i][1], answers[i][3], answers[i][20], answers[i][21], answers[i][22], answers[i][23], answers[i][24], answers[i][25], answers[i][26], answers[i][27], answers[i][28], answers[i][29], answers[i][30], answers[i][31], answers[i][32], answers[i][33], answers[i][34], answers[i][35], "Сформирован"])
        } /*else if (answers[i][3] == "Mass") {
          massInvoices.push([answers[i][0], answers[i][1], answers[i][3], answers[i][36], answers[i][37], answers[i][38], answers[i][39], answers[i][40], answers[i][41], answers[i][42], answers[i][43], answers[i][44], answers[i][45], answers[i][46], answers[i][47], "Сформирован"])
        }*/
      }
    }
  
    var recruitmentNew = []
    for (var i = 0; i < recruitmentInvoices.length; i++) {
      var recruitmentCnt = 0
      for (var j = 0; j < recruitmentDataInSheet.length; j++)
        if (recruitmentInvoices[i][0].valueOf() == recruitmentDataInSheet[j][0].valueOf() && recruitmentInvoices[i][1] == recruitmentDataInSheet[j][1] && recruitmentInvoices[i][2] == recruitmentDataInSheet[j][2] && recruitmentInvoices[i][3] == recruitmentDataInSheet[j][3] && recruitmentInvoices[i][4] == recruitmentDataInSheet[j][4] && recruitmentInvoices[i][5] == recruitmentDataInSheet[j][5] && recruitmentInvoices[i][6].valueOf() == recruitmentDataInSheet[j][6].valueOf() && recruitmentInvoices[i][7].valueOf() == recruitmentDataInSheet[j][7].valueOf() && recruitmentInvoices[i][8] == recruitmentDataInSheet[j][8] && recruitmentInvoices[i][9] == recruitmentDataInSheet[j][9] && recruitmentInvoices[i][10] == recruitmentDataInSheet[j][10] && recruitmentInvoices[i][11].valueOf() == recruitmentDataInSheet[j][11].valueOf() && recruitmentInvoices[i][12] == recruitmentDataInSheet[j][12] && recruitmentInvoices[i][13].toString() == recruitmentDataInSheet[j][13].toString() && recruitmentInvoices[i][14].toString() == recruitmentDataInSheet[j][14].toString() && recruitmentInvoices[i][15].toString() == recruitmentDataInSheet[j][15].toString() && recruitmentInvoices[i][16] == recruitmentDataInSheet[j][16] && recruitmentInvoices[i][17] == recruitmentDataInSheet[j][17] && recruitmentInvoices[i][18] == recruitmentDataInSheet[j][18]) recruitmentCnt++
      if (recruitmentCnt == 0) recruitmentNew.push(recruitmentInvoices[i])
    }
  
    if (recruitmentNew.length > 0) {
      SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Реестр счетов по рекрутменту").getRange("A" + getFirstEmptyRowByColumnArray("1RaKEkcdozQmz-GYAw4nW9otYYqfgF3DhrpGJHv_MNw4", "Реестр счетов по рекрутменту", "A") + ":T" + (getFirstEmptyRowByColumnArray("1RaKEkcdozQmz-GYAw4nW9otYYqfgF3DhrpGJHv_MNw4", "Реестр счетов по рекрутменту", "A") + recruitmentNew.length - 1)).setValues(recruitmentNew)
      GmailApp.sendEmail("elena@4ssistance.com", "Новый счёт", "В реестре счетов по рекрутменту появился новый счет: https://docs.google.com/spreadsheets/d/1RaKEkcdozQmz-GYAw4nW9otYYqfgF3DhrpGJHv_MNw4/edit#gid=1465472462")
    }
  
    for (var j = 0; j < recruitmentDataInSheet.length; j++) {
      if (recruitmentDataInSheet[j][19] == "Отклонен" && recruitmentDataInSheet[j][21] != "Письмо отправлено") {
        var img = DriveApp.getFileById("1zEQq1E3wGtATRRenrzfuASspk0cQK2DK").getBlob()
        MailApp.sendEmail({
          to: recruitmentDataInSheet[j][1],
          subject: "Проёб",
          htmlBody: "<p>" + recruitmentDataInSheet[j][20] + "</p><br><img src=\"cid:sampleImage\">",
          inlineImages: { sampleImage: img }
        })
        SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Реестр счетов по рекрутменту").getRange("V" + (j + 2)).setValue("Письмо отправлено")
      }
    }
  }