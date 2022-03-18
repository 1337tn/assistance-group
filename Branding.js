function createSheet(name) {
    /*
    Функция создает в активном файле новый лист с названием name и активирует его
    */
    console.log("Creating new sheet...")
    var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet()
    var yourNewSheet = activeSpreadsheet.getSheetByName(name)
  
    if (yourNewSheet != null) {
      activeSpreadsheet.deleteSheet(yourNewSheet)
    }
    SpreadsheetApp.flush()
    yourNewSheet = activeSpreadsheet.insertSheet()
    yourNewSheet.setName(name)
    SpreadsheetApp.flush()
    console.log("Finished sheet creation")
  }
  
  function getFirstEmptyRowByColumnArray(letter) {
    var spr = SpreadsheetApp.getActiveSpreadsheet()
    var column = spr.getRange(letter + ':' + letter)
    var values = column.getValues(); // get all data in one call
    var ct = 0;
    while (values[ct] && values[ct][0] != "") {
      ct++;
    }
    return (ct + 1)
  }
  
  function getFirstEmptyRowByColumnArraySheet(sheetName, letter) {
    var spr = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName)
    var column = spr.getRange(letter + ':' + letter)
    var values = column.getValues(); // get all data in one call
    var ct = 0;
    while (values[ct] && values[ct][0] != "") {
      ct++;
    }
    return (ct + 1)
  }
  
  function monthToRussianString(month) {
    switch (month) {
      case 'Jan':
        month = 'января';
        break;
      case 'Feb':
        month = 'февраля';
        break;
      case 'Mar':
        month = 'марта';
        break;
      case 'Apr':
        month = 'апреля';
        break;
      case 'May':
        month = 'мая';
        break;
      case 'Jun':
        month = 'июня';
        break;
      case 'Jul':
        month = 'июля';
        break;
      case 'Aug':
        month = 'августа';
        break;
      case 'Sep':
        month = 'сентября';
        break;
      case 'Oct':
        month = 'октября';
        break;
      case 'Nov':
        month = 'ноября';
        break;
      case 'Dec':
        month = 'декабря';
        break;
    }
    return month
  }
  
  function monthToRussianString2(month) {
    switch (month) {
      case 'Jan':
        month = 'январь';
        break;
      case 'Feb':
        month = 'февраль';
        break;
      case 'Mar':
        month = 'март';
        break;
      case 'Apr':
        month = 'апрель';
        break;
      case 'May':
        month = 'май';
        break;
      case 'Jun':
        month = 'июнь';
        break;
      case 'Jul':
        month = 'июль';
        break;
      case 'Aug':
        month = 'август';
        break;
      case 'Sep':
        month = 'сентябрь';
        break;
      case 'Oct':
        month = 'октябрь';
        break;
      case 'Nov':
        month = 'ноябрь';
        break;
      case 'Dec':
        month = 'декабрь';
        break;
    }
    return month
  }
  
  function monthToRussianString3(month) {
    switch (month) {
      case 0:
        month = 'Январь';
        break;
      case 1:
        month = 'Февраль';
        break;
      case 2:
        month = 'Март';
        break;
      case 3:
        month = 'Апрель';
        break;
      case 4:
        month = 'Май';
        break;
      case 5:
        month = 'Июнь';
        break;
      case 6:
        month = 'Июль';
        break;
      case 7:
        month = 'Август';
        break;
      case 8:
        month = 'Сентябрь';
        break;
      case 9:
        month = 'Октябрь';
        break;
      case 10:
        month = 'Ноябрь';
        break;
      case 11:
        month = 'Декабрь';
        break;
    }
    return month
  }
  
  function numberToString(_number) {
    var _arr_numbers = new Array();
    _arr_numbers[1] = new Array('', 'один', 'два', 'три', 'четыре', 'пять', 'шесть', 'семь', 'восемь', 'девять', 'десять', 'одиннадцать', 'двенадцать', 'тринадцать', 'четырнадцать', 'пятнадцать', 'шестнадцать', 'семнадцать', 'восемнадцать', 'девятнадцать');
    _arr_numbers[2] = new Array('', '', 'двадцать', 'тридцать', 'сорок', 'пятьдесят', 'шестьдесят', 'семьдесят', 'восемьдесят', 'девяносто');
    _arr_numbers[3] = new Array('', 'сто', 'двести', 'триста', 'четыреста', 'пятьсот', 'шестьсот', 'семьсот', 'восемьсот', 'девятьсот');
    function number_parser(_num, _desc) {
      var _string = '';
      var _num_hundred = '';
      if (_num.length == 3) {
        _num_hundred = _num.substr(0, 1);
        _num = _num.substr(1, 3);
        _string = _arr_numbers[3][_num_hundred] + ' ';
      }
      if (_num < 20) _string += _arr_numbers[1][parseFloat(_num)] + ' ';
      else {
        var _first_num = _num.substr(0, 1);
        var _second_num = _num.substr(1, 2);
        _string += _arr_numbers[2][_first_num] + ' ' + _arr_numbers[1][_second_num] + ' ';
      }
      switch (_desc) {
        case 0:
          var _last_num = parseFloat(_num.substr(-1));
          if (_last_num == 1) _string += 'рубль';
          else if (_last_num > 1 && _last_num < 5) _string += 'рубля';
          else _string += 'рублей';
          break;
        case 1:
          var _last_num = parseFloat(_num.substr(-1));
          if (_last_num == 1 && thousandsBugFixingFlag) _string += 'тысяча ';
          else if (_last_num > 1 && _last_num < 5 && thousandsBugFixingFlag) _string += 'тысячи ';
          else _string += 'тысяч ';
          _string = _string.replace('один ', 'одна ');
          _string = _string.replace('два ', 'две ');
          break;
        case 2:
          var _last_num = parseFloat(_num.substr(-1));
          if (_last_num == 1 && millionsBugFixingFlag) _string += 'миллион ';
          else if (_last_num > 1 && _last_num < 5 && millionsBugFixingFlag) _string += 'миллиона ';
          else _string += 'миллионов ';
          break;
        case 3:
          var _last_num = parseFloat(_num.substr(-1));
          if (_last_num == 1 && billionsFixingFlag) _string += 'миллиард ';
          else if (_last_num > 1 && _last_num < 5 && billionsBugFixingFlag) _string += 'миллиарда ';
          else _string += 'миллиардов ';
          break;
      }
      _string = _string.replace('  ', ' ');
      return _string;
    }
    function decimals_parser(_num) {
      var _first_num = _num.substr(0, 1);
      var _second_num = parseFloat(_num.substr(1, 2));
      var _string = ' ' + _first_num + _second_num;
      if (_second_num == 1) _string += ' копейка';
      else if (_second_num > 1 && _second_num < 5) _string += ' копейки';
      else _string += ' копеек';
      return _string;
    }
    if (!_number || _number == 0) return false;
    if (typeof _number !== 'number') {
      _number = _number.replace(',', '.');
      _number = parseFloat(_number);
      if (isNaN(_number)) return false;
    }
    _number = _number.toFixed(2);
    var thousandsBugFixingFlag = true
    var millionsBugFixingFlag = true
    var billionsFixingFlag = true
    if (_number.indexOf('.') != -1) {
      var _number_arr = _number.split('.');
      var _number = _number_arr[0];
      var _number_decimals = _number_arr[1];
      if (_number.length >= 5) if (_number[_number.length - 5] == 1 && (_number[_number.length - 4] >= 0 || _number[_number.length - 4] <= 9)) thousandsBugFixingFlag = false
      if (_number.length >= 8) if (_number[_number.length - 8] == 1 && (_number[_number.length - 7] >= 0 || _number[_number.length - 7] <= 9)) millionsBugFixingFlag = false
      if (_number.length >= 11) if (_number[_number.length - 11] == 1 && (_number[_number.length - 10] >= 0 || _number[_number.length - 10] <= 9)) billionsBugFixingFlag = false
    }
    var _number_length = _number.length;
    var _string = '';
    var _num_parser = '';
    var _count = 0;
    for (var _p = (_number_length - 1); _p >= 0; _p--) {
      var _num_digit = _number.substr(_p, 1);
      _num_parser = _num_digit + _num_parser;
      if ((_num_parser.length == 3 || _p == 0) && !isNaN(parseFloat(_num_parser))) {
        _string = number_parser(_num_parser, _count) + _string;
        _num_parser = '';
        _count++;
      }
    }
    if (_number_decimals) _string += decimals_parser(_number_decimals);
    //console.log(_string)
    return _string;
  }
  
  function PDF(dateFrom, action, contractor, city) {
    var pdfName = 'БО ' + contractor + ' ' + city + ' ' + monthToRussianString2(Utilities.formatDate(dateFrom, Session.getScriptTimeZone(), "MMM")) + ' 2022 (' + action + ')' //Need to set the values to another sheet
    console.log('Saving PDF ' + pdfName + ' started...')
    var folderID = "1HKNIW8YvCqm7FRKvufOZogO8uFy13PV3"; // Folder id to save in a Drive folder.
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Бумажный отчет').activate()
    var ss = SpreadsheetApp.openByUrl(
      'https://docs.google.com/spreadsheets/d/1H92njsZuOkB_z7Sxbx3DoGOcg8py-ky6C4abKO5kkfY/edit?ts=60d16e87').getActiveSheet()
    var sourceSpreadsheet = SpreadsheetApp.getActive()
    var url = 'https://docs.google.com/spreadsheets/d/' + sourceSpreadsheet.getId() + '/export?exportFormat=pdf&format=pdf' // export as pdf / csv / xls / xlsx
      + '&size=A4' // paper size legal / letter / A4
      + '&portrait=true' // orientation, false for landscape
      + '&fitw=true' // fit to page width, false for actual size
      + '&sheetnames=true&printtitle=false' // hide optional headers and footers
      + '&pagenum=RIGHT&gridlines=false' // hide page numbers and gridlines
      + '&fzr=false' // do not repeat row headers (frozen rows) on each page
      + '&horizontal_alignment=CENTER' //LEFT/CENTER/RIGHT
      + '&vertical_alignment=TOP' //TOP/MIDDLE/BOTTOM
      + '&gid=' + sourceSpreadsheet.getSheetId()
    var options = {
      method: "GET",
      headers: { Authorization: "Bearer " + ScriptApp.getOAuthToken() },
      muteHttpExceptions: true
    }
    var response = UrlFetchApp.fetch(url, options).getBlob();
    var existing = DriveApp.getFolderById(folderID).getFilesByName(pdfName)
    var hasFile = existing.hasNext()
    if (hasFile) {
      var duplicate = existing.next()
      var durl = 'https://www.googleapis.com/drive/v3/files/' + duplicate.getId();
      var dres = UrlFetchApp.fetch(durl, {
        method: 'delete',
        muteHttpExceptions: true,
        headers: { 'Authorization': 'Bearer ' + ScriptApp.getOAuthToken() }
      });
      if (dres.getResponseCode() >= 400) {
        //handle errors;
      }
    }
    var fileId = DriveApp.getFolderById(folderID).createFile(response).setName(pdfName).getId()
    return fileId
  }
  
  function getLastRowSpecial(range) {
    var rowNum = 0;
    var blank = false;
    for (var row = 0; row < range.length; row++) {
      if (range[row][0] == "" && !blank) {
        rowNum = row;
        blank = true;
      } else if (range[row][0] !== "") {
        blank = false;
      };
    };
    return rowNum;
  }
  
  function main(rowNumber) {
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Providers_list').activate()
    var providersList = SpreadsheetApp.openByUrl(SpreadsheetApp.getActiveSpreadsheet().getRange('C' + rowNumber).getValues()[0][0])
    SpreadsheetApp.setActiveSpreadsheet(providersList)
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName('2022').activate()
    data = SpreadsheetApp.getActiveSpreadsheet().getRange('A2:N' + (getFirstEmptyRowByColumnArray("A") - 1)).getValues()
    if (data[data.length - 1][0] <= dateFrom) return
    var files = []
    var actionsDict = {}
    SpreadsheetApp.setActiveSpreadsheet(SpreadsheetApp.openById("1wLKoe1JtQTDqwcZfrJUPZ-5pTYjQ15xHTJG3lcY_nfA"))
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName('БО управление скриптом').activate()
    var contractor = SpreadsheetApp.getActiveSpreadsheet().getRange('B' + rowNumber).getValues()[0][0]
    var city = SpreadsheetApp.getActiveSpreadsheet().getRange('A' + rowNumber).getValues()[0][0]
    var dateFrom = SpreadsheetApp.getActiveSpreadsheet().getRange('C' + rowNumber).getValues()[0][0]
    var dateTo = SpreadsheetApp.getActiveSpreadsheet().getRange('D' + rowNumber).getValues()[0][0]
    var action = SpreadsheetApp.getActiveSpreadsheet().getRange('E' + rowNumber).getValues()[0][0]
    var iteration = SpreadsheetApp.getActiveSpreadsheet().getRange('H2').getValues()[0][0]
    const allActions = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange('J2:J' + (getFirstEmptyRowByColumnArray("J") - 1)).getValues()
    for (var i = 0; i < allActions.length; i++) {
      actionsDict[allActions[i]] = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange('I2:I' + (getFirstEmptyRowByColumnArray("J") - 1)).getValues()[i]
    }
    SpreadsheetApp.getActiveSpreadsheet().getRange('G2').setValue(rowNumber)
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Providers_list').activate()
    var KPP = SpreadsheetApp.getActiveSpreadsheet().getRange('L' + rowNumber).getValues()[0][0]
    var INN = SpreadsheetApp.getActiveSpreadsheet().getRange('M' + rowNumber).getValues()[0][0]
    var abbr = SpreadsheetApp.getActiveSpreadsheet().getRange('Z' + rowNumber).getValues()[0][0]
    var email = SpreadsheetApp.getActiveSpreadsheet().getRange('F' + rowNumber).getValues()[0][0]
    var hasNDS = SpreadsheetApp.getActiveSpreadsheet().getRange('X' + rowNumber).getValues()[0][0]
    var fileId
    SpreadsheetApp.setActiveSpreadsheet(providersList)
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName('2022').activate()
    var test = SpreadsheetApp.getActiveSpreadsheet().getRange('D2:D' + (getFirstEmptyRowByColumnArray("A") - 1)).getValues()
    var actions = []
    if (action == 'Все акции') {
      for (var i = 0; i < test.length; i++)
        for (var j = 0; j < allActions.length; j++)
          if (allActions[j].toString() == test[i].toString() && actions.indexOf(test[i].toString()) === -1) actions.push(test[i].toString())
    } else actions.push(action)
    var brandingMasterdata = SpreadsheetApp.openById("1wLKoe1JtQTDqwcZfrJUPZ-5pTYjQ15xHTJG3lcY_nfA")
    SpreadsheetApp.setActiveSpreadsheet(brandingMasterdata)
    for (var action of actions) {
      // Предобработка листа БО
      createSheet('Бумажный отчет')
      var range = SpreadsheetApp.getActiveSpreadsheet().getRange("A1:L")
      range.setWrap(true)
      range.setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP)
      range.setVerticalAlignment("middle")
      SpreadsheetApp.getActiveSpreadsheet().getRange('A3:J3').merge()
      SpreadsheetApp.getActiveSpreadsheet().getRange('A4:J4').merge()
      if (hasNDS == 'YES') {
        SpreadsheetApp.getActiveSpreadsheet().getRange('A3: L3').merge()
        SpreadsheetApp.getActiveSpreadsheet().getRange('A4:L4').merge()
      }
      SpreadsheetApp.getActiveSpreadsheet().getRange('A6:J6').merge()
      SpreadsheetApp.getActiveSpreadsheet().getRange('A7:J7').merge()
      SpreadsheetApp.getActiveSpreadsheet().getRange('A8:J8').merge()
      SpreadsheetApp.getActiveSpreadsheet().getRange('A10:A11').merge()
      SpreadsheetApp.getActiveSpreadsheet().getRange('B10:I10').merge()
      SpreadsheetApp.getActiveSpreadsheet().getRange('J10:J11').merge()
      range = SpreadsheetApp.getActiveSpreadsheet().getRange("A1:J9")
      range.setFontSize(12);
      range = SpreadsheetApp.getActiveSpreadsheet().getRange("A1:L11")
      range.setFontWeight("bold");
      range = SpreadsheetApp.getActiveSpreadsheet().getRange("A3:J4")
      range.setHorizontalAlignment("center")
      range = SpreadsheetApp.getActiveSpreadsheet().getRange("A10:L11")
      range.setHorizontalAlignment("center")
      SpreadsheetApp.getActiveSpreadsheet().setRowHeight(11, 70)
      SpreadsheetApp.getActiveSpreadsheet().getRange('A4').setValue('за  период с «' + dateFrom.getDate() + '» ' + monthToRussianString(Utilities.formatDate(dateFrom, Session.getScriptTimeZone(), "MMM")) + ' 2022 г. по «' + dateTo.getDate() + '» ' + monthToRussianString(Utilities.formatDate(dateTo, Session.getScriptTimeZone(), "MMM")) + ' 2022 г.')
      /* TODO: Распарсить часть названия, чтобы была расшифровка ООО, пример: Общество с ограниченной ответственностью "АртПроект" */
      SpreadsheetApp.getActiveSpreadsheet().getRange('A6').setValue('Исполнитель: ' + contractor)
      if (KPP != "") SpreadsheetApp.getActiveSpreadsheet().getRange('A7').setValue('ИНН ' + INN + '/ КПП ' + KPP)
      else SpreadsheetApp.getActiveSpreadsheet().getRange('A7').setValue('ИНН ' + INN)
      SpreadsheetApp.getActiveSpreadsheet().getRange('A8').setValue('Заказчик: Общество с ограниченной ответственностью "Эссистэнс"')
      SpreadsheetApp.getActiveSpreadsheet().getRange('A9').setValue('ИНН 7704331668/ КПП 771401001')
      SpreadsheetApp.getActiveSpreadsheet().getRange('A10').setValue('Дата монтажа')
      SpreadsheetApp.getActiveSpreadsheet().getRange('B10').setValue('Цена за 1 единицу')
      SpreadsheetApp.getActiveSpreadsheet().getRange('J10').setValue('Итого Стоимость услуг по заказу без НДС')
      SpreadsheetApp.getActiveSpreadsheet().getRange('B11').setValue('Монтаж заднего стекла')
      SpreadsheetApp.getActiveSpreadsheet().getRange('C11').setValue('Стоимость наклеек на заднее стекло')
      SpreadsheetApp.getActiveSpreadsheet().getRange('D11').setValue('Монтаж наклеек на борта')
      SpreadsheetApp.getActiveSpreadsheet().getRange('E11').setValue('Стоимость наклеек на борта')
      SpreadsheetApp.getActiveSpreadsheet().getRange('F11').setValue('Монтаж коробов')
      SpreadsheetApp.getActiveSpreadsheet().getRange('G11').setValue('Монтаж шашечного пояса')
      SpreadsheetApp.getActiveSpreadsheet().getRange('H11').setValue('Стоимость демонтажа')
      SpreadsheetApp.getActiveSpreadsheet().getRange('I11').setValue('Бренд')
      SpreadsheetApp.getActiveSpreadsheet().getRange('A12:A').setNumberFormat("dd.MM.yyyy")
      if (hasNDS == "ДА") {
        SpreadsheetApp.getActiveSpreadsheet().getRange('A9:L9').merge()
        SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange('K10:K11').merge()
        SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange('L10:L11').merge()
        SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange('K10').setValue('НДС 20%')
        SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange('L10').setValue('Итого Стоимость услуг по заказу c НДС')
      }
      else SpreadsheetApp.getActiveSpreadsheet().getRange('A9:J9').merge()
      var cnt = 12
      var num = iteration + '/' + abbr + '/' + Utilities.formatDate(dateFrom, Session.getScriptTimeZone(), "MM") + '-' + Utilities.formatDate(dateFrom, Session.getScriptTimeZone(), "YYYY") + '/' + actionsDict[action]
      SpreadsheetApp.getActiveSpreadsheet().getRange('A3').setValue('Отчет о проделанной работе № ' + num)
      for (var i = 0; i < data.length; i++) {
        if (data[i][4].toString() == 'ЗАКРЫТ' || data[i][4].toString() == 'СПЕЦКВОТА ЗАКРЫТ') {
          if (data[i][0] >= dateFrom && data[i][0] <= dateTo && data[i][1] == contractor && data[i][3].toString().indexOf(action) != -1) {
            SpreadsheetApp.getActiveSpreadsheet().getRange('A' + cnt).setValue(data[i][0])
            SpreadsheetApp.getActiveSpreadsheet().getRange('B' + cnt).setValue(data[i][6])
            SpreadsheetApp.getActiveSpreadsheet().getRange('C' + cnt).setValue(data[i][7])
            SpreadsheetApp.getActiveSpreadsheet().getRange('D' + cnt).setValue(data[i][8])
            SpreadsheetApp.getActiveSpreadsheet().getRange('E' + cnt).setValue(data[i][9])
            SpreadsheetApp.getActiveSpreadsheet().getRange('F' + cnt).setValue(data[i][10])
            SpreadsheetApp.getActiveSpreadsheet().getRange('G' + cnt).setValue(data[i][11])
            SpreadsheetApp.getActiveSpreadsheet().getRange('H' + cnt).setValue(data[i][12])
            if (data[i][3].toString().indexOf('Uber') != -1) SpreadsheetApp.getActiveSpreadsheet().getRange('I' + cnt).setValue('Uber')
            else if (data[i][3].toString().indexOf('Yandex') != -1) SpreadsheetApp.getActiveSpreadsheet().getRange('I' + cnt).setValue('Yandex')
            else SpreadsheetApp.getActiveSpreadsheet().getRange('I' + cnt).setValue(data[i][3])
            SpreadsheetApp.getActiveSpreadsheet().getRange('J' + cnt).setValue(data[i][13])
            if (hasNDS == 'ДА') {
              SpreadsheetApp.getActiveSpreadsheet().getRange('K' + cnt).setValue(data[i][13] * 0.2)
              SpreadsheetApp.getActiveSpreadsheet().getRange('L' + cnt).setValue(data[i][13] * 1.2)
            }
            cnt += 1
          }
        }
      }
      if (hasNDS == 'ДА') range = SpreadsheetApp.getActiveSpreadsheet().getRange('A9:L' + (cnt - 1))
      else range = SpreadsheetApp.getActiveSpreadsheet().getRange('A9:J' + (cnt - 1))
      range.setBorder(true, true, true, true, true, true)
      if (hasNDS == 'ДА') {
        SpreadsheetApp.getActiveSpreadsheet().getRange('K' + cnt).setValue('Итоговая сумма:').setFontWeight("bold")
        var netSquare = SpreadsheetApp.getActiveSpreadsheet().getRange('L12:L' + (cnt - 1)).getValues().flat().filter(v => v != '').map(v => parseInt(v))
        if (netSquare.length != 0) {
          var sum = netSquare.reduce((a, b) => a + b)
        }
        SpreadsheetApp.getActiveSpreadsheet().getRange('A' + (cnt + 2) + ':L' + (cnt + 7)).setWrap(true)
        SpreadsheetApp.getActiveSpreadsheet().getRange('A' + (cnt + 2) + ':L' + (cnt + 7)).setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP)
        SpreadsheetApp.getActiveSpreadsheet().getRange('A' + (cnt + 2) + ':L' + (cnt + 7)).setFontSize(12)
        SpreadsheetApp.getActiveSpreadsheet().getRange('L12:L' + (cnt + 1)).setNumberFormat("# ##0.00").setHorizontalAlignment("center")
        SpreadsheetApp.getActiveSpreadsheet().getRange('L' + cnt).setValue(sum).setFontWeight("bold")
        SpreadsheetApp.getActiveSpreadsheet().getRange('A' + (cnt + 2) + ':L' + (cnt + 2)).merge()
        SpreadsheetApp.getActiveSpreadsheet().getRange('A' + (cnt + 3) + ':L' + (cnt + 3)).merge()
      }
      else {
        SpreadsheetApp.getActiveSpreadsheet().getRange('I' + cnt).setValue('Итоговая сумма:').setFontWeight("bold")
        var netSquare = SpreadsheetApp.getActiveSpreadsheet().getRange('J12:J' + (cnt - 1)).getValues().flat().filter(v => v != '').map(v => parseInt(v))
        if (netSquare.length != 0) {
          var sum = netSquare.reduce((a, b) => a + b)
        }
        SpreadsheetApp.getActiveSpreadsheet().getRange('A' + (cnt + 2) + ':J' + (cnt + 7)).setWrap(true)
        SpreadsheetApp.getActiveSpreadsheet().getRange('A' + (cnt + 2) + ':J' + (cnt + 7)).setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP)
        SpreadsheetApp.getActiveSpreadsheet().getRange('A' + (cnt + 2) + ':J' + (cnt + 7)).setFontSize(12)
        SpreadsheetApp.getActiveSpreadsheet().getRange('J' + cnt).setValue(sum).setFontWeight("bold")
        SpreadsheetApp.getActiveSpreadsheet().getRange('A' + (cnt + 2) + ':J' + (cnt + 2)).merge()
        SpreadsheetApp.getActiveSpreadsheet().getRange('A' + (cnt + 3) + ':J' + (cnt + 3)).merge()
      }
      if (netSquare.length != 0) {
        const sum = netSquare.reduce((a, b) => a + b)
        SpreadsheetApp.getActiveSpreadsheet().getRange('A' + (cnt + 6) + ':B' + (cnt + 6)).merge().setHorizontalAlignment("center")
        SpreadsheetApp.getActiveSpreadsheet().getRange('A' + (cnt + 7) + ':B' + (cnt + 7)).merge().setHorizontalAlignment("center")
        SpreadsheetApp.getActiveSpreadsheet().getRange('E' + (cnt + 6)).setHorizontalAlignment("center")
        SpreadsheetApp.getActiveSpreadsheet().getRange('E' + (cnt + 7)).setHorizontalAlignment("center")
        SpreadsheetApp.getActiveSpreadsheet().getRange('A' + (cnt + 2)).setValue('Всего оказано услуг на сумму: ' + numberToString(sum) + '.').setFontWeight("bold")
        SpreadsheetApp.getActiveSpreadsheet().getRange('A' + (cnt + 3)).setValue('Вышеперечисленные работы (услуги) выполнены полностью и в срок. Заказчик претензий по объему, качеству и срокам оказания услуг не имеет.').setFontWeight("bold")
        if (cnt >= 59 && cnt <= 67) {
          cnt += 2
        }
        console.log(cnt)
        SpreadsheetApp.getActiveSpreadsheet().getRange('A' + (cnt + 6)).setValue('Исполнитель').setFontWeight("bold")
        SpreadsheetApp.getActiveSpreadsheet().getRange('A' + (cnt + 7)).setValue('М.П.').setFontWeight("bold")
        SpreadsheetApp.getActiveSpreadsheet().getRange('E' + (cnt + 6)).setValue('Заказчик').setFontWeight("bold")
        SpreadsheetApp.getActiveSpreadsheet().getRange('E' + (cnt + 7)).setValue('М.П.').setFontWeight("bold")
        SpreadsheetApp.getActiveSpreadsheet().getRange('A12:A').setHorizontalAlignment("center")
        SpreadsheetApp.getActiveSpreadsheet().getRange('B12:H' + cnt).setNumberFormat("# ##0.00").setHorizontalAlignment("center")
        SpreadsheetApp.getActiveSpreadsheet().getRange('J12:L' + cnt).setNumberFormat("# ##0.00").setHorizontalAlignment("center")
        SpreadsheetApp.getActiveSpreadsheet().getRange('I12:I' + cnt).setHorizontalAlignment("center")
        SpreadsheetApp.getActiveSpreadsheet().getRange('A' + cnt).setHorizontalAlignment('left')
        SpreadsheetApp.getActiveSpreadsheet().getRange('A' + (cnt + 1)).setHorizontalAlignment('left')
        SpreadsheetApp.getActiveSpreadsheet().getRange('A' + (cnt + 2)).setHorizontalAlignment('left')
        SpreadsheetApp.getActiveSpreadsheet().getRange('A' + (cnt + 3)).setHorizontalAlignment('left')
        console.log('Ínserting image...')
        var blob = DriveApp.getFileById('1l-5-q1VqdqM7Ecym7VNxuXoW8GQVBCF-').getBlob()
        SpreadsheetApp.getActiveSpreadsheet().insertImage(blob, 6, cnt + 6)
        blob = DriveApp.getFileById('1REiHEJcT_NgGfENhY3gpjFfJTQHy8jup').getBlob()
        SpreadsheetApp.getActiveSpreadsheet().insertImage(blob, 5, cnt + 7)
        if (!(SpreadsheetApp.getActiveSpreadsheet().getRange('A' + cnt).isBlank()) || !(SpreadsheetApp.getActiveSpreadsheet().getRange('A' + (cnt + 3)).isBlank())) fileId = PDF(dateFrom, action, contractor, city)
        files.push(fileId)
      }
    }
    if (data[data.length - 1][0] >= dateFrom) {
      var attachments = []
      console.log(files)
      for (var f = 0; f < files.length; f++) {
        var file = DriveApp.getFileById(files[f]).getAs(MimeType.PDF)
        attachments.push(file)
      }
      GmailApp.createDraft(email, 'Эссистэнс: БО г. ' + city + ' ' + monthToRussianString2(Utilities.formatDate(dateFrom, Session.getScriptTimeZone(), "MMM")) + ' 2022', "", { attachments: attachments, htmlBody: "<p>Добрый день!</p><p>Во вложении БО по городу " + city + ". Жду от Вас закрывающие документы <u>той же датой, что и БО</u> + подписанный БО.</p><p>Заранее спасибо!</p>" + '<table border="0" cellspacing="0" cellpadding="0" width="220" style="font-size:12.8px;color:rgb(80,0,80);width:165pt;border-collapse:collapse"><tbody><tr style="height:53.15pt"><td width="104" valign="bottom" style="width:77.95pt;border-top:none;border-bottom:none;border-left:none;border-right:1.5pt solid rgb(31,183,174);padding:0cm 5.4pt;height:53.15pt"><p align="center" style="text-align:center"><span style="font-size:9.5pt"><u></u><u></u></span></p></td><td width="189" valign="top" style="width:5cm;padding:0cm 5.4pt;height:53.15pt"><p style="background-image:initial;background-position:initial;background-repeat:initial"><font color="#0d2b4c" face="Arial, sans-serif"><span style="font-size:12.6667px">Команда Эссистэнс<br><br>@:&nbsp;</span></font><span style="color:rgb(228,175,10);font-family:&quot;Helvetica Neue&quot;;font-size:12px"><a href="mailto:ya@4ssistance.com" style="color:rgb(17,85,204)" target="_blank">ya@4ssistance.com</a></span><span style="font-family:&quot;Helvetica Neue&quot;;font-size:12px">&nbsp;<br><br><span style="font-size:9pt;font-family:Arial;color:rgb(13,43,76);background-image:initial;background-position:initial;background-repeat:initial">Cell:&nbsp;</span><span style="font-size:9.5pt;font-family:Arial;color:rgb(13,43,76);background-image:initial;background-position:initial;background-repeat:initial">+7&nbsp;</span></span><font color="#0d2b4c" face="Arial"><span style="font-size:12.6667px">905 746 40 48</span></font><span style="font-family:&quot;Helvetica Neue&quot;;font-size:12px"><br></span></p></td></tr><tr style="height:8.3pt"><td width="293" colspan="2" valign="bottom" style="width:219.7pt;padding:0cm 5.4pt;height:8.3pt"><p style="background-image:initial;background-position:initial;background-repeat:initial"><b style="font-size:12.8px"><span style="font-size:11pt;font-family:Arial,sans-serif;color:rgb(31,183,174)">——————————————————</span></b></p></td></tr></tbody></table><p style="font-size:12.8px;text-indent:7.1pt;background-image:initial;background-position:initial;background-repeat:initial"><span lang="EN-US" style="font-size:8pt;font-family:Arial;color:rgb(13,43,76)">A</span><span style="font-size:8pt;font-family:Arial;color:rgb(13,43,76)">ssistance</span><span style="font-size:8pt;font-family:Arial">&nbsp;<span style="color:rgb(13,43,76)">| 3 ul. Yamskogo Polya, 20/1, 502 |&nbsp;Moscow | Russia</span></span></p><p style="font-size:12.8px;margin-top:6pt;text-indent:7.1pt;background-image:initial;background-position:initial;background-repeat:initial"></p><p style="font-size:12.8px;margin-top:0cm;margin-bottom:0.0001pt;text-align:justify"><span style="font-size:12.8px;text-align:start"></span></p><p style="margin-top:6pt;text-indent:7.1pt;background-image:initial;background-position:initial;background-repeat:initial"><font color="#0d2b4c" face="Arial"><span style="font-size:8pt">Tel.&nbsp;</span><span style="font-size:10.6667px">+7 499 322 22 21 доб. 999&nbsp;</span><span style="font-size:8pt">|&nbsp;</span></font><a href="http://www.4ssistance.com/" style="color:rgb(17,85,204);font-family:Arial;font-size:8pt" target="_blank">www.4ssistance.com</a></p>' })
    }
  }
  
  function createReport2() { main(2) }
  function createReport3() { main(3) }
  function createReport4() { main(4) }
  function createReport5() { main(5) }
  function createReport6() { main(6) }
  function createReport7() { main(7) }
  function createReport8() { main(8) }
  function createReport9() { main(9) }
  function createReport10() { main(10) }
  function createReport11() { main(11) }
  function createReport12() { main(12) }
  function createReport13() { main(13) }
  function createReport14() { main(14) }
  function createReport15() { main(15) }
  function createReport16() { main(16) }
  function createReport17() { main(17) }
  function createReport18() { main(18) }
  function createReport19() { main(19) }
  function createReport20() { main(20) }
  function createReport21() { main(21) }
  function createReport22() { main(22) }
  function createReport23() { main(23) }
  function createReport24() { main(24) }
  function createReport25() { main(25) }
  function createReport26() { main(26) }
  function createReport27() { main(27) }
  function createReport28() { main(28) }
  function createReport29() { main(29) }
  function createReport30() { main(30) }
  function createReport31() { main(31) }
  function createReport32() { main(32) }
  function createReport33() { main(33) }
  function createReport34() { main(34) }
  function createReport35() { main(35) }
  function createReport36() { main(36) }
  function createReport37() { main(37) }
  function createReport38() { main(38) }
  function createReport39() { main(39) }
  function createReport40() { main(40) }
  function createReport41() { main(41) }
  function createReport42() { main(42) }
  function createReport43() { main(43) }
  function createReport44() { main(44) }
  function createReport45() { main(45) }
  function createReport46() { main(46) }
  function createReport47() { main(47) }
  function createReport48() { main(48) }
  function createReport49() { main(49) }
  function createReport50() { main(50) }
  function createReport51() { main(51) }
  function createReport52() { main(52) }
  function createReport53() { main(53) }
  function createReport54() { main(54) }
  function createReport55() { main(55) }
  function createReport56() { main(56) }
  function createReport57() { main(57) }
  function createReport58() { main(58) }
  function createReport59() { main(59) }
  function createReport60() { main(60) }
  function createReport61() { main(61) }
  function createReport62() { main(62) }
  function createReport63() { main(63) }
  function createReport64() { main(64) }
  function createReport65() { main(65) }
  function createReport66() { main(66) }
  function createReport67() { main(67) }
  function createReport68() { main(68) }
  function createReport69() { main(69) }
  function createReport70() { main(70) }
  function createReport71() { main(71) }
  function createReport72() { main(72) }
  function createReport73() { main(73) }
  function createReport74() { main(74) }
  function createReport75() { main(75) }
  function createReport76() { main(76) }
  function createReport77() { main(77) }
  function createReport78() { main(78) }
  function createReport79() { main(79) }
  function createReport80() { main(80) }
  function createReport81() { main(81) }
  function createReport82() { main(82) }
  function createReport83() { main(83) }
  function createReport84() { main(84) }
  function createReport85() { main(85) }
  function createReport86() { main(86) }
  function createReport87() { main(87) }
  function createReport88() { main(88) }
  function createReport89() { main(89) }
  function createReport90() { main(90) }
  function createReport91() { main(91) }
  function createReport92() { main(92) }
  function createReport93() { main(93) }
  function createReport94() { main(94) }
  function createReport95() { main(95) }
  function createReport96() { main(96) }
  function createReport97() { main(97) }
  function createReport98() { main(98) }
  function createReport99() { main(99) }
  function createReport100() { main(100) }
  function createReport101() { main(101) }
  function createReport102() { main(102) }
  function createReport103() { main(103) }
  function createReport104() { main(104) }
  function createReport105() { main(105) }
  function createReport106() { main(106) }
  function createReport107() { main(107) }
  function createReport108() { main(108) }
  function createReport109() { main(109) }
  function createReport110() { main(110) }
  function createReport111() { main(111) }
  function createReport112() { main(112) }
  function createReport113() { main(113) }
  function createReport114() { main(114) }
  function createReport115() { main(115) }
  function createReport116() { main(116) }
  function createReport117() { main(117) }
  function createReport118() { main(118) }
  function createReport119() { main(119) }
  function createReport120() { main(120) }
  function createReport121() { main(121) }
  function createReport122() { main(122) }
  function createReport123() { main(123) }
  function createReport124() { main(124) }
  function createReport125() { main(125) }
  function createReport126() { main(126) }
  function createReport127() { main(127) }
  function createReport128() { main(128) }
  function createReport129() { main(129) }
  function createReport130() { main(130) }
  function createReport131() { main(131) }
  function createReport132() { main(132) }
  function createReport133() { main(133) }
  function createReport134() { main(134) }
  function createReport135() { main(135) }
  function createReport136() { main(136) }
  function createReport137() { main(137) }
  function createReport138() { main(138) }
  function createReport139() { main(139) }
  function createReport140() { main(140) }
  function createReport141() { main(141) }
  function createReport142() { main(142) }
  function makeAll() {
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName('БО управление скриптом').activate()
    var lastRow = getFirstEmptyRowByColumnArray("B")
    if (SpreadsheetApp.getActiveSpreadsheet().getRange('G2').getValue() == '')
      for (var i = 2; i < lastRow; i++) main(i)
    else for (var i = SpreadsheetApp.getActiveSpreadsheet().getRange('G2').getValue(); i < lastRow; i++) main(i)
  }
  
  function makeSummary() {
    var summaryMonth = new Date(SpreadsheetApp.getActiveSpreadsheet().getSheetByName("БО управление скриптом").getRange("G7").getValue()).getMonth()
    var summaryYear = new Date(SpreadsheetApp.getActiveSpreadsheet().getSheetByName("БО управление скриптом").getRange("G7").getValue()).getFullYear()
    createSheet(monthToRussianString3(summaryMonth) + " " + summaryYear + " Summary")
    SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange("A1").setFormula('=QUERY(ARRAYFORMULA({IMPORTRANGE("1H92njsZuOkB_z7Sxbx3DoGOcg8py-ky6C4abKO5kkfY", "Massive!A2:N")}), "SELECT Col4, Col3, Col2, SUM(Col14) WHERE Col1 >= DATE ' + "'" + '"' + '&TEXT(DATEVALUE("' + Utilities.formatDate(new Date(new Date(SpreadsheetApp.getActiveSpreadsheet().getSheetByName("БО управление скриптом").getRange("G7").getValue()).getFullYear(), new Date(SpreadsheetApp.getActiveSpreadsheet().getSheetByName("БО управление скриптом").getRange("G7").getValue()).getMonth(), 1), Session.getScriptTimeZone(), "yyyy-MM-dd") + '"),"yyyy-MM-dd")&"' + "'" + 'AND Col1 <= DATE ' + "'" + '"' + '&TEXT(DATEVALUE("' + Utilities.formatDate(new Date(new Date(SpreadsheetApp.getActiveSpreadsheet().getSheetByName("БО управление скриптом").getRange("G7").getValue()).getFullYear(), new Date(SpreadsheetApp.getActiveSpreadsheet().getSheetByName("БО управление скриптом").getRange("G7").getValue()).getMonth() + 1, 0), Session.getScriptTimeZone(), "yyyy-MM-dd") + '"),"yyyy-MM-dd")&"' + "'" + 'AND Col5 CONTAINS ' + "'" + 'ЗАКРЫТ' + "'" + ' GROUP BY Col2, Col3, Col4 LABEL SUM(Col14) ' + "'" + "Итого без НДС', Col4 'Акция', Col3 'Город', Col2 'Подрядчик'" + '")')
    SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange("E1:Q1").setValues([["НДС 20%", "Итого с НДС 20%", "Агентское вознаграждение", "Полная сумма", "Подписанные документы", "Дата получения документов", "Платежка", "Дата ПП", "Закрывающие документы", "Комментарии", "ЮЛ Клиента", "Номер счета", "№ отчета агента"]])
    SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange("E2:H2").setFormulas([['=IFERROR(IF(VLOOKUP(R[0]C[-2], IMPORTRANGE("1H92njsZuOkB_z7Sxbx3DoGOcg8py-ky6C4abKO5kkfY", "Providers_list 2022!B2:X"), 23, FALSE) = "ДА", R[0]C[-1]*0.2, 0))', '=IFERROR(R[0]C[-2]+R[0]C[-1])', '=IFERROR(ifs(5%*R[0]C[-1]<=0,0,5%*R[0]C[-1]>=25000,25000,5%*R[0]C[-1]>0,5%*R[0]C[-1]))', '=IFERROR(R[0]C[-2]+R[0]C[-1])']])
    var filldownRange = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange(2, 5, SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getLastRow() - 1, 4)
    var lastRow = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getLastRow()
    SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange("E2:H2").copyTo(filldownRange)
    SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange("C" + (lastRow + 20)).setValue("Итого:").setFontWeight("bold")
    SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange("D" + (lastRow + 20) + ":H" + (lastRow + 20)).setFormulas([["=SUM(D2:D" + lastRow + ")", "=SUM(E2:E" + lastRow + ")", "=SUM(F2:F" + lastRow + ")", "=SUM(G2:G" + lastRow + ")", "=SUM(H2:H" + lastRow + ")"]]).setFontWeight("bold")
    var range = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange("A1:Q1")
    range.setWrap(true)
    range.setFontSize(10)
    range.setFontWeight("bold")
    range.setHorizontalAlignment("center")
    range.setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP)
    range.setVerticalAlignment("middle")
    SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().setColumnWidths(1, 22, 120)
    SpreadsheetApp.getActiveSpreadsheet().getRange('D2:H' + SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getLastRow()).setNumberFormat("# ### ### ##0.00").setHorizontalAlignment("center")
    var rule = SpreadsheetApp.newDataValidation().requireValueInList(["Яндекс.Такси Лавка ООО", "Яндекс.Такси Маркет ООО", "Яндекс.Такси Еда ООО", "Яндекс.Такси ООО", "Яндекс.Такси Доставка ООО", "Яндекс.Такси BV"], true).setAllowInvalid(false).build()
    SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange("O2:O").setDataValidation(rule).setHorizontalAlignment("center")
    SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().setColumnWidths(15, 1, 200)
    rule = SpreadsheetApp.newDataValidation().requireDate().setAllowInvalid(false).setHelpText("Введите дату в формате ДД/ММ/ГГ").build()
    SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange("J2:J").setDataValidation(rule)
    SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange("J2:J").setNumberFormat("dd.mm.yy").setHorizontalAlignment("center")
    SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange("L2:L").setDataValidation(rule)
    SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange("L2:L").setNumberFormat("dd.mm.yy").setHorizontalAlignment("center")
    rule = SpreadsheetApp.newDataValidation().requireFormulaSatisfied("=ISURL(I2)").setAllowInvalid(false).setHelpText("Введите ссылку").build()
    SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange("I2:I").setDataValidation(rule)
    SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange("I1:I").setFontColor("blue")
    rule = SpreadsheetApp.newDataValidation().requireFormulaSatisfied("=ISURL(K2)").setAllowInvalid(false).setHelpText("Введите ссылку").build()
    SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange("K2:K").setDataValidation(rule)
    SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange("K1:K").setFontColor("green")
    rule = SpreadsheetApp.newDataValidation().requireFormulaSatisfied("=ISURL(M2)").setAllowInvalid(false).setHelpText("Введите ссылку").build()
    SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange("M2:M").setDataValidation(rule)
    SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange("M1:M").setFontColor("red")
    rule = SpreadsheetApp.newDataValidation().requireNumberBetween(-999999999999, 999999999999).setAllowInvalid(false).setHelpText("Введите численное значение").build()
  }
  
  function freezeValues(range) {
    var range = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange(range)
    range.copyTo(range, { contentsOnly: true })
  }
  
  function createAndFillPredictionYandex(date) {
    var predictionMonth = date.getMonth()
    var predictionYear = date.getFullYear()
    var lastDay = function (y, m) {
      return new Date(y, m + 1, 0).getDate()
    }
    SpreadsheetApp.setActiveSpreadsheet(SpreadsheetApp.openById("1SoWMPeLfp-zrFBSOoTngwmBuGXougWiNoEdLfiTxNNc"))
    createSheet("Yandex " + monthToRussianString3(predictionMonth) + " " + predictionYear + " прогноз")
    SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange("I1").setFormula('=WORKDAY("01.' + (predictionMonth + 1) + "." + predictionYear + '";' + 'NETWORKDAYS("01.' + + (predictionMonth + 1) + "." + predictionYear + '"; EOMONTH("01.' + (predictionMonth + 1) + "." + predictionYear + '";0)) - 4)')
    SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange("I2").setFormula("=I1+1")
    SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange("A1").setFormula('=QUERY(IMPORTRANGE("1inuctJeM-6eflNtNyyDwdkCU3oFsyoHJNvPclStLQ7w"; "Массив квот!B2:V"); "SELECT Col4, (SUM(Col8) + SUM(Col16)), (SUM(Col11) + SUM(Col19)), (SUM(Col10) + SUM(Col18)), (SUM(Col13) + SUM(Col21)) WHERE Col1 >= DATE ' + "'" + '"&TEXT(DATEVALUE("1/' + (predictionMonth + 1) + '/' + predictionYear + '"' + '); "yyyy-mm-dd")&"' + "' AND Col7 = 'Yandex Go' GROUP BY Col4 LABEL Col4 'Город', (SUM(Col8) + SUM(Col16)) 'Квота', (SUM(Col11) + SUM(Col19)) 'Факт', (SUM(Col10) + SUM(Col18)) 'Остаток', (SUM(Col13) + SUM(Col21)) 'В записи'" + '")')
    SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange("F1").setValue("Прогноз")
    SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange("F2").setFormula('=ARRAYFORMULA(ROUND(IF(QUERY(IMPORTRANGE("1inuctJeM-6eflNtNyyDwdkCU3oFsyoHJNvPclStLQ7w"; "Массив квот!B2:V"); "SELECT (SUM(Col11) + SUM(Col19))/(SUM(Col8) + SUM(Col16)) WHERE Col1 >= DATE ' + "'" + '"' + '&TEXT(DATEVALUE("1/' + (predictionMonth + 1) + '/' + predictionYear + '"' + '); "yyyy-mm-dd")&"' + "' AND Col7 = 'Yandex Go' GROUP BY Col4 LABEL (SUM(Col11) + SUM(Col19))/(SUM(Col8) + SUM(Col16)) ''" + '"' + ') >= 0,5; QUERY(IMPORTRANGE("1inuctJeM-6eflNtNyyDwdkCU3oFsyoHJNvPclStLQ7w"; "Массив квот!B2:V"); "SELECT (SUM(Col10) + SUM(Col18)) WHERE Col1 >= DATE ' + "'" + '"' + '&TEXT(DATEVALUE("1/' + (predictionMonth + 1) + '/' + predictionYear + '"' + '); "yyyy-mm-dd")&"' + "' AND Col7 = 'Yandex Go' GROUP BY Col4 LABEL (SUM(Col10) + SUM(Col18)) ''" + '"); QUERY(IMPORTRANGE("1inuctJeM-6eflNtNyyDwdkCU3oFsyoHJNvPclStLQ7w"; "Массив квот!B2:V"); "SELECT (SUM(Col11) + SUM(Col19))/(SUM(Col8) + SUM(Col16)) * (SUM(Col10) + SUM(Col18)) WHERE Col1 >= DATE ' + "'" + '"&TEXT(DATEVALUE("1/' + (predictionMonth + 1) + '/' + predictionYear + '"' + '); "yyyy-mm-dd")&"' + "' AND Col7 = 'Yandex Go' GROUP BY Col4 LABEL (SUM(Col11) + SUM(Col19))/(SUM(Col8) + SUM(Col16)) * (SUM(Col10) + SUM(Col18)) ''" + '"' + "))))")
    SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange("G1").setFormula('=QUERY(IMPORTRANGE("1inuctJeM-6eflNtNyyDwdkCU3oFsyoHJNvPclStLQ7w"; "Массив квот!B2:V"); "SELECT ((SUM(Col12) + SUM(Col20))*1.05) WHERE Col1 >= DATE ' + "'" + '"&TEXT(DATEVALUE("1/' + (predictionMonth + 1) + '/' + predictionYear + '"' + '); "yyyy-mm-dd")&"' + "' AND Col7 = 'Yandex Go' GROUP BY Col4 LABEL ((SUM(Col12) + SUM(Col20))*1.05) 'Факт 01." + ((predictionMonth + 1) < 10 ? "0" + (predictionMonth + 1) : (predictionMonth + 1)) + "." + predictionYear + " - " + SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange("I1").getDisplayValue() + "'" + '")')
    SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange("H1").setValue("Прогноз " + SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange("I2").getDisplayValue() + " - " + lastDay(predictionYear, predictionMonth) + "." + ((predictionMonth + 1) < 10 ? "0" + (predictionMonth + 1) : (predictionMonth + 1)) + "." + predictionYear)
    SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange("H2").setFormula("=ARRAYFORMULA(ROUND(IF(C2:C = 0; 0; G2:G/C2:C*F2:F)))")
    return SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getLastRow()
  }
  
  function createAndFillPredictionUber(date) {
    var predictionMonth = date.getMonth()
    var predictionYear = date.getFullYear()
    var lastDay = function (y, m) {
      return new Date(y, m + 1, 0).getDate()
    }
    SpreadsheetApp.setActiveSpreadsheet(SpreadsheetApp.openById("1SoWMPeLfp-zrFBSOoTngwmBuGXougWiNoEdLfiTxNNc"))
    createSheet("Uber " + monthToRussianString3(predictionMonth) + " " + predictionYear + " прогноз")
    SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange("I1").setFormula('=WORKDAY("01.' + (predictionMonth + 1) + "." + predictionYear + '";' + 'NETWORKDAYS("01.' + + (predictionMonth + 1) + "." + predictionYear + '"; EOMONTH("01.' + (predictionMonth + 1) + "." + predictionYear + '";0)) - 4)')
    SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange("I2").setFormula("=I1+1")
    SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange("A1").setFormula('=QUERY(IMPORTRANGE("1inuctJeM-6eflNtNyyDwdkCU3oFsyoHJNvPclStLQ7w"; "Массив квот!B2:V"); "SELECT Col4, (SUM(Col8) + SUM(Col16)), (SUM(Col11) + SUM(Col19)), (SUM(Col10) + SUM(Col18)), (SUM(Col13) + SUM(Col21)) WHERE Col1 >= DATE ' + "'" + '"&TEXT(DATEVALUE("1/' + (predictionMonth + 1) + '/' + predictionYear + '"' + '); "yyyy-mm-dd")&"' + "' AND Col7 CONTAINS 'Uber' GROUP BY Col4 LABEL Col4 'Город', (SUM(Col8) + SUM(Col16)) 'Квота', (SUM(Col11) + SUM(Col19)) 'Факт', (SUM(Col10) + SUM(Col18)) 'Остаток', (SUM(Col13) + SUM(Col21)) 'В записи'" + '")')
    SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange("F1").setValue("Прогноз")
    SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange("F2").setFormula('=ARRAYFORMULA(ROUND(IF(QUERY(IMPORTRANGE("1inuctJeM-6eflNtNyyDwdkCU3oFsyoHJNvPclStLQ7w"; "Массив квот!B2:V"); "SELECT (SUM(Col11) + SUM(Col19))/(SUM(Col8) + SUM(Col16)) WHERE Col1 >= DATE ' + "'" + '"' + '&TEXT(DATEVALUE("1/' + (predictionMonth + 1) + '/' + predictionYear + '"' + '); "yyyy-mm-dd")&"' + "' AND Col7 CONTAINS 'Uber' GROUP BY Col4 LABEL (SUM(Col11) + SUM(Col19))/(SUM(Col8) + SUM(Col16)) ''" + '"' + ') >= 0,5; QUERY(IMPORTRANGE("1inuctJeM-6eflNtNyyDwdkCU3oFsyoHJNvPclStLQ7w"; "Массив квот!B2:V"); "SELECT (SUM(Col10) + SUM(Col18)) WHERE Col1 >= DATE ' + "'" + '"' + '&TEXT(DATEVALUE("1/' + (predictionMonth + 1) + '/' + predictionYear + '"' + '); "yyyy-mm-dd")&"' + "' AND Col7 CONTAINS 'Uber' GROUP BY Col4 LABEL (SUM(Col10) + SUM(Col18)) ''" + '"); QUERY(IMPORTRANGE("1inuctJeM-6eflNtNyyDwdkCU3oFsyoHJNvPclStLQ7w"; "Массив квот!B2:V"); "SELECT (SUM(Col11) + SUM(Col19))/(SUM(Col8) + SUM(Col16)) * (SUM(Col10) + SUM(Col18)) WHERE Col1 >= DATE ' + "'" + '"&TEXT(DATEVALUE("1/' + (predictionMonth + 1) + '/' + predictionYear + '"' + '); "yyyy-mm-dd")&"' + "' AND Col7 CONTAINS 'Uber' GROUP BY Col4 LABEL (SUM(Col11) + SUM(Col19))/(SUM(Col8) + SUM(Col16)) * (SUM(Col10) + SUM(Col18)) ''" + '"' + "))))")
    SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange("G1").setFormula('=QUERY(IMPORTRANGE("1inuctJeM-6eflNtNyyDwdkCU3oFsyoHJNvPclStLQ7w"; "Массив квот!B2:V"); "SELECT ((SUM(Col12) + SUM(Col20))*1.05) WHERE Col1 >= DATE ' + "'" + '"&TEXT(DATEVALUE("1/' + (predictionMonth + 1) + '/' + predictionYear + '"' + '); "yyyy-mm-dd")&"' + "' AND Col7 CONTAINS 'Uber' GROUP BY Col4 LABEL ((SUM(Col12) + SUM(Col20))*1.05) 'Факт 01." + ((predictionMonth + 1) < 10 ? "0" + (predictionMonth + 1) : (predictionMonth + 1)) + "." + predictionYear + " - " + SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange("I1").getDisplayValue() + "'" + '")')
    SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange("H1").setValue("Прогноз " + SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange("I2").getDisplayValue() + " - " + lastDay(predictionYear, predictionMonth) + "." + ((predictionMonth + 1) < 10 ? "0" + (predictionMonth + 1) : (predictionMonth + 1)) + "." + predictionYear)
    SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange("H2").setFormula("=ARRAYFORMULA(ROUND(IF(C2:C = 0; 0; G2:G/C2:C*F2:F)))")
    return SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getLastRow()
  }
  
  function createAndFillPredictionVezet(date) {
    var predictionMonth = date.getMonth()
    var predictionYear = date.getFullYear()
    var lastDay = function (y, m) {
      return new Date(y, m + 1, 0).getDate()
    }
    SpreadsheetApp.setActiveSpreadsheet(SpreadsheetApp.openById("1SoWMPeLfp-zrFBSOoTngwmBuGXougWiNoEdLfiTxNNc"))
    createSheet("Везет " + monthToRussianString3(predictionMonth) + " " + predictionYear + " прогноз")
    SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange("I1").setFormula('=WORKDAY("01.' + (predictionMonth + 1) + "." + predictionYear + '";' + 'NETWORKDAYS("01.' + + (predictionMonth + 1) + "." + predictionYear + '"; EOMONTH("01.' + (predictionMonth + 1) + "." + predictionYear + '";0)) - 4)')
    SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange("I2").setFormula("=I1+1")
    SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange("A1").setFormula('=QUERY(IMPORTRANGE("1inuctJeM-6eflNtNyyDwdkCU3oFsyoHJNvPclStLQ7w"; "Массив квот!B2:V"); "SELECT Col4, (SUM(Col8) + SUM(Col16)), (SUM(Col11) + SUM(Col19)), (SUM(Col10) + SUM(Col18)), (SUM(Col13) + SUM(Col21)) WHERE Col1 >= DATE ' + "'" + '"&TEXT(DATEVALUE("1/' + (predictionMonth + 1) + '/' + predictionYear + '"' + '); "yyyy-mm-dd")&"' + "' AND Col7 = 'Везет' GROUP BY Col4 LABEL Col4 'Город', (SUM(Col8) + SUM(Col16)) 'Квота', (SUM(Col11) + SUM(Col19)) 'Факт', (SUM(Col10) + SUM(Col18)) 'Остаток', (SUM(Col13) + SUM(Col21)) 'В записи'" + '")')
    SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange("F1").setValue("Прогноз")
    SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange("F2").setFormula('=ARRAYFORMULA(ROUND(IF(QUERY(IMPORTRANGE("1inuctJeM-6eflNtNyyDwdkCU3oFsyoHJNvPclStLQ7w"; "Массив квот!B2:V"); "SELECT (SUM(Col11) + SUM(Col19))/(SUM(Col8) + SUM(Col16)) WHERE Col1 >= DATE ' + "'" + '"' + '&TEXT(DATEVALUE("1/' + (predictionMonth + 1) + '/' + predictionYear + '"' + '); "yyyy-mm-dd")&"' + "' AND Col7 = 'Везет' GROUP BY Col4 LABEL (SUM(Col11) + SUM(Col19))/(SUM(Col8) + SUM(Col16)) ''" + '"' + ') >= 0,5; QUERY(IMPORTRANGE("1inuctJeM-6eflNtNyyDwdkCU3oFsyoHJNvPclStLQ7w"; "Массив квот!B2:V"); "SELECT (SUM(Col10) + SUM(Col18)) WHERE Col1 >= DATE ' + "'" + '"' + '&TEXT(DATEVALUE("1/' + (predictionMonth + 1) + '/' + predictionYear + '"' + '); "yyyy-mm-dd")&"' + "' AND Col7 = 'Везет' GROUP BY Col4 LABEL (SUM(Col10) + SUM(Col18)) ''" + '"); QUERY(IMPORTRANGE("1inuctJeM-6eflNtNyyDwdkCU3oFsyoHJNvPclStLQ7w"; "Массив квот!B2:V"); "SELECT (SUM(Col11) + SUM(Col19))/(SUM(Col8) + SUM(Col16)) * (SUM(Col10) + SUM(Col18)) WHERE Col1 >= DATE ' + "'" + '"&TEXT(DATEVALUE("1/' + (predictionMonth + 1) + '/' + predictionYear + '"' + '); "yyyy-mm-dd")&"' + "' AND Col7 = 'Везет' GROUP BY Col4 LABEL (SUM(Col11) + SUM(Col19))/(SUM(Col8) + SUM(Col16)) * (SUM(Col10) + SUM(Col18)) ''" + '"' + "))))")
    SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange("G1").setFormula('=QUERY(IMPORTRANGE("1inuctJeM-6eflNtNyyDwdkCU3oFsyoHJNvPclStLQ7w"; "Массив квот!B2:V"); "SELECT ((SUM(Col12) + SUM(Col20))*1.05) WHERE Col1 >= DATE ' + "'" + '"&TEXT(DATEVALUE("1/' + (predictionMonth + 1) + '/' + predictionYear + '"' + '); "yyyy-mm-dd")&"' + "' AND Col7 = 'Везет' GROUP BY Col4 LABEL ((SUM(Col12) + SUM(Col20))*1.05) 'Факт 01." + ((predictionMonth + 1) < 10 ? "0" + (predictionMonth + 1) : (predictionMonth + 1)) + "." + predictionYear + " - " + SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange("I1").getDisplayValue() + "'" + '")')
    SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange("H1").setValue("Прогноз " + SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange("I2").getDisplayValue() + " - " + lastDay(predictionYear, predictionMonth) + "." + ((predictionMonth + 1) < 10 ? "0" + (predictionMonth + 1) : (predictionMonth + 1)) + "." + predictionYear)
    SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange("H2").setFormula("=ARRAYFORMULA(ROUND(IF(C2:C = 0; 0; G2:G/C2:C*F2:F)))")
    return SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getLastRow()
  }
  
  function makePrediction() {
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("БО управление скриптом").activate()
    var date = new Date(SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange("H7").getValue())
    var i = createAndFillPredictionYandex(date)
    if (i >= 100) {
      freezeValues("A1:H")
      SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().deleteColumns(2, 5)
      SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange("A1:C1").setFontWeight("bold").setHorizontalAlignment("center").setWrap(true)
      SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange("B2:C").setHorizontalAlignment("center").setNumberFormat("# ### ### ##0")
      SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange("A" + getFirstEmptyRowByColumnArray("A")).setValue("Total").setFontWeight("bold")
      SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange("B" + (getFirstEmptyRowByColumnArray("A") - 1)).setFormula("=SUM(B2:B" + (getFirstEmptyRowByColumnArray("A") - 2) + ")").setFontWeight("bold")
      SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange("C" + (getFirstEmptyRowByColumnArray("A") - 1)).setFormula("=SUM(C2:C" + (getFirstEmptyRowByColumnArray("A") - 2) + ")").setFontWeight("bold")
      SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange("D1:D2").clearContent()
      SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().deleteRows(getFirstEmptyRowByColumnArray("A"), SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getMaxRows() - getFirstEmptyRowByColumnArray("A") + 1)
    }
    i = createAndFillPredictionUber(date)
    if (i >= 100) {
      freezeValues("A1:H")
      SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().deleteColumns(2, 5)
      SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange("A1:C1").setFontWeight("bold").setHorizontalAlignment("center").setWrap(true)
      SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange("B2:C").setHorizontalAlignment("center").setNumberFormat("# ### ### ##0")
      SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange("A" + getFirstEmptyRowByColumnArray("A")).setValue("Total").setFontWeight("bold")
      SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange("B" + (getFirstEmptyRowByColumnArray("A") - 1)).setFormula("=SUM(B2:B" + (getFirstEmptyRowByColumnArray("A") - 2) + ")").setFontWeight("bold")
      SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange("C" + (getFirstEmptyRowByColumnArray("A") - 1)).setFormula("=SUM(C2:C" + (getFirstEmptyRowByColumnArray("A") - 2) + ")").setFontWeight("bold")
      SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange("D1:D2").clearContent()
      SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().deleteRows(getFirstEmptyRowByColumnArray("A"), SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getMaxRows() - getFirstEmptyRowByColumnArray("A") + 1)
    }
    i = createAndFillPredictionVezet(date)
    if (i >= 100) {
      freezeValues("A1:H")
      SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().deleteColumns(2, 5)
      SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange("A1:C1").setFontWeight("bold").setHorizontalAlignment("center").setWrap(true)
      SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange("B2:C").setHorizontalAlignment("center").setNumberFormat("# ### ### ##0")
      SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange("A" + getFirstEmptyRowByColumnArray("A")).setValue("Total").setFontWeight("bold")
      SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange("B" + (getFirstEmptyRowByColumnArray("A") - 1)).setFormula("=SUM(B2:B" + (getFirstEmptyRowByColumnArray("A") - 2) + ")").setFontWeight("bold")
      SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange("C" + (getFirstEmptyRowByColumnArray("A") - 1)).setFormula("=SUM(C2:C" + (getFirstEmptyRowByColumnArray("A") - 2) + ")").setFontWeight("bold")
      SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange("D1:D2").clearContent()
      SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().deleteRows(getFirstEmptyRowByColumnArray("A"), SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getMaxRows() - getFirstEmptyRowByColumnArray("A") + 1)
    }
  }