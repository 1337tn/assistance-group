function createSheet(name) {
    // Функция создания нового листа в активном файле, при создании сразу активирует новый лист
    // Входной аргумент: название листа
    console.log("Creating new sheet...")
    var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet()
    var yourNewSheet = activeSpreadsheet.getSheetByName(name)
  
    if (yourNewSheet != null) {
      activeSpreadsheet.deleteSheet(yourNewSheet)
    }
    yourNewSheet = activeSpreadsheet.insertSheet()
    yourNewSheet.setName(name)
    console.log("Finished sheet creation")
  }
  
  function getFirstEmptyRowByColumnArraySheet(sheetName, letter) {
    // Функция нахождения первой пустой строки на листе
    // Входные аргументы: название листа; буква столбца, по которому смотрим заполненную строку
    // Возвращает индекс первой пустой строки
    var spr = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName)
    var column = spr.getRange(letter + ':' + letter)
    var values = column.getValues(); // get all data in one call
    var ct = 16
    while (values[ct] && values[ct][0] != "") {
      ct++
    }
    return (ct + 1)
  }
  
  function monthToRussianString(month) {
    // Функция перевода номера месяца (в JS) в название месяца на русском
    // Входной аргумент: номер месяца (в формате JS, т.е. январь это 0)
    // Возвращает строку - месяц на русском языке
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
    // Функция перевода суммы в прописной вариант
    // Входной аргумент: численная сумма
    // Возвращает строку сумма прописью
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
  
  function createInvoice() {
    // Функция формирования счёта
    var data = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Формирование счетов").getRange("A2:B6").getValues()
    var project = data[0][0]
    var invoiceNo = data[2][0]
    var invoiceDate = Utilities.formatDate(new Date(data[2][1]), Session.getScriptTimeZone(), "dd.MM.yyyy")
    // TODO: Добавить ЮЛ
  
    // Если проект - Брендинг
    if (project == "Branding") {
      // Выставить счет по Брендингу за период:
      var brandingInvoiceMonth = monthToRussianString(new Date(data[4][0]).getMonth())
      // Активирует файл, где хранятся summary по брендингу https://docs.google.com/spreadsheets/d/1wLKoe1JtQTDqwcZfrJUPZ-5pTYjQ15xHTJG3lcY_nfA/edit#gid=1748580169
      SpreadsheetApp.setActiveSpreadsheet(SpreadsheetApp.openById("1wLKoe1JtQTDqwcZfrJUPZ-5pTYjQ15xHTJG3lcY_nfA"))
      // Забирает всё с листа summary за указанный месяц
      var massive = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(brandingInvoiceMonth + " 2021 Summary").getRange("A2:H" + (getFirstEmptyRowByColumnArraySheet(brandingInvoiceMonth + " 2021 Summary", "A") - 1)).getValues()
      // Сортировка по НДС, акции, городу
      var sortedMassive = massive.sort((a, b) => {
        return a[4] == b[4] ? (a[0].toLocaleString() === b[0].toLocaleString() ? a[1].toLocaleString().localeCompare(b[1].toLocaleString()) : a[0].toLocaleString().localeCompare(b[0].toLocaleString())) : a[4] - b[4]
      })
      // Открывает файл со счетами, ПЕРЕДЕЛАТЬ на отдельный файл "Формирование счетов"
      SpreadsheetApp.setActiveSpreadsheet(SpreadsheetApp.openById("18W8wAtlE-TkBnk-tdUeOqUROTT0ALwGw-Ntgg2hYQsc"))
      createSheet(invoiceNo)
      // Настройка шрифтов и форматов
      // TODO: Доделать рамочки
      SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange("A1:G14").setFontSize(9)
      SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().setColumnWidth(1, 33).setColumnWidth(2, 315).setColumnWidth(3, 175).setColumnWidth(4, 150).setColumnWidth(5, 150).setColumnWidth(6, 115).setColumnWidth(7, 175).setColumnWidth(8, 150)
      SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange("B2:G2").merge().setValue('Общество с ограниченной ответственностью "ЭССИСТЭНС"').setFontWeight('bold')
      SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange("B4:G4").merge().setValue('125124, г. Москва, 3-я улица Ямского Поля, д.2, стр.13, помещение XI, комната 41')
      SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange("B6:C6").merge().setValue('ИНН 7704331668')
      SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange("B7:C7").merge().setValue('Получатель')
      SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange("B8:C8").merge().setValue('Общество с ограниченной ответственностью "ЭССИСТЭНС"')
      SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange("B9:C9").merge().setValue('Банк получателя:')
      SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange("B10:C10").merge().setValue('МОСКОВСКИЙ ФИЛИАЛ АО КБ "МОДУЛЬБАНК" Г.МОСКВА')
      SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange("D6:D8").merge().setValue("Сч. №")
      SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange("E6:H8").merge().setValue("40702810670010002298").setNumberFormat('@STRING@')
      SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange("B9:C9").merge().setValue("Банк получателя:")
      SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange("D9").setValue("БИК")
      SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange("E9:H9").merge().setValue("044525092").setHorizontalAlignment('left').setNumberFormat('@STRING@')
      SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange("B10:C10").merge().setValue('МОСКОВСКИЙ ФИЛИАЛ АО КБ "МОДУЛЬБАНК" Г.МОСКВА')
      SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange("D10").setValue("Сч. №")
      SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange("E10:H10").merge().setValue("30101810645250000092").setNumberFormat('@STRING@')
      SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange("B12:H12").merge().setValue("Счёт на оплату № " + invoiceNo + " от " + invoiceDate + " г.").setHorizontalAlignment('center').setFontSize(16).setFontWeight('bold')
      SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange("B14").setValue("Плательщик:").setVerticalAlignment('top')
      // TODO: Тянуть из базы ЮЛ и реквизитов
      SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange("C14:H14").merge().setValue('ООО "Яндекс.Такси", ИНН 7704340310, КПП 770301001, р/с 40702810800001005378 в ИНГ БАНК (ЕВРАЗИЯ) АО, кор.счет 30101810500000000222, БИК 044525222').setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP)
      SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange("B16:H16").setValues([["Статья расходов", "Город", "Сумма без НДС", "НДС", "Включая НДС", "Вознаграждение", "Всего к оплате:"]]).setFontWeight('bold').setHorizontalAlignment('center')
      SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange("B17:H17").merge().setValue("Услуги, не облагаемые НДС").setFontWeight('bold')
      // Суммирование по Акции - Городу без учёта Подрядчика
      var result = []
      var resultWithNDS = []
      var sums = [0, 0, 0, 0]
      var sumsWithNDS = [0, 0, 0, 0]
      var j = 0
      var k = 0
      for (var i = 0; i < sortedMassive.length; i++) {
        console.log(sortedMassive[i])
        if (result.length != 0 && j > 0 && result[j - 1][0] == sortedMassive[i][0] && result[j - 1][1] == sortedMassive[i][1] && result[j - 1][3] == sortedMassive[i][4] && sortedMassive[i][4] == 0) {
          result[j - 1][2] += sortedMassive[i][3]
          result[j - 1][4] += sortedMassive[i][4]
          result[j - 1][5] += sortedMassive[i][5]
          result[j - 1][6] += sortedMassive[i][6]
          sums[0] += sortedMassive[i][3]
          sums[1] += sortedMassive[i][4]
          sums[2] += sortedMassive[i][5]
          sums[3] += sortedMassive[i][6]
        } else if (sortedMassive[i][4] == 0) {
          result.push([sortedMassive[i][0], sortedMassive[i][1], sortedMassive[i][3], sortedMassive[i][4], sortedMassive[i][5], sortedMassive[i][6], sortedMassive[i][7]])
          sums[0] += sortedMassive[i][3]
          sums[1] += sortedMassive[i][4]
          sums[2] += sortedMassive[i][5]
          sums[3] += sortedMassive[i][6]
          j += 1
        } else if (resultWithNDS.length != 0 && k > 0 && resultWithNDS[k - 1][0] == sortedMassive[i][0] && resultWithNDS[k - 1][1] == sortedMassive[i][1] && resultWithNDS[k - 1][3] == sortedMassive[i][4]) {
          resultWithNDS[k - 1][2] += sortedMassive[i][3]
          resultWithNDS[k - 1][4] += sortedMassive[i][4]
          resultWithNDS[k - 1][5] += sortedMassive[i][5]
          resultWithNDS[k - 1][6] += sortedMassive[i][6]
          sumsWithNDS[0] += sortedMassive[i][3]
          sumsWithNDS[1] += sortedMassive[i][4]
          sumsWithNDS[2] += sortedMassive[i][5]
          sumsWithNDS[3] += sortedMassive[i][6]
        } else {
          resultWithNDS.push([sortedMassive[i][0], sortedMassive[i][1], sortedMassive[i][3], sortedMassive[i][4], sortedMassive[i][5], sortedMassive[i][6], sortedMassive[i][7]])
          sumsWithNDS[0] += sortedMassive[i][3]
          sumsWithNDS[1] += sortedMassive[i][4]
          sumsWithNDS[2] += sortedMassive[i][5]
          sumsWithNDS[3] += sortedMassive[i][6]
          k += 1
        }
      }
      console.log(result)
      SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange("B18:H" + (17 + result.length)).setValues(result)
      SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange("B" + (18 + result.length) + ":H" + (18 + result.length)).merge().setValue("Услуги, облагаемые НДС 20%")
      SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange("B" + (19 + result.length) + ":H" + (18 + result.length + resultWithNDS.length)).setValues(resultWithNDS)
      console.log(sums, sumsWithNDS)
      // Если проект - ЯТ закупки
    } else if (project == "ЯТ закупки") {
      var invoiceDate = new Date(data[2][1])
      SpreadsheetApp.setActiveSpreadsheet(SpreadsheetApp.openById("1voKQaqYMQ3j5lamGS3BivDQ5QLIWFaxIdOJ_EiEaj-Q"))
      var massive = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Costs 2021").getRange("A2:AE").getValues()
      var resultWithNDS = {}
      var result = {}
      for (var i = 0; i < massive.length; i++) {
        if (massive[i][29] == invoiceNo && massive[i][30].valueOf() == invoiceDate.valueOf()) {
          if (massive[i][14].toFixed(1) == 0) {
            if (massive[i][10] in result) {
              result[massive[i][10]][0] += massive[i][13]
              result[massive[i][10]][1] += massive[i][15]
              result[massive[i][10]][2] += massive[i][12]
              result[massive[i][10]][3] += massive[i][17]
              result[massive[i][10]][4] += massive[i][18]
            }
            else {
              result[massive[i][10]] = [massive[i][13], massive[i][15], massive[i][12], massive[i][17], massive[i][18]]
            }
          }
          else {
            if (massive[i][10] in resultWithNDS) {
              resultWithNDS[massive[i][10]][0] += massive[i][13]
              resultWithNDS[massive[i][10]][1] += massive[i][15]
              resultWithNDS[massive[i][10]][2] += massive[i][12]
              resultWithNDS[massive[i][10]][3] += massive[i][17]
              resultWithNDS[massive[i][10]][4] += massive[i][18]
            }
            else {
              resultWithNDS[massive[i][10]] = [massive[i][13], massive[i][15], massive[i][12], massive[i][17], massive[i][18]]
            }
          }
        }
      }
      console.log(resultWithNDS)
      console.log(result)
      SpreadsheetApp.setActiveSpreadsheet(SpreadsheetApp.openById("18W8wAtlE-TkBnk-tdUeOqUROTT0ALwGw-Ntgg2hYQsc"))
      createSheet(invoiceNo)
      SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange("A1:G14").setFontSize(9)
      SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().setColumnWidth(1, 33).setColumnWidth(2, 315).setColumnWidth(3, 175).setColumnWidth(4, 150).setColumnWidth(5, 150).setColumnWidth(6, 115).setColumnWidth(7, 175).setColumnWidth(8, 150)
      SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange("B2:G2").merge().setValue('Общество с ограниченной ответственностью "ЭССИСТЭНС"').setFontWeight('bold')
      SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange("B4:G4").merge().setValue('125124, г. Москва, 3-я улица Ямского Поля, д.2, стр.13, помещение XI, комната 41')
      SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange("B6:C6").merge().setValue('ИНН 7704331668')
      SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange("B7:C7").merge().setValue('Получатель')
      SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange("B8:C8").merge().setValue('Общество с ограниченной ответственностью "ЭССИСТЭНС"')
      SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange("B9:C9").merge().setValue('Банк получателя:')
      SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange("B10:C10").merge().setValue('МОСКОВСКИЙ ФИЛИАЛ АО КБ "МОДУЛЬБАНК" Г.МОСКВА')
      SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange("D6:D8").merge().setValue("Сч. №")
      SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange("E6:G8").merge().setValue("40702810670010002298").setNumberFormat('@STRING@')
      SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange("B9:C9").merge().setValue("Банк получателя:")
      SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange("D9").setValue("БИК")
      SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange("E9:H9").merge().setValue("044525092").setHorizontalAlignment('left').setNumberFormat('@STRING@')
      SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange("B10:C10").merge().setValue('МОСКОВСКИЙ ФИЛИАЛ АО КБ "МОДУЛЬБАНК" Г.МОСКВА')
      SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange("D10").setValue("Сч. №")
      SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange("E10:G10").merge().setValue("30101810645250000092").setNumberFormat('@STRING@')
      SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange("B12:G12").merge().setValue("Счёт на оплату № " + invoiceNo + " от " + Utilities.formatDate(invoiceDate, Session.getScriptTimeZone(), "dd.MM.yyyy") + " г.").setHorizontalAlignment('center').setFontSize(16).setFontWeight('bold')
      SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange("B14").setValue("Плательщик:").setVerticalAlignment('top')
      SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange("C14:G14").merge().setValue('ООО "Яндекс.Такси", ИНН 7704340310, КПП 770301001, р/с 40702810800001005378 в ИНГ БАНК (ЕВРАЗИЯ) АО, кор.счет 30101810500000000222, БИК 044525222')
      SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange("C14").setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP)
      SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange("B16:G16").setValues([["Статья расходов", "Сумма без НДС", "НДС", "Включая НДС", "Вознаграждение", "Всего к оплате:"]]).setFontWeight('bold').setHorizontalAlignment('center')
      console.log(Object.keys(result).length)
      console.log(Object.keys(resultWithNDS).length)
      var sum = (r, a) => r.map((b, i) => a[i] + b)
      var i1 = -1
      var i2 = -1
      if (Object.keys(result).length > 0) {
        SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange("B17:G17").merge().setValue("Услуги, не облагаемые НДС:").setFontWeight('bold')
        SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange("B18:B" + (Object.entries(result).length + 17)).setValues(Object.keys(result).map(x => [x]))
        SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange("C18:G" + (Object.values(result).length + 17)).setValues(Object.values(result))
        SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange("B" + (Object.values(result).length + 18)).setValue("Итого услуги, не облагаемые НДС:").setFontWeight("bold")
        SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange("C" + (Object.values(result).length + 18) + ":G" + (Object.values(result).length + 18)).setValues([Object.values(result).reduce(sum)]).setFontWeight("bold")
        i1 = Object.values(result).length + 18
      }
      if (Object.keys(resultWithNDS).length > 0) {
        var index = getFirstEmptyRowByColumnArraySheet(invoiceNo, "B")
        SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange("B" + index + ":G" + index).merge().setValue("Услуги, облагаемые НДС:").setFontWeight('bold')
        index += 1
        SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange("B" + index + ":B" + (Object.entries(resultWithNDS).length - 1 + index)).setValues(Object.keys(resultWithNDS).map(x => [x]))
        SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange("C" + index + ":G" + (index + Object.keys(resultWithNDS).length - 1)).setValues(Object.values(resultWithNDS))
        SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange("B" + (index + Object.keys(resultWithNDS).length)).setValue("Итого услуги, облагаемые НДС:").setFontWeight("bold")
        SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange("C" + (index + Object.keys(resultWithNDS).length) + ":G" + (index + Object.keys(resultWithNDS).length)).setValues([Object.values(resultWithNDS).reduce(sum)]).setFontWeight("bold")
        i2 = index + Object.keys(resultWithNDS).length
      }
      var totals = []
      if (i1 > -1) {
        if (i2 > -1) {
          totals.push([SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange("C" + i1).getValue() + SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange("C" + i2).getValue()])
          totals.push([SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange("D" + i1).getValue() + SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange("D" + i2).getValue()])
          totals.push([SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange("G" + i1).getValue() + SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange("G" + i2).getValue()])
        }
        else {
          totals.push([SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange("C" + i1).getValue()])
          totals.push([SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange("D" + i1).getValue()])
          totals.push([SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange("G" + i1).getValue()])
        }
      } else {
        if (i2 > -1) {
          totals.push([SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange("C" + i2).getValue()])
          totals.push([SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange("D" + i2).getValue()])
          totals.push([SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange("G" + i2).getValue()])
        }
      }
      SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange("F" + (getFirstEmptyRowByColumnArraySheet(invoiceNo, "B") + 1) + ":F" + (getFirstEmptyRowByColumnArraySheet(invoiceNo, "B") + 3)).setValues([["Всего без НДС:"], ["НДС 20%:"], ["Всего к оплате:"]]).setFontWeight("bold")
      SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange("G" + (getFirstEmptyRowByColumnArraySheet(invoiceNo, "B") + 1) + ":G" + (getFirstEmptyRowByColumnArraySheet(invoiceNo, "B") + 3)).setValues(totals).setFontWeight("bold")
      if (totals[1][0] != 0) SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange("B" + (getFirstEmptyRowByColumnArraySheet(invoiceNo, "B") + 6)).setValue(numberToString(totals[2][0]) + ", в т.ч. НДС " + numberToString(totals[1][0]))
      else SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange("B" + (getFirstEmptyRowByColumnArraySheet(invoiceNo, "B") + 6)).setValue(numberToString(totals[2][0]) + ", НДС не облагается")
    }
  }