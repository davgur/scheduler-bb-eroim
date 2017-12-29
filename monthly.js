//var monthly = new LoadMonthly(1);
//monthly.run();

function LoadMonthly(_resultCurrentRowId) {
  var ss           = SpreadsheetApp.getActiveSpreadsheet();
  var resultSheet  = ss.getSheetByName('board');
  var configsSheet = ss.getSheetByName('config');
  _resultCurrentRowId++;

  var dataValues = SpreadsheetApp.openById('1LlRo5Ob5Bw8penUEakQj_1NyfmmD-T_Dama-k81dohQ').getSheetByName('Sheet1').getRange('A2:AG1500').getValues();
  dataValues     = _filter(dataValues);

  return {
    run: run
  };

  function _filter(arr) {
    var result = { yearsList: [] }, data = [];
    var _now   = new Date();
    arr.filter(function (val, key) {
      return val[5] == 3 && val[0] >= _now;
    }).sort(function (a, b) {
      var a2Num = new Number(a[0]);
      var b2Num = new Number(b[0]);
      var a5Num = new Number(a[1].split(':')[0]);
      var b5Num = new Number(b[1].split(':')[0]);
      if (a2Num < b2Num) {
        return -1;
      }
      if (a2Num > b2Num) {
        return 1;
      }
      if (a5Num < a5Num) {
        return -1;
      }
      if (a5Num > a5Num) {
        return 1;
      }
      return 0;
    }).forEach(function (val) {
      var key = new Date(val[0]).getFullYear();
      if (!result[key]) {
        result[key] = [];
        result.yearsList.push(key);
      }
      result[key].push(val);

    });
    return result;
  }

  function run() {
    dataValues.yearsList.forEach(function (year) {
      printYear(year, dataValues[year]);
    });
    resultSheet.getRange(_resultCurrentRowId + 1, 1, 50, 50).clear();
  }

  function printYear(year, data) {
    _printTableTitle(year);
    var now_month;
    data.forEach(function (row, i) {
      _resultCurrentRowId++;
      var data  = rowToObject(row);
      var _date = new Date(data.date);
      _date.setHours(_date.getHours() + 8);
      var temp_month = _date.getMonth();
      if (now_month !== temp_month) {
        _formatMonth(temp_month);
        now_month = temp_month;
        _printSeparators();
        _resultCurrentRowId++;
      }
      resultSheet.setRowHeight(_resultCurrentRowId, 24);
      //heb
      _addContent(data, 2, data.heb, 0);
      //english
      _addContent(data, 6, data.eng, 1);
      //rus
      _addContent(data, 10, data.rus, 2);
      //esp
      _addContent(data, 14, data.esp, 3);

      _printSeparators();
    });
    _resultCurrentRowId++;
    _toBlack(resultSheet.getRange(_resultCurrentRowId, 1, 1, 17));
  }

  function _printTableTitle(year) {
    _resultCurrentRowId++;
    var hebText = 'לוח אירועים שנתי ';
    var engText = 'Annual Board of Events ';
    var rusText = 'Расписание событий на ';
    var espText = 'Tabla Anual de Eventos Para ';

    //heb
    resultSheet.getRange(_resultCurrentRowId, 2, 1, 3).merge().setFontSize(24).setFontWeight('bold').setHorizontalAlignment('center').setValue(hebText + year).setBackground('#6d9eeb');
    //eng
    resultSheet.getRange(_resultCurrentRowId, 6, 1, 3).merge().setFontSize(24).setFontWeight('bold').setHorizontalAlignment('center').setValue(engText + year).setBackground('#00ff00');
    //rus
    resultSheet.getRange(_resultCurrentRowId, 10, 1, 3).merge().setFontSize(24).setFontWeight('bold').setHorizontalAlignment('center').setValue(rusText + year + ' год').setBackground('#ffff00');
    //esp
    resultSheet.getRange(_resultCurrentRowId, 14, 1, 3).merge().setFontSize(24).setFontWeight('bold').setHorizontalAlignment('center').setValue(espText + year).setBackground('#e06666');
    resultSheet.setRowHeight(_resultCurrentRowId, 50);

    _printSeparators();
    _toBlack(resultSheet.getRange(_resultCurrentRowId + 1, 1, 1, 17));
    _resultCurrentRowId++;
  }

  function _formatMonth(monthIndex) {
    var _bgColors = configsSheet.getRange(3, 11, 4, 1).getValues();
    var _months   = configsSheet.getRange(3, 1, 12, 4).getValues();
    resultSheet.setRowHeight(_resultCurrentRowId, 24);

    //heb
    resultSheet.getRange(_resultCurrentRowId, 2, 1, 3).merge().setFontSize(16).setFontWeight('bold').setHorizontalAlignment('center').setValue(_months[monthIndex][0]).setBackground(_bgColors[0][0]);
    //rus
    resultSheet.getRange(_resultCurrentRowId, 6, 1, 3).merge().setFontSize(16).setFontWeight('bold').setHorizontalAlignment('center').setValue(_months[monthIndex][1]).setBackground(_bgColors[1][0]);
    //eng
    resultSheet.getRange(_resultCurrentRowId, 10, 1, 3).merge().setFontSize(16).setFontWeight('bold').setHorizontalAlignment('center').setValue(_months[monthIndex][2]).setBackground(_bgColors[2][0]);
    //esp
    resultSheet.getRange(_resultCurrentRowId, 14, 1, 3).merge().setFontSize(16).setFontWeight('bold').setHorizontalAlignment('center').setValue(_months[monthIndex][3]).setBackground(_bgColors[3][0]);
  }

  function _addContent(data, startIndex, value, langIndex) {
    var _values   = [];
    var _weekDays = configsSheet.getRange(3, 6, 7, 4).getValues();
    var range     = resultSheet.getRange(_resultCurrentRowId, startIndex, 1, 3);
    _clearFormat(range, langIndex === 0);
    range.setFontSize(12);

    _values[2] = Utilities.formatDate(new Date(data.date), 'Asia/Jerusalem', 'dd/MM/yyyy');
    _values[1] = _weekDays[data.weekDay][langIndex];
    _values[0] = value;
    if (langIndex === 0) {
      _values.reverse();
      resultSheet.getRange(_resultCurrentRowId, startIndex, 1, 2).setFontWeight('bold');
    } else {
      resultSheet.getRange(_resultCurrentRowId, startIndex + 1, 1, 2).setFontWeight('bold');
    }

    range.setValues([_values]);
  }

  function rowToObject(row) {
    return {
      date: row[0],
      weekDay: row[0].getDay(),
      heb: row[3],
      eng: row[8],
      rus: row[11],
      esp: row[14]
    };
  }

  function _printSeparators() {
    _toBlack(resultSheet.getRange(_resultCurrentRowId, 1));
    _toBlack(resultSheet.getRange(_resultCurrentRowId, 5));
    _toBlack(resultSheet.getRange(_resultCurrentRowId, 9));
    _toBlack(resultSheet.getRange(_resultCurrentRowId, 13));
    _toBlack(resultSheet.getRange(_resultCurrentRowId, 17));
  }

  function _toBlack(range) {
    range.clear().setBackground('#cccccc');
  }

  function _clearFormat(range, isHeb) {
    range.clearFormat();
    var dir = isHeb ? 'right' : 'left';
    range.setHorizontalAlignment(dir);
  }
}