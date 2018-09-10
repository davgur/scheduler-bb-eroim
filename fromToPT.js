function runscript() {
  var weekly = new BuilderFromToPT();
  weekly.run();
}

function BuilderFromToPT() {
  var _resultCurrentRowId = 1;
  var ss                  = SpreadsheetApp.getActiveSpreadsheet();
  var resultSheet         = ss.getSheetByName('test');
  var configurations      = ss.getSheetByName('config').getRange('B2:B100').getValues();
  var _weekDays           = SpreadsheetApp.openById('1B76zdIX2p48FEA1fvJr36DHKsQaIODQT9kWUZ8n0c7o').getSheetByName('config').getRange(3, 6, 7, 4).getValues();
  var MAIN_TITLES         = {
    heb: { color: '#6d9eeb', text: configurations[3] },
    rus: { color: '#ffff00', text: configurations[2] },
  };
  var dataValues          = SpreadsheetApp.openById('1LlRo5Ob5Bw8penUEakQj_1NyfmmD-T_Dama-k81dohQ').getSheetByName('Sheet1').getRange('A3:AG3000').getValues();
  dataValues              = _filter(dataValues).map(function (row) {
    return (row[5] == 6 || row[5] == 6.1) ? _parseTitle(row) : _parseEvent(row);
  });
  dataValues.sort(_sort);
  dataValues.push({});
  return {
    run: function () {
      _printMainTitles(_resultCurrentRowId, 'heb');

      dataValues.forEach(function (obj, key, arr) {
        printIterator(obj, key, arr, _printEventHeb, _printDateHeb, obj.heb);
      });

      _printMainTitles(_resultCurrentRowId, 'rus');

      dataValues.forEach(function (obj, key, arr) {
        printIterator(obj, key, arr, _printEventRus, _printDateRus, obj.rus);
      });
      nextRow();
      clearRange(resultSheet.getRange(_resultCurrentRowId, 1, 100, 6));
      _separator(resultSheet.getRange(_resultCurrentRowId, 1, 1, 6));
    }
  };

  function printIterator(obj, key, arr, eventPrinter, dataPrinter, valueText) {

    nextRow();

    if (!obj.date) {
      _printRowSeparators();
      return;
    }

    if (obj.isTitle && obj.isDaily) {
      _printTitle(obj, valueText);
      if (obj.date.getTime() != arr[key - 1].date.getTime()) {
        nextRow();
      }
    }

    if (key === 0 || obj.date.getTime() != arr[key - 1].date.getTime()) {
      dataPrinter(obj);
      _printRowSeparators();
      if (!(obj.isTitle && obj.isDaily)) {
        nextRow();
      }
    }

    if (obj.isTitle && obj.isDaily) {
      return;
    }

    if (obj.isTitle) {
      _printTitle(obj, valueText);
    } else {
      eventPrinter(obj);
    }
    _printRowSeparators();
  }

  function _filter(arr) {
    var the_week = _getWeekBoards();
    return arr.filter(function (val, key) {
      if (!(val[5] == 2 || parseInt(val[5]) === 1 || val[5] == 6 || val[5] == 6.1)) {
        return false;
      }

      var _theDay = new Date(val[0]);
      return _theDay <= the_week.end && _theDay > the_week.start;
    });
  }

  function _sort(a, b) {
    if (a.date < b.date) {
      return -1;
    }
    if (a.date > b.date) {
      return 1;
    }
    if (a.start < b.start) {
      return -1;
    }
    if (a.start > b.start) {
      return 1;
    }
    return 0;
  }

  function _getWeekBoards() {
    var start = new Date(configurations[0]);
    var end   = new Date(configurations[1]);
    return { start: start, end: end };
  }

  /*functions*/
  function _printMainTitles(rowId, lang) {
    resultSheet.getRange(rowId, 1, 1, 6)
      .merge()
      .setBackground(MAIN_TITLES[lang].color)
      .setFontSize(36)
      .setFontWeight('bold')
      .setValue(MAIN_TITLES[lang].text)
      .setHorizontalAlignment('center');

    _separator(resultSheet.getRange(_resultCurrentRowId + 1, 1, 1, 6));
  }

  function _printTitle(obj, value) {
    var range = resultSheet.getRange(_resultCurrentRowId, 2, 1, 4);
    range.setBackground(obj.color).setFontSize(20).merge().setValue(value);
    range.setHorizontalAlignment('center');
    _printRowSeparators();
  }

  function _printEventHeb(e) {
    resultSheet.setRowHeight(_resultCurrentRowId, 24);
    resultSheet.getRange(_resultCurrentRowId, 3).setValue(e.heb.name).setFontWeight('bold');
    resultSheet.getRange(_resultCurrentRowId, 2).setValue(e.end + '-' + e.start).setFontWeight('bold');
    resultSheet.getRange(_resultCurrentRowId, 4).setValue(e.heb.manPlace);
    if (e.heb.womanPlace) {
      resultSheet.getRange(_resultCurrentRowId, 5).setValue(e.heb.womanPlace);
    } else {
      resultSheet.getRange(_resultCurrentRowId, 4, 1, 2).merge();
    }
  }

  function _printEventRus(e) {
    resultSheet.setRowHeight(_resultCurrentRowId, 24);
    resultSheet.getRange(_resultCurrentRowId, 4).setValue(e.rus.name).setFontWeight('bold');
    resultSheet.getRange(_resultCurrentRowId, 5).setValue(e.start + '-' + e.end).setFontWeight('bold');
    resultSheet.getRange(_resultCurrentRowId, 3).setValue(e.rus.manPlace);
    if (e.rus.womanPlace) {
      resultSheet.getRange(_resultCurrentRowId, 2).setValue(e.rus.womanPlace);
    } else {
      resultSheet.getRange(_resultCurrentRowId, 2, 1, 2).merge();
    }
  }

  function _printDateHeb(date) {
    clearRange(resultSheet.getRange(_resultCurrentRowId, 1, 1, 6));
    resultSheet.setRowHeight(_resultCurrentRowId, 28);
    _printDateByLang(0, date.date, '#cfe2f3');
    _printCellTitle([['שעה', 'פעילות', 'גברים', 'נשים']]);
  }

  function _printDateRus(date) {
    clearRange(resultSheet.getRange(_resultCurrentRowId, 1, 1, 6));
    resultSheet.setRowHeight(_resultCurrentRowId, 28);
    _printDateByLang(2, date.date, '#ffe599');
    _printCellTitle([['женщины', 'мужчины', 'мероприятия', 'время']]);
  }

  function _printCellTitle(values) {
    nextRow();
    resultSheet
      .getRange(_resultCurrentRowId, 2, 1, 4)
      .clear()
      .setBackground('#eeeeee')
      .setFontSize(12)
      .setFontWeight('bold')
      .setValues(values)
      .setHorizontalAlignment('center');
    _separator(resultSheet.getRange(_resultCurrentRowId, 6, 1, 1));
  }

  function _printDateByLang(langIndex, date, bgColor) {
    var a       = Utilities.formatDate(date, SpreadsheetApp.getActive().getSpreadsheetTimeZone(), 'dd/MM');
    var b       = _weekDays[date.getUTCDay()][langIndex];
    var dateStr = Utilities.formatString('\'%s %s', a, b);
    _separator(resultSheet.getRange(_resultCurrentRowId, 1, 1, 6));
    return resultSheet.getRange(_resultCurrentRowId, 2, 1, 4).merge().setBackground(bgColor).setFontSize(16).setFontWeight('bold').setValue(dateStr.toString()).setHorizontalAlignment('center');
  }

  function _parseEvent(row) {
    var _placeHeb = row[4].split('|@|');
    var _placeRus = row[10].split('|@|');
    return {
      heb: {
        name: row[3],
        manPlace: _placeHeb[0],
        womanPlace: _placeHeb[1]
      },
      rus: {
        name: row[11],
        manPlace: _placeRus[0],
        womanPlace: _placeRus[1]
      },
      start: row[1],
      end: row[2],
      date: row[0]
    };
  }

  function _parseTitle(row) {
    return {
      heb: row[3],
      rus: row[11],
      start: row[1],
      end: row[2],
      date: row[0],
      color: row[4],
      isTitle: true,
      isDaily: row[5] == 6
    };
  }

  function _printRowSeparators() {
    _separator(resultSheet.getRange(_resultCurrentRowId, 1));
    _separator(resultSheet.getRange(_resultCurrentRowId, 6));
  }

  function _separator(range) {
    range.clear().setBackground('#cccccc');
  }

  function clearRange(range) {
    range.clear();
    var _merged = range.getMergedRanges();
    if (_merged.length > 0) {
      range.getMergedRanges().breakApart();
    }
    return range;
  }

  function nextRow() {
    _resultCurrentRowId++;
    clearRange(resultSheet.getRange(_resultCurrentRowId, 1, 1, 6)).setFontSize(12).setHorizontalAlignment('center');
  }
}