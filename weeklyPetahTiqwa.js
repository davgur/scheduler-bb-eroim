function runscript() {
  var weekly = new LoadWeeklyPT();
  weekly.run();
}

function LoadWeeklyPT() {
  var _resultCurrentRowId = 2;
  var ss                  = SpreadsheetApp.getActiveSpreadsheet();
  var resultSheet         = ss.getSheetByName('board');
  var titleValues         = ss.getSheetByName('titles').getRange('A2:F100');
  var _weekDays           = SpreadsheetApp.openById('1B76zdIX2p48FEA1fvJr36DHKsQaIODQT9kWUZ8n0c7o').getSheetByName('config').getRange(3, 6, 7, 4).getValues();

  var dataValues = SpreadsheetApp.openById('1LlRo5Ob5Bw8penUEakQj_1NyfmmD-T_Dama-k81dohQ').getSheetByName('Sheet1').getRange('A3:AG700').getValues();
  dataValues     = _filter(dataValues).map(function (row) {
    return _parseEvent(row);
  });
  dataValues     = _filter(_parseTitles(titleValues)).concat(dataValues).sort(_sort);

  return {
    run: function () {
      dataValues.forEach(function (obj, key, arr) {
        _resultCurrentRowId++;
        if (key === 0 || obj.date.getTime() != arr[key - 1].date.getTime()) {
          _printDate(obj);
          _printRowSeparators();
          _resultCurrentRowId++;
        }
        clearRange(resultSheet.getRange(_resultCurrentRowId, 1, 1, 20)).setFontSize(12).setHorizontalAlignment('center');
        if (obj.isTitle) {
          _printTitle(obj);
        } else {
          _printEvent(obj);
        }
        _printRowSeparators();
      });
      clearRange(resultSheet.getRange(_resultCurrentRowId, 1, 50, 20));
      _separator(resultSheet.getRange(_resultCurrentRowId, 1, 1, 11));
    }
  };

  function _filter(arr) {
    var the_week = _getWeekBoards();
    return arr.filter(function (val, key) {
      if (val[8] !== 2) {
        return false;
      }
      var _theDay = _dateToNumber(val[2]);
      if (_theDay > the_week.end || _theDay < the_week.start) {
        return false;
      }
      return true;
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

  function _parseTitles(range) {
    var numRows = range.getNumRows();
    var arr     = [];
    for (i = 1; i < numRows; i++) {
      var data = {
        isTitle: true,
        heb: range.getCell(i, 3).getValue(),
        rus: range.getCell(i, 4).getValue(),
        start: range.getCell(i, 2).getValue(),
        date: range.getCell(i, 1).getValue(),
        range: range.getCell(i, 3)
      };
      arr.push(data);

    }
    return arr.filter(function (t) {
      return t.date;
    });
  }

  function _getWeekBoards() {
    var _now   = new Date();
    //start from sunday
    var _start = new Date(_now.getYear(), _now.getMonth(), _now.getDate());
    var _end   = new Date(_now.getYear(), _now.getMonth(), _now.getDate() + (7 - _now.getDay()));
    return { start: _dateToNumber(_start), end: _dateToNumber(_end) };
  }

  /*functions*/
  function _printTitle(date) {
    resultSheet.setRowHeight(_resultCurrentRowId, 28);
    var hebRange = resultSheet.getRange(_resultCurrentRowId, 2, 1, 4);
    var rusRange = resultSheet.getRange(_resultCurrentRowId, 7, 1, 4);

    date.range.copyTo(hebRange, { formatOnly: true });
    date.range.copyTo(rusRange, { formatOnly: true });

    hebRange.merge().setValue(date.heb);
    rusRange.merge().setValue(date.rus);
  }

  function _printEvent(e) {
    resultSheet.setRowHeight(_resultCurrentRowId, 24);
    _ptintHeb(e);
    _ptintRus(e);
  }

  function _ptintHeb(e) {
    resultSheet.getRange(_resultCurrentRowId, 2).setValue(e.heb.name).setFontWeight('bold');
    resultSheet.getRange(_resultCurrentRowId, 3).setValue(e.start + '-' + e.end).setFontWeight('bold');
    resultSheet.getRange(_resultCurrentRowId, 4).setValue(e.heb.manPlace);
    if (e.heb.womanPlace) {
      resultSheet.getRange(_resultCurrentRowId, 5).setValue(e.heb.womanPlace);
    } else {
      resultSheet.getRange(_resultCurrentRowId, 4, 1, 2).merge();
    }
  }

  function _ptintRus(e) {
    resultSheet.getRange(_resultCurrentRowId, 10).setValue(e.rus.name).setFontWeight('bold');
    resultSheet.getRange(_resultCurrentRowId, 9).setValue(e.start + '-' + e.end).setFontWeight('bold');
    resultSheet.getRange(_resultCurrentRowId, 8).setValue(e.rus.manPlace);
    if (e.rus.womanPlace) {
      resultSheet.getRange(_resultCurrentRowId, 7).setValue(e.rus.womanPlace);
    } else {
      resultSheet.getRange(_resultCurrentRowId, 7, 1, 2).merge();
    }
  }

  function _printDate(date) {
    clearRange(resultSheet.getRange(_resultCurrentRowId, 1, 1, 20));
    resultSheet.setRowHeight(_resultCurrentRowId, 28);

    _printDateByLang(2, 0, date.date, '#cfe2f3');
    _printDateByLang(7, 2, date.date, '#ffe599');
  }

  function _printDateByLang(colStartNum, langIndex, date, bgColor) {
    var dateStr = Utilities.formatString('%s %s', Utilities.formatDate(date, SpreadsheetApp.getActive().getSpreadsheetTimeZone(), 'dd/MM'), _weekDays[date.getDay()][langIndex]);
    var range   = resultSheet.getRange(_resultCurrentRowId, colStartNum, 1, 4).merge().setBackground(bgColor).setFontSize(16).setFontWeight('bold').setValue(dateStr.toString()).setHorizontalAlignment('center');
    _separator(resultSheet.getRange(_resultCurrentRowId, colStartNum - 1, 1, 1));
    return range;
  }

  function _parseEvent(row) {
    var _placeHeb = row[7].split('|@|');
    var _placeRus = row[18].split('|@|');
    return {
      heb: {
        name: row[4],
        manPlace: _placeHeb[0],
        womanPlace: _placeHeb[1]
      },
      rus: {
        name: row[21],
        manPlace: _placeRus[0],
        womanPlace: _placeRus[1]
      },
      start: row[5],
      end: row[6],
      date: row[2]
    };
  }

  function _dateToNumber(d) {
    return parseInt(Utilities.formatDate(new Date(d), 'EST', 'D'));
  }

  function _printRowSeparators() {
    _separator(resultSheet.getRange(_resultCurrentRowId, 1));
    _separator(resultSheet.getRange(_resultCurrentRowId, 6));
    _separator(resultSheet.getRange(_resultCurrentRowId, 11));
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
}