function runscript() {
  var weekly = new BuilderFromToPT();
  weekly.run();
}

function BuilderFromToPT() {
  var _resultCurrentRowId = 1;
  var ss                  = SpreadsheetApp.getActiveSpreadsheet();
  var resultSheet         = ss.getSheetByName('fromTo');
  var titleValues         = ss.getSheetByName('config').getRange('D2:P100');
  var configurations      = ss.getSheetByName('config').getRange('B2:B100').getValues();
  var _weekDays           = SpreadsheetApp.openById('1B76zdIX2p48FEA1fvJr36DHKsQaIODQT9kWUZ8n0c7o').getSheetByName('config').getRange(3, 6, 7, 4).getValues();
  var MAIN_TITLES         = {
    heb: { color: '#6d9eeb', text: configurations[3] },
    rus: { color: '#ffff00', text: configurations[2] },
  };
  var dataValues          = SpreadsheetApp.openById('1LlRo5Ob5Bw8penUEakQj_1NyfmmD-T_Dama-k81dohQ').getSheetByName('Sheet1').getRange('A3:AG3000').getValues();
  dataValues              = _filter(dataValues);
  dataValues              = dataValues.map(function (row) {
    return _parseEvent(row);
  });
  dataValues              = _filterObjects(_parseTitles(titleValues)).concat(dataValues).sort(_sort);
  dataValues.push({});
  return {
    run: function () {
      clearRange(resultSheet.getRange(1, 1, 1000, 6));
      _printMainTitles(_resultCurrentRowId, 'heb');
      dataValues.forEach(function (obj, key, arr) {
        _resultCurrentRowId++;
        if (!obj.date) {
          _printRowSeparators();
          return;
        }
        if (key === 0 || obj.date.getTime() != arr[key - 1].date.getTime()) {
          _printDateHeb(obj);
          _printRowSeparators();
          _resultCurrentRowId++;
        }
        clearRange(resultSheet.getRange(_resultCurrentRowId, 1, 1, 6)).setFontSize(12).setHorizontalAlignment('center');
        if (obj.isTitle) {
          var range = resultSheet.getRange(_resultCurrentRowId, 2, 1, 4);
          obj.style.copyTo(range, { formatOnly: true });
          if (obj.isMarged) {
            range.merge().setValue(obj.heb);
            range.setHorizontalAlignment('center');
          } else {
            _printEventHeb(obj);
            range.setHorizontalAlignment('center');
          }
        } else {
          _printEventHeb(obj);
        }
        _printRowSeparators();
      });

      _printMainTitles(_resultCurrentRowId, 'rus');
      dataValues.forEach(function (obj, key, arr) {
        _resultCurrentRowId++;
        if (!obj.date) {
          _printRowSeparators();
          return;
        }
        if (key === 0 || obj.date.getTime() != arr[key - 1].date.getTime()) {
          _printDateRus(obj);
          _printRowSeparators();
          _resultCurrentRowId++;
        }
        clearRange(resultSheet.getRange(_resultCurrentRowId, 1, 1, 6)).setFontSize(12).setHorizontalAlignment('center');
        if (obj.isTitle) {
          var range = resultSheet.getRange(_resultCurrentRowId, 2, 1, 4);
          obj.style.copyTo(range, { formatOnly: true });
          if (obj.isMarged) {
            range.merge().setValue(obj.rus);
            range.setHorizontalAlignment('center');
          } else {
            _printEventRus(obj);
            range.setHorizontalAlignment('center');
          }
        } else {
          _printEventRus(obj);
        }
        _printRowSeparators();
      });
      clearRange(resultSheet.getRange(_resultCurrentRowId, 1, 50, 6));
      _separator(resultSheet.getRange(_resultCurrentRowId, 1, 1, 6));
    }
  };

  function _filter(arr) {
    var the_week = _getWeekBoards();
    return arr.filter(function (val, key) {
      if (val[5] != 2 && parseInt(val[5]) !== 1) {
        return false;
      }
      var _theDay = new Date(val[0]);

      return _theDay <= the_week.end && _theDay > the_week.start;
    });
  }

  function _filterObjects(arr) {
    var the_week = _getWeekBoards();
    return arr.filter(function (val, key) {
      return val.date <= the_week.end && val.date > the_week.start;
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
    var _values = range.getValues();
    var arr     = [];
    for (var i = 0; i < numRows; i++) {
      var _value = _values[i];
      var _data;
      if (_value[12].toLowerCase() !== 'v') {
        _data = _parseEvent(_value);
      } else {
        _data = {
          isMarged: true,
          heb: _value[3],
          rus: _value[11],
          start: _value[1],
          date: _value[0],
        };
      }
      _data.isTitle = true;
      _data.style   = range.getCell(i + 1, 1);
      arr.push(_data);
    }
    return arr.filter(function (t) {
      return t.date;
    });
  }

  function _getWeekBoards() {
    var start = new Date(configurations[0]);
    var end   = new Date(configurations[1]);
    return { start: start, end: end };
  }

  function _getWeekBoardsByWeekDay() {
    var _start, _end;
    var THE_DAY = 1;
    var _now    = new Date();
    var delta   = THE_DAY - _now.getDay();
    delta       = (delta < 0) ? 6 + delta : delta;

    _end   = new Date(_now.getYear(), _now.getMonth(), _now.getDate() + delta);
    _start = new Date(_now.getYear(), _now.getMonth(), _now.getDate() + delta - 7);
    return { start: _dateToNumber(_start), end: _dateToNumber(_end) };
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
    _resultCurrentRowId++;
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
    var range = resultSheet.getRange(_resultCurrentRowId, 2, 1, 4).merge().setBackground(bgColor).setFontSize(16).setFontWeight('bold').setValue(dateStr.toString()).setHorizontalAlignment('center');
    return range;
  }

  function _parseEvent(row) {
    if (row[4].search('_IS_JSON_') !== -1) {
      return _parseEventFromJSON(row);
    }
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

  function _parseEventFromJSON(row) {

    var _jsonHeb = JSON.parse(row[4].split('_IS_JSON_')[1]);
    var _jsonRus = JSON.parse(row[10].split('_IS_JSON_')[1]);
    return {
      heb: {
        name: row[3],
        manPlace: _jsonHeb.manPlace,
        womanPlace: _jsonHeb.womanPlace
      },
      rus: {
        name: row[11],
        manPlace: _jsonRus.manPlace,
        womanPlace: _jsonRus.womanPlace
      },
      titleParams: _jsonHeb.titleParams,
      start: row[1],
      end: row[2],
      date: row[0]
    };
  }

  function _dateToNumber(d) {
    return parseInt(Utilities.formatDate(new Date(d), 'EST', 'D'));
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
}