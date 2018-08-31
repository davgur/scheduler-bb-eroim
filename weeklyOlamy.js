//var weekly = new LoadWeekly(1);
//weekly.run();

function LoadWeekly(_resultCurrentRowId) {
  var ss             = SpreadsheetApp.getActiveSpreadsheet();
  var resultSheet    = ss.getSheetByName('board');
  var configsSheet   = ss.getSheetByName('config');
  var _theDay        = {};
  var _weekDays      = configsSheet.getRange(3, 6, 7, 4).getValues();
  var _configOfTable = configsSheet.getRange(3, 14, 100, 1).getValues();

  var dataValues = SpreadsheetApp.openById('1LlRo5Ob5Bw8penUEakQj_1NyfmmD-T_Dama-k81dohQ').getSheetByName('Sheet1').getRange('A3:AG3000').getValues();
  dataValues     = _filter(dataValues);
  dataValues.push({ isLast: true });
  return {
    run: function (runAfterAll) {
      clearRange(resultSheet.getRange(1, 1, 1000, 17));
      _separator(resultSheet.getRange(_resultCurrentRowId, 1, 1, 17));
      _printTableTitle();
      dataValues.forEach(function (row, key) {
        if (row.isLast) {
          _printDay(_theDay);
          return;
        }
        if (_theDay.date && row[0] && _theDay.date.date.getTime() != row[0].getTime()) {
          _printDay(_theDay);
          _theDay = {};
        }
        if (!_theDay.date) {
          _theDay.date   = _parseDate(row);
          _theDay.events = [];
        }
        var _event = row[5] == 6 ? _parseTitle(row) : _parseEvent(row);
        _theDay.events.push(_event);
      });
      _separator(resultSheet.getRange(_resultCurrentRowId + 1, 1, 1, 17));
      runAfterAll(_resultCurrentRowId);
    }
  };

  function _filter(arr) {
    var the_week = _getToWeek();
    var _result  = arr.filter(function (val) {
      if (!(val[5] == 2 || val[5] == 3.1 || val[5] == 6)) {
        return false;
      }
      var _theDay = new Date(val[0]);
      return _theDay <= the_week.end && _theDay > the_week.start;
    });

    _result.sort(_equalTime);
    return _result;
  }

  function _getToWeek() {
    var _start, _end;
    var THE_DAY = 4;
    var _now    = new Date();
    var delta   = THE_DAY - _now.getDay();
    delta       = (delta < 0) ? 6 + delta : delta;
    _end        = new Date(_now.getYear(), _now.getMonth(), _now.getDate() + delta + parseInt(_configOfTable[0][0]));
    _start      = new Date(_now.getYear(), _now.getMonth(), _now.getDate() - 1);
    return { start: _start, end: _end };
  }

  /*functions*/
  function _printTableTitle() {
    _printDateByLang(2, 'לוח אירועים שבועי', '#6d9eeb').setFontSize(24).setFontWeight('bold');
    _printDateByLang(6, 'Weekly Events Board', '#00ff00').setFontSize(24).setFontWeight('bold');
    _printDateByLang(10, 'Расписание на неделю', '#ffff00').setFontSize(24).setFontWeight('bold');
    _printDateByLang(14, 'Lista de Eventos Semanales', '#e06666').setFontSize(24).setFontWeight('bold');
    resultSheet.setRowHeight(_resultCurrentRowId, 50);
  }

  function _printTitle(title) {
    resultSheet.getRange(_resultCurrentRowId, 1, 1, 50).clear();
    resultSheet.setRowHeight(_resultCurrentRowId, 30);

    //heb
    _printDateByLang(2, title.heb, title.color);
    //english
    _printDateByLang(6, title.eng, title.color);
    //rus
    _printDateByLang(10, title.rus, title.color);
    //esp
    _printDateByLang(14, title.esp, title.color);

    _separator(resultSheet.getRange(_resultCurrentRowId, 17, 1, 1));
  }

  function _printDay(day) {
    _printDate(day.date);
    day.events.forEach(function (e) {
      if (!e.start && !e.end) {
        return;
      }
      _resultCurrentRowId++;
      resultSheet.getRange(_resultCurrentRowId, 1, 1, 50).clear();
      if (e.isTitle) {
        _printTitle(e);
        return;
      }

      resultSheet.setRowHeight(_resultCurrentRowId, 24);

      _ptintHeb(e);
      _ptintEng(e);
      _ptintRus(e);
      _ptintEsp(e);
      _separator(resultSheet.getRange(_resultCurrentRowId, 17, 1, 1));
    });
  }

  function _printDate(data) {
    _resultCurrentRowId++;
    resultSheet.getRange(_resultCurrentRowId, 1, 1, 50).clear();
    resultSheet.setRowHeight(_resultCurrentRowId, 28);
    var dateStr = Utilities.formatDate(data.date, 'Asia/Jerusalem', 'dd/MM');
    //heb
    _printDateByLang(2, _weekDays[data.weekDay][0] + ' ' + dateStr, '#cfe2f3');
    //english
    _printDateByLang(6, _weekDays[data.weekDay][1] + ' ' + dateStr, '#b6d7a8');
    //rus
    _printDateByLang(10, _weekDays[data.weekDay][2] + ' ' + dateStr, '#ffe599');
    //esp
    _printDateByLang(14, _weekDays[data.weekDay][3] + ' ' + dateStr, '#f4cccc');
    _separator(resultSheet.getRange(_resultCurrentRowId, 17, 1, 1));
  }

  function _printDateByLang(colStartNum, textLang, bgColor) {
    var range = resultSheet.getRange(_resultCurrentRowId, colStartNum, 1, 1);
    range.setBackground(bgColor);
    range.setFontSize(16);
    range.setFontWeight('bold');
    range.setHorizontalAlignment('center');
    range.setValue(textLang);

    resultSheet.getRange(_resultCurrentRowId, colStartNum, 1, 3).merge();

    _separator(resultSheet.getRange(_resultCurrentRowId, colStartNum - 1, 1, 1));
    return range;
  }

  function _ptintHeb(e) {
    var timeStr = e.end + '-' + e.start;
    var textStr = e.heb;
    _printEvent(timeStr, textStr, 2, true);
  }

  function _ptintEng(e) {
    var timeStr = e.start + '-' + e.end;
    var textStr = e.eng;
    _printEvent(timeStr, textStr, 6);
  }

  function _ptintRus(e) {
    var timeStr = e.start + '-' + e.end;
    var textStr = e.rus;
    _printEvent(timeStr, textStr, 10);
  }

  function _ptintEsp(e) {
    var timeStr = e.start + '-' + e.end;
    var textStr = e.esp;
    _printEvent(timeStr, textStr, 14);
  }

  function _printEvent(timeStr, textStr, colStartNum, isHeb) {
    var textColNum = colStartNum;
    var timeColNum = colStartNum + 1;
    var timeDir    = 'left';

    if (isHeb) {
      timeDir    = 'right';
      textColNum = colStartNum + 2;
      timeColNum = colStartNum;
    }

    var rangeText = resultSheet.getRange(_resultCurrentRowId, textColNum);
    var rangeTime = resultSheet.getRange(_resultCurrentRowId, timeColNum, 1, 2);
    rangeTime.setHorizontalAlignment(timeDir);

    rangeTime.setValue(timeStr).merge();
    rangeTime.setFontWeight('bold');
    rangeTime.setFontSize(12);

    rangeText.setValue(textStr);
    rangeText.setFontSize(12);
    rangeText.setFontWeight('normal');

    _separator(resultSheet.getRange(_resultCurrentRowId, colStartNum - 1, 1, 1));
  }

  function _separator(range) {
    range.clear().setBackground('#cccccc');
  }

  function _parseDate(row) {
    var weekDay = Utilities.formatDate(row[0], 'Asia/Jerusalem', 'u') - 1;
    return {
      weekDay: weekDay,
      date: row[0],
      heb: row[3],
      eng: row[8],
      rus: row[11],
      esp: row[14]
    };
  }

  function _parseEvent(row) {
    return {
      heb: row[3],
      eng: row[8],
      rus: row[11],
      esp: row[14],
      start: row[1],
      end: row[2]
    };
  }

  function _parseTitle(row) {
    var result     = _parseEvent(row);
    result.color   = row[4];
    result.isTitle = true;
    return result;
  }

  function _equalTime(a, b) {
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
    if (a5Num < b5Num) {
      return -1;
    }
    if (a5Num > b5Num) {
      return 1;
    }
    return 0;
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