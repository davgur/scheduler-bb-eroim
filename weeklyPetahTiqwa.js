function runscript() {
    var weekly = new LoadWeeklyPT();
    weekly.run();
}

function LoadWeeklyPT() {
    var _resultCurrentRowId = 1;
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var resultSheet = ss.getSheetByName('board');
    var titleValues = ss.getSheetByName('titles').getRange('A2:F100');
    var _weekDays = SpreadsheetApp.openById('1B76zdIX2p48FEA1fvJr36DHKsQaIODQT9kWUZ8n0c7o').getSheetByName('config').getRange(3, 6, 7, 4).getValues();
    var MAIN_TITLES = {
        heb: {color: '#6d9eeb', text: 'לוח אירועים שבועי - פתח תקווה'},
        rus: {color: '#ffff00', text: 'Расписание на неделю'},
    };
    var dataValues = SpreadsheetApp.openById('1LlRo5Ob5Bw8penUEakQj_1NyfmmD-T_Dama-k81dohQ').getSheetByName('Sheet1').getRange('A3:AG1500').getValues();
    dataValues = _filter(dataValues);
    dataValues = dataValues.map(function (row) {
        return _parseEvent(row);
    });
    dataValues = _filterObjects(_parseTitles(titleValues)).concat(dataValues).sort(_sort);
    return {
        run: function () {
            _printMainTitles(_resultCurrentRowId, 'heb');
            dataValues.forEach(function (obj, key, arr) {
                _resultCurrentRowId++;
                if (key === 0 || obj.date.getTime() != arr[key - 1].date.getTime()) {
                    _printDateHeb(obj);
                    _printRowSeparators();
                    _resultCurrentRowId++;
                }
                clearRange(resultSheet.getRange(_resultCurrentRowId, 1, 1, 12)).setFontSize(12).setHorizontalAlignment('center');
                if (obj.isTitle) {
                    _printTitleHeb(obj);
                } else {
                    _printEventHeb(obj);
                }
                _printRowSeparators();
            });

            _printMainTitles(_resultCurrentRowId, 'rus');
            dataValues.forEach(function (obj, key, arr) {
                _resultCurrentRowId++;
                if (key === 0 || obj.date.getTime() != arr[key - 1].date.getTime()) {
                    _printDateRus(obj);
                    _printRowSeparators();
                    _resultCurrentRowId++;
                }
                clearRange(resultSheet.getRange(_resultCurrentRowId, 1, 1, 12)).setFontSize(12).setHorizontalAlignment('center');
                if (obj.isTitle) {
                    _printTitleRus(obj);
                } else {
                    _printEventRus(obj);
                }
                _printRowSeparators();
            });
            clearRange(resultSheet.getRange(_resultCurrentRowId, 1, 50, 12));
            _separator(resultSheet.getRange(_resultCurrentRowId, 1, 1, 6));
        }
    };

    function _filter(arr) {
        var the_week = _getWeekBoards();
        return arr.filter(function (val, key) {
            if (val[5] != 2 && val[5] !== new Number(1)) {
                return false;
            }
            var _theDay = _dateToNumber(val[0]);

            if (the_week.start > the_week.end) {
                return _theDay < the_week.end || _theDay > the_week.start;
            }
            return _theDay < the_week.end && _theDay > the_week.start;
        });
    }

    function _filterObjects(arr) {
        var the_week = _getWeekBoards();
        return arr.filter(function (val, key) {
            var _theDay = _dateToNumber(val.date);

            if (the_week.start > the_week.end) {
                return _theDay < the_week.end || _theDay > the_week.start;
            }
            return _theDay < the_week.end && _theDay > the_week.start;
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
        var arr = [];
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
        var _now = new Date();
        _end = new Date(_now.getYear(), _now.getMonth(), _now.getDate() + 7);
        return {start: _dateToNumber(_now), end: _dateToNumber(_end)};
    }

    function _getWeekBoardsByWeekDay() {
        var _start, _end;
        var THE_DAY = 1;
        var _now = new Date();
        var delta = THE_DAY - _now.getDay();
        delta = (delta < 0) ? 6 + delta : delta;

        _end = new Date(_now.getYear(), _now.getMonth(), _now.getDate() + delta);
        _start = new Date(_now.getYear(), _now.getMonth(), _now.getDate() + delta - 7);
        return {start: _dateToNumber(_start), end: _dateToNumber(_end)};
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

    function _printTitleHeb(date) {
        resultSheet.setRowHeight(_resultCurrentRowId, 28);
        var range = resultSheet.getRange(_resultCurrentRowId, 2, 1, 4);
        date.range.copyTo(range, {formatOnly: true});
        range.merge().setValue(date.heb);
    }

    function _printTitleRus(date) {
        resultSheet.setRowHeight(_resultCurrentRowId, 28);
        var range = resultSheet.getRange(_resultCurrentRowId, 2, 1, 4);
        date.range.copyTo(range, {formatOnly: true});
        range.merge().setValue(date.rus);
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
        clearRange(resultSheet.getRange(_resultCurrentRowId, 1, 1, 10));
        resultSheet.setRowHeight(_resultCurrentRowId, 28);
        _printDateByLang(0, date.date, '#cfe2f3');
        _printCellTitle([['שעה', 'פעילות', 'גברים', 'נשים']]);
    }

    function _printDateRus(date) {
        clearRange(resultSheet.getRange(_resultCurrentRowId, 1, 1, 10));
        resultSheet.setRowHeight(_resultCurrentRowId, 28);
        _printDateByLang(2, date.date, '#ffe599');
        _printCellTitle([['женщины', 'мужчины', 'мероприятия', 'время']]);
    }

    function _printCellTitle(values) {
        _resultCurrentRowId++;
        resultSheet
            .getRange(_resultCurrentRowId, 2, 1, 4)
            .clear()
            .setBackground("#eeeeee")
            .setFontSize(12)
            .setFontWeight('bold')
            .setValues(values)
            .setHorizontalAlignment('center');
        _separator(resultSheet.getRange(_resultCurrentRowId, 6, 1, 1));
    }

    function _printDateByLang(langIndex, date, bgColor) {
        var a = Utilities.formatDate(date, SpreadsheetApp.getActive().getSpreadsheetTimeZone(), 'dd/MM');
        var b = _weekDays[date.getUTCDay()][langIndex];
        var dateStr = Utilities.formatString('\'%s %s', a, b);
        _separator(resultSheet.getRange(_resultCurrentRowId, 1, 1, 6));
        var range = resultSheet.getRange(_resultCurrentRowId, 2, 1, 4).merge().setBackground(bgColor).setFontSize(16).setFontWeight('bold').setValue(dateStr.toString()).setHorizontalAlignment('center');
        return range;
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