var weekly = new LoadWeeklyPT();
weekly.run();

function LoadWeeklyPT() {
    var _resultCurrentRowId = 2;
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var resultSheet = ss.getSheetByName('board');
    var resultValues = resultSheet.getDataRange().getValues();

    var dataValues = SpreadsheetApp.openById("1LlRo5Ob5Bw8penUEakQj_1NyfmmD-T_Dama-k81dohQ").getSheetByName("Sheet1").getRange('A3:AG700').getValues();
    dataValues = _filter(dataValues).map(function (row) {
        return _parseEvent(row);
    });
    dataValues = _sort(dataValues);
    dataValues = _insertTitles(dataValues);
    dataValues = _sort(dataValues);

    return {
        run: function () {
            resultSheet.getRange(3, 1, 200, 15).clear().setFontSize(12);
            dataValues.forEach(function (obj, key, arr) {
                if (key === 0 || obj.date !== arr[key - 1].date) {
                    _printDay(obj);
                }
                if (obj.isTitle) {
                    _parseTitle(obj);
                } else {
                    _parseEvent(obj);
                }
            });
            _separator(resultSheet.getRange(_resultCurrentRowId, 1, 1, 11));
        }
    }

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
        return _result;
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
    function _insertTitles(arr) {

    }

    function _sort(arr) {
        arr.sort(function (a, b) {
            if (a[2] < b[2]) {
                return -1;
            }
            if (a[2] > b[2]) {
                return 1;
            }
            return 0;
        });
        return arr;
    }

    function _getWeekBoards() {
        var _now = new Date();
        //start from sunday
        var _start = new Date(_now.getYear(), _now.getMonth(), _now.getDate());
        var _end = new Date(_now.getYear(), _now.getMonth(), _now.getDate() + (7 - _now.getDay()));
        return {start: _dateToNumber(_start), end: _dateToNumber(_end)};
    }


    /*functions*/

    function _printDay(day) {
        _printDate(day.date);
        day.events.forEach(function (e) {
            if (!e.start && !e.end) {
                return;
            }
            _resultCurrentRowId++;
            resultSheet.setRowHeight(_resultCurrentRowId, 24);

            _ptintHeb(e);
            _ptintRus(e);
            _printRowSeparators()
        });
    }

    function _printDate(date) {
        _resultCurrentRowId++;
        resultSheet.getRange(_resultCurrentRowId, 1, 1, 50).clear();
        resultSheet.setRowHeight(_resultCurrentRowId, 28);
        var dateStr = Utilities.formatDate(date.date, SpreadsheetApp.getActive().getSpreadsheetTimeZone(), "dd/MM");

        _printDateByLang(2, date.heb + " " + dateStr, "#cfe2f3");
        _printDateByLang(7, date.rus + " " + dateStr, "#ffe599");
        _printRowSeparators();
    }


    function _printDateByLang(colStartNum, textLang, bgColor) {
        var range = resultSheet.getRange(_resultCurrentRowId, colStartNum, 1, 1);
        range.setBackground(bgColor).setFontSize(16).setFontWeight("bold").setHorizontalAlignment("center").setValue(textLang);
        resultSheet.getRange(_resultCurrentRowId, colStartNum, 1, 4).merge();
        _separator(resultSheet.getRange(_resultCurrentRowId, colStartNum - 1, 1, 1));
        return range;
    }


    function _ptintHeb(e) {
        resultSheet.getRange(_resultCurrentRowId, 2).setValue(e.heb.name);
        resultSheet.getRange(_resultCurrentRowId, 3).setValue(e.start + "-" + e.end);
        resultSheet.getRange(_resultCurrentRowId, 4).setValue(e.heb.manPlace);
        if (e.heb.womanPlace) {
            resultSheet.getRange(_resultCurrentRowId, 5).setValue(e.heb.womanPlace);
        } else {
            resultSheet.getRange(_resultCurrentRowId, 4, 1, 2).merge().setHorizontalAlignment("center");
        }
    }

    function _ptintRus(e) {
        resultSheet.getRange(_resultCurrentRowId, 10).setValue(e.heb.name);
        resultSheet.getRange(_resultCurrentRowId, 9).setValue(e.start + "-" + e.end);
        resultSheet.getRange(_resultCurrentRowId, 8).setValue(e.heb.manPlace);
        if (e.heb.womanPlace) {
            resultSheet.getRange(_resultCurrentRowId, 7).setValue(e.heb.womanPlace);
        } else {
            resultSheet.getRange(_resultCurrentRowId, 7, 1, 2).merge().setHorizontalAlignment("center");
        }
    }


    function _parseDate(row) {
        return {
            date: row[2],
            heb: row[3],
            rus: row[22]
        };
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
            end: row[6]
        };
    }

    function _dateToNumber(d) {
        return parseInt(Utilities.formatDate(new Date(d), "EST", "D"));
    }

    function _printRowSeparators() {
        _separator(resultSheet.getRange(_resultCurrentRowId, 1));
        _separator(resultSheet.getRange(_resultCurrentRowId, 6));
        _separator(resultSheet.getRange(_resultCurrentRowId, 11));
    }

    function _separator(range) {
        range.clear().setBackground("#cccccc");
    }
}