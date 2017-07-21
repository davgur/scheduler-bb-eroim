//var monthly = new LoadMonthly(1);
//monthly.run();

function LoadMonthly(_resultCurrentRowId) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var resultSheet = ss.getSheetByName('board');
    var resultValues = resultSheet.getDataRange().getValues();
    var _theDay = {};
    _resultCurrentRowId++;

    var dataValues = SpreadsheetApp.openById("1LlRo5Ob5Bw8penUEakQj_1NyfmmD-T_Dama-k81dohQ").getSheetByName("Sheet1").getRange('A3:AG1000').getValues();
    dataValues = _filter(dataValues);

    var MONTH_NAMES = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('config').getRange('A1:D11').getValues();

    return {
        run: run
    };

    function _filter(arr) {
        var result = {yearsList: []}, data = [];
        var _now = new Date();
        arr.filter(function (val, key) {
            if (val[8] !== 3) {
                return false;
            }
            if (val[2] < _now) {
                return false;
            }
            return true;
        }).sort(function (a, b) {
            if (a[2] < b[2]) {
                return -1;
            }
            if (a[2] > b[2]) {
                return 1;
            }
            if (a[5] < b[5]) {
                return -1;
            }
            if (a[5] > b[5]) {
                return 1;
            }
            return 0;
        }).forEach(function (val) {
            var key = new Date(val[2]).getFullYear();
            if (!result[key]) {
                result[key] = [];
                result.yearsList.push(key);
            }
            result[key].push(val);

        });
        return result;
    }

    function _defineAutoResize(colsId) {
        colsId.forEach(function (i) {
            resultSheet.autoResizeColumn(i);
        });
    }

    function run() {
        dataValues.yearsList.forEach(function (year) {
            printYear(year, dataValues[year]);
        });
        resultSheet.getRange(_resultCurrentRowId + 1, 1, 50, 50).clear();
        //_defineAutoResize([4,6, 10, 14]);
    }

    function printYear(year, data) {
        _printTableTitle(year);
        var _theMonth;
        data.forEach(function (row, i) {
            _resultCurrentRowId++;
            if (_theMonth !== row[2].getMonth()) {
                _theMonth = row[2].getMonth();
                resultSheet.setRowHeight(_resultCurrentRowId, 30);
                //heb
                _formatMonth(resultSheet.getRange(_resultCurrentRowId, 2, 1, 3), MONTH_NAMES[_theMonth + 1][1], "#cfe2f3");
                //eng
                _formatMonth(resultSheet.getRange(_resultCurrentRowId, 6, 1, 3), MONTH_NAMES[_theMonth + 1][2], "#b6d7a8");
                //rus
                _formatMonth(resultSheet.getRange(_resultCurrentRowId, 10, 1, 3), MONTH_NAMES[_theMonth + 1][3], "#ffe599");
                //esp
                _formatMonth(resultSheet.getRange(_resultCurrentRowId, 14, 1, 3), MONTH_NAMES[_theMonth + 1][4], "#f4cccc");
                _resultCurrentRowId++;
            }

            resultSheet.setRowHeight(_resultCurrentRowId, 24);
            //heb
            _addContent(resultSheet.getRange(_resultCurrentRowId, 2, 1, 3), row.slice(2, 5), true);
            resultSheet.getRange(_resultCurrentRowId, 2, 1, 2).setFontWeight("bold");
            //english
            _addContent(resultSheet.getRange(_resultCurrentRowId, 6, 1, 3), row.slice(13, 16));
            resultSheet.getRange(_resultCurrentRowId, 7, 1, 2).setFontWeight("bold");
            //rus
            _addContent(resultSheet.getRange(_resultCurrentRowId, 10, 1, 3), row.slice(21, 24));
            resultSheet.getRange(_resultCurrentRowId, 11, 1, 2).setFontWeight("bold");
            //esp
            _addContent(resultSheet.getRange(_resultCurrentRowId, 14, 1, 3), row.slice(29, 32));
            resultSheet.getRange(_resultCurrentRowId, 15, 1, 2).setFontWeight("bold");


            _toBlack(resultSheet.getRange(_resultCurrentRowId, 1));
            _toBlack(resultSheet.getRange(_resultCurrentRowId, 5));
            _toBlack(resultSheet.getRange(_resultCurrentRowId, 9));
            _toBlack(resultSheet.getRange(_resultCurrentRowId, 13));
            _toBlack(resultSheet.getRange(_resultCurrentRowId, 17));
        });
        _resultCurrentRowId++;
        _toBlack(resultSheet.getRange(_resultCurrentRowId, 1, 1, 17));
    }


    function _printTableTitle(year) {
        _resultCurrentRowId++;
        _formatMonth(resultSheet.getRange(_resultCurrentRowId, 2, 1, 3), ['לוח אירועים שנתי ' + year, '', ''], "#6d9eeb").setFontSize(24).setFontWeight("bold");
        _formatMonth(resultSheet.getRange(_resultCurrentRowId, 6, 1, 3), ["Annual Board of Events " + year, '', ''], "#00ff00").setFontSize(24).setFontWeight("bold");
        _formatMonth(resultSheet.getRange(_resultCurrentRowId, 10, 1, 3), ["Расписание событий на " + year + " год", '', ''], "#ffff00").setFontSize(24).setFontWeight("bold");
        _formatMonth(resultSheet.getRange(_resultCurrentRowId, 14, 1, 3), ["Tabla Anual de Eventos Para " + year, '', ''], "#e06666").setFontSize(24).setFontWeight("bold");
        resultSheet.setRowHeight(_resultCurrentRowId, 50);

        _toBlack(resultSheet.getRange(_resultCurrentRowId, 1));
        _toBlack(resultSheet.getRange(_resultCurrentRowId, 5));
        _toBlack(resultSheet.getRange(_resultCurrentRowId, 9));
        _toBlack(resultSheet.getRange(_resultCurrentRowId, 13));
        _toBlack(resultSheet.getRange(_resultCurrentRowId, 17));
        _toBlack(resultSheet.getRange(_resultCurrentRowId + 1, 1, 1, 17));
        _resultCurrentRowId++;
    }

    function _formatMonth(range, value, bgColor) {
        range.merge()
            .setValue(value)
            .setFontSize(16)
            .setFontWeight("bold")
            .setHorizontalAlignment("center")
            .setBackground(bgColor);
        return range;
    }

    function _toBlack(range) {
        range.clear().setBackground("#cccccc");
    }

    function _clearFormat(range, isHeb) {
        range.clearFormat();
        var dir = isHeb ? "right" : "left";
        range.setHorizontalAlignment(dir);
    }

    function _toCenter(range) {
        range.setHorizontalAlignment("center");
    }

    function _addContent(range, values, isHeb) {
        var _dateStr, _time, _values = [];
        _clearFormat(range, isHeb);

        if (values.length === 0) {
            return;
        }
        range.setFontSize(12);
        if (isHeb) {
            if (!values[0]) {
                return;
            }
            _values[0] = Utilities.formatDate(new Date(values[0]), "GMT+0200", "dd/MM/yyyy");
            _values[1] = values[1];
            _values[2] = values[2];
        } else {
            if (!values[2]) {
                return;
            }

            _values[2] = Utilities.formatDate(new Date(values[2]), "GMT+0200", "dd/MM/yyyy");
            _values[1] = values[1];
            _values[0] = values[0];
        }
        range.setValues([_values]);
    }

    function _buildArray(from, to) {
        var _arr = [];
        for (var i = from; i <= to; i++) {
            (function (j) {
                _arr.push(j);
            }(i));
        }
        return _arr;
    }
}