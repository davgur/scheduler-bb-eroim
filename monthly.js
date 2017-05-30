//var monthly = new LoadMonthly(1);
//monthly.run();

function LoadMonthly(_resultCurrentRowId) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var resultSheet = ss.getSheetByName('board');
    var resultValues = resultSheet.getDataRange().getValues();
    var _theDay = {};
    _resultCurrentRowId++;

    var dataValues = SpreadsheetApp.openById("1LlRo5Ob5Bw8penUEakQj_1NyfmmD-T_Dama-k81dohQ").getSheetByName("Sheet1").getRange('A3:AG500').getValues();
    dataValues = _filter(dataValues);

    return {
        run: run
    }

    function _filter(arr) {
        var _result = [];

        _result = arr.filter(function (val, key) {
            if (val[8] === 0 || (val[8] === 3 && val[2] > new Date())) {
                return true;
            }
            return false;
        });
        _result = _result.filter(function (val, key, arr) {
            if (val[8] === 0) {
                return !!arr[key + 1] && arr[key + 1][8] !== 0;
            }
            return !!val[8];
        });
        return _result;
    }

    function _defineAutoResize(colsId) {
        colsId.forEach(function (i) {
            resultSheet.autoResizeColumn(i);
        });
    }

    function run() {
        _printTableTitle()
        dataValues.forEach(function (row, i) {
            _resultCurrentRowId++;
            if (row[8] === 0) {
                //heb
                _formatMonth(resultSheet.getRange(_resultCurrentRowId, 2, 1, 3), row.slice(1, 4), "#cfe2f3");
                //eng
                _formatMonth(resultSheet.getRange(_resultCurrentRowId, 6, 1, 3), row.slice(10, 13), "#b6d7a8");
                //rus
                _formatMonth(resultSheet.getRange(_resultCurrentRowId, 10, 1, 3), row.slice(18, 21), "#ffe599");
                //esp
                _formatMonth(resultSheet.getRange(_resultCurrentRowId, 14, 1, 3), row.slice(26, 29), "#f4cccc");

            } else {
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

            }

            _toBlack(resultSheet.getRange(_resultCurrentRowId, 1));
            _toBlack(resultSheet.getRange(_resultCurrentRowId, 5));
            _toBlack(resultSheet.getRange(_resultCurrentRowId, 9));
            _toBlack(resultSheet.getRange(_resultCurrentRowId, 13));
            _toBlack(resultSheet.getRange(_resultCurrentRowId, 17));


        });
        resultSheet.getRange(_resultCurrentRowId + 1, 1, 50, 50).clear();
        _toBlack(resultSheet.getRange(_resultCurrentRowId + 1, 1, 1, 17));
        //_defineAutoResize([4,6, 10, 14]);
    }


    function _printTableTitle() {
        _resultCurrentRowId++;
        _formatMonth(resultSheet.getRange(_resultCurrentRowId, 2, 1, 3), ['לוח אירועים שנתי 2017', '', ''], "#6d9eeb").setFontSize(36).setFontWeight("normal");
        _formatMonth(resultSheet.getRange(_resultCurrentRowId, 6, 1, 3), ["Annual Board of Events 2017", '', ''], "#00ff00").setFontSize(36).setFontWeight("normal");
        _formatMonth(resultSheet.getRange(_resultCurrentRowId, 10, 1, 3), ["Расписание событий на 2017 год", '', ''], "#ffff00").setFontSize(36).setFontWeight("normal");
        _formatMonth(resultSheet.getRange(_resultCurrentRowId, 14, 1, 3), ["TABLA ANUAL DE EVENTOS PARA 2017", '', ''], "#e06666").setFontSize(36).setFontWeight("normal");
        resultSheet.setRowHeight(_resultCurrentRowId, 50);

        _toBlack(resultSheet.getRange(_resultCurrentRowId + 1, 1, 1, 17));
        _resultCurrentRowId++;
    }

    function _formatMonth(range, values, bgColor) {
        if (values.length > 0) {
            range.setValues([values]);
        }
        range.merge();
        range.setFontSize(16);
        range.setFontWeight("bold");
        range.setHorizontalAlignment("center");
        range.setBackground(bgColor);
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
            _values[0] = Utilities.formatDate(values[0], Session.getScriptTimeZone(), "dd/MM/yyyy");
            _values[1] = values[1];
            _values[2] = values[2];
        } else {
            if (!values[2]) {
                return;
            }
            _values[2] = Utilities.formatDate(values[2], Session.getScriptTimeZone(), "dd/MM/yyyy");
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