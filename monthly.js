//var monthly = new LoadMonthly(1);
//monthly.run();

function LoadMonthly(_resultCurrentRowId) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var resultSheet = ss.getSheetByName('Sheet1');
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
                var __c = !!arr[key + 1] && arr[key + 1][8] !== 0;
                var __b = arr[key + 1][8];
                var __d = arr[key + 1];
                var __dd = arr[key];
                if (__c) {
                    var __cc = 1;
                }
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
        dataValues.forEach(function (row, i) {
            _resultCurrentRowId++;
            if (row[8] === 0) {
                //heb
                _formatMonth(resultSheet.getRange(_resultCurrentRowId, 2, 1, 2), row.slice(1, 3), "#cfe2f3");
                //eng
                _formatMonth(resultSheet.getRange(_resultCurrentRowId, 5, 1, 2), row.slice(10, 12), "#b6d7a8");
                //rus
                _formatMonth(resultSheet.getRange(_resultCurrentRowId, 8, 1, 2), row.slice(18, 20), "#ffe599");
                //esp
                _formatMonth(resultSheet.getRange(_resultCurrentRowId, 11, 1, 2), row.slice(26, 28), "#f4cccc");

            } else {
                //heb
                _addContent(resultSheet.getRange(_resultCurrentRowId, 2, 1, 2), row.slice(2, 5), true);
                //english
                _addContent(resultSheet.getRange(_resultCurrentRowId, 5, 1, 2), row.slice(13, 16));
                //rus
                _addContent(resultSheet.getRange(_resultCurrentRowId, 8, 1, 2), row.slice(21, 24));
                //esp
                _addContent(resultSheet.getRange(_resultCurrentRowId, 11, 1, 2), row.slice(29, 32));

            }

            _toBlack(resultSheet.getRange(_resultCurrentRowId, 1));
            _toBlack(resultSheet.getRange(_resultCurrentRowId, 4));
            _toBlack(resultSheet.getRange(_resultCurrentRowId, 7));
            _toBlack(resultSheet.getRange(_resultCurrentRowId, 10));
            _toBlack(resultSheet.getRange(_resultCurrentRowId, 13));


        });
        resultSheet.getRange(_resultCurrentRowId + 1, 1, 50, 50).clear();
        _toBlack(resultSheet.getRange(_resultCurrentRowId + 1, 1, 1, 13));
        _defineAutoResize([2, 3, 5, 6, 8, 9, 11, 12]);
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
    }

    function _toBlack(range) {
        range.setBackground("#cccccc");
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
        if (isHeb) {
            if (!values[0]) {
                return;
            }
            _values[0] = values[1] + ' ' + Utilities.formatDate(values[0], Session.getScriptTimeZone(), "dd/MM/yyyy");
            _values[1] = values[2];
        } else {

            if (!values[2]) {
                return;
            }
            _values[0] = values[1] + ' ' + Utilities.formatDate(values[2], Session.getScriptTimeZone(), "dd/MM/yyyy");
            _values[1] = values[0];
        }
        range.setValues([_values]);
    }
}