//var monthly = new LoadMonthly(2);
//monthly.run();

function LoadMonthly(_resultCurrentRowId) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var resultSheet = ss.getSheetByName('Sheet1');
    var resultValues = resultSheet.getDataRange().getValues();
    var _theDay = {};
    _resultCurrentRowId++;

    var dataValues = SpreadsheetApp.openById("1LlRo5Ob5Bw8penUEakQj_1NyfmmD-T_Dama-k81dohQ").getSheetByName("test").getRange('A3:AG500').getValues();
    dataValues = _filter(dataValues);
   _defineAutoResize([2,3,4,8,9,10,13,14,15,18,19,20,22,23,24]);
  
  return {
      run: run
    }
    
    function _filter(arr) {
      var _result = [];
      
      _result = arr.filter(function (val, key) {
          if (val[8] === 0 || val[8] === 3) {
              return true;
            }
          
          return false;
        });
      return _result;
    }
  function _defineAutoResize(colsId){
    colsId.forEach(function(i){
      resultSheet.autoResizeColumn(i);
    });
  }
  
  function run() {
        for (var row = 0; row < 10; row++) {
            _resultCurrentRowId++;
            var __bb = dataValues[row];
            if (dataValues[row][8] === 0) {              
                //heb
                _formatMonth(resultSheet.getRange(_resultCurrentRowId, 2, 1, 3), dataValues[row].slice(1, 4), "#cfe2f3");
                //eng
                _formatMonth(resultSheet.getRange(_resultCurrentRowId, 6, 1, 3), dataValues[row].slice(10, 13), "#b6d7a8");
                //rus
                _formatMonth(resultSheet.getRange(_resultCurrentRowId, 10, 1, 3), dataValues[row].slice(18, 21), "#ffe599");
                //esp
                _formatMonth(resultSheet.getRange(_resultCurrentRowId, 14, 1, 3), dataValues[row].slice(26, 29), "#f4cccc");

            } else {
                //heb
                _addContent(resultSheet.getRange(_resultCurrentRowId, 2, 1, 3), dataValues[row].slice(2, 7), true);
                _toCenter(resultSheet.getRange(_resultCurrentRowId, 4));
                _toCenter(resultSheet.getRange(_resultCurrentRowId, 6, 1, 2));
                //english
                _addContent(resultSheet.getRange(_resultCurrentRowId, 6, 1, 3), dataValues[row].slice(11, 16));
                _toCenter(resultSheet.getRange(_resultCurrentRowId, 12, 1, 2));
                _toCenter(resultSheet.getRange(_resultCurrentRowId, 16));
                //rus
                _addContent(resultSheet.getRange(_resultCurrentRowId, 10, 1, 3), dataValues[row].slice(19, 24));
                _toCenter(resultSheet.getRange(_resultCurrentRowId, 20, 1, 2));
                _toCenter(resultSheet.getRange(_resultCurrentRowId, 24));
                //esp
                _addContent(resultSheet.getRange(_resultCurrentRowId, 14, 1, 3), dataValues[row].slice(27, 32));
                _toCenter(resultSheet.getRange(row, 28, 1, 2));
                _toCenter(resultSheet.getRange(row, 32));

            }

            _toBlack(resultSheet.getRange(_resultCurrentRowId, 1));
            _toBlack(resultSheet.getRange(_resultCurrentRowId, 5));
            _toBlack(resultSheet.getRange(_resultCurrentRowId, 9));
            _toBlack(resultSheet.getRange(_resultCurrentRowId, 13));
            _toBlack(resultSheet.getRange(_resultCurrentRowId, 17));          
        }
        resultSheet.getRange(_resultCurrentRowId + 1, 1, 50, 50).clear();
        _toBlack(resultSheet.getRange(_resultCurrentRowId + 1, 1, 1, 17));
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
      var _dateStr,_time, _values = [];
        _clearFormat(range, isHeb);
      
        if (values.length === 0) {
            return;
        }
        if (isHeb) {
           _values[0] =  Utilities.formatDate(values[0],  Session.getScriptTimeZone(), "dd/MM/yyyy") + ' ' + values[1];
          _values[1] = values[2];
          _values[2] = values[3] + " - " + values[4];
            if (_values[2] == " - ") {
                _values[2] = "";
            }
        } else {            
           _values[0] = values[3] + ' ' + Utilities.formatDate(values[4],  Session.getScriptTimeZone(), "dd/MM/yyyy");
          _values[1] = values[2];
          _values[2] = values[1] + " - " + values[0];
            if (_values[2] == " - ") {
                _values[2] = "";
            }
        }
        range.setValues([_values]);
    }
}