function runscript() {
  var builder = new Builder();
  builder.clean();
  builder.run();
}

function Builder() {

  var calendar = CalendarApp.getCalendarById('64pcduapg3t78237ffvh079t7c@group.calendar.google.com');

  var dataValues2018 = SpreadsheetApp.openById('1LlRo5Ob5Bw8penUEakQj_1NyfmmD-T_Dama-k81dohQ').getSheetByName('Sheet1').getRange('A3:AG3000').getValues();
  var dataValues2019 = SpreadsheetApp.openById('1HFI4KlEY72xfuEfzvYqYZ_gEKAAwiQ6aBRa3kWc44iY').getSheetByName('Sheet1').getRange('A2:O1500').getValues();

  return {
    clean: function () {
      var from   = new Date(2000, 1);
      var to     = new Date(2030, 1);
      var events = calendar.getEvents(from, to);
      events.forEach(function (e) {
        try {
          e.deleteEvent();
        } catch (e) {
        }
      });
    },
    run: function () {

      dataValues2018.filter(function (val) {
        return (val[5] == 2 || val[5] == 6);
      }).map(function (r) {
        return rowToObject(r);
      }).forEach(eventFromData);

      function eventFromData(data) {
        if (!data.date) {
          return;
        }
        var options = {
          'color': 'ORANGE',
          'description': 'description',
          'location': 'location',
        };
        try {
          calendar.createEvent(data.heb, data.start, data.end, options);
        } catch (e) {
        }
      }

      function rowToObject(row) {

        var date = dateByRow(row);
        return {
          start: date.start,
          end: date.end,
          date: dateByRow(row),
          heb: row[3],
          eng: row[8],
          rus: row[11],
          esp: row[14],
          colorId: colorByType(row[5])
        };
      }

      function dateByRow(row) {
        var start = new Date(row[0]);
        var end   = new Date(row[0]);

        var _startStr = row[1].split(':');
        start.setHours(_startStr[0], _startStr[1]);

        var _endStr = row[2].split(':');
        end.setHours(_endStr[0], _endStr[1]);

        return { start: start, end: end };
      }

      function colorByType(type) {
        switch (type) {
        case 2:
          return CalendarApp.EventColor.ORANGE;
        case 6:
          return CalendarApp.EventColor.GREEN;
        default:
          return CalendarApp.EventColor.RED;
        }
      }

    }
  };
}