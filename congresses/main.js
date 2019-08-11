var currentRowByLang, configurations, theDay, weekDays, resultSheet;

function run() {
  init();
  printTableTitles();
  printColumnTitles();
  var data = mapToDay(collect());
  data.forEach(printDay);
}

function collect() {
  var tab   = SpreadsheetApp.openById('1jCTKzHt2k4YRGBL9b3jigYuhSY4BNNBg00ejU_jIfQ8').getSheetByName('data').getRange('A2:O1500').getValues();
  var langs = SpreadsheetApp.openById('1jCTKzHt2k4YRGBL9b3jigYuhSY4BNNBg00ejU_jIfQ8').getSheetByName('data').getRange('C1:O1').getValues()[0].filter(function (l) {
    return !!l;
  });
  return tab
    .filter(function (d) {
      return !!d[0];
    })
    .map(function (d) {
      var obj = {
        start: d[0],
        end: d[1],
        duration: d[1] - d[0],
        titles: {}
      };

      langs.forEach(function (l, i) {
        obj.titles[l] = d[i + 2];
      });
      return obj;
    });
}

function mapToDay(data) {
  return data
    .sort(function (a, b) {
      return a.start - b.start;
    })
    .reduce(function (result, val, i, arr) {
      var eventNum = result.length;
      if (i === 0 || isNextDayByTimeZone(arr[i - 1].start, val.start)) {
        eventNum++;
        result.push({ order: eventNum, events: [] });
      }

      result.forEach(function (item) {
        if (item.order !== eventNum) {
          return;
        }
        item.events.push(val);
      });
      return result;
    }, []);
}

function isNextDayByTimeZone(prev, now, tzone) {
  if (!tzone) {
    tzone = configurations[0].tzone;
  }

  var data1 = convertDateByZone(prev, tzone);
  var data2 = convertDateByZone(now, tzone);

  return data1.getUTCDay() !== data2.getUTCDay();
}

function printTableTitles() {
  configurations.forEach(function (x) {
    increesCurrentRowByLang(x.id);
    var range = resultSheet.getRange(currentRowByLang[x.id], rangeStart(x.order, 0), 1, 3);

    range
      .merge()
      .setBackground(x.h1c)
      .setFontSize(24)
      .setFontWeight('bold')
      .setValue(x.h1text)
      .setHorizontalAlignment('center');
  });
}

function printColumnTitles() {
  configurations.forEach(function (x) {
    increesCurrentRowByLang(x.id);
    resultSheet
      .getRange(currentRowByLang[x.id], rangeStart(x.order, 0), 1, 1)
      .merge()
      .setValue(x.tzone)
      .setHorizontalAlignment('center');
    resultSheet
      .getRange(currentRowByLang[x.id], rangeStart(x.order, 1), 1, 1)
      .merge()
      .setValue(x.col2Name)
      .setHorizontalAlignment('center');
    resultSheet
      .getRange(currentRowByLang[x.id], rangeStart(x.order, 2), 1, 1)
      .merge()
      .setValue(x.col3Name)
      .setHorizontalAlignment('center');
  });
}

function printDay(day) {
  equalizeCurrentRows();
  configurations.forEach(function (x) {
    printDayTitle(day, x);
  });

  day.events.forEach(printEvent);
}

function printDayTitle(day, langConfig) {
  increesCurrentRowByLang(langConfig.id);
  var range = resultSheet.getRange(currentRowByLang[langConfig.id], rangeStart(langConfig.order, 0), 1, 3);

  range
    .merge()
    .setBackground(langConfig.h2c)
    .setFontSize(16)
    .setFontWeight('bold')
    .setValue(langConfig.dayTitle + ' ' + day.order)
    .setHorizontalAlignment('center');
}

function printEvent(e, i, arr) {
  var addTitleTZone;
  configurations.forEach(function (x) {
    if (i !== 0 && isNextDayByTimeZone(arr[i - 1].start, e.start, x.tzone)) {
      addTitleTZone = x.tzone;
    }
  });

  configurations.forEach(function (x) {
    if (i === 0) {
      printDateTitle(e, x);
      printDayEvent(e, x);
      return;
    }

    if (!addTitleTZone) {
      printDayEvent(e, x);
    } else if (x.tzone === addTitleTZone) {
      printDateTitle(e, x);
      printDayEvent(e, x);
    } else {
      increesCurrentRowByLang(x.id);
      printDayEvent(e, x);
    }
  });

}

function printDateTitle(event, langConfig) {
  increesCurrentRowByLang(langConfig.id);

  var range      = resultSheet.getRange(currentRowByLang[langConfig.id], rangeStart(langConfig.order, 0), 1, 3);
  var dateByZone = convertDateByZone(event.end, langConfig.tzone);
  var title      = getWeekDayName(dateByZone.getUTCDay() - 1, langConfig.id) + ' ' + Utilities.formatDate(event.end, langConfig.tzone, 'yyyy-MM-dd');

  range
    .merge()
    .setFontWeight('bold')
    .setFontSize(12)
    .setValue(title)
    .setHorizontalAlignment('center');
}

function convertDateByZone(date, zone) {
  var str = Utilities.formatDate(new Date(date), zone, 'yyyy-MM-dd\'T\'HH:mm:ss.SSS\'Z\'');
  return new Date(str);
}

function printDayEvent(event, langConfig) {
  increesCurrentRowByLang(langConfig.id);
  var timeRange = resultSheet.getRange(currentRowByLang[langConfig.id], rangeStart(langConfig.order, 0), 1, 1);
  timeRange.setValue(formatTimeByZone(event.start, event.end, langConfig.tzone, langConfig.id === 'heb')).setHorizontalAlignment('center');

  var durationRange = resultSheet.getRange(currentRowByLang[langConfig.id], rangeStart(langConfig.order, 1), 1, 1);
  durationRange.setValue(Utilities.formatDate(new Date(event.duration), 'UTC', 'HH:mm').toString()).setHorizontalAlignment('center');

  var titleRange = resultSheet.getRange(currentRowByLang[langConfig.id], rangeStart(langConfig.order, 2), 1, 1).setHorizontalAlignment('center');
  titleRange.setValue(event.titles[langConfig.id]);
}

function formatTimeByZone(start, end, zone, isRtl) {
  var sStr = Utilities.formatDate(start, zone, 'HH:mm');
  var eStr = Utilities.formatDate(end, zone, 'HH:mm');
  return isRtl ? eStr + '-' + sStr : sStr + '-' + eStr;
}

function init() {
  var ss           = SpreadsheetApp.getActiveSpreadsheet();
  resultSheet      = ss.getSheetByName('board');
  configurations   = getConfiguration().sort(function (a, b) {
    return a.order - b.order;
  });
  currentRowByLang = {};
  theDay           = {};
  clearAll();
}

function increesCurrentRowByLang(lang) {
  if (!currentRowByLang[lang]) {
    currentRowByLang[lang] = 0;
  }
  currentRowByLang[lang] = 1 + currentRowByLang[lang];
  colorSeparator(lang);
}

function colorSeparator(lang) {
  var params = configurations.filter(function (x) {
    return x.id === lang;
  })[0];

  if (!params) {
    return;
  }
  var range = resultSheet.getRange(currentRowByLang[params.id], rangeStart(params.order, 3), 1, 1);

  range.setBackground('#EEEEEE');
}

function equalizeCurrentRows() {
  var max = 1;

  configurations
    .forEach(function (x) {
      if (currentRowByLang[x.id] > max) {
        max = currentRowByLang[x.id];
      }
    });

  configurations
    .forEach(function (x) {
      currentRowByLang[x.id] = max;
    });

}

function getConfiguration() {
  var tab    = SpreadsheetApp.openById('1jCTKzHt2k4YRGBL9b3jigYuhSY4BNNBg00ejU_jIfQ8').getSheetByName('configuration');
  var result = [];
  var i      = 2;
  do {
    var r      = {};
    r.order    = tab.getRange('A' + i).getValue();
    r.id       = tab.getRange('B' + i).getValue();
    r.h1c      = tab.getRange('C' + i).getBackground();
    r.h2c      = tab.getRange('D' + i).getBackground();
    r.h1text   = tab.getRange('E' + i).getValue();
    r.tzone    = tab.getRange('F' + i).getValue() || 'Asia/Jerusalem';
    r.col1Name = tab.getRange('G' + i).getValue() + ' (' + r.tzone + ')';
    r.col2Name = tab.getRange('H' + i).getValue();
    r.col3Name = tab.getRange('I' + i).getValue();
    r.dayTitle = tab.getRange('J' + i).getValue();
    result.push(r);
    i++;
  } while (tab.getRange('B' + i).getValue());
  return result;
}

function clearAll() {
  var width = configurations.length * 4;
  var range = resultSheet.getRange(1, 1, 100, width);
  range.clear();
  var merged = range.getMergedRanges();
  if (merged.length > 0) {
    range.getMergedRanges().breakApart();
  }
  return range;
}

function getWeekDayName(day, lang) {
  if (!weekDays) {
    weekDays = SpreadsheetApp.openById('1B76zdIX2p48FEA1fvJr36DHKsQaIODQT9kWUZ8n0c7o').getSheetByName('config').getRange(3, 6, 7, 4).getValues();
  }

  if (day < 0) {
    day = 7 + day;
  }

  switch (lang) {
  case 'heb':
    return weekDays[day][0];
  case 'rus':
    return weekDays[day][2];
  case 'esp':
    return weekDays[day][3];
  default:
    return weekDays[day][1];
  }
}

function rangeStart(order, skip) {
  return order * 4 - 3 + skip;
}