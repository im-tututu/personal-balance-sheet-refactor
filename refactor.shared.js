var REFACTOR_SHEET_NAMES = {
  assets: '资产清单',
  flows: '资金流水',
  overview: '概览',
  balance: '总资产负债',
  snapshots: '净值快照',
  config: '脚本配置'
};

var REFACTOR_COLUMNS = {
  assets: {
    code: 3,
    amount: 10
  },
  flows: {
    type: 1,
    date: 3,
    cashflow: 4
  },
  snapshots: {
    date: 1,
    totalAssets: 2,
    totalLiabilities: 3,
    netAssets: 4,
    netFlow: 5,
    pnl: 6,
    nav: 7,
    drawdown: 8
  }
};

var REFACTOR_PRICE_SOURCES = ['sh', 'sz', 'bj', 'hk', 'of'];

function getCurrentTotalAssets_(ss) {
  return sumAmounts_(ss, function(amount) { return amount > 0; });
}

function getCurrentTotalLiabilities_(ss) {
  return -sumAmounts_(ss, function(amount) { return amount < 0; });
}

function getCurrentNetAssets_(ss) {
  return getCurrentTotalAssets_(ss) - getCurrentTotalLiabilities_(ss);
}

function sumAmounts_(ss, predicate) {
  var assetSheet = ss.getSheetByName(REFACTOR_SHEET_NAMES.assets);
  if (!assetSheet || assetSheet.getLastRow() <= 1) return 0;

  var values = assetSheet.getRange(2, REFACTOR_COLUMNS.assets.amount, assetSheet.getLastRow() - 1, 1).getValues();
  return values.reduce(function(sum, row) {
    var amount = toNumber_(row[0]);
    if (!isFinite(amount) || !predicate(amount)) return sum;
    return sum + amount;
  }, 0);
}

// 投资流水是净值、XIRR 和快照的共同基础，统一在这里按天归集。
function buildDailyInvestmentFlowMap_(ss) {
  var flowSheet = ss.getSheetByName(REFACTOR_SHEET_NAMES.flows);
  if (!flowSheet || flowSheet.getLastRow() <= 1) return {};

  var values = flowSheet.getRange(2, 1, flowSheet.getLastRow() - 1, REFACTOR_COLUMNS.flows.cashflow).getValues();
  var map = {};

  values.forEach(function(row) {
    var type = row[REFACTOR_COLUMNS.flows.type - 1];
    var dateValue = row[REFACTOR_COLUMNS.flows.date - 1];
    var amount = toNumber_(row[REFACTOR_COLUMNS.flows.cashflow - 1]);

    if (type !== '投资' || !dateValue || !isFinite(amount)) return;
    var key = formatDateKey_(normalizeDate_(dateValue));
    map[key] = (map[key] || 0) + amount;
  });

  return map;
}

function refactorUpdateAssetPrices() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(REFACTOR_SHEET_NAMES.assets);
  if (!sheet || sheet.getLastRow() <= 1) return;

  var values = sheet.getRange(2, 1, sheet.getLastRow() - 1, 17).getValues();
  var quoteMap = fetchQuoteMap_(values);

  for (var i = 0; i < values.length; i++) {
    var code = values[i][2];
    if (!code || !quoteMap[code]) continue;

    values[i][3] = quoteMap[code].price;
    values[i][4] = quoteMap[code].date;
    values[i][6] = quoteMap[code].price;
    values[i][7] = quoteMap[code].price * toNumber_(values[i][5] || 0);
  }

  for (var row = 0; row < values.length; row++) {
    var totalPrice = toNumber_(values[row][7]);
    var fx = toNumber_(values[row][8]);
    values[row][9] = isFinite(totalPrice) && isFinite(fx) ? totalPrice * fx : values[row][9];
  }

  sheet.getRange(2, 1, values.length, values[0].length).setValues(values);
  rebuildAssetDerivedColumns_(ss);
}

function fetchQuoteMap_(assetRows) {
  var codes = assetRows
    .map(function(row) { return row[2]; })
    .filter(function(code) {
      return code && REFACTOR_PRICE_SOURCES.indexOf(String(code).slice(0, 2)) > -1;
    });

  if (!codes.length) return {};

  var sinaResponse = '';
  var tencentResponse = '';
  try {
    sinaResponse = UrlFetchApp.fetch('https://hq.sinajs.cn/list=' + codes.join(','), {
      headers: { Referer: 'https://finance.sina.com.cn' }
    }).getContentText('GBK');
  } catch (error) {
    Logger.log('新浪行情请求失败: ' + error);
  }

  try {
    var tencentCodes = codes.map(function(code) {
      return String(code).indexOf('of') === 0 ? String(code).replace('of', 'jj') : code;
    });
    tencentResponse = UrlFetchApp.fetch('https://qt.gtimg.cn/q=' + tencentCodes.join(',')).getContentText('GBK');
  } catch (error) {
    Logger.log('腾讯行情请求失败: ' + error);
  }

  var sinaMap = parseSinaQuotes_(sinaResponse);
  var tencentMap = parseTencentQuotes_(tencentResponse);
  var quoteMap = {};

  codes.forEach(function(code) {
    var codeStr = String(code);
    var market = codeStr.slice(0, 2);
    var sinaQuote = sinaMap[codeStr];
    var tencentKey = market === 'of' ? codeStr.replace('of', 'jj') : codeStr;
    var tencentQuote = tencentMap[tencentKey];
    var finalQuote = null;

    if (market === 'of' && tencentQuote) {
      finalQuote = {
        price: tencentQuote.price,
        date: tencentQuote.date
      };
    } else if (sinaQuote) {
      finalQuote = sinaQuote;
    }

    if (finalQuote && isFinite(finalQuote.price)) {
      quoteMap[codeStr] = finalQuote;
    }
  });

  return quoteMap;
}

function parseSinaQuotes_(responseText) {
  var map = {};
  if (!responseText) return map;

  var lines = responseText.split(';');
  lines.forEach(function(line) {
    var match = line.match(/var hq_str_(\w+)="([^"]*)"/);
    if (!match) return;

    var code = match[1];
    var fields = match[2].split(',');
    var market = code.slice(0, 2);
    var quote = null;

    if (market === 'sh' || market === 'sz' || market === 'bj') {
      quote = { price: Number(fields[3]), date: buildDateTime_(fields[30], fields[31]) };
    } else if (market === 'hk') {
      quote = { price: Number(fields[6]), date: buildDateTime_(fields[17], fields[18]) };
    } else if (market === 'of') {
      quote = { price: Number(fields[1]), date: fields[5] || '' };
    }

    if (quote) map[code] = quote;
  });

  return map;
}

function parseTencentQuotes_(responseText) {
  var map = {};
  if (!responseText) return map;

  var lines = responseText.split(';');
  lines.forEach(function(line) {
    var match = line.match(/v_(\w+)="([^"]*)"/);
    if (!match) return;
    var code = match[1];
    var fields = match[2].split('~');
    map[code] = {
      price: Number(fields[5]),
      date: fields[4] || ''
    };
  });

  return map;
}

function rebuildAssetDerivedColumns_(ss) {
  var sheet = ss.getSheetByName(REFACTOR_SHEET_NAMES.assets);
  if (!sheet) return;

  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) return;

  for (var row = 2; row <= lastRow; row++) {
    sheet.getRange(row, 8).setFormula('=IFERROR(F' + row + '*G' + row + ',)');
    sheet.getRange(row, 10).setFormula('=IFERROR(H' + row + '*I' + row + ',)');
    sheet.getRange(row, 11).setFormula('=IFERROR(J' + row + '/SUM(FILTER($J$2:$J,$J$2:$J>0)),)');
  }
}

function finalizeSnapshotRows_(rows) {
  if (!rows.length) return;

  var firstNetAssets = toNumber_(rows[0][3]) || 1;
  var runningPeak = 1;

  for (var i = 0; i < rows.length; i++) {
    var currentNetAssets = toNumber_(rows[i][3]) || 0;
    var currentFlow = toNumber_(rows[i][4]) || 0;
    var previousNetAssets = i > 0 ? (toNumber_(rows[i - 1][3]) || 0) : currentNetAssets - currentFlow;

    var pnl = i === 0 ? 0 : currentNetAssets - previousNetAssets - currentFlow;
    var nav = firstNetAssets ? currentNetAssets / firstNetAssets : 1;
    runningPeak = Math.max(runningPeak, nav || 0);
    var drawdown = runningPeak ? (nav / runningPeak) - 1 : 0;

    rows[i][5] = pnl;
    rows[i][6] = nav;
    rows[i][7] = drawdown;
  }
}

function calculatePortfolioXirr_(ss, snapshotRows) {
  var dailyFlowMap = buildDailyInvestmentFlowMap_(ss);
  var keys = Object.keys(dailyFlowMap).sort();
  if (!keys.length || !snapshotRows.length) return '';

  var values = [];
  var dates = [];

  keys.forEach(function(key) {
    var amount = dailyFlowMap[key];
    if (!amount) return;
    values.push(amount);
    dates.push(parseDateKey_(key));
  });

  var lastRow = snapshotRows[snapshotRows.length - 1];
  values.push(toNumber_(lastRow[3]) || 0);
  dates.push(normalizeDate_(lastRow[0]));

  var result = XIRR(values, dates, 0.1);
  return typeof result === 'number' && isFinite(result) ? result : '';
}

function getMaxDrawdownMeta_(rows) {
  var runningPeakNav = -Infinity;
  var runningPeakDate = '';
  var maxDrawdown = 0;
  var maxFrom = '';
  var maxTo = '';

  rows.forEach(function(row) {
    var date = normalizeDate_(row[0]);
    var nav = toNumber_(row[6]) || 0;
    if (nav > runningPeakNav) {
      runningPeakNav = nav;
      runningPeakDate = date;
    }
    if (!runningPeakNav) return;

    var drawdown = (nav / runningPeakNav) - 1;
    if (drawdown < maxDrawdown) {
      maxDrawdown = drawdown;
      maxFrom = runningPeakDate;
      maxTo = date;
    }
  });

  return {
    maxDrawdown: maxDrawdown,
    fromDate: maxFrom || '',
    toDate: maxTo || ''
  };
}

function refactorUpdateSummary_(summary) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(REFACTOR_SHEET_NAMES.config);
  if (!sheet) throw new Error('找不到工作表: ' + REFACTOR_SHEET_NAMES.config);

  var rows = [
    ['latest_snapshot_date', summary.latestDate || '', 'GAS 更新'],
    ['latest_total_assets', summary.totalAssets || 0, 'GAS 更新'],
    ['latest_net_assets', summary.netAssets || 0, 'GAS 更新'],
    ['latest_daily_return', summary.dailyReturn || 0, 'GAS 更新'],
    ['xirr_all', summary.xirr === '' ? '' : summary.xirr, 'GAS 更新'],
    ['max_drawdown', summary.maxDrawdown || 0, 'GAS 更新'],
    ['drawdown_from', summary.drawdownFrom || '', 'GAS 更新'],
    ['drawdown_to', summary.drawdownTo || '', 'GAS 更新']
  ];

  sheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
}

function getConfigValue_(ss, key) {
  var sheet = ss.getSheetByName(REFACTOR_SHEET_NAMES.config);
  if (!sheet || sheet.getLastRow() <= 1) return '';
  var values = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();
  for (var i = 0; i < values.length; i++) {
    if (values[i][0] === key) return values[i][1];
  }
  return '';
}

function buildDateTime_(datePart, timePart) {
  if (!datePart) return '';
  return timePart ? String(datePart) + ' ' + String(timePart) : String(datePart);
}

function GetDateDiff(startDate, endDate) {
  var startTime = new Date(startDate).getTime();
  var endTime = new Date(endDate).getTime();
  var dates = Math.abs((startTime - endTime)) / (1000 * 60 * 60 * 24);
  return Math.trunc(dates) + 1;
}

function XIRR(values, dates, guess) {
  var irrResult = function(valuesInner, datesInner, rate) {
    var r = rate + 1;
    var result = valuesInner[0];
    for (var i = 1; i < valuesInner.length; i++) {
      var datediff = GetDateDiff(datesInner[i], datesInner[0]);
      result += valuesInner[i] / Math.pow(r, datediff / 365);
    }
    return result;
  };

  var irrResultDeriv = function(valuesInner, datesInner, rate) {
    var r = rate + 1;
    var result = 0;
    for (var i = 1; i < valuesInner.length; i++) {
      var datediff = GetDateDiff(datesInner[i], datesInner[0]);
      var frac = Math.pow(r, datediff) / 365;
      result -= frac * valuesInner[i] / Math.pow(r, frac + 1);
    }
    return result;
  };

  var positive = false;
  var negative = false;
  for (var j = 0; j < values.length; j++) {
    if (values[j] > 0) positive = true;
    if (values[j] < 0) negative = true;
  }
  if (!positive || !negative) return '#NUM!';

  var startGuess = (typeof guess === 'undefined') ? 0.1 : guess;
  var resultRate = startGuess;
  var epsMax = 1e-10;
  var iterMax = 50;
  var newRate;
  var epsRate;
  var resultValue;
  var iteration = 0;
  var contLoop = true;

  do {
    resultValue = irrResult(values, dates, resultRate);
    newRate = resultRate - resultValue / irrResultDeriv(values, dates, resultRate);
    epsRate = Math.abs(newRate - resultRate);
    resultRate = newRate;
    contLoop = (epsRate > epsMax) && (Math.abs(resultValue) > epsMax);
  } while (contLoop && (++iteration < iterMax));

  if (contLoop) return '#NUM!';
  return resultRate;
}

function normalizeDate_(value) {
  if (!value) return null;
  var date = value instanceof Date ? new Date(value) : new Date(value);
  if (isNaN(date.getTime())) return null;
  date.setHours(0, 0, 0, 0);
  return date;
}

function formatDateKey_(date) {
  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function parseDateKey_(key) {
  var parts = key.split('-');
  return new Date(Number(parts[0]), Number(parts[1]) - 1, Number(parts[2]));
}

function toNumber_(value) {
  if (typeof value === 'number') return value;
  if (value === null || value === '') return NaN;
  var cleaned = String(value).replace(/[,\s￥]/g, '');
  return Number(cleaned);
}
