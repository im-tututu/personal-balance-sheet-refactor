var REFACTOR_SHEET_NAMES = {
  assets: '资产清单',
  assetSnapshots: '资产清单快照',
  flows: '资金流水',
  overview: '概览',
  balance: '总资产负债',
  snapshots: '净值快照',
  marketValueHistory: '历史分项快照',
  config: '脚本配置',
  validation: '重算校验'
};

var REFACTOR_COLUMNS = {
  assets: {
    name: 1,
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
    dailyReturn: 6,
    nav: 7,
    drawdown: 8,
    cumulativeNetFlow: 9,
    shares: 10
  }
};
var REFACTOR_PRICE_SOURCES = ['sh', 'sz', 'bj', 'hk', 'of'];
var REFACTOR_SNAPSHOT_START_DATE = new Date(2019, 0, 1);
var REFACTOR_SOURCE_SPREADSHEET_ID = '1m8l-5XBg5wUcR1fFaRFU2fKYqVHTCsjcDBWIFPznOm4';
var REFACTOR_SOURCE_MARKET_VALUE_SHEET = '市值记录';
var REFACTOR_MARKET_VALUE_ARCHIVE_COLUMNS = [
  '日期',
  '市值W',
  '券商',
  '支付宝',
  '网商银行',
  '股权类',
  '债权类',
  '商品',
  '现金',
  '费率',
  '日费用',
  '总收益',
  '券商收益',
  '支付宝基金收益'
];

var REFACTOR_ASSET_SNAPSHOT_DATE_HEADER = '快照日期';

function logRefactor_(message, payload) {
  if (typeof payload === 'undefined') {
    Logger.log('[refactor] ' + message);
    return;
  }
  Logger.log('[refactor] ' + message + ' | ' + JSON.stringify(payload));
}

function getCurrentTotalLiabilities_(ss) {
  return -sumAmounts_(ss, function(amount) { return amount < 0; });
}

function getCurrentNetAssets_(ss) {
  return getCurrentTotalAssets_(ss);
}

function ensureSnapshotSheetLayout_(sheet) {
  var headers = [[
    '日期',
    '总资产',
    '',
    '',
    '当日净现金流',
    '日收益',
    '净值',
    '回撤',
    '净现金流合计',
    '份额'
  ]];
  sheet.getRange(1, 1, 1, headers[0].length).setValues(headers);
}

function ensureAssetSnapshotSheetLayout_(sheet, assetHeaders) {
  var headers = [[REFACTOR_ASSET_SNAPSHOT_DATE_HEADER].concat(assetHeaders)];
  sheet.clearContents();
  sheet.getRange(1, 1, 1, headers[0].length).setValues(headers);
}

function sumAmounts_(ss, predicate) {
  var assetSheet = ss.getSheetByName(REFACTOR_SHEET_NAMES.assets);
  if (!assetSheet || assetSheet.getLastRow() <= 1) return 0;

  var values = assetSheet.getRange(2, 1, assetSheet.getLastRow() - 1, REFACTOR_COLUMNS.assets.amount).getValues();
  return values.reduce(function(sum, row) {
    var name = row[REFACTOR_COLUMNS.assets.name - 1];
    var amount = toNumber_(row[REFACTOR_COLUMNS.assets.amount - 1]);
    if (!isRealAssetRow_(name, amount)) return sum;
    if (!isFinite(amount) || !predicate(amount)) return sum;
    return sum + amount;
  }, 0);
}

function buildDailyInvestmentFlowMap_(ss) {
  var flowSheet = ss.getSheetByName(REFACTOR_SHEET_NAMES.flows);
  if (!flowSheet || flowSheet.getLastRow() <= 1) return {};

  var values = flowSheet.getRange(
    2,
    1,
    flowSheet.getLastRow() - 1,
    REFACTOR_COLUMNS.flows.cashflow
  ).getValues();

  var map = {};

  values.forEach(function(row) {
    var rawType = row[REFACTOR_COLUMNS.flows.type - 1];
    var type = String(rawType == null ? '' : rawType).replace(/\s+/g, '');
    var dateValue = row[REFACTOR_COLUMNS.flows.date - 1];
    var amount = toNumber_(row[REFACTOR_COLUMNS.flows.cashflow - 1]);
    var normalizedDate = normalizeDate_(dateValue);

    if (type !== '投资') return;
    if (!normalizedDate || !isFinite(amount)) return;

    var key = formatDateKey_(normalizedDate);

    // 关键：从“账户现金视角”翻成“组合资金流入视角”
    map[key] = (map[key] || 0) - amount;
  });

  return map;
}

function getCumulativeInvestmentFlowAsOf_(dailyMap, date) {
  var targetKey = formatDateKey_(date);
  var keys = Object.keys(dailyMap).sort();
  var sum = 0;

  for (var i = 0; i < keys.length; i++) {
    if (keys[i] <= targetKey) {
      sum += dailyMap[keys[i]];
    }
  }
  return sum;
}

function getSnapshotOpeningCumulativeFlow_(ss) {
  var dailyMap = buildDailyInvestmentFlowMap_(ss);
  var dayBeforeStart = new Date(REFACTOR_SNAPSHOT_START_DATE);
  dayBeforeStart.setDate(dayBeforeStart.getDate() - 1);
  return getCumulativeInvestmentFlowAsOf_(dailyMap, dayBeforeStart);
}


function buildCumulativeInvestmentFlowMap_(ss) {
  var dailyMap = buildDailyInvestmentFlowMap_(ss);
  var keys = Object.keys(dailyMap).sort();
  var cumulativeMap = {};
  var running = 0;

  keys.forEach(function(key) {
    running += dailyMap[key];
    cumulativeMap[key] = running;
  });

  return cumulativeMap;
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

function isRealAssetRow_(name, amount) {
  if (!name) return false;
  if (!isFinite(amount)) return false;
  var text = String(name).trim();
  if (!text) return false;
  if (text === '`') return false;
  return true;
}

function finalizeSnapshotRows_(rows, openingCumulativeNetFlow) {
  if (!rows.length) return;

  rows.sort(function(a, b) {
    return normalizeDate_(a[0]).getTime() - normalizeDate_(b[0]).getTime();
  });

  var runningPeak = 1;
  var openingFlow = isFinite(toNumber_(openingCumulativeNetFlow))
    ? toNumber_(openingCumulativeNetFlow)
    : 0;

  for (var i = 0; i < rows.length; i++) {
    var currentTotalAssets = toNumber_(rows[i][1]) || 0;

    var cumulativeNetFlow = toNumber_(rows[i][8]);
    if (!isFinite(cumulativeNetFlow)) cumulativeNetFlow = 0;

    var previousCumulativeNetFlow = i > 0
      ? (toNumber_(rows[i - 1][8]) || 0)
      : openingFlow;

    var dayFlow = cumulativeNetFlow - previousCumulativeNetFlow;

    var previousTotalAssets = i > 0 ? (toNumber_(rows[i - 1][1]) || 0) : currentTotalAssets;
    var dailyReturn = i === 0 ? 0 : (currentTotalAssets - previousTotalAssets - dayFlow);

    var prevNav = i > 0 ? (toNumber_(rows[i - 1][6]) || 1) : 1;
    if (!isFinite(prevNav) || prevNav <= 0) prevNav = 1;

    var prevShares = i > 0 ? (toNumber_(rows[i - 1][9]) || 0) : 0;

    var shares;
    var nav;

    if (i === 0) {
      nav = 1;
      shares = currentTotalAssets / nav;
    } else {
      shares = prevShares + dayFlow / prevNav;
      nav = shares ? (currentTotalAssets / shares) : 0;
    }

    runningPeak = Math.max(runningPeak, nav || 0);
    var drawdown = runningPeak ? (nav / runningPeak) - 1 : 0;

    rows[i][2] = '';
    rows[i][3] = '';
    rows[i][4] = dayFlow;
    rows[i][5] = dailyReturn;
    rows[i][6] = nav;
    rows[i][7] = drawdown;
    rows[i][8] = cumulativeNetFlow;
    rows[i][9] = shares;
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
  values.push(toNumber_(lastRow[1]) || 0);
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
    ['latest_net_assets', summary.netAssets === '' ? '' : (summary.netAssets || 0), 'GAS 更新'],
    ['latest_daily_return', summary.dailyReturn || 0, 'GAS 更新'],
    ['xirr_all', summary.xirr === '' ? '' : summary.xirr, 'GAS 更新'],
    ['max_drawdown', summary.maxDrawdown || 0, 'GAS 更新'],
    ['drawdown_from', summary.drawdownFrom || '', 'GAS 更新'],
    ['drawdown_to', summary.drawdownTo || '', 'GAS 更新']
  ];

  sheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
}

function refactorRefreshValidationSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(REFACTOR_SHEET_NAMES.validation);
  if (!sheet) {
    sheet = ss.insertSheet(REFACTOR_SHEET_NAMES.validation);
  } else {
    sheet.clearContents();
  }

  var snapshotSheet = ss.getSheetByName(REFACTOR_SHEET_NAMES.snapshots);
  if (!snapshotSheet || snapshotSheet.getLastRow() <= 1) return;

  var newRows = snapshotSheet
    .getRange(2, 1, snapshotSheet.getLastRow() - 1, REFACTOR_COLUMNS.snapshots.cumulativeNetFlow)
    .getValues()
    .filter(function(row) { return row[0]; });

  var historicalRows = getHistoricalMarketValueRows_();
  var historicalMap = {};
  historicalRows.forEach(function(row) {
    historicalMap[formatDateKey_(row.date)] = row.totalAssets;
  });

  var output = [[
    '日期',
    '旧表总资产',
    '新表总资产',
    '总资产差额',
    '净现金流合计',
    '当日净现金流',
    '新表日收益',
    '新表净值',
    '新表回撤',
    '新表份额'
  ]];

  newRows.forEach(function(row) {
    var dateKey = formatDateKey_(normalizeDate_(row[0]));
    var oldTotalAssets = historicalMap[dateKey];
    var newTotalAssets = toNumber_(row[1]);
    output.push([
      row[0],
      oldTotalAssets,
      newTotalAssets,
      isFinite(oldTotalAssets) && isFinite(newTotalAssets) ? newTotalAssets - oldTotalAssets : '',
      row[8],
      row[4],
      row[5],
      row[6],
      row[7],
      row[9]
    ]);
  });

  sheet.getRange(1, 1, output.length, output[0].length).setValues(output);
}

function refactorArchiveHistoricalMarketValue_() {
  var targetSs = SpreadsheetApp.getActiveSpreadsheet();
  var targetSheet = getOrCreateSheet_(targetSs, REFACTOR_SHEET_NAMES.marketValueHistory);
  var sourceSs = SpreadsheetApp.openById(REFACTOR_SOURCE_SPREADSHEET_ID);
  var sourceSheet = sourceSs.getSheetByName(REFACTOR_SOURCE_MARKET_VALUE_SHEET);
  if (!sourceSheet || sourceSheet.getLastRow() < 2 || sourceSheet.getLastColumn() < 1) return;

  var headerRowIndex = findMarketValueHeaderRow_(sourceSheet);
  if (!headerRowIndex) {
    throw new Error('旧表“市值记录”中未找到表头行，无法迁移历史分项。');
  }

  var headers = sourceSheet.getRange(headerRowIndex, 1, 1, sourceSheet.getLastColumn()).getValues()[0];
  var lastDataRow = sourceSheet.getLastRow();
  if (lastDataRow <= headerRowIndex) {
    if (targetSheet.getLastRow() < 1) {
      targetSheet.getRange(1, 1, 1, REFACTOR_MARKET_VALUE_ARCHIVE_COLUMNS.length)
        .setValues([REFACTOR_MARKET_VALUE_ARCHIVE_COLUMNS]);
    }
    return;
  }

  var data = sourceSheet.getRange(headerRowIndex + 1, 1, lastDataRow - headerRowIndex, sourceSheet.getLastColumn()).getValues();
  var headerIndexMap = buildHeaderIndexMap_(headers);
  var selectedHeaders = REFACTOR_MARKET_VALUE_ARCHIVE_COLUMNS.filter(function(header) {
    return headerIndexMap.hasOwnProperty(header);
  });
  var existingDateMap = getExistingSheetDateMap_(targetSheet);
  var summaries = [];

  data.forEach(function(row) {
    var date = normalizeDate_(row[headerIndexMap['日期']]);
    if (!date || date < REFACTOR_SNAPSHOT_START_DATE) return;
    var dateKey = formatDateKey_(date);
    if (existingDateMap[dateKey]) return;

    var summary = createMarketValueHistorySummary_(date);
    selectedHeaders.forEach(function(header) {
      summary[header] = row[headerIndexMap[header]];
    });
    summaries.push(summary);
  });

  if (!summaries.length) {
    if (targetSheet.getLastRow() < 1) {
      targetSheet.getRange(1, 1, 1, REFACTOR_MARKET_VALUE_ARCHIVE_COLUMNS.length)
        .setValues([REFACTOR_MARKET_VALUE_ARCHIVE_COLUMNS]);
    }
    logRefactor_('历史分项快照补历史完成', { inserted_dates: 0 });
    return { insertedDates: 0 };
  }

  var result = upsertMarketValueHistoryRows_(targetSs, summaries);
  logRefactor_('历史分项快照补历史完成', { inserted_dates: result.upsertedDates });
  return { insertedDates: result.upsertedDates };
}

function getHistoricalMarketValueRows_() {
  var sourceSs = SpreadsheetApp.openById(REFACTOR_SOURCE_SPREADSHEET_ID);
  var sheet = sourceSs.getSheetByName(REFACTOR_SOURCE_MARKET_VALUE_SHEET);
  if (!sheet || sheet.getLastRow() < 2) return [];

  var headerRowIndex = findMarketValueHeaderRow_(sheet);
  if (!headerRowIndex) return [];

  var headers = sheet.getRange(headerRowIndex, 1, 1, sheet.getLastColumn()).getValues()[0];
  var headerIndexMap = buildHeaderIndexMap_(headers);
  if (!headerIndexMap.hasOwnProperty('日期') || !headerIndexMap.hasOwnProperty('市值W')) {
    throw new Error('旧表“市值记录”缺少“日期”或“市值W”列。');
  }

  var values = sheet.getRange(headerRowIndex + 1, 1, sheet.getLastRow() - headerRowIndex, sheet.getLastColumn()).getValues();
  return values
    .map(function(row) {
      var date = normalizeDate_(row[headerIndexMap['日期']]);
      var totalAssets = toNumber_(row[headerIndexMap['市值W']]);
      return {
        date: date,
        totalAssets: totalAssets
      };
    })
    .filter(function(row) {
      return row.date && isFinite(row.totalAssets) && row.date >= REFACTOR_SNAPSHOT_START_DATE;
    })
    .sort(function(a, b) {
      return a.date.getTime() - b.date.getTime();
    });
}

function findMarketValueHeaderRow_(sheet) {
  var maxRows = Math.min(sheet.getLastRow(), 10);
  if (!maxRows) return 0;
  var values = sheet.getRange(1, 1, maxRows, sheet.getLastColumn()).getValues();
  for (var i = 0; i < values.length; i++) {
    var row = values[i];
    if (row.indexOf('日期') > -1 && row.indexOf('市值W') > -1) {
      return i + 1;
    }
  }
  return 0;
}

function buildHeaderIndexMap_(headers) {
  var map = {};
  headers.forEach(function(header, index) {
    var key = String(header || '').trim();
    if (key) map[key] = index;
  });
  return map;
}

function getOrCreateSheet_(ss, sheetName) {
  var sheet = ss.getSheetByName(sheetName);
  return sheet || ss.insertSheet(sheetName);
}

function getNormalizedHeaderValues_(sheet) {
  if (!sheet || sheet.getLastColumn() < 1) return [];
  return sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(function(header) {
    return String(header || '').trim();
  });
}

function getExistingSheetDateMap_(sheet) {
  var map = {};
  if (!sheet || sheet.getLastRow() <= 1) return map;

  var values = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
  values.forEach(function(row) {
    var date = normalizeDate_(row[0]);
    if (!date) return;
    map[formatDateKey_(date)] = true;
  });
  return map;
}

// 每晚保留一份“非零资产行”快照，后续所有分项汇总都基于这里继续演进。
function refactorUpsertCurrentAssetSnapshot_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var assetSheet = ss.getSheetByName(REFACTOR_SHEET_NAMES.assets);
  if (!assetSheet || assetSheet.getLastRow() <= 1) return;

  var assetHeaders = getNormalizedHeaderValues_(assetSheet);
  var assetData = assetSheet.getRange(2, 1, assetSheet.getLastRow() - 1, assetSheet.getLastColumn()).getValues();
  var snapshotSheet = getOrCreateSheet_(ss, REFACTOR_SHEET_NAMES.assetSnapshots);
  var snapshotHeaders = getNormalizedHeaderValues_(snapshotSheet);
  var expectedHeaders = [REFACTOR_ASSET_SNAPSHOT_DATE_HEADER].concat(assetHeaders);

  if (!snapshotHeaders.length || snapshotHeaders.join('\t') !== expectedHeaders.join('\t')) {
    ensureAssetSnapshotSheetLayout_(snapshotSheet, assetHeaders);
  }

  var today = normalizeDate_(new Date());
  var todayKey = formatDateKey_(today);
  var keptRows = [];

  if (snapshotSheet.getLastRow() > 1) {
    var existingRows = snapshotSheet.getRange(2, 1, snapshotSheet.getLastRow() - 1, expectedHeaders.length).getValues();
    keptRows = existingRows.filter(function(row) {
      var rowDate = normalizeDate_(row[0]);
      return !rowDate || formatDateKey_(rowDate) !== todayKey;
    });
  }

  var todayRows = assetData
    .filter(function(row) {
      var name = row[REFACTOR_COLUMNS.assets.name - 1];
      var amount = toNumber_(row[REFACTOR_COLUMNS.assets.amount - 1]);
      return isRealAssetRow_(name, amount) && amount !== 0;
    })
    .map(function(row) {
      return [today].concat(row);
    });

  var output = [expectedHeaders].concat(keptRows, todayRows);
  snapshotSheet.clearContents();
  snapshotSheet.getRange(1, 1, output.length, output[0].length).setValues(output);
  var result = {
    date: todayKey,
    snapshotRows: todayRows.length,
    totalRows: output.length - 1
  };
  logRefactor_('资产清单快照更新完成', result);
  return result;
}

// 历史分项快照按日期聚合资产清单快照；同一天重复跑时直接覆盖当天结果。
function refactorUpdateMarketValueHistoryFromAssetSnapshots_(targetDateKeys) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var snapshotSheet = ss.getSheetByName(REFACTOR_SHEET_NAMES.assetSnapshots);
  if (!snapshotSheet || snapshotSheet.getLastRow() <= 1) {
    logRefactor_('历史分项快照更新跳过', { reason: 'no_asset_snapshots' });
    return { upsertedDates: 0 };
  }

  var brokerFlowMap = buildAccountCumulativeFlowMap_(ss, function(account) {
    return /^投资-证券/.test(account);
  });
  var alipayFundFlowMap = buildAccountCumulativeFlowMap_(ss, function(account) {
    return /支付宝-基金$/.test(account);
  });
  var investmentFlowMap = buildCumulativeInvestmentFlowMap_(ss);
  var headers = getNormalizedHeaderValues_(snapshotSheet);
  var headerMap = buildHeaderIndexMap_(headers);
  var rows = snapshotSheet.getRange(2, 1, snapshotSheet.getLastRow() - 1, snapshotSheet.getLastColumn()).getValues();
  var grouped = {};
  var targetMap = null;

  if (targetDateKeys && targetDateKeys.length) {
    targetMap = {};
    targetDateKeys.forEach(function(dateKey) {
      targetMap[dateKey] = true;
    });
  }

  rows.forEach(function(row) {
    var date = normalizeDate_(row[headerMap[REFACTOR_ASSET_SNAPSHOT_DATE_HEADER]]);
    if (!date) return;
    var dateKey = formatDateKey_(date);
    if (targetMap && !targetMap[dateKey]) return;

    var amount = getAssetSnapshotAmount_(row, headerMap);
    if (!isFinite(amount) || amount === 0) return;
    var annualFee = getAssetSnapshotAnnualFee_(row, headerMap);

    if (!grouped[dateKey]) {
      grouped[dateKey] = createMarketValueHistorySummary_(date);
    }

    if (amount > 0) {
      grouped[dateKey]['市值W'] += amount;
    }

    var accountBucket = classifyAccountBucketFromAssetSnapshot_(row, headerMap);
    if (accountBucket && amount > 0) {
      grouped[dateKey][accountBucket] += amount;
    }

    var assetClassBucket = classifyAssetClassBucketFromAssetSnapshot_(row, headerMap);
    if (assetClassBucket && amount > 0) {
      grouped[dateKey][assetClassBucket] += amount;
    }

    if (isFinite(annualFee)) {
      grouped[dateKey].annualFeeSum += annualFee;
      grouped[dateKey]['日费用'] = grouped[dateKey].annualFeeSum / 365;
    }
  });

  var summaries = Object.keys(grouped).sort().map(function(dateKey) {
    var summary = grouped[dateKey];
    var totalAssets = toNumber_(summary['市值W']) || 0;
    summary['费率'] = totalAssets ? (summary.annualFeeSum / totalAssets) : 0;
    summary['总收益'] = totalAssets - getCumulativeValueFromMapAsOf_(investmentFlowMap, dateKey);
    summary['券商收益'] = (toNumber_(summary['券商']) || 0) + getCumulativeValueFromMapAsOf_(brokerFlowMap, dateKey);
    summary['支付宝基金收益'] = (toNumber_(summary['支付宝']) || 0) + getCumulativeValueFromMapAsOf_(alipayFundFlowMap, dateKey);
    delete summary.annualFeeSum;
    return grouped[dateKey];
  });
  var result = upsertMarketValueHistoryRows_(ss, summaries);
  logRefactor_('历史分项快照更新完成', {
    target_dates: targetDateKeys && targetDateKeys.length ? targetDateKeys : 'all',
    upserted_dates: result.upsertedDates
  });
  return result;
}

function getAssetSnapshotAmount_(row, headerMap) {
  var amount = getNumericValueByHeaderCandidates_(row, headerMap, ['金额', '总金额', '资产金额', '市值']);
  if (isFinite(amount)) return amount;
  return toNumber_(row[REFACTOR_COLUMNS.assets.amount]);
}

function createMarketValueHistorySummary_(date) {
  var summary = {};
  REFACTOR_MARKET_VALUE_ARCHIVE_COLUMNS.forEach(function(header) {
    summary[header] = header === '日期' ? date : 0;
  });
  summary.annualFeeSum = 0;
  return summary;
}

function sumMetricIntoSummary_(summary, targetKey, row, headerMap, candidates) {
  var value = getNumericValueByHeaderCandidates_(row, headerMap, candidates);
  if (isFinite(value)) {
    summary[targetKey] += value;
  }
}

function getNumericValueByHeaderCandidates_(row, headerMap, candidates) {
  for (var i = 0; i < candidates.length; i++) {
    if (headerMap.hasOwnProperty(candidates[i])) {
      var value = toNumber_(row[headerMap[candidates[i]]]);
      if (isFinite(value)) return value;
    }
  }
  return NaN;
}

function classifyAccountBucketFromAssetSnapshot_(row, headerMap) {
  var institution = getTextByHeaderCandidates_(row, headerMap, ['机构']);
  if (institution === '券商' || institution === '支付宝' || institution === '网商银行') {
    return institution;
  }

  var text = getCombinedSnapshotText_(row, headerMap, [
    '机构',
    '名称',
    '平台',
    '账户',
    '归属账户',
    '一级分类',
    '二级分类',
    '资产类型'
  ]);

  if (!text) return '';
  if (text.indexOf('支付宝') > -1) return '支付宝';
  if (text.indexOf('网商银行') > -1) return '网商银行';
  if (text.indexOf('券商') > -1 || text.indexOf('证券') > -1) return '券商';
  return '';
}

function classifyAssetClassBucketFromAssetSnapshot_(row, headerMap) {
  var text = getCombinedSnapshotText_(row, headerMap, [
    '资产类型',
    '一级分类',
    '二级分类',
    '名称'
  ]);

  if (!text) return '';
  if (text.indexOf('商品') > -1 || text.indexOf('黄金') > -1) return '商品';
  if (text.indexOf('债') > -1 || text.indexOf('固收') > -1) return '债权类';
  if (text.indexOf('现金') > -1 || text.indexOf('活期') > -1 || text.indexOf('存款') > -1) return '现金';
  if (text.indexOf('股') > -1 || text.indexOf('权益') > -1 || text.indexOf('股票') > -1 || text.indexOf('基金') > -1) return '股权类';
  return '';
}

function getCombinedSnapshotText_(row, headerMap, headers) {
  return headers.map(function(header) {
    if (!headerMap.hasOwnProperty(header)) return '';
    return String(row[headerMap[header]] || '').trim();
  }).join(' ');
}

function getTextByHeaderCandidates_(row, headerMap, candidates) {
  for (var i = 0; i < candidates.length; i++) {
    if (headerMap.hasOwnProperty(candidates[i])) {
      return String(row[headerMap[candidates[i]]] || '').trim();
    }
  }
  return '';
}

function getAssetSnapshotAnnualFee_(row, headerMap) {
  return getNumericValueByHeaderCandidates_(row, headerMap, ['年化费用']);
}

function buildAccountCumulativeFlowMap_(ss, accountMatcher) {
  var flowSheet = ss.getSheetByName(REFACTOR_SHEET_NAMES.flows);
  if (!flowSheet || flowSheet.getLastRow() <= 1) return {};

  var headers = getNormalizedHeaderValues_(flowSheet);
  var headerMap = buildHeaderIndexMap_(headers);
  var accountIndex = getHeaderIndexByCandidates_(headerMap, ['账号', '账户']);
  var dateIndex = getHeaderIndexByCandidates_(headerMap, ['日期']);
  var amountIndex = getHeaderIndexByCandidates_(headerMap, ['金额', '现金流', '净流入', '发生金额']);
  if (accountIndex < 0 || dateIndex < 0 || amountIndex < 0) return {};

  var values = flowSheet.getRange(2, 1, flowSheet.getLastRow() - 1, flowSheet.getLastColumn()).getValues();
  var dailyMap = {};

  values.forEach(function(row) {
    var date = normalizeDate_(row[dateIndex]);
    if (!date) return;
    var account = String(row[accountIndex] || '').trim();
    if (!accountMatcher(account)) return;
    var amount = toNumber_(row[amountIndex]);
    if (!isFinite(amount)) return;

    var dateKey = formatDateKey_(date);
    dailyMap[dateKey] = (dailyMap[dateKey] || 0) + amount;
  });

  var keys = Object.keys(dailyMap).sort();
  var cumulativeMap = {};
  var running = 0;
  keys.forEach(function(key) {
    running += dailyMap[key];
    cumulativeMap[key] = running;
  });

  return cumulativeMap;
}

function getCumulativeValueFromMapAsOf_(cumulativeMap, targetDateKey) {
  var keys = Object.keys(cumulativeMap).sort();
  var value = 0;
  for (var i = 0; i < keys.length; i++) {
    if (keys[i] <= targetDateKey) {
      value = cumulativeMap[keys[i]];
    } else {
      break;
    }
  }
  return value;
}

function getHeaderIndexByCandidates_(headerMap, candidates) {
  for (var i = 0; i < candidates.length; i++) {
    if (headerMap.hasOwnProperty(candidates[i])) {
      return headerMap[candidates[i]];
    }
  }
  return -1;
}

function upsertMarketValueHistoryRows_(ss, summaries) {
  if (!summaries.length) return { upsertedDates: 0 };

  var sheet = getOrCreateSheet_(ss, REFACTOR_SHEET_NAMES.marketValueHistory);
  var headers = REFACTOR_MARKET_VALUE_ARCHIVE_COLUMNS;
  var incomingMap = {};

  summaries.forEach(function(summary) {
    incomingMap[formatDateKey_(summary['日期'])] = headers.map(function(header) {
      return summary[header];
    });
  });

  var existingRows = [];
  if (sheet.getLastRow() > 1) {
    var existingData = sheet.getRange(2, 1, sheet.getLastRow() - 1, headers.length).getValues();
    existingRows = existingData.filter(function(row) {
      var date = normalizeDate_(row[0]);
      return date && !incomingMap.hasOwnProperty(formatDateKey_(date));
    });
  }

  var newRows = Object.keys(incomingMap).sort().map(function(dateKey) {
    return incomingMap[dateKey];
  });

  var mergedRows = existingRows.concat(newRows).sort(function(a, b) {
    return normalizeDate_(a[0]).getTime() - normalizeDate_(b[0]).getTime();
  });

  var output = [headers].concat(mergedRows);
  sheet.clearContents();
  sheet.getRange(1, 1, output.length, headers.length).setValues(output);
  return { upsertedDates: newRows.length };
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


function getCurrentTotalAssets_(ss) {

  return sumAmounts_(ss, function(amount) { return amount > 0; });

}
