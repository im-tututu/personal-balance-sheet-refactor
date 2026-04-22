var REFACTOR_SHEET_NAMES = {
  assets: '资产清单',
  flows: '资金流水',
  snapshots: '净值快照'
};

var REFACTOR_COLUMNS = {
  assets: {
    name: 1,
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
    netFlow: 3,
    dailyReturn: 4,
    cumulativeNetFlow: 5,
    shares: 6,
    recalculatedNav: 7,
    broker: 8,
    alipay: 9,
    mybank: 10,
    equity: 11,
    debt: 12,
    commodity: 13,
    cash: 14,
    feeRate: 15,
    dailyFee: 16,
    totalProfit: 17,
    brokerProfit: 18,
    alipayFundProfit: 19
  }
};

var REFACTOR_SNAPSHOT_START_DATE = new Date(2019, 0, 1);
var REFACTOR_SOURCE_SPREADSHEET_ID = '1m8l-5XBg5wUcR1fFaRFU2fKYqVHTCsjcDBWIFPznOm4';
var REFACTOR_SOURCE_MARKET_VALUE_SHEET = '市值记录';
var REFACTOR_SNAPSHOT_COLUMNS = [
  '日期',
  '总资产',
  '当日净现金流',
  '日收益',
  '净现金流合计',
  '份额',
  '净值重算',
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

function logRefactor_(message, payload) {
  if (typeof payload === 'undefined') {
    Logger.log('[refactor] ' + message);
    return;
  }
  Logger.log('[refactor] ' + message + ' | ' + JSON.stringify(payload));
}

function ensureSnapshotSheetLayout_(sheet) {
  var headers = [REFACTOR_SNAPSHOT_COLUMNS];
  ensureSheetColumnCount_(sheet, headers[0].length);
  sheet.getRange(1, 1, 1, sheet.getMaxColumns()).clearContent();
  sheet.getRange(1, 1, 1, headers[0].length).setValues(headers);
}

function getSnapshotColumnCount_() {
  return REFACTOR_SNAPSHOT_COLUMNS.length;
}

function ensureSheetColumnCount_(sheet, expectedColumns) {
  var currentColumns = sheet.getMaxColumns();
  if (currentColumns < expectedColumns) {
    sheet.insertColumnsAfter(currentColumns, expectedColumns - currentColumns);
  } else if (currentColumns > expectedColumns) {
    sheet.deleteColumns(expectedColumns + 1, currentColumns - expectedColumns);
  }
}

function buildInvestmentFlowTimeline_(ss) {
  var flowSheet = ss.getSheetByName(REFACTOR_SHEET_NAMES.flows);
  if (!flowSheet || flowSheet.getLastRow() <= 1) return [];

  var headers = getNormalizedHeaderValues_(flowSheet);
  var headerMap = buildHeaderIndexMap_(headers);
  var typeIndex = getHeaderIndexByCandidates_(headerMap, ['类型', '流水类型']);
  var dateIndex = getHeaderIndexByCandidates_(headerMap, ['日期']);
  var amountIndex = getHeaderIndexByCandidates_(headerMap, ['现金流', '发生金额', '金额', '净流入']);
  if (typeIndex < 0 || dateIndex < 0 || amountIndex < 0) return [];

  var flowRange = flowSheet.getRange(2, 1, flowSheet.getLastRow() - 1, flowSheet.getLastColumn());
  var values = flowRange.getValues();
  var displayValues = flowRange.getDisplayValues();

  return values
    .map(function(row, rowIndex) {
      var type = String(row[typeIndex] == null ? '' : row[typeIndex]).replace(/\s+/g, '');
      var displayRow = displayValues[rowIndex] || [];
      var amount = toNumberWithFallback_(row[amountIndex], displayRow[amountIndex]);
      var date = normalizeDateTimeWithFallback_(row[dateIndex], displayRow[dateIndex]);
      if (type !== '投资' || !date || !isFinite(amount)) return null;
      return {
        date: date,
        amount: amount
      };
    })
    .filter(function(item) { return item; })
    .sort(function(a, b) {
      return a.date.getTime() - b.date.getTime();
    });
}

function getUndatedInvestmentFlowTotal_(ss) {
  var flowSheet = ss.getSheetByName(REFACTOR_SHEET_NAMES.flows);
  if (!flowSheet || flowSheet.getLastRow() <= 1) return 0;

  var headers = getNormalizedHeaderValues_(flowSheet);
  var headerMap = buildHeaderIndexMap_(headers);
  var typeIndex = getHeaderIndexByCandidates_(headerMap, ['类型', '流水类型']);
  var dateIndex = getHeaderIndexByCandidates_(headerMap, ['日期']);
  var amountIndex = getHeaderIndexByCandidates_(headerMap, ['现金流', '发生金额', '金额', '净流入']);
  if (typeIndex < 0 || dateIndex < 0 || amountIndex < 0) return 0;

  var flowRange = flowSheet.getRange(2, 1, flowSheet.getLastRow() - 1, flowSheet.getLastColumn());
  var values = flowRange.getValues();
  var displayValues = flowRange.getDisplayValues();
  var total = 0;

  values.forEach(function(row, rowIndex) {
    var type = String(row[typeIndex] == null ? '' : row[typeIndex]).replace(/\s+/g, '');
    if (type !== '投资') return;

    var displayRow = displayValues[rowIndex] || [];
    if (normalizeDateTimeWithFallback_(row[dateIndex], displayRow[dateIndex])) return;

    var amount = toNumberWithFallback_(row[amountIndex], displayRow[amountIndex]);
    if (isFinite(amount)) total += amount;
  });

  return total;
}

function buildAccountCumulativeFlowMap_(ss, accountMatcher) {
  var flowSheet = ss.getSheetByName(REFACTOR_SHEET_NAMES.flows);
  if (!flowSheet || flowSheet.getLastRow() <= 1) return {};

  var headers = getNormalizedHeaderValues_(flowSheet);
  var headerMap = buildHeaderIndexMap_(headers);
  var accountIndex = getHeaderIndexByCandidates_(headerMap, ['账号', '账户']);
  var dateIndex = getHeaderIndexByCandidates_(headerMap, ['日期']);
  var amountIndex = getHeaderIndexByCandidates_(headerMap, ['现金流', '发生金额', '金额', '净流入']);
  if (accountIndex < 0 || dateIndex < 0 || amountIndex < 0) return {};

  var flowRange = flowSheet.getRange(2, 1, flowSheet.getLastRow() - 1, flowSheet.getLastColumn());
  var values = flowRange.getValues();
  var displayValues = flowRange.getDisplayValues();
  var dailyMap = {};

  values.forEach(function(row, rowIndex) {
    var account = String(row[accountIndex] || '').trim();
    if (!accountMatcher(account)) return;

    var displayRow = displayValues[rowIndex] || [];
    var date = normalizeDateTimeWithFallback_(row[dateIndex], displayRow[dateIndex]);
    if (!date) return;

    var amount = toNumberWithFallback_(row[amountIndex], displayRow[amountIndex]);
    if (!isFinite(amount)) return;

    var dateKey = formatDateKey_(date);
    dailyMap[dateKey] = (dailyMap[dateKey] || 0) + amount;
  });

  var cumulativeMap = {};
  var running = 0;
  Object.keys(dailyMap).sort().forEach(function(dateKey) {
    running += dailyMap[dateKey];
    cumulativeMap[dateKey] = running;
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

function getCumulativeInvestmentFlowAsOfTimeline_(timeline, date) {
  var targetTime = normalizeDate_(date);
  if (!targetTime) return 0;
  targetTime.setHours(23, 59, 59, 999);

  var sum = 0;
  for (var i = 0; i < timeline.length; i++) {
    if (timeline[i].date.getTime() <= targetTime.getTime()) {
      sum += timeline[i].amount;
    } else {
      break;
    }
  }
  return sum;
}

function buildSnapshotSeedRows_(marketRows, flowTimeline, undatedFlowTotal) {
  if (!marketRows.length) return [];
  var openingUndatedFlow = isFinite(toNumber_(undatedFlowTotal)) ? toNumber_(undatedFlowTotal) : 0;

  return marketRows.map(function(row) {
    var totalAssets = toNumber_(row.totalAssets) || 0;
    var cumulativeNetFlow = getCumulativeInvestmentFlowAsOfTimeline_(flowTimeline, row.date) + openingUndatedFlow;
    return [
      row.date,
      totalAssets,
      isFinite(toNumber_(row.netFlow)) ? toNumber_(row.netFlow) : '',
      isFinite(toNumber_(row.dailyReturn)) ? toNumber_(row.dailyReturn) : '',
      cumulativeNetFlow,
      '',
      '',
      toNumberOrZero_(row.broker),
      toNumberOrZero_(row.alipay),
      toNumberOrZero_(row.mybank),
      toNumberOrZero_(row.equity),
      toNumberOrZero_(row.debt),
      toNumberOrZero_(row.commodity),
      toNumberOrZero_(row.cash),
      isFinite(toNumber_(row.feeRate)) ? toNumber_(row.feeRate) : '',
      isFinite(toNumber_(row.dailyFee)) ? toNumber_(row.dailyFee) : '',
      isFinite(toNumber_(row.totalProfit)) ? toNumber_(row.totalProfit) : '',
      toNumberOrZero_(row.brokerProfit),
      toNumberOrZero_(row.alipayFundProfit)
    ];
  });
}

function buildCurrentSnapshotRow_(ss, date) {
  var summary = summarizeCurrentAssets_(ss);
  var flowTimeline = buildInvestmentFlowTimeline_(ss);
  var cumulativeNetFlow = getCumulativeInvestmentFlowAsOfTimeline_(flowTimeline, date) + getUndatedInvestmentFlowTotal_(ss);
  var dateKey = formatDateKey_(normalizeDate_(date));
  var brokerFlowMap = buildAccountCumulativeFlowMap_(ss, function(account) {
    return /^投资-证券/.test(account);
  });
  var alipayFundFlowMap = buildAccountCumulativeFlowMap_(ss, function(account) {
    return /支付宝-基金$/.test(account);
  });

  return [
    normalizeDate_(date),
    summary.totalAssets,
    '',
    '',
    cumulativeNetFlow,
    '',
    '',
    summary.broker,
    summary.alipay,
    summary.mybank,
    summary.equity,
    summary.debt,
    summary.commodity,
    summary.cash,
    summary.feeRate,
    summary.dailyFee,
    summary.totalAssets - cumulativeNetFlow,
    summary.broker + getCumulativeValueFromMapAsOf_(brokerFlowMap, dateKey),
    summary.alipay + getCumulativeValueFromMapAsOf_(alipayFundFlowMap, dateKey)
  ];
}

function createSnapshotSummary_(date) {
  return {
    date: date,
    totalAssets: 0,
    netFlow: NaN,
    dailyReturn: NaN,
    broker: 0,
    alipay: 0,
    mybank: 0,
    equity: 0,
    debt: 0,
    commodity: 0,
    cash: 0,
    feeRate: 0,
    dailyFee: 0,
    totalProfit: NaN,
    brokerProfit: 0,
    alipayFundProfit: 0
  };
}

function summarizeCurrentAssets_(ss) {
  var sheet = ss.getSheetByName(REFACTOR_SHEET_NAMES.assets);
  var summary = createSnapshotSummary_(new Date());
  if (!sheet || sheet.getLastRow() <= 1) return summary;

  var headers = getNormalizedHeaderValues_(sheet);
  var headerMap = buildHeaderIndexMap_(headers);
  var values = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
  var annualFeeSum = 0;

  values.forEach(function(row) {
    var name = row[REFACTOR_COLUMNS.assets.name - 1];
    var amount = getAssetAmount_(row, headerMap);
    if (!isRealAssetRow_(name, amount) || amount <= 0) return;

    summary.totalAssets += amount;

    var accountBucket = classifyAccountBucket_(row, headerMap);
    if (accountBucket) summary[accountBucket] += amount;

    var assetClassBucket = classifyAssetClassBucket_(row, headerMap);
    if (assetClassBucket) summary[assetClassBucket] += amount;

    var annualFee = getNumericValueByHeaderCandidates_(row, headerMap, ['年化费用']);
    if (isFinite(annualFee)) annualFeeSum += annualFee;
  });

  summary.feeRate = summary.totalAssets ? annualFeeSum / summary.totalAssets : 0;
  summary.dailyFee = annualFeeSum / 365;
  return summary;
}

function isRealAssetRow_(name, amount) {
  if (!name || !isFinite(amount)) return false;
  var text = String(name).trim();
  return !!text && text !== '`';
}

function finalizeSnapshotRows_(rows) {
  if (!rows.length) return;

  rows.sort(function(a, b) {
    return normalizeDateTime_(a[0]).getTime() - normalizeDateTime_(b[0]).getTime();
  });

  var previousNav = 1;
  var previousShares = 0;

  for (var i = 0; i < rows.length; i++) {
    var currentTotalAssets = toNumber_(rows[i][REFACTOR_COLUMNS.snapshots.totalAssets - 1]) || 0;
    var cumulativeNetFlow = toNumber_(rows[i][REFACTOR_COLUMNS.snapshots.cumulativeNetFlow - 1]);
    if (!isFinite(cumulativeNetFlow)) cumulativeNetFlow = 0;

    var previousCumulativeNetFlow = i > 0
      ? (toNumber_(rows[i - 1][REFACTOR_COLUMNS.snapshots.cumulativeNetFlow - 1]) || 0)
      : cumulativeNetFlow;
    var dayFlow = cumulativeNetFlow - previousCumulativeNetFlow;
    if (isFinite(toNumber_(rows[i][REFACTOR_COLUMNS.snapshots.netFlow - 1]))) {
      dayFlow = toNumber_(rows[i][REFACTOR_COLUMNS.snapshots.netFlow - 1]);
    }

    var previousTotalAssets = i > 0
      ? (toNumber_(rows[i - 1][REFACTOR_COLUMNS.snapshots.totalAssets - 1]) || 0)
      : currentTotalAssets;
    var dailyReturn = i === 0 ? 0 : (currentTotalAssets - previousTotalAssets - dayFlow);
    if (isFinite(toNumber_(rows[i][REFACTOR_COLUMNS.snapshots.dailyReturn - 1]))) {
      dailyReturn = toNumber_(rows[i][REFACTOR_COLUMNS.snapshots.dailyReturn - 1]);
    }

    var shares;
    var nav;
    if (i === 0) {
      nav = 1;
      shares = currentTotalAssets;
    } else {
      shares = previousShares + dayFlow / previousNav;
      nav = shares ? (currentTotalAssets / shares) : 0;
    }

    previousNav = nav;
    previousShares = shares;

    rows[i][REFACTOR_COLUMNS.snapshots.netFlow - 1] = dayFlow;
    rows[i][REFACTOR_COLUMNS.snapshots.dailyReturn - 1] = dailyReturn;
    rows[i][REFACTOR_COLUMNS.snapshots.cumulativeNetFlow - 1] = cumulativeNetFlow;
    rows[i][REFACTOR_COLUMNS.snapshots.shares - 1] = shares;
    rows[i][REFACTOR_COLUMNS.snapshots.recalculatedNav - 1] = nav;
  }
}

function upsertSnapshotRows_(sheet, incomingRows, replaceAll) {
  ensureSnapshotSheetLayout_(sheet);
  if (!incomingRows.length && !replaceAll) return { snapshotRows: 0, totalRows: Math.max(sheet.getLastRow() - 1, 0) };

  var incomingMap = {};
  incomingRows.forEach(function(row) {
    var date = normalizeDate_(row[REFACTOR_COLUMNS.snapshots.date - 1]);
    if (date) incomingMap[formatDateKey_(date)] = row;
  });

  var rows = [];
  if (!replaceAll && sheet.getLastRow() > 1) {
    rows = sheet
      .getRange(2, 1, sheet.getLastRow() - 1, getSnapshotColumnCount_())
      .getValues()
      .filter(function(row) {
        var date = normalizeDate_(row[REFACTOR_COLUMNS.snapshots.date - 1]);
        return date && !incomingMap.hasOwnProperty(formatDateKey_(date));
      });
  }

  Object.keys(incomingMap).sort().forEach(function(dateKey) {
    rows.push(incomingMap[dateKey]);
  });
  finalizeSnapshotRows_(rows);

  if (sheet.getLastRow() > 1) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getMaxColumns()).clearContent();
  }
  if (rows.length) {
    sheet.getRange(2, 1, rows.length, getSnapshotColumnCount_()).setValues(rows);
  }

  return {
    snapshotRows: incomingRows.length,
    totalRows: rows.length
  };
}

function importRawSnapshotRows_(sheet, rows) {
  ensureSnapshotSheetLayout_(sheet);
  if (sheet.getLastRow() > 1) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getMaxColumns()).clearContent();
  }
  if (rows.length) {
    sheet.getRange(2, 1, rows.length, getSnapshotColumnCount_()).setValues(rows);
  }
  return { importedRows: rows.length };
}

function recalculateSnapshotRows_(sheet) {
  ensureSnapshotSheetLayout_(sheet);
  if (sheet.getLastRow() <= 1) return { snapshotRows: 0, totalRows: 0 };

  var rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, getSnapshotColumnCount_()).getValues();
  finalizeSnapshotRows_(rows);
  sheet.getRange(2, 1, rows.length, getSnapshotColumnCount_()).setValues(rows);
  return {
    snapshotRows: rows.length,
    totalRows: rows.length
  };
}

function getHistoricalMarketValueRows_() {
  var sourceSs = SpreadsheetApp.openById(REFACTOR_SOURCE_SPREADSHEET_ID);
  var sheet = sourceSs.getSheetByName(REFACTOR_SOURCE_MARKET_VALUE_SHEET);
  if (!sheet || sheet.getLastRow() < 2) return [];

  var headerRowIndex = findMarketValueHeaderRow_(sheet);
  if (!headerRowIndex) return [];

  var headers = sheet.getRange(headerRowIndex, 1, 1, sheet.getLastColumn()).getValues()[0];
  var headerIndexMap = buildFlexibleHeaderIndexMap_(headers);
  var dateIndex = getFlexibleHeaderIndex_(headerIndexMap, ['日期']);
  var totalAssetsIndex = getFlexibleHeaderIndex_(headerIndexMap, ['市值W', '市值', '总资产']);
  if (dateIndex < 0 || totalAssetsIndex < 0) {
    throw new Error('旧表“市值记录”缺少“日期”或“市值W”列。');
  }

  var values = sheet.getRange(headerRowIndex + 1, 1, sheet.getLastRow() - headerRowIndex, sheet.getLastColumn()).getValues();
  return values
    .map(function(row) {
      var date = normalizeDateTime_(row[dateIndex]);
      var totalAssets = toNumber_(row[totalAssetsIndex]);
      var summary = createSnapshotSummary_(date);
      summary.totalAssets = totalAssets;
      summary.netFlow = getHistoricalOptionalNumber_(row, headerIndexMap, ['净流入', '当日净现金流']);
      summary.dailyReturn = getHistoricalOptionalNumber_(row, headerIndexMap, ['日收益']);
      summary.broker = getHistoricalNumber_(row, headerIndexMap, ['券商']);
      summary.alipay = getHistoricalNumber_(row, headerIndexMap, ['支付宝']);
      summary.mybank = getHistoricalNumber_(row, headerIndexMap, ['网商银行']);
      summary.equity = getHistoricalNumber_(row, headerIndexMap, ['股权类', '权益类']);
      summary.debt = getHistoricalNumber_(row, headerIndexMap, ['债权类', '固收类']);
      summary.commodity = getHistoricalNumber_(row, headerIndexMap, ['商品']);
      summary.cash = getHistoricalNumber_(row, headerIndexMap, ['现金']);
      summary.feeRate = getHistoricalNumber_(row, headerIndexMap, ['费率']);
      summary.dailyFee = getHistoricalNumber_(row, headerIndexMap, ['日费用']);
      summary.totalProfit = getHistoricalNumber_(row, headerIndexMap, ['总收益']);
      summary.brokerProfit = getHistoricalNumber_(row, headerIndexMap, ['券商收益']);
      summary.alipayFundProfit = getHistoricalNumber_(row, headerIndexMap, ['支付宝基金收益', '支付宝收益']);
      return summary;
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
    var headerIndexMap = buildFlexibleHeaderIndexMap_(values[i]);
    if (getFlexibleHeaderIndex_(headerIndexMap, ['日期']) > -1 &&
        getFlexibleHeaderIndex_(headerIndexMap, ['市值W', '市值', '总资产']) > -1) {
      return i + 1;
    }
  }
  return 0;
}

function getHistoricalNumber_(row, headerIndexMap, headers) {
  return toNumberOrZero_(getHistoricalOptionalNumber_(row, headerIndexMap, headers));
}

function getHistoricalOptionalNumber_(row, headerIndexMap, headers) {
  var index = getFlexibleHeaderIndex_(headerIndexMap, headers);
  return index > -1 ? toNumber_(row[index]) : NaN;
}

function getAssetAmount_(row, headerMap) {
  var amount = getNumericValueByHeaderCandidates_(row, headerMap, ['金额', '总金额', '资产金额', '市值']);
  if (isFinite(amount)) return amount;
  return toNumber_(row[REFACTOR_COLUMNS.assets.amount - 1]);
}

function classifyAccountBucket_(row, headerMap) {
  var institution = getTextByHeaderCandidates_(row, headerMap, ['机构']);
  if (institution === '券商') return 'broker';
  if (institution === '支付宝') return 'alipay';
  if (institution === '网商银行') return 'mybank';

  var text = getCombinedTextByHeaders_(row, headerMap, [
    '机构',
    '名称',
    '平台',
    '账户',
    '归属账户',
    '一级分类',
    '二级分类',
    '资产类型'
  ]);

  if (text.indexOf('支付宝') > -1) return 'alipay';
  if (text.indexOf('网商银行') > -1) return 'mybank';
  if (text.indexOf('券商') > -1 || text.indexOf('证券') > -1) return 'broker';
  return '';
}

function classifyAssetClassBucket_(row, headerMap) {
  var text = getCombinedTextByHeaders_(row, headerMap, [
    '资产类型',
    '一级分类',
    '二级分类',
    '名称'
  ]);

  if (text.indexOf('商品') > -1 || text.indexOf('黄金') > -1) return 'commodity';
  if (text.indexOf('债') > -1 || text.indexOf('固收') > -1) return 'debt';
  if (text.indexOf('现金') > -1 || text.indexOf('活期') > -1 || text.indexOf('存款') > -1) return 'cash';
  if (text.indexOf('股') > -1 || text.indexOf('权益') > -1 || text.indexOf('股票') > -1 || text.indexOf('基金') > -1) return 'equity';
  return '';
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

function getTextByHeaderCandidates_(row, headerMap, candidates) {
  for (var i = 0; i < candidates.length; i++) {
    if (headerMap.hasOwnProperty(candidates[i])) {
      return String(row[headerMap[candidates[i]]] || '').trim();
    }
  }
  return '';
}

function getCombinedTextByHeaders_(row, headerMap, headers) {
  return headers.map(function(header) {
    if (!headerMap.hasOwnProperty(header)) return '';
    return String(row[headerMap[header]] || '').trim();
  }).join(' ');
}

function buildHeaderIndexMap_(headers) {
  var map = {};
  headers.forEach(function(header, index) {
    var key = String(header || '').trim();
    if (key) map[key] = index;
  });
  return map;
}

function buildFlexibleHeaderIndexMap_(headers) {
  var map = {};
  headers.forEach(function(header, index) {
    var key = normalizeHeaderKey_(header);
    if (key && !map.hasOwnProperty(key)) map[key] = index;
  });
  return map;
}

function getFlexibleHeaderIndex_(headerIndexMap, headers) {
  for (var i = 0; i < headers.length; i++) {
    var key = normalizeHeaderKey_(headers[i]);
    if (headerIndexMap.hasOwnProperty(key)) {
      return headerIndexMap[key];
    }
  }
  return -1;
}

function normalizeHeaderKey_(header) {
  return String(header || '').replace(/\s+/g, '').trim();
}

function getHeaderIndexByCandidates_(headerMap, candidates) {
  for (var i = 0; i < candidates.length; i++) {
    if (headerMap.hasOwnProperty(candidates[i])) {
      return headerMap[candidates[i]];
    }
  }
  return -1;
}

function getNormalizedHeaderValues_(sheet) {
  if (!sheet || sheet.getLastColumn() < 1) return [];
  return sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(function(header) {
    return String(header || '').trim();
  });
}

function normalizeDate_(value) {
  var date = parseDateValue_(value);
  if (!date || isNaN(date.getTime())) return null;
  date.setHours(0, 0, 0, 0);
  return date;
}

function normalizeDateTime_(value) {
  var date = parseDateValue_(value);
  if (!date || isNaN(date.getTime())) return null;
  return date;
}

function normalizeDateWithFallback_(value, displayValue) {
  var date = normalizeDate_(value);
  if (date) return date;
  return normalizeDate_(displayValue);
}

function normalizeDateTimeWithFallback_(value, displayValue) {
  var date = normalizeDateTime_(value);
  if (date) return date;
  return normalizeDateTime_(displayValue);
}

function parseDateValue_(value) {
  if (value == null || value === '') return null;
  if (value instanceof Date) return new Date(value);
  if (typeof value === 'number' && isFinite(value)) {
    return new Date(Math.round((value - 25569) * 86400 * 1000));
  }

  var text = String(value).trim();
  if (!text) return null;

  var normalized = text
    .replace(/[年\/.]/g, '-')
    .replace(/月/g, '-')
    .replace(/日/g, '')
    .replace(/\s+/g, ' ')
    .trim();

  var direct = new Date(normalized);
  if (!isNaN(direct.getTime())) return direct;

  var match = normalized.match(/^(\d{4})-(\d{1,2})-(\d{1,2})(?:\s+(\d{1,2})(?::(\d{1,2}))?(?::(\d{1,2}))?)?$/);
  if (!match) return null;

  return new Date(
    Number(match[1]),
    Number(match[2]) - 1,
    Number(match[3]),
    Number(match[4] || 0),
    Number(match[5] || 0),
    Number(match[6] || 0)
  );
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

function toNumberOrZero_(value) {
  var numberValue = toNumber_(value);
  return isFinite(numberValue) ? numberValue : 0;
}

function toNumberWithFallback_(value, displayValue) {
  var numberValue = toNumber_(value);
  if (isFinite(numberValue)) return numberValue;
  return toNumber_(displayValue);
}
