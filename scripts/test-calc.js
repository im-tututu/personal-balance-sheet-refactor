const fs = require('fs');
const path = require('path');
const vm = require('vm');

const root = path.resolve(__dirname, '..');
const source = fs.readFileSync(path.join(root, 'refactor.shared.js'), 'utf8');

const context = {
  console,
  Logger: { log() {} },
  SpreadsheetApp: {},
  Utilities: {
    formatDate(date, timeZone, pattern) {
      if (pattern !== 'yyyy-MM-dd') {
        throw new Error(`Unsupported pattern in local test: ${pattern}`);
      }
      const utc = new Date(date.getTime() + (date.getTimezoneOffset() * 60000));
      const year = utc.getUTCFullYear();
      const month = String(utc.getUTCMonth() + 1).padStart(2, '0');
      const day = String(utc.getUTCDate()).padStart(2, '0');
      return `${year}-${month}-${day}`;
    }
  },
  Session: {
    getScriptTimeZone() {
      return 'Asia/Shanghai';
    }
  }
};

vm.createContext(context);
vm.runInContext(source, context, { filename: 'refactor.shared.js' });

function assert(condition, message) {
  if (!condition) {
    throw new Error(message);
  }
}

function almostEqual(actual, expected, epsilon, message) {
  if (Math.abs(actual - expected) > epsilon) {
    throw new Error(`${message}: expected ${expected}, got ${actual}`);
  }
}

function createFakeSpreadsheet(rows, displayRows) {
  return {
    getSheetByName(name) {
      if (name !== context.REFACTOR_SHEET_NAMES.flows) return null;
      return {
        getLastRow() { return rows.length; },
        getLastColumn() { return rows[0].length; },
        getRange(row, col, numRows, numCols) {
          return {
            getValues() {
              return rows.slice(row - 1, row - 1 + numRows).map(function(r) {
                return r.slice(col - 1, col - 1 + numCols);
              });
            },
            getDisplayValues() {
              return (displayRows || rows).slice(row - 1, row - 1 + numRows).map(function(r) {
                return r.slice(col - 1, col - 1 + numCols).map(function(cell) {
                  return cell instanceof Date ? '2026-01-01 12:00:00' : String(cell);
                });
              });
            }
          };
        }
      };
    }
  };
}

assert(context.toNumber_('1,234.5') === 1234.5, 'toNumber_ should parse comma number');
assert(context.toNumber_('￥9,876') === 9876, 'toNumber_ should parse RMB number');
assert(/^\d{4}-\d{2}-\d{2}$/.test(context.formatDateKey_(new Date())), 'formatDateKey_ should return yyyy-mm-dd');
assert(context.parseDateKey_('2026-04-18').getFullYear() === 2026, 'parseDateKey_ should parse year');

const fakeSs = createFakeSpreadsheet([
  ['类型', '备注', '日期', '金额'],
  ['投资', '', new Date(2026, 0, 1, 12), -100],
  ['投资', '', new Date(2026, 0, 1, 12), 40],
  ['投资', '', new Date(2026, 0, 2, 12), -60],
  ['转账', '', new Date(2026, 0, 2, 12), 999]
]);
const timeline = context.buildInvestmentFlowTimeline_(fakeSs);
almostEqual(context.getCumulativeInvestmentFlowAsOfTimeline_(timeline, new Date(2026, 0, 1, 23, 59, 59)), -60, 1e-9, 'timeline cumulative flow should include same-day investment rows');
almostEqual(context.getCumulativeInvestmentFlowAsOfTimeline_(timeline, new Date(2026, 0, 2, 23, 59, 59)), -120, 1e-9, 'timeline cumulative flow should be the raw running sum');

const fakeSsCashflowPreferred = createFakeSpreadsheet([
  ['类型', '日期', '金额', '现金流', '净流入'],
  ['投资', new Date(2026, 0, 1, 12), 999999, -88, 777777],
  ['投资', new Date(2026, 0, 1, 12), 999999, -12, 777777]
]);
const preferredTimeline = context.buildInvestmentFlowTimeline_(fakeSsCashflowPreferred);
almostEqual(context.getCumulativeInvestmentFlowAsOfTimeline_(preferredTimeline, new Date(2026, 0, 1, 23, 59, 59)), -100, 1e-9, 'investment flow should prefer the cashflow column when present');

const fakeSsDisplayDateFallback = createFakeSpreadsheet(
  [
    ['类型', '日期', '现金流'],
    ['投资', '', '-50,000.25'],
    ['投资', '', '100.25']
  ],
  [
    ['类型', '日期', '现金流'],
    ['投资', '2026/04/18 09:30:00', '-50,000.25'],
    ['投资', '2026年4月18日', '100.25']
  ]
);
const displayFallbackTimeline = context.buildInvestmentFlowTimeline_(fakeSsDisplayDateFallback);
almostEqual(context.getCumulativeInvestmentFlowAsOfTimeline_(displayFallbackTimeline, new Date(2026, 3, 18, 23, 59, 59)), -49900, 1e-9, 'timeline should include rows whose dates come from display text');

const fakeSsUndatedInvestment = createFakeSpreadsheet(
  [
    ['类型', '日期', '现金流'],
    ['投资', '', -1000],
    ['投资', new Date(2026, 0, 2, 12), -200],
    ['转账', '', 999]
  ],
  [
    ['类型', '日期', '现金流'],
    ['投资', '', '-1000'],
    ['投资', '2026-01-02 12:00:00', '-200'],
    ['转账', '', '999']
  ]
);
almostEqual(context.getUndatedInvestmentFlowTotal_(fakeSsUndatedInvestment), -1000, 1e-9, 'undated investment flow should be accumulated separately');
const seedRows = context.buildSnapshotSeedRows_([
  {
    date: new Date(2026, 0, 2),
    totalAssets: 5000,
    netFlow: NaN,
    dailyReturn: NaN,
    broker: 1000,
    alipay: 800,
    mybank: 200,
    equity: 3000,
    debt: 1000,
    commodity: 500,
    cash: 500,
    feeRate: 0.002,
    dailyFee: 3,
    totalProfit: 400,
    brokerProfit: 100,
    alipayFundProfit: 80
  }
], context.buildInvestmentFlowTimeline_(fakeSsUndatedInvestment), context.getUndatedInvestmentFlowTotal_(fakeSsUndatedInvestment));
assert(seedRows[0].length === context.REFACTOR_SNAPSHOT_COLUMNS.length, 'combined snapshot seed rows should use the merged column count');
almostEqual(seedRows[0][4], -1200, 1e-9, 'snapshot seed rows should include undated investment flow in cumulative net flow');
almostEqual(seedRows[0][7], 1000, 1e-9, 'snapshot seed rows should include broker column');
almostEqual(seedRows[0][14], 0.002, 1e-9, 'snapshot seed rows should include fee rate column');
almostEqual(seedRows[0][16], 400, 1e-9, 'snapshot seed rows should include total profit column');

const rows = [
  [new Date(2026, 0, 1), 100, 100, 0, 100, '', ''],
  [new Date(2026, 0, 2), 220, '', '', 200, '', '']
];
context.finalizeSnapshotRows_(rows);
almostEqual(rows[0][5], 100, 1e-9, 'first row shares should equal first row assets');
almostEqual(rows[0][6], 1, 1e-9, 'first row recalculated nav should start from 1');
almostEqual(rows[1][2], 100, 1e-9, 'second row net flow should come from cumulative flow delta');
almostEqual(rows[1][3], 20, 1e-9, 'second row daily return should use snapshot-only assets and flow');
almostEqual(rows[1][5], 200, 1e-9, 'second row shares should use previous snapshot nav');
almostEqual(rows[1][6], 1.1, 1e-9, 'second row recalculated nav should use only snapshot rows');

const rowsWithoutSource = [
  [new Date(2026, 0, 1), 100, '', '', 100, '', ''],
  [new Date(2026, 0, 2), 220, '', '', 200, '', '']
];
context.finalizeSnapshotRows_(rowsWithoutSource);
almostEqual(rowsWithoutSource[0][6], 1, 1e-9, 'first row nav should reset to 1');
almostEqual(rowsWithoutSource[0][5], 100, 1e-9, 'first row shares should equal first row assets');
almostEqual(rowsWithoutSource[1][6], 1.1, 1e-9, 'pure recalculated nav should work from snapshot columns');

console.log('calc smoke ok');
