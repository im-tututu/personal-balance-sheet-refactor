const fs = require('fs');
const path = require('path');
const vm = require('vm');

const root = path.resolve(__dirname, '..');
const source = fs.readFileSync(path.join(root, 'refactor.shared.js'), 'utf8');

const context = {
  console,
  Logger: { log() {} },
  SpreadsheetApp: {},
  UrlFetchApp: {},
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

assert(context.toNumber_('1,234.5') === 1234.5, 'toNumber_ should parse comma number');
assert(context.toNumber_('￥9,876') === 9876, 'toNumber_ should parse RMB number');
const formattedKey = context.formatDateKey_(new Date());
assert(/^\d{4}-\d{2}-\d{2}$/.test(formattedKey), 'formatDateKey_ should return yyyy-mm-dd');
assert(context.parseDateKey_('2026-04-18').getFullYear() === 2026, 'parseDateKey_ should parse year');
assert(context.GetDateDiff('2026-04-18', '2026-04-18') === 1, 'GetDateDiff should be inclusive');

const xirr = context.XIRR(
  [-1000, 1100],
  [new Date(2025, 0, 1), new Date(2026, 0, 1)],
  0.1
);
almostEqual(xirr, 0.1, 0.02, 'XIRR smoke test failed');

console.log('calc smoke ok');
