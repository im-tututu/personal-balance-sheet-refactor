const fs = require('fs');
const path = require('path');

const root = path.resolve(__dirname, '..');
const files = [
  'refactor.shared.js',
  'refactor.setup.js',
  'refactor.daily.js',
  'refactor.format.js',
  'refactor.test.js'
];

files.forEach((file) => {
  const fullPath = path.join(root, file);
  const source = fs.readFileSync(fullPath, 'utf8');
  new Function(source);
});

console.log(`syntax ok (${files.length} files)`);
