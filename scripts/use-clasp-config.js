const fs = require('fs');
const path = require('path');

const projectRoot = path.resolve(__dirname, '..');
const sourceName = process.argv[2];

if (!sourceName) {
  console.error('Usage: node scripts/use-clasp-config.js <config-file>');
  process.exit(1);
}

const sourcePath = path.join(projectRoot, sourceName);
const targetPath = path.join(projectRoot, '.clasp.json');

if (!fs.existsSync(sourcePath)) {
  console.error(`Config not found: ${sourceName}`);
  process.exit(1);
}

const sourceText = fs.readFileSync(sourcePath, 'utf8');
JSON.parse(sourceText);
fs.writeFileSync(targetPath, sourceText);

console.log(`Updated .clasp.json from ${sourceName}`);
