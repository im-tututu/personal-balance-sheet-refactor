const fs = require('fs');
const path = require('path');

const configPath = path.resolve(__dirname, '..', '.clasp.json');

if (!fs.existsSync(configPath)) {
  console.error('Missing .clasp.json. Run `npm run gas:env:prod` or `npm run gas:env:dev` first.');
  process.exit(1);
}

const config = JSON.parse(fs.readFileSync(configPath, 'utf8'));
const scriptId = String(config.scriptId || '').trim();

if (!scriptId || scriptId === 'REPLACE_WITH_DEV_SCRIPT_ID') {
  console.error('Invalid scriptId in .clasp.json. Run `npm run gas:env:prod` or create `.clasp.dev.json` and then run `npm run gas:env:dev`.');
  process.exit(1);
}

console.log(`clasp config ok: ${scriptId}`);
