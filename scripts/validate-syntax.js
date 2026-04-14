const fs = require('fs');
const path = require('path');
const acorn = require('acorn');

const root = path.resolve(__dirname, '..');
const files = fs.readdirSync(root)
  .filter((name) => name.endsWith('.gs'))
  .concat(['js.html'])
  .map((name) => path.join(root, name))
  .filter((file) => fs.existsSync(file));

function extractSource(filePath, content) {
  if (filePath.endsWith('js.html')) {
    const match = content.match(/<script>([\s\S]*)<\/script>\s*$/);
    if (!match) {
      throw new Error('js.html から script 本体を抽出できません');
    }
    return match[1];
  }
  return content;
}

files.forEach((filePath) => {
  const raw = fs.readFileSync(filePath, 'utf8');
  const source = extractSource(filePath, raw);
  acorn.parse(source, { ecmaVersion: 'latest', sourceType: 'script' });
  console.log(`ok ${path.basename(filePath)}`);
});

console.log(`syntax ok: ${files.length} files`);
