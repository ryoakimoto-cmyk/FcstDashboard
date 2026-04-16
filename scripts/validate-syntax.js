const fs = require('fs');
const path = require('path');
const acorn = require('acorn');

function resolveRoot() {
  const cliRoot = process.argv[2];
  const envRoot = process.env.VALIDATE_SYNTAX_ROOT || process.env.VALIDATE_ROOT;
  const candidate = cliRoot || envRoot;
  return candidate ? path.resolve(candidate) : path.resolve(__dirname, '..');
}

const root = resolveRoot();
const files = fs.readdirSync(root)
  .filter((name) => name.endsWith('.gs'))
  .concat(['js.html', 'index.html', 'mrr-index.html'])
  .map((name) => path.join(root, name))
  .filter((file) => fs.existsSync(file));

function extractSources(filePath, content) {
  if (filePath.endsWith('js.html')) {
    const match = content.match(/<script>([\s\S]*)<\/script>\s*$/);
    if (!match) {
      throw new Error('js.html から script 本体を抽出できません');
    }
    return [match[1]];
  }
  if (filePath.endsWith('.html')) {
    return extractInlineScripts(filePath, content);
  }
  return [content];
}

function extractInlineScripts(filePath, content) {
  validateHtmlTemplate(filePath, content);
  const scripts = [];
  const re = /<script\b([^>]*)>([\s\S]*?)<\/script>/gi;
  let match;
  while ((match = re.exec(content)) !== null) {
    const attrs = match[1] || '';
    const body = match[2] || '';
    if (/\bsrc\s*=/.test(attrs)) continue;
    if (/\btype\s*=\s*["']application\/json["']/i.test(attrs)) continue;
    if (!body.trim()) continue;
    scripts.push(body);
  }
  return scripts;
}

function validateHtmlTemplate(filePath, content) {
  const name = path.basename(filePath);
  ['</html>', '</body>', '</head>', '</title>'].forEach((marker) => {
    if (!content.includes(marker)) {
      throw new Error(`${name}: missing "${marker}"`);
    }
  });

  if (name === 'index.html') {
    [
      'FCST Dashboard',
      '案件リスト',
      '推移',
      'id="main-content"',
      "id=\"gas-dept-config-json\"",
      "id=\"gas-embedded-init-data-json\"",
      "<?!= include('js') ?>"
    ].forEach((marker) => {
      if (!content.includes(marker)) {
        throw new Error(`index.html: missing "${marker}"`);
      }
    });
  }

  if (name === 'mrr-index.html') {
    [
      'MRR進捗ダッシュボード',
      '週次',
      '部署別',
      '読込中',
      'getMrrDashboardData',
      '<h1>',
      '</h1>'
    ].forEach((marker) => {
      if (!content.includes(marker)) {
        throw new Error(`mrr-index.html: missing "${marker}"`);
      }
    });
  }
}

files.forEach((filePath) => {
  const raw = fs.readFileSync(filePath, 'utf8');
  const sources = extractSources(filePath, raw);
  sources.forEach((source, index) => {
    acorn.parse(source, { ecmaVersion: 'latest', sourceType: 'script' });
    if (sources.length > 1) {
      console.log(`ok ${path.basename(filePath)}#${index + 1}`);
    }
  });
  if (sources.length === 1) {
    console.log(`ok ${path.basename(filePath)}`);
  }
});

console.log(`syntax ok: ${files.length} files`);
