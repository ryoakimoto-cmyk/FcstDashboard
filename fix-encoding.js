const fs = require('fs');
const path = require('path');
const { TextDecoder } = require('util');

const ROOT = process.cwd();
const UTF8_FATAL = new TextDecoder('utf-8', { fatal: true });
const TARGET_EXTENSIONS = new Set(['.gs', '.html', '.js', '.json', '.md', '.ps1', '.css']);
const IGNORE_DIRS = new Set(['.git', 'node_modules']);
const IGNORE_FILES = new Set([
  '_from_commit_js.html',
  '_js_from_commit.html',
  '__tmp_js_extract.js',
  '_tmp_js_syntax_check.js',
  '_tmp_js_syntax_check_remote.js',
  '.codex-changed-files.json',
  'debug.log'
]);
const SUSPICIOUS_FRAGMENTS = [
  '郢',
  '郢ｧ',
  '隴ｯ莠',
  '陞ｳ蠕',
  '鬩幢ｽｨ',
  '驍ｨ',
  '陝ｷ'
];

function walk(dir, out) {
  fs.readdirSync(dir, { withFileTypes: true }).forEach((entry) => {
    const fullPath = path.join(dir, entry.name);
    if (entry.isDirectory()) {
      if (!IGNORE_DIRS.has(entry.name)) walk(fullPath, out);
      return;
    }
    if (IGNORE_FILES.has(entry.name)) return;
    if (TARGET_EXTENSIONS.has(path.extname(entry.name))) out.push(fullPath);
  });
}

function toRelative(filePath) {
  return path.relative(ROOT, filePath).replace(/\\/g, '/');
}

function validateUtf8(buffer, relativePath, issues) {
  try {
    UTF8_FATAL.decode(buffer);
  } catch (error) {
    issues.push(`${relativePath}: invalid UTF-8`);
    return null;
  }
  return buffer.toString('utf8');
}

const files = [];
walk(ROOT, files);

const errors = [];
const warnings = [];
let normalizedCount = 0;

files.forEach((filePath) => {
  const relativePath = toRelative(filePath);
  const raw = fs.readFileSync(filePath);
  const text = validateUtf8(raw, relativePath, errors);
  if (text === null) return;

  if (text.charCodeAt(0) === 0xfeff) {
    errors.push(`${relativePath}: UTF-8 BOM is not allowed`);
  }

  const normalized = text.replace(/^\ufeff/, '').replace(/\r\n/g, '\n');
  if (normalized !== text) {
    fs.writeFileSync(filePath, normalized, 'utf8');
    normalizedCount += 1;
  }

  if (normalized.includes('\ufffd')) {
    errors.push(`${relativePath}: contains replacement character U+FFFD`);
  }

  const foundFragments = relativePath === 'fix-encoding.js'
    ? []
    : SUSPICIOUS_FRAGMENTS.filter((fragment) => normalized.includes(fragment));
  if (foundFragments.length) {
    warnings.push(`${relativePath}: suspicious mojibake fragments detected (${foundFragments.join(', ')})`);
  }
});

if (warnings.length) {
  console.warn('Encoding warnings:');
  warnings.forEach((warning) => console.warn(`- ${warning}`));
}

if (errors.length) {
  console.error('Encoding check failed:');
  errors.forEach((error) => console.error(`- ${error}`));
  process.exit(1);
}

console.log(`Encoding check passed. ${files.length} files scanned, ${normalizedCount} normalized.`);
