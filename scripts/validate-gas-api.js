// Validates that every google.script.run.XXX() call in HTML files
// has a corresponding top-level function in a .gs file.
// Usage: node scripts/validate-gas-api.js
const fs = require('fs');
const path = require('path');

const ROOT = path.join(__dirname, '..');
const RESERVED = new Set([
  'withSuccessHandler',
  'withFailureHandler',
  'withUserObject',
  'run',
]);

function extractCalledFunctions(htmlSrc) {
  const called = new Set();
  // Find every occurrence of "google.script.run" and walk forward collecting .NAME( tokens
  // until we hit a terminator (;) or a non-chaining character.
  const runToken = 'google.script.run';
  let idx = 0;
  while ((idx = htmlSrc.indexOf(runToken, idx)) !== -1) {
    let pos = idx + runToken.length;
    const names = [];
    const chainRe = /\s*\.([a-zA-Z_][a-zA-Z0-9_]*)\s*\(/y;
    while (true) {
      chainRe.lastIndex = pos;
      const m = chainRe.exec(htmlSrc);
      if (!m) break;
      names.push(m[1]);
      // Skip the balanced parenthesis group
      const openParen = chainRe.lastIndex - 1;
      let depth = 1;
      let p = openParen + 1;
      let inStr = null;
      while (p < htmlSrc.length && depth > 0) {
        const ch = htmlSrc[p];
        if (inStr) {
          if (ch === '\\') { p += 2; continue; }
          if (ch === inStr) inStr = null;
        } else if (ch === '"' || ch === "'" || ch === '`') {
          inStr = ch;
        } else if (ch === '(') depth++;
        else if (ch === ')') depth--;
        p++;
      }
      pos = p;
    }
    for (let i = names.length - 1; i >= 0; i--) {
      if (!RESERVED.has(names[i])) {
        called.add(names[i]);
        break;
      }
    }
    idx = pos > idx ? pos : idx + runToken.length;
  }
  return called;
}

function extractDefinedFunctions(gsSrc) {
  const defined = new Set();
  const pattern = /^function\s+([a-zA-Z_][a-zA-Z0-9_]*)\s*\(/gm;
  let m;
  while ((m = pattern.exec(gsSrc)) !== null) {
    const name = m[1];
    if (!name.endsWith('_')) defined.add(name);
  }
  return defined;
}

function main() {
  const htmlFiles = fs.readdirSync(ROOT).filter((f) => f.endsWith('.html'));
  const gsFiles = fs.readdirSync(ROOT).filter((f) => f.endsWith('.gs'));

  const called = new Set();
  for (const f of htmlFiles) {
    const src = fs.readFileSync(path.join(ROOT, f), 'utf8');
    for (const name of extractCalledFunctions(src)) called.add(name);
  }

  const defined = new Set();
  for (const f of gsFiles) {
    const src = fs.readFileSync(path.join(ROOT, f), 'utf8');
    for (const name of extractDefinedFunctions(src)) defined.add(name);
  }

  const missing = [...called].filter((n) => !defined.has(n)).sort();
  if (missing.length > 0) {
    console.error('Missing GAS functions referenced by google.script.run:');
    for (const name of missing) console.error('  - ' + name);
    console.error('\nDefine these functions in a .gs file (without trailing underscore) before pushing.');
    process.exit(1);
  }

  const calledList = [...called].sort().join(', ');
  console.log('All good: every google.script.run call resolves to a GAS function.');
  console.log('Validated calls: ' + (calledList || '(none)'));
  process.exit(0);
}

main();
