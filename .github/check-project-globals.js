const fs = require('fs');
const path = require('path');

function walk(dir) {
  const result = [];
  for (const entry of fs.readdirSync(dir, { withFileTypes: true })) {
    if (entry.name === '.git' || entry.name === 'node_modules') continue;
    const full = path.join(dir, entry.name);
    if (entry.isDirectory()) result.push(...walk(full));
    else if (entry.isFile() && entry.name.endsWith('.gs')) result.push(full);
  }
  return result;
}

function maskStringsAndComments(source) {
  // Good-enough lexical masking for static name checks. Preserve newlines/length
  // so diagnostics still point to the correct source line.
  return source
    .replace(/\/\*[\s\S]*?\*\//g, m => m.replace(/[^\n]/g, ' '))
    .replace(/\/\/[^\n]*/g, m => ' '.repeat(m.length))
    .replace(/`(?:\\.|[^`])*`/gs, m => m.replace(/[^\n]/g, ' '))
    .replace(/'(?:\\.|[^'\\])*'/g, m => ' '.repeat(m.length))
    .replace(/"(?:\\.|[^"\\])*"/g, m => ' '.repeat(m.length));
}

const files = walk('.').sort();
const sources = files.map(file => ({ file, source: fs.readFileSync(file, 'utf8') }));
const all = sources.map(x => x.source).join('\n');
const maskedAll = maskStringsAndComments(all);

const definitions = new Set();
for (const match of maskedAll.matchAll(/\bfunction\s+([A-Za-z_$][\w$]*)\s*\(/g)) definitions.add(match[1]);
for (const match of maskedAll.matchAll(/\b(?:const|let|var)\s+([A-Za-z_$][\w$]*)\b/g)) definitions.add(match[1]);

const errors = [];
for (const { file, source } of sources) {
  const masked = maskStringsAndComments(source);
  const localReported = new Set();

  // Project-private helpers conventionally end in _. This catches missing
  // cross-file functions that JavaScript syntax checking cannot see.
  for (const match of masked.matchAll(/\b([A-Za-z_$][\w$]*_)\s*\(/g)) {
    const name = match[1];
    if (definitions.has(name) || localReported.has(name)) continue;
    localReported.add(name);
    const line = masked.slice(0, match.index).split('\n').length;
    errors.push(`${file}:${line}: undefined project helper ${name}`);
  }

  // Stable UI constants caused a real production failure when Helper.gs and
  // Config.gs were synced partially. Verify every ROLE_COL_/COL_ reference.
  for (const match of masked.matchAll(/\b((?:ROLE_COL|COL)_[A-Z0-9_]+)\b/g)) {
    const name = match[1];
    if (definitions.has(name) || localReported.has(name)) continue;
    localReported.add(name);
    const line = masked.slice(0, match.index).split('\n').length;
    errors.push(`${file}:${line}: undefined project constant ${name}`);
  }
}

if (errors.length) {
  console.error(errors.join('\n'));
  process.exit(1);
}
console.log(`Cross-file globals OK (${files.length} Apps Script files).`);
