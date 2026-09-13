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

const files = walk('.').sort();
const sources = files.map(file => ({ file, source: fs.readFileSync(file, 'utf8') }));
const all = sources.map(x => x.source).join('\n');

// Definitions are intentionally collected from raw source. This checker is a
// narrow cross-file guard, not a full JS parser; combined `node --check` already
// validates syntax and duplicate lexical declarations.
const definitions = new Set();
for (const match of all.matchAll(/\bfunction\s+([A-Za-z_$][\w$]*)\s*\(/g)) definitions.add(match[1]);
for (const match of all.matchAll(/\b(?:const|let|var)\s+([A-Za-z_$][\w$]*)\b/g)) definitions.add(match[1]);

const errors = [];
for (const { file, source } of sources) {
  const localReported = new Set();

  // Project-private helpers conventionally end in _. This catches missing
  // cross-file functions such as handleLaunchSnapshot_ that syntax checking
  // cannot detect. References inside comments/strings are rare and, if they
  // occur, should normally name a real project helper anyway.
  for (const match of source.matchAll(/\b([A-Za-z_$][\w$]*_)\s*\(/g)) {
    const name = match[1];
    if (definitions.has(name) || localReported.has(name)) continue;
    localReported.add(name);
    const line = source.slice(0, match.index).split('\n').length;
    errors.push(`${file}:${line}: undefined project helper ${name}`);
  }

  // Stable UI constants caused a real runtime failure after partial file sync.
  for (const match of source.matchAll(/\b((?:ROLE_COL|COL)_[A-Z0-9_]+)\b/g)) {
    const name = match[1];
    if (definitions.has(name) || localReported.has(name)) continue;
    localReported.add(name);
    const line = source.slice(0, match.index).split('\n').length;
    errors.push(`${file}:${line}: undefined project constant ${name}`);
  }
}

if (errors.length) {
  console.error(errors.join('\n'));
  process.exit(1);
}
console.log(`Cross-file globals OK (${files.length} Apps Script files).`);
