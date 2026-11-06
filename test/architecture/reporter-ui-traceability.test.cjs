const { readFile } = require('node:fs/promises');
const test = require('node:test');
const assert = require('node:assert/strict');

test('reporter UI traceability / every accepted-design scenario has one classified evidence row', async () => {
  const spec = await readFile('openspec/changes/redesign-reporter-ui-demo/specs/reporter-ui-demo/spec.md', 'utf8');
  const matrix = await readFile('docs/reporter-ui-traceability.md', 'utf8');
  const names = [...spec.matchAll(/^#### Scenario: (.+)$/gm)].map((m) => m[1].trim());
  const rows = matrix.split(/\r?\n/).filter((line) => line.startsWith('| ')).slice(2)
    .map((line) => line.split('|').slice(1,-1).map((cell) => cell.trim()));
  assert.equal(rows.length, names.length);
  for (const name of names) {
    const matches = rows.filter((row) => row[0] === name);
    assert.equal(matches.length, 1, name);
    assert.ok(matches[0][1] && matches[0][2], name);
    assert.equal(matches[0][3], 'Automated', name);
  }
});
