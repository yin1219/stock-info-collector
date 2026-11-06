const { readFile } = require('node:fs/promises');
const test = require('node:test');
const assert = require('node:assert/strict');

test('guided onboarding / all scenarios have classified test evidence', async () => {
  const spec = await readFile('openspec/changes/add-guided-onboarding/specs/guided-onboarding/spec.md', 'utf8');
  const matrix = await readFile('docs/onboarding-traceability.md', 'utf8');
  const scenarios = [...spec.matchAll(/^#### Scenario: (.+)$/gm)].map(match => match[1].trim());
  const rows = matrix.split(/\r?\n/).filter(line => line.startsWith('| ')).slice(2)
    .map(line => line.split('|').slice(1, -1).map(cell => cell.trim()));
  assert.equal(rows.length, scenarios.length);
  for (const scenario of scenarios) {
    const matching = rows.filter(row => row[0] === scenario);
    assert.equal(matching.length, 1, scenario);
    assert.match(matching[0][1], /test\//);
    assert.equal(matching[0][2], 'Automated', scenario);
  }
});
