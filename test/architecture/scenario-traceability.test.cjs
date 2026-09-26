const { readdir, readFile } = require('node:fs/promises');
const path = require('node:path');
const test = require('node:test');
const assert = require('node:assert/strict');

const root = path.resolve(__dirname, '..', '..');
const specsRoot = path.join(root, 'openspec', 'changes', 'build-stock-reporter-assistant-v2', 'specs');
const matrixPath = path.join(root, 'docs', 'scenario-traceability.md');

test('OpenSpec traceability / classifies every capability Scenario exactly once', async () => {
  const capabilities = await readdir(specsRoot, { withFileTypes: true });
  const scenarios = [];
  for (const capability of capabilities.filter((entry) => entry.isDirectory())) {
    const specification = await readFile(path.join(specsRoot, capability.name, 'spec.md'), 'utf8');
    for (const match of specification.matchAll(/^#### Scenario: (.+)$/gm)) {
      scenarios.push({ capability: capability.name, name: match[1] });
    }
  }

  const matrix = await readFile(matrixPath, 'utf8');
  const scenarioSection = matrix.split('## Characterization-only invariants')[0];
  const rows = scenarioSection.split(/\r?\n/).filter((line) => line.startsWith('| '));
  const scenarioRows = rows.slice(2).map((line) => line.split('|').slice(1, -1).map((cell) => cell.trim()));
  const classified = scenarioRows.map((cells) => ({ capability: cells[0], name: cells[2], cells }));

  for (const scenario of scenarios) {
    const matches = classified.filter((row) => row.capability === scenario.capability && row.name === scenario.name);
    assert.equal(matches.length, 1, `${scenario.capability}: ${scenario.name} must have exactly one row`);
    assert.ok(matches[0].cells[3], `${scenario.name} must declare a test layer`);
    assert.ok(matches[0].cells[4], `${scenario.name} must declare a unique test or acceptance name`);
    assert.ok(['Planned', 'Automated', 'Manual accepted'].includes(matches[0].cells[5]), `${scenario.name} has invalid status`);
  }

  assert.equal(classified.length, scenarios.length, 'matrix must not contain unclassified or duplicate rows');
  assert.ok(scenarios.length > 0);
});
