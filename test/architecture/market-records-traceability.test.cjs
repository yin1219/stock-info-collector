const { readFile } = require("node:fs/promises");
const test = require("node:test");
const assert = require("node:assert/strict");
test("market-records traceability / every scenario has exactly one classified row", async () => {
  const spec = await readFile(
    "openspec/changes/add-market-records/specs/market-records/spec.md",
    "utf8",
  );
  const matrix = await readFile("docs/market-records-traceability.md", "utf8");
  const rows = matrix
    .split(/\r?\n/)
    .filter((line) => line.startsWith("| "))
    .slice(2)
    .map((line) =>
      line
        .split("|")
        .slice(1, -1)
        .map((x) => x.trim()),
    );
  const scenarios = [...spec.matchAll(/^#### Scenario: (.+)$/gm)].map((m) =>
    m[1].trim(),
  );
  assert.equal(rows.length, scenarios.length);
  for (const name of scenarios) {
    const matches = rows.filter((r) => r[0] === name);
    assert.equal(matches.length, 1, name);
    assert.ok(matches[0][1] && matches[0][2]);
    assert.equal(matches[0][3], "Automated");
  }
});
