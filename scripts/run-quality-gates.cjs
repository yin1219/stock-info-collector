const { readFile } = require('node:fs/promises');
const path = require('node:path');
const { evaluateCoverage } = require('./quality-gates.cjs');

const root = path.resolve(__dirname, '..');
const criticalRulePaths = [
  'src/domain/dedup.ts',
  'src/domain/schedule.ts',
  'src/domain/disclosure-state.ts',
  'src/services/notification-outbox.ts',
  'src/repositories/migrations.ts',
];

async function main() {
  const coveragePath = path.join(root, 'coverage', 'coverage-summary.json');
  let summary;
  try {
    summary = JSON.parse(await readFile(coveragePath, 'utf8'));
  } catch {
    console.error('找不到 coverage/coverage-summary.json，請先執行 npm run coverage');
    process.exitCode = 1;
    return;
  }

  const files = Object.entries(summary).filter(([file]) => file !== 'total');
  const bySuffix = new Map(files.map(([file, metrics]) => [
    file.replaceAll('\\', '/').split('/').slice(-3).join('/'),
    metrics,
  ]));
  const criticalRules = {};
  for (const rulePath of criticalRulePaths) {
    const metrics = bySuffix.get(rulePath);
    if (!metrics) {
      criticalRules[rulePath] = { branches: { total: 1, covered: 0 } };
    } else {
      criticalRules[rulePath] = { branches: metrics.branches };
    }
  }

  const domainsAndServices = Object.fromEntries(files
    .filter(([file]) => /[/\\]src[/\\](domain|services)[/\\]/.test(file))
    .map(([file, metrics]) => [file, { lines: metrics.lines, branches: metrics.branches }]));
  const result = evaluateCoverage({ criticalRules, domainsAndServices });
  if (!result.passed) {
    for (const failure of result.failures) {
      console.error(`FAIL ${failure}`);
    }
    process.exitCode = 1;
    return;
  }
  console.log('Coverage quality gates passed.');
}

void main();
