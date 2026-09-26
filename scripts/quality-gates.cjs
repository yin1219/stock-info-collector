function percent(coverage) {
  if (!coverage || coverage.total === 0) {
    return 0;
  }
  return Math.floor((coverage.covered / coverage.total) * 100);
}

function evaluateScenarioCoverage(required, classified) {
  const classifiedSet = new Set(classified);
  const failures = required
    .filter((scenario) => !classifiedSet.has(scenario))
    .map((scenario) => `未追蹤 Scenario：${scenario}`);
  return { passed: failures.length === 0, failures };
}

function evaluateCoverage(input) {
  const failures = [];
  for (const [file, coverage] of Object.entries(input.criticalRules)) {
    if (percent(coverage.branches) < 100) {
      failures.push(`關鍵規則 branch coverage 未達 100%：${file} (${percent(coverage.branches)}%)`);
    }
  }
  for (const [file, coverage] of Object.entries(input.domainsAndServices)) {
    if (percent(coverage.lines) < 90) {
      failures.push(`domain/services line coverage 未達 90%：${file} (${percent(coverage.lines)}%)`);
    }
    if (percent(coverage.branches) < 85) {
      failures.push(`domain/services branch coverage 未達 85%：${file} (${percent(coverage.branches)}%)`);
    }
  }
  return { passed: failures.length === 0, failures };
}

module.exports = { evaluateCoverage, evaluateScenarioCoverage };
