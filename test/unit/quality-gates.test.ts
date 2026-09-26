import { describe, expect, it } from 'vitest';
import { evaluateCoverage, evaluateScenarioCoverage } from '../../scripts/quality-gates.cjs';

describe('quality gates / OpenSpec scenario coverage', () => {
  it('fails when a required Scenario has no classified test', () => {
    expect(evaluateScenarioCoverage(['one', 'two'], ['one'])).toEqual({
      passed: false,
      failures: ['未追蹤 Scenario：two'],
    });
  });

  it('passes when every Scenario has a classified test', () => {
    expect(evaluateScenarioCoverage(['one', 'two'], ['one', 'two'])).toEqual({
      passed: true,
      failures: [],
    });
  });
});

describe('quality gates / behavior coverage', () => {
  it('requires every critical rule branch to be covered', () => {
    expect(evaluateCoverage({
      criticalRules: { 'src/domain/dedup.ts': { branches: { total: 5, covered: 4 } } },
      domainsAndServices: {},
    })).toEqual({
      passed: false,
      failures: ['關鍵規則 branch coverage 未達 100%：src/domain/dedup.ts (80%)'],
    });
  });

  it('requires at least 90% line and 85% branch coverage for domain and services', () => {
    expect(evaluateCoverage({
      criticalRules: {},
      domainsAndServices: {
        'src/services/monitor.ts': {
          lines: { total: 100, covered: 89 },
          branches: { total: 100, covered: 85 },
        },
      },
    })).toEqual({
      passed: false,
      failures: ['domain/services line coverage 未達 90%：src/services/monitor.ts (89%)'],
    });
  });

  it('passes when all thresholds are met', () => {
    expect(evaluateCoverage({
      criticalRules: { 'src/domain/schedule.ts': { branches: { total: 8, covered: 8 } } },
      domainsAndServices: {
        'src/services/monitor.ts': {
          lines: { total: 100, covered: 92 },
          branches: { total: 100, covered: 86 },
        },
      },
    })).toEqual({ passed: true, failures: [] });
  });
});
