interface CoverageCount {
  total: number;
  covered: number;
}

interface FileCoverage {
  lines?: CoverageCount;
  branches?: CoverageCount;
}

interface CoverageInput {
  criticalRules: Record<string, FileCoverage>;
  domainsAndServices: Record<string, FileCoverage>;
}

interface GateResult {
  passed: boolean;
  failures: string[];
}

export function evaluateCoverage(input: CoverageInput): GateResult;
export function evaluateScenarioCoverage(required: string[], classified: string[]): GateResult;
