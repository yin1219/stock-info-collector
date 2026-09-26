import { mkdtemp, mkdir, readFile, rm } from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';

export interface TestPaths {
  root: string;
  userData: string;
  database: string;
  cleanup: () => Promise<void>;
}

export class FakeClock {
  private instant: Date;

  constructor(instant: Date | string) {
    this.instant = new Date(instant);
    if (Number.isNaN(this.instant.valueOf())) {
      throw new Error('FakeClock 需要有效的初始時間');
    }
  }

  now(): Date {
    return new Date(this.instant);
  }

  advance(duration: { milliseconds?: number; seconds?: number; minutes?: number; hours?: number }): void {
    const milliseconds = (duration.milliseconds ?? 0)
      + (duration.seconds ?? 0) * 1_000
      + (duration.minutes ?? 0) * 60_000
      + (duration.hours ?? 0) * 3_600_000;
    this.instant = new Date(this.instant.valueOf() + milliseconds);
  }
}

export class FakeHttpClient {
  readonly requests: string[] = [];

  constructor(private readonly responses: Record<string, unknown> = {}) {}

  async get(url: string): Promise<unknown> {
    assertOfflineUrl(url);
    this.requests.push(url);
    if (!(url in this.responses)) {
      throw new Error(`沒有 HTTP fixture：${url}`);
    }
    const response = this.responses[url];
    if (response instanceof Error) {
      throw response;
    }
    return response;
  }
}

export class FakeProvider<T> {
  calls = 0;

  constructor(private readonly result: T | Error) {}

  async fetch(): Promise<T> {
    this.calls += 1;
    if (this.result instanceof Error) {
      throw this.result;
    }
    return structuredClone(this.result);
  }
}

export class FakeNotificationChannel<T = unknown> {
  readonly deliveries: T[] = [];
  failure: Error | undefined;

  async send(notification: T): Promise<void> {
    if (this.failure) {
      throw this.failure;
    }
    this.deliveries.push(structuredClone(notification));
  }
}

export class FakeCalendar<T extends { summary: string } = { summary: string }> {
  readonly insertions: T[] = [];
  private readonly existingSummaries: Set<string>;

  constructor(options: { existingSummaries?: string[] } = {}) {
    this.existingSummaries = new Set(options.existingSummaries ?? []);
  }

  async hasEvent(query: { summary: string }): Promise<boolean> {
    return this.existingSummaries.has(query.summary);
  }

  async insert(event: T): Promise<void> {
    this.insertions.push(structuredClone(event));
    this.existingSummaries.add(event.summary);
  }
}

export function assertOfflineUrl(url: string): void {
  const parsed = new URL(url);
  if (parsed.protocol !== 'fixture:') {
    throw new Error(`標準測試禁止 live network：${url}`);
  }
}

export function assertTestPath(candidate: string, allowedRoot: string): void {
  const resolvedCandidate = path.resolve(candidate);
  const resolvedRoot = path.resolve(allowedRoot);
  const relative = path.relative(resolvedRoot, resolvedCandidate);
  if (relative === '' || relative === '..' || relative.startsWith(`..${path.sep}`) || path.isAbsolute(relative)) {
    throw new Error(`標準測試禁止正式路徑：${resolvedCandidate}`);
  }
}

export async function createTestPaths(): Promise<TestPaths> {
  const root = await mkdtemp(path.join(os.tmpdir(), 'stock-reporter-test-'));
  const userData = path.join(root, 'userData');
  await mkdir(userData, { recursive: true });
  const database = path.join(root, 'reporter.test.sqlite3');
  assertTestPath(database, root);
  return {
    root,
    userData,
    database,
    cleanup: () => rm(root, { recursive: true, force: true }),
  };
}

export async function loadFixture(relativePath: string): Promise<string> {
  const fixtureRoot = path.resolve(process.cwd(), 'test/fixtures');
  const candidate = path.resolve(fixtureRoot, relativePath);
  const relative = path.relative(fixtureRoot, candidate);
  if (relative === '..' || relative.startsWith(`..${path.sep}`) || path.isAbsolute(relative)) {
    throw new Error(`fixture 路徑越界：${relativePath}`);
  }
  return readFile(candidate, 'utf8');
}
