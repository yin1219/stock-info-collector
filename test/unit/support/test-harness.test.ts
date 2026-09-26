import { existsSync } from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { afterEach, describe, expect, it } from 'vitest';
import {
  FakeCalendar,
  FakeClock,
  FakeHttpClient,
  FakeNotificationChannel,
  FakeProvider,
  assertOfflineUrl,
  assertTestPath,
  createTestPaths,
  loadFixture,
  type TestPaths,
} from '../../support';

let paths: TestPaths | undefined;

afterEach(async () => {
  await paths?.cleanup();
  paths = undefined;
});

describe('test infrastructure / deterministic fakes', () => {
  it('controls time and provider responses without system time or network', async () => {
    const clock = new FakeClock('2026-09-25T10:30:00+08:00');
    const http = new FakeHttpClient({
      'fixture://mops/today': { status: 200, body: 'fixture body' },
    });
    const provider = new FakeProvider([{ id: 'event-1' }]);

    expect(clock.now().toISOString()).toBe('2026-09-25T02:30:00.000Z');
    clock.advance({ minutes: 30 });
    expect(clock.now().toISOString()).toBe('2026-09-25T03:00:00.000Z');
    await expect(http.get('fixture://mops/today')).resolves.toEqual({
      status: 200,
      body: 'fixture body',
    });
    await expect(provider.fetch()).resolves.toEqual([{ id: 'event-1' }]);
  });

  it('records fake notification and Calendar interactions', async () => {
    const notifications = new FakeNotificationChannel();
    const calendar = new FakeCalendar({ existingSummaries: ['2454-聯發科 法說會'] });

    await notifications.send({ title: '測試通知', body: '不會送到 Windows' });
    await expect(calendar.hasEvent({ summary: '2454-聯發科 法說會' })).resolves.toBe(true);
    await calendar.insert({ summary: '2330-台積電 法說會' });

    expect(notifications.deliveries).toHaveLength(1);
    expect(calendar.insertions).toEqual([{ summary: '2330-台積電 法說會' }]);
  });
});

describe('test infrastructure / safety guards', () => {
  it('rejects live network URLs and paths outside the temporary test root', async () => {
    expect(() => assertOfflineUrl('https://mops.twse.com.tw/mops/web/index')).toThrow(
      /標準測試禁止 live network/,
    );
    expect(() => assertOfflineUrl('fixture://mops/today')).not.toThrow();

    paths = await createTestPaths();
    expect(paths.root.startsWith(os.tmpdir())).toBe(true);
    expect(paths.database).toBe(path.join(paths.root, 'reporter.test.sqlite3'));
    expect(existsSync(paths.userData)).toBe(true);
    expect(() => assertTestPath(paths!.database, paths!.root)).not.toThrow();
    expect(() => assertTestPath(path.resolve('reporter.sqlite3'), paths!.root)).toThrow(
      /標準測試禁止正式路徑/,
    );
  });

  it('loads fixtures only from the fixture sandbox', async () => {
    await expect(loadFixture('conferences/mops-conference.html')).resolves.toContain(
      '公司代號',
    );
    await expect(loadFixture('../package.json')).rejects.toThrow(/fixture 路徑越界/);
  });
});
