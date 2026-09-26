const { readFile } = require('node:fs/promises');
const path = require('node:path');
const test = require('node:test');
const assert = require('node:assert/strict');

const {
  buildCalendarEvent,
  parseConferenceHtml,
  planConferenceSync,
} = require('../../src/providers/legacy-conference.cjs');

const fixturePath = path.join(
  __dirname,
  '..',
  'fixtures',
  'conferences',
  'mops-conference.html',
);

test('conference-calendar-sync / legacy parser / parses the current MOPS selectors', async () => {
  const html = await readFile(fixturePath, 'utf8');

  assert.deepEqual(parseConferenceHtml(html), {
    CompId: '2454',
    CompName: '聯發科技股份有限公司',
    Time: '115/09/28 時間： 14 點 30 分 (24小時制)',
    Location: '線上法人說明會',
    Content: '說明公司營運概況',
  });
});

test('conference-calendar-sync / legacy event / preserves title, timezone, duration and reminders', () => {
  const event = buildCalendarEvent({
    CompId: '2454',
    CompName: '聯發科技股份有限公司',
    Time: '115/09/28 時間： 14 點 30 分 (24小時制)',
    Location: '線上法人說明會',
    Content: '說明公司營運概況',
  });

  assert.deepEqual(event, {
    summary: '2454-聯發科技股份有限公司 法說會',
    location: '線上法人說明會',
    description: '說明公司營運概況',
    start: { dateTime: '2026-09-28T06:30:00.000Z', timeZone: 'Asia/Taipei' },
    end: { dateTime: '2026-09-28T08:30:00.000Z', timeZone: 'Asia/Taipei' },
    recurrence: [],
    attendees: [],
    reminders: {
      useDefault: false,
      overrides: [
        { method: 'email', minutes: 1440 },
        { method: 'popup', minutes: 30 },
      ],
    },
  });
});

test('conference-calendar-sync / legacy dedup / queries summary in the exact event window', async () => {
  const queries = [];
  const calendar = {
    async hasEvent(query) {
      queries.push(query);
      return true;
    },
  };

  const plan = await planConferenceSync({
    conferences: [{
      CompId: '2454',
      CompName: '聯發科技股份有限公司',
      Time: '115/09/28 時間： 14 點 30 分 (24小時制)',
      Location: '線上法人說明會',
      Content: '說明公司營運概況',
    }],
    now: new Date('2026-09-25T00:00:00.000Z'),
    calendar,
  });

  assert.deepEqual(queries, [{
    calendarId: 'primary',
    query: '2454-聯發科技股份有限公司 法說會',
    start: new Date('2026-09-28T06:30:00.000Z'),
    end: new Date('2026-09-28T08:30:00.000Z'),
  }]);
  assert.deepEqual(plan, [{ status: 'existing', summary: '2454-聯發科技股份有限公司 法說會' }]);
});

test('conference-calendar-sync / legacy filtering / skips conferences before now', async () => {
  let queried = false;
  const plan = await planConferenceSync({
    conferences: [{
      CompId: '2454',
      CompName: '聯發科技股份有限公司',
      Time: '115/09/20 時間： 14 點 30 分 (24小時制)',
      Location: '線上法人說明會',
      Content: '說明公司營運概況',
    }],
    now: new Date('2026-09-25T00:00:00.000Z'),
    calendar: { async hasEvent() { queried = true; return false; } },
  });

  assert.equal(queried, false);
  assert.deepEqual(plan, [{ status: 'expired', summary: '2454-聯發科技股份有限公司 法說會' }]);
});
