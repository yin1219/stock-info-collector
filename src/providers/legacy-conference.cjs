const cheerio = require('cheerio');

const DATE_TIME_PATTERN = /(?<ydm>\d*\/\d*\/\d*)(?:\D|\d)*?時間：(?<hour>\d*?)點(?<min>\d*?)分/;
const COMPANY_PATTERN = /<b>公司代號：<\/b>(?<compId>\d+).*?<b>公司名稱：<\/b>(?<compName>.+?)<br>/;

function parseConferenceHtml(html) {
  const $ = cheerio.load(html);
  const contentHtml = $('center').html()?.replace(/\s+/g, '');
  const match = COMPANY_PATTERN.exec(contentHtml);
  if (!match?.groups) {
    return null;
  }

  return {
    CompId: match.groups.compId,
    CompName: match.groups.compName,
    Time: $('center > form > table > tbody > tr:nth-child(1) > td:nth-child(3)').text(),
    Location: $('center > form > table > tbody > tr:nth-child(2) > td:nth-child(2)').text(),
    Content: $('center > form > table > tbody > tr:nth-child(3) > td:nth-child(2)').text(),
  };
}

function conferenceWindow(conference) {
  const compactTime = conference.Time?.replaceAll('\n', '').replaceAll(' ', '');
  const match = DATE_TIME_PATTERN.exec(compactTime);
  if (!match?.groups) {
    throw new Error(`無法解析法說會時間：${conference.CompId}`);
  }

  const [taiwanYear, month, day] = match.groups.ydm.split('/');
  const year = Number(taiwanYear) + 1911;
  const hour = Number(match.groups.hour);
  const minute = Number(match.groups.min);
  const start = new Date(year, Number(month) - 1, Number(day), hour, minute);
  const end = new Date(year, Number(month) - 1, Number(day), hour + 2, minute);
  return { start, end };
}

function buildCalendarEvent(conference) {
  const { start, end } = conferenceWindow(conference);
  return {
    summary: `${conference.CompId}-${conference.CompName} 法說會`,
    location: conference.Location,
    description: conference.Content,
    start: { dateTime: start.toISOString(), timeZone: 'Asia/Taipei' },
    end: { dateTime: end.toISOString(), timeZone: 'Asia/Taipei' },
    recurrence: [],
    attendees: [],
    reminders: {
      useDefault: false,
      overrides: [
        { method: 'email', minutes: 24 * 60 },
        { method: 'popup', minutes: 30 },
      ],
    },
  };
}

async function planConferenceSync({ conferences, now, calendar }) {
  const plan = [];
  for (const conference of conferences) {
    const event = buildCalendarEvent(conference);
    const start = new Date(event.start.dateTime);
    const end = new Date(event.end.dateTime);

    if (start < now) {
      plan.push({ status: 'expired', summary: event.summary });
      continue;
    }

    const exists = await calendar.hasEvent({
      calendarId: 'primary',
      query: event.summary,
      start,
      end,
    });
    plan.push({ status: exists ? 'existing' : 'pending', summary: event.summary });
  }
  return plan;
}

module.exports = {
  buildCalendarEvent,
  parseConferenceHtml,
  planConferenceSync,
};
