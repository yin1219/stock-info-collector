import { useCallback, useEffect, useRef, useState, type FormEvent } from 'react';
import type { GoogleCalendarStatus, LoginStartupSettings, ReporterRoute, ScheduleSettings } from '../preload/api';

interface WatchlistRow {
  id: string;
  companyId: string;
  market: 'TWSE' | 'TPEX';
  stockCode: string;
  name: string;
  active: boolean;
  category: string;
  notes: string;
}

interface ImportReport {
  summary: { added: number; skipped: number; failed: number };
  items: Array<{ stockCode: string; status: string; reason: string }>;
}

interface DisclosureRow {
  id: string;
  market: 'TWSE' | 'TPEX';
  stockCode: string;
  companyName: string | null;
  disclosureDate: string;
  brokerCode: string;
  isWatched: boolean;
  contentJson: string;
}

interface DisclosureMonitorStatus {
  status: 'active' | 'complete' | 'degraded' | 'stale' | 'failed';
  finishedAt: string | null;
  errorMessage: string | null;
  markets: Record<string, { status: string; dataDate: string | null; recordCount: number; retryAt: string | null; errorMessage?: string }>;
}

interface MaterialEventRow {
  id: string;
  companyId: string;
  market: 'TWSE' | 'TPEX';
  stockCode: string;
  companyName: string;
  title: string;
  content: string;
  source: string;
  sourceUrl: string;
  publishedAt: string;
  discoveredAt: string;
  readAt: string | null;
  eventType: 'announcement' | 'correction' | 'supplement';
  revisionOf: string | null;
}

interface MaterialEventDetail extends MaterialEventRow {
  original: MaterialEventRow | null;
  relatedRevisions: MaterialEventRow[];
}

interface MaterialMonitorStatus {
  status: 'active' | 'complete' | 'degraded' | 'stale' | 'failed';
  finishedAt: string | null;
  errorMessage: string | null;
  sources: Array<{ source: string; status: string; dataDate: string | null; errorMessage: string | null }>;
}

interface ConferenceRow {
  id: string;
  stockCode: string;
  companyName: string;
  startsAt: string;
  location: string;
  content: string;
  sourceUrl?: string;
  syncStatus: 'pending' | 'existing' | 'synced' | 'failed';
  lastError: string | null;
}

type ScheduleJobStatus = { enabled: boolean; status: string; lastSuccessAt: string | null; lastResult: string | null; errorMessage: string | null; nextRunAt: string | null };
type ScheduleStatus = { available?: boolean; material: ScheduleJobStatus; disclosure: ScheduleJobStatus };
type Page = 'overview' | 'watchlist' | 'disclosures' | 'material-events' | 'conferences' | 'settings';

function disclosureStatusLabel(status: string): string {
  switch (status) {
    case 'complete': return '資料完整';
    case 'empty_success': return '本日無揭露';
    case 'stale': return '官方尚未更新';
    case 'failed': return '來源失敗';
    case 'active': return '檢查中';
    default: return status;
  }
}

function disclosureEmptyMessage(
  monitor: DisclosureMonitorStatus | null,
  date: string,
  market: 'all' | 'TWSE' | 'TPEX',
): string {
  const selectedMarkets = market === 'all' ? ['TWSE', 'TPEX'] : [market];
  const states = selectedMarkets.map((source) => monitor?.markets[source]);
  if (isConfirmedEmpty(monitor, date, market)) {
    return '當日無個股達違約資訊揭露標準';
  }
  if (states.some((state) => state?.status === 'failed' || state?.status === 'stale')
    || monitor?.status === 'failed' || monitor?.status === 'stale') {
    return '來源檢查失敗或官方資料尚未更新，不能判定為零筆。';
  }
  return '此日期目前沒有可顯示的違約交割資料，尚未確認是否為零筆。';
}

function isConfirmedEmpty(monitor: DisclosureMonitorStatus | null, date: string, market: 'all' | 'TWSE' | 'TPEX'): boolean {
  const selectedMarkets = market === 'all' ? ['TWSE', 'TPEX'] : [market];
  return selectedMarkets.every((source) => monitor?.markets[source]?.status === 'empty_success'
    && monitor.markets[source].dataDate === date);
}

function disclosureCountLabel(count: number, monitor: DisclosureMonitorStatus | null, date: string, market: 'all' | 'TWSE' | 'TPEX'): string {
  if (count > 0 || isConfirmedEmpty(monitor, date, market)) return `${count} 筆`;
  return monitor ? '未確認' : '尚未檢查';
}

function formatTaipeiDateTime(value: string): string {
  const instant = new Date(value);
  if (Number.isNaN(instant.getTime())) return '時間待確認';
  return new Intl.DateTimeFormat('zh-TW', {
    timeZone: 'Asia/Taipei', year: 'numeric', month: '2-digit', day: '2-digit',
    hour: '2-digit', minute: '2-digit', hourCycle: 'h23',
  }).format(instant);
}

function materialSourceName(source: string): string {
  if (source === 'TWSE-reconciliation') return '上市結構化每日對帳';
  if (source === 'TPEX-reconciliation') return '上櫃結構化每日對帳';
  return source;
}

function monitorSourceStatusLabel(status: string): string {
  if (status === 'complete') return '資料完整';
  if (status === 'degraded') return '部分降級';
  if (status === 'stale') return '官方尚未更新';
  if (status === 'failed') return '來源失敗';
  return status;
}

function taipeiDateToday(): string {
  const parts = new Intl.DateTimeFormat('en-CA', { timeZone: 'Asia/Taipei', year: 'numeric', month: '2-digit', day: '2-digit' }).formatToParts(new Date());
  const values = Object.fromEntries(parts.map(({ type, value }) => [type, value]));
  return `${values.year}-${values.month}-${values.day}`;
}

export function App() {
  const [rows, setRows] = useState<WatchlistRow[]>([]);
  const [query, setQuery] = useState('');
  const [activeOnly, setActiveOnly] = useState(false);
  const [market, setMarket] = useState<'TWSE' | 'TPEX'>('TWSE');
  const [stockCode, setStockCode] = useState('');
  const [error, setError] = useState('');
  const [configText, setConfigText] = useState('');
  const [importReport, setImportReport] = useState<ImportReport | null>(null);
  const [busy, setBusy] = useState(false);
  const [page, setPage] = useState<Page>('watchlist');
  const [disclosures, setDisclosures] = useState<DisclosureRow[]>([]);
  const [disclosureMonitorStatus, setDisclosureMonitorStatus] = useState<DisclosureMonitorStatus | null>(null);
  const [disclosureDate, setDisclosureDate] = useState(taipeiDateToday);
  const [disclosureMarket, setDisclosureMarket] = useState<'all' | 'TWSE' | 'TPEX'>('all');
  const [materialEvents, setMaterialEvents] = useState<MaterialEventRow[]>([]);
  const [materialQuery, setMaterialQuery] = useState('');
  const [unreadOnly, setUnreadOnly] = useState(false);
  const [showAllMaterial, setShowAllMaterial] = useState(false);
  const [expandedMaterialRowId, setExpandedMaterialRowId] = useState<string | null>(null);
  const [selectedMaterialEvent, setSelectedMaterialEvent] = useState<MaterialEventDetail | null>(null);
  const materialDetailRequest = useRef(0);
  const [notificationEventIds, setNotificationEventIds] = useState<string[] | null>(null);
  const [materialMonitorStatus, setMaterialMonitorStatus] = useState<MaterialMonitorStatus | null>(null);
  const [scheduleSettings, setScheduleSettings] = useState<ScheduleSettings | null>(null);
  const [scheduleStatus, setScheduleStatus] = useState<ScheduleStatus | null>(null);
  const [scheduleMessage, setScheduleMessage] = useState('');
  const [scheduleMessageKind, setScheduleMessageKind] = useState<'material' | 'disclosure' | null>(null);
  const [dataExportMessage, setDataExportMessage] = useState('');
  const [dataExportBusy, setDataExportBusy] = useState(false);
  const [runningSchedule, setRunningSchedule] = useState<'material' | 'disclosure' | null>(null);
  const [loginStartupSettings, setLoginStartupSettings] = useState<LoginStartupSettings | null>(null);
  const [overviewDisclosureCount, setOverviewDisclosureCount] = useState<number | null>(null);
  const [overviewMaterialCount, setOverviewMaterialCount] = useState<number | null>(null);
  const [overviewMaterialEvents, setOverviewMaterialEvents] = useState<MaterialEventRow[]>([]);
  const [overviewConferences, setOverviewConferences] = useState<ConferenceRow[]>([]);
  const [overviewScheduleStatus, setOverviewScheduleStatus] = useState<ScheduleStatus | null>(null);
  const [googleCalendarStatus, setGoogleCalendarStatus] = useState<GoogleCalendarStatus | null>(null);
  const [conferenceRows, setConferenceRows] = useState<ConferenceRow[]>([]);
  const [conferenceBusy, setConferenceBusy] = useState(false);
  const [conferenceMessage, setConferenceMessage] = useState('');

  useEffect(() => {
    if (page !== 'settings') return;
    void Promise.all([window.reporterApi.getScheduleSettings(), window.reporterApi.getScheduleStatus(), window.reporterApi.getLoginStartupSettings()])
      .then(([settings, status, startup]) => { setScheduleSettings(settings); setScheduleStatus(status as ScheduleStatus); setLoginStartupSettings(startup); })
      .catch((cause: unknown) => setError(cause instanceof Error ? cause.message : String(cause)));
  }, [page]);

  useEffect(() => {
    if (page !== 'overview') return;
    void Promise.all([
      window.reporterApi.listDisclosures({ disclosureDate: taipeiDateToday() }),
      window.reporterApi.listMaterialEvents({ unreadOnly: true, watchedOnly: true }),
      window.reporterApi.getMaterialMonitorStatus(),
      window.reporterApi.getDisclosureMonitorStatus(),
      window.reporterApi.listConferences(),
      window.reporterApi.getScheduleStatus(),
    ]).then(([disclosureRows, materialRows, materialStatus, disclosureStatus, conferenceRows, scheduleStatus]) => {
      setOverviewDisclosureCount((disclosureRows as DisclosureRow[]).length);
      setOverviewMaterialCount((materialRows as MaterialEventRow[]).length);
      setOverviewMaterialEvents(materialRows as MaterialEventRow[]);
      const upcomingConferences = (conferenceRows as ConferenceRow[]).filter((conference) => {
        const startsAt = new Date(conference.startsAt).getTime();
        return startsAt >= Date.now() && startsAt <= Date.now() + 14 * 24 * 60 * 60 * 1000;
      });
      setOverviewConferences(upcomingConferences);
      setOverviewScheduleStatus(scheduleStatus as ScheduleStatus);
      setMaterialMonitorStatus(materialStatus as MaterialMonitorStatus | null);
      setDisclosureMonitorStatus(disclosureStatus as DisclosureMonitorStatus | null);
    }).catch((cause: unknown) => setError(cause instanceof Error ? cause.message : String(cause)));
  }, [page]);

  useEffect(() => {
    if (page !== 'conferences') return;
    void Promise.all([window.reporterApi.getGoogleCalendarStatus(), window.reporterApi.listConferences()])
      .then(([status, rows]) => { setGoogleCalendarStatus(status); setConferenceRows(rows as ConferenceRow[]); })
      .catch((cause: unknown) => setError(cause instanceof Error ? cause.message : String(cause)));
  }, [page]);

  const refresh = useCallback(async () => {
    const result = await window.reporterApi.listWatchlist({ query, activeOnly });
    setRows(result as WatchlistRow[]);
  }, [query, activeOnly]);

  useEffect(() => {
    if (page !== 'material-events') return;
    void window.reporterApi.getMaterialMonitorStatus()
      .then((result) => setMaterialMonitorStatus(result as MaterialMonitorStatus | null))
      .catch((cause: unknown) => setError(cause instanceof Error ? cause.message : String(cause)));
    void window.reporterApi.listMaterialEvents({
      query: materialQuery, unreadOnly, ...(showAllMaterial ? {} : { watchedOnly: true }),
      ...(notificationEventIds ? { eventIds: notificationEventIds } : {}),
    }).then((result) => {
      setMaterialEvents(result as MaterialEventRow[]);
    }).catch((cause: unknown) => setError(cause instanceof Error ? cause.message : String(cause)));
  }, [page, materialQuery, unreadOnly, showAllMaterial, notificationEventIds]);

  useEffect(() => window.reporterApi.onNotificationRoute((route: ReporterRoute) => {
    setError('');
    if (route.type === 'disclosure-list') {
      setNotificationEventIds(null);
      setPage('disclosures');
      void loadDisclosures(disclosureDate, disclosureMarket);
      return;
    }
    setPage('material-events');
    setMaterialQuery('');
    setUnreadOnly(false);
    if (route.type === 'source-status') {
      setNotificationEventIds(null);
      setSelectedMaterialEvent(null);
      setExpandedMaterialRowId(null);
      return;
    }
    if (route.type === 'event-list') {
      setNotificationEventIds(route.eventIds);
      setSelectedMaterialEvent(null);
      setExpandedMaterialRowId(null);
      return;
    }
    setNotificationEventIds([route.eventId]);
    setExpandedMaterialRowId(route.eventId);
    void window.reporterApi.getMaterialEvent(route.eventId).then((result) => setSelectedMaterialEvent(result as MaterialEventDetail));
  }), []);

  useEffect(() => {
    void refresh().catch((cause: unknown) => setError(cause instanceof Error ? cause.message : String(cause)));
  }, [refresh]);

  async function addCompany(event: FormEvent<HTMLFormElement>) {
    event.preventDefault();
    const normalizedCode = stockCode.trim();
    if (!/^\d{4,6}$/.test(normalizedCode)) {
      setError('股票代號必須是 4 至 6 位數字');
      return;
    }
    setBusy(true);
    setError('');
    try {
      await window.reporterApi.addWatchlist({ market, stockCode: normalizedCode });
      setStockCode('');
      await refresh();
    } catch (cause) {
      setError(cause instanceof Error ? cause.message : String(cause));
    } finally {
      setBusy(false);
    }
  }

  async function previewImport() {
    setError('');
    try {
      setImportReport(await window.reporterApi.previewWatchlistImport(configText) as ImportReport);
    } catch (cause) {
      setError(cause instanceof Error ? cause.message : String(cause));
    }
  }

  async function applyImport() {
    setBusy(true);
    setError('');
    try {
      setImportReport(await window.reporterApi.applyWatchlistImport(configText) as ImportReport);
      await refresh();
    } catch (cause) {
      setError(cause instanceof Error ? cause.message : String(cause));
    } finally {
      setBusy(false);
    }
  }

  async function loadDisclosures(date = disclosureDate, selectedMarket = disclosureMarket) {
    setError('');
    try {
      const [result, status] = await Promise.all([window.reporterApi.listDisclosures({
        disclosureDate: date,
        ...(selectedMarket === 'all' ? {} : { market: selectedMarket }),
      }), window.reporterApi.getDisclosureMonitorStatus()]);
      setDisclosures(result as DisclosureRow[]);
      setDisclosureMonitorStatus(status as DisclosureMonitorStatus | null);
    } catch (cause) {
      setError(cause instanceof Error ? cause.message : String(cause));
    }
  }

  async function runSchedule(kind: 'material' | 'disclosure') {
    const title = kind === 'material' ? '重大訊息' : '違約交割';
    setRunningSchedule(kind);
    setScheduleMessageKind(kind);
    setScheduleMessage(`正在檢查${title}…`);
    try {
      const result = await window.reporterApi.runScheduledCheck(kind) as { status: string };
      const latest = await window.reporterApi.getScheduleStatus() as ScheduleStatus;
      setScheduleStatus(latest);
      const job = latest[kind];
      setScheduleMessage(result.status === 'already-running' ? '此工作目前仍在執行，未重複啟動。'
        : job.status === 'failed' ? `${title}檢查失敗：${job.errorMessage ?? '來源錯誤'}`
          : job.status === 'stale' ? `${title}官方資料尚未更新。`
            : job.status === 'degraded' ? `${title}部分來源檢查失敗，請查看最新狀態。`
              : `${title}檢查完成，請查看最新狀態。`);
    } catch (cause) {
      setScheduleMessage(cause instanceof Error ? cause.message : String(cause));
    } finally {
      setRunningSchedule(null);
    }
  }

  async function saveSchedule() {
    if (!scheduleSettings) return;
    setError('');
    try {
      setScheduleSettings(await window.reporterApi.saveScheduleSettings(scheduleSettings));
      setScheduleMessage('排程設定已儲存。');
      setScheduleStatus(await window.reporterApi.getScheduleStatus() as ScheduleStatus);
    } catch (cause) {
      setError(cause instanceof Error ? cause.message : String(cause));
    }
  }

  async function saveLoginStartup() {
    if (!loginStartupSettings) return;
    setError('');
    try {
      setLoginStartupSettings(await window.reporterApi.saveLoginStartupSettings(loginStartupSettings));
      setScheduleMessage('登入啟動設定已儲存。');
    } catch (cause) {
      setError(cause instanceof Error ? cause.message : String(cause));
    }
  }

  async function exportLocalData() {
    setDataExportBusy(true);
    setDataExportMessage('');
    try {
      const result = await window.reporterApi.exportUserData();
      setDataExportMessage(result.status === 'cancelled' ? '已取消資料匯出。' : `資料已匯出至 ${result.destination}`);
    } catch (cause) {
      setDataExportMessage(cause instanceof Error ? cause.message : '資料匯出失敗，請查看診斷紀錄。');
    } finally {
      setDataExportBusy(false);
    }
  }

  async function importGoogleCredentials() {
    setError('');
    try {
      const imported = await window.reporterApi.importGoogleCredentials();
      if (imported) {
        setGoogleCalendarStatus(await window.reporterApi.getGoogleCalendarStatus());
        setConferenceMessage('OAuth 設定已安全匯入。');
      }
    } catch (cause) {
      setError(cause instanceof Error ? cause.message : String(cause));
    }
  }

  async function migrateLegacyGoogleToken() {
    setError('');
    try {
      const migrated = await window.reporterApi.migrateLegacyGoogleToken();
      if (migrated) {
        setGoogleCalendarStatus(await window.reporterApi.getGoogleCalendarStatus());
        setConferenceMessage('舊版 token 已安全遷移；原始檔案仍保留。');
      }
    } catch (cause) {
      setError(cause instanceof Error ? cause.message : String(cause));
    }
  }

  async function connectGoogleCalendar() {
    setConferenceBusy(true);
    setError('');
    setConferenceMessage('請在系統瀏覽器完成 Google 授權；完成後此頁會更新連線狀態。');
    try {
      setGoogleCalendarStatus(await window.reporterApi.connectGoogleCalendar());
      setConferenceMessage('Google Calendar 已連線。');
    } catch (cause) {
      setConferenceMessage('');
      setError(cause instanceof Error ? cause.message : String(cause));
      setGoogleCalendarStatus(await window.reporterApi.getGoogleCalendarStatus());
    } finally {
      setConferenceBusy(false);
    }
  }

  async function disconnectGoogleCalendar() {
    setConferenceBusy(true);
    setError('');
    try {
      setGoogleCalendarStatus(await window.reporterApi.disconnectGoogleCalendar());
      setConferenceMessage('已中斷 Google Calendar 連線。');
    } catch (cause) {
      setError(cause instanceof Error ? cause.message : String(cause));
    } finally {
      setConferenceBusy(false);
    }
  }

  async function syncConferences() {
    setConferenceBusy(true);
    setError('');
    try {
      const result = await window.reporterApi.syncConferences() as { status: string; results: Array<{ status: string }> };
      setGoogleCalendarStatus(await window.reporterApi.getGoogleCalendarStatus());
      setConferenceRows(await window.reporterApi.listConferences() as ConferenceRow[]);
      setConferenceMessage(result.status === 'reauthorization-required' ? 'Google 憑證失效，請重新授權。' : `法說會同步完成，處理 ${result.results.length} 筆。`);
    } catch (cause) {
      setError(cause instanceof Error ? cause.message : String(cause));
    } finally {
      setConferenceBusy(false);
    }
  }

  async function openMaterialEvent(id: string, anchorId = id) {
    if (expandedMaterialRowId === anchorId && selectedMaterialEvent?.id === id) {
      materialDetailRequest.current += 1;
      setExpandedMaterialRowId(null);
      setSelectedMaterialEvent(null);
      return;
    }
    const requestId = ++materialDetailRequest.current;
    setExpandedMaterialRowId(anchorId);
    setSelectedMaterialEvent(null);
    setError('');
    try {
      const detail = await window.reporterApi.getMaterialEvent(id) as MaterialEventDetail;
      if (requestId === materialDetailRequest.current) setSelectedMaterialEvent(detail);
    } catch (cause) {
      if (requestId === materialDetailRequest.current) {
        setExpandedMaterialRowId(null);
        setError(cause instanceof Error ? cause.message : String(cause));
      }
    }
  }

  async function markSelectedMaterialEventRead() {
    if (!selectedMaterialEvent) return;
    await window.reporterApi.markMaterialEventRead(selectedMaterialEvent.id);
    setSelectedMaterialEvent((current) => current ? { ...current, readAt: new Date().toISOString() } : current);
    await window.reporterApi.listMaterialEvents({ query: materialQuery, unreadOnly,
      ...(showAllMaterial ? {} : { watchedOnly: true }), ...(notificationEventIds ? { eventIds: notificationEventIds } : {}),
    }).then((result) => setMaterialEvents(result as MaterialEventRow[]));
  }

  return (
    <main className="app-shell">
      <div className="sidebar">
      <header className="app-header">
        <div>
          <p className="eyebrow">STOCK REPORTER ASSISTANT</p>
          <h1>股市記者小幫手</h1>
        </div>
        <span className="version">v2 · 本機資料</span>
      </header>

      <nav aria-label="主要功能" aria-orientation="horizontal" className="main-nav" role="tablist">
        <button aria-selected={page === 'overview'} role="tab" type="button" onClick={() => setPage('overview')}>總覽</button>
        <button aria-selected={page === 'watchlist'} role="tab" type="button" onClick={() => setPage('watchlist')}>關注公司</button>
        <button aria-selected={page === 'material-events'} role="tab" type="button" onClick={() => { setNotificationEventIds(null); setPage('material-events'); }}>重大訊息</button>
        <button aria-selected={page === 'disclosures'} role="tab" type="button" onClick={() => { setPage('disclosures'); void loadDisclosures(); }}>違約交割</button>
        <button aria-selected={page === 'conferences'} role="tab" type="button" onClick={() => setPage('conferences')}>法說會</button>
        <button aria-selected={page === 'settings'} role="tab" type="button" onClick={() => setPage('settings')}>設定</button>
      </nav>
      <aside aria-label="背景執行狀態" className="sidebar-status">
        <strong>背景監控服務</strong>
        <p>{scheduleStatus?.available === false || overviewScheduleStatus?.available === false ? '隔離測試模式：背景監控已停用。' : '關閉視窗後會留在系統匣執行。'}</p>
        <small>來源健康與下次排程請見總覽和設定。</small>
      </aside>
      </div>

      {page === 'overview' ? <section aria-labelledby="overview-title" className="panel overview-panel">
        <div className="section-heading"><div><p className="eyebrow">TODAY</p><h2 id="overview-title">今日監控總覽</h2></div><span className="count">{taipeiDateToday()}</span></div>
        {error && <p role="alert" className="error-message">{error}</p>}
        <div className="overview-grid">
          <article className="overview-card"><h3>關注清單</h3><p className="overview-value">關注公司 {rows.filter((row) => row.active).length} 家</p><button className="text-button" type="button" onClick={() => setPage('watchlist')}>管理關注公司</button></article>
          <article className="overview-card"><h3>未讀重大訊息</h3><p className="overview-value">{overviewMaterialCount === null ? '載入中…' : `${overviewMaterialCount} 筆`}</p><p className="muted">來源：{materialMonitorStatus ? monitorSourceStatusLabel(materialMonitorStatus.status) : '尚未檢查'}</p><button className="text-button" type="button" onClick={() => setPage('material-events')}>查看重大訊息</button></article>
          <article className="overview-card"><h3>今日違約揭露</h3><p className="overview-value">{overviewDisclosureCount === null ? '載入中…' : disclosureCountLabel(overviewDisclosureCount, disclosureMonitorStatus, taipeiDateToday(), 'all')}</p><p className="muted">{disclosureMonitorStatus ? disclosureStatusLabel(disclosureMonitorStatus.status) : '尚未檢查'}</p><button className="text-button" type="button" onClick={() => { setPage('disclosures'); void loadDisclosures(); }}>查看全市場揭露</button></article>
          <article className="overview-card"><h3>近期法說會</h3><p className="overview-value">{overviewConferences.length} 場</p><p className="muted">未來 14 天</p><button className="text-button" type="button" onClick={() => setPage('conferences')}>查看法說會</button></article>
          <article className="overview-card"><h3>資料來源</h3><p className="overview-value">{materialMonitorStatus || disclosureMonitorStatus ? `${[...(materialMonitorStatus?.sources.map((source) => source.status) ?? []), ...Object.values(disclosureMonitorStatus?.markets ?? {}).map((item) => item.status)].filter((status) => status === 'complete' || status === 'empty_success').length} / ${ (materialMonitorStatus?.sources.length ?? 0) + Object.keys(disclosureMonitorStatus?.markets ?? {}).length}` : '尚未檢查'}</p><p className="muted">{materialMonitorStatus?.status === 'degraded' || disclosureMonitorStatus?.status === 'degraded' ? '部分來源降級' : '檢查各來源狀態'}</p><button className="text-button" type="button" onClick={() => setPage('settings')}>查看監控狀態</button></article>
        </div>
        <div className="overview-details">
          <section className="overview-section" aria-labelledby="overview-recent-title">
            <div className="section-heading"><h3 id="overview-recent-title">剛收到的重大訊息</h3><button className="text-button" type="button" onClick={() => setPage('material-events')}>查看全部</button></div>
            {overviewMaterialEvents.length === 0 ? <p className="empty-state">目前沒有未讀重大訊息。</p> : <ul className="overview-event-list">
              {overviewMaterialEvents.slice(0, 3).map((event) => <li key={event.id}><button className="overview-event-button" type="button" onClick={() => { setPage('material-events'); void openMaterialEvent(event.id); }}><strong>{event.stockCode} {event.companyName}｜{event.title}</strong><small>{new Date(event.publishedAt).toLocaleString('zh-TW', { timeZone: 'Asia/Taipei' })} · {event.eventType === 'correction' ? '更正' : event.eventType === 'supplement' ? '補充' : '公告'}</small></button></li>)}
            </ul>}
          </section>
          <section className="overview-section" aria-labelledby="overview-schedule-title">
            <div className="section-heading"><h3 id="overview-schedule-title">監控排程</h3><button className="text-button" type="button" onClick={() => setPage('settings')}>設定</button></div>
            {!overviewScheduleStatus ? <p role="status">正在載入排程狀態…</p> : <>
              {overviewScheduleStatus.available === false && <p role="status">隔離測試模式不會連線官方網站；立即檢查已停用。</p>}
              {(['material', 'disclosure'] as const).map((kind) => {
                const job = overviewScheduleStatus[kind];
                const title = kind === 'material' ? '重大訊息' : '違約交割';
                return <div className="overview-health" key={kind}><span>{title}</span><span className={`status ${job.status === 'failed' ? 'failed' : job.status === 'active' ? 'active' : ''}`}>{job.status === 'failed' ? '檢查失敗' : job.status === 'active' ? '檢查中' : job.lastSuccessAt ? `上次成功 ${new Date(job.lastSuccessAt).toLocaleTimeString('zh-TW', { timeZone: 'Asia/Taipei', hour: '2-digit', minute: '2-digit' })}` : '尚無執行紀錄'}</span><button className="text-button" type="button" disabled={overviewScheduleStatus.available === false || runningSchedule === kind} onClick={() => void runSchedule(kind)}>{runningSchedule === kind ? '檢查中…' : '立即檢查'}</button>{job.errorMessage && <small role="alert">{job.errorMessage}</small>}{scheduleMessageKind === kind && scheduleMessage && <small role="status">{scheduleMessage}</small>}</div>;
              })}
            </>}
          </section>
          <section className="overview-section" aria-labelledby="overview-next-conference-title">
            <div className="section-heading"><h3 id="overview-next-conference-title">下一場法說會</h3><button className="text-button" type="button" onClick={() => setPage('conferences')}>查看全部</button></div>
            {overviewConferences[0] ? <><strong>{overviewConferences[0].stockCode} {overviewConferences[0].companyName}</strong><p>{new Date(overviewConferences[0].startsAt).toLocaleString('zh-TW', { timeZone: 'Asia/Taipei' })}<br /><span className="muted">{overviewConferences[0].location || '地點未提供'}</span></p></> : <p className="empty-state">未來 14 天沒有已保存場次。</p>}
          </section>
        </div>
      </section> : page === 'watchlist' ? <>
      <section aria-labelledby="watchlist-title" className="panel">
        <div className="section-heading">
          <div><p className="eyebrow">WATCHLIST</p><h2 id="watchlist-title">關注公司</h2></div>
          <span className="count">{rows.length} 家</span>
        </div>

        <form aria-label="新增關注公司" className="add-form" onSubmit={addCompany}>
          <label>市場
            <select aria-label="市場" value={market} onChange={(event) => setMarket(event.target.value as 'TWSE' | 'TPEX')}>
              <option value="TWSE">上市</option><option value="TPEX">上櫃</option>
            </select>
          </label>
          <label>股票代號
            <input aria-label="股票代號" autoComplete="off" inputMode="numeric" maxLength={6} value={stockCode} onChange={(event) => setStockCode(event.target.value)} />
          </label>
          <button disabled={busy} type="submit">新增公司</button>
        </form>

        <div className="list-controls">
          <label className="search-label">搜尋
            <input aria-label="搜尋關注公司" type="search" placeholder="公司名稱或股票代號" value={query} onChange={(event) => setQuery(event.target.value)} />
          </label>
          <label className="check-label"><input aria-label="只顯示啟用公司" type="checkbox" checked={activeOnly} onChange={(event) => setActiveOnly(event.target.checked)} />只顯示啟用公司</label>
        </div>

        {error && <p role="alert" className="error-message">{error}</p>}
        {rows.length === 0 ? <p className="empty-state">目前沒有符合條件的公司。</p> : (
          <ul className="company-list">
            {rows.map((row) => (
              <li key={row.id} className="company-row">
                <div className="company-avatar" aria-hidden="true">{row.stockCode.slice(0, 2)}</div>
                <div className="company-details"><strong>{row.name}</strong><span>{row.stockCode} · {row.market === 'TWSE' ? '上市' : '上櫃'}</span>
                  <div className="detail-fields"><input aria-label={`分類${row.name}`} value={row.category} placeholder="分類" onChange={(event) => setRows((current) => current.map((entry) => entry.id === row.id ? { ...entry, category: event.target.value } : entry))} /><input aria-label={`備註${row.name}`} value={row.notes} placeholder="備註" onChange={(event) => setRows((current) => current.map((entry) => entry.id === row.id ? { ...entry, notes: event.target.value } : entry))} /></div>
                </div>
                <span className={`status ${row.active ? 'active' : ''}`}>{row.active ? '監控中' : '已停用'}</span>
                <button className="text-button" type="button" aria-label={`儲存${row.name}`} onClick={async () => { await window.reporterApi.updateWatchlist({ companyId: row.companyId, category: row.category, notes: row.notes }); await refresh(); }}>儲存</button>
                <button className="text-button" type="button" aria-label={`${row.active ? '停用' : '啟用'}${row.name}`} onClick={async () => { await window.reporterApi.setWatchlistActive({ companyId: row.companyId, active: !row.active }); await refresh(); }}>{row.active ? '停用' : '啟用'}</button>
                <button className="text-button danger" type="button" aria-label={`移除${row.name}`} onClick={async () => { await window.reporterApi.removeWatchlist({ companyId: row.companyId }); await refresh(); }}>移除</button>
              </li>
            ))}
          </ul>
        )}
      <section aria-labelledby="import-title" className="import-panel">
        <div className="section-heading"><div><p className="eyebrow">MIGRATION</p><h2 id="import-title">匯入舊版股票清單</h2></div></div>
        <p className="muted">貼上 config/default.json 內容。預覽不會修改清單；確認後才會匯入。</p>
        <label className="config-label">設定 JSON
          <textarea aria-label="舊設定 JSON" rows={4} placeholder={'{"StockNumbers":"2330,2317"}'} value={configText} onChange={(event) => { setConfigText(event.target.value); setImportReport(null); }} />
        </label>
        <div className="button-row"><button className="secondary" type="button" onClick={() => void previewImport()}>預覽匯入</button><button disabled={!importReport || busy || importReport.summary.added === 0} type="button" onClick={() => void applyImport()}>確認並匯入</button></div>
        {importReport && <div className="import-result" aria-live="polite"><strong>新增 {importReport.summary.added}、略過 {importReport.summary.skipped}、失敗 {importReport.summary.failed}</strong>{importReport.items.length > 0 && <ul>{importReport.items.map((item, index) => <li key={`${item.stockCode}-${index}`}>{item.stockCode}：{item.reason}</li>)}</ul>}</div>}
      </section>
      </section>
      </> : page === 'disclosures' ? <section aria-labelledby="disclosure-title" className="panel">
        <div className="section-heading"><div><p className="eyebrow">MARKET DISCLOSURES</p><h2 id="disclosure-title">全市場違約交割</h2></div><span className="count">{disclosureCountLabel(disclosures.length, disclosureMonitorStatus, disclosureDate, disclosureMarket)}</span></div>
        {disclosureMonitorStatus && <div role="status" aria-label="違約交割來源狀態" className={`source-health ${disclosureMonitorStatus.status}`}>
          <strong>{disclosureMonitorStatus.status === 'complete' ? '資料來源同步完成' : disclosureMonitorStatus.status === 'degraded' ? '資料來源部分降級' : disclosureMonitorStatus.status === 'stale' ? '官方資料尚未更新' : disclosureMonitorStatus.status === 'failed' ? '資料來源檢查失敗' : '資料來源檢查中'}</strong>
          {disclosureMonitorStatus.errorMessage && <p>{disclosureMonitorStatus.errorMessage}</p>}
          <ul>{Object.entries(disclosureMonitorStatus.markets).map(([source, item]) => <li key={source}>{source}：{disclosureStatusLabel(item.status)}{item.dataDate ? `（資料日期 ${item.dataDate}）` : ''}{item.retryAt ? `；預計重試 ${formatTaipeiDateTime(item.retryAt)}` : ''}{item.errorMessage ? ` — ${item.errorMessage}` : ''}</li>)}</ul>
        </div>}
        <div className="disclosure-filters">
          <label>資料日期<input aria-label="違約交割資料日期" type="date" value={disclosureDate} onChange={(event) => { setDisclosureDate(event.target.value); void loadDisclosures(event.target.value); }} /></label>
          <label>市場<select aria-label="違約交割市場" value={disclosureMarket} onChange={(event) => { const value = event.target.value as 'all' | 'TWSE' | 'TPEX'; setDisclosureMarket(value); void loadDisclosures(disclosureDate, value); }}><option value="all">全部市場</option><option value="TWSE">上市</option><option value="TPEX">上櫃</option></select></label>
        </div>
        {error && <p role="alert" className="error-message">{error}</p>}
        {disclosures.length === 0 ? <p className="empty-state">{disclosureEmptyMessage(disclosureMonitorStatus, disclosureDate, disclosureMarket)}</p> : <div className="disclosure-list" role="list">
          {disclosures.map((entry) => <article key={entry.id} className={`disclosure-row ${entry.isWatched ? 'watched' : ''}`} role="listitem">
            <div className="disclosure-market">{entry.market === 'TWSE' ? '上市' : '上櫃'}</div>
            <div className="company-details"><strong>{entry.companyName || '公司名稱未提供'}</strong><span>{entry.stockCode} · 券商 {entry.brokerCode || '未提供'}</span></div>
            {entry.isWatched && <span className="watched-badge">關注公司</span>}
            <span className="disclosure-date">{entry.disclosureDate}</span>
          </article>)}
        </div>}
      </section> : page === 'conferences' ? <section aria-labelledby="conference-title" className="panel">
        <div className="section-heading"><div><p className="eyebrow">CONFERENCES</p><h2 id="conference-title">法說會與 Google Calendar</h2></div><span className={`status ${googleCalendarStatus?.status === 'connected' ? 'active' : 'failed'}`}>{googleCalendarStatus?.status === 'connected' ? 'Google 已連線' : googleCalendarStatus?.status === 'reauthorization-required' ? '需要重新授權' : googleCalendarStatus?.status === 'disconnected' ? '尚未連線' : '尚未設定'}</span></div>
        {error && <p role="alert" className="error-message">{error}</p>}
        <div className="conference-actions">
          {googleCalendarStatus?.status === 'disconnected' && <button className="secondary" disabled={conferenceBusy} type="button" onClick={() => void migrateLegacyGoogleToken()}>匯入舊版 token（不刪除原檔）</button>}
          {googleCalendarStatus?.status !== 'connected' && <button disabled={conferenceBusy} type="button" onClick={() => void connectGoogleCalendar()}>{googleCalendarStatus?.status === 'reauthorization-required' ? '重新授權 Google Calendar' : '連線 Google Calendar'}</button>}
          {googleCalendarStatus?.status === 'connected' && <><button disabled={conferenceBusy} type="button" onClick={() => void syncConferences()}>立即擷取並同步法說會</button><button className="secondary" disabled={conferenceBusy} type="button" onClick={() => void disconnectGoogleCalendar()}>中斷 Google Calendar</button></>}
        </div>
        {conferenceMessage && <p role="status">{conferenceMessage}</p>}
        {conferenceRows.length === 0 ? <p className="empty-state">目前沒有已保存的法說會。連線後可擷取關注公司的近期場次。</p> : <div className="conference-list" role="list">
          {conferenceRows.map((conference) => <article className="conference-row" key={conference.id} role="listitem">
            <div className="company-details"><strong>{conference.companyName} · {conference.stockCode}</strong><span>{new Date(conference.startsAt).toLocaleString('zh-TW', { timeZone: 'Asia/Taipei' })} · {conference.location || '地點未提供'}</span>{conference.sourceUrl && <a href={conference.sourceUrl} target="_blank" rel="noreferrer">開啟法說會來源</a>}</div>
            <span className={`status ${conference.syncStatus === 'failed' ? 'failed' : conference.syncStatus === 'synced' ? 'active' : ''}`}>{conference.syncStatus === 'pending' ? '待同步' : conference.syncStatus === 'existing' ? '已存在' : conference.syncStatus === 'synced' ? '已同步' : '同步失敗'}</span>
            {conference.lastError && <p role="alert" className="error-message">{conference.lastError}</p>}
          </article>)}
        </div>}
      </section> : page === 'settings' ? <section aria-labelledby="settings-title" className="panel">
        <div className="section-heading"><div><p className="eyebrow">SCHEDULE</p><h2 id="settings-title">監控排程</h2></div></div>
        {error && <p role="alert" className="error-message">{error}</p>}
        {!scheduleSettings || !scheduleStatus ? <p role="status">正在載入排程設定…</p> : <>
          {scheduleStatus.available === false && <p role="status">隔離測試模式不會連線官方網站；立即檢查已停用。</p>}
          <div className="schedule-grid">
            {(['material', 'disclosure'] as const).map((kind) => {
              const status = scheduleStatus[kind];
              const title = kind === 'material' ? '重大訊息' : '違約交割';
              return <article className="schedule-card" key={kind}>
                <div className="section-heading"><h3>{title}</h3><span className={`status ${status.status === 'failed' ? 'failed' : status.status === 'active' ? 'active' : ''}`}>{status.status === 'idle' ? '尚無執行紀錄' : status.status}</span></div>
                <p>上次成功：{status.lastSuccessAt ? new Date(status.lastSuccessAt).toLocaleString('zh-TW', { timeZone: 'Asia/Taipei' }) : '尚無'}</p>
                <p>上次結果：{status.lastResult ?? '尚無'}</p>
                <p>預計下次：{scheduleStatus.available === false ? '隔離測試模式不執行' : status.nextRunAt ? new Date(status.nextRunAt).toLocaleString('zh-TW', { timeZone: 'Asia/Taipei' }) : '已停用'}</p>
                {status.errorMessage && <p role="alert" className="error-message">{status.errorMessage}</p>}
                <button disabled={scheduleStatus.available === false || runningSchedule === kind} type="button" onClick={() => void runSchedule(kind)}>{runningSchedule === kind ? `檢查${title}中…` : `立即檢查${title}`}</button>
                {scheduleMessageKind === kind && scheduleMessage && <p role="status">{scheduleMessage}</p>}
              </article>;
            })}
          </div>
          <div className="schedule-settings">
            <h3>重大訊息檢查</h3>
            <label className="check-label"><input type="checkbox" aria-label="啟用重大訊息監控" checked={scheduleSettings.monitoring.enabled} onChange={(event) => setScheduleSettings({ ...scheduleSettings, monitoring: { ...scheduleSettings.monitoring, enabled: event.target.checked } })} />啟用</label>
            <label>開始時間<input aria-label="重大訊息開始時間" type="time" value={scheduleSettings.monitoring.start} onChange={(event) => setScheduleSettings({ ...scheduleSettings, monitoring: { ...scheduleSettings.monitoring, start: event.target.value } })} /></label>
            <label>結束時間<input aria-label="重大訊息結束時間" type="time" value={scheduleSettings.monitoring.end} onChange={(event) => setScheduleSettings({ ...scheduleSettings, monitoring: { ...scheduleSettings.monitoring, end: event.target.value } })} /></label>
            <label>檢查頻率<select aria-label="重大訊息檢查頻率" value={scheduleSettings.monitoring.intervalMinutes} onChange={(event) => setScheduleSettings({ ...scheduleSettings, monitoring: { ...scheduleSettings.monitoring, intervalMinutes: Number(event.target.value) as ScheduleSettings['monitoring']['intervalMinutes'] } })}><option value={15}>15 分鐘</option><option value={30}>30 分鐘</option><option value={60}>60 分鐘</option><option value={120}>120 分鐘</option></select></label>
            <h3>違約交割檢查</h3>
            <label className="check-label"><input type="checkbox" aria-label="啟用違約交割監控" checked={scheduleSettings.disclosure.enabled} onChange={(event) => setScheduleSettings({ ...scheduleSettings, disclosure: { ...scheduleSettings.disclosure, enabled: event.target.checked } })} />啟用</label>
            <label>每日檢查時間<input aria-label="違約交割檢查時間" type="time" value={scheduleSettings.disclosure.runAt} onChange={(event) => setScheduleSettings({ ...scheduleSettings, disclosure: { ...scheduleSettings.disclosure, runAt: event.target.value } })} /></label>
            <label className="check-label"><input type="checkbox" aria-label="收到本日無違約揭露通知" checked={scheduleSettings.notifyEmptyDefaultDisclosures} onChange={(event) => setScheduleSettings({ ...scheduleSettings, notifyEmptyDefaultDisclosures: event.target.checked })} />本日無違約揭露時通知</label>
            <button type="button" onClick={() => void saveSchedule()}>儲存排程</button>
          </div>
          <div className="schedule-settings desktop-startup-settings">
            <h3>Windows 登入啟動</h3>
            {!loginStartupSettings ? <p role="status">正在載入登入啟動設定…</p> : <>
              <label className="check-label"><input type="checkbox" aria-label="登入 Windows 時啟動" checked={loginStartupSettings.enabled} onChange={(event) => setLoginStartupSettings({ ...loginStartupSettings, enabled: event.target.checked })} />登入 Windows 時啟動</label>
              <label className="check-label"><input type="checkbox" aria-label="啟動後縮到系統匣" checked={loginStartupSettings.startHidden} disabled={!loginStartupSettings.enabled} onChange={(event) => setLoginStartupSettings({ ...loginStartupSettings, startHidden: event.target.checked })} />啟動後縮到系統匣</label>
              <button type="button" onClick={() => void saveLoginStartup()}>儲存登入啟動設定</button>
            </>}
          </div>
          <div className="schedule-settings data-export-settings">
            <h3>本機資料備份</h3>
            <p>匯出版本化 JSON；可重用的 OAuth 憑證與 token 不會包含在檔案中。</p>
            <button type="button" disabled={dataExportBusy} onClick={() => void exportLocalData()}>{dataExportBusy ? '正在匯出…' : '匯出資料'}</button>
            {dataExportMessage && <p role={dataExportMessage.includes('失敗') || dataExportMessage.includes('無法') ? 'alert' : 'status'}>{dataExportMessage}</p>}
          </div>
        </>}
      </section> : <section aria-labelledby="material-title" className="panel material-panel">
        <div className="section-heading"><div><p className="eyebrow">MATERIAL EVENTS</p><h2 id="material-title">重大訊息</h2></div><span className="count">{showAllMaterial ? '全部' : '關注'} {materialEvents.length} 筆</span></div>
        {materialMonitorStatus && <div role="status" aria-label="重大訊息來源狀態" className={`source-health ${materialMonitorStatus.status}`}>
          <strong>{materialMonitorStatus.status === 'complete' ? '資料來源同步完成' : materialMonitorStatus.status === 'degraded' ? '資料來源部分降級' : materialMonitorStatus.status === 'stale' ? '官方資料尚未更新' : materialMonitorStatus.status === 'failed' ? '資料來源檢查失敗' : '資料來源檢查中'}</strong>
          {materialMonitorStatus.errorMessage && <p>{materialMonitorStatus.errorMessage}</p>}
          <ul>{materialMonitorStatus.sources.map((source) => <li key={source.source}>{materialSourceName(source.source)}：{monitorSourceStatusLabel(source.status)}{source.dataDate ? `（資料日期 ${source.dataDate}）` : ''}{source.errorMessage ? ` — ${source.errorMessage}` : ''}</li>)}</ul>
        </div>}
        <div className="list-controls material-controls">
          <label className="search-label">搜尋
            <input aria-label="搜尋重大訊息" type="search" placeholder="公司、代號或主旨" value={materialQuery} onChange={(event) => setMaterialQuery(event.target.value)} />
          </label>
          <label className="check-label"><input aria-label="只顯示未讀重大訊息" type="checkbox" checked={unreadOnly} onChange={(event) => setUnreadOnly(event.target.checked)} />只顯示未讀</label>
          <label className="check-label"><input aria-label="顯示全部公告" type="checkbox" checked={showAllMaterial} onChange={(event) => setShowAllMaterial(event.target.checked)} />顯示全部公告</label>
        </div>
        {error && <p role="alert" className="error-message">{error}</p>}
        {materialEvents.length === 0 ? <p className="empty-state">{showAllMaterial ? '目前沒有符合條件的重大訊息。' : '目前沒有符合條件的關注公司重大訊息；可勾選「顯示全部公告」。'}</p> : <div className="material-list" role="list">
          {materialEvents.map((event) => <article className={`material-row ${event.readAt ? 'read' : 'unread'} ${expandedMaterialRowId === event.id ? 'expanded' : ''}`} key={event.id} role="listitem">
            <button className="material-summary" type="button" aria-label={`${event.companyName} ${event.title}`} aria-expanded={expandedMaterialRowId === event.id} aria-controls={expandedMaterialRowId === event.id ? `material-detail-${event.id}` : undefined} onClick={() => void openMaterialEvent(event.id)}>
              <span className="material-code">{event.stockCode}</span><span className="material-copy"><strong>{event.companyName} · {event.title}</strong><small>{new Date(event.publishedAt).toLocaleString('zh-TW', { timeZone: 'Asia/Taipei' })} · {event.eventType === 'correction' ? '更正' : event.eventType === 'supplement' ? '補充' : '公告'} · {event.readAt ? '已讀' : '未讀'}</small></span>
            </button>
            {expandedMaterialRowId === event.id && !selectedMaterialEvent && <p id={`material-detail-${event.id}`} className="material-detail-loading" role="status">載入完整公告…</p>}
            {expandedMaterialRowId === event.id && selectedMaterialEvent && <aside id={`material-detail-${event.id}`} aria-label="重大訊息完整內容" className="event-detail" role="complementary">
              <div className="detail-heading"><div><p className="eyebrow">{selectedMaterialEvent.eventType === 'correction' ? 'CORRECTION' : selectedMaterialEvent.eventType === 'supplement' ? 'SUPPLEMENT' : 'ANNOUNCEMENT'}</p><h3>{selectedMaterialEvent.companyName} · {selectedMaterialEvent.title}</h3></div><button className="text-button" type="button" aria-label="關閉重大訊息詳情" onClick={() => { materialDetailRequest.current += 1; setExpandedMaterialRowId(null); setSelectedMaterialEvent(null); }}>關閉</button></div>
              <p className="detail-meta">{selectedMaterialEvent.stockCode} · {selectedMaterialEvent.market === 'TWSE' ? '上市' : '上櫃'} · 發布 {new Date(selectedMaterialEvent.publishedAt).toLocaleString('zh-TW', { timeZone: 'Asia/Taipei' })}</p>
              <div className="event-content">{selectedMaterialEvent.content}</div>
              <div className="event-provenance"><span>首次發現：{new Date(selectedMaterialEvent.discoveredAt).toLocaleString('zh-TW', { timeZone: 'Asia/Taipei' })}</span><span>來源：{selectedMaterialEvent.source} · 本機已保存</span><a href={selectedMaterialEvent.sourceUrl} target="_blank" rel="noreferrer">開啟來源公告</a></div>
              {selectedMaterialEvent.original && <button className="original-link" type="button" onClick={() => void openMaterialEvent(selectedMaterialEvent.original!.id, event.id)}>原公告：{selectedMaterialEvent.original.title}</button>}
              {selectedMaterialEvent.relatedRevisions.length > 0 && <p>後續版本：{selectedMaterialEvent.relatedRevisions.map((revision) => revision.title).join('、')}</p>}
              {!selectedMaterialEvent.readAt && <button className="secondary" type="button" onClick={() => void markSelectedMaterialEventRead()}>標記為已讀</button>}
            </aside>}
          </article>)}
        </div>}
      </section>}
      <footer>資料儲存在此電腦。外部名錄僅於新增或匯入時查詢。</footer>
    </main>
  );
}
