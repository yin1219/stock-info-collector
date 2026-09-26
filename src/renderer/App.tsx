import { useCallback, useEffect, useState, type FormEvent } from 'react';
import type { ReporterRoute } from '../preload/api';

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

type Page = 'watchlist' | 'disclosures' | 'material-events';

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
  const [disclosureDate, setDisclosureDate] = useState(taipeiDateToday);
  const [disclosureMarket, setDisclosureMarket] = useState<'all' | 'TWSE' | 'TPEX'>('all');
  const [materialEvents, setMaterialEvents] = useState<MaterialEventRow[]>([]);
  const [materialQuery, setMaterialQuery] = useState('');
  const [unreadOnly, setUnreadOnly] = useState(false);
  const [selectedMaterialEvent, setSelectedMaterialEvent] = useState<MaterialEventDetail | null>(null);
  const [notificationEventIds, setNotificationEventIds] = useState<string[] | null>(null);

  const refresh = useCallback(async () => {
    const result = await window.reporterApi.listWatchlist({ query, activeOnly });
    setRows(result as WatchlistRow[]);
  }, [query, activeOnly]);

  useEffect(() => {
    if (page !== 'material-events') return;
    void window.reporterApi.listMaterialEvents({
      query: materialQuery, unreadOnly, ...(notificationEventIds ? { eventIds: notificationEventIds } : {}),
    }).then((result) => {
      setMaterialEvents(result as MaterialEventRow[]);
    }).catch((cause: unknown) => setError(cause instanceof Error ? cause.message : String(cause)));
  }, [page, materialQuery, unreadOnly, notificationEventIds]);

  useEffect(() => window.reporterApi.onNotificationRoute((route: ReporterRoute) => {
    setError('');
    if (route.type === 'disclosure-list') {
      setNotificationEventIds(null);
      setPage('disclosures');
      void window.reporterApi.listDisclosures({ disclosureDate }).then((result) => setDisclosures(result as DisclosureRow[]));
      return;
    }
    setPage('material-events');
    setMaterialQuery('');
    setUnreadOnly(false);
    if (route.type === 'event-list') {
      setNotificationEventIds(route.eventIds);
      setSelectedMaterialEvent(null);
      return;
    }
    setNotificationEventIds([route.eventId]);
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
      const result = await window.reporterApi.listDisclosures({
        disclosureDate: date,
        ...(selectedMarket === 'all' ? {} : { market: selectedMarket }),
      });
      setDisclosures(result as DisclosureRow[]);
    } catch (cause) {
      setError(cause instanceof Error ? cause.message : String(cause));
    }
  }

  async function openMaterialEvent(id: string) {
    setError('');
    try {
      setSelectedMaterialEvent(await window.reporterApi.getMaterialEvent(id) as MaterialEventDetail);
    } catch (cause) {
      setError(cause instanceof Error ? cause.message : String(cause));
    }
  }

  async function markSelectedMaterialEventRead() {
    if (!selectedMaterialEvent) return;
    await window.reporterApi.markMaterialEventRead(selectedMaterialEvent.id);
    setSelectedMaterialEvent((current) => current ? { ...current, readAt: new Date().toISOString() } : current);
    await window.reporterApi.listMaterialEvents({ query: materialQuery, unreadOnly }).then((result) => setMaterialEvents(result as MaterialEventRow[]));
  }

  return (
    <main className="app-shell">
      <header className="app-header">
        <div>
          <p className="eyebrow">STOCK REPORTER ASSISTANT</p>
          <h1>股市記者小幫手</h1>
        </div>
        <span className="version">v2 · 本機資料</span>
      </header>

      <nav aria-label="主要功能" aria-orientation="horizontal" className="main-nav" role="tablist">
        <button aria-selected={page === 'watchlist'} role="tab" type="button" onClick={() => setPage('watchlist')}>關注公司</button>
        <button aria-selected={page === 'disclosures'} role="tab" type="button" onClick={() => { setPage('disclosures'); void loadDisclosures(); }}>違約交割</button>
        <button aria-selected={page === 'material-events'} role="tab" type="button" onClick={() => { setNotificationEventIds(null); setPage('material-events'); }}>重大訊息</button>
      </nav>

      {page === 'watchlist' ? <>
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
      </section>

      <section aria-labelledby="import-title" className="panel import-panel">
        <div className="section-heading"><div><p className="eyebrow">MIGRATION</p><h2 id="import-title">匯入舊版股票清單</h2></div></div>
        <p className="muted">貼上 config/default.json 內容。預覽不會修改清單；確認後才會匯入。</p>
        <label className="config-label">設定 JSON
          <textarea aria-label="舊設定 JSON" rows={4} placeholder={'{"StockNumbers":"2330,2317"}'} value={configText} onChange={(event) => { setConfigText(event.target.value); setImportReport(null); }} />
        </label>
        <div className="button-row"><button className="secondary" type="button" onClick={() => void previewImport()}>預覽匯入</button><button disabled={!importReport || busy || importReport.summary.added === 0} type="button" onClick={() => void applyImport()}>確認並匯入</button></div>
        {importReport && <div className="import-result" aria-live="polite"><strong>新增 {importReport.summary.added}、略過 {importReport.summary.skipped}、失敗 {importReport.summary.failed}</strong>{importReport.items.length > 0 && <ul>{importReport.items.map((item, index) => <li key={`${item.stockCode}-${index}`}>{item.stockCode}：{item.reason}</li>)}</ul>}</div>}
      </section>
      </> : page === 'disclosures' ? <section aria-labelledby="disclosure-title" className="panel">
        <div className="section-heading"><div><p className="eyebrow">MARKET DISCLOSURES</p><h2 id="disclosure-title">全市場違約交割</h2></div><span className="count">{disclosures.length} 筆</span></div>
        <div className="disclosure-filters">
          <label>資料日期<input aria-label="違約交割資料日期" type="date" value={disclosureDate} onChange={(event) => { setDisclosureDate(event.target.value); void loadDisclosures(event.target.value); }} /></label>
          <label>市場<select aria-label="違約交割市場" value={disclosureMarket} onChange={(event) => { const value = event.target.value as 'all' | 'TWSE' | 'TPEX'; setDisclosureMarket(value); void loadDisclosures(disclosureDate, value); }}><option value="all">全部市場</option><option value="TWSE">上市</option><option value="TPEX">上櫃</option></select></label>
        </div>
        {error && <p role="alert" className="error-message">{error}</p>}
        {disclosures.length === 0 ? <p className="empty-state">此日期目前沒有可顯示的違約交割資料。</p> : <div className="disclosure-list" role="list">
          {disclosures.map((entry) => <article key={entry.id} className={`disclosure-row ${entry.isWatched ? 'watched' : ''}`} role="listitem">
            <div className="disclosure-market">{entry.market === 'TWSE' ? '上市' : '上櫃'}</div>
            <div className="company-details"><strong>{entry.companyName || '公司名稱未提供'}</strong><span>{entry.stockCode} · 券商 {entry.brokerCode || '未提供'}</span></div>
            {entry.isWatched && <span className="watched-badge">關注公司</span>}
            <span className="disclosure-date">{entry.disclosureDate}</span>
          </article>)}
        </div>}
      </section> : <section aria-labelledby="material-title" className="panel material-panel">
        <div className="section-heading"><div><p className="eyebrow">MATERIAL EVENTS</p><h2 id="material-title">重大訊息</h2></div><span className="count">{materialEvents.length} 筆</span></div>
        <div className="list-controls material-controls">
          <label className="search-label">搜尋
            <input aria-label="搜尋重大訊息" type="search" placeholder="公司、代號或主旨" value={materialQuery} onChange={(event) => setMaterialQuery(event.target.value)} />
          </label>
          <label className="check-label"><input aria-label="只顯示未讀重大訊息" type="checkbox" checked={unreadOnly} onChange={(event) => setUnreadOnly(event.target.checked)} />只顯示未讀</label>
        </div>
        {error && <p role="alert" className="error-message">{error}</p>}
        {materialEvents.length === 0 ? <p className="empty-state">目前沒有符合條件的重大訊息。</p> : <div className="material-list" role="list">
          {materialEvents.map((event) => <article className={`material-row ${event.readAt ? 'read' : 'unread'}`} key={event.id} role="listitem">
            <button className="material-summary" type="button" aria-label={`${event.companyName} ${event.title}`} onClick={() => void openMaterialEvent(event.id)}>
              <span className="material-code">{event.stockCode}</span><span className="material-copy"><strong>{event.companyName} · {event.title}</strong><small>{new Date(event.publishedAt).toLocaleString('zh-TW', { timeZone: 'Asia/Taipei' })} · {event.eventType === 'correction' ? '更正' : event.eventType === 'supplement' ? '補充' : '公告'} · {event.readAt ? '已讀' : '未讀'}</small></span>
            </button>
          </article>)}
        </div>}
        {selectedMaterialEvent && <aside aria-label="重大訊息完整內容" className="event-detail" role="complementary">
          <div className="detail-heading"><div><p className="eyebrow">{selectedMaterialEvent.eventType === 'correction' ? 'CORRECTION' : selectedMaterialEvent.eventType === 'supplement' ? 'SUPPLEMENT' : 'ANNOUNCEMENT'}</p><h3>{selectedMaterialEvent.companyName} · {selectedMaterialEvent.title}</h3></div><button className="text-button" type="button" aria-label="關閉重大訊息詳情" onClick={() => setSelectedMaterialEvent(null)}>關閉</button></div>
          <p className="detail-meta">{selectedMaterialEvent.stockCode} · {selectedMaterialEvent.market === 'TWSE' ? '上市' : '上櫃'} · 發布 {new Date(selectedMaterialEvent.publishedAt).toLocaleString('zh-TW', { timeZone: 'Asia/Taipei' })}</p>
          <div className="event-content">{selectedMaterialEvent.content}</div>
          <div className="event-provenance"><span>首次發現：{new Date(selectedMaterialEvent.discoveredAt).toLocaleString('zh-TW', { timeZone: 'Asia/Taipei' })}</span><span>來源：{selectedMaterialEvent.source} · 本機已保存</span><a href={selectedMaterialEvent.sourceUrl} target="_blank" rel="noreferrer">開啟來源公告</a></div>
          {selectedMaterialEvent.original && <button className="original-link" type="button" onClick={() => void openMaterialEvent(selectedMaterialEvent.original!.id)}>原公告：{selectedMaterialEvent.original.title}</button>}
          {selectedMaterialEvent.relatedRevisions.length > 0 && <p>後續版本：{selectedMaterialEvent.relatedRevisions.map((revision) => revision.title).join('、')}</p>}
          {!selectedMaterialEvent.readAt && <button className="secondary" type="button" onClick={() => void markSelectedMaterialEventRead()}>標記為已讀</button>}
        </aside>}
      </section>}
      <footer>資料儲存在此電腦。外部名錄僅於新增或匯入時查詢。</footer>
    </main>
  );
}
