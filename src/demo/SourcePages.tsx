import { useState } from "react";
import {
  DateControl,
  Icon,
  Modal,
  PageHeader,
  SourceDialog,
  TextAction,
} from "./components";
import { conferences } from "./data";
type Notice = (text: string) => void;
export function DisclosuresPage({ notify }: { notify: Notice }) {
  const [date, setDate] = useState("2026-11-04"),
    [market, setMarket] = useState("全部市場"),
    [state, setState] = useState("confirmed"),
    [source, setSource] = useState(false),
    [log, setLog] = useState(false);
  const confirmed = state === "confirmed" && date === "2026-11-04";
  const rows =
    market === "全部市場"
      ? ["上市 TWSE", "上櫃 TPEX"]
      : [market === "上市" ? "上市 TWSE" : "上櫃 TPEX"];
  return (
    <>
      <PageHeader title="違約交割" subtitle="全市場個股違約資訊揭露" />
      <div className="disclosure-toolbar">
        <DateControl value={date} onChange={setDate} />
        <select
          aria-label="市場"
          value={market}
          onChange={(e) => setMarket(e.target.value)}
        >
          <option>全部市場</option>
          <option>上市</option>
          <option>上櫃</option>
        </select>
        <button
          onClick={() =>
            notify(
              confirmed
                ? "示意檢查完成：當日確認無揭露"
                : "示意來源尚未更新，不能判定為零筆",
            )
          }
        >
          立即檢查
        </button>
      </div>
      <section
        className={
          confirmed ? "disclosure-result" : "disclosure-result unconfirmed"
        }
      >
        <Icon name={confirmed ? "checkCircle" : "warning"} size={50} />
        <div>
          <h2>
            {confirmed ? "當日無個股達違約資訊揭露標準" : "今日資料尚未確認"}
          </h2>
          <p>
            {confirmed
              ? "上市與上櫃來源均已完成當日檢查。"
              : "來源日期落後或檢查失敗，不能視為沒有揭露。"}
          </p>
        </div>
      </section>
      <h2>來源狀態</h2>
      <div className="table-scroll">
        <table className="source-table">
          <thead>
            <tr>
              <th>市場</th>
              <th>資料日期</th>
              <th>檢查結果</th>
              <th>來源</th>
            </tr>
          </thead>
          <tbody>
            {rows.map((row) => (
              <tr key={row}>
                <td>{row}</td>
                <td>{confirmed ? "2026/11/04" : "2026/11/03"}</td>
                <td className={confirmed ? "positive" : "muted"}>
                  {confirmed
                    ? "已確認 · 無揭露"
                    : state === "failed"
                      ? "檢查失敗"
                      : "尚未確認 · 資料過期"}
                </td>
                <td>
                  <TextAction onClick={() => setSource(true)}>
                    查看官方資料
                  </TextAction>
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
      <div className="table-foot">
        <span>最近檢查 2026/11/04 18:31</span>
        <TextAction onClick={() => setLog(true)}>查看檢查紀錄</TextAction>
      </div>
      <footer className="page-foot">
        <span>來源失敗或日期落後時，會顯示「尚未確認」，不視為無揭露。</span>
      </footer>
      <details className="demo-scenario">
        <summary>切換示意狀態</summary>
        <label>
          來源情境
          <select value={state} onChange={(e) => setState(e.target.value)}>
            <option value="confirmed">當日確認無揭露</option>
            <option value="stale">來源日期落後</option>
            <option value="failed">來源檢查失敗</option>
          </select>
        </label>
      </details>
      {source && <SourceDialog onClose={() => setSource(false)} />}
      {log && (
        <Modal title="最近檢查紀錄" onClose={() => setLog(false)}>
          <dl className="detail-list">
            <div>
              <dt>示意執行時間</dt>
              <dd>2026/11/04 18:31</dd>
            </div>
            <div>
              <dt>TWSE／TPEX</dt>
              <dd>{confirmed ? "當日確認無揭露" : "尚未確認"}</dd>
            </div>
          </dl>
          <div className="dialog-actions">
            <button onClick={() => setLog(false)}>關閉</button>
          </div>
        </Modal>
      )}
    </>
  );
}
export function ConferencesPage({
  connected,
  onConnect,
  notify,
  settings,
}: {
  connected: boolean;
  onConnect: (v: boolean) => void;
  notify: Notice;
  settings: () => void;
}) {
  const [menu, setMenu] = useState(false),
    [source, setSource] = useState(false),
    [synced, setSynced] = useState(false),
    [busy, setBusy] = useState(false);
  async function sync() {
    setBusy(true);
    notify("正在示意擷取與查重…");
    await new Promise((resolve) => setTimeout(resolve, 500));
    setSynced(true);
    setBusy(false);
    notify("示意同步完成：3 筆已存在，未重複新增；沒有寫入 Google Calendar");
  }
  return (
    <>
      <PageHeader title="法說會" subtitle="關注公司的近期場次與行事曆同步">
        <button
          disabled={busy}
          onClick={() => {
            if (!connected) {
              onConnect(true);
              notify("已模擬連線，未開啟 Google 授權");
            } else void sync();
          }}
        >
          {connected
            ? busy
              ? "同步示意中…"
              : "立即擷取並同步法說會"
            : "連線 Google Calendar"}
        </button>
      </PageHeader>
      <div className="calendar-connection">
        <span>
          <Icon name="calendar" size={28} />
          Google Calendar{" "}
          <strong className={connected ? "positive" : "muted"}>
            {connected ? "已連線" : "尚未連線"}
          </strong>
          <span className="connection-scope">主要行事曆</span>
        </span>
        <div className="connection-menu">
          <button className="text-action" onClick={() => setMenu(!menu)}>
            管理連線
            <Icon name="down" size={18} />
          </button>
          {menu && (
            <div className="dropdown-menu">
              <button
                onClick={() => {
                  onConnect(!connected);
                  setMenu(false);
                  notify(
                    connected
                      ? "已中斷示意連線，未刪除任何正式 token"
                      : "已模擬連線",
                  );
                }}
              >
                {connected ? "中斷 Google Calendar" : "連線 Google Calendar"}
              </button>
            </div>
          )}
        </div>
      </div>
      <div className="schedule-line">
        <span>每日 19:00 自動同步 · v2 排程運作中</span>
        <TextAction onClick={settings} icon={false}>
          查看排程
        </TextAction>
      </div>
      <p className="sync-result">
        上次同步 11/05 14:30 · 新增 {synced ? "0" : "1"} · 已存在{" "}
        {synced ? "3" : "2"} · 失敗 0
      </p>
      <h2 className="upcoming-title">即將舉行</h2>
      <div className="table-scroll">
        <table className="conference-table">
          <thead>
            <tr>
              <th>日期與時間</th>
              <th>公司</th>
              <th>地點／方式</th>
              <th>行事曆</th>
            </tr>
          </thead>
          <tbody>
            {conferences.map((row, i) => (
              <tr key={row.code}>
                <td>
                  <strong>
                    {row.date} <small>{row.day}</small>
                  </strong>
                  <br />
                  {row.time}
                </td>
                <td>
                  <strong>{row.name}</strong>
                  <br />
                  <span className="muted number">{row.code}</span>
                </td>
                <td>
                  {row.place}
                  <br />
                  <TextAction onClick={() => setSource(true)} icon={false}>
                    開啟官方來源
                  </TextAction>
                </td>
                <td>
                  <span
                    className={
                      i === 0 && !synced
                        ? "calendar-state positive"
                        : "calendar-state"
                    }
                  >
                    <Icon name="checkCircle" size={22} />
                    {i === 0 && !synced ? "已同步" : "已存在"}
                  </span>
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
      <footer className="page-foot">
        會先確認 Google 行事曆是否已有相同場次，避免重複新增。
      </footer>
      {source && <SourceDialog onClose={() => setSource(false)} />}
    </>
  );
}
