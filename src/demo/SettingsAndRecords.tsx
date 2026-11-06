import { useState } from "react";
import {
  DateControl,
  FileDemo,
  Icon,
  Modal,
  PageHeader,
  Switch,
  Tabs,
  TextAction,
} from "./components";
import { indexRanks, investorRanks, type Preferences } from "./data";
import { CMoneyImportGuide } from "../renderer/CMoneyImportGuide";
type Notice = (message: string) => void;
export function RecordsPage({ notify }: { notify: Notice }) {
  const [tab, setTab] = useState("加權指數"),
    [investor, setInvestor] = useState("投信"),
    [direction, setDirection] = useState("最大漲點"),
    [date, setDate] = useState("2026-07-31"),
    [manage, setManage] = useState(false),
    [formula, setFormula] = useState(false);
  const selected = investorRanks[investor];
  const rise = direction === "最大漲點";
  function changeTab(value: string) {
    setTab(value);
    setDate(value === "加權指數" ? "2026-07-31" : "2026-11-04");
  }
  return (
    <>
      <PageHeader title="盤面紀錄" subtitle="以 2020 年以來的日資料比對">
        <DateControl value={date} onChange={setDate} />
        <button onClick={() => notify("已更新示意資料；沒有呼叫資料來源 API")}>
          更新資料
        </button>
        <button className="text-action" onClick={() => setManage(true)}>
          資料管理
        </button>
      </PageHeader>
      <Tabs
        labels={["加權指數", "三大法人"]}
        value={tab}
        onChange={changeTab}
      />
      {tab === "加權指數" ? (
        <>
          <p className="data-confirmed">
            {date === "2026-07-31" ? "盤後資料已確認" : "此日期尚無示意資料"}
          </p>
          <div className="index-metrics">
            <div>
              <h2>盤中最大漲點</h2>
              <p className="metric-value gain">
                {date === "2026-07-31" ? "+3,281.06" : "—"} <small>點</small>
              </p>
              <p className="metric-rank">
                {date === "2026-07-31" ? "2020 年以來第 1 名" : "待確認"}
              </p>
            </div>
            <div>
              <h2>盤中最大跌點</h2>
              <p className="metric-value">—</p>
              <p>
                {date === "2026-07-31" ? "當日未跌破前收" : "尚無可計算資料"}
              </p>
            </div>
          </div>
          <div className="section-title rankings-title">
            <h2>歷史排名</h2>
            <Tabs
              small
              labels={["最大漲點", "最大跌點"]}
              value={direction}
              onChange={setDirection}
            />
          </div>
          <div className="table-scroll">
            <table className="index-table">
              <thead>
                <tr>
                  <th>名次</th>
                  <th>日期</th>
                  <th className="right">盤中{direction}</th>
                </tr>
              </thead>
              <tbody>
                {indexRanks[rise ? "rise" : "fall"].map((row, i) => (
                  <tr key={row[0]} className={i === 0 ? "highlight" : ""}>
                    <td className="strong">{String(i + 1).padStart(2, "0")}</td>
                    <td>{row[0]}</td>
                    <td
                      className={
                        rise ? "right gain strong" : "right loss strong"
                      }
                    >
                      {rise ? "+" : "−"}
                      {row[1]} 點
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </>
      ) : (
        <>
          <Tabs
            small
            labels={["外資", "投信", "自營商"]}
            value={investor}
            onChange={setInvestor}
          />
          <div className="investor-metrics">
            <div>
              <h2>
                {investor}連續{investor === "投信" ? "買" : "賣"}超
              </h2>
              <p
                className={
                  investor === "投信"
                    ? "metric-value gain"
                    : "metric-value loss"
                }
              >
                {selected.days}
                <small>個交易日</small>
              </p>
            </div>
            <div>
              <h2>累計{investor === "投信" ? "買" : "賣"}超</h2>
              <p
                className={
                  investor === "投信"
                    ? "metric-value gain"
                    : "metric-value loss"
                }
              >
                {selected.amount}
                <small>億元</small>
              </p>
            </div>
            <p className="metric-rank">2020 年以來第 {selected.rank} 名</p>
          </div>
          <p className="statistics-date">統計至 2026/11/04</p>
          <div className="investor-tables">
            {(["buy", "sell"] as const).map((d) => (
              <section key={d}>
                <h2>
                  連續{d === "buy" ? "買" : "賣"}超前三名{" "}
                  <small>（累計金額：億元）</small>
                </h2>
                <div className="table-scroll">
                  <table>
                    <thead>
                      <tr>
                        <th>名次</th>
                        <th>期間</th>
                        <th>天數</th>
                        <th className="right">
                          累計金額
                          <br />
                          <small>（億元）</small>
                        </th>
                      </tr>
                    </thead>
                    <tbody>
                      {selected[d].map((row, i) => (
                        <tr
                          key={row[0]}
                          className={
                            investor === "投信" && d === "buy" && i === 1
                              ? "highlight"
                              : ""
                          }
                        >
                          <td className="strong">
                            {String(i + 1).padStart(2, "0")}
                          </td>
                          <td className="period">
                            {row[0]}—<br />
                            {row[1]}
                          </td>
                          <td>{row[2]} 日</td>
                          <td
                            className={`right strong ${d === "buy" ? "gain" : "loss"}`}
                          >
                            {row[3]}
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </section>
            ))}
          </div>
        </>
      )}
      <footer className="page-foot">
        <span>
          資料範圍 2020/01/02—2026/11/04 <b className="divider">|</b> 來源：
          {tab === "加權指數" ? "CMoney／TWSE" : "CMoney"}
        </span>
        <TextAction onClick={() => setFormula(true)}>查看計算方式</TextAction>
      </footer>
      {manage && (
        <Modal title="盤面資料管理" onClose={() => setManage(false)} wide>
          <p>這是示意匯入，不會讀取檔案內容或修改正式資料。</p>
          <CMoneyImportGuide
            renderImport={(_kind, label) => (
              <FileDemo
                label={`選擇${label} Excel`}
                accept=".xlsx"
                onSelect={(name) =>
                  notify(`已選擇${label}：${name}，僅示意匯入`)
                }
              />
            )}
          />
          <h3>每日法人更新（TWSE）</h3>
          <FileDemo
            label="選擇 TWSE 當日 CSV"
            onSelect={(name) => notify(`已選擇 ${name}，僅示意匯入`)}
          />
          <div className="dialog-actions">
            <button onClick={() => setManage(false)}>完成</button>
          </div>
        </Modal>
      )}
      {formula && (
        <Modal title="計算方式" onClose={() => setFormula(false)}>
          <dl className="formula-list">
            <dt>盤中最大漲點</dt>
            <dd>當日最高點 − 前一交易日收盤</dd>
            <dt>盤中最大跌點</dt>
            <dd>前一交易日收盤 − 當日最低點</dd>
            <dt>法人連續買／賣超</dt>
            <dd>
              至少連續兩個交易日。零值、反向或已知缺日會中斷期間，以累計金額排名。
            </dd>
          </dl>
          <p className="muted">
            排名限2020年以來的日資料，不能據此證明2020年前沒有更大紀錄。
          </p>
          <div className="dialog-actions">
            <button onClick={() => setFormula(false)}>關閉</button>
          </div>
        </Modal>
      )}
    </>
  );
}
export function SettingsPage({
  timerDue,
  schedule,
  cancelSchedule,
  preferences: p,
  onChange,
  connected,
  onConnect,
  notify,
  reset,
}: {
  timerDue: string | null;
  schedule: () => void;
  cancelSchedule: () => void;
  preferences: Preferences;
  onChange: (v: Preferences) => void;
  connected: boolean;
  onConnect: (v: boolean) => void;
  notify: Notice;
  reset: () => void;
}) {
  const [tab, setTab] = useState("監控排程"),
    [log, setLog] = useState(false),
    [resetting, setResetting] = useState(false);
  function exportDemo() {
    const url = URL.createObjectURL(
      new Blob(
        [
          JSON.stringify(
            { format: "reporter-ui-demo", preferences: p },
            null,
            2,
          ),
        ],
        { type: "application/json" },
      ),
    );
    const a = document.createElement("a");
    a.href = url;
    a.download = "reporter-demo-settings.json";
    a.click();
    setTimeout(() => URL.revokeObjectURL(url), 0);
    notify("已匯出示意設定，不含正式資料");
  }
  return (
    <>
      <PageHeader title="設定" subtitle="依照工作節奏安排背景檢查。">
        <button onClick={() => notify("已儲存示意設定；正式排程未變更")}>
          儲存變更
        </button>
      </PageHeader>
      <Tabs
        labels={["監控排程", "通知", "啟動與系統匣", "資料管理"]}
        value={tab}
        onChange={setTab}
      />
      {tab === "監控排程" && (
        <div className="settings-groups">
          <section className="settings-group">
            <header>
              <h2>重大訊息</h2>
              <div className="switch-label">
                <Switch
                  label="啟用重大訊息"
                  checked={p.material}
                  onChange={(material) => onChange({ ...p, material })}
                />
                啟用
              </div>
            </header>
            <div className="schedule-fields">
              <label>
                開始時間
                <input
                  type="time"
                  value={p.start}
                  disabled={!p.material}
                  onChange={(e) => onChange({ ...p, start: e.target.value })}
                />
              </label>
              <label>
                結束時間
                <input
                  type="time"
                  value={p.end}
                  disabled={!p.material}
                  onChange={(e) => onChange({ ...p, end: e.target.value })}
                />
              </label>
              <label>
                檢查頻率
                <select
                  value={p.frequency}
                  disabled={!p.material}
                  onChange={(e) =>
                    onChange({ ...p, frequency: e.target.value })
                  }
                >
                  {["15", "30", "60", "120"].map((v) => (
                    <option value={v} key={v}>
                      {v} 分鐘
                    </option>
                  ))}
                </select>
              </label>
            </div>
            <p className="muted field-help">
              頻率依上次檢查起算，無新公告時不通知。
            </p>
          </section>
          <section className="settings-group">
            <header>
              <h2>違約交割</h2>
              <div className="switch-label">
                <Switch
                  label="啟用違約交割"
                  checked={p.disclosure}
                  onChange={(disclosure) => onChange({ ...p, disclosure })}
                />
                啟用
              </div>
            </header>
            <div className="disclosure-fields">
              <label>
                每日檢查時間
                <input
                  type="time"
                  disabled={!p.disclosure}
                  value={p.disclosureTime}
                  onChange={(e) =>
                    onChange({ ...p, disclosureTime: e.target.value })
                  }
                />
              </label>
              <label className="check-field">
                <input
                  type="checkbox"
                  checked={p.notifyEmpty}
                  onChange={(e) =>
                    onChange({ ...p, notifyEmpty: e.target.checked })
                  }
                />
                本日無揭露時也通知
              </label>
            </div>
          </section>
          <section className="settings-group">
            <header>
              <h2>法說會</h2>
            </header>
            <dl className="detail-list">
              <div>
                <dt>排程</dt>
                <dd>
                  每日 19:00 <b className="divider">|</b>{" "}
                  <span className="positive">v2 排程已就緒</span>
                </dd>
                <TextAction onClick={() => setLog(true)}>
                  檢視工作排程
                </TextAction>
              </div>
              <div>
                <dt>連線狀態</dt>
                <dd>
                  Google Calendar {connected ? "已連線" : "尚未連線"}{" "}
                  <b className="divider">|</b>{" "}
                  <TextAction
                    onClick={() => {
                      onConnect(!connected);
                      notify("已切換示意連線狀態");
                    }}
                  >
                    管理連線
                  </TextAction>
                </dd>
              </div>
            </dl>
          </section>
        </div>
      )}
      {tab === "通知" && (
        <div className="settings-groups">
          <section className="settings-group">
            <header>
              <h2>通知偏好</h2>
              <Switch
                label="啟用通知"
                checked={p.notifications}
                onChange={(notifications) => onChange({ ...p, notifications })}
              />
            </header>
            <p>有新公告或盤面紀錄時提醒。</p>
            <label className="check-field">
              <input
                type="checkbox"
                checked={p.notifyEmpty}
                onChange={(e) =>
                  onChange({ ...p, notifyEmpty: e.target.checked })
                }
              />
              本日無違約揭露時也通知
            </label>
          </section>
          <section className="settings-group">
            <header>
              <h2>測試通知</h2>
            </header>
            <p className="muted">Demo 只在畫面顯示回饋，不會發送系統通知。</p>
            <div className="button-row">
              <button
                disabled={!p.notifications}
                onClick={() =>
                  notify("示意通知：通知操作正常，未發送 Windows 通知")
                }
              >
                發送測試通知
              </button>
              <button
                className="secondary"
                disabled={!p.notifications}
                onClick={schedule}
              >
                一分鐘後發送測試通知
              </button>
            </div>
            {timerDue && (
              <p>
                已排定於 {timerDue}
                <button className="text-action" onClick={cancelSchedule}>
                  取消排定
                </button>
              </p>
            )}
          </section>
        </div>
      )}
      {tab === "啟動與系統匣" && (
        <div className="settings-groups">
          <section className="settings-group">
            <header>
              <h2>Windows 登入啟動</h2>
            </header>
            <div className="preference-row">
              <div>
                <h3>登入 Windows 時啟動</h3>
                <p className="muted">登入後自動開始背景監控。</p>
              </div>
              <Switch
                label="登入 Windows 時啟動"
                checked={p.login}
                onChange={(login) => onChange({ ...p, login })}
              />
            </div>
            <div className="preference-row">
              <div>
                <h3>啟動後縮到系統匣</h3>
                <p className="muted">先在背景執行，需要時再開啟視窗。</p>
              </div>
              <Switch
                label="啟動後縮到系統匣"
                checked={p.hidden}
                onChange={(hidden) => onChange({ ...p, hidden })}
              />
            </div>
          </section>
          <p className="muted">
            這些控制只改變Demo狀態，不會修改Windows登入項目。
          </p>
        </div>
      )}
      {tab === "資料管理" && (
        <div className="settings-groups">
          <section className="settings-group">
            <header>
              <h2>匯入與匯出</h2>
            </header>
            <p className="muted">本機示意資料會在重新整理時重置。</p>
            <div className="button-row">
              <button onClick={exportDemo}>匯出示意設定</button>
              <FileDemo
                label="選擇匯入檔案"
                onSelect={(name) =>
                  notify(`已選擇 ${name}；Demo 未寫入正式資料`)
                }
              />
            </div>
          </section>
          <section className="settings-group">
            <header>
              <h2>重置Demo</h2>
            </header>
            <p className="muted">還原所有示意公司、公告及設定。</p>
            <button className="secondary" onClick={() => setResetting(true)}>
              重置示意資料
            </button>
          </section>
        </div>
      )}
      <footer className="page-foot">
        <span>已儲存的排程會在背景執行。</span>
        <TextAction onClick={() => setLog(true)}>查看最近執行紀錄</TextAction>
      </footer>
      {log && (
        <Modal title="示意排程紀錄" onClose={() => setLog(false)}>
          <dl className="detail-list">
            <div>
              <dt>重大訊息</dt>
              <dd>15:00 檢查完成</dd>
            </div>
            <div>
              <dt>違約交割</dt>
              <dd>每日 {p.disclosureTime}</dd>
            </div>
            <div>
              <dt>法說會</dt>
              <dd>每日 19:00，v2 排程已就緒</dd>
            </div>
          </dl>
          <p className="muted">未讀取真實 Windows 工作或執行紀錄。</p>
          <div className="dialog-actions">
            <button onClick={() => setLog(false)}>關閉</button>
          </div>
        </Modal>
      )}
      {resetting && (
        <Modal title="重置示意資料" onClose={() => setResetting(false)}>
          <p>只還原此Demo，正式資料不受影響。</p>
          <div className="dialog-actions">
            <button className="secondary" onClick={() => setResetting(false)}>
              取消
            </button>
            <button
              onClick={() => {
                setResetting(false);
                reset();
              }}
            >
              確認重置
            </button>
          </div>
        </Modal>
      )}
    </>
  );
}
