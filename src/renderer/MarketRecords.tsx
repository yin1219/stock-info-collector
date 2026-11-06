import { useEffect, useRef, useState } from "react";
import { CMoneyImportGuide } from "./CMoneyImportGuide";
import { PageHeader, Tabs, Modal, TextAction } from "./ui";
import type { Investor } from "../domain/market-records";
import type {
  MarketImportKind,
  MarketSnapshot,
} from "../domain/market-records";
interface Api {
  getMarketRecords(): Promise<unknown>;
  refreshMarketRecords(): Promise<unknown>;
  importMarketRecords(kind: MarketImportKind): Promise<unknown>;
}
export function MarketRecords({ api }: { api: Api }) {
  const [data, setData] = useState<MarketSnapshot | null>(null),
    [busy, setBusy] = useState(false),
    [message, setMessage] = useState("");
  const [tab, setTab] = useState("加權指數"),
    [investor, setInvestor] = useState<Investor>("trust"),
    [direction, setDirection] = useState<"rise" | "fall">("rise"),
    [date, setDate] = useState(""),
    [manage, setManage] = useState(false),
    [formula, setFormula] = useState(false);
  const dataRequest = useRef(0);
  const operationBusy = useRef(false);
  useEffect(() => {
    let live = true;
    let reading = false;
    const readSnapshot = () => {
      if (reading || operationBusy.current) return;
      reading = true;
      const requestId = ++dataRequest.current;
      void api
        .getMarketRecords()
        .then((value) => {
          if (live && requestId === dataRequest.current) {
            const snapshot = value as MarketSnapshot;
            setData(snapshot);
            setDate((current) => current || snapshot.today);
          }
        })
        .catch((error) => {
          if (live && requestId === dataRequest.current)
            setMessage(String(error));
        })
        .finally(() => {
          reading = false;
        });
    };
    const readWhenVisible = () => {
      if (document.visibilityState === "visible") readSnapshot();
    };
    readSnapshot();
    const timer = window.setInterval(readWhenVisible, 60_000);
    window.addEventListener("focus", readWhenVisible);
    document.addEventListener("visibilitychange", readWhenVisible);
    return () => {
      live = false;
      window.clearInterval(timer);
      window.removeEventListener("focus", readWhenVisible);
      document.removeEventListener("visibilitychange", readWhenVisible);
    };
  }, [api]);
  async function run(kind: MarketImportKind | "refresh") {
    if (operationBusy.current) return;
    operationBusy.current = true;
    dataRequest.current += 1;
    setBusy(true);
    setMessage(kind === "refresh" ? "正在更新官方指數…" : "請選擇匯入檔案…");
    try {
      if (kind === "refresh") {
        const updated = (await api.refreshMarketRecords()) as MarketSnapshot;
        setData(updated);
        setDate((current) => current || updated.today);
        setMessage(
          updated.error
            ? "官方指數檢查未完成，請查看來源狀態"
            : "官方指數檢查完成",
        );
      } else {
        const result = (await api.importMarketRecords(kind)) as {
          status: string;
          count?: number;
        };
        if (result.status === "cancelled") setMessage("已取消匯入");
        else {
          const updated = (await api.getMarketRecords()) as MarketSnapshot;
          setData(updated);
          setDate((current) => current || updated.today);
          setMessage(`已匯入 ${result.count} 筆日資料`);
        }
      }
    } catch (e) {
      setMessage(e instanceof Error ? e.message : String(e));
    } finally {
      operationBusy.current = false;
      setBusy(false);
    }
  }
  const number = (value: number) =>
    value.toLocaleString("zh-TW", {
      minimumFractionDigits: 2,
      maximumFractionDigits: 2,
    });
  const labels: Record<Investor, string> = {
    foreign: "外資",
    trust: "投信",
    dealer: "自營商",
  };
  const latest = data?.latestFlows?.find((row) => row.investor === investor);
  const latestStreak = latest?.streak;
  const selected = data
    ? {
        rise: data.indices.rise.find((row) => row.date === date),
        fall: data.indices.fall.find((row) => row.date === date),
      }
    : null;
  const feedback = message && (
    <p role="status" aria-live="polite">
      {message}
    </p>
  );
  return (
    <section aria-label="盤面紀錄資料">
      <PageHeader title="盤面紀錄" subtitle="以 2020 年以來的日資料比對">
        {tab === "加權指數" && (
          <input
            className="date-control"
            aria-label="盤面資料日期"
            type="date"
            value={date}
            max={data?.today}
            onChange={(event) => setDate(event.target.value)}
          />
        )}
        <button disabled={busy} onClick={() => void run("refresh")}>
          更新官方指數
        </button>
        <button className="text-action" onClick={() => setManage(true)}>
          資料管理
        </button>
      </PageHeader>
      {!manage && feedback}
      <Tabs labels={["加權指數", "三大法人"]} value={tab} onChange={setTab} />
      {!data ? (
        <p role="status">正在載入盤面日資料…</p>
      ) : (
        <>
          {tab === "加權指數" && data.indexUpdate && (
            <section aria-label="指數資料更新" className="market-source">
              <p>
                {data.indexUpdate.automatic
                  ? "每日 15:30 起每小時自動檢查 OpenAPI"
                  : "隔離預覽使用本機資料，自動更新停用"}
                {" · 官方資料日 "}
                {data.indexUpdate.latestOfficialDate ?? "尚無"}
                {" · 最近檢查 "}
                {data.indexUpdate.checkedAt
                  ? new Date(data.indexUpdate.checkedAt).toLocaleString(
                      "zh-TW",
                      { timeZone: "Asia/Taipei", hour12: false },
                    )
                  : "本次啟動尚未檢查"}
              </p>
              {data.indexUpdate.knownMissingDates.length > 0 && (
                <p className="source-warning">
                  已知指數缺 {data.indexUpdate.knownMissingDates.length}{" "}
                  個交易日 （
                  {data.indexUpdate.knownMissingDates.slice(0, 3).join("、")}
                  {data.indexUpdate.knownMissingDates.length > 3 ? "等" : ""}
                  ）。
                  <TextAction onClick={() => setManage(true)}>
                    匯入 CMoney 補齊
                  </TextAction>
                </p>
              )}
            </section>
          )}
          {!data.coverage.start && (
            <p className="empty-line">
              尚未匯入歷史日資料；不能判定是否有創紀錄。
            </p>
          )}
          {data.error && (
            <p role="alert" className="source-warning">
              官方指數來源失敗：{data.error}；保留既有資料。
            </p>
          )}
          {data.foreignWarning && (
            <p role="alert" className="source-warning">
              {data.foreignWarning}
            </p>
          )}
          {data.manualRequired && (
            <p className="source-warning">
              法人日資料待更新，請下載當日 CSV 後匯入。
            </p>
          )}
          {tab === "加權指數" ? (
            <>
              <p className="data-confirmed">
                {date === data.today
                  ? data.todayConfirmed
                    ? "指數盤後確認"
                    : "指數待盤後確認"
                  : selected?.rise || selected?.fall
                    ? "歷史盤後資料"
                    : "此日期尚無可計算資料"}
              </p>
              <div className="index-metrics">
                {(["rise", "fall"] as const).map((kind) => (
                  <div key={kind}>
                    <h2>盤中最大{kind === "rise" ? "漲" : "跌"}點</h2>
                    <p
                      className={`metric-value ${kind === "rise" ? "gain" : "loss"}`}
                    >
                      {selected?.[kind]
                        ? `${kind === "rise" ? "+" : "−"}${number(selected[kind]!.amount)}`
                        : "—"}
                      <small>點</small>
                    </p>
                    <p className="metric-rank">
                      {selected?.[kind]
                        ? `2020 年以來第 ${selected[kind]!.rank} 名`
                        : date === data.today && !data.todayConfirmed
                          ? "待盤後確認"
                          : "未形成正點數紀錄，或缺少前一交易日收盤，請核對日資料"}
                    </p>
                  </div>
                ))}
              </div>
              <div className="section-title rankings-title">
                <h2>歷史排名</h2>
                <Tabs
                  small
                  labels={["最大漲點", "最大跌點"]}
                  value={direction === "rise" ? "最大漲點" : "最大跌點"}
                  onChange={(value) =>
                    setDirection(value === "最大漲點" ? "rise" : "fall")
                  }
                />
              </div>
              <div className="table-scroll">
                <table className="index-table">
                  <thead>
                    <tr>
                      <th>名次</th>
                      <th>日期</th>
                      <th className="right">
                        盤中最大{direction === "rise" ? "漲" : "跌"}點
                      </th>
                    </tr>
                  </thead>
                  <tbody>
                    {data.indices[direction]
                      .filter((row) => row.rank <= 3)
                      .map((row) => (
                        <tr
                          key={row.date}
                          className={row.date === date ? "highlight" : ""}
                        >
                          <td className="strong">
                            {String(row.rank).padStart(2, "0")}
                          </td>
                          <td>{row.date}</td>
                          <td
                            className={`right strong ${direction === "rise" ? "gain" : "loss"}`}
                          >
                            {direction === "rise" ? "+" : "−"}
                            {number(row.amount)} 點
                          </td>
                        </tr>
                      ))}
                  </tbody>
                </table>
              </div>
              {!data.indices[direction].length && (
                <p className="empty-line">尚無可計算排名</p>
              )}
            </>
          ) : (
            <>
              <Tabs
                small
                labels={["外資", "投信", "自營商"]}
                value={labels[investor]}
                onChange={(value) =>
                  setInvestor(
                    value === "外資"
                      ? "foreign"
                      : value === "投信"
                        ? "trust"
                        : "dealer",
                  )
                }
              />
              <div className="investor-metrics">
                <div>
                  <h2>
                    {labels[investor]}最新連續
                    {latest?.direction === "sell" ? "賣" : "買"}超
                  </h2>
                  <p
                    className={`metric-value ${latest?.direction === "sell" ? "loss" : "gain"}`}
                  >
                    {latestStreak?.days ?? "—"}
                    <small>個交易日</small>
                  </p>
                </div>
                <div>
                  <h2>累計金額</h2>
                  <p
                    className={`metric-value ${latest?.direction === "sell" ? "loss" : "gain"}`}
                  >
                    {latestStreak ? number(latestStreak.amount / 1e8) : "—"}
                    <small>億元</small>
                  </p>
                </div>
                <p className="metric-rank">
                  {latestStreak
                    ? `2020 年以來第 ${latestStreak.rank} 名`
                    : "尚無可計算連續期間"}
                </p>
              </div>
              <p className="statistics-date">
                統計至 {latest?.dataDate ?? "尚無已確認資料"} ·
                僅使用已確認且口徑相符的資料
              </p>
              <div className="investor-tables">
                {(["buy", "sell"] as const).map((kind) => {
                  const group = data.flowGroups.find(
                    (group) => group.metric === `${investor}-${kind}`,
                  );
                  return (
                    <section key={kind}>
                      <h2>
                        連續{kind === "buy" ? "買" : "賣"}超前三名{" "}
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
                            {group?.rows.map((row) => (
                              <tr
                                key={row.start}
                                className={
                                  row.end === latest?.dataDate
                                    ? "highlight"
                                    : ""
                                }
                              >
                                <td className="strong">
                                  {String(row.rank).padStart(2, "0")}
                                </td>
                                <td className="period">
                                  {row.start}—<br />
                                  {row.end}
                                </td>
                                <td>{row.days} 日</td>
                                <td
                                  className={`right strong ${kind === "buy" ? "gain" : "loss"}`}
                                >
                                  {number(row.amount / 1e8)}
                                </td>
                              </tr>
                            ))}
                          </tbody>
                        </table>
                      </div>
                      {!group?.rows.length && (
                        <p className="empty-line">尚無可計算期間</p>
                      )}
                      <p className="market-source">
                        來源：{group?.source ?? "尚無資料"}
                      </p>
                    </section>
                  );
                })}
              </div>
            </>
          )}
          <footer className="page-foot">
            <span>
              資料範圍 {data.coverage.start ?? "尚無"}—
              {data.coverage.end ?? "尚無"} · 來源依日資料保存
            </span>
            <TextAction onClick={() => setFormula(true)}>
              查看計算方式
            </TextAction>
          </footer>
        </>
      )}
      {manage && (
        <Modal title="盤面資料管理" onClose={() => setManage(false)} wide>
          {feedback}
          <CMoneyImportGuide
            renderImport={(kind, label) => (
              <button disabled={busy} onClick={() => void run(kind)}>
                匯入{label} Excel
              </button>
            )}
          />
          <h3>每日法人更新（TWSE）</h3>
          <p>
            <a
              href="https://www.twse.com.tw/zh/trading/foreign/bfi82u.html"
              target="_blank"
              rel="noreferrer"
            >
              前往 TWSE 下載當日 CSV
            </a>
          </p>
          <button disabled={busy} onClick={() => void run("csv")}>
            匯入 TWSE 當日 CSV
          </button>
          {data && (
            <details>
              <summary>來源與資料完整性</summary>
              <p>
                指數資料來源為金融監督管理委員會證券期貨局／臺灣證券交易所的
                <a
                  href="https://data.gov.tw/dataset/11755"
                  target="_blank"
                  rel="noreferrer"
                >
                  發行量加權股價指數歷史資料
                </a>
                ， 依
                <a
                  href="https://data.gov.tw/license"
                  target="_blank"
                  rel="noreferrer"
                >
                  政府資料開放授權條款第1版
                </a>
                提供。
              </p>
              <p>{data.sourceStatus}</p>
              <p>歷史資料完整性依匯入檔案；系統不會以缺資料推定無紀錄。</p>
              <p>修訂紀錄 {data.revisions ?? 0} 筆</p>
            </details>
          )}
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
              至少連續兩個交易日。零值、反向或已知缺日中斷，以累計金額排名。僅以保存且口徑相符的日資料計算。
            </dd>
          </dl>
          <p className="muted">
            2020 年以來的排名不代表2020年前沒有更大的紀錄。
          </p>
          <div className="dialog-actions">
            <button onClick={() => setFormula(false)}>關閉</button>
          </div>
        </Modal>
      )}
    </section>
  );
}
