import { useEffect, useState } from "react";
import type {
  GoogleCalendarStatus,
  LoginStartupSettings,
  ReporterApi,
  ScheduleSettings,
} from "../preload/api";
import type { GuidePage } from "./onboarding";
import { Modal } from "./ui";

export type SetupApi = Pick<
  ReporterApi,
  | "listWatchlist"
  | "getScheduleSettings"
  | "getLoginStartupSettings"
  | "getGoogleCalendarStatus"
  | "getMarketRecords"
>;
export type SetupTab = "監控排程" | "通知" | "啟動與系統匣";
interface SetupData {
  watchlist: unknown[] | null;
  schedule: ScheduleSettings | null;
  startup: LoginStartupSettings | null;
  google: GoogleCalendarStatus | null;
  market: { coverage?: { start: string | null; end: string | null } } | null;
}
export function SetupChecklist({
  api,
  onNavigate,
  onClose,
}: {
  api: SetupApi;
  onNavigate: (page: GuidePage, tab?: SetupTab) => void;
  onClose: () => void;
}) {
  const [data, setData] = useState<SetupData | null>(null);
  const [revision, setRevision] = useState(0);
  useEffect(() => {
    let current = true;
    setData(null);
    void Promise.allSettled([
      api.listWatchlist(),
      api.getScheduleSettings(),
      api.getLoginStartupSettings(),
      api.getGoogleCalendarStatus(),
      api.getMarketRecords(),
    ]).then(([watchlist, schedule, startup, google, market]) => {
      if (!current) return;
      setData({
        watchlist:
          watchlist.status === "fulfilled" && Array.isArray(watchlist.value)
            ? watchlist.value
            : null,
        schedule: schedule.status === "fulfilled" ? schedule.value : null,
        startup: startup.status === "fulfilled" ? startup.value : null,
        google: google.status === "fulfilled" ? google.value : null,
        market:
          market.status === "fulfilled"
            ? (market.value as SetupData["market"])
            : null,
      });
    });
    return () => {
      current = false;
    };
  }, [api, revision]);
  const activeCompanies =
    data?.watchlist?.filter(
      (row) =>
        row &&
        typeof row === "object" &&
        "active" in row &&
        row.active === true,
    ).length ?? 0;
  const unknown = "無法讀取目前狀態";
  const monitoring = data?.schedule?.monitoring;
  const startup = data?.startup;
  const coverage = data?.market?.coverage;
  const items: {
    title: string;
    status: string;
    detail: string;
    action: string;
    page: GuidePage;
    tab?: SetupTab;
  }[] = [
    {
      title: "關注公司",
      status: data?.watchlist
        ? activeCompanies
          ? `已啟用 ${activeCompanies} 家公司`
          : "尚無啟用中的公司"
        : unknown,
      detail:
        "新增或匯入要追蹤的股號。啟用中的公司決定重大訊息通知與法說會查詢範圍。",
      action: "管理關注公司",
      page: "watchlist",
    },
    {
      title: "檢查排程",
      status: monitoring
        ? monitoring.enabled
          ? "重大訊息監控已啟用"
          : "重大訊息監控已停用"
        : unknown,
      detail: monitoring
        ? `重大訊息 ${monitoring.start}–${monitoring.end}，每 ${monitoring.intervalMinutes} 分鐘檢查。違約交割${data?.schedule?.disclosure.enabled ? "每日 " + data.schedule.disclosure.runAt : "目前停用"}。請核對時段並儲存變更。`
        : "請確認重大訊息的開始／結束時間、頻率，以及違約交割每日檢查時間。",
      action: "設定檢查排程",
      page: "settings",
      tab: "監控排程",
    },
    {
      title: "登入與背景運作",
      status: !startup
        ? unknown
        : startup.blockedByWindows
          ? "Windows 已阻擋登入啟動"
          : startup.enabled
            ? "登入啟動已開啟"
            : "登入啟動尚未開啟",
      detail: startup
        ? `啟動後縮到系統匣${startup.startHidden ? "已開啟" : "尚未開啟"}。關窗留在系統匣可持續監控；從系統匣結束會停止。${startup.blockedByWindows ? "請至 Windows 啟動應用程式設定允許啟動。" : ""}`
        : "可設定登入 Windows 時啟動與啟動後縮到系統匣，依使用習慣選擇。",
      action: "設定背景啟動",
      page: "settings",
      tab: "啟動與系統匣",
    },
    {
      title: "Google Calendar",
      status: !data?.google
        ? unknown
        : data.google.status === "connected"
          ? "Google 已連線"
          : data.google.status === "reauthorization-required"
            ? "需要重新授權"
            : data.google.status === "not-configured"
              ? "應用程式 Google 設定尚未提供"
              : "尚未連線",
      detail:
        "需要同步法說會時，再前往連線並在瀏覽器授權。請同時確認每日工作狀態與關注公司。",
      action: "設定 Google Calendar",
      page: "conferences",
    },
    {
      title: "Windows 通知",
      status: "需手動測試",
      detail:
        "在通知分頁發送測試通知，確認 Windows 有顯示並可點回程式。也可排定一分鐘後通知，再關窗測試系統匣運作。",
      action: "前往測試通知",
      page: "settings",
      tab: "通知",
    },
    {
      title: "盤面歷史資料",
      status: !data?.market
        ? unknown
        : coverage?.start
          ? "已有日資料"
          : "尚未匯入日資料",
      detail: coverage?.start
        ? `${coverage.start}—${coverage.end ?? "尚無結束日期"}。已有日資料不代表歷史完整，請核對缺日並用 CMoney 補匯。`
        : "需要盤面紀錄時，再至資料管理匯入 CMoney 指數與法人 Excel。每日指數由官方更新，歷史缺日仍需補齊。",
      action: "匯入歷史日資料",
      page: "market-records",
    },
  ];
  return (
    <Modal title="首次使用設定" onClose={onClose}>
      <p className="guide-intro">
        先核對基本設定，再依採訪需求連線行事曆或匯入歷史資料。每個項目都可稍後回來調整。
      </p>
      <button
        className="text-action"
        onClick={() => setRevision((value) => value + 1)}
      >
        重新檢查
      </button>
      {!data && (
        <p role="status" className="setup-loading">
          正在讀取目前設定…
        </p>
      )}
      {[
        { title: "先確認這三項", rows: items.slice(0, 3) },
        { title: "依工作需要選擇", rows: items.slice(3) },
      ].map((group) => (
        <section key={group.title} className="setup-group">
          <h3>{group.title}</h3>
          <ol className="setup-list">
            {group.rows.map((item) => (
              <li key={item.title}>
                <section aria-label={item.title}>
                  <h4>{item.title}</h4>
                  <p className="setup-state">
                    {data ? item.status : "正在確認…"}
                  </p>
                  <p>{item.detail}</p>
                  <button
                    className="text-action"
                    onClick={() => {
                      onClose();
                      onNavigate(item.page, item.tab);
                    }}
                  >
                    {item.action}
                  </button>
                </section>
              </li>
            ))}
          </ol>
        </section>
      ))}
      <div className="dialog-actions guide-actions">
        <button onClick={onClose}>先使用小幫手</button>
      </div>
    </Modal>
  );
}
