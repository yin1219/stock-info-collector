export type GuidePage =
  | "overview"
  | "watchlist"
  | "material-events"
  | "disclosures"
  | "conferences"
  | "market-records"
  | "settings";
export type GuidePreference = "complete" | "skipped";
export interface GuideStep {
  target: string;
  title: string;
  description: string;
}
export interface TourStep extends GuideStep {
  page: GuidePage;
}
export const GUIDE_STORAGE_KEY = "reporter.guided-onboarding.v1";

export function readGuidePreference(
  storage?: Pick<Storage, "getItem">,
): GuidePreference | null {
  try {
    const value = JSON.parse(
      (storage ?? window.localStorage).getItem(GUIDE_STORAGE_KEY) ?? "null",
    );
    return value?.version === 1 &&
      ["complete", "skipped"].includes(value.status)
      ? value.status
      : null;
  } catch {
    return null;
  }
}
export function saveGuidePreference(
  storage: Pick<Storage, "setItem"> | undefined,
  status: GuidePreference,
): boolean {
  try {
    (storage ?? window.localStorage).setItem(
      GUIDE_STORAGE_KEY,
      JSON.stringify({ version: 1, status }),
    );
    return true;
  } catch {
    return false;
  }
}

export const pageGuides: Record<
  GuidePage,
  { title: string; summary: string; steps: GuideStep[] }
> = {
  overview: {
    title: "總覽",
    summary: "先看今天的監控狀態，再進入需要處理的資訊。",
    steps: [
      {
        target: '[data-guide="page-header"]',
        title: "掌握今日狀態",
        description:
          "總覽整理關注公司公告、違約揭露與近期法說會，協助安排今天的採訪工作。",
      },
      {
        target: ".overview-columns",
        title: "前往需要的資訊",
        description:
          "從各區的查看入口進入完整清單。重大訊息預設只列啟用中的關注公司。",
      },
      {
        target: ".page-foot",
        title: "留意來源與檢查時間",
        description:
          "請一起查看資料日期、上次檢查與來源狀態。尚未檢查或來源降級時，仍可能有未取得的資訊。",
      },
    ],
  },
  watchlist: {
    title: "關注公司",
    summary: "管理公告通知與法說會需要追蹤的公司。",
    steps: [
      {
        target: '[data-guide="page-actions"]',
        title: "建立關注公司",
        description:
          "使用新增關注公司輸入股號，或匯入舊版股票清單。完成後才開始追蹤這些公司的公告與法說會。",
      },
      {
        target: ".list-toolbar",
        title: "尋找與篩選公司",
        description:
          "可搜尋股號或公司名稱，也可切換監控中與已停用，確認目前要接收通知的公司。",
      },
      {
        target: ".table-foot",
        title: "停用與移除的影響",
        description:
          "停用會保留公司設定；移除會從關注清單消失。已保存的公告仍保留，關注清單影響後續通知與法說會查詢。",
      },
    ],
  },
  "material-events": {
    title: "重大訊息",
    summary: "閱讀關注公司的公告，辨識來源完整性與通知範圍。",
    steps: [
      {
        target: ".list-toolbar",
        title: "關注公司與全部公告",
        description:
          "預設顯示啟用中的關注公司。顯示全部公告可查本機保存的其他公司，通知仍只針對關注公司。",
      },
      {
        target: ".material-list",
        title: "整列展開公告",
        description:
          "點擊公告摘要列可原位展開全文，切換另一列會收起前一列。可標為已讀，或開啟官方來源。",
      },
      {
        target: '[data-guide="page-actions"]',
        title: "檢查、刪除與重新通知",
        description:
          "立即檢查會重新抓取來源。本機刪除經確認後隱藏該筆，官方再次取得時可重新通知。來源降級時仍可能漏筆。",
      },
    ],
  },
  disclosures: {
    title: "違約交割",
    summary: "查全市場個股違約揭露，先辨識資料是否已確認。",
    steps: [
      {
        target: ".disclosure-toolbar",
        title: "選擇日期與市場",
        description:
          "可依交易日期查看上市、上櫃或全部市場，立即檢查會重新取得兩個市場的官方揭露。",
      },
      {
        target: ".disclosure-result",
        title: "零筆與未確認",
        description:
          "只有當日有效資料確認零筆時，才顯示當日無個股達違約資訊揭露標準。資料過期或取得失敗時顯示未確認。",
      },
      {
        target: ".page-foot",
        title: "查看來源日期",
        description:
          "請核對各市場的來源日期與檢查結果。官方資料更新可能延遲，來源失敗與沒有揭露是不同狀態。",
      },
    ],
  },
  conferences: {
    title: "法說會",
    summary: "將關注公司的未來法說會同步到主要 Google Calendar。",
    steps: [
      {
        target: ".calendar-connection",
        title: "連線 Google Calendar",
        description:
          "先按連線並在瀏覽器完成 Google 授權。中斷連線會清除 v2 本機授權，之後可重新連線。",
      },
      {
        target: '[data-guide="page-actions"]',
        title: "同步與避免重複",
        description:
          "同步只處理關注公司的未來場次，每次核對主要行事曆相同時段與摘要。已存在就保留，被刪除的活動可補建。",
      },
      {
        target: ".conference-schedule",
        title: "確認每日排程",
        description:
          "頁面會顯示目前由 v1 或 v2 管理。v2 每日 19:00 由 Windows 工作執行，請在設定確認唯一排程已就緒。",
      },
    ],
  },
  "market-records": {
    title: "盤面紀錄",
    summary: "以保存的日資料查看指數漲跌與三大法人連買連賣紀錄。",
    steps: [
      {
        target: '[data-guide="page-header"]',
        title: "2020 年以來的排名",
        description:
          "排名依已保存的 2020 年以來日資料計算。缺日可能影響名次與連續期間，資料範圍不代表完整歷史。",
      },
      {
        target: ".market-source",
        title: "官方日資料與確認狀態",
        description:
          "程式背景運作時，15:30 起每小時檢查 TWSE 指數。今日官方日資料出現後才盤後確認，請留意來源日期與已知缺日。",
      },
      {
        target: '[data-guide="page-actions"]',
        title: "用 CMoney 補齊歷史",
        description:
          "資料管理可分別匯入指數與法人 CMoney Excel，也可匯入自行下載的 TWSE 法人 CSV。補匯後重算排名，同日指數優先保留官方資料。",
      },
    ],
  },
  settings: {
    title: "設定",
    summary: "依工作節奏安排監控、通知與登入啟動。",
    steps: [
      {
        target: '[role="tablist"][aria-label="設定分類"]',
        title: "四類設定",
        description:
          "監控排程、通知、啟動與系統匣、資料管理分開管理。修改監控排程後，請按儲存變更。",
      },
      {
        target: '[role="tablist"][aria-label="設定分類"]',
        title: "背景檢查與系統匣",
        description:
          "重大訊息頻率由前次檢查起算。關窗後留在系統匣可繼續監控，從系統匣結束才停止。登入啟動可在啟動與系統匣調整。",
      },
      {
        target: '[role="tablist"][aria-label="設定分類"]',
        title: "測試通知與保存資料",
        description:
          "通知分頁可立即測試或排定一分鐘後通知，方便驗證關窗留匣。資料管理可匯出本機資料，授權與 token 會排除。",
      },
    ],
  },
};
export const quickTour: TourStep[] = [
  { page: "watchlist", ...pageGuides.watchlist.steps[0] },
  { page: "material-events", ...pageGuides["material-events"].steps[0] },
  { page: "disclosures", ...pageGuides.disclosures.steps[1] },
  {
    page: "conferences",
    target: ".calendar-connection",
    title: "連線與同步 Google Calendar",
    description:
      "先連線並在瀏覽器完成 Google 授權，再手動同步未來法說會。每次同步會核對相同時段與摘要，已有活動會保留，被刪除的活動可補建。",
  },
  { page: "settings", ...pageGuides.settings.steps[1] },
];
