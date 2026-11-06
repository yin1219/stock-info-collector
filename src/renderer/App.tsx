import {
  useCallback,
  useEffect,
  useRef,
  useState,
  type FormEvent,
} from "react";
import { GuideProvider, GuideWelcome, GuideLauncher } from "./Guide";
import { MarketRecords } from "./MarketRecords";
import { MaterialSourceReport } from "./SourceReport";
import { Icon, PageHeader, TextAction, Switch, Tabs, Modal } from "./ui";
import type {
  ConferenceTaskStatus,
  GoogleCalendarStatus,
  LegacyConferenceConfigReview,
  LoginStartupSettings,
  ReporterRoute,
  ScheduleSettings,
} from "../preload/api";

interface WatchlistRow {
  id: string;
  companyId: string;
  market: "TWSE" | "TPEX";
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
  market: "TWSE" | "TPEX";
  stockCode: string;
  companyName: string | null;
  disclosureDate: string;
  brokerCode: string;
  isWatched: boolean;
  contentJson: string;
}

interface DisclosureMonitorStatus {
  status: "active" | "complete" | "degraded" | "stale" | "failed";
  finishedAt: string | null;
  errorMessage: string | null;
  markets: Record<
    string,
    {
      status: string;
      dataDate: string | null;
      recordCount: number;
      retryAt: string | null;
      errorMessage?: string;
    }
  >;
}

interface MaterialEventRow {
  id: string;
  companyId: string;
  market: "TWSE" | "TPEX";
  stockCode: string;
  companyName: string;
  title: string;
  content: string;
  source: string;
  sourceUrl: string;
  publishedAt: string;
  discoveredAt: string;
  readAt: string | null;
  eventType: "announcement" | "correction" | "supplement";
  revisionOf: string | null;
}

interface MaterialEventDetail extends MaterialEventRow {
  original: MaterialEventRow | null;
  relatedRevisions: MaterialEventRow[];
}

interface MaterialMonitorStatus {
  status: "active" | "complete" | "degraded" | "stale" | "failed";
  finishedAt: string | null;
  errorMessage: string | null;
  sources: Array<{
    source: string;
    status: string;
    dataDate: string | null;
    recordCount?: number;
    errorMessage: string | null;
  }>;
}

interface ConferenceRow {
  id: string;
  stockCode: string;
  companyName: string;
  startsAt: string;
  location: string;
  content: string;
  sourceUrl?: string;
  syncStatus: "pending" | "existing" | "synced" | "failed";
  lastError: string | null;
}

type ScheduleJobStatus = {
  enabled: boolean;
  status: string;
  lastSuccessAt: string | null;
  lastResult: string | null;
  errorMessage: string | null;
  nextRunAt: string | null;
};
type ScheduleStatus = {
  available?: boolean;
  material: ScheduleJobStatus;
  disclosure: ScheduleJobStatus;
};
type Page =
  | "overview"
  | "watchlist"
  | "disclosures"
  | "material-events"
  | "conferences"
  | "settings"
  | "market-records";

function disclosureStatusLabel(status: string): string {
  switch (status) {
    case "complete":
      return "資料完整";
    case "empty_success":
      return "本日無揭露";
    case "stale":
      return "官方尚未更新";
    case "failed":
      return "來源失敗";
    case "active":
      return "檢查中";
    default:
      return status;
  }
}

function disclosureEmptyMessage(
  monitor: DisclosureMonitorStatus | null,
  date: string,
  market: "all" | "TWSE" | "TPEX",
): string {
  const selectedMarkets = market === "all" ? ["TWSE", "TPEX"] : [market];
  const states = selectedMarkets.map((source) => monitor?.markets[source]);
  if (isConfirmedEmpty(monitor, date, market)) {
    return "當日無個股達違約資訊揭露標準";
  }
  if (
    states.some(
      (state) => state?.status === "failed" || state?.status === "stale",
    ) ||
    monitor?.status === "failed" ||
    monitor?.status === "stale"
  ) {
    return "來源檢查失敗或官方資料尚未更新，不能判定為零筆。";
  }
  return "此日期目前沒有可顯示的違約交割資料，尚未確認是否為零筆。";
}

function isConfirmedEmpty(
  monitor: DisclosureMonitorStatus | null,
  date: string,
  market: "all" | "TWSE" | "TPEX",
): boolean {
  const selectedMarkets = market === "all" ? ["TWSE", "TPEX"] : [market];
  return selectedMarkets.every(
    (source) =>
      monitor?.markets[source]?.status === "empty_success" &&
      monitor.markets[source].dataDate === date,
  );
}

function disclosureCountLabel(
  count: number,
  monitor: DisclosureMonitorStatus | null,
  date: string,
  market: "all" | "TWSE" | "TPEX",
): string {
  if (count > 0 || isConfirmedEmpty(monitor, date, market))
    return `${count} 筆`;
  return monitor ? "未確認" : "尚未檢查";
}

function formatTaipeiDateTime(value: string): string {
  const instant = new Date(value);
  if (Number.isNaN(instant.getTime())) return "時間待確認";
  return new Intl.DateTimeFormat("zh-TW", {
    timeZone: "Asia/Taipei",
    year: "numeric",
    month: "2-digit",
    day: "2-digit",
    hour: "2-digit",
    minute: "2-digit",
    hourCycle: "h23",
  }).format(instant);
}

function materialSourceName(source: string): string {
  if (source === "TWSE-reconciliation") return "上市結構化每日對帳";
  if (source === "TPEX-reconciliation") return "上櫃結構化每日對帳";
  return source;
}

function monitorSourceStatusLabel(status: string): string {
  if (status === "complete") return "資料完整";
  if (status === "degraded") return "部分降級";
  if (status === "stale") return "官方尚未更新";
  if (status === "failed") return "來源失敗";
  return status;
}

function taipeiDateToday(): string {
  const parts = new Intl.DateTimeFormat("en-CA", {
    timeZone: "Asia/Taipei",
    year: "numeric",
    month: "2-digit",
    day: "2-digit",
  }).formatToParts(new Date());
  const values = Object.fromEntries(
    parts.map(({ type, value }) => [type, value]),
  );
  return `${values.year}-${values.month}-${values.day}`;
}

function conferenceTaskHealth(
  status: ConferenceTaskStatus | null,
): { message: string; role: "status" | "alert" } | null {
  if (
    (status?.kind !== "legacy" && status?.kind !== "v2") ||
    !status.lastTaskResult
  )
    return null;
  const owner = status.kind === "legacy" ? "既有 v1" : "v2";
  if (status.lastTaskResult === 0x41303) {
    return {
      message: `${owner} 法說會工作尚未執行；請在首次排程後再核對結果。`,
      role: "status",
    };
  }
  const code = status.lastTaskResult
    .toString(16)
    .padStart(8, "0")
    .toUpperCase();
  return {
    message: `${owner} 法說會工作最近一次 Windows 工作結果異常（0x${code}）；請檢查工作排程器${status.kind === "legacy" ? "，v2 不會建立第二個工作" : ""}。`,
    role: "alert",
  };
}

export function App() {
  const [appVersion, setAppVersion] = useState<string | null>(null);
  const [rows, setRows] = useState<WatchlistRow[]>([]);
  const [query, setQuery] = useState("");
  const [activeOnly, setActiveOnly] = useState(false);
  const [market, setMarket] = useState<"TWSE" | "TPEX">("TWSE");
  const [stockCode, setStockCode] = useState("");
  const [error, setError] = useState("");
  const [watchlistSyncMessage, setWatchlistSyncMessage] = useState("");
  const [configText, setConfigText] = useState("");
  const [importReport, setImportReport] = useState<ImportReport | null>(null);
  const [busy, setBusy] = useState(false);
  const [page, setPage] = useState<Page>("overview");
  const [navOpen, setNavOpen] = useState(false);
  const [settingsTab, setSettingsTab] = useState("監控排程");
  const [addingCompany, setAddingCompany] = useState(false);
  const [editingCompany, setEditingCompany] = useState<WatchlistRow | null>(
    null,
  );
  const [removingCompany, setRemovingCompany] = useState<WatchlistRow | null>(
    null,
  );
  const [importingWatchlist, setImportingWatchlist] = useState(false);
  const [watchlistView, setWatchlistView] = useState("全部");
  const [allRows, setAllRows] = useState<WatchlistRow[]>([]);
  const [showSources, setShowSources] = useState(false);
  const [overviewRefresh, setOverviewRefresh] = useState(0);
  const [disclosures, setDisclosures] = useState<DisclosureRow[]>([]);
  const [disclosureMonitorStatus, setDisclosureMonitorStatus] =
    useState<DisclosureMonitorStatus | null>(null);
  const [disclosureDate, setDisclosureDate] = useState(taipeiDateToday);
  const [disclosureMarket, setDisclosureMarket] = useState<
    "all" | "TWSE" | "TPEX"
  >("all");
  const [materialEvents, setMaterialEvents] = useState<MaterialEventRow[]>([]);
  const [materialQuery, setMaterialQuery] = useState("");
  const [unreadOnly, setUnreadOnly] = useState(false);
  const [showAllMaterial, setShowAllMaterial] = useState(false);
  const [expandedMaterialRowId, setExpandedMaterialRowId] = useState<
    string | null
  >(null);
  const [selectedMaterialEvent, setSelectedMaterialEvent] =
    useState<MaterialEventDetail | null>(null);
  const [materialActionMessage, setMaterialActionMessage] = useState("");
  const [materialDeleteBusy, setMaterialDeleteBusy] = useState(false);
  const materialDetailRequest = useRef(0);
  const [notificationEventIds, setNotificationEventIds] = useState<
    string[] | null
  >(null);
  const [materialMonitorStatus, setMaterialMonitorStatus] =
    useState<MaterialMonitorStatus | null>(null);
  const [scheduleSettings, setScheduleSettings] =
    useState<ScheduleSettings | null>(null);
  const [scheduleStatus, setScheduleStatus] = useState<ScheduleStatus | null>(
    null,
  );
  const [scheduleMessage, setScheduleMessage] = useState("");
  const [scheduleMessageKind, setScheduleMessageKind] = useState<
    "material" | "disclosure" | null
  >(null);
  const [dataExportMessage, setDataExportMessage] = useState("");
  const [dataExportFailed, setDataExportFailed] = useState(false);
  const [dataExportBusy, setDataExportBusy] = useState(false);
  const [testNotificationMessage, setTestNotificationMessage] = useState("");
  const [testNotificationBusy, setTestNotificationBusy] = useState(false);
  const [runningSchedule, setRunningSchedule] = useState<
    "material" | "disclosure" | null
  >(null);
  const [loginStartupSettings, setLoginStartupSettings] =
    useState<LoginStartupSettings | null>(null);
  const [loginStartupBusy, setLoginStartupBusy] = useState(false);
  const [loginStartupMessage, setLoginStartupMessage] = useState("");
  const [overviewDisclosureCount, setOverviewDisclosureCount] = useState<
    number | null
  >(null);
  const [overviewMaterialCount, setOverviewMaterialCount] = useState<
    number | null
  >(null);
  const [overviewMaterialEvents, setOverviewMaterialEvents] = useState<
    MaterialEventRow[]
  >([]);
  const [overviewConferences, setOverviewConferences] = useState<
    ConferenceRow[]
  >([]);
  const [overviewScheduleStatus, setOverviewScheduleStatus] =
    useState<ScheduleStatus | null>(null);
  const [googleCalendarStatus, setGoogleCalendarStatus] =
    useState<GoogleCalendarStatus | null>(null);
  const [conferenceRows, setConferenceRows] = useState<ConferenceRow[]>([]);
  const [conferenceBusy, setConferenceBusy] = useState(false);
  const [conferenceMessage, setConferenceMessage] = useState("");
  const [conferenceTaskStatus, setConferenceTaskStatus] =
    useState<ConferenceTaskStatus | null>(null);
  const [legacyConfigReview, setLegacyConfigReview] =
    useState<LegacyConferenceConfigReview | null>(null);
  const [conferenceTaskBusy, setConferenceTaskBusy] = useState(false);
  const [conferenceTaskMessage, setConferenceTaskMessage] = useState("");

  useEffect(() => {
    void window.reporterApi
      .getVersion()
      .then(setAppVersion)
      .catch(() => setAppVersion(null));
  }, []);

  useEffect(() => {
    if (page !== "settings") return;
    void Promise.all([
      window.reporterApi.getScheduleSettings(),
      window.reporterApi.getScheduleStatus(),
      window.reporterApi.getLoginStartupSettings(),
      window.reporterApi.getGoogleCalendarStatus(),
    ])
      .then(([settings, status, startup, google]) => {
        setScheduleSettings(settings);
        setScheduleStatus(status as ScheduleStatus);
        setLoginStartupSettings(startup);
        setGoogleCalendarStatus(google);
      })
      .catch((cause: unknown) =>
        setError(cause instanceof Error ? cause.message : String(cause)),
      );
    setConferenceTaskStatus(null);
    setLegacyConfigReview(null);
    void window.reporterApi
      .getConferenceTaskStatus()
      .then(setConferenceTaskStatus)
      .catch((cause: unknown) =>
        setConferenceTaskStatus({
          kind: "blocked",
          reason: cause instanceof Error ? cause.message : String(cause),
        }),
      );
  }, [page]);

  useEffect(() => {
    if (page !== "overview") return;
    void Promise.all([
      window.reporterApi.listDisclosures({ disclosureDate: taipeiDateToday() }),
      window.reporterApi.listMaterialEvents({
        unreadOnly: true,
        watchedOnly: true,
      }),
      window.reporterApi.getMaterialMonitorStatus(),
      window.reporterApi.getDisclosureMonitorStatus(),
      window.reporterApi.listConferences(),
      window.reporterApi.getScheduleStatus(),
    ])
      .then(
        ([
          disclosureRows,
          materialRows,
          materialStatus,
          disclosureStatus,
          conferenceRows,
          scheduleStatus,
        ]) => {
          setOverviewDisclosureCount(
            (disclosureRows as DisclosureRow[]).length,
          );
          setOverviewMaterialCount((materialRows as MaterialEventRow[]).length);
          setOverviewMaterialEvents(materialRows as MaterialEventRow[]);
          const upcomingConferences = (
            conferenceRows as ConferenceRow[]
          ).filter((conference) => {
            const startsAt = new Date(conference.startsAt).getTime();
            return (
              startsAt >= Date.now() &&
              startsAt <= Date.now() + 14 * 24 * 60 * 60 * 1000
            );
          });
          setOverviewConferences(upcomingConferences);
          setOverviewScheduleStatus(scheduleStatus as ScheduleStatus);
          setMaterialMonitorStatus(
            materialStatus as MaterialMonitorStatus | null,
          );
          setDisclosureMonitorStatus(
            disclosureStatus as DisclosureMonitorStatus | null,
          );
        },
      )
      .catch((cause: unknown) =>
        setError(cause instanceof Error ? cause.message : String(cause)),
      );
  }, [page, overviewRefresh]);

  useEffect(() => {
    if (page !== "conferences") return;
    setConferenceTaskStatus(null);
    void Promise.all([
      window.reporterApi.getGoogleCalendarStatus(),
      window.reporterApi.listConferences(),
    ])
      .then(([status, rows]) => {
        setGoogleCalendarStatus(status);
        setConferenceRows(rows as ConferenceRow[]);
      })
      .catch((cause: unknown) =>
        setError(cause instanceof Error ? cause.message : String(cause)),
      );
    void window.reporterApi
      .getConferenceTaskStatus()
      .then(setConferenceTaskStatus)
      .catch((cause: unknown) =>
        setConferenceTaskStatus({
          kind: "blocked",
          reason: cause instanceof Error ? cause.message : String(cause),
        }),
      );
  }, [page]);

  const refresh = useCallback(async () => {
    const [result, all] = await Promise.all([
      window.reporterApi.listWatchlist({ query, activeOnly }),
      window.reporterApi.listWatchlist(),
    ]);
    setRows(result as WatchlistRow[]);
    setAllRows(all as WatchlistRow[]);
  }, [query, activeOnly]);

  useEffect(() => {
    // A legacy-config warning can follow a committed SQLite removal.
    // Close only after the refreshed list confirms the company is gone.
    if (
      removingCompany &&
      !busy &&
      !allRows.some((row) => row.companyId === removingCompany.companyId)
    ) {
      setRemovingCompany(null);
    }
  }, [allRows, busy, removingCompany]);

  async function reportLegacyWatchlistSync() {
    try {
      const task = await window.reporterApi.getConferenceTaskStatus();
      if (
        task.kind !== "legacy" &&
        task.kind !== "legacy-disabled" &&
        !(task.kind === "v2" && task.legacyConfigPath)
      )
        return;
      const review = await window.reporterApi.reviewLegacyConferenceConfig();
      const current = [...review.currentCodes].sort();
      const active = [...review.activeCodes].sort();
      if (
        !review.confirmed ||
        current.length !== active.length ||
        current.some((code, index) => code !== active[index])
      ) {
        throw new Error("設定與啟用股號不一致，請到「設定」重新核對");
      }
      setWatchlistSyncMessage("v1 法說會股號已套用");
    } catch (cause) {
      setError(
        `無法確認 v1 法說會股號已套用：${cause instanceof Error ? cause.message : String(cause)}`,
      );
    }
  }

  useEffect(() => {
    if (page !== "material-events") return;
    void window.reporterApi
      .getMaterialMonitorStatus()
      .then((result) =>
        setMaterialMonitorStatus(result as MaterialMonitorStatus | null),
      )
      .catch((cause: unknown) =>
        setError(cause instanceof Error ? cause.message : String(cause)),
      );
    void window.reporterApi
      .listMaterialEvents({
        query: materialQuery,
        unreadOnly,
        ...(showAllMaterial ? {} : { watchedOnly: true }),
        ...(notificationEventIds ? { eventIds: notificationEventIds } : {}),
      })
      .then((result) => {
        setMaterialEvents(result as MaterialEventRow[]);
      })
      .catch((cause: unknown) =>
        setError(cause instanceof Error ? cause.message : String(cause)),
      );
  }, [page, materialQuery, unreadOnly, showAllMaterial, notificationEventIds]);

  useEffect(
    () =>
      window.reporterApi.onNotificationRoute((route: ReporterRoute) => {
        setError("");
        if (route.type === "market-records") {
          setPage("market-records");
          return;
        }
        if (route.type === "test-notification") {
          setPage("settings");
          setSettingsTab("通知");
          setTestNotificationMessage("已從測試通知返回設定。");
          return;
        }
        if (route.type === "disclosure-list") {
          setNotificationEventIds(null);
          setPage("disclosures");
          const today = taipeiDateToday();
          setDisclosureDate(today);
          setDisclosureMarket("all");
          void loadDisclosures(today, "all");
          return;
        }
        setPage("material-events");
        setMaterialQuery("");
        setUnreadOnly(false);
        if (route.type === "source-status") {
          setShowSources(true);
          setNotificationEventIds(null);
          setSelectedMaterialEvent(null);
          setExpandedMaterialRowId(null);
          return;
        }
        if (route.type === "event-list") {
          setNotificationEventIds(route.eventIds);
          setSelectedMaterialEvent(null);
          setExpandedMaterialRowId(null);
          return;
        }
        setNotificationEventIds([route.eventId]);
        setExpandedMaterialRowId(route.eventId);
        void window.reporterApi
          .getMaterialEvent(route.eventId)
          .then((result) =>
            setSelectedMaterialEvent(result as MaterialEventDetail),
          );
      }),
    [],
  );

  useEffect(() => {
    if (page !== "watchlist" && page !== "overview") return;
    void refresh().catch((cause: unknown) =>
      setError(cause instanceof Error ? cause.message : String(cause)),
    );
  }, [refresh, page]);

  async function addCompany(event: FormEvent<HTMLFormElement>) {
    event.preventDefault();
    const normalizedCode = stockCode.trim();
    if (!/^\d{4,6}$/.test(normalizedCode)) {
      setError("股票代號必須是 4 至 6 位數字");
      return;
    }
    setBusy(true);
    setError("");
    setWatchlistSyncMessage("");
    try {
      await window.reporterApi.addWatchlist({
        market,
        stockCode: normalizedCode,
      });
      setStockCode("");
      setAddingCompany(false);
      await refresh();
      await reportLegacyWatchlistSync();
    } catch (cause) {
      await refresh().catch(() => undefined);
      setError(cause instanceof Error ? cause.message : String(cause));
    } finally {
      setBusy(false);
    }
  }

  async function previewImport() {
    setError("");
    try {
      setImportReport(
        (await window.reporterApi.previewWatchlistImport(
          configText,
        )) as ImportReport,
      );
    } catch (cause) {
      setError(cause instanceof Error ? cause.message : String(cause));
    }
  }

  async function applyImport() {
    setBusy(true);
    setError("");
    setWatchlistSyncMessage("");
    try {
      setImportReport(
        (await window.reporterApi.applyWatchlistImport(
          configText,
        )) as ImportReport,
      );
      await refresh();
      await reportLegacyWatchlistSync();
    } catch (cause) {
      await refresh().catch(() => undefined);
      setError(cause instanceof Error ? cause.message : String(cause));
    } finally {
      setBusy(false);
    }
  }

  async function changeWatchlist(
    mutation: () => Promise<unknown>,
  ): Promise<boolean> {
    if (busy) return false;
    setBusy(true);
    setError("");
    setWatchlistSyncMessage("");
    try {
      await mutation();
      await reportLegacyWatchlistSync();
      return true;
    } catch (cause) {
      setError(cause instanceof Error ? cause.message : String(cause));
      return false;
    } finally {
      await refresh().catch(() => undefined);
      setBusy(false);
    }
  }

  async function loadDisclosures(
    date = disclosureDate,
    selectedMarket = disclosureMarket,
  ) {
    setError("");
    try {
      const [result, status] = await Promise.all([
        window.reporterApi.listDisclosures({
          disclosureDate: date,
          ...(selectedMarket === "all" ? {} : { market: selectedMarket }),
        }),
        window.reporterApi.getDisclosureMonitorStatus(),
      ]);
      setDisclosures(result as DisclosureRow[]);
      setDisclosureMonitorStatus(status as DisclosureMonitorStatus | null);
    } catch (cause) {
      setError(cause instanceof Error ? cause.message : String(cause));
    }
  }

  async function runSchedule(kind: "material" | "disclosure") {
    const title = kind === "material" ? "重大訊息" : "違約交割";
    setRunningSchedule(kind);
    setScheduleMessageKind(kind);
    setScheduleMessage(`正在檢查${title}…`);
    try {
      const result = (await window.reporterApi.runScheduledCheck(kind)) as {
        status: string;
      };
      const latest =
        (await window.reporterApi.getScheduleStatus()) as ScheduleStatus;
      setScheduleStatus(latest);
      const job = latest[kind];
      setScheduleMessage(
        result.status === "already-running"
          ? "此工作目前仍在執行，未重複啟動。"
          : job.status === "failed"
            ? `${title}檢查失敗：${job.errorMessage ?? "來源錯誤"}`
            : job.status === "stale"
              ? `${title}官方資料尚未更新。`
              : job.status === "degraded"
                ? kind === "material"
                  ? "重大訊息來源部分降級；資料可能不完整，請查看已保存公告與來源狀態。"
                  : `${title}部分來源檢查失敗，請查看最新狀態。`
                : `${title}檢查完成，請查看最新狀態。`,
      );
      if (kind === "disclosure") await loadDisclosures();
      else {
        setMaterialEvents(
          (await window.reporterApi.listMaterialEvents({
            query: materialQuery,
            unreadOnly,
            ...(showAllMaterial ? {} : { watchedOnly: true }),
            ...(notificationEventIds ? { eventIds: notificationEventIds } : {}),
          })) as MaterialEventRow[],
        );
        setMaterialMonitorStatus(
          (await window.reporterApi.getMaterialMonitorStatus()) as MaterialMonitorStatus | null,
        );
      }
      setOverviewRefresh((value) => value + 1);
    } catch (cause) {
      setScheduleMessage(
        cause instanceof Error ? cause.message : String(cause),
      );
    } finally {
      setRunningSchedule(null);
    }
  }

  async function saveSchedule() {
    if (!scheduleSettings) return;
    setError("");
    try {
      setScheduleSettings(
        await window.reporterApi.saveScheduleSettings(scheduleSettings),
      );
      setScheduleMessage("排程設定已儲存。");
      setScheduleStatus(
        (await window.reporterApi.getScheduleStatus()) as ScheduleStatus,
      );
    } catch (cause) {
      setError(cause instanceof Error ? cause.message : String(cause));
    }
  }

  async function reviewLegacyConferenceConfig() {
    setConferenceTaskBusy(true);
    setConferenceTaskMessage("");
    try {
      setLegacyConfigReview(
        await window.reporterApi.reviewLegacyConferenceConfig(),
      );
    } catch (cause) {
      setLegacyConfigReview(null);
      setConferenceTaskMessage(
        `核對舊版股號失敗：${cause instanceof Error ? cause.message : String(cause)}`,
      );
    } finally {
      setConferenceTaskBusy(false);
    }
  }

  async function confirmLegacyConferenceConfig() {
    if (!legacyConfigReview) return;
    setConferenceTaskBusy(true);
    setConferenceTaskMessage("");
    try {
      await window.reporterApi.confirmLegacyConferenceConfig(
        legacyConfigReview.hash,
      );
      setLegacyConfigReview(
        await window.reporterApi.reviewLegacyConferenceConfig(),
      );
      setConferenceTaskMessage(
        "已套用到舊版工作；之後啟用的關注股號會同步至已確認的舊版設定。",
      );
    } catch (cause) {
      setLegacyConfigReview(null);
      setConferenceTaskMessage(
        `套用舊版股號失敗：${cause instanceof Error ? cause.message : String(cause)}；請重新核對。`,
      );
    } finally {
      setConferenceTaskBusy(false);
    }
  }

  async function ensureConferenceTask() {
    setConferenceTaskBusy(true);
    setConferenceTaskMessage("");
    try {
      const status = await window.reporterApi.ensureConferenceTask();
      setConferenceTaskStatus(status);
      setConferenceTaskMessage(
        status.kind === "v2"
          ? "v2 每日工作已就緒。"
          : (status.reason ?? "每日工作尚未就緒。"),
      );
    } catch (cause) {
      setConferenceTaskMessage(
        `建立每日工作失敗：${cause instanceof Error ? cause.message : String(cause)}`,
      );
    } finally {
      setConferenceTaskBusy(false);
    }
  }

  async function switchFromDisabledLegacy() {
    if (
      !window.confirm(
        "已完成正式法說會與 Google Calendar 驗收，確認切換為 v2 每日 19:00 排程嗎？已停用的 v1 工作會保留供回退。",
      )
    )
      return;
    setConferenceTaskBusy(true);
    setConferenceTaskMessage("");
    try {
      const status = await window.reporterApi.switchFromDisabledLegacy();
      setConferenceTaskStatus(status);
      setConferenceTaskMessage(
        status.kind === "v2"
          ? "v2 每日工作已就緒。"
          : (status.reason ?? "切換未完成；請重新確認 Windows 工作狀態。"),
      );
    } catch (cause) {
      setConferenceTaskMessage(
        `切換每日工作失敗：${cause instanceof Error ? cause.message : String(cause)}`,
      );
    } finally {
      setConferenceTaskBusy(false);
    }
  }

  async function saveLoginStartup(next: LoginStartupSettings) {
    if (loginStartupBusy) return;
    setLoginStartupBusy(true);
    setLoginStartupMessage("正在儲存登入啟動設定…");
    try {
      const actual = await window.reporterApi.saveLoginStartupSettings(next);
      setLoginStartupSettings(actual);
      setLoginStartupMessage(
        actual.blockedByWindows
          ? "已登錄啟動項，但 Windows 已停用此應用程式的登入啟動。"
          : "登入啟動設定已儲存。",
      );
    } catch (cause) {
      setLoginStartupMessage(
        `登入啟動設定失敗：${cause instanceof Error ? cause.message : String(cause)}`,
      );
    } finally {
      setLoginStartupBusy(false);
    }
  }

  async function sendTestNotification() {
    setTestNotificationBusy(true);
    setTestNotificationMessage("");
    try {
      const result = await window.reporterApi.sendTestNotification();
      setTestNotificationMessage(
        result.status === "suppressed"
          ? "隔離測試模式不顯示 Windows 通知。"
          : "Windows 已接受測試通知顯示要求；若未看到，請檢查系統通知設定。",
      );
    } catch (cause) {
      setTestNotificationMessage(
        `測試通知失敗：${cause instanceof Error ? cause.message : String(cause)}`,
      );
    } finally {
      setTestNotificationBusy(false);
    }
  }

  async function scheduleTestNotification() {
    setTestNotificationBusy(true);
    setTestNotificationMessage("");
    try {
      const result = await window.reporterApi.scheduleTestNotification();
      setTestNotificationMessage(
        result.status === "suppressed"
          ? "隔離測試模式不顯示 Windows 通知。"
          : `已排定 ${new Date(result.scheduledAt).toLocaleTimeString("zh-TW", { timeZone: "Asia/Taipei" })} 發送測試通知。可關閉視窗並讓程式留在系統匣；若選擇「結束」就不會發送。`,
      );
    } catch (cause) {
      setTestNotificationMessage(
        `測試通知排程失敗：${cause instanceof Error ? cause.message : String(cause)}`,
      );
    } finally {
      setTestNotificationBusy(false);
    }
  }

  async function exportLocalData() {
    setDataExportBusy(true);
    setDataExportMessage("");
    setDataExportFailed(false);
    try {
      const result = await window.reporterApi.exportUserData();
      setDataExportMessage(
        result.status === "cancelled"
          ? "已取消資料匯出。"
          : `資料已匯出至 ${result.destination}`,
      );
    } catch (cause) {
      setDataExportFailed(true);
      setDataExportMessage(
        cause instanceof Error
          ? cause.message
          : "資料匯出失敗，請查看診斷紀錄。",
      );
    } finally {
      setDataExportBusy(false);
    }
  }

  async function importGoogleCredentials() {
    setError("");
    try {
      const imported = await window.reporterApi.importGoogleCredentials();
      if (imported) {
        setGoogleCalendarStatus(
          await window.reporterApi.getGoogleCalendarStatus(),
        );
        setConferenceMessage("OAuth 設定已安全匯入。");
      }
    } catch (cause) {
      setError(cause instanceof Error ? cause.message : String(cause));
    }
  }

  async function migrateLegacyGoogleToken() {
    setError("");
    try {
      const migrated = await window.reporterApi.migrateLegacyGoogleToken();
      if (migrated) {
        setGoogleCalendarStatus(
          await window.reporterApi.getGoogleCalendarStatus(),
        );
        setConferenceMessage("舊版 token 已安全遷移；原始檔案仍保留。");
      }
    } catch (cause) {
      setError(cause instanceof Error ? cause.message : String(cause));
    }
  }

  async function connectGoogleCalendar() {
    setConferenceBusy(true);
    setError("");
    setConferenceMessage(
      "請在系統瀏覽器完成 Google 授權；完成後此頁會更新連線狀態。",
    );
    try {
      setGoogleCalendarStatus(await window.reporterApi.connectGoogleCalendar());
      setConferenceMessage("Google Calendar 已連線。");
    } catch (cause) {
      setConferenceMessage("");
      setError(cause instanceof Error ? cause.message : String(cause));
      setGoogleCalendarStatus(
        await window.reporterApi.getGoogleCalendarStatus(),
      );
    } finally {
      setConferenceBusy(false);
    }
  }

  async function disconnectGoogleCalendar() {
    setConferenceBusy(true);
    setError("");
    try {
      setGoogleCalendarStatus(
        await window.reporterApi.disconnectGoogleCalendar(),
      );
      setConferenceMessage("已中斷 Google Calendar 連線。");
    } catch (cause) {
      setError(cause instanceof Error ? cause.message : String(cause));
      setGoogleCalendarStatus(
        await window.reporterApi.getGoogleCalendarStatus(),
      );
    } finally {
      setConferenceBusy(false);
    }
  }

  async function syncConferences() {
    setConferenceBusy(true);
    setError("");
    try {
      const result = (await window.reporterApi.syncConferences()) as {
        status: string;
        results: Array<{ status: string }>;
      };
      setGoogleCalendarStatus(
        await window.reporterApi.getGoogleCalendarStatus(),
      );
      setConferenceRows(
        (await window.reporterApi.listConferences()) as ConferenceRow[],
      );
      const results = result.results ?? [];
      const failed = results.filter((row) => row.status === "failed").length;
      const added = results.filter((row) => row.status === "synced").length;
      const existing = results.filter(
        (row) => row.status === "existing",
      ).length;
      setConferenceMessage(
        result.status === "reauthorization-required"
          ? "Google 憑證失效，請重新授權。"
          : result.status === "already-running"
            ? "此工作目前仍在執行，未重複啟動。"
            : `${failed || result.status === "failed" || result.status === "partial" ? "法說會同步未全部成功" : "法說會同步完成"}：新增 ${added}、已存在 ${existing}、失敗 ${failed}（處理 ${results.length} 筆）。`,
      );
    } catch (cause) {
      setError(cause instanceof Error ? cause.message : String(cause));
    } finally {
      setConferenceBusy(false);
    }
  }

  async function openMaterialEvent(id: string, anchorId = id) {
    if (
      expandedMaterialRowId === anchorId &&
      selectedMaterialEvent?.id === id
    ) {
      materialDetailRequest.current += 1;
      setExpandedMaterialRowId(null);
      setSelectedMaterialEvent(null);
      return;
    }
    const requestId = ++materialDetailRequest.current;
    setExpandedMaterialRowId(anchorId);
    setSelectedMaterialEvent(null);
    setError("");
    try {
      const detail = (await window.reporterApi.getMaterialEvent(
        id,
      )) as MaterialEventDetail;
      if (requestId === materialDetailRequest.current)
        setSelectedMaterialEvent(detail);
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
    setSelectedMaterialEvent((current) =>
      current ? { ...current, readAt: new Date().toISOString() } : current,
    );
    await window.reporterApi
      .listMaterialEvents({
        query: materialQuery,
        unreadOnly,
        ...(showAllMaterial ? {} : { watchedOnly: true }),
        ...(notificationEventIds ? { eventIds: notificationEventIds } : {}),
      })
      .then((result) => setMaterialEvents(result as MaterialEventRow[]));
  }

  async function softDeleteSelectedMaterialEvent() {
    if (!selectedMaterialEvent || materialDeleteBusy) return;
    if (
      !window.confirm(
        "要從本機清單移除此筆公告嗎？原通知與刪除紀錄會保留；下次官方來源再次取得相同公告時，可能重新通知。",
      )
    )
      return;
    setError("");
    setMaterialActionMessage("");
    setMaterialDeleteBusy(true);
    try {
      await window.reporterApi.softDeleteMaterialEvent(
        selectedMaterialEvent.id,
      );
      materialDetailRequest.current += 1;
      setExpandedMaterialRowId(null);
      setSelectedMaterialEvent(null);
      const result = await window.reporterApi.listMaterialEvents({
        query: materialQuery,
        unreadOnly,
        ...(showAllMaterial ? {} : { watchedOnly: true }),
        ...(notificationEventIds ? { eventIds: notificationEventIds } : {}),
      });
      setMaterialEvents(result as MaterialEventRow[]);
      setMaterialActionMessage(
        "已從本機清單移除；下次官方來源再次取得此公告時，可能重新通知。",
      );
    } catch (cause) {
      setError(cause instanceof Error ? cause.message : String(cause));
    } finally {
      setMaterialDeleteBusy(false);
    }
  }

  const taskHealth = conferenceTaskHealth(conferenceTaskStatus);
  const pages: Array<[Page, string, string]> = [
    ["overview", "總覽", "home"],
    ["watchlist", "關注公司", "building"],
    ["material-events", "重大訊息", "megaphone"],
    ["disclosures", "違約交割", "warning"],
    ["conferences", "法說會", "conference"],
    ["market-records", "盤面紀錄", "bars"],
    ["settings", "設定", "gear"],
  ];
  const navigate = (next: Page, tab?: string) => {
    if (tab) setSettingsTab(tab);
    setPage(next);
    setNavOpen(false);
    if (next === "material-events") setNotificationEventIds(null);
    if (next === "disclosures") void loadDisclosures();
  };
  const visibleWatchlist = rows.filter(
    (row) => watchlistView !== "已停用" || !row.active,
  );
  const jobStatus =
    page === "overview"
      ? overviewScheduleStatus
      : (scheduleStatus ?? overviewScheduleStatus);
  const calendarLabel =
    googleCalendarStatus?.status === "connected"
      ? "Google 已連線"
      : googleCalendarStatus?.status === "reauthorization-required"
        ? "需要重新授權"
        : googleCalendarStatus?.status === "not-configured"
          ? "尚未設定"
          : "尚未連線";
  const taskText = !conferenceTaskStatus
    ? "正在確認法說會每日工作…"
    : conferenceTaskStatus.kind === "v2"
      ? "v2 法說會每日工作已就緒；Windows 每日 19:00 執行。"
      : conferenceTaskStatus.kind === "legacy"
        ? "既有 v1 工作仍是唯一受管法說會排程；v2 手動同步停用，避免同時寫入 Calendar。"
        : conferenceTaskStatus.kind === "legacy-disabled"
          ? "v1 工作已停用，目前尚無每日排程；可手動驗收法說會與 Calendar，驗收後才切換 v2。"
          : conferenceTaskStatus.kind === "upgrade-v2-time"
            ? "偵測到舊版 v2 每日 17:45 工作，正在遷移至 19:00；完成前不執行同步。"
            : conferenceTaskStatus.kind === "create-v2"
              ? "尚未建立 v2 法說會每日工作；完成設定後才可手動同步。"
              : `法說會每日工作未就緒：${conferenceTaskStatus.reason ?? "無法確認唯一寫入者"}`;
  const renderCheck = (kind: "material" | "disclosure") => (
    <button
      className="secondary"
      disabled={jobStatus?.available === false || runningSchedule !== null}
      onClick={() => void runSchedule(kind)}
    >
      {runningSchedule === kind
        ? "檢查中…"
        : `立即檢查${kind === "material" ? "重大訊息" : "違約交割"}`}
    </button>
  );
  const renderJob = (kind: "material" | "disclosure") => {
    const job = jobStatus?.[kind];
    return (
      <div className="job-feedback">
        {job && (
          <>
            <span>
              {job.status === "idle"
                ? "尚無執行紀錄"
                : job.status === "active"
                  ? "檢查中"
                  : monitorSourceStatusLabel(job.status)}
            </span>
            <span>
              預計下次：
              {jobStatus?.available === false
                ? "隔離測試模式不執行"
                : job.nextRunAt
                  ? formatTaipeiDateTime(job.nextRunAt)
                  : "尚未排定"}
            </span>
            {job.errorMessage && <p role="alert">{job.errorMessage}</p>}
          </>
        )}
        {scheduleMessageKind === kind && scheduleMessage && (
          <p role="status">{scheduleMessage}</p>
        )}
      </div>
    );
  };
  return (
    <GuideProvider page={page} onNavigate={navigate} api={window.reporterApi}>
    <div className="reporter-app">
      <button
        className="mobile-menu icon-button"
        aria-label="開啟導覽"
        onClick={() => setNavOpen(!navOpen)}
      >
        <Icon name="menu" />
      </button>
      {navOpen && (
        <button
          className="nav-scrim"
          aria-label="關閉導覽"
          onClick={() => setNavOpen(false)}
        />
      )}
      <aside className={navOpen ? "sidebar open" : "sidebar"}>
        <div className="brand">股市記者小幫手</div>
        <nav className="main-nav" aria-label="主要功能" role="tablist">
          {pages.map(([key, title, icon]) => (
            <button
              type="button"
              key={key}
              role="tab"
              aria-selected={page === key}
              className={page === key ? "nav-item selected" : "nav-item"}
              onClick={() => navigate(key)}
            >
              <Icon name={icon} size={27} />
              <span>{title}</span>
            </button>
          ))}
        </nav>
        <footer aria-label="背景執行狀態">
          <GuideLauncher />
          <p>
            <i className="status-dot" />
            {jobStatus?.available === false
              ? "隔離測試模式：背景監控已停用。"
              : "背景監控服務"}
          </p>
          <span className="version">
            v2 · {appVersion ? `版本 ${appVersion}` : "正在讀取版本…"}
          </span>
        </footer>
      </aside>
      <main className="reporter-main" id="reporter-content">
        <div className="demo-label">本機資料 · Asia/Taipei</div>
        <GuideWelcome />
        {error &&
          !addingCompany &&
          !editingCompany &&
          !removingCompany &&
          !importingWatchlist && (
            <p role="alert" className="error-message">
              {error}
            </p>
          )}
        {page === "market-records" ? (
          <MarketRecords api={window.reporterApi} />
        ) : page === "overview" ? (
          <>
            <PageHeader title="總覽" subtitle="今日監控與採訪行程">
              <span className="overview-date">{taipeiDateToday()}</span>
              {renderCheck("material")}
            </PageHeader>
            <p className="running-line">
              <i className="status-dot" />
              {overviewScheduleStatus?.available === false
                ? "隔離測試模式：背景監控已停用。"
                : "關閉視窗後會留在系統匣執行。"}
            </p>
            <div className="overview-summary">
              <div>
                關注公司{" "}
                <strong>{allRows.filter((row) => row.active).length}</strong> 家
              </div>
              <div>
                未讀重大訊息 <strong>{overviewMaterialCount ?? "—"}</strong> 筆
              </div>
              <div>
                今日違約揭露{" "}
                <span>
                  {overviewDisclosureCount === null
                    ? "載入中…"
                    : disclosureCountLabel(
                        overviewDisclosureCount,
                        disclosureMonitorStatus,
                        taipeiDateToday(),
                        "all",
                      )}
                </span>
              </div>
            </div>
            <div className="overview-columns">
              <section>
                <div className="section-title">
                  <h2>最新重大訊息</h2>
                  <TextAction onClick={() => navigate("material-events")}>
                    查看全部
                  </TextAction>
                </div>
                <div className="recent-list">
                  {overviewMaterialEvents.slice(0, 4).map((event) => (
                    <button
                      key={event.id}
                      onClick={() => {
                        navigate("material-events");
                        void openMaterialEvent(event.id);
                      }}
                    >
                      <time>{formatTaipeiDateTime(event.publishedAt)}</time>
                      <span className="recent-company">
                        <strong>{event.companyName}</strong> {event.stockCode}
                      </span>
                      <span>{event.title}</span>
                    </button>
                  ))}
                </div>
                {overviewMaterialEvents.length === 0 && (
                  <p className="empty-line">目前沒有未讀重大訊息。</p>
                )}
              </section>
              <aside className="overview-aside">
                <h2>下一場法說會</h2>
                {overviewConferences[0] ? (
                  <>
                    <div className="next-conference">
                      <div>
                        <strong>
                          {overviewConferences[0].companyName}{" "}
                          {overviewConferences[0].stockCode}
                        </strong>
                        <p>
                          {formatTaipeiDateTime(
                            overviewConferences[0].startsAt,
                          )}
                          <br />
                          {overviewConferences[0].location || "地點未提供"}
                        </p>
                      </div>
                    </div>
                    <TextAction onClick={() => navigate("conferences")}>
                      查看法說會
                    </TextAction>
                  </>
                ) : (
                  <p className="muted">未來 14 天沒有已保存場次。</p>
                )}
                <h2 className="next-check-heading">接下來的檢查</h2>
                {!overviewScheduleStatus ? (
                  <p role="status">正在載入排程狀態…</p>
                ) : (
                  <dl className="next-checks">
                    {(["material", "disclosure"] as const).map((kind) => (
                      <div key={kind}>
                        <dt>{kind === "material" ? "重大訊息" : "違約交割"}</dt>
                        <dd>
                          {overviewScheduleStatus[kind].nextRunAt
                            ? formatTaipeiDateTime(
                                overviewScheduleStatus[kind].nextRunAt!,
                              )
                            : "尚未排定"}
                        </dd>
                      </div>
                    ))}
                  </dl>
                )}
              </aside>
            </div>
            {renderJob("material")}
            <footer className="page-foot">
              <span>
                {materialMonitorStatus
                  ? `MOPS ${monitorSourceStatusLabel(materialMonitorStatus.status)}`
                  : "來源尚未檢查"}
              </span>
              <TextAction onClick={() => setShowSources(true)}>
                查看來源狀態
              </TextAction>
              <TextAction onClick={() => navigate("settings")}>設定</TextAction>
            </footer>
          </>
        ) : page === "watchlist" ? (
          <>
            <PageHeader
              title="關注公司"
              subtitle="管理需要追蹤公告與法說會的公司"
            >
              <button
                onClick={() => {
                  setAddingCompany(true);
                  setError("");
                }}
              >
                新增關注公司
              </button>
              <button
                className="secondary"
                onClick={() => setImportingWatchlist(true)}
              >
                匯入舊版股票清單
              </button>
            </PageHeader>
            <div className="list-toolbar">
              <label className="search-field">
                <Icon name="search" />
                <input
                  type="search"
                  aria-label="搜尋關注公司"
                  placeholder="搜尋股號或公司名稱"
                  value={query}
                  onChange={(event) => setQuery(event.target.value)}
                />
              </label>
              <div className="filter-segments">
                {["全部", "監控中", "已停用"].map((view) => (
                  <button
                    key={view}
                    aria-pressed={watchlistView === view}
                    onClick={() => {
                      setWatchlistView(view);
                      setActiveOnly(view === "監控中");
                    }}
                  >
                    {view}{" "}
                    {
                      allRows.filter(
                        (row) =>
                          view === "全部" ||
                          (view === "監控中" ? row.active : !row.active),
                      ).length
                    }
                  </button>
                ))}
              </div>
            </div>
            {watchlistSyncMessage && (
              <p role="status">{watchlistSyncMessage}</p>
            )}
            <div className="table-scroll">
              <table className="watch-table">
                <thead>
                  <tr>
                    {["股號", "公司名稱", "市場", "分類", "監控", "操作"].map(
                      (label) => (
                        <th key={label}>{label}</th>
                      ),
                    )}
                  </tr>
                </thead>
                <tbody>
                  {visibleWatchlist.map((row) => (
                    <tr key={row.id}>
                      <td className="strong">{row.stockCode}</td>
                      <td className="strong">{row.name}</td>
                      <td>{row.market === "TWSE" ? "上市" : "上櫃"}</td>
                      <td>{row.category || "—"}</td>
                      <td>
                        <Switch
                          label={`監控${row.name}`}
                          checked={row.active}
                          disabled={busy}
                          onChange={(active) =>
                            void changeWatchlist(() =>
                              window.reporterApi.setWatchlistActive({
                                companyId: row.companyId,
                                active,
                              }),
                            )
                          }
                        />
                      </td>
                      <td>
                        <div className="row-actions">
                          <button
                            className="icon-button"
                            aria-label={`編輯${row.name}`}
                            onClick={() => setEditingCompany({ ...row })}
                          >
                            <Icon name="edit" />
                          </button>
                          <button
                            className="icon-button"
                            aria-label={`移除${row.name}`}
                            onClick={() => setRemovingCompany(row)}
                          >
                            <Icon name="more" />
                          </button>
                        </div>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
            {!visibleWatchlist.length && (
              <p className="empty-line">目前沒有符合條件的公司。</p>
            )}
            <footer className="table-foot">
              <span>停用保留清單；移除不會刪除已保存的公告。</span>
              <span>顯示 {visibleWatchlist.length} 家公司</span>
            </footer>
          </>
        ) : page === "disclosures" ? (
          <section aria-label="違約交割資料">
            <PageHeader title="違約交割" subtitle="全市場個股違約資訊揭露" />
            <div className="disclosure-toolbar">
              <input
                type="date"
                aria-label="違約交割資料日期"
                value={disclosureDate}
                onChange={(event) => {
                  setDisclosureDate(event.target.value);
                  void loadDisclosures(event.target.value);
                }}
              />
              <select
                aria-label="違約交割市場"
                value={disclosureMarket}
                onChange={(event) => {
                  const value = event.target.value as typeof disclosureMarket;
                  setDisclosureMarket(value);
                  void loadDisclosures(disclosureDate, value);
                }}
              >
                <option value="all">全部市場</option>
                <option value="TWSE">上市</option>
                <option value="TPEX">上櫃</option>
              </select>
              {renderCheck("disclosure")}
            </div>
            {renderJob("disclosure")}
            {!disclosures.length ? (
              <section
                className={
                  isConfirmedEmpty(
                    disclosureMonitorStatus,
                    disclosureDate,
                    disclosureMarket,
                  )
                    ? "disclosure-result"
                    : "disclosure-result unconfirmed"
                }
              >
                <Icon
                  name={
                    isConfirmedEmpty(
                      disclosureMonitorStatus,
                      disclosureDate,
                      disclosureMarket,
                    )
                      ? "checkCircle"
                      : "warning"
                  }
                  size={50}
                />
                <div>
                  <h2>
                    {disclosureEmptyMessage(
                      disclosureMonitorStatus,
                      disclosureDate,
                      disclosureMarket,
                    )}
                  </h2>
                  <p>
                    {isConfirmedEmpty(
                      disclosureMonitorStatus,
                      disclosureDate,
                      disclosureMarket,
                    )
                      ? "所選市場來源均已完成當日檢查。"
                      : "來源日期落後或檢查失敗時，不視為無揭露。"}
                  </p>
                </div>
              </section>
            ) : (
              <div className="table-scroll">
                <table>
                  <thead>
                    <tr>
                      <th>公司／股號</th>
                      <th>市場</th>
                      <th>申報日期</th>
                      <th>券商</th>
                    </tr>
                  </thead>
                  <tbody>
                    {disclosures.map((row) => (
                      <tr
                        key={row.id}
                        className={row.isWatched ? "highlight" : ""}
                      >
                        <td>
                          <strong>{row.companyName ?? row.stockCode}</strong>
                          <br />
                          {row.stockCode}{" "}
                          {row.isWatched && (
                            <small className="watched-badge">關注公司</small>
                          )}
                          <details>
                            <summary>詳細資訊</summary>
                            <pre className="disclosure-content">
                              {row.contentJson}
                            </pre>
                          </details>
                        </td>
                        <td>{row.market === "TWSE" ? "上市" : "上櫃"}</td>
                        <td>{row.disclosureDate}</td>
                        <td>{row.brokerCode}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            )}
            <h2 id="disclosure-title">來源狀態</h2>
            <div
              className="table-scroll"
              role="status"
              aria-label="違約交割來源狀態"
            >
              <p>
                {disclosureMonitorStatus?.status === "degraded"
                  ? "資料來源部分降級"
                  : disclosureMonitorStatus
                    ? disclosureStatusLabel(disclosureMonitorStatus.status)
                    : "尚未檢查"}
              </p>
              <table className="source-table">
                <thead>
                  <tr>
                    <th>市場</th>
                    <th>資料日期</th>
                    <th>檢查結果</th>
                    <th>筆數／補查</th>
                  </tr>
                </thead>
                <tbody>
                  {["TWSE", "TPEX"]
                    .filter(
                      (market) =>
                        disclosureMarket === "all" ||
                        market === disclosureMarket,
                    )
                    .map((market) => {
                      const source = disclosureMonitorStatus?.markets[market];
                      return (
                        <tr key={market}>
                          <td>{market}</td>
                          <td>{source?.dataDate ?? "尚無資料日期"}</td>
                          <td>
                            {source
                              ? disclosureStatusLabel(source.status)
                              : "尚未檢查"}
                            {source?.errorMessage && (
                              <p role="alert">{source.errorMessage}</p>
                            )}
                          </td>
                          <td>
                            {source?.recordCount ?? "—"}
                            <br />
                            {source?.retryAt &&
                              `原定重試 ${formatTaipeiDateTime(source.retryAt)}`}
                          </td>
                        </tr>
                      );
                    })}
                </tbody>
              </table>
            </div>
            {disclosureMonitorStatus?.errorMessage && (
              <p role="alert">{disclosureMonitorStatus.errorMessage}</p>
            )}
            <footer className="page-foot">
              <span>
                來源失敗或日期落後時，會顯示「尚未確認」，不視為無揭露。
              </span>
            </footer>
          </section>
        ) : page === "conferences" ? (
          <>
            <PageHeader
              title="法說會"
              subtitle="關注公司的近期場次與行事曆同步"
            >
              {googleCalendarStatus?.status === "connected" ? (
                <button
                  disabled={
                    conferenceBusy ||
                    !["v2", "legacy-disabled"].includes(
                      conferenceTaskStatus?.kind ?? "",
                    )
                  }
                  onClick={() => void syncConferences()}
                >
                  {conferenceBusy ? "正在同步…" : "立即擷取並同步法說會"}
                </button>
              ) : (
                <button
                  disabled={conferenceBusy}
                  onClick={() => void connectGoogleCalendar()}
                >
                  {googleCalendarStatus?.status === "reauthorization-required"
                    ? "重新授權 Google Calendar"
                    : "連線 Google Calendar"}
                </button>
              )}
            </PageHeader>
            <div className="calendar-connection">
              <span>
                <Icon name="calendar" size={28} />
                Google Calendar{" "}
                <strong className="positive">{calendarLabel}</strong>
                <b className="divider">|</b>主要行事曆
              </span>
              {googleCalendarStatus?.status === "connected" && (
                <TextAction onClick={() => void disconnectGoogleCalendar()}>
                  中斷 Google Calendar
                </TextAction>
              )}
            </div>
            <div className="conference-schedule">
              <p role="status">{taskText}</p>
              <TextAction onClick={() => navigate("settings")}>
                查看排程
              </TextAction>
            </div>
            {taskHealth && <p role={taskHealth.role}>{taskHealth.message}</p>}
            {conferenceMessage && <p role="status">{conferenceMessage}</p>}
            <h2>已保存法說會</h2>
            {!conferenceRows.length ? (
              <p className="empty-line">
                目前沒有已保存的法說會。連線後可擷取關注公司的近期場次。
              </p>
            ) : (
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
                    {conferenceRows.map((row) => (
                      <tr key={row.id}>
                        <td>{formatTaipeiDateTime(row.startsAt)}</td>
                        <td>
                          <strong>{row.companyName}</strong>
                          <br />
                          {row.stockCode}
                        </td>
                        <td>
                          {row.location || "地點未提供"}
                          <br />
                          {row.sourceUrl && (
                            <a
                              href={row.sourceUrl}
                              target="_blank"
                              rel="noreferrer"
                            >
                              開啟法說會來源
                            </a>
                          )}
                        </td>
                        <td>
                          <span
                            className={
                              row.syncStatus === "synced" ? "positive" : ""
                            }
                          >
                            {row.syncStatus === "pending"
                              ? "待同步"
                              : row.syncStatus === "existing"
                                ? "已存在"
                                : row.syncStatus === "synced"
                                  ? "已同步"
                                  : "同步失敗"}
                          </span>
                          {row.lastError && <p role="alert">{row.lastError}</p>}
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            )}
            <details className="advanced-settings">
              <summary>舊版授權與進階設定</summary>
              <p>一般使用者請使用上方 Google 連線；以下僅供原有設定遷移。</p>
              {googleCalendarStatus?.status === "disconnected" && (
                <button
                  className="secondary"
                  disabled={conferenceBusy}
                  onClick={() => void migrateLegacyGoogleToken()}
                >
                  匯入舊版 token（不刪除原檔）
                </button>
              )}
              <button
                className="secondary"
                onClick={() => void importGoogleCredentials()}
              >
                匯入應用程式 OAuth 設定
              </button>
            </details>
            <footer className="page-foot">
              會先確認 Google 行事曆是否已有相同場次，避免重複新增。
            </footer>
          </>
        ) : page === "settings" ? (
          <>
            <PageHeader title="設定" subtitle="依照工作節奏安排背景檢查。">
              {settingsTab === "監控排程" && (
                <button
                  disabled={!scheduleSettings}
                  onClick={() => void saveSchedule()}
                >
                  儲存變更
                </button>
              )}
            </PageHeader>
            <Tabs
              label="設定分類"
              labels={["監控排程", "通知", "啟動與系統匣", "資料管理"]}
              value={settingsTab}
              onChange={setSettingsTab}
            />
            {settingsTab === "監控排程" ? (
              <>
                {scheduleStatus?.available === false && (
                  <p role="status">
                    隔離測試模式不會連線官方網站；立即檢查已停用。
                  </p>
                )}
                {!scheduleSettings ? (
                  <p role="status">正在載入排程設定…</p>
                ) : (
                  <div className="settings-groups">
                    <section
                      className="settings-group"
                      aria-label="重大訊息檢查"
                    >
                      <header>
                        <h2>重大訊息</h2>
                        <div className="switch-label">
                          <Switch
                            label="啟用重大訊息監控"
                            checked={scheduleSettings.monitoring.enabled}
                            onChange={(enabled) =>
                              setScheduleSettings({
                                ...scheduleSettings,
                                monitoring: {
                                  ...scheduleSettings.monitoring,
                                  enabled,
                                },
                              })
                            }
                          />
                          啟用
                        </div>
                      </header>
                      <div className="schedule-fields">
                        <label>
                          開始時間
                          <input
                            aria-label="重大訊息開始時間"
                            type="time"
                            value={scheduleSettings.monitoring.start}
                            onChange={(event) =>
                              setScheduleSettings({
                                ...scheduleSettings,
                                monitoring: {
                                  ...scheduleSettings.monitoring,
                                  start: event.target.value,
                                },
                              })
                            }
                          />
                        </label>
                        <label>
                          結束時間
                          <input
                            aria-label="重大訊息結束時間"
                            type="time"
                            value={scheduleSettings.monitoring.end}
                            onChange={(event) =>
                              setScheduleSettings({
                                ...scheduleSettings,
                                monitoring: {
                                  ...scheduleSettings.monitoring,
                                  end: event.target.value,
                                },
                              })
                            }
                          />
                        </label>
                        <label>
                          檢查頻率
                          <select
                            aria-label="重大訊息檢查頻率"
                            value={scheduleSettings.monitoring.intervalMinutes}
                            onChange={(event) =>
                              setScheduleSettings({
                                ...scheduleSettings,
                                monitoring: {
                                  ...scheduleSettings.monitoring,
                                  intervalMinutes: Number(
                                    event.target.value,
                                  ) as 15 | 30 | 60 | 120,
                                },
                              })
                            }
                          >
                            {[15, 30, 60, 120].map((n) => (
                              <option key={n} value={n}>
                                {n} 分鐘
                              </option>
                            ))}
                          </select>
                        </label>
                      </div>
                      <p className="muted field-help">
                        檢查間隔從上次檢查起算，不是固定整點或每刻鐘。無新公告時不通知。
                      </p>
                      {renderCheck("material")}
                      {renderJob("material")}
                    </section>
                    <section
                      className="settings-group"
                      aria-label="違約交割檢查"
                    >
                      <header>
                        <h2>違約交割</h2>
                        <div className="switch-label">
                          <Switch
                            label="啟用違約交割監控"
                            checked={scheduleSettings.disclosure.enabled}
                            onChange={(enabled) =>
                              setScheduleSettings({
                                ...scheduleSettings,
                                disclosure: {
                                  ...scheduleSettings.disclosure,
                                  enabled,
                                },
                              })
                            }
                          />
                          啟用
                        </div>
                      </header>
                      <div className="disclosure-fields">
                        <label>
                          每日檢查時間
                          <input
                            type="time"
                            aria-label="違約交割檢查時間"
                            value={scheduleSettings.disclosure.runAt}
                            onChange={(event) =>
                              setScheduleSettings({
                                ...scheduleSettings,
                                disclosure: {
                                  ...scheduleSettings.disclosure,
                                  runAt: event.target.value,
                                },
                              })
                            }
                          />
                        </label>
                        <label className="check-field">
                          <input
                            type="checkbox"
                            checked={
                              scheduleSettings.notifyEmptyDefaultDisclosures
                            }
                            onChange={(event) =>
                              setScheduleSettings({
                                ...scheduleSettings,
                                notifyEmptyDefaultDisclosures:
                                  event.target.checked,
                              })
                            }
                          />
                          本日無違約揭露時也通知
                        </label>
                      </div>
                      {renderCheck("disclosure")}
                      {renderJob("disclosure")}
                    </section>
                    <section className="settings-group">
                      <header>
                        <h2>法說會</h2>
                      </header>
                      {conferenceTaskStatus?.kind === "v2" &&
                        !conferenceTaskStatus.legacyConfigPath &&
                        !legacyConfigReview &&
                        !conferenceTaskMessage && (
                          <p>v2 每日工作已建立；Windows 每日 19:00 執行。</p>
                        )}
                      <p>
                        {calendarLabel} ·{" "}
                        <TextAction onClick={() => navigate("conferences")}>
                          管理連線
                        </TextAction>
                      </p>
                      <details
                        open={
                          conferenceTaskStatus?.kind !== "v2" ||
                          Boolean(legacyConfigReview) ||
                          Boolean(conferenceTaskStatus?.legacyConfigPath) ||
                          Boolean(conferenceTaskMessage)
                        }
                      >
                        <summary>每日工作與舊版管理</summary>{" "}
                        <section
                          className="schedule-settings"
                          aria-labelledby="conference-task-title"
                        >
                          <h3 id="conference-task-title">法說會每日工作</h3>
                          {!conferenceTaskStatus ? (
                            <p role="status">正在確認 Windows 每日工作…</p>
                          ) : conferenceTaskStatus.kind ===
                            "legacy-disabled" ? (
                            <>
                              <p role="status">
                                既有 v1 工作已停用並保留供回退；v2
                                每日工作尚未啟用。完成正式 Calendar
                                驗收後可明確切換。
                              </p>
                              {!conferenceTaskStatus.setupReady && (
                                <p role="status">
                                  切換前請先設定至少一間啟用的關注公司，並完成
                                  Google Calendar 連線。
                                </p>
                              )}
                              <button
                                type="button"
                                disabled={
                                  conferenceTaskBusy ||
                                  !conferenceTaskStatus.setupReady
                                }
                                onClick={() => void switchFromDisabledLegacy()}
                              >
                                切換為 v2 每日排程
                              </button>
                            </>
                          ) : conferenceTaskStatus.kind === "legacy" ? (
                            <>
                              <p>
                                沿用既有 v1 工作；不會另建 v2
                                每日工作。排程與授權維持原狀。
                              </p>
                              {taskHealth && (
                                <p
                                  role={taskHealth.role}
                                  className={
                                    taskHealth.role === "alert"
                                      ? "error-message"
                                      : undefined
                                  }
                                >
                                  {taskHealth.message}
                                </p>
                              )}
                            </>
                          ) : conferenceTaskStatus.kind === "v2" ? (
                            <>
                              <p>
                                v2 每日工作已建立；Windows 每日 19:00 執行。
                              </p>
                              {taskHealth && (
                                <p
                                  role={taskHealth.role}
                                  className={
                                    taskHealth.role === "alert"
                                      ? "error-message"
                                      : undefined
                                  }
                                >
                                  {taskHealth.message}
                                </p>
                              )}
                            </>
                          ) : conferenceTaskStatus.kind ===
                            "upgrade-v2-time" ? (
                            <>
                              <p role="status">
                                偵測到舊版 v2 每日 17:45
                                工作；會先備份並核對唯一寫入者，再遷移至 19:00。
                              </p>
                              <button
                                type="button"
                                disabled={conferenceTaskBusy}
                                onClick={() => void ensureConferenceTask()}
                              >
                                重試遷移至 19:00
                              </button>
                            </>
                          ) : conferenceTaskStatus.kind === "create-v2" ? (
                            <>
                              <p>
                                未偵測到既有 v1 工作。完成關注公司與 Google
                                連線後，v2 會自動建立每日 19:00
                                的法說會工作；若尚未建立，可在此重試。
                              </p>
                              {!conferenceTaskStatus.setupReady && (
                                <p role="status">
                                  請先設定關注公司並完成 Google Calendar 連線。
                                </p>
                              )}
                              <button
                                type="button"
                                disabled={
                                  conferenceTaskBusy ||
                                  !conferenceTaskStatus.setupReady
                                }
                                onClick={() => void ensureConferenceTask()}
                              >
                                建立 v2 每日工作
                              </button>
                            </>
                          ) : (
                            <p role="alert">
                              {conferenceTaskStatus.reason ??
                                "無法確認 Windows 每日工作；為避免重複建立，已停止排程設定。"}
                            </p>
                          )}
                          {(conferenceTaskStatus?.kind === "legacy" ||
                            conferenceTaskStatus?.kind === "legacy-disabled" ||
                            (conferenceTaskStatus?.kind === "v2" &&
                              conferenceTaskStatus.legacyConfigPath)) && (
                            <>
                              <p>
                                修改 v2
                                關注公司前，請先核對並確認同步至舊版股號設定；尚未確認時不會改動舊版設定。
                              </p>
                              <button
                                type="button"
                                disabled={conferenceTaskBusy}
                                onClick={() =>
                                  void reviewLegacyConferenceConfig()
                                }
                              >
                                核對舊版股號
                              </button>
                              {legacyConfigReview && (
                                <div role="group" aria-label="舊版股號核對">
                                  <p>
                                    舊版目前：
                                    {legacyConfigReview.currentCodes.join(
                                      "、",
                                    ) || "無"}
                                  </p>
                                  <p>
                                    v2 啟用：
                                    {legacyConfigReview.activeCodes.join(
                                      "、",
                                    ) || "無"}
                                  </p>
                                  <button
                                    type="button"
                                    disabled={
                                      conferenceTaskBusy ||
                                      legacyConfigReview.activeCodes.length ===
                                        0
                                    }
                                    onClick={() =>
                                      void confirmLegacyConferenceConfig()
                                    }
                                  >
                                    確認套用到舊版工作
                                  </button>
                                  {legacyConfigReview.activeCodes.length ===
                                    0 && (
                                    <p role="status">
                                      至少保留一間啟用公司；舊版工作不可寫入空股號。
                                    </p>
                                  )}
                                </div>
                              )}
                            </>
                          )}
                          {conferenceTaskMessage && (
                            <p
                              role={
                                conferenceTaskMessage.includes("失敗")
                                  ? "alert"
                                  : "status"
                              }
                            >
                              {conferenceTaskMessage}
                            </p>
                          )}
                        </section>
                      </details>
                    </section>
                  </div>
                )}
                {scheduleMessageKind === null && scheduleMessage && (
                  <p role="status">{scheduleMessage}</p>
                )}
              </>
            ) : settingsTab === "通知" ? (
              <div className="settings-groups">
                <section className="settings-group">
                  <header>
                    <h2>通知偏好</h2>
                  </header>
                  <p>新公告與盤面紀錄由背景工作依既有通知規則處理。</p>
                  <p className="muted">無違約揭露通知可在「監控排程」設定。</p>
                </section>
                <section className="settings-group">
                  <header>
                    <h2>測試通知</h2>
                  </header>
                  <div className="button-row">
                    <button
                      disabled={testNotificationBusy}
                      onClick={() => void sendTestNotification()}
                    >
                      發送測試通知
                    </button>
                    <button
                      className="secondary"
                      disabled={testNotificationBusy}
                      onClick={() => void scheduleTestNotification()}
                    >
                      一分鐘後發送測試通知
                    </button>
                  </div>
                  {testNotificationMessage && (
                    <p
                      role={
                        testNotificationMessage.includes("失敗")
                          ? "alert"
                          : "status"
                      }
                    >
                      {testNotificationMessage}
                    </p>
                  )}
                </section>
              </div>
            ) : settingsTab === "啟動與系統匣" ? (
              <div className="settings-groups">
                <section className="settings-group">
                  <header>
                    <h2>Windows 登入啟動</h2>
                  </header>
                  {!loginStartupSettings ? (
                    <p role="status">正在載入登入啟動設定…</p>
                  ) : (
                    <>
                      <div className="preference-row">
                        <div>
                          <h3>登入 Windows 時啟動</h3>
                          <p className="muted">登入後自動開始背景監控。</p>
                        </div>
                        <Switch
                          label="登入 Windows 時啟動"
                          disabled={loginStartupBusy}
                          checked={loginStartupSettings.enabled}
                          onChange={(enabled) =>
                            void saveLoginStartup({
                              ...loginStartupSettings,
                              enabled,
                            })
                          }
                        />
                      </div>
                      <div className="preference-row">
                        <div>
                          <h3>啟動後縮到系統匣</h3>
                          <p className="muted">
                            先在背景執行，需要時再開啟視窗。
                          </p>
                        </div>
                        <Switch
                          label="啟動後縮到系統匣"
                          disabled={loginStartupBusy}
                          checked={loginStartupSettings.startHidden}
                          onChange={(startHidden) =>
                            void saveLoginStartup({
                              ...loginStartupSettings,
                              startHidden,
                            })
                          }
                        />
                      </div>
                      {loginStartupSettings.blockedByWindows && (
                        <p role="alert">Windows 已停用此應用程式的登入啟動。</p>
                      )}
                      <button
                        className="secondary"
                        disabled={loginStartupBusy}
                        onClick={() =>
                          void saveLoginStartup(loginStartupSettings)
                        }
                      >
                        儲存登入啟動設定
                      </button>
                    </>
                  )}
                  {loginStartupMessage && (
                    <p role="status">{loginStartupMessage}</p>
                  )}
                </section>
              </div>
            ) : (
              <div className="settings-groups">
                <section className="settings-group">
                  <header>
                    <h2>本機資料備份</h2>
                  </header>
                  <p>
                    匯出非敏感本機資料，不包含 Google
                    授權資訊。不會覆寫已有檔案。
                  </p>
                  <button
                    disabled={dataExportBusy}
                    onClick={() => void exportLocalData()}
                  >
                    匯出本機資料
                  </button>
                  {dataExportMessage && (
                    <p role={dataExportFailed ? "alert" : "status"}>
                      {dataExportMessage}
                    </p>
                  )}
                </section>
                <section className="settings-group">
                  <header>
                    <h2>盤面資料</h2>
                  </header>
                  <TextAction onClick={() => navigate("market-records")}>
                    管理 CMoney／TWSE 日資料
                  </TextAction>
                </section>
              </div>
            )}
            <footer className="page-foot">
              <span>已儲存的排程會在背景執行。</span>
              <TextAction onClick={() => setShowSources(true)}>
                查看來源狀態
              </TextAction>
            </footer>
          </>
        ) : (
          <>
            <PageHeader title="重大訊息" subtitle="關注公司的最新公告">
              {renderCheck("material")}
            </PageHeader>
            {materialMonitorStatus &&
              materialMonitorStatus.status !== "complete" && (
                <div className="source-warning">
                  <Icon name="warning" />
                  <span>部分來源待確認；已取得的公告仍可閱讀</span>
                  <TextAction onClick={() => setShowSources(true)}>
                    展開詳情
                  </TextAction>
                </div>
              )}
            {renderJob("material")}
            <div className="list-toolbar">
              <label className="search-field">
                <Icon name="search" />
                <input
                  type="search"
                  aria-label="搜尋重大訊息"
                  placeholder="搜尋公司或公告內容"
                  value={materialQuery}
                  onChange={(event) => setMaterialQuery(event.target.value)}
                />
              </label>
              <label className="check-field">
                <input
                  type="checkbox"
                  aria-label="顯示全部公告"
                  checked={showAllMaterial}
                  onChange={(event) => setShowAllMaterial(event.target.checked)}
                />
                顯示全部公告
              </label>
              <label className="check-field">
                <input
                  type="checkbox"
                  aria-label="只顯示未讀重大訊息"
                  checked={unreadOnly}
                  onChange={(event) => setUnreadOnly(event.target.checked)}
                />
                只看未讀
              </label>
            </div>
            {materialActionMessage && (
              <p role="status" aria-label="重大訊息操作結果">
                {materialActionMessage}
              </p>
            )}
            {!materialEvents.length ? (
              <p className="empty-line">
                {showAllMaterial
                  ? "目前沒有符合條件的重大訊息。"
                  : "目前沒有符合條件的關注公司重大訊息；可勾選「顯示全部公告」。"}
              </p>
            ) : (
              <div className="announcement-list material-list" role="list">
                <div className="announcement-head">
                  <span>發布時間</span>
                  <span>公司</span>
                  <span>公告主旨</span>
                </div>
                {materialEvents.map((event) => (
                  <article
                    className={`announcement material-row ${expandedMaterialRowId === event.id ? "expanded" : ""}`}
                    role="listitem"
                    key={event.id}
                  >
                    <button
                      className="announcement-summary material-summary"
                      aria-label={`${event.companyName} ${event.title}`}
                      aria-expanded={expandedMaterialRowId === event.id}
                      onClick={() => void openMaterialEvent(event.id)}
                    >
                      <span className="announcement-time">
                        {!event.readAt && <i className="small-dot" />}
                        {formatTaipeiDateTime(event.publishedAt)}
                      </span>
                      <span>
                        <strong>{event.companyName}</strong>{" "}
                        <small>{event.stockCode}</small>
                      </span>
                      <span className="announcement-subject">
                        {event.title}
                        <Icon name="down" size={18} />
                      </span>
                    </button>
                    {expandedMaterialRowId === event.id &&
                      !selectedMaterialEvent && (
                        <p role="status">載入完整公告…</p>
                      )}
                    {expandedMaterialRowId === event.id &&
                      selectedMaterialEvent && (
                        <aside
                          aria-label="重大訊息完整內容"
                          className="announcement-detail event-detail"
                        >
                          <h3>公告內容</h3>
                          <p className="muted">
                            {selectedMaterialEvent.companyName} ·{" "}
                            {selectedMaterialEvent.title}
                          </p>
                          <div className="event-content">
                            {selectedMaterialEvent.content}
                          </div>
                          <p className="muted">
                            來源：{selectedMaterialEvent.source} · 發布{" "}
                            {formatTaipeiDateTime(
                              selectedMaterialEvent.publishedAt,
                            )}{" "}
                            · 首次發現{" "}
                            {formatTaipeiDateTime(
                              selectedMaterialEvent.discoveredAt,
                            )}
                          </p>
                          {selectedMaterialEvent.original && (
                            <button
                              className="text-action"
                              onClick={() =>
                                void openMaterialEvent(
                                  selectedMaterialEvent.original!.id,
                                  event.id,
                                )
                              }
                            >
                              原公告：{selectedMaterialEvent.original.title}
                            </button>
                          )}
                          {selectedMaterialEvent.relatedRevisions.length >
                            0 && (
                            <p>
                              後續版本：
                              {selectedMaterialEvent.relatedRevisions
                                .map((row) => row.title)
                                .join("、")}
                            </p>
                          )}
                          <footer>
                            <a
                              href={selectedMaterialEvent.sourceUrl}
                              target="_blank"
                              rel="noreferrer"
                            >
                              開啟來源公告
                            </a>
                            <div>
                              {!selectedMaterialEvent.readAt && (
                                <button
                                  className="text-action"
                                  onClick={() =>
                                    void markSelectedMaterialEventRead()
                                  }
                                >
                                  標記為已讀
                                </button>
                              )}
                              <button
                                className="danger-text"
                                disabled={materialDeleteBusy}
                                onClick={() =>
                                  void softDeleteSelectedMaterialEvent()
                                }
                              >
                                刪除此筆本機公告
                              </button>
                              <button
                                className="text-action close-detail"
                                aria-label="關閉重大訊息詳情"
                                onClick={() => {
                                  materialDetailRequest.current += 1;
                                  setExpandedMaterialRowId(null);
                                  setSelectedMaterialEvent(null);
                                }}
                              >
                                關閉
                                <Icon name="down" size={17} />
                              </button>
                            </div>
                          </footer>
                        </aside>
                      )}
                  </article>
                ))}
              </div>
            )}
            <footer className="table-foot">
              <span>顯示 {materialEvents.length} 筆公告</span>
              <TextAction onClick={() => setShowSources(true)}>
                查看來源狀態
              </TextAction>
            </footer>
          </>
        )}
      </main>
      {addingCompany && (
        <Modal title="新增關注公司" onClose={() => setAddingCompany(false)}>
          <form
            aria-label="新增關注公司"
            className="company-form"
            onSubmit={(event) => void addCompany(event)}
          >
            <label>
              市場
              <select
                aria-label="市場"
                value={market}
                onChange={(event) =>
                  setMarket(event.target.value as "TWSE" | "TPEX")
                }
              >
                <option value="TWSE">上市</option>
                <option value="TPEX">上櫃</option>
              </select>
            </label>
            <label>
              股票代號
              <input
                aria-label="股票代號"
                value={stockCode}
                inputMode="numeric"
                autoComplete="off"
                maxLength={6}
                onChange={(event) => setStockCode(event.target.value)}
              />
            </label>
            <p className="muted">公司名稱由官方名錄查詢。</p>
            {error && <p role="alert">{error}</p>}
            <div className="dialog-actions">
              <button
                type="button"
                className="secondary"
                onClick={() => setAddingCompany(false)}
              >
                取消
              </button>
              <button disabled={busy}>新增公司</button>
            </div>
          </form>
        </Modal>
      )}
      {editingCompany && (
        <Modal title="編輯關注公司" onClose={() => setEditingCompany(null)}>
          <form
            className="company-form"
            onSubmit={(event) => {
              event.preventDefault();
              void changeWatchlist(() =>
                window.reporterApi.updateWatchlist({
                  companyId: editingCompany.companyId,
                  category: editingCompany.category,
                  notes: editingCompany.notes,
                }),
              ).then((ok) => {
                if (ok) setEditingCompany(null);
              });
            }}
          >
            <p>
              {editingCompany.stockCode} · {editingCompany.name}
            </p>
            <label>
              分類
              <input
                value={editingCompany.category}
                onChange={(event) =>
                  setEditingCompany({
                    ...editingCompany,
                    category: event.target.value,
                  })
                }
              />
            </label>
            <label>
              備註
              <textarea
                value={editingCompany.notes}
                onChange={(event) =>
                  setEditingCompany({
                    ...editingCompany,
                    notes: event.target.value,
                  })
                }
              />
            </label>
            {error && <p role="alert">{error}</p>}
            <div className="dialog-actions">
              <button
                type="button"
                className="secondary"
                onClick={() => setEditingCompany(null)}
              >
                取消
              </button>
              <button disabled={busy}>儲存</button>
            </div>
          </form>
        </Modal>
      )}
      {removingCompany && (
        <Modal title="移除關注公司" onClose={() => setRemovingCompany(null)}>
          <p>
            確認移除 {removingCompany.name}（{removingCompany.stockCode}
            ）？已保存的公告會保留。
          </p>
          {error && <p role="alert">{error}</p>}
          <div className="dialog-actions">
            <button
              className="secondary"
              onClick={() => setRemovingCompany(null)}
            >
              取消
            </button>
            <button
              disabled={busy}
              onClick={() =>
                void changeWatchlist(() =>
                  window.reporterApi.removeWatchlist({
                    companyId: removingCompany.companyId,
                  }),
                ).then((ok) => {
                  if (ok) setRemovingCompany(null);
                })
              }
            >
              確認移除
            </button>
          </div>
        </Modal>
      )}
      {importingWatchlist && (
        <Modal
          title="匯入舊版股票清單"
          onClose={() => setImportingWatchlist(false)}
          wide
        >
          <label>
            舊版設定 JSON
            <textarea
              aria-label="舊版設定 JSON"
              value={configText}
              onChange={(event) => {
                setConfigText(event.target.value);
                setImportReport(null);
              }}
            />
          </label>
          <div className="button-row">
            <button className="secondary" onClick={() => void previewImport()}>
              預覽匯入
            </button>
            <button
              disabled={busy || !importReport}
              onClick={() => void applyImport()}
            >
              確認匯入
            </button>
          </div>
          {error && <p role="alert">{error}</p>}
          {importReport && (
            <>
              <p role="status">
                新增 {importReport.summary.added}、略過{" "}
                {importReport.summary.skipped}、失敗{" "}
                {importReport.summary.failed}
              </p>
              <ul>
                {importReport.items.map((item, index) => (
                  <li key={index}>
                    {item.stockCode}：{item.status} · {item.reason}
                  </li>
                ))}
              </ul>
            </>
          )}
        </Modal>
      )}
      {showSources && (
        <Modal title="資料來源狀態" onClose={() => setShowSources(false)} wide>
          <h3>重大訊息</h3>
          <MaterialSourceReport monitor={materialMonitorStatus} />
          <h3>違約交割</h3>
          {!disclosureMonitorStatus ? (
            <p>尚未檢查</p>
          ) : (
            Object.entries(disclosureMonitorStatus.markets).map(
              ([market, source]) => (
                <p key={market}>
                  {market}：{disclosureStatusLabel(source.status)} ·{" "}
                  {source.dataDate ?? "尚無資料日期"}
                  {source.errorMessage && ` · ${source.errorMessage}`}
                </p>
              ),
            )
          )}
          <div className="dialog-actions">
            <button onClick={() => setShowSources(false)}>關閉</button>
          </div>
        </Modal>
      )}
    </div>
    </GuideProvider>
  );
}
