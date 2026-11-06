import { useEffect, useRef, useState } from "react";
import { Icon } from "./components";
import {
  pages,
  companies as seedCompanies,
  announcements as seedAnnouncements,
  initialPreferences,
  type Page,
} from "./data";
import { OverviewPage, WatchlistPage, AnnouncementsPage } from "./WorkPages";
import { DisclosuresPage, ConferencesPage } from "./SourcePages";
import { RecordsPage, SettingsPage } from "./SettingsAndRecords";
const icons = [
  "home",
  "building",
  "megaphone",
  "warning",
  "conference",
  "bars",
  "gear",
];
export function DemoApp() {
  const [page, setPage] = useState<Page>("總覽"),
    [navOpen, setNavOpen] = useState(false),
    [companies, setCompanies] = useState(() =>
      seedCompanies.map((c) => ({ ...c })),
    ),
    [announcements, setAnnouncements] = useState(() =>
      seedAnnouncements.map((a) => ({ ...a })),
    ),
    [preferences, setPreferences] = useState({ ...initialPreferences }),
    [feedback, setFeedback] = useState(""),
    [connected, setConnected] = useState(true),
    [resetKey, setResetKey] = useState(0),
    [timerDue, setTimerDue] = useState<string | null>(null);
  const noticeTimer = useRef<ReturnType<typeof setTimeout> | null>(null);
  useEffect(
    () => () => {
      if (noticeTimer.current) clearTimeout(noticeTimer.current);
    },
    [],
  );
  function cancelNotice() {
    if (noticeTimer.current) clearTimeout(noticeTimer.current);
    noticeTimer.current = null;
    setTimerDue(null);
  }
  function scheduleNotice() {
    cancelNotice();
    setTimerDue(new Date(Date.now() + 60000).toLocaleTimeString("zh-TW"));
    noticeTimer.current = setTimeout(() => {
      setFeedback("一分鐘測試：示意通知已到期；未發送 Windows 通知");
      noticeTimer.current = null;
      setTimerDue(null);
    }, 60000);
    setFeedback("已排定一分鐘後的示意通知");
  }
  function navigate(p: Page) {
    setPage(p);
    setNavOpen(false);
    setFeedback("");
  }
  function reset() {
    cancelNotice();
    setCompanies(seedCompanies.map((c) => ({ ...c })));
    setAnnouncements(seedAnnouncements.map((a) => ({ ...a })));
    setPreferences({ ...initialPreferences });
    setConnected(true);
    setResetKey((k) => k + 1);
    setFeedback("已重置示意資料");
  }
  return (
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
        <nav aria-label="主要導覽">
          {pages.map((p, i) => (
            <button
              key={p}
              aria-current={page === p ? "page" : undefined}
              className={page === p ? "nav-item selected" : "nav-item"}
              onClick={() => navigate(p)}
            >
              <Icon name={icons[i]} size={27} />
              <span>{p}</span>
            </button>
          ))}
        </nav>
        <footer>
          <p>
            <i className="status-dot" />
            背景監控中
          </p>
          <span>v2 · 版本 1.0.23</span>
        </footer>
      </aside>
      <main className="reporter-main" id="demo-content">
        <div className="demo-label">設計預覽 · 示意資料</div>
        <div key={resetKey}>
          {page === "總覽" && (
            <OverviewPage
              companies={companies}
              announcements={announcements.filter((a) =>
                companies.some((c) => c.active && c.code === a.code),
              )}
              navigate={navigate}
              notify={setFeedback}
            />
          )}
          {page === "關注公司" && (
            <WatchlistPage
              companies={companies}
              onChange={setCompanies}
              notify={setFeedback}
            />
          )}
          {page === "重大訊息" && (
            <AnnouncementsPage
              items={announcements}
              companies={companies}
              onChange={setAnnouncements}
              notify={setFeedback}
            />
          )}
          {page === "違約交割" && <DisclosuresPage notify={setFeedback} />}
          {page === "法說會" && (
            <ConferencesPage
              connected={connected}
              onConnect={setConnected}
              notify={setFeedback}
              settings={() => navigate("設定")}
            />
          )}
          {page === "盤面紀錄" && <RecordsPage notify={setFeedback} />}
          {page === "設定" && (
            <SettingsPage
              timerDue={timerDue}
              schedule={scheduleNotice}
              cancelSchedule={() => {
                cancelNotice();
                setFeedback("已取消示意通知");
              }}
              preferences={preferences}
              onChange={setPreferences}
              connected={connected}
              onConnect={setConnected}
              notify={setFeedback}
              reset={reset}
            />
          )}
        </div>
        <p
          className={feedback ? "feedback visible" : "feedback"}
          role="status"
          aria-live="polite"
        >
          {feedback}
        </p>
      </main>
    </div>
  );
}
