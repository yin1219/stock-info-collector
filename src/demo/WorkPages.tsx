import { useState } from "react";
import {
  DateControl,
  FileDemo,
  Icon,
  Modal,
  PageHeader,
  SourceDialog,
  Switch,
  TextAction,
} from "./components";
import type { Company, Announcement, Page } from "./data";
type Notice = (text: string) => void;
export function OverviewPage({
  companies,
  announcements,
  navigate,
  notify,
}: {
  companies: Company[];
  announcements: Announcement[];
  navigate: (p: Page) => void;
  notify: Notice;
}) {
  const [date, setDate] = useState("2026-11-05");
  return (
    <>
      <PageHeader title="總覽" subtitle="今日監控與採訪行程">
        <DateControl value={date} onChange={setDate} />
        <button onClick={() => notify("示意檢查完成，沒有連線官方來源")}>
          立即檢查
        </button>
      </PageHeader>
      <p className="running-line">
        <i className="status-dot" />
        背景監控中 · 關閉視窗後仍會繼續檢查
      </p>
      <div className="overview-summary">
        <div>
          關注公司 <strong>{companies.length}</strong> 家
        </div>
        <div>
          未讀重大訊息{" "}
          <strong>{announcements.filter((a) => a.unread).length}</strong> 筆
        </div>
        <div>
          今日違約揭露 <span>待確認</span>
        </div>
      </div>
      <div className="overview-columns">
        <section>
          <div className="section-title">
            <h2>最新重大訊息</h2>
            <TextAction onClick={() => navigate("重大訊息")}>
              查看全部
            </TextAction>
          </div>
          <div className="recent-list">
            {announcements.slice(0, 4).map((a) => (
              <button key={a.id} onClick={() => navigate("重大訊息")}>
                <time>{a.time.slice(6)}</time>
                <span className="recent-company">
                  {a.unread && <i className="small-dot" />}
                  <strong>{a.company}</strong> {a.code}
                </span>
                <span>{a.title}</span>
              </button>
            ))}
          </div>
        </section>
        <aside className="overview-aside">
          <h2>下一場法說會</h2>
          <div className="next-conference">
            <strong>11/12</strong>
            <div>
              14:00–16:00
              <br />
              台積電 2330
              <br />
              <span className="muted">線上法說會</span>
            </div>
          </div>
          <div className="next-calendar">
            <span>
              <Icon name="checkCircle" size={18} />
              已加入 Google Calendar
            </span>
            <button onClick={() => navigate("法說會")}>查看法說會</button>
          </div>
          <h2 className="next-check-heading">接下來的檢查</h2>
          <dl className="next-checks">
            <div>
              <dt>重大訊息</dt>
              <dd>15:15</dd>
            </div>
            <div>
              <dt>違約交割</dt>
              <dd>18:30</dd>
            </div>
            <div>
              <dt>法說會</dt>
              <dd>19:00</dd>
            </div>
          </dl>
        </aside>
      </div>
      <footer className="page-foot">
        <span>
          <Icon name="warning" size={17} />
          MOPS 部分來源待確認
        </span>
        <TextAction onClick={() => navigate("重大訊息")}>
          查看來源狀態
        </TextAction>
      </footer>
    </>
  );
}
export function WatchlistPage({
  companies,
  onChange,
  notify,
}: {
  companies: Company[];
  onChange: (v: Company[]) => void;
  notify: Notice;
}) {
  const [query, setQuery] = useState(""),
    [filter, setFilter] = useState("全部"),
    [all, setAll] = useState(false),
    [editing, setEditing] = useState<Company | null>(null),
    [isNew, setIsNew] = useState(false),
    [removing, setRemoving] = useState<Company | null>(null),
    [menu, setMenu] = useState(""),
    [error, setError] = useState("");
  const filtered = companies.filter(
    (c) =>
      (filter === "全部" || (filter === "監控中" ? c.active : !c.active)) &&
      (c.code.includes(query) || c.name.includes(query)),
  );
  const shown = all ? filtered : filtered.slice(0, 4);
  function save() {
    if (!editing) return;
    if (!/^\d{4,6}$/.test(editing.code) || !editing.name.trim()) {
      setError("請填寫4至6位股票代號與公司名稱");
      return;
    }
    if (isNew && companies.some((c) => c.code === editing.code)) {
      setError("這個股票代號已在關注清單中");
      return;
    }
    onChange(
      isNew
        ? [...companies, editing]
        : companies.map((c) => (c.code === editing.code ? editing : c)),
    );
    setAll(true);
    setQuery("");
    setFilter("全部");
    setEditing(null);
    notify(isNew ? "已新增示意關注公司" : "已更新示意公司資料");
  }
  return (
    <>
      <PageHeader title="關注公司" subtitle="管理需要追蹤公告與法說會的公司">
        <button
          onClick={() => {
            setEditing({
              code: "",
              name: "",
              market: "上市",
              category: "",
              active: true,
              notes: "",
            });
            setIsNew(true);
            setError("");
          }}
        >
          新增關注公司
        </button>
        <FileDemo
          label="匯入舊版股票清單"
          onSelect={(name) => notify(`已選擇 ${name}，Demo 不會修改正式清單`)}
        />
      </PageHeader>
      <div className="list-toolbar">
        <label className="search-field">
          <Icon name="search" />
          <input
            aria-label="搜尋股號或公司名稱"
            placeholder="搜尋股號或公司名稱"
            value={query}
            onChange={(e) => {
              setQuery(e.target.value);
              setAll(true);
            }}
          />
        </label>
        <div className="filter-segments">
          {["全部", "監控中", "已停用"].map((f) => (
            <button
              key={f}
              aria-pressed={filter === f}
              onClick={() => {
                setFilter(f);
                setAll(true);
              }}
            >
              {f}{" "}
              {
                companies.filter(
                  (c) =>
                    f === "全部" || (f === "監控中" ? c.active : !c.active),
                ).length
              }
            </button>
          ))}
        </div>
      </div>
      <div className="table-scroll">
        <table className="watch-table">
          <thead>
            <tr>
              {["股號", "公司名稱", "市場", "分類", "監控", "操作"].map((v) => (
                <th key={v}>{v}</th>
              ))}
            </tr>
          </thead>
          <tbody>
            {shown.map((c) => (
              <tr
                key={c.code}
                className={menu === c.code ? "highlight" : undefined}
              >
                <td className="strong number">{c.code}</td>
                <td className="strong">{c.name}</td>
                <td>{c.market}</td>
                <td>{c.category || "—"}</td>
                <td>
                  <Switch
                    label={`監控${c.name}`}
                    checked={c.active}
                    onChange={(active) =>
                      onChange(
                        companies.map((row) =>
                          row.code === c.code ? { ...row, active } : row,
                        ),
                      )
                    }
                  />
                </td>
                <td>
                  <div className="row-actions">
                    <button
                      className="icon-button"
                      aria-label={`編輯${c.name}`}
                      onClick={() => {
                        setEditing({ ...c });
                        setIsNew(false);
                        setError("");
                      }}
                    >
                      <Icon name="edit" />
                    </button>
                    {menu === c.code ? (
                      <button
                        className="text-action"
                        aria-label={`移除${c.name}`}
                        onClick={() => setRemoving(c)}
                      >
                        移除
                      </button>
                    ) : (
                      <button
                        className="icon-button"
                        aria-label={`更多${c.name}操作`}
                        onClick={() => setMenu(c.code)}
                      >
                        <Icon name="more" />
                      </button>
                    )}
                  </div>
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
      {!shown.length && <p className="empty-line">目前沒有符合條件的公司。</p>}
      <footer className="table-foot">
        <span>停用保留清單；移除不會刪除已保存的公告。</span>
        <span>
          顯示 {shown.length} 家公司
          {!all && filtered.length > 4 && (
            <TextAction onClick={() => setAll(true)}>顯示全部</TextAction>
          )}
        </span>
      </footer>
      {editing && (
        <Modal
          title={isNew ? "新增關注公司" : "編輯關注公司"}
          onClose={() => setEditing(null)}
        >
          <form
            onSubmit={(e) => {
              e.preventDefault();
              save();
            }}
            className="company-form"
          >
            <label>
              股票代號
              <input
                value={editing.code}
                disabled={!isNew}
                inputMode="numeric"
                onChange={(e) =>
                  setEditing({ ...editing, code: e.target.value })
                }
              />
            </label>
            <label>
              公司名稱
              <input
                value={editing.name}
                onChange={(e) =>
                  setEditing({ ...editing, name: e.target.value })
                }
              />
            </label>
            <div className="two-fields">
              <label>
                市場
                <select
                  value={editing.market}
                  onChange={(e) =>
                    setEditing({ ...editing, market: e.target.value })
                  }
                >
                  <option>上市</option>
                  <option>上櫃</option>
                </select>
              </label>
              <label>
                分類
                <input
                  value={editing.category}
                  onChange={(e) =>
                    setEditing({ ...editing, category: e.target.value })
                  }
                />
              </label>
            </div>
            <label>
              備註
              <textarea
                value={editing.notes}
                onChange={(e) =>
                  setEditing({ ...editing, notes: e.target.value })
                }
              />
            </label>
            {error && (
              <p role="alert" className="error">
                {error}
              </p>
            )}
            <div className="dialog-actions">
              <button
                type="button"
                className="secondary"
                onClick={() => setEditing(null)}
              >
                取消
              </button>
              <button type="submit">{isNew ? "新增" : "儲存"}</button>
            </div>
          </form>
        </Modal>
      )}
      {removing && (
        <Modal title="移除關注公司" onClose={() => setRemoving(null)}>
          <p>
            確認移除 {removing.name}（{removing.code}）？已保存的公告會保留。
          </p>
          <div className="dialog-actions">
            <button className="secondary" onClick={() => setRemoving(null)}>
              取消
            </button>
            <button
              onClick={() => {
                onChange(companies.filter((c) => c.code !== removing.code));
                setRemoving(null);
                notify("已從示意清單移除");
              }}
            >
              確認移除
            </button>
          </div>
        </Modal>
      )}
    </>
  );
}
export function AnnouncementsPage({
  items,
  companies,
  onChange,
  notify,
}: {
  items: Announcement[];
  companies: Company[];
  onChange: (v: Announcement[]) => void;
  notify: Notice;
}) {
  const [expanded, setExpanded] = useState<number | null>(null),
    [watched, setWatched] = useState(true),
    [unread, setUnread] = useState(false),
    [query, setQuery] = useState(""),
    [all, setAll] = useState(false),
    [details, setDetails] = useState(false),
    [source, setSource] = useState(false),
    [deleting, setDeleting] = useState<number | null>(null);
  const rows = items.filter(
    (a) =>
      (!watched || companies.some((c) => c.active && c.code === a.code)) &&
      (!unread || a.unread || expanded === a.id) &&
      `${a.code}${a.company}${a.title}${a.body}`.includes(query),
  );
  const visible = all ? rows : rows.slice(0, 4);
  function expand(a: Announcement) {
    setExpanded(expanded === a.id ? null : a.id);
    onChange(
      items.map((row) => (row.id === a.id ? { ...row, unread: false } : row)),
    );
  }
  return (
    <>
      <PageHeader title="重大訊息" subtitle="關注公司的最新公告">
        <button onClick={() => notify("示意檢查完成；部分來源仍待確認")}>
          立即檢查
        </button>
      </PageHeader>
      <div className="source-warning">
        <Icon name="warning" />
        <span>部分來源待確認；已取得的公告仍可閱讀</span>
        <button className="text-action" onClick={() => setDetails(!details)}>
          展開詳情
          <Icon name="down" size={18} />
        </button>
      </div>
      {details && (
        <div className="source-detail">
          MOPS
          已取得部分公告；另一來源尚待確認。此為示意狀態，不代表官方目前狀態。
        </div>
      )}
      <div className="list-toolbar">
        <label className="search-field">
          <Icon name="search" />
          <input
            aria-label="搜尋公司或公告內容"
            placeholder="搜尋公司或公告內容"
            value={query}
            onChange={(e) => setQuery(e.target.value)}
          />
        </label>
        <label className="check-field">
          <input
            type="checkbox"
            checked={watched}
            onChange={(e) => setWatched(e.target.checked)}
          />
          僅顯示關注公司
        </label>
        <label className="check-field">
          <input
            type="checkbox"
            checked={unread}
            onChange={(e) => setUnread(e.target.checked)}
          />
          只看未讀
        </label>
      </div>
      <div className="announcement-list">
        <div className="announcement-head">
          <span>發布時間</span>
          <span>公司</span>
          <span>公告主旨</span>
        </div>
        {visible.map((a) => (
          <article
            key={a.id}
            className={
              expanded === a.id ? "announcement expanded" : "announcement"
            }
          >
            <button
              className="announcement-summary"
              aria-label={a.title}
              aria-expanded={expanded === a.id}
              onClick={() => expand(a)}
            >
              <span className="announcement-time">
                {a.unread && <i className="small-dot" />}
                {a.time}
              </span>
              <span>
                <strong>{a.company}</strong> <small>{a.code}</small>
              </span>
              <span className="announcement-subject">
                {a.title}
                <Icon name="down" size={18} />
              </span>
            </button>
            {expanded === a.id && (
              <div className="announcement-detail">
                <h3>公告內容</h3>
                <p>{a.body}</p>
                <p className="muted">來源：公開資訊觀測站</p>
                <footer>
                  <TextAction onClick={() => setSource(true)}>
                    開啟官方來源
                  </TextAction>
                  <div>
                    <button
                      className="danger-text"
                      onClick={() => setDeleting(a.id)}
                    >
                      刪除此筆本機公告
                    </button>
                    <button
                      className="text-action close-detail"
                      aria-label="關閉公告內容"
                      onClick={() => setExpanded(null)}
                    >
                      關閉
                      <Icon name="down" size={17} />
                    </button>
                  </div>
                </footer>
              </div>
            )}
          </article>
        ))}
      </div>
      {!visible.length && <p className="empty-line">沒有符合條件的公告。</p>}
      <footer className="table-foot">
        <span>顯示 {visible.length} 筆公告</span>
        {!all && rows.length > 4 && (
          <TextAction onClick={() => setAll(true)}>查看更多</TextAction>
        )}
      </footer>
      {source && <SourceDialog onClose={() => setSource(false)} />}
      {deleting !== null && (
        <Modal title="刪除此筆本機公告" onClose={() => setDeleting(null)}>
          <p>刪除後會從示意清單移除。此操作不會變更正式公告。</p>
          <div className="dialog-actions">
            <button className="secondary" onClick={() => setDeleting(null)}>
              取消
            </button>
            <button
              onClick={() => {
                onChange(items.filter((a) => a.id !== deleting));
                setDeleting(null);
                setExpanded(null);
                notify("已刪除示意公告");
              }}
            >
              確認刪除
            </button>
          </div>
        </Modal>
      )}
    </>
  );
}
