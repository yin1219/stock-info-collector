import { useContext, useEffect, useRef, useState, type ReactNode } from "react";
import { driver } from "driver.js";
import "driver.js/dist/driver.css";
import "./guide.css";
import { SetupChecklist, type SetupApi, type SetupTab } from "./SetupChecklist";
import { GuideContext } from "./guide-context";
import { Modal, Icon } from "./ui";
import {
  pageGuides,
  quickTour,
  readGuidePreference,
  saveGuidePreference,
  type GuidePage,
  type GuidePreference,
  type TourStep,
} from "./onboarding";

interface TourSession {
  steps: TourStep[];
  index: number;
  quick: boolean;
}
export function GuideProvider({
  page,
  onNavigate,
  children,
  api,
}: {
  page: GuidePage;
  onNavigate: (page: GuidePage, tab?: SetupTab) => void;
  api?: SetupApi;
  children: ReactNode;
}) {
  const [welcome, setWelcome] = useState(() => readGuidePreference() === null);
  const [message, setMessage] = useState<string | null>(null);
  const [help, setHelp] = useState<GuidePage | "index" | "setup" | null>(null);
  const [session, setSession] = useState<TourSession | null>(null);
  const focusPending = useRef(false);
  const navigate = useRef(onNavigate);
  navigate.current = onNavigate;

  function remember(status: GuidePreference) {
    setWelcome(false);
    setMessage(
      saveGuidePreference(undefined, status)
        ? null
        : "導覽偏好無法保存，下次啟動可能再次顯示歡迎卡。",
    );
  }
  function finish(complete: boolean) {
    if (session?.quick) remember(complete ? "complete" : "skipped");
    focusPending.current = true;
    setSession(null);
  }
  function start(steps: TourStep[], quick: boolean) {
    setHelp(null);
    setWelcome(false);
    navigate.current(steps[0].page);
    setSession({ steps, quick, index: 0 });
  }
  useEffect(() => {
    if (!session) {
      if (focusPending.current) {
        focusPending.current = false;
        document
          .querySelector<HTMLButtonElement>('[data-guide="page-help"]')
          ?.focus();
      }
      return;
    }
    const step = session.steps[session.index];
    if (page !== step.page) {
      finish(false);
      return;
    }
    function move(index: number) {
      if (index === session!.steps.length) {
        finish(true);
        return;
      }
      navigate.current(session!.steps[index].page);
      setSession({ ...session!, index });
    }
    const reducedMotion =
      window.matchMedia?.("(prefers-reduced-motion: reduce)").matches ?? false;
    const tour = driver({
      steps: session.steps.map((item) => ({
        element:
          item.page === page
            ? (document.querySelector(item.target) ?? undefined)
            : undefined,
        popover: {
          title: item.title,
          description: item.description,
          side: "bottom" as const,
          align: "start" as const,
        },
      })),
      animate: !reducedMotion,
      smoothScroll: false,
      allowClose: true,
      allowKeyboardControl: true,
      disableActiveInteraction: true,
      overlayClickBehavior: "none",
      overlayOpacity: 0.42,
      stagePadding: 6,
      stageRadius: 6,
      popoverClass: "reporter-tour",
      showButtons: ["previous", "next", "close"],
      disableButtons: session.index === 0 ? ["previous"] : [],
      showProgress: true,
      progressText: `第 ${session.index + 1} / ${session.steps.length} 步`,
      prevBtnText: "上一步",
      nextBtnText: "下一步",
      doneBtnText:
        session.index === session.steps.length - 1 ? "完成導覽" : "下一步",
      closeBtnLabel: "結束導覽",
      onPrevClick: () => move(Math.max(0, session.index - 1)),
      onNextClick: () => move(session.index + 1),
      onDestroyStarted: () => finish(false),
      onPopoverRender: ({ wrapper, nextButton, previousButton }) => {
        wrapper.setAttribute("role", "dialog");
        wrapper.setAttribute("aria-modal", "true");
        wrapper.removeAttribute("aria-labelledby");
        wrapper.setAttribute("aria-label", "功能導覽");
        previousButton.disabled = session.index === 0;
        wrapper.addEventListener("keydown", (event) => {
          if (event.key !== "Tab") return;
          event.preventDefault();
          event.stopPropagation();
          const buttons = Array.from(
            wrapper.querySelectorAll<HTMLButtonElement>(
              "button:not(:disabled)",
            ),
          );
          const index = buttons.indexOf(
            document.activeElement as HTMLButtonElement,
          );
          const offset = event.shiftKey ? -1 : 1;
          buttons[(index + offset + buttons.length) % buttons.length]?.focus();
        });
        nextButton.focus();
      },
    });
    tour.drive(session.index);
    return () => tour.destroy();
  }, [session, page]);

  const guide =
    help && help !== "index" && help !== "setup" ? pageGuides[help] : null;
  return (
    <GuideContext.Provider
      value={{
        welcome,
        message,
        openSetup: api ? () => setHelp("setup") : undefined,
        openPage: () => setHelp(page),
        openIndex: () => setHelp("index"),
        startQuick: () => start(quickTour, true),
        skip: () => remember("skipped"),
      }}
    >
      {children}
      {help === "setup" && api && (
        <SetupChecklist
          api={api}
          onNavigate={(target, tab) => navigate.current(target, tab)}
          onClose={() => {
            setHelp(null);
            if (welcome) remember("skipped");
          }}
        />
      )}
      {help === "index" && (
        <Modal title="使用指南" onClose={() => setHelp(null)}>
          <p className="guide-intro">
            從關注公司開始，了解公告、行事曆與盤面資料的工作流程。
          </p>
          <div className="button-row guide-index-actions">
            {api && (
              <button onClick={() => setHelp("setup")}>首次使用設定</button>
            )}
            <button
              className="secondary"
              onClick={() => start(quickTour, true)}
            >
              重看快速導覽
            </button>
          </div>
          <h3 className="guide-topics-title">各頁功能說明</h3>
          <div className="guide-topics">
            {(
              Object.entries(pageGuides) as [
                GuidePage,
                (typeof pageGuides)[GuidePage],
              ][]
            ).map(([key, item]) => (
              <button
                className="guide-topic"
                key={key}
                aria-label={item.title}
                onClick={() => setHelp(key)}
              >
                <strong>{item.title}</strong>
                <span>{item.summary}</span>
                <Icon name="chevron" size={18} />
              </button>
            ))}
          </div>
        </Modal>
      )}
      {guide && help && help !== "index" && help !== "setup" && (
        <Modal title={`${guide.title}使用說明`} onClose={() => setHelp(null)}>
          <p className="guide-intro">{guide.summary}</p>
          <ol className="guide-explanations">
            {guide.steps.map((step) => (
              <li key={step.title}>
                <h3>{step.title}</h3>
                <p>{step.description}</p>
              </li>
            ))}
          </ol>
          <div className="dialog-actions guide-actions">
            <button className="secondary" onClick={() => setHelp("index")}>
              所有功能說明
            </button>
            <button
              onClick={() =>
                start(
                  guide.steps.map((step) => ({ ...step, page: help })),
                  false,
                )
              }
            >
              帶我操作
            </button>
          </div>
        </Modal>
      )}
    </GuideContext.Provider>
  );
}
export function GuideWelcome() {
  const guide = useContext(GuideContext);
  if (!guide) return null;
  return (
    <>
      {guide.welcome && (
        <section className="guide-welcome" aria-label="歡迎使用小幫手">
          <div>
            <h2>歡迎使用小幫手</h2>
            <p>先設定關注公司與背景檢查，也可用 5 步導覽認識功能。</p>
          </div>
          <div className="button-row">
            {guide.openSetup && (
              <button onClick={guide.openSetup}>開始設定</button>
            )}
            <button className="secondary" onClick={guide.startQuick}>
              開始導覽
            </button>
            <button className="text-action" onClick={guide.skip}>
              稍後
            </button>
          </div>
        </section>
      )}
      {guide.message && (
        <p role="status" className="source-warning">
          {guide.message}
        </p>
      )}
    </>
  );
}
export function GuideLauncher() {
  const guide = useContext(GuideContext);
  return (
    guide && (
      <button className="text-action guide-launcher" onClick={guide.openIndex}>
        <Icon name="info" size={20} />
        使用指南
      </button>
    )
  );
}
