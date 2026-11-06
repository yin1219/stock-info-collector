import { GuideContext } from "./guide-context";
import { useContext, useEffect, useRef, type ReactNode } from "react";
export function Icon({ name, size = 22 }: { name: string; size?: number }) {
  const paths: Record<string, ReactNode> = {
    info: <><circle cx="12" cy="12" r="9" /><path d="M12 11v6m0-10v1" /></>,
    home: (
      <>
        <path d="m3 10 9-8 9 8v11h-7v-7h-4v7H3Z" />
      </>
    ),
    building: (
      <>
        <path d="M5 22V3h14v19M2 22h20" />
        <path d="M9 7h1m4 0h1M9 11h1m4 0h1M9 15h1m4 0h1M10 22v-4h4v4" />
      </>
    ),
    megaphone: (
      <>
        <path d="M3 9h5l12-5v16L8 15H3ZM8 15l2 6H6l-2-6M22 9v6" />
      </>
    ),
    warning: (
      <>
        <path d="M10 3 2 19a1 1 0 0 0 1 2h18a1 1 0 0 0 1-2L14 3a2 2 0 0 0-4 0Z" />
        <path d="M12 8v5m0 4v.01" />
      </>
    ),
    calendar: (
      <>
        <rect x="3" y="5" width="18" height="16" rx="2" />
        <path d="M7 2v6m10-6v6M3 11h18" />
      </>
    ),
    conference: (
      <>
        <path d="M9 3h12v12h-5M9 7h7m-2 14h7" />
        <circle cx="7" cy="10" r="3" />
        <path d="M2 22v-3a5 5 0 0 1 10 0v3ZM16 15v6" />
      </>
    ),
    bars: (
      <>
        <path d="M5 21V14m7 7V7m7 14V3" strokeWidth="3.5" />
      </>
    ),
    gear: (
      <>
        <path d="m9 3 1-2h4l1 2 3 2 2 1 2 4-2 2v3l1 2-3 4-3-1-3 1-2 1-4-2v-3l-2-2-2-1 1-5 3-1Z" />
        <circle cx="12" cy="12" r="3" />
      </>
    ),
    search: (
      <>
        <circle cx="10" cy="10" r="7" />
        <path d="m15 15 6 6" />
      </>
    ),
    chevron: <path d="m8 4 8 8-8 8" />,
    down: <path d="m5 8 7 7 7-7" />,
    check: <path d="m5 12 4 4L19 6" />,
    checkCircle: (
      <>
        <circle cx="12" cy="12" r="10" />
        <path d="m6 12 4 4 8-9" />
      </>
    ),
    edit: (
      <>
        <path d="m4 16-1 5 5-1L21 7l-4-4ZM14 6l4 4" />
      </>
    ),
    more: (
      <>
        <circle cx="4" cy="12" r="1" />
        <circle cx="12" cy="12" r="1" />
        <circle cx="20" cy="12" r="1" />
      </>
    ),
    clock: (
      <>
        <circle cx="12" cy="12" r="10" />
        <path d="M12 5v7l5 2" />
      </>
    ),
    menu: <path d="M3 5h18M3 12h18M3 19h18" />,
    close: <path d="m5 5 14 14M19 5 5 19" />,
  };
  return (
    <svg
      width={size}
      height={size}
      viewBox="0 0 24 24"
      fill="none"
      stroke="currentColor"
      strokeWidth="1.7"
      strokeLinecap="round"
      strokeLinejoin="round"
      aria-hidden="true"
    >
      {paths[name] ?? paths.bars}
    </svg>
  );
}
export function PageHeader({
  title,
  subtitle,
  children,
}: {
  title: string;
  subtitle: string;
  children?: ReactNode;
}) {
  const guide = useContext(GuideContext);
  return (
    <header className="page-header" data-guide="page-header">
      <div>
        <div className="page-title">
          <h1>{title}</h1>
          {guide && (
            <button
              className="text-action page-help"
              data-guide="page-help"
              onClick={guide.openPage}
            >
              <Icon name="info" size={18} />使用說明
            </button>
          )}
        </div>
        <p>{subtitle}</p>
      </div>
      <div className="header-actions" data-guide="page-actions">
        {children}
      </div>
    </header>
  );
}
export function TextAction({
  children,
  onClick,
  icon = true,
}: {
  children: ReactNode;
  onClick: () => void;
  icon?: boolean;
}) {
  return (
    <button className="text-action" onClick={onClick}>
      {children}
      {icon && <Icon name="chevron" size={17} />}
    </button>
  );
}
export function Switch({
  label,
  checked,
  onChange,
  disabled = false,
}: {
  label: string;
  checked: boolean;
  onChange: (value: boolean) => void;
  disabled?: boolean;
}) {
  return (
    <button
      type="button"
      className="switch"
      role="switch"
      aria-label={label}
      aria-checked={checked}
      disabled={disabled}
      onClick={() => onChange(!checked)}
    >
      <span />
    </button>
  );
}
export function Tabs({
  labels,
  value,
  onChange,
  small = false,
  label,
}: {
  labels: string[];
  value: string;
  onChange: (value: string) => void;
  small?: boolean;
  label?: string;
}) {
  return (
    <div
      className={small ? "tabs small-tabs" : "tabs"}
      role="tablist"
      aria-label={label}
    >
      {labels.map((label) => (
        <button
          key={label}
          role="tab"
          aria-selected={value === label}
          onClick={() => onChange(label)}
        >
          {label}
        </button>
      ))}
    </div>
  );
}
export function Modal({
  title,
  onClose,
  children,
  wide = false,
}: {
  title: string;
  onClose: () => void;
  children: ReactNode;
  wide?: boolean;
}) {
  const ref = useRef<HTMLDivElement>(null);
  const close = useRef(onClose);
  close.current = onClose;
  useEffect(() => {
    const previous = document.activeElement as HTMLElement;
    const el = ref.current!;
    el.querySelector<HTMLElement>("input,select,button")?.focus();
    function key(e: KeyboardEvent) {
      if (e.key === "Escape") close.current();
      if (e.key === "Tab") {
        const items = Array.from(
          el.querySelectorAll<HTMLElement>(
            "button:not(:disabled),input:not(:disabled),select,textarea,a[href]",
          ),
        );
        const first = items[0],
          last = items.at(-1);
        if (e.shiftKey && document.activeElement === first) {
          e.preventDefault();
          last?.focus();
        } else if (!e.shiftKey && document.activeElement === last) {
          e.preventDefault();
          first?.focus();
        }
      }
    }
    el.addEventListener("keydown", key);
    return () => {
      el.removeEventListener("keydown", key);
      previous?.focus();
    };
  }, []);
  return (
    <div
      className="modal-backdrop"
      onMouseDown={(e) => {
        if (e.target === e.currentTarget) onClose();
      }}
    >
      <div
        ref={ref}
        role="dialog"
        aria-modal="true"
        aria-label={title}
        className={wide ? "modal modal-wide" : "modal"}
      >
        <header>
          <h2>{title}</h2>
          <button
            className="icon-button"
            onClick={onClose}
            aria-label="關閉對話框"
          >
            <Icon name="close" />
          </button>
        </header>
        {children}
      </div>
    </div>
  );
}
export function DateControl({
  value,
  onChange,
}: {
  value: string;
  onChange: (value: string) => void;
}) {
  return (
    <input
      className="date-control"
      type="date"
      aria-label="資料日期"
      value={value}
      onChange={(e) => onChange(e.target.value)}
    />
  );
}
