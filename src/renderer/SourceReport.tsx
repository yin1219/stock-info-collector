interface Source {
  source: string;
  status: string;
  dataDate: string | null;
  recordCount?: number;
  errorMessage?: string | null;
}
const statuses: Record<string, string> = {
  complete: "資料完整",
  degraded: "部分降級",
  stale: "官方尚未更新",
  failed: "來源失敗",
  active: "檢查中",
};
export function MaterialSourceReport({
  monitor,
}: {
  monitor: {
    status: string;
    errorMessage: string | null;
    sources: Source[];
  } | null;
}) {
  return (
    <section role="status" aria-label="重大訊息來源狀態">
      <p>
        {!monitor
          ? "尚未檢查"
          : monitor.status === "degraded"
            ? "資料來源部分降級"
            : (statuses[monitor.status] ?? monitor.status)}
      </p>
      {monitor?.errorMessage && <p>{monitor.errorMessage}</p>}
      <ul>
        {monitor?.sources.map((source) => {
          const name =
            source.source === "TWSE-reconciliation"
              ? "上市結構化每日對帳"
              : source.source === "TPEX-reconciliation"
                ? "上櫃結構化每日對帳"
                : source.source;
          const details = [
            source.dataDate
              ? `${source.source.includes("reconciliation") ? "出表日期" : "資料日期"} ${source.dataDate}`
              : "",
            source.recordCount !== undefined
              ? `來源回傳 ${source.recordCount} 筆`
              : "",
          ]
            .filter(Boolean)
            .join("，");
          return (
            <li key={source.source}>
              {name}：{statuses[source.status] ?? source.status}
              {details && `（${details}）`}
              {source.errorMessage && <p>{source.errorMessage}</p>}
            </li>
          );
        })}
      </ul>
      {monitor?.sources.some((row) =>
        row.source.includes("reconciliation"),
      ) && (
        <p className="muted">
          對帳的出表日期可能晚於公告發言日期；對帳筆數不等於當日新公告筆數。
        </p>
      )}
    </section>
  );
}
