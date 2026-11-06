import type { ReactNode } from "react";

const imports = [
  {
    kind: "index",
    label: "加權指數",
    columns:
      "A欄：日期；C欄：最高價；D欄：最低價；E欄：收盤價。單位：指數點數。",
  },
  {
    kind: "foreign",
    label: "外資",
    columns: "A欄：日期；B欄：上市買賣超金額(百萬)。單位：百萬元。",
  },
  {
    kind: "trust",
    label: "投信",
    columns: "A欄：日期；B欄：上市買賣超金額(百萬)。單位：百萬元。",
  },
  {
    kind: "dealer",
    label: "自營商",
    columns: "A欄：日期；B欄：上市買賣超金額(百萬)。單位：百萬元。",
  },
] as const;

export function CMoneyImportGuide({
  renderImport,
}: {
  renderImport: (
    kind: (typeof imports)[number]["kind"],
    label: string,
  ) => ReactNode;
}) {
  return (
    <section className="cmoney-import-guide" aria-label="CMoney資料匯入說明">
      <h3>CMoney 匯出資料</h3>
      <p>不同分頁共用同一份資料；加權指數與三種法人分別保存，不會互相覆蓋。</p>
      <p>
        首次建立歷史基準時，匯入需要納入排名的歷史區間。已有歷史資料後，只需補匯新增日期或修訂資料，不必每次重匯完整歷史。
      </p>
      <p>
        請使用 CMoney 預設表格匯出的
        .xlsx，第一張工作表的欄名、順序及單位須符合下列格式。檔名與年份不限定，僅需逐份選擇類型匯入，不必每次提供全部四種類型。
      </p>
      <p>
        法人金額不需自行換算成億元；當日資料不會因匯入就視為盤後定稿。目前不支援混合類型檔案一次匯入。
      </p>
      <div className="cmoney-import-options">
        {imports.map((item) => (
          <section key={item.kind} className="cmoney-import-item">
            <h4>{item.label}</h4>
            <p className="muted">{item.columns}</p>
            {renderImport(item.kind, item.label)}
          </section>
        ))}
      </div>
    </section>
  );
}
