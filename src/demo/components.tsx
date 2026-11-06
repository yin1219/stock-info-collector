export {
  Icon,
  PageHeader,
  TextAction,
  Switch,
  Tabs,
  Modal,
  DateControl,
} from "../renderer/ui";
import { Modal } from "../renderer/ui";
export function SourceDialog({ onClose }: { onClose: () => void }) {
  return (
    <Modal title="資料來源" onClose={onClose}>
      <p>這是示意資料，沒有連線官方網站或帳號。</p>
      <p className="muted">正式版本會在這裡顯示公告來源與資料日期。</p>
      <div className="dialog-actions">
        <button onClick={onClose}>關閉</button>
      </div>
    </Modal>
  );
}
export function FileDemo({
  label,
  onSelect,
  accept = ".xlsx,.csv,.json",
}: {
  label: string;
  onSelect: (name: string) => void;
  accept?: string;
}) {
  return (
    <label className="file-button">
      {label}
      <input
        type="file"
        accept={accept}
        onChange={(e) => {
          if (e.target.files?.[0]) onSelect(e.target.files[0].name);
          e.target.value = "";
        }}
      />
    </label>
  );
}
