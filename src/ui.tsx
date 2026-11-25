// 纯展示组件：不包含飞书 SDK、工作簿填充或打印状态，界面调整集中在这里。
import type { CSSProperties, RefObject } from "react";
import type { EditableField, ImageField, SheetDef } from "./mapping";
import type { BitableRecord } from "./sdk";
import { pageCss } from "./preview";

const button: CSSProperties = {
  padding: "8px 14px", border: "1px solid #d0d0d5", borderRadius: 6,
  background: "#fff", cursor: "pointer", fontSize: 13,
};
const primaryButton: CSSProperties = {
  ...button, background: "#3370ff", color: "#fff", border: "1px solid #3370ff",
};

export const uiStyles = {
  button,
  primaryButton,
  primaryButtonDisabled: {
    ...primaryButton, background: "#f2f3f5", color: "#bbbfc4",
    border: "1px solid #e5e6eb", cursor: "not-allowed",
  } as CSSProperties,
  linkButton: {
    padding: 0, border: "none", background: "none", color: "#3370ff",
    cursor: "pointer", fontSize: 12, whiteSpace: "nowrap",
  } as CSSProperties,
  chip: {
    display: "inline-block", maxWidth: 150, padding: "4px 9px", borderRadius: 10,
    background: "#f0f2f5", color: "#646a73", fontSize: 11, lineHeight: 1.4,
    overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap", verticalAlign: "middle",
  } as CSSProperties,
};

export interface PrintPage {
  html: string;
  margins: { l: number; r: number; t: number; b: number };
  pageW: number;
  label: string;
}

interface TemplatePickerProps {
  sheets: SheetDef[];
  selectedId: string;
  onChange: (id: string) => void;
}

export function TemplatePicker({ sheets, selectedId, onChange }: TemplatePickerProps) {
  return (
    <div style={{ marginBottom: 12, flexShrink: 0 }}>
      <div style={{ color: "#646a73", marginBottom: 6 }}>选择模板</div>
      {sheets.map((sheet) => (
        <label
          key={sheet.id}
          style={{
            display: "flex", gap: 6, alignItems: "center", padding: "6px 8px",
            border: "1px solid #e0e0e5", borderRadius: 6, marginBottom: 6,
            cursor: "pointer", background: sheet.id === selectedId ? "#eff4ff" : "#fff",
          }}
        >
          <input type="radio" name="tpl" checked={sheet.id === selectedId} onChange={() => onChange(sheet.id)} />
          <div>
            <div style={{ fontWeight: 500 }}>{sheet.name}</div>
            <div style={{ fontSize: 11, color: "#8f959e" }}>{sheet.desc}</div>
          </div>
        </label>
      ))}
    </div>
  );
}

interface EditableInputsProps {
  defs: EditableField[];
  values: Record<string, string>;
  onChange: (key: string, value: string) => void;
}

export function EditableInputs({ defs, values, onChange }: EditableInputsProps) {
  if (!defs.length) return null;
  return (
    <div style={{ marginBottom: 12, flexShrink: 0 }}>
      <div style={{ color: "#646a73", marginBottom: 6 }}>表头填写</div>
      {defs.map((def) => (
        <div key={def.key} style={{ display: "flex", alignItems: "center", gap: 8 }}>
          <span style={{ color: "#646a73", whiteSpace: "nowrap" }}>{def.label}</span>
          <input
            value={values[def.key] ?? ""}
            placeholder={def.defaultValue || ""}
            onChange={(event) => onChange(def.key, event.target.value)}
            style={{
              flex: 1, minWidth: 0, padding: "5px 8px", border: "1px solid #d0d0d5",
              borderRadius: 6, fontSize: 13, color: "#1f2329",
            }}
          />
          {def.defaultValue && (
            <button style={uiStyles.linkButton} onClick={() => onChange(def.key, def.defaultValue!)}>默认</button>
          )}
        </div>
      ))}
      <div style={{ fontSize: 11, color: "#8f959e", marginTop: 5, lineHeight: 1.5 }}>
        打印时写进表单表头；清空则改用表格里的「{defs[0].fields?.[0] ?? "同名字段"}」。
      </div>
    </div>
  );
}

interface SignatureSettingsProps {
  slots: ImageField[];
  signs: Record<string, string>;
  open: boolean;
  onToggle: () => void;
  onUpload: (key: string, file: File) => void;
  onClear: (key: string) => void;
}

export function SignatureSettings({ slots, signs, open, onToggle, onUpload, onClear }: SignatureSettingsProps) {
  const count = slots.filter((slot) => !!signs[slot.uploadKey!]).length;
  return (
    <div style={{ marginBottom: 12, flexShrink: 0 }}>
      <div style={{ display: "flex", alignItems: "center", marginBottom: 6 }}>
        <span style={{ color: "#646a73" }}>签名</span>
        <span style={{ marginLeft: 6, fontSize: 11, color: count > 0 ? "#3370ff" : "#8f959e" }}>
          {count}/{slots.length} 已上传
        </span>
        <span style={{ flex: 1 }} />
        <button style={uiStyles.linkButton} onClick={onToggle}>{open ? "收起" : "管理签名"}</button>
      </div>
      {open && (
        <div style={{ border: "1px solid #e0e0e5", borderRadius: 6, background: "#fff" }}>
          {slots.map((slot) => {
            const key = slot.uploadKey!;
            return (
              <div
                key={key}
                style={{ display: "flex", alignItems: "center", gap: 8, padding: "6px 10px", borderBottom: "1px solid #f0f0f2" }}
              >
                <span style={{ flex: 1, fontSize: 12, minWidth: 0, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>
                  {slot.label || slot.fields[0]}
                </span>
                {signs[key] ? (
                  <>
                    <img
                      src={signs[key]} alt=""
                      style={{ maxWidth: 56, maxHeight: 22, objectFit: "contain", border: "1px solid #f0f0f2", background: "#fff" }}
                    />
                    <button style={uiStyles.linkButton} onClick={() => onClear(key)}>清除</button>
                  </>
                ) : <span style={{ fontSize: 11, color: "#bbbfc4" }}>未上传</span>}
                <label style={uiStyles.linkButton}>
                  上传
                  <input
                    type="file" accept="image/*" style={{ display: "none" }}
                    onChange={(event) => {
                      const file = event.target.files?.[0];
                      event.target.value = "";
                      if (file) onUpload(key, file);
                    }}
                  />
                </label>
              </div>
            );
          })}
          <div style={{ padding: "6px 10px", fontSize: 11, color: "#8f959e", lineHeight: 1.6 }}>
            表格里有真实签名图时优先用它；没有、且签批状态为「已确认」时，自动填这里上传的签名。
          </div>
        </div>
      )}
    </div>
  );
}

interface RecordSelectorProps {
  records: BitableRecord[];
  filtered: BitableRecord[];
  selected: Set<string>;
  search: string;
  loading: boolean;
  recordLabel: (record: BitableRecord) => string;
  onSearch: (value: string) => void;
  onToggle: (id: string) => void;
  onToggleAll: (all: boolean) => void;
}

export function RecordSelector(props: RecordSelectorProps) {
  const { records, filtered, selected, search, loading, recordLabel, onSearch, onToggle, onToggleAll } = props;
  return (
    <>
      <div style={{ marginBottom: 8, display: "flex", gap: 8, alignItems: "center", flexShrink: 0 }}>
        <label style={{ display: "flex", gap: 5, alignItems: "center", cursor: "pointer" }}>
          <input
            type="checkbox"
            checked={filtered.length > 0 && filtered.every((record) => selected.has(record.recordId))}
            onChange={(event) => onToggleAll(event.target.checked)}
          />
          <span style={{ color: "#646a73" }}>全选</span>
        </label>
        <span style={{ color: "#646a73" }}>已选 {selected.size}/{records.length}</span>
        <span style={{ flex: 1 }} />
        <input
          placeholder="搜索…" value={search} onChange={(event) => onSearch(event.target.value)}
          style={{ padding: "4px 8px", border: "1px solid #d0d0d5", borderRadius: 6, width: 130 }}
        />
      </div>
      <div style={{ flex: 1, minHeight: 90, overflowY: "auto", border: "1px solid #e0e0e5", borderRadius: 6, marginBottom: 10 }}>
        {filtered.map((record) => (
          <label
            key={record.recordId}
            style={{ display: "flex", gap: 8, alignItems: "center", padding: "7px 10px", borderBottom: "1px solid #f0f0f2", cursor: "pointer" }}
          >
            <input type="checkbox" checked={selected.has(record.recordId)} onChange={() => onToggle(record.recordId)} />
            <span style={{ overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{recordLabel(record)}</span>
          </label>
        ))}
        {filtered.length === 0 && !loading && <div style={{ padding: 14, color: "#8f959e" }}>当前数据表没有记录</div>}
      </div>
    </>
  );
}

export function LoadingOverlay({ tableName }: { tableName: string }) {
  return (
    <div
      style={{
        position: "absolute", inset: 0, display: "flex", alignItems: "center", justifyContent: "center",
        gap: 8, background: "rgba(255,255,255,.78)", color: "#646a73", fontSize: 12, zIndex: 5,
      }}
    >
      <span className="pl-spinner" />
      <span>{tableName ? `正在读取「${tableName}」…` : "正在读取多维表格…"}</span>
    </div>
  );
}

interface PreviewPanelProps {
  pages: PrintPage[];
  tableName: string;
  sheetName: string;
  signHint: string;
  error: string;
  frameZoom: number;
  frameBox: RefObject<HTMLDivElement>;
  onFrameLoad: () => void;
  onBack: () => void;
  onPrint: () => void;
}

export function PreviewPanel(props: PreviewPanelProps) {
  const { pages, tableName, sheetName, signHint, error, frameZoom, frameBox, onFrameLoad, onBack, onPrint } = props;
  const { margins, pageW } = pages[0];
  const srcDoc = `<html><head><meta charset="utf-8"><style>${pageCss(margins, pageW, frameZoom)}</style></head><body>${pages.map((page) => page.html).join("")}</body></html>`;
  return (
    <div style={{ display: "flex", flexDirection: "column", height: "100vh", background: "#f2f3f5" }}>
      <div style={{ padding: 10, display: "flex", gap: 8, alignItems: "center", borderBottom: "1px solid #e0e0e5", background: "#fff" }}>
        <strong style={{ fontSize: 14 }}>{tableName ? `${tableName} · ` : ""}{sheetName} × {pages.length} 份</strong>
        <button style={uiStyles.button} onClick={onBack}>返回</button>
        <span style={{ flex: 1 }} />
        {signHint && (
          <span style={{ fontSize: 11, color: signHint.includes("空白") || signHint.includes("仅有地址") ? "#e24b4a" : "#8f959e" }}>
            {signHint}
          </span>
        )}
        <span style={{ fontSize: 11, color: "#8f959e" }}>打印设置：A4、纵向、{pageW}mm 宽</span>
        <button style={uiStyles.primaryButton} onClick={onPrint}>打印</button>
      </div>
      {error && <div style={{ color: "#e24b4a", padding: "6px 12px", fontSize: 12 }}>{error}</div>}
      <div ref={frameBox} style={{ flex: 1, overflow: "auto" }}>
        <iframe
          id="print-frame" title="preview" onLoad={onFrameLoad} srcDoc={srcDoc}
          style={{ width: "100%", height: "100%", border: "none" }}
        />
      </div>
    </div>
  );
}
