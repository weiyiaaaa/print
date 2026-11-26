// ===== 主界面：选模板 -> 勾选记录 -> 直接打印（点一次调起打印对话框） =====
// 模板与字段映射全部来自 mapping.ts；打印页由填好的工作簿实时渲染。
// 多维表格里切换数据表时，插件自动刷新数据并把表名显示在标题与打印按钮旁。
import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import { SHEETS, sheetsForTable, matchStatus } from "./mapping";
import type { SheetDef, ImageField, EditableField } from "./mapping";
import { resolveField, getPrintTableData, getImageUrls, fetchImageBytes, onActiveTableChange, switchToTable } from "./sdk";
import type { FieldMeta, BitableRecord, TableChoice } from "./sdk";
import { fillSheet, loadWorkbook } from "./excel";
import { sheetToHtml } from "./preview";
import { loadSignatures, saveSignature, removeSignature, dataUrlToBytes, fileToSignatureDataUrl } from "./signatures";
import {
  EditableInputs,
  LoadingOverlay,
  PreviewPanel,
  RecordSelector,
  SignatureSettings,
  TableChoicePanel,
  TemplatePicker,
  uiStyles,
} from "./ui";
import type { PrintPage } from "./ui";

/** 读字段文本（字符串 / 选项 / 人员等多形态） */
function fieldText(record: BitableRecord, name: string): string {
  const v = record.fields[name];
  if (v == null) return "";
  if (typeof v === "string") return v;
  if (typeof v === "number") return String(v);
  if (Array.isArray(v)) {
    return v.map((x) => (typeof x === "string" ? x : (x && (x.text || x.name)) || "")).join(" ");
  }
  if (typeof v === "object") return String((v as any).text ?? (v as any).value ?? (v as any).name ?? "");
  return "";
}

/** 该签名位是否已具备填写条件：没配状态字段就总是可填；配了就要求状态命中关键词 */
function isConfirmed(record: BitableRecord, image: ImageField): boolean {
  const fields = image.statusField;
  if (!fields?.length) return true;
  return fields.some((field) => matchStatus(fieldText(record, field), image.statusOk, image.statusNot));
}

function errorMessage(error: unknown, fallback: string): string {
  return error instanceof Error && error.message ? error.message : fallback;
}


// 每个模板的每一项分开存，下次打开插件还是上次填的值。

const editKey = (sheetId: string, key: string) => `pp.edit.${sheetId}.${key}`;

function loadEdits(sheetId: string, defs: EditableField[]): Record<string, string> {
  const out: Record<string, string> = {};
  for (const d of defs) {
    let saved: string | null = null;
    try { saved = localStorage.getItem(editKey(sheetId, d.key)); } catch { /* 隐私模式忽略 */ }
    out[d.key] = saved ?? d.defaultValue ?? "";
  }
  return out;
}

function saveEdit(sheetId: string, key: string, value: string): void {
  try {
    if (value === "") localStorage.removeItem(editKey(sheetId, key));
    else localStorage.setItem(editKey(sheetId, key), value);
  } catch { /* 忽略 */ }
}

export default function App() {
  const [sheetId, setSheetId] = useState(SHEETS[0].id);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState("");
  const [fields, setFields] = useState<FieldMeta[]>([]);
  const [records, setRecords] = useState<BitableRecord[]>([]);
  const [table, setTable] = useState<any>(null);
  const [tableName, setTableName] = useState("");
  const [tableChoiceRequired, setTableChoiceRequired] = useState(false);
  const [tableChoices, setTableChoices] = useState<TableChoice[]>([]);
  const [switchingTableId, setSwitchingTableId] = useState<string | null>(null);
  const [selected, setSelected] = useState<Set<string>>(new Set());
  const [search, setSearch] = useState("");
  const [preview, setPreview] = useState<PrintPage[] | null>(null);
  const [busy, setBusy] = useState(false);
  const [frameZoom, setFrameZoom] = useState(1);
  const frameBox = useRef<HTMLDivElement>(null);
  const pendingPrint = useRef(false);
  const seqRef = useRef(0); // 数据请求序号：切表时旧的请求结果直接丢弃
  // 签名取图统计（仅用于预览页给出提示，不影响打印）
  const signStat = useRef({ bytes: 0, urlOnly: 0, uploaded: 0, none: 0 });
  const [signHint, setSignHint] = useState("");
  // 上传的签名图（uploadKey -> dataURL），存 localStorage
  const [signs, setSigns] = useState<Record<string, string>>({});
  const [signPanel, setSignPanel] = useState(false);
  // 表头可编辑文本（单据编号等）：key -> 当前值
  const [edits, setEdits] = useState<Record<string, string>>({});

  // 飞书侧栏比 A4（794px）窄，屏幕预览整体等比缩小，避免被横向裁掉
  useEffect(() => {
    if (!preview) return;
    const el = frameBox.current;
    if (!el) return;
    const calc = () => {
      const w = el.clientWidth - 16;
      if (w > 0) setFrameZoom(Math.min(1, w / 794));
    };
    calc();
    const ro = new ResizeObserver(calc);
    ro.observe(el);
    return () => ro.disconnect();
  }, [preview]);

  /** 读取当前数据表并重置界面（首次加载与切换表都走这里） */
  const refresh = useCallback(async () => {
    const seq = ++seqRef.current;
    setLoading(true);
    setError("");
    setPreview(null);
    try {
      const result = await getPrintTableData();
      if (seq !== seqRef.current) return; // 快速连续切表时丢弃过期结果
      setTableName(result.currentTable.name);
      setTableChoiceRequired(!result.supported);
      setTableChoices(result.targetTables);
      setSelected(new Set()); // 默认不勾选，由用户自己挑记录
      setSearch("");
      if (!result.supported) {
        setFields([]);
        setRecords([]);
        setTable(null);
        return;
      }

      const t = result.data!;
      setFields(t.fields);
      setRecords(t.records);
      setTable(t.table ?? null);
      // 模板跟随数据表：当前选中的模板不在该表的可用范围内时，切到该表的模板
      const forTable = sheetsForTable(result.currentTable.name);
      setSheetId((cur) => (forTable.some((s) => s.id === cur) ? cur : forTable[0].id));
    } catch (error: unknown) {
      if (seq !== seqRef.current) return;
      setError(errorMessage(error, "读取多维表格数据失败"));
      setFields([]);
      setRecords([]);
      setTable(null);
      setTableName("");
      setTableChoices([]);
      setTableChoiceRequired(false);
    } finally {
      if (seq === seqRef.current) setLoading(false);
    }
  }, []);

  // 首次加载 + 监听多维表格页签切换：换表即重新读取并刷新界面
  useEffect(() => {
    let alive = true;
    let off: (() => void) | null = null;

    setSigns(loadSignatures()); // 载入本地上传的签名
    refresh();
    (async () => {
      try {
        const un = await onActiveTableChange(() => { if (alive) refresh(); });
        if (alive) off = un;
        else un();
      } catch { /* 环境不支持事件监听时静默降级，仍可手动重开插件 */ }
    })();

    return () => {
      alive = false;
      off?.();
    };
  }, [refresh]);

  // 可用模板由当前数据表名决定：切到哪张表，就只出现那张表对应的模板
  const availSheets = useMemo(() => sheetsForTable(tableName), [tableName]);
  const sheet: SheetDef = availSheets.find((s) => s.id === sheetId) || availSheets[0];
  const fieldsByName = useMemo(() => new Map(fields.map((field) => [field.name, field])), [fields]);
  const recordsById = useMemo(() => new Map(records.map((record) => [record.recordId, record])), [records]);
  const primary = fields.find((field) => field.isPrimary)?.name ?? fields[0]?.name;
  // 可上传签名的位置（映射里配了 uploadKey 的）
  const signSlots = (sheet?.images || []).filter((im) => !!im.uploadKey);
  // 表头可编辑项（当前模板配了才有）
  const editDefs: EditableField[] = sheet?.editable ?? [];

  // 切模板时载入该模板上次填的表头文本（没存过就用默认值）
  const sheetKey = sheet?.id ?? "";
  useEffect(() => {
    if (!sheetKey) return;
    setEdits(loadEdits(sheetKey, editDefs));
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [sheetKey]);

  const setEdit = (key: string, value: string) => {
    setEdits((cur) => ({ ...cur, [key]: value }));
    saveEdit(sheetKey, key, value);
  };

  const recordLabel = useCallback((record: BitableRecord): string => {
    if (!primary) return "(无标题)";
    return resolveField(record, [primary], fieldsByName).text || "(无标题)";
  }, [primary, fieldsByName]);

  const filtered = useMemo(
    () => records.filter((record) => !search || recordLabel(record).includes(search)),
    [records, search, recordLabel]
  );

  const toggle = (id: string) => {
    setSelected((current) => {
      const next = new Set(current);
      if (next.has(id)) next.delete(id);
      else next.add(id);
      return next;
    });
  };
  const toggleAll = (all: boolean) => setSelected(all ? new Set(filtered.map((r) => r.recordId)) : new Set());

  /**
   * 取签名图，三级优先：
   * 1) 多维表格里的真实签名附件（逐个候选地址试下载）
   * 2) 取不到且有地址 → 交给预览用 <img src> 直接显示
   * 3) 都没有且签批状态已确认 → 用本地上传的签名图兜底
   */
  async function loadSignature(
    record: BitableRecord,
    names: string[],
    im?: ImageField
  ): Promise<{ bytes?: Uint8Array; url?: string } | null> {
    for (const name of names) {
      const urls = await getImageUrls(table, record, name, fieldsByName);
      if (!urls.length) continue;
      for (const u of urls) {
        const bytes = await fetchImageBytes(u);
        if (bytes && bytes.length) {
          signStat.current.bytes++;
          return { bytes };
        }
      }
      signStat.current.urlOnly++;      // 有地址但下不到二进制：预览用 URL 兜底
      return { url: urls[0] };
    }

    // 3) 上传签名兜底：签批状态已确认（或该位不设状态字段）才填
    if (im?.uploadKey && isConfirmed(record, im)) {
      const dataUrl = signs[im.uploadKey];
      const bytes = dataUrl ? dataUrlToBytes(dataUrl) : null;
      if (bytes) {
        signStat.current.uploaded++;
        return { bytes };
      }
    }

    signStat.current.none++;           // 字段没值或字段名不匹配
    return null;
  }

  /** 上传某个签名位的签名图（压缩后存本地） */
  async function uploadSign(key: string, file: File | null | undefined) {
    if (!file) return;
    try {
      const dataUrl = await fileToSignatureDataUrl(file);
      saveSignature(key, dataUrl);
      setSigns(loadSignatures());
    } catch (error: unknown) {
      setError(errorMessage(error, "签名图片处理失败"));
    }
  }

  function clearSign(key: string) {
    removeSignature(key);
    setSigns(loadSignatures());
  }

  /** 一条记录 -> 填好的工作簿（不落盘，直接用于渲染打印页） */
  async function buildWorkbook(record: BitableRecord) {
    const wb = await loadWorkbook(sheet.file);
    await fillSheet(wb, sheet, record, fieldsByName, (names, im) => loadSignature(record, names, im), edits);
    return wb;
  }

  /** 选中的记录逐条填充，生成打印页 */
  async function buildPages(): Promise<PrintPage[]> {
    const results: PrintPage[] = [];
    signStat.current = { bytes: 0, urlOnly: 0, uploaded: 0, none: 0 };
    for (const id of selected) {
      const rec = recordsById.get(id)!;
      const wb = await buildWorkbook(rec);
      const { html, margins, pageWmm } = sheetToHtml(wb, wb.worksheets[0], { shift: sheet.textShift });
      results.push({ html, margins, pageW: pageWmm, label: recordLabel(rec) });
    }
    const st = signStat.current;
    const parts: string[] = [];
    if (st.bytes) parts.push(`${st.bytes} 张表格图`);
    if (st.urlOnly) parts.push(`${st.urlOnly} 张仅有地址`);
    if (st.uploaded) parts.push(`${st.uploaded} 张上传签名`);
    if (st.none) parts.push(`${st.none} 张空白`);
    setSignHint(`签名：${parts.join(" / ") || "—"}`);
    return results;
  }

  /** 直接打印：生成页面 -> 载入预览框 -> 自动调起打印对话框 */
  async function doPrint() {
    setBusy(true);
    setError("");
    try {
      const pages = await buildPages();
      if (!pages.length) return;
      pendingPrint.current = true;
      setPreview(pages);
    } catch (error: unknown) {
      setError(errorMessage(error, "生成打印页失败"));
    } finally {
      setBusy(false);
    }
  }

  function iframePrint() {
    const f = document.getElementById("print-frame") as HTMLIFrameElement | null;
    try {
      f?.contentWindow?.focus();
      f?.contentWindow?.print();
    } catch {
      setError("当前环境不支持直接调起打印，请改用浏览器菜单打印本页");
    }
  }

  /** 预览页加载完成后自动打印一次（用户点了「打印」时） */
  function onFrameLoad() {
    if (!pendingPrint.current) return;
    pendingPrint.current = false;
    setTimeout(iframePrint, 120); // 等图片完成解码
  }

  async function choosePrintTable(tableId: string) {
    if (switchingTableId) return; // 同一时刻只允许一次跳转；刷新请求由 seqRef 执行替换策略
    setSwitchingTableId(tableId);
    setError("");
    try {
      const switched = await switchToTable(tableId);
      if (!switched) throw new Error("未能跳转到目标数据表，请重试");
      // 某些环境会同时触发 SelectionChange；seqRef 会丢弃其中一次过期读取。
      await refresh();
    } catch (error: unknown) {
      setError(errorMessage(error, "跳转数据表失败，请重试"));
    } finally {
      setSwitchingTableId(null);
    }
  }

  if (tableChoiceRequired) {
    return (
      <TableChoicePanel
        currentTableName={tableName}
        tables={tableChoices}
        switchingId={switchingTableId}
        error={error}
        onSelect={choosePrintTable}
      />
    );
  }

  if (preview) {
    return (
      <PreviewPanel
        pages={preview}
        tableName={tableName}
        sheetName={sheet.name}
        signHint={signHint}
        error={error}
        frameZoom={frameZoom}
        frameBox={frameBox}
        onFrameLoad={onFrameLoad}
        onBack={() => setPreview(null)}
        onPrint={iframePrint}
      />
    );
  }

  const printDisabled = busy || loading || selected.size === 0;

  return (
    <div
      style={{
        display: "flex",
        flexDirection: "column",
        height: "100vh",
        boxSizing: "border-box",
        padding: "14px 14px 0",
        fontSize: 13,
        color: "#1f2329",
        fontFamily: "-apple-system,'PingFang SC','Microsoft YaHei',sans-serif",
      }}
    >
      <h3 style={{ margin: "0 0 10px", fontSize: 15, display: "flex", alignItems: "center", gap: 6, minWidth: 0, flexShrink: 0 }}>
        <span>打印</span>
        {tableName && (
          <span style={{ fontSize: 12, fontWeight: 400, color: "#8f959e", overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>
            · {tableName}
          </span>
        )}
      </h3>

      {error && (
        <div style={{ color: "#e24b4a", marginBottom: 10, lineHeight: 1.6, flexShrink: 0 }}>
          <div>{error}</div>
          <div style={{ color: "#8f959e", fontSize: 11, marginTop: 4 }}>提示：本插件需在飞书多维表格的插件面板中运行。</div>
        </div>
      )}

      {/* 主体：高度随面板自适应，记录多时只有列表滚动；读取/切表时盖一层加载动画 */}
      <div style={{ flex: 1, minHeight: 0, position: "relative", display: "flex", flexDirection: "column", overflowY: "auto" }}>
        <TemplatePicker sheets={availSheets} selectedId={sheet.id} onChange={setSheetId} />

        <EditableInputs defs={editDefs} values={edits} onChange={setEdit} />

        <SignatureSettings
          slots={signSlots}
          signs={signs}
          open={signPanel}
          onToggle={() => setSignPanel((value) => !value)}
          onUpload={uploadSign}
          onClear={clearSign}
        />

        <RecordSelector
          records={records}
          filtered={filtered}
          selected={selected}
          search={search}
          loading={loading}
          recordLabel={recordLabel}
          onSearch={setSearch}
          onToggle={toggle}
          onToggleAll={toggleAll}
        />

        {loading && <LoadingOverlay tableName={tableName} />}
      </div>

      {/* 页尾置底：打印按钮与当前表名 */}
      <div
        style={{
          flexShrink: 0,
          display: "flex",
          gap: 8,
          alignItems: "center",
          padding: "10px 0 12px",
          marginTop: 8,
          borderTop: "1px solid #eef0f2",
          background: "#fff",
        }}
      >
        <button style={{ ...(printDisabled ? uiStyles.primaryButtonDisabled : uiStyles.primaryButton), flex: 1 }} disabled={printDisabled} onClick={doPrint}>
          {busy ? "生成中…" : `打印（${selected.size} 份 · A4）`}
        </button>
        <span style={uiStyles.chip} title={tableName}>当前表：{tableName || "—"}</span>
      </div>
    </div>
  );
}
