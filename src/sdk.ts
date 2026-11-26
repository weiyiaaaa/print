// ===== 数据层：多维表格访问、表路由、字段值归一化 =====
import { PRINT_TABLE_NAMES, isSupportedPrintTable } from "./mapping";

export { PRINT_TABLE_NAMES, isSupportedPrintTable };

export interface FieldMeta {
  fieldId: string;
  name: string;
  type: number;
  isPrimary?: boolean;
}

export interface BitableRecord {
  recordId: string;
  fields: Record<string, unknown>;
}

// 动态加载 SDK：优先 js-sdk（支持「自定义插件/边栏」本地调试），失败则回退 block-bitable-api（opdev 上传链路）
let cachedApi: any = null;
async function getBitable(): Promise<any> {
  if (cachedApi) return cachedApi;
  try {
    const mod: any = await import("@lark-base-open/js-sdk");
    cachedApi = mod.bitable ?? mod.default?.bitable;
  } catch {
    /* ignore */
  }
  if (!cachedApi) {
    try {
      const mod: any = await import("@lark-opdev/block-bitable-api");
      cachedApi = mod.bitable ?? mod.default?.bitable;
    } catch {
      /* ignore */
    }
  }
  if (!cachedApi) throw new Error("无法加载多维表格 SDK，请在飞书多维表格的插件面板中打开");
  return cachedApi;
}

function pick(obj: any, path: string): any {
  return path.split(".").reduce((o, k) => (o == null ? undefined : o[k]), obj);
}

function need(obj: any, path: string): any {
  const val = pick(obj, path);
  if (val == null) throw new Error(`SDK 缺少接口 ${path}，请检查插件环境/SDK 版本`);
  return val;
}

/**
 * 统一表接口适配器：
 * - js-sdk 1.x（对象式）：bitable.getActiveTable() -> table.getFieldMetaList()/getRecordListByPage()
 * - 旧版/官方链路：bitable.base.getActiveTableId()/getTableFields()/getRecords()
 */
export interface TableData {
  name: string;
  tableId?: string;
  fields: FieldMeta[];
  records: BitableRecord[];
  /** js-sdk 的表对象：签名图片等需要按 recordId/fieldId 单独取值时用 */
  table?: any;
}

export interface TableChoice {
  id: string;
  name: string;
}

export interface PrintTableData {
  supported: boolean;
  currentTable: TableChoice;
  targetTables: TableChoice[];
  /** 仅目标表读取字段和记录；其他表不加载业务数据 */
  data?: TableData;
}

function toFieldMeta(meta: any): FieldMeta {
  return {
    fieldId: meta.id ?? meta.fieldId,
    name: meta.name,
    type: meta.type,
    isPrimary: !!meta.isPrimary,
  };
}

/** 兼容不同 SDK 版本返回的数组、同步/异步可迭代记录集合。 */
async function toRecordArray(list: any, response: any): Promise<any[]> {
  if (Array.isArray(list)) return list;
  const records: any[] = [];
  if (typeof list?.[Symbol.asyncIterator] === "function") {
    for await (const record of list) records.push(record);
  } else if (typeof list?.[Symbol.iterator] === "function") {
    records.push(...list);
  }
  if (!records.length && response && !Array.isArray(response) && typeof response[Symbol.iterator] === "function") {
    records.push(...response);
  }
  return records;
}

/** 获取当前表（字段+全部记录，记录字段统一按字段名索引） */
export async function getTableData(): Promise<TableData> {
  const api = await getBitable();

  // ---- js-sdk 1.x：bitable.base.getActiveTable() 返回表对象；字段按 fieldId 索引，这里统一转成按字段名 ----
  const getActiveTable = pick(api, "base.getActiveTable") ?? (typeof api.getActiveTable === "function" ? api.getActiveTable : null);
  if (getActiveTable) {
    const table = await getActiveTable.call(pick(api, "base") ?? api);
    const metas: any[] = (await table.getFieldMetaList()) || [];
    const fields = metas.map(toFieldMeta);
    const id2name = new Map(fields.map((f) => [f.fieldId, f.name]));

    const records: BitableRecord[] = [];
    // 1) getRecordListByPage（stringValue: true -> 直接返回界面显示字符串，与参考插件 getCellString 同效）
    let pageToken: any = undefined;
    let guard = 0;
    do {
      const resp: any =
        typeof table.getRecordListByPage === "function"
          ? await table.getRecordListByPage(
              pageToken == null ? { pageSize: 200, stringValue: true } : { pageSize: 200, pageToken, stringValue: true }
            )
          : await table.getRecordList();
      const list = resp?.records ?? resp?.recordList ?? [];
      const pageRecords = await toRecordArray(list, resp);
      for (const r of pageRecords) {
        const byName: Record<string, unknown> = {};
        for (const [fid, val] of Object.entries(r.fields || {})) {
          const nm = id2name.get(fid);
          byName[nm || fid] = val;       // 字段名键（或退化用 fieldId 键）
          if (nm && fid) byName[fid] = val; // fieldId 双键兜底
        }
        records.push({ recordId: r.recordId ?? r.id, fields: byName });
      }
      pageToken = resp?.hasMore ? resp.pageToken ?? resp.nextPageToken : undefined;
      if (++guard > 100) break;
    } while (pageToken != null);

    // 2) 兜底：fields 为空的记录，逐格 getCellString 取显示文本（参考插件同方案）
    const emptyRecs = records.filter((r) => Object.keys(r.fields).length === 0);
    if (emptyRecs.length > 0 && typeof table.getCellString === "function") {
      for (const rec of emptyRecs.slice(0, 200)) {
        const byName: Record<string, unknown> = {};
        for (const f of fields) {
          try {
            const s = await table.getCellString(f.fieldId, rec.recordId);
            if (s != null && s !== "") {
              byName[f.name] = s;
              byName[f.fieldId] = s;
            }
          } catch { /* ignore */ }
        }
        rec.fields = byName;
      }
    }

    let name = "数据表";
    try { name = await table.getName(); } catch { /* ignore */ }
    const tableId: string | undefined = (table as any)?.id;
    return { name, tableId, fields, records, table };
  }

  // ---- 旧式 API：bitable.base.* ----
  const tableId = await need(api, "base.getActiveTableId")();
  const fnFields = pick(api, "base.getTableFields") ?? pick(api, "base.getFieldMetas");
  if (!fnFields) throw new Error("SDK 缺少接口 base.getTableFields，请检查插件环境/SDK 版本");
  const metas = await fnFields(tableId);
  const fields: FieldMeta[] = (metas || []).map(toFieldMeta);

  const fnRecords = pick(api, "base.getRecords") ?? pick(api, "base.getRecordList");
  if (!fnRecords) throw new Error("SDK 缺少接口 base.getRecords，请检查插件环境/SDK 版本");
  const records: BitableRecord[] = [];
  let offset: number | undefined = undefined;
  let guard = 0;
  do {
    const page: any = offset == null ? await fnRecords(tableId) : await fnRecords(tableId, { offset });
    const list = page.records || page.recordList || [];
    for (const r of list) {
      records.push({ recordId: r.recordId ?? r.id, fields: r.fields || {} });
    }
    offset = typeof page.offset === "number" && records.length < page.total ? page.offset : undefined;
    if (++guard > 200) break;
  } while (offset != null);

  // 表名（旧式 API 没有 getName，借 getTableById 取）
  let name = "数据表";
  try {
    const getById = pick(api, "base.getTableById");
    if (typeof getById === "function") {
      const t = await getById.call(pick(api, "base"), tableId);
      name = (await t?.getName?.()) || name;
    }
  } catch { /* ignore */ }

  return { name, tableId, fields, records };
}

/** 轻量读取当前表信息，避免仅为了判断表名就加载全部字段和记录。 */
async function getActiveTableInfo(): Promise<TableChoice> {
  const api = await getBitable();
  const base = pick(api, "base") ?? api;
  const getActiveTable = base?.getActiveTable;
  if (typeof getActiveTable === "function") {
    const table = await getActiveTable.call(base);
    let id = table?.id ? String(table.id) : "";
    if (!id && typeof table?.getMeta === "function") {
      const meta = await table.getMeta();
      id = meta?.id ? String(meta.id) : "";
    }
    let name = "数据表";
    try { name = (await table.getName()) || name; } catch { /* ignore */ }
    return { id, name };
  }

  const getActiveTableId = need(api, "base.getActiveTableId");
  const id = String(await getActiveTableId.call(base));
  let name = "数据表";
  try {
    if (typeof base?.getTableMetaById === "function") {
      const meta = await base.getTableMetaById(id);
      name = meta?.name || name;
    } else if (typeof base?.getTableById === "function") {
      const table = await base.getTableById(id);
      name = (await table?.getName?.()) || name;
    }
  } catch { /* ignore */ }
  return { id, name };
}

async function getTableChoices(api: any): Promise<TableChoice[]> {
  const base = pick(api, "base") ?? api;
  if (typeof base?.getTableMetaList === "function") {
    const metas = await base.getTableMetaList();
    return (metas || []).map((meta: any) => ({ id: String(meta.id), name: String(meta.name) }));
  }
  if (typeof base?.getTableList === "function") {
    const tables = await base.getTableList();
    return Promise.all((tables || []).map(async (table: any) => ({
      id: String(table.id ?? (await table.getMeta?.())?.id ?? ""),
      name: String((await table.getName?.()) || "数据表"),
    })));
  }
  throw new Error("当前 SDK 不支持读取数据表列表，请升级插件 SDK");
}

/** 目标表直接读取打印数据；其他表只返回两个指定表的跳转入口。 */
export async function getPrintTableData(): Promise<PrintTableData> {
  const currentTable = await getActiveTableInfo();
  if (isSupportedPrintTable(currentTable.name)) {
    return {
      supported: true,
      currentTable,
      targetTables: [],
      data: await getTableData(),
    };
  }

  const api = await getBitable();
  const allTables = await getTableChoices(api);
  const targetTables = PRINT_TABLE_NAMES.map((targetName) => {
    const table = allTables.find((item) => item.name.replace(/\s+/g, "") === targetName);
    return table ? { id: table.id, name: targetName } : null;
  }).filter((table): table is TableChoice => !!table);
  return { supported: false, currentTable, targetTables };
}

/** 切换多维表格中的激活数据表。 */
export async function switchToTable(tableId: string): Promise<boolean> {
  const api = await getBitable();
  const ui = pick(api, "ui");
  if (typeof ui?.switchToTable === "function") return !!(await ui.switchToTable(tableId));
  if (typeof ui?.switchBlock === "function") return !!(await ui.switchBlock(tableId));
  throw new Error("当前 SDK 不支持切换数据表，请升级插件 SDK");
}

/** 读取当前激活表的 ID（js-sdk 走 getActiveTable().id，旧式走 getActiveTableId） */
async function readActiveTableId(): Promise<string | null> {
  try {
    const api = await getBitable();
    const base = pick(api, "base") ?? api;
    const getActiveTable = base?.getActiveTable;
    if (typeof getActiveTable === "function") {
      const t = await getActiveTable.call(base);
      const id = t?.id;
      if (id) return String(id);
    }
    const fn = base?.getActiveTableId;
    if (typeof fn === "function") {
      const id = await fn.call(base);
      if (id) return String(id);
    }
  } catch { /* ignore */ }
  return null;
}

/**
 * 监听「切换数据表」：用户在多维表格页签里换表时回调。
 * - 以 SelectionChange 为主（事件带 tableId），TableAdd/TableDelete 作补充
 * - 同一张表内的选区变化不触发
 * 返回取消监听的函数。
 */
export async function onActiveTableChange(cb: (tableId: string) => void): Promise<() => void> {
  const api = await getBitable();
  const base = pick(api, "base") ?? api;
  let last = await readActiveTableId();
  let lastProbe = 0;

  const handle = (e: any) => {
    const id = e?.data?.tableId;
    if (id) {
      if (id === last) return;
      last = id;
      cb(id);
      return;
    }
    // 部分环境的事件不带 tableId：低频兜底比对一次当前表，避免频繁穿透 SDK
    const now = Date.now();
    if (now - lastProbe < 500) return;
    lastProbe = now;
    readActiveTableId()
      .then((cur) => {
        if (cur && cur !== last) {
          last = cur;
          cb(cur);
        }
      })
      .catch(() => { /* ignore */ });
  };

  const offs: (() => void)[] = [];
  for (const ev of ["onSelectionChange", "onTableAdd", "onTableDelete"]) {
    const fn = (base as any)?.[ev];
    if (typeof fn !== "function") continue;
    try {
      const off = fn.call(base, handle);
      if (typeof off === "function") offs.push(off);
    } catch { /* 该事件不受支持，跳过 */ }
  }

  return () => offs.forEach((off) => { try { off(); } catch { /* ignore */ } });
}

function pad(n: number) { return n < 10 ? "0" + n : "" + n; }

function fmtDate(ts: number, fmt = "YYYY-MM-DD"): string {
  const d = new Date(ts);
  if (isNaN(d.getTime())) return "";
  return fmt
    .replace("YYYY", "" + d.getFullYear())
    .replace("MM", pad(d.getMonth() + 1))
    .replace("DD", pad(d.getDate()))
    .replace("HH", pad(d.getHours()))
    .replace("mm", pad(d.getMinutes()));
}

/** 文本段数组 [{type:'text',text:'...'}] -> string */
function joinTextSegs(v: any): string {
  if (Array.isArray(v)) return v.map((s) => (typeof s === "string" ? s : String(s?.text ?? ""))).join("");
  return String(v ?? "");
}

/** 任意字段值 -> 可打印文本，失败返回空串 */
function normalizeValue(value: unknown, fieldType?: number, dateFmt?: string): string {
  if (value == null || value === "") return "";
  try {
    if (typeof value === "string") return value;
    if (typeof value === "number") {
      // 日期时间(5)/创建时间(1001)/修改时间(1002) 字段值为毫秒时间戳；1005 是自动编号
      if (fieldType === 5 || fieldType === 1001 || fieldType === 1002) return fmtDate(value, dateFmt);
      return String(value);
    }
    if (typeof value === "boolean") return value ? "√" : "";
    if (Array.isArray(value)) {
      // 附件 [{name}] / 人员 [{name? id}] / 文本段
      if (value.length === 0) return "";
      if (value[0] && typeof value[0] === "object" && "name" in value[0]) {
        return value.map((o: any) => String(o.name ?? o.text ?? "")).filter(Boolean).join("、");
      }
      return joinTextSegs(value);
    }
    if (typeof value === "object") {
      const obj = value as any;
      // 单选/多选 {text}
      if ("text" in obj) {
        if (fieldType === 15 && obj.link) return String(obj.text ?? obj.link);
        return joinTextSegs(obj.text) || String(obj.text ?? "");
      }
      // 公式/查找引用 {value: ...}
      if ("value" in obj) return normalizeValue(obj.value, fieldType, dateFmt);
      // 双向关联 {recordIds, text}
      if ("recordIds" in obj) return obj.text ? joinTextSegs(obj.text) : (obj.recordIds || []).join(", ");
      if ("link" in obj) return String(obj.text ?? obj.link ?? "");
    }
    return String(value);
  } catch {
    return "";
  }
}

/** 按候选链在记录中取第一个非空值（兼容字段名键 / fieldId 键两种记录结构） */
export function resolveField(
  record: BitableRecord,
  candidates: string[],
  fieldsByName: Map<string, FieldMeta>,
  dateFmt?: string
): { text: string; fieldName: string } {
  for (const name of candidates) {
    const meta = fieldsByName.get(name);
    const keys = [name, meta?.fieldId];
    for (const k of keys) {
      if (k == null || !(k in record.fields)) continue;
      const text = normalizeValue(record.fields[k], meta?.type, dateFmt);
      if (text !== "") return { text, fieldName: name };
    }
  }
  return { text: "", fieldName: "" };
}

/** 字段值里所有可能的图片地址（直链 url 优先，其次临时下载地址） */
function pickUrls(v: any, out: string[] = []): string[] {
  if (!v) return out;
  if (typeof v === "string") {
    const s = v.trim();
    if (/^https?:\/\//.test(s) && !out.includes(s)) out.push(s);
    return out;
  }
  if (Array.isArray(v)) {
    for (const it of v) pickUrls(it, out);
    return out;
  }
  if (typeof v === "object") {
    for (const k of ["url", "tmp_url", "tmpUrl", "downloadUrl", "preview_url", "link"]) {
      pickUrls((v as any)[k], out);
    }
    if ("value" in v) pickUrls((v as any).value, out);
  }
  return out;
}

/**
 * 取签名等图片字段的全部候选下载地址。
 * 按可靠性依次收集：getAttachmentUrls -> getCellValue -> 已读取的记录原始值。
 * 返回多个地址，供调用方逐个尝试（不同接口给的链接有效期/鉴权方式不同）。
 */
export async function getImageUrls(table: any, record: BitableRecord, fieldName: string, fieldsByName: Map<string, FieldMeta>): Promise<string[]> {
  const meta = fieldsByName.get(fieldName);
  const recordId = record.recordId;
  const out: string[] = [];
  const add = (v: any) => {
    for (const u of pickUrls(v)) if (!out.includes(u)) out.push(u);
  };

  if (table && meta) {
    for (const call of [
      () => (typeof table.getAttachmentUrls === "function" ? table.getAttachmentUrls(recordId, meta.fieldId) : null),
      () => (typeof table.getCellValue === "function" ? table.getCellValue(meta.fieldId, recordId) : null),
    ]) {
      try {
        const r = await call();
        if (r) add(r);
      } catch { /* 继续下一种方式 */ }
    }
  }

  add(record.fields[fieldName]);
  if (meta) add(record.fields[meta.fieldId]);
  return out;
}

/** 下载图片二进制（相对地址按当前页面 origin 补全） */
export async function fetchImageBytes(url: string): Promise<Uint8Array | null> {
  try {
    const abs = /^https?:\/\//.test(url) ? url : new URL(url, location.href).href;
    // 飞书附件链接需要带上页面会话凭据，否则可能返回 401/403
    const res = await fetch(abs, { credentials: "include" });
    if (!res.ok) return null;
    const buf = await res.arrayBuffer();
    return buf.byteLength ? new Uint8Array(buf) : null;
  } catch {
    return null;
  }
}
