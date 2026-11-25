// ===== 字段映射表（唯一需要维护的文件）=====
// 一个模板 = 一份 xlsx + 一张映射表。以后加字段/加模板，只改这里。
//
// 四种映射：
//   cells  单个单元格：把字段值写进该格（自动居中）
//   blocks 合并区：先把这段空白单元格合并成一个整体，再整体填入（明细/长文本用）
//   checks 复选框：值等于 when 时，把该格第 box 个 □ 涂成 ■（判定结果/等级/评审结果）
//   images 签名：从多维表格附件里取图片，按单元格大小插入（不是写文字）
//
// 坐标写合并区的左上单元格即可（如 C10:D15 写 "C10:D15"）。
// 模板行高、列宽、页面设置一律不动 —— 打印出来的表单长度与源文件一致。
import reviewXlsx from "./templates/template_review.xlsx";
import fatXlsx from "./templates/template_fat.xlsx";

// ---------- 值转换（按需在条目上挂 fmt） ----------

/** 日期：时间戳 -> yyyy-MM-dd（sdk 已按字段类型转日期，此处兜底文本型时间戳） */
export function fmtDate(raw: string): string {
  const t = (raw ?? "").trim();
  if (!/^\d{10,13}$/.test(t)) return t;
  const d = new Date(t.length === 13 ? Number(t) : Number(t) * 1000);
  if (isNaN(d.getTime())) return t;
  const p = (n: number) => String(n).padStart(2, "0");
  return `${d.getFullYear()}-${p(d.getMonth() + 1)}-${p(d.getDate())}`;
}

/**
 * FAT检开始时间 → [月, 日, 时]：兼容时间戳 / 日期字符串（含上午下午）。
 * 用于填「现申请：___月___日___时」的三处空白；没有时间时第三段留空。
 */
export function fmtApplyTime(raw: string): string[] {
  const t = (raw ?? "").trim();
  if (!t) return [];
  const p = (n: number) => String(n).padStart(2, "0");
  const ts = t.match(/^(\d{10,13})$/);
  if (ts) {
    const d = new Date(ts[1].length === 13 ? Number(ts[1]) : Number(ts[1]) * 1000);
    return isNaN(d.getTime()) ? [] : [p(d.getMonth() + 1), p(d.getDate()), p(d.getHours())];
  }
  const hh12 = (h: number, text: string) => {
    if (/下午|PM/i.test(text) && h < 12) return h + 12;
    if (/上午|AM/i.test(text) && h === 12) return 0;
    return h;
  };
  const m = t.match(/(\d{4})[-/.年](\d{1,2})[-/.月](\d{1,2})(?:日)?[\sT]*(?:上午|下午|AM|PM)?\s*(\d{1,2})?/i);
  if (m) return [p(Number(m[2])), p(Number(m[3])), m[4] ? p(hh12(Number(m[4]), t)) : ""];
  const m2 = t.match(/(\d{1,2})月(\d{1,2})日[\s]*(?:上午|下午|AM|PM)?\s*(\d{1,2})?/i);
  if (m2) return [p(Number(m2[1])), p(Number(m2[2])), m2[3] ? p(hh12(Number(m2[3]), t)) : ""];
  const d = new Date(t);
  return isNaN(d.getTime()) ? [] : [p(d.getMonth() + 1), p(d.getDate()), p(d.getHours())];
}

/** 多行文本：按换行/逗号/顿号拆成一行一条（合并区里直接显示为多行） */
export function fmtLines(raw: string): string {
  return (raw ?? "")
    .split(/[\r\n,，、]+/)
    .map((x) => x.replace(/^\s*\d+[.、)）]\s*/, "").trim())
    .filter(Boolean)
    .join("\n");
}

/** 取多行文本的第 n 行：把「附条件事项清单」按条填进 1/2/3/4/5 号格子 */
export function nthLine(n: number): (raw: string) => string {
  return (raw: string) => {
    const lines = fmtLines(raw).split("\n").filter(Boolean);
    return lines[n] ?? "";
  };
}

/** 判定结果归一化：合格 / 不合格 / 附条件合格（顺序：先认附条件，再认不合格） */
export function fmtVerdict(raw: string): string {
  const t = raw ?? "";
  if (/附条件/.test(t)) return "附条件合格";
  if (/不合格|禁止/.test(t)) return "不合格";
  if (/合格|允许/.test(t)) return "合格";
  return t;
}

/** 影响度等级归一化：一/二/三/四级 或罗马数字 -> I/II/III/IV */
export function fmtGrade(raw: string): string {
  const t = (raw ?? "").trim().toUpperCase();
  if (t.includes("四")) return "IV";
  if (t.includes("三")) return "III";
  if (t.includes("二")) return "II";
  if (t.includes("一")) return "I";
  if (/IV|Ⅳ/.test(t)) return "IV";
  if (/III|Ⅲ/.test(t)) return "III";
  if (/II|Ⅱ/.test(t)) return "II";
  if (/I|Ⅰ/.test(t)) return "I";
  return t;
}

/** 评审结论归一化：允许发机 / 不允许发机 */
export function fmtReview(raw: string): string {
  const t = raw ?? "";
  if (/不允许/.test(t)) return "不允许发机";
  if (/允许/.test(t)) return "允许发机";
  return t;
}

/** 确认结果归一化：完成 / 附条件完成（先认附条件，避免「附条件完成」被「完成」抢先命中） */
export function fmtDone(raw: string): string {
  const t = (raw ?? "").trim();
  if (!t) return "";
  if (/附条件/.test(t)) return "附条件完成";
  if (/完成|已完/.test(t)) return "完成";
  return t;
}

/** 添附资料归一化：有 / 无 */
export function fmtAttach(raw: string): string {
  const t = (raw ?? "").trim();
  if (!t) return "";
  if (/无/.test(t) && !/有/.test(t)) return "无";
  if (/有/.test(t)) return "有";
  return t;
}

// ---------- 类型 ----------

/** 单元格映射：值写入该格并居中；append=true 时保留模板原文字并追加 */
export interface CellField {
  cell: string;
  fields: string[];
  append?: boolean;
  wrap?: boolean;
  fmt?: (raw: string) => string;
  /** 水平对齐，默认居中 */
  align?: "left" | "center" | "right";
  /** 条件写入：whenField 的取值命中 whenValues 之一时才写（同一字段对应两个分支格子时用） */
  whenField?: string[];
  whenValues?: string[];
  /** 填空式写入：把模板原文中这两段之间的空白替换成值（如「第　　　回」填申请回数） */
  insert?: [string, string];
  /** 多段填空：模板里连续几处「前文+空白+后文」，把字段值拆成几段按序填入（如「__月__日__时」） */
  inserts?: { slots: [string, string][]; fmt: (raw: string) => string[] };
  /** 日期字段的取文本格式；默认 YYYY-MM-DD，需要小时时传 "YYYY-MM-DD HH" */
  dateFmt?: string;
  /** 超长自动缩字号（不换行、不改行高），保证文字落在本格 / 真实合并区内 */
  fit?: boolean;
}

/** 合并区映射：先合并 range 覆盖的空白单元格，再把整段文本居中填入 */
export interface BlockField {
  range: string;
  fields: string[];
  fmt?: (raw: string) => string;
  /** 超长自动适配：按合并区宽度换行，放不下再缩字号（最小 6pt） */
  fit?: boolean;
}

/** 签名图片：取字段对应的附件图片，缩放到格子内插入 */
export interface ImageField {
  cell: string;
  fields: string[];
  maxW?: number;   // 最大宽（px），默认按格子宽度
  maxH?: number;   // 最大高（px），默认按格子高度
  /** 靠右放置：左侧有标签文字时用，避免图与文字重叠 */
  align?: "left" | "right";
  /** 距单元格左边界的固定偏移（px）：标签在左、图紧跟其后时用 */
  offsetX?: number;
  /** 界面显示名（签名设置里用） */
  label?: string;
  /** 签批状态字段候选链（如「品管经理签批」）。命中 statusOk 才算已签，才填上传的签名 */
  statusField?: string[];
  /** 状态值包含任一关键词即视为已确认（默认见 DEFAULT_STATUS_OK） */
  statusOk?: string[];
  /** 出现这些词一律视为未确认（默认见 DEFAULT_STATUS_NOT） */
  statusNot?: string[];
  /** 上传签名的存储键（每个签名位一个） */
  uploadKey?: string;
}

/**
 * 表头可编辑文本（如「单据编号」）：界面上给一个输入框，打印时写进指定单元格。
 * 输入框留空时回退到 fields 里的表格字段，两者都没有就只写 prefix（标签）。
 */
export interface EditableField {
  key: string;          // 存储键（同一模板内唯一）
  label: string;        // 界面显示名，如「单据编号」
  cell: string;         // 写入的单元格
  prefix?: string;      // 写入时带上的固定标签文字，如「单据编号：」
  defaultValue?: string;// 输入框默认值
  fields?: string[];    // 输入框留空时的回退字段
  /** 文本超出单元格宽度时自动缩小字号（只缩到放得下，正常长度仍是原字号） */
  fit?: boolean;
}

/** 复选框：字段值（经 fmt 归一化后）等于 when 时，把该格第 box 个 □ 涂成 ■ */
export interface CheckField {
  cell: string;
  fields: string[];
  when: string;
  box?: number;   // 第几个 □（从 0 数），默认 0
  fmt?: (raw: string) => string;
  /** 该格本身就是一个画好的方框（没有 □ 字符）：命中时整格填黑 */
  fill?: boolean;
}

/** 签批状态命中这些词即视为已确认 */
export const DEFAULT_STATUS_OK = ["已确认", "确认", "同意", "通过", "允许", "已完成"];

/** 出现这些词一律视为未确认（先于命中词判断，避免「不允许发机」被「允许发机」误判） */
export const DEFAULT_STATUS_NOT = ["不允许", "不同意", "不通过", "否决", "退回", "未确认", "待确认"];

/** 状态文本是否命中"已确认"类关键词（未指定词表时用默认表） */
export function matchStatus(text: string, words?: string[], notWords?: string[]): boolean {
  const t = (text || "").trim();
  if (!t) return false;
  const deny = notWords && notWords.length ? notWords : DEFAULT_STATUS_NOT;
  if (deny.some((w) => t.includes(w))) return false;
  const list = words && words.length ? words : DEFAULT_STATUS_OK;
  return list.some((w) => t.includes(w));
}

export interface SheetDef {
  id: string;
  name: string;
  desc: string;
  file: string;
  /** 结构修正：仅合并单元格、不写入值（如源模板里重复的「签名：」两格合成一格） */
  merges?: string[];
  cells: CellField[];
  blocks: BlockField[];
  checks?: CheckField[];
  images: ImageField[];
  /** 表头可编辑文本（界面上给输入框，打印时写入指定格子） */
  editable?: EditableField[];
  /** 文字位移（单位：字符）：单元格 -> 偏移量，负数向左。仅打印渲染生效，不改模板本身 */
  textShift?: Record<string, number>;
}

// ---------- 模板 1：产品发机评审单（2022版） ----------
// 字段别名沿用「发机评审单打印插件」的 SLOTS 定义，按表单分区排列。
// 同一业务字段只维护一份候选名，后续改别名不用逐格查找。
const REVIEW_ATTACH = ["添附资料", "附件", "附件情况"];
const REVIEW_VERDICT = ["判定结果", "品管部门判定结果", "品管判定"];
const REVIEW_GRADE = ["附条件合格影响等级", "影响度等级", "附条件合格影响度等级", "等级"];
const REVIEW_DIRECTOR_RESULT = ["主管总监评审结果", "总经理监评审结果", "总监评审结果"];
const REVIEW_SALES_RESULT = ["销售代表评审结果", "销售评审结果"];
const REVIEW_BU_RESULT = ["事业部总监评审结果", "事业部评审结果"];

const review: SheetDef = {
  id: "review",
  name: "产品发机评审单",
  desc: "产品中心 · TL/TR-RC008-A0",
  file: reviewXlsx,
  merges: ["F37:G37"],
  // 两个签名标签整体左移 4 个字符，右侧让出签名图的位置
  textShift: { D29: -4, F29: -4 },

  // 表头「单据编号」：源模板写死 C/E-20-，现改为可编辑，默认 TL/TR-RC008-A0。
  // 界面上留空时回退到表格里的「单据编号」字段（保持原来的追加效果）。
  editable: [
    {
      key: "docNo", label: "单据编号", cell: "F2", prefix: "单据编号：",
      defaultValue: "TL/TR-RC008-A0", fields: ["单据编号", "评审单号", "单号"],
    },
  ],

  cells: [
    // ── 基本信息 ──
    { cell: "D5", fields: ["项目号", "FAT发机评审信息", "项目编号"], fit: true },
    { cell: "D6", fields: ["客户名称", "客户"], fit: true },
    { cell: "E6", fields: ["机型", "设备型号", "型号"], fit: true },
    { cell: "F6", fields: ["申请评审部门", "申请部门", "评审部门"], fit: true },
    { cell: "F3", fields: ["申请日期", "评审日期", "日期"], append: true, fmt: fmtDate },
    // 单据编号是表单受控编号（原模板写死 C/E-20-），改由界面输入框填写，见下方 editable

    // ── 品管判定 ──
    { cell: "F23", fields: ["问题对应的主管总监", "主管总监", "品管总监", "总监"], fit: true },
    { cell: "F24", fields: ["销售代表", "销售代表签名", "销售"], fit: true },

    // ── 评审环节：处理意见（签名见 images） ──
    { cell: "C33", fields: ["主管总监评审意见", "总经理监处理意见", "总监评审意见"], append: true, fit: true },
    { cell: "C37", fields: ["销售代表评审意见", "销售评审意见"], append: true, fit: true },
    { cell: "C41", fields: ["事业部总监评审意见", "事业部评审意见"], append: true, fit: true },

    // ── 结果跟踪 ──
    { cell: "G46", fields: ["结案时间", "问题点结案时间", "结案日期"], fmt: fmtDate, fit: true },
    // 整改状态：模板该区无标签格，需要时指定坐标后打开
    // { cell: "C44", fields: ["整改状态", "整改情况"] },
  ],

  // ── 残留问题点 / 判定理由：先合并空白区，再整段填入 ──
  blocks: [
    { range: "C10:D15", fields: ["残留问题点", "关联问题项", "问题点"], fmt: fmtLines, fit: true },
    { range: "E10:E15", fields: ["原因&应对方案", "检测标准&要求", "发生原因说明"], fmt: fmtLines, fit: true },
    { range: "F10:F15", fields: ["预完成日期", "预期处理完成时间"], fmt: fmtDate, fit: true },
    { range: "E20:G22", fields: ["判定理由/说明/处置建议", "判定的理由/说明/处置建议", "判定理由", "处置建议", "理由"], fit: true },
    // { range: "G10:G15", fields: ["主担当部门", "担当部门"] },  // 模板此列未合并，需要时打开
  ],

  // ── 复选框：值命中时把对应 □ 涂成 ■ ──
  checks: [
    // 添附资料（有 □（共 张）；无 □）
    { cell: "E8", fields: REVIEW_ATTACH, when: "有", box: 0, fmt: fmtAttach },
    { cell: "E8", fields: REVIEW_ATTACH, when: "无", box: 1, fmt: fmtAttach },

    // 品管部门判定结果（三个独立框）
    { cell: "C20", fields: REVIEW_VERDICT, when: "合格", fmt: fmtVerdict },
    { cell: "C21", fields: REVIEW_VERDICT, when: "不合格", fmt: fmtVerdict },
    { cell: "C22", fields: REVIEW_VERDICT, when: "附条件合格", fmt: fmtVerdict },

    // 附条件合格影响度等级（同一格内 4 个框：I / II / III / IV）
    { cell: "C24", fields: REVIEW_GRADE, when: "I", box: 0, fmt: fmtGrade },
    { cell: "C24", fields: REVIEW_GRADE, when: "II", box: 1, fmt: fmtGrade },
    { cell: "C24", fields: REVIEW_GRADE, when: "III", box: 2, fmt: fmtGrade },
    { cell: "C24", fields: REVIEW_GRADE, when: "IV", box: 3, fmt: fmtGrade },

    // 主管总监评审结论
    { cell: "C32", fields: REVIEW_DIRECTOR_RESULT, when: "允许发机", box: 0, fmt: fmtReview },
    { cell: "C32", fields: REVIEW_DIRECTOR_RESULT, when: "不允许发机", box: 1, fmt: fmtReview },

    // 销售代表评审结论
    { cell: "C36", fields: REVIEW_SALES_RESULT, when: "允许发机", box: 0, fmt: fmtReview },
    { cell: "C36", fields: REVIEW_SALES_RESULT, when: "不允许发机", box: 1, fmt: fmtReview },

    // 事业部总监评审结论
    { cell: "C40", fields: REVIEW_BU_RESULT, when: "允许发机", box: 0, fmt: fmtReview },
    { cell: "C40", fields: REVIEW_BU_RESULT, when: "不允许发机", box: 1, fmt: fmtReview },
  ],

  // ── 签名：优先用多维表格里的真实签名图；取不到且签批已确认时，用自己上传的签名兜底 ──
  images: [
    { cell: "D16", label: "申请者签名", fields: ["申请者签名", "申请者", "申请人"], uploadKey: "applicant" },
    {
      cell: "F16", label: "申请部门负责人确认", fields: ["申请部门负责人确认", "部门负责人", "负责人确认"],
      statusField: ["部门负责人签批", "部门负责人确认"], uploadKey: "dept_head",
    },
    // 品管经理签名：紧贴「品管经理签名：」标签右侧（标签左移 4 字符后，D 列内还有 113px 空位）
    {
      cell: "D29", label: "品管经理签名", fields: ["品管经理签名", "品管经理"], offsetX: 68,
      statusField: ["品管经理签批", "品管经理确认"], uploadKey: "qc_manager",
    },
    // 品管总监签名（图片字段）：「品管总监签批」是文字状态，手写签名在「品管总监签名」
    {
      cell: "G29", label: "品管总监签名", fields: ["品管总监签名", "品管主管总监签名"],
      statusField: ["品管总监签批", "品管总监确认"], uploadKey: "qc_director",
    },
    {
      cell: "G33", label: "主管总监评审签名", fields: ["主管总监评审签名", "总经理监签名", "总监评审签名"],
      statusField: ["主管总监评审结果"], statusOk: ["允许发机"], uploadKey: "director",
    },
    {
      cell: "G37", label: "销售代表评审签名", fields: ["销售代表评审签名", "销售评审签名"],
      statusField: ["销售代表评审结果"], statusOk: ["允许发机"], uploadKey: "sales",
    },
    {
      cell: "G41", label: "事业部总监评审签名", fields: ["事业部总监评审签名", "事业部评审签名"],
      statusField: ["事业部总监评审结果"], statusOk: ["允许发机"], uploadKey: "bu_director",
    },
    { cell: "G45", label: "记录人签名", fields: ["记录人签名", "问题点跟踪记录人", "记录人"], uploadKey: "recorder" },
  ],
};

// ---------- 模板 2：内部FAT实施申请单 ----------

// 检验结果相关字段的候选名（多维表格里叫什么都能认）
const FAT_RESULT = ["内部FAT检验结果", "内部FAT检验情况", "FAT检验结果", "检验结果", "品管部门检验结果"];
const FAT_COND = ["附条件事项清单", "附条件事项", "附条件清单", "附条件项", "附条件说明"];
const FAT_NO = ["关连整改通知书编号", "关联整改通知书编号", "整改通知书编号", "关连整改通知书No", "关连整改通知书"];
const FAT_FIX = ["整改处理结果", "整改通知书处理结果", "整改结果", "处理结果", "整改处理情况"];
// 申请次数：多维表格里一般叫「申请次数」，模板写作「第    回」
const FAT_TIMES = ["申请次数", "本项目申请回数", "申请回数", "申请批次", "第几次申请"];
// 已确认项目：三行各一个单选字段（完成 / 附条件完成）
const FAT_OK1 = ["验收记录确认", "验收记录", "验收确认"];
const FAT_OK2 = ["工艺调试记录确认", "工艺调试记录", "工艺调试确认"];
const FAT_OK3 = ["瓶样检测合格表确认", "瓶样检测合格表", "瓶样检测确认", "瓶样检测"];
// 发行确认：人员字段
const FAT_PERSON = ["FAT检验人", "FAT检验人员", "检验人", "FAT检验员", "发行确认"];
// 申请人：人员字段（模板上「申请人：」标签在 C23）
const FAT_APPLICANT = ["申请人", "申请者", "申请人姓名", "发起人", "创建人"];
// FAT检开始时间：填 C16「现申请：___月___日___时」
const FAT_START = ["FAT检开始时间", "FAT检验开始时间", "FAT检开始日期", "申请开始时间", "开始时间"];

const fat: SheetDef = {
  id: "fat",
  name: "内部FAT申请表",
  desc: "产品中心内部FAT实施申请单",
  file: fatXlsx,
  // 表头标签是按右对齐设计、超宽时向左溢出到空白的 J 列（Excel 的行为）。
  // 网页渲染会把超宽文字从格子左边向右溢出，压进 L 列与值重叠，所以按实测字宽补偿：
  //   K 列 58px：发行日期：/发行确认：宽 80px，需 -3ch(24px)；申请单编号：宽 96px，需 -5ch(40px)
  //   三行标签右端对齐在 ~620px，值从 L 列的 622px 开始，正好接上，不重叠也不出框。
  textShift: { K1: -3, K2: -5, K3: -3 },

  // 申请单编号：模板里写死的 C/E-22- 前缀不用了，整格重写为多维表格的编号字段；
  // 编号偏长时自动缩小字号，保证不出右边框。
  editable: [
    {
      key: "applyNo", label: "申请单编号", cell: "L2",
      fields: ["申请单编号", "申请单号", "FAT申请编号"], fit: true,
    },
  ],
  cells: [
    { cell: "D5", fields: ["项目号", "FAT发机评审信息"] },
    { cell: "G5", fields: ["客户/机型", "客户名称"] },
    { cell: "L5", fields: FAT_TIMES, insert: ["第", "回"] },
    // 发行日期：标签在 K1，值写在 L1（该格模板为空，append 只为保持左对齐、不改成居中）
    { cell: "L1", fields: ["发行日期"], append: true, fmt: fmtDate },
    // 现申请：___月___日___时 —— 三处空白各填一段，总宽度与模板保持一致
    {
      cell: "C16", fields: FAT_START, dateFmt: "YYYY-MM-DD HH",
      inserts: { slots: [["现申请：", "月"], ["月", "日"], ["日", "时"]], fmt: fmtApplyTime },
    },

    // ── 附以下条件：1~5 行，一行一条（F31:L31 … F35:L35 已合并） ──
    { cell: "F31", fields: FAT_COND, fmt: nthLine(0), wrap: true, align: "left" },
    { cell: "F32", fields: FAT_COND, fmt: nthLine(1), wrap: true, align: "left" },
    { cell: "F33", fields: FAT_COND, fmt: nthLine(2), wrap: true, align: "left" },
    { cell: "F34", fields: FAT_COND, fmt: nthLine(3), wrap: true, align: "left" },
    { cell: "F35", fields: FAT_COND, fmt: nthLine(4), wrap: true, align: "left" },

    // ── 关连整改通知书 No.：附条件合格填上面那行，不合格填下面那行，合格两边都不填 ──
    {
      cell: "H29", fields: FAT_NO, append: true,
      whenField: FAT_RESULT, whenValues: ["附条件合格"],
    },
    {
      cell: "H37", fields: FAT_NO, append: true,
      whenField: FAT_RESULT, whenValues: ["不合格"],
    },

    // ── 整改处理结果：接在「整改通知书处理及附条件事项处理的结果：」标签后面 ──
    { cell: "C40", fields: FAT_FIX, append: true },

    // ── 发行确认：填 FAT 检验人姓名（标签在 K3，值写 L3） ──
    { cell: "L3", fields: FAT_PERSON },

    // ── 申请人：接在「申请人：」标签后面（C23，标签宽 64px 已溢出，追加后仍留 60px 余量到 F23） ──
    { cell: "C23", fields: FAT_APPLICANT, append: true },
  ],

  // ── 检验结果三选一：C27/C29/C37 本身就是画好的方框，命中时整格涂黑 ──
  checks: [
    { cell: "C27", fields: FAT_RESULT, when: "合格", fmt: fmtVerdict, fill: true },
    { cell: "C29", fields: FAT_RESULT, when: "附条件合格", fmt: fmtVerdict, fill: true },
    { cell: "C37", fields: FAT_RESULT, when: "不合格", fmt: fmtVerdict, fill: true },

    // 已确认三项：每项「完成 / 附条件完成」各一个方框（H 列），命中就涂黑
    { cell: "H9", fields: FAT_OK1, when: "完成", fmt: fmtDone, fill: true },
    { cell: "H10", fields: FAT_OK1, when: "附条件完成", fmt: fmtDone, fill: true },
    { cell: "H11", fields: FAT_OK2, when: "完成", fmt: fmtDone, fill: true },
    { cell: "H12", fields: FAT_OK2, when: "附条件完成", fmt: fmtDone, fill: true },
    { cell: "H13", fields: FAT_OK3, when: "完成", fmt: fmtDone, fill: true },
    { cell: "H14", fields: FAT_OK3, when: "附条件完成", fmt: fmtDone, fill: true },
  ],
  blocks: [
    { range: "C20:H22", fields: ["缺件及未达事项", "关联问题项", "问题点"], fmt: fmtLines },
    { range: "I20:K22", fields: ["发生原因说明", "原因&应对方案"] },
    { range: "L20:M22", fields: ["预期处理完成时间", "预完成日期"], fmt: fmtDate },
  ],
  images: [
    // 申请人的手签图：姓名已打印在「申请人：」后面（C23，占 33~145px），
    // 签名图顺延到 E23（149px 起），避免盖住打印出来的姓名。
    { cell: "E23", fields: ["申请者签名", "申请人签名"] },
    { cell: "G23", fields: ["制造部经理", "制造部经理签名"] },
    { cell: "H43", fields: ["确认签名", "最终合格确认签名"] },
  ],
};

/** 界面上的模板列表：加模板 = 往这里加一项 */
export const SHEETS: SheetDef[] = [review, fat];

// ===== 数据表名 -> 模板 =====
// 多维表格里切到哪张数据表，插件就只显示对应的那一个模板。
// 按「评审」/「申请」区分，不能只看 FAT：评审表也可能叫「FAT发机评审单」。
// 规则自上而下命中即停；都不命中时显示全部模板，避免新表名没法用。
// 以后新增模板/新表名，在这里加一条规则即可。
const TABLE_RULES: { test: RegExp; sheetId: string }[] = [
  { test: /评审/, sheetId: "review" },      // FAT发机评审单 / 产品发机评审单 / 发机评审信息
  { test: /申请/, sheetId: "fat" },         // 内部FAT实施申请单 / FAT申请单 / FAT申请表
  { test: /fat/i, sheetId: "fat" },         // 兜底：表名里只写了 FAT
];

/** 按当前数据表名筛选可用模板 */
export function sheetsForTable(tableName: string): SheetDef[] {
  const name = (tableName || "").trim();
  if (!name) return SHEETS;
  for (const r of TABLE_RULES) {
    if (!r.test.test(name)) continue;
    const s = SHEETS.find((x) => x.id === r.sheetId);
    if (s) return [s];
  }
  return SHEETS;
}
