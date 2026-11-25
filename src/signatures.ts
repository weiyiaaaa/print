// 上传签名管理：把签好的手写签名图存到浏览器本地（localStorage），打印时按签批状态自动填入
const KEY = "print-plugin-signatures:v1";

export type SavedSignatures = Record<string, string>; // uploadKey -> dataURL

export function loadSignatures(): SavedSignatures {
  try {
    const raw = localStorage.getItem(KEY);
    const obj = raw ? JSON.parse(raw) : {};
    return obj && typeof obj === "object" ? obj : {};
  } catch {
    return {};
  }
}

function writeSignatures(signatures: SavedSignatures): void {
  try {
    localStorage.setItem(KEY, JSON.stringify(signatures));
  } catch { /* 隐私模式或配额不足时忽略，与原有行为一致 */ }
}

export function saveSignature(key: string, dataUrl: string): void {
  writeSignatures({ ...loadSignatures(), [key]: dataUrl });
}

export function removeSignature(key: string): void {
  const signatures = loadSignatures();
  delete signatures[key];
  writeSignatures(signatures);
}

/** dataURL -> 二进制（写进 xlsx 用） */
export function dataUrlToBytes(dataUrl: string): Uint8Array | null {
  try {
    const i = dataUrl.indexOf(",");
    if (i < 0) return null;
    const bin = atob(dataUrl.slice(i + 1));
    const out = new Uint8Array(bin.length);
    for (let k = 0; k < bin.length; k++) out[k] = bin.charCodeAt(k);
    return out.length ? out : null;
  } catch {
    return null;
  }
}

/**
 * 上传的图片文件 -> 压缩后的 dataURL。
 * 签名图不需要很大：默认限制 420×220，保留透明底（PNG）。
 */
export async function fileToSignatureDataUrl(file: File, maxW = 420, maxH = 220): Promise<string> {
  const bitmap = await createImageBitmap(file);
  const scale = Math.min(1, maxW / bitmap.width, maxH / bitmap.height);
  const w = Math.max(1, Math.round(bitmap.width * scale));
  const h = Math.max(1, Math.round(bitmap.height * scale));
  const canvas = document.createElement("canvas");
  canvas.width = w;
  canvas.height = h;
  const ctx = canvas.getContext("2d");
  if (!ctx) throw new Error("浏览器不支持画布，无法处理签名图");
  ctx.drawImage(bitmap, 0, 0, w, h);
  bitmap.close?.();
  return canvas.toDataURL("image/png");
}
