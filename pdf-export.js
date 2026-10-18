// 提出用 PDF の一括作成（1人1つのPDFをフォルダに入れてzipで保存）
// Excelと同じ様式（A4縦1枚）を画像として描き、PDFに貼り付けます。外部ライブラリは使いません。
import { WD, pad, ymParts, daysIn, wday, ymLabel, toMin, fmt, workMin, summarize } from "./calc.js";
import { zip } from "./xlsx-writer.js";

const DPI = 200;
const PX = mm => Math.round(mm / 25.4 * DPI);          // ミリ → ピクセル
const W = PX(210), H = PX(297);                         // A4
const FONT = '"MS PGothic","Hiragino Sans","Hiragino Kaku Gothic ProN","Yu Gothic","Meiryo","Noto Sans JP",sans-serif';
const LINE = { hair: 1.5, thin: 2.5, medium: 5 };

function newCanvas(){
  const c = document.createElement("canvas"); c.width = W; c.height = H;
  const g = c.getContext("2d");
  g.fillStyle = "#fff"; g.fillRect(0, 0, W, H);
  g.fillStyle = "#000"; g.strokeStyle = "#000"; g.textBaseline = "middle";
  return { c, g };
}
function line(g, x1, y1, x2, y2, w){ if (!w) return; g.lineWidth = LINE[w]; g.beginPath(); g.moveTo(x1, y1); g.lineTo(x2, y2); g.stroke(); }
function text(g, s, x, y, w, { size = 4.2, bold = false, align = "center" } = {}){
  if (s == null || s === "") return;
  g.font = `${bold ? "bold " : ""}${PX(size)}px ${FONT}`;
  g.textAlign = align;
  const tx = align === "center" ? x + w / 2 : align === "left" ? x + PX(1.5) : x + w - PX(1.5);
  g.fillText(String(s), tx, y, w - PX(2));
}
// 枠：b = {l,r,t,b} の線の太さ
function box(g, x, y, w, h, b = {}, fill){
  if (fill){ g.fillStyle = fill; g.fillRect(x, y, w, h); g.fillStyle = "#000"; }
  line(g, x, y, x + w, y, b.t); line(g, x, y + h, x + w, y + h, b.b);
  line(g, x, y, x, y + h, b.l); line(g, x + w, y, x + w, y + h, b.r);
}

/* ---------- 勤務表1枚 ---------- */
export function drawTimecard(p, m, ym){
  const { c, g } = newCanvas();
  const { y, m: mo } = ymParts(ym), ds = (m && m.days) || {}, n = daysIn(ym);
  const ratio = [7, 7, 14, 14, 14, 14], tw = PX(160), x0 = Math.round((W - tw) / 2);
  const cw = ratio.map(r => tw * r / 70), cx = [x0]; cw.forEach((w, i) => cx.push(cx[i] + w));
  let top = PX(16);

  text(g, `${y}年${mo}月`, x0, top + PX(5), tw, { size: 6, bold: true });
  top += PX(14);

  // 社員番号・名前
  const half = cx[3] - cx[0], half2 = cx[6] - cx[3], lh = PX(6), vh = PX(7.5);
  box(g, cx[0], top, half, lh, { l: "thin", t: "thin", b: "hair" }, "#d8d8d8");
  box(g, cx[3], top, half2, lh, { l: "thin", r: "thin", t: "thin", b: "hair" }, "#d8d8d8");
  text(g, "社員番号", cx[0], top + lh / 2, half); text(g, "名前", cx[3], top + lh / 2, half2);
  box(g, cx[0], top + lh, half, vh, { l: "thin", b: "thin" }); box(g, cx[3], top + lh, half2, vh, { l: "thin", r: "thin", b: "thin" });
  text(g, p.no || "", cx[0], top + lh + vh / 2, half, { size: 4.6 }); text(g, p.name, cx[3], top + lh + vh / 2, half2, { size: 4.6 });
  top += lh + vh + PX(6);

  // 見出し（2段）
  const hh = PX(6);
  box(g, cx[0], top, cw[0], hh * 2, { l: "medium", r: "hair", t: "medium", b: "medium" }); text(g, "日", cx[0], top + hh, cw[0]);
  box(g, cx[1], top, cw[1], hh * 2, { r: "hair", t: "medium", b: "medium" }); text(g, "曜", cx[1], top + hh, cw[1]);
  box(g, cx[2], top, cx[6] - cx[2], hh, { r: "medium", t: "medium", b: "hair" }); text(g, "タイムカード", cx[2], top + hh / 2, cx[6] - cx[2]);
  ["入", "退", "休憩", "勤務時間"].forEach((t, i) => {
    box(g, cx[2 + i], top + hh, cw[2 + i], hh, { r: i === 3 ? "medium" : "hair", b: "medium" });
    text(g, t, cx[2 + i], top + hh * 1.5, cw[2 + i]);
  });
  top += hh * 2;

  // 31行
  const rh = PX(6.3); let total = 0;
  for (let i = 0; i < 31; i++){
    const d = i + 1, inMonth = d <= n, day = inMonth ? (ds[pad(d)] || {}) : {}, yy = top + rh * i, mid = yy + rh / 2;
    const w = workMin(day); if (w != null) total += w;
    const vals = [inMonth ? d : "", inMonth ? WD[wday(ym, d)] : "", day.in ? fmt(toMin(day.in)) : "", day.out ? fmt(toMin(day.out)) : "", day.brk ? fmt(toMin(day.brk)) : "", w == null ? "" : fmt(w)];
    for (let k = 0; k < 6; k++){
      box(g, cx[k], yy, cw[k], rh, { l: k === 0 ? "medium" : null, r: k === 5 ? "medium" : "hair", b: i === 30 ? "medium" : "hair" });
      text(g, vals[k], cx[k], mid, cw[k]);
    }
  }
  top += rh * 31;

  // 合計
  const th = PX(7);
  box(g, cx[0], top, cx[5] - cx[0], th, { l: "medium", r: "thin", b: "medium" }); text(g, "合計", cx[0], top + th / 2, cx[5] - cx[0], { bold: true });
  box(g, cx[5], top, cw[5], th, { r: "medium", b: "medium" }); text(g, fmt(total), cx[5], top + th / 2, cw[5], { bold: true });
  return c;
}

/* ---------- 集計1枚 ---------- */
export function drawSummary(people, months, ym){
  const { c, g } = newCanvas();
  const ratio = [12, 18, 10, 14], tw = PX(150), x0 = Math.round((W - tw) / 2);
  const cw = ratio.map(r => tw * r / 54), cx = [x0]; cw.forEach((w, i) => cx.push(cx[i] + w));
  let top = PX(16);
  text(g, ymLabel(ym) + "　勤務時間集計", x0, top + PX(4), tw, { size: 5.5, bold: true, align: "left" });
  top += PX(12);
  const rh = PX(6.5), all = { l: "thin", r: "thin", t: "thin", b: "thin" };
  ["社員番号", "名前", "出勤日数", "勤務時間合計"].forEach((t, k) => { box(g, cx[k], top, cw[k], rh, all, "#d8d8d8"); text(g, t, cx[k], top + rh / 2, cw[k]); });
  top += rh;
  let sd = 0, sm = 0;
  for (const p of people){
    if (top > H - PX(30)) break;        // 1枚に収まる範囲（約35人）まで
    const s = summarize(months[p.id]); sd += s.days; sm += s.total;
    const v = [p.no || "", p.name, s.days, fmt(s.total)];
    v.forEach((t, k) => { box(g, cx[k], top, cw[k], rh, all); text(g, t, cx[k], top + rh / 2, cw[k], { align: k === 1 ? "left" : "center" }); });
    top += rh;
  }
  box(g, cx[0], top, cx[2] - cx[0], rh, all); text(g, "合計", cx[0], top + rh / 2, cx[2] - cx[0], { bold: true });
  box(g, cx[2], top, cw[2], rh, all); text(g, sd, cx[2], top + rh / 2, cw[2], { bold: true });
  box(g, cx[3], top, cw[3], rh, all); text(g, fmt(sm), cx[3], top + rh / 2, cw[3], { bold: true });
  return c;
}

/* ---------- 画像 → PDF ---------- */
async function deflate(bytes){
  const cs = new CompressionStream("deflate");                 // zlib形式（PDFの FlateDecode と同じ）
  const out = new Response(new Blob([bytes]).stream().pipeThrough(cs));
  return new Uint8Array(await out.arrayBuffer());
}
async function canvasImage(c){
  if (typeof CompressionStream !== "undefined"){
    const px = c.getContext("2d").getImageData(0, 0, c.width, c.height).data;
    const gray = new Uint8Array(c.width * c.height);
    for (let i = 0, j = 0; j < gray.length; i += 4, j++) gray[j] = px[i];   // 白黒（グレー）で保存
    return { data: await deflate(gray), filter: "FlateDecode", cs: "DeviceGray" };
  }
  const blob = await new Promise(r => c.toBlob(r, "image/jpeg", 0.92));    // 古いブラウザ向け
  return { data: new Uint8Array(await blob.arrayBuffer()), filter: "DCTDecode", cs: "DeviceRGB" };
}
export async function canvasToPdf(c, title){
  const img = await canvasImage(c);
  const enc = new TextEncoder(), chunks = [], offs = [];
  let len = 0;
  const push = x => { const b = typeof x === "string" ? enc.encode(x) : x; chunks.push(b); len += b.length; };
  const obj = (n, body) => { offs[n] = len; push(`${n} 0 obj\n`); body(); push("\nendobj\n"); };
  const content = `q 595.28 0 0 841.89 0 0 cm /Im0 Do Q`;
  const hexTitle = "FEFF" + Array.from(title || "").map(ch => ch.charCodeAt(0).toString(16).padStart(4, "0")).join("").toUpperCase();
  push("%PDF-1.4\n%\xE2\xE3\xCF\xD3\n");
  obj(1, () => push("<< /Type /Catalog /Pages 2 0 R >>"));
  obj(2, () => push("<< /Type /Pages /Kids [3 0 R] /Count 1 >>"));
  obj(3, () => push("<< /Type /Page /Parent 2 0 R /MediaBox [0 0 595.28 841.89] /Resources << /XObject << /Im0 4 0 R >> >> /Contents 5 0 R >>"));
  obj(4, () => { push(`<< /Type /XObject /Subtype /Image /Width ${c.width} /Height ${c.height} /ColorSpace /${img.cs} /BitsPerComponent 8 /Filter /${img.filter} /Length ${img.data.length} >>\nstream\n`); push(img.data); push("\nendstream"); });
  obj(5, () => push(`<< /Length ${content.length} >>\nstream\n${content}\nendstream`));
  obj(6, () => push(`<< /Title <${hexTitle}> /Producer (kintai) >>`));
  const xref = len;
  push(`xref\n0 7\n0000000000 65535 f \n${offs.slice(1).map(o => String(o).padStart(10, "0") + " 00000 n \n").join("")}`);
  push(`trailer\n<< /Size 7 /Root 1 0 R /Info 6 0 R >>\nstartxref\n${xref}\n%%EOF\n`);
  const out = new Uint8Array(len); let p = 0; for (const ch of chunks){ out.set(ch, p); p += ch.length; }
  return out;
}

/* ---------- 一括ダウンロード ---------- */
const safe = s => String(s).replace(/[\\/:*?"<>|]/g, "").trim() || "従業員";

export async function buildPdfZip(people, months, ym, onProgress){
  const { y, m } = ymParts(ym), folder = `出勤時間_${y}年${m}月`;
  const files = {}, used = new Set();
  for (let i = 0; i < people.length; i++){
    const p = people[i];
    let base = safe((p.no ? p.no + "_" : "") + p.name), name = base, k = 2;
    while (used.has(name)) name = `${base}(${k++})`;
    used.add(name);
    files[`${folder}/${name}.pdf`] = await canvasToPdf(drawTimecard(p, months[p.id], ym), `${ymLabel(ym)} ${p.name}`);
    if (onProgress) onProgress(i + 1, people.length);
    await new Promise(r => setTimeout(r, 0));             // 画面が固まらないよう1人ずつ区切る
  }
  if (people.length > 1) files[`${folder}/集計.pdf`] = await canvasToPdf(drawSummary(people, months, ym), `${ymLabel(ym)} 集計`);
  return { bytes: zip(files), filename: `${folder}_PDF.zip` };
}

export async function downloadPdfZip(people, months, ym, onProgress){
  const { bytes, filename } = await buildPdfZip(people, months, ym, onProgress);
  const url = URL.createObjectURL(new Blob([bytes], { type: "application/zip" }));
  const a = document.createElement("a"); a.href = url; a.download = filename;
  document.body.appendChild(a); a.click(); a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 5000);
}
