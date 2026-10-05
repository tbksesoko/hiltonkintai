// 日付と勤務時間の計算（画面・Excel共通）
export const WD = "日月火水木金土";

/* ---------- 日付・時間 ---------- */
export function pad(n){ return String(n).padStart(2, "0"); }
export function thisYM(){ const d = new Date(); return d.getFullYear() + "-" + pad(d.getMonth() + 1); }
export function ymParts(ym){ const [y, m] = ym.split("-").map(Number); return { y, m }; }
export function shiftYM(ym, k){ let { y, m } = ymParts(ym); m += k; while (m < 1){ m += 12; y--; } while (m > 12){ m -= 12; y++; } return y + "-" + pad(m); }
export function daysIn(ym){ const { y, m } = ymParts(ym); return new Date(y, m, 0).getDate(); }
export function wday(ym, d){ const { y, m } = ymParts(ym); return new Date(y, m - 1, d).getDay(); }
export function ymLabel(ym){ const { y, m } = ymParts(ym); return y + "年" + m + "月"; }
export function toMin(t){ if (!t) return null; const [h, m] = String(t).split(":").map(Number); return isNaN(h) ? null : h * 60 + (m || 0); }
export function fmt(min){ if (min == null) return ""; const s = min < 0 ? "-" : ""; min = Math.abs(min); return s + Math.floor(min / 60) + ":" + pad(min % 60); }

// エクセル見本と同じ計算：入は15分単位で切り上げ、退は15分単位で切り捨て、休憩を差し引く
export function workMin(day){
  if (!day) return null;
  const a = toMin(day.in), b0 = toMin(day.out);
  if (a == null || b0 == null) return null;
  const b = b0 <= a ? b0 + 1440 : b0;            // 退が入より早い＝翌日の退勤
  const inR = Math.ceil(a / 15) * 15, outR = Math.floor(b / 15) * 15;
  return Math.max(0, outR - inR - (toMin(day.brk) || 0));
}
export function summarize(m){
  let total = 0, days = 0, edits = 0;
  const ds = (m && m.days) || {};
  for (const k in ds){ const w = workMin(ds[k]); if (w != null){ total += w; days++; } if (ds[k].adm) edits++; }
  return { total, days, edits };
}
