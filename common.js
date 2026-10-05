// 出勤時間申請 共通処理（Firebase初期化・勤務時間の計算・勤務表の表示・Excel出力）
import { initializeApp } from "https://www.gstatic.com/firebasejs/10.12.2/firebase-app.js";
import { getFirestore, doc, getDoc, setDoc, updateDoc } from "https://www.gstatic.com/firebasejs/10.12.2/firebase-firestore.js";
import { firebaseConfig } from "./firebase-config.js";

export const configured = !String(firebaseConfig.apiKey).startsWith("YOUR_");
export const fbApp = configured ? initializeApp(firebaseConfig) : null;
export const db = configured ? getFirestore(fbApp) : null;

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

/* ---------- 画面の小物 ---------- */
export function esc(s){ return String(s ?? "").replace(/[&<>"']/g, c => ({ "&":"&amp;", "<":"&lt;", ">":"&gt;", '"':"&quot;", "'":"&#39;" }[c])); }
export function toast(msg){
  const t = document.getElementById("toast"); if (!t) return;
  t.textContent = msg; t.hidden = false; clearTimeout(toast._t); toast._t = setTimeout(() => t.hidden = true, 2200);
}
export function statusPill(m){
  if (!m || !m.days || !Object.keys(m.days).length) return '<span class="pill none">未入力</span>';
  return m.status === "submitted" ? '<span class="pill sub">申請済</span>' : '<span class="pill draft">入力中</span>';
}

/* ---------- 勤務表（従業員・管理者 共通） ---------- */
export function buildSheet(host, empId, ym, asAdmin){
  const n = daysIn(ym);
  const opts = ['<option value="">—</option>'];
  for (let m = 0; m <= 180; m += 15){ const v = Math.floor(m / 60) + ":" + pad(m % 60); opts.push(`<option value="${v}">${v}</option>`); }
  let rows = "";
  for (let d = 1; d <= n; d++){
    const w = wday(ym, d), dd = pad(d);
    rows += `<tr data-d="${dd}" class="${w === 0 ? "sun" : w === 6 ? "sat" : ""}">
      <td class="d">${d}</td><td class="w">${WD[w]}</td>
      <td><input type="time" id="t-${dd}-in" data-f="in" aria-label="${d}日 入"></td>
      <td><input type="time" id="t-${dd}-out" data-f="out" aria-label="${d}日 退"></td>
      <td><select id="t-${dd}-brk" data-f="brk" aria-label="${d}日 休憩">${opts.join("")}</select></td>
      <td class="h"></td></tr>`;
  }
  host.innerHTML = `<div class="tablebox"><table>
    <thead><tr class="grp"><th rowspan="2">日</th><th rowspan="2">曜</th><th colspan="4">タイムカード</th></tr>
    <tr><th>入</th><th>退</th><th>休憩</th><th>勤務時間</th></tr></thead>
    <tbody>${rows}</tbody>
    <tfoot><tr><td colspan="5" style="text-align:right">合計</td><td class="h tot"></td></tr></tfoot>
  </table></div>`;
  host.onchange = e => {
    const el = e.target; if (!el.dataset.f) return;
    const tr = el.closest("tr");
    const day = {}; tr.querySelectorAll("[data-f]").forEach(x => day[x.dataset.f] = x.value);
    const w = workMin(day); tr.querySelector("td.h").textContent = w == null ? "" : fmt(w);
    saveDay(empId, ym, tr.dataset.d, { [el.dataset.f]: el.value }, asAdmin);
  };
}
export function fillSheet(host, m, locked, asAdmin){
  const ds = (m && m.days) || {};
  host.querySelectorAll("tbody tr").forEach(tr => {
    const day = ds[tr.dataset.d] || {};
    tr.querySelectorAll("[data-f]").forEach(el => {
      if (document.activeElement !== el) el.value = day[el.dataset.f] || "";
      el.disabled = locked;
    });
    const w = workMin(day), td = tr.querySelector("td.h");
    td.textContent = w == null ? "" : fmt(w); td.classList.toggle("zero", w === 0);
    tr.classList.toggle("adm", !!day.adm);
    let o = tr.querySelector(".orig");
    if (asAdmin && day.adm && day.orig){
      if (!o){ o = document.createElement("span"); o.className = "orig"; tr.children[2].appendChild(o); }
      o.textContent = `申請時 ${day.orig.in || "—"}〜${day.orig.out || "—"} 休${day.orig.brk || "—"}`;
    } else if (o) o.remove();
  });
  host.querySelector(".tot").textContent = fmt(summarize(m).total) || "0:00";
}

const queues = {};
export function saveDay(empId, ym, d, patch, asAdmin){
  const key = empId + "_" + ym;
  queues[key] = (queues[key] || Promise.resolve()).then(async () => {
    const ref = doc(db, "months", key);
    try {
      const snap = await getDoc(ref);
      const cur = snap.exists() ? snap.data() : null;
      const prev = (cur && cur.days && cur.days[d]) || {};
      const next = { ...prev, ...patch };
      if (asAdmin && !prev.adm){ next.adm = true; next.orig = { in: prev.in || "", out: prev.out || "", brk: prev.brk || "" }; }
      const now = new Date().toISOString();
      if (cur) await updateDoc(ref, { ["days." + d]: next, updatedAt: now });
      else await setDoc(ref, { empId, ym, status: "draft", days: { [d]: next }, updatedAt: now });
      toast(asAdmin ? "修正を保存しました" : "保存しました");
    } catch (e) {
      console.error(e);
      toast(e && e.code === "permission-denied" ? "保存できません（申請済み、または権限がありません）" : "保存できませんでした。通信を確認してもう一度入力してください");
    }
  });
  return queues[key];
}
export async function setStatus(empId, ym, status){
  const patch = { status, updatedAt: new Date().toISOString() };
  if (status === "submitted") patch.submittedAt = patch.updatedAt;
  await updateDoc(doc(db, "months", empId + "_" + ym), patch);
}

/* ---------- Excel 出力（見本のA〜F列と同じ並び） ---------- */
function serial(ym, d){ const { y, m } = ymParts(ym); return (Date.UTC(y, m - 1, d) - Date.UTC(1899, 11, 30)) / 86400000; }
function tcell(t){ const m = toMin(t); return m == null ? null : { t: "n", v: m / 1440, z: "h:mm" }; }
function sheetFor(p, m, ym){
  const { y, m: mo } = ymParts(ym), ds = (m && m.days) || {};
  const aoa = [[y, "年", mo, "月"], [], ["社員番号", p.no || "", "名前", p.name], [], ["日", "曜", "入", "退", "休憩", "勤務時間"]];
  let total = 0;
  for (let d = 1; d <= daysIn(ym); d++){
    const day = ds[pad(d)] || {}, w = workMin(day); if (w != null) total += w;
    aoa.push([{ t: "n", v: serial(ym, d), z: "d" }, WD[wday(ym, d)], tcell(day.in), tcell(day.out), tcell(day.brk), w == null ? null : { t: "n", v: w / 1440, z: "[h]:mm" }]);
  }
  aoa.push(["合計", "", "", "", "", { t: "n", v: total / 1440, z: "[h]:mm" }]);
  const ws = XLSX.utils.aoa_to_sheet(aoa);
  ws["!cols"] = [{ wch: 7 }, { wch: 6 }, { wch: 9 }, { wch: 9 }, { wch: 9 }, { wch: 11 }];
  ws["!merges"] = [{ s: { r: aoa.length - 1, c: 0 }, e: { r: aoa.length - 1, c: 4 } }];
  return ws;
}
function sheetName(s, used){
  const base = String(s).replace(/[\[\]:*?\/\\]/g, "").slice(0, 31) || "従業員";
  let k = base, i = 2; while (used.has(k)) k = base.slice(0, 28) + "(" + (i++) + ")";
  used.add(k); return k;
}
// people: [{id,no,name}]  months: {empId: monthDoc}
export function downloadExcel(people, months, ym){
  if (typeof XLSX === "undefined"){ toast("Excel出力の準備ができていません。ページを再読み込みしてください"); return; }
  const wb = XLSX.utils.book_new(), used = new Set(["集計"]);
  if (people.length > 1){
    const rows = [["年月", ymLabel(ym)], [], ["社員番号", "名前", "出勤日数", "勤務時間合計", "勤務時間(時間数)", "状態", "管理者修正(日)"]];
    people.forEach(p => {
      const m = months[p.id], s = summarize(m);
      rows.push([p.no || "", p.name, s.days, { t: "n", v: s.total / 1440, z: "[h]:mm" }, Math.round(s.total / 60 * 100) / 100,
        !s.days ? "未入力" : m.status === "submitted" ? "申請済" : "入力中", s.edits]);
    });
    const ws = XLSX.utils.aoa_to_sheet(rows);
    ws["!cols"] = [{ wch: 10 }, { wch: 14 }, { wch: 9 }, { wch: 12 }, { wch: 14 }, { wch: 8 }, { wch: 14 }];
    XLSX.utils.book_append_sheet(wb, ws, "集計");
  }
  people.forEach(p => XLSX.utils.book_append_sheet(wb, sheetFor(p, months[p.id], ym), sheetName(p.name, used)));
  const { y, m } = ymParts(ym);
  XLSX.writeFile(wb, people.length > 1 ? `出勤時間_${y}年${m}月.xlsx` : `出勤時間_${y}年${m}月_${people[0].name}.xlsx`);
}
