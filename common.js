// 出勤時間申請 共通処理（Firebase初期化・勤務表の表示・保存）
import { initializeApp } from "https://www.gstatic.com/firebasejs/10.12.2/firebase-app.js";
import { getFirestore, doc, getDoc, setDoc, updateDoc } from "https://www.gstatic.com/firebasejs/10.12.2/firebase-firestore.js";
import { firebaseConfig } from "./firebase-config.js";

export const configured = !String(firebaseConfig.apiKey).startsWith("YOUR_");
export const fbApp = configured ? initializeApp(firebaseConfig) : null;
export const db = configured ? getFirestore(fbApp) : null;

import { WD, pad, ymParts, daysIn, wday, ymLabel, toMin, fmt, workMin, summarize } from "./calc.js";
export * from "./calc.js";
export { downloadExcel } from "./excel.js";

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
