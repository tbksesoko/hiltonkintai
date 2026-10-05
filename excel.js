// 提出用 Excel の作成（エクセル見本のA〜F列の様式・A4縦1枚に印刷）
import { WD, pad, ymParts, daysIn, wday, ymLabel, toMin, workMin, summarize } from "./calc.js";
import { Workbook } from "./xlsx-writer.js";

const FIRST = 6, LAST = 36, TOTAL = 37;        // データ行（0始まり）：7〜37行目、合計は38行目
const GRAY = "D8D8D8";

function serial(ym, d){ const { y, m } = ymParts(ym); return (Date.UTC(y, m - 1, d) - Date.UTC(1899, 11, 30)) / 86400000; }

// 表（5〜38行目）の罫線：外枠は太線、内側は細線（見本と同じ）
function grid(r, c){
  return {
    l: c === 0 ? "medium" : "hair",
    r: c === 5 ? "medium" : "hair",
    t: r === 4 || r === FIRST || r === TOTAL ? "medium" : "hair",
    b: r === 5 || r === LAST || r === TOTAL ? "medium" : "hair",
  };
}

function timecardSheet(wb, name, p, m, ym){
  const ws = wb.addSheet(name);
  const { y, m: mo } = ymParts(ym), ds = (m && m.days) || {}, n = daysIn(ym);
  const st = spec => wb.style(spec);
  ws.cols = [7, 7, 14, 14, 14, 14];

  // 1行目：年月
  ws.set(0, 0, `${y}年${mo}月`, st({ h: "center", size: 14, bold: true }));
  for (let c = 1; c < 6; c++) ws.set(0, c, null, st({}));
  ws.merge(0, 0, 0, 5);
  ws.rowH[0] = 24;

  // 3〜4行目：社員番号・名前
  const box = (b, h) => st({ h, fill: b.fill, border: { l: b.l, r: b.r, t: b.t, b: b.b } });
  for (let c = 0; c < 6; c++){
    const l = c === 0 || c === 3 ? "thin" : null, r = c === 2 || c === 5 ? "thin" : null;
    ws.set(2, c, c === 0 ? "社員番号" : c === 3 ? "名前" : null, box({ fill: GRAY, l, r, t: "thin", b: "hair" }, "center"));
    ws.set(3, c, c === 0 ? (p.no || "") : c === 3 ? p.name : null, box({ l, r, t: "hair", b: "thin" }, "center"));
  }
  ws.merge(2, 0, 2, 2); ws.merge(2, 3, 2, 5); ws.merge(3, 0, 3, 2); ws.merge(3, 3, 3, 5);
  ws.rowH[2] = 16; ws.rowH[3] = 20;

  // 5〜6行目：見出し
  const hd = (r, c, v) => ws.set(r, c, v, st({ h: "center", border: grid(r, c) }));
  hd(4, 0, "日"); hd(5, 0, null); ws.merge(4, 0, 5, 0);
  hd(4, 1, "曜"); hd(5, 1, null); ws.merge(4, 1, 5, 1);
  hd(4, 2, "タイムカード"); hd(4, 3, null); hd(4, 4, null); hd(4, 5, null); ws.merge(4, 2, 4, 5);
  ["入", "退", "休憩", "勤務時間"].forEach((t, i) => hd(5, 2 + i, t));
  ws.rowH[4] = 16; ws.rowH[5] = 16;

  // 7〜37行目：日ごと（月の日数に関係なく31行、見本と同じ）
  let total = 0;
  for (let i = 0; i < 31; i++){
    const r = FIRST + i, d = i + 1, inMonth = d <= n, day = inMonth ? (ds[pad(d)] || {}) : {};
    const R = r + 1;   // Excelの行番号
    ws.set(r, 0, inMonth ? serial(ym, d) : null, st({ h: "center", fmt: "d", border: grid(r, 0) }));
    ws.set(r, 1, inMonth ? WD[wday(ym, d)] : null, st({ h: "center", border: grid(r, 1) }));
    [day.in, day.out, day.brk].forEach((t, k) => {
      const v = toMin(t);
      ws.set(r, 2 + k, v == null ? null : v / 1440, st({ h: "center", fmt: "h:mm", border: grid(r, 2 + k) }));
    });
    const w = workMin(day); if (w != null) total += w;
    // 勤務時間は数式（入は15分切り上げ・退は15分切り捨て・休憩を引く）。Excelで時間を直しても再計算されます
    const f = `IF(OR(C${R}="",D${R}=""),"",MAX(0,FLOOR(ROUND((D${R}+(D${R}<=C${R}))*1440,0),15)-CEILING(ROUND(C${R}*1440,0),15)-ROUND(N(E${R})*1440,0))/1440)`;
    ws.set(r, 5, w == null ? "" : w / 1440, st({ h: "center", fmt: "[h]:mm", border: grid(r, 5) }), f);
    ws.rowH[r] = 18;
  }

  // 38行目：合計
  ws.set(TOTAL, 0, "合計", st({ h: "center", bold: true, border: grid(TOTAL, 0) }));
  for (let c = 1; c < 5; c++) ws.set(TOTAL, c, null, st({ border: grid(TOTAL, c) }));
  ws.merge(TOTAL, 0, TOTAL, 4);
  ws.set(TOTAL, 5, total / 1440, st({ h: "center", bold: true, fmt: "[h]:mm", border: grid(TOTAL, 5) }), `SUM(F${FIRST + 1}:F${LAST + 1})`);
  ws.rowH[TOTAL] = 20;

  ws.printArea = "$A$1:$F$38";
  ws.page = { fit: true, center: true, margins: { l: 0.71, r: 0.71, t: 0.75, b: 0.75 } };
  return ws;
}

function summarySheet(wb, people, months, ym){
  const ws = wb.addSheet("集計");
  const st = spec => wb.style(spec);
  const cell = (r, c, v, extra = {}) => ws.set(r, c, v, st({ border: { l: "thin", r: "thin", t: "thin", b: "thin" }, ...extra }));
  ws.cols = [12, 18, 10, 14];
  ws.set(0, 0, ymLabel(ym) + "　勤務時間集計", st({ bold: true, size: 12 })); ws.rowH[0] = 20;
  ["社員番号", "名前", "出勤日数", "勤務時間合計"].forEach((t, c) => cell(2, c, t, { h: "center", fill: GRAY }));
  let r = 3, sumDays = 0, sumMin = 0;
  people.forEach(p => {
    const s = summarize(months[p.id]); sumDays += s.days; sumMin += s.total;
    cell(r, 0, p.no || "", { h: "center" }); cell(r, 1, p.name); cell(r, 2, s.days, { h: "center" });
    cell(r, 3, s.total / 1440, { h: "center", fmt: "[h]:mm" }); ws.rowH[r] = 18; r++;
  });
  cell(r, 0, "合計", { h: "center", bold: true }); cell(r, 1, null); ws.merge(r, 0, r, 1);
  cell(r, 2, sumDays, { h: "center", bold: true }, );
  cell(r, 3, sumMin / 1440, { h: "center", bold: true, fmt: "[h]:mm" });
  ws.rowH[r] = 20;
  ws.printArea = `$A$1:$D$${r + 1}`;
  ws.page = { fit: true, center: true };
  return ws;
}

function sheetName(s, used){
  const base = String(s).replace(/[\[\]:*?\/\\']/g, "").trim().slice(0, 31) || "従業員";
  let k = base, i = 2; while (used.has(k)) k = base.slice(0, 27) + "(" + (i++) + ")";
  used.add(k); return k;
}

// people: [{id,no,name}]  months: {empId: monthDoc}
export function buildWorkbook(people, months, ym){
  const wb = new Workbook(), used = new Set(["集計"]);
  people.forEach(p => timecardSheet(wb, sheetName(p.name, used), p, months[p.id], ym));
  if (people.length > 1) summarySheet(wb, people, months, ym);
  return wb;
}

export function downloadExcel(people, months, ym){
  const bytes = buildWorkbook(people, months, ym).toBytes();
  const { y, m } = ymParts(ym);
  const name = people.length > 1 ? `出勤時間_${y}年${m}月.xlsx` : `出勤時間_${y}年${m}月_${people[0].name}.xlsx`;
  const url = URL.createObjectURL(new Blob([bytes], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }));
  const a = document.createElement("a"); a.href = url; a.download = name;
  document.body.appendChild(a); a.click(); a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 5000);
}
