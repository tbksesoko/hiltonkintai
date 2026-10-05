// 書式付き .xlsx を作る小さな書き出し処理（外部ライブラリ不要）
// 罫線・塗り・フォント・結合・列幅・印刷設定・数式（計算済みの値つき）に対応

const enc = new TextEncoder();
const xmlEsc = s => String(s).replace(/[&<>"]/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]));

/* ---------- スタイル登録 ---------- */
class Styles {
  constructor(){
    this.fonts = ['<font><sz val="11"/><name val="MS PGothic"/><family val="3"/><charset val="128"/></font>'];
    this.fills = ['<fill><patternFill patternType="none"/></fill>', '<fill><patternFill patternType="gray125"/></fill>'];
    this.borders = ['<border><left/><right/><top/><bottom/><diagonal/></border>'];
    this.numFmts = {}; this.nextFmt = 164;
    this.xfs = ['<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'];
    this.cache = new Map();
  }
  _idx(list, xml){ let i = list.indexOf(xml); if (i < 0){ list.push(xml); i = list.length - 1; } return i; }
  // spec: {bold, size, fill:"D8D8D8", border:{l,r,t,b}, fmt:"h:mm", h:"center", v:"center", wrap}
  id(spec = {}){
    const key = JSON.stringify(spec);
    if (this.cache.has(key)) return this.cache.get(key);
    const font = this._idx(this.fonts, `<font>${spec.bold ? "<b/>" : ""}<sz val="${spec.size || 11}"/><name val="MS PGothic"/><family val="3"/><charset val="128"/></font>`);
    const fill = spec.fill ? this._idx(this.fills, `<fill><patternFill patternType="solid"><fgColor rgb="FF${spec.fill}"/><bgColor indexed="64"/></patternFill></fill>`) : 0;
    const b = spec.border || {};
    const side = (n, s) => s ? `<${n} style="${s}"><color auto="1"/></${n}>` : `<${n}/>`;
    const border = this._idx(this.borders, `<border>${side("left", b.l)}${side("right", b.r)}${side("top", b.t)}${side("bottom", b.b)}<diagonal/></border>`);
    let fmt = 0;
    if (spec.fmt){ if (!(spec.fmt in this.numFmts)) this.numFmts[spec.fmt] = this.nextFmt++; fmt = this.numFmts[spec.fmt]; }
    const al = (spec.h || spec.v || spec.wrap) ? `<alignment${spec.h ? ` horizontal="${spec.h}"` : ""} vertical="${spec.v || "center"}"${spec.wrap ? ' wrapText="1"' : ""}/>` : '<alignment vertical="center"/>';
    const xf = `<xf numFmtId="${fmt}" fontId="${font}" fillId="${fill}" borderId="${border}" xfId="0"${fmt ? ' applyNumberFormat="1"' : ""} applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1">${al}</xf>`;
    const id = this._idx(this.xfs, xf);
    this.cache.set(key, id); return id;
  }
  xml(){
    const nf = Object.entries(this.numFmts).map(([c, i]) => `<numFmt numFmtId="${i}" formatCode="${xmlEsc(c)}"/>`).join("");
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">${nf ? `<numFmts count="${Object.keys(this.numFmts).length}">${nf}</numFmts>` : ""}<fonts count="${this.fonts.length}">${this.fonts.join("")}</fonts><fills count="${this.fills.length}">${this.fills.join("")}</fills><borders count="${this.borders.length}">${this.borders.join("")}</borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="${this.xfs.length}">${this.xfs.join("")}</cellXfs><cellStyles count="1"><cellStyle name="標準" xfId="0" builtinId="0"/></cellStyles></styleSheet>`;
  }
}

/* ---------- シート ---------- */
export const colName = c => { let s = ""; c++; while (c){ const m = (c - 1) % 26; s = String.fromCharCode(65 + m) + s; c = Math.floor((c - 1) / 26); } return s; };
export const ref = (r, c) => colName(c) + (r + 1);

export class Sheet {
  constructor(name){
    this.name = name; this.cells = {}; this.merges = []; this.cols = []; this.rowH = {};
    this.page = null; this.printArea = null;
  }
  // value: number | string | null ; opts: {f: formula}
  set(r, c, value, style = 0, f){ (this.cells[r] = this.cells[r] || {})[c] = { value, style, f }; }
  merge(r1, c1, r2, c2){ this.merges.push(ref(r1, c1) + ":" + ref(r2, c2)); }
  xml(){
    const rows = Object.keys(this.cells).map(Number).sort((a, b) => a - b).map(r => {
      const cs = Object.keys(this.cells[r]).map(Number).sort((a, b) => a - b).map(c => {
        const { value, style, f } = this.cells[r][c], a = ref(r, c), s = style ? ` s="${style}"` : "";
        if (f != null){
          const isStr = typeof value === "string";
          return `<c r="${a}"${s}${isStr ? ' t="str"' : ""}><f>${xmlEsc(f)}</f>${value == null || value === "" ? (isStr ? "<v></v>" : "") : `<v>${xmlEsc(value)}</v>`}</c>`;
        }
        if (value == null || value === "") return `<c r="${a}"${s}/>`;
        if (typeof value === "number") return `<c r="${a}"${s}><v>${value}</v></c>`;
        return `<c r="${a}"${s} t="inlineStr"><is><t xml:space="preserve">${xmlEsc(value)}</t></is></c>`;
      }).join("");
      const h = this.rowH[r] ? ` ht="${this.rowH[r]}" customHeight="1"` : "";
      return `<row r="${r + 1}"${h}>${cs}</row>`;
    }).join("");
    const cols = this.cols.length ? `<cols>${this.cols.map((w, i) => `<col min="${i + 1}" max="${i + 1}" width="${w}" customWidth="1"/>`).join("")}</cols>` : "";
    const merges = this.merges.length ? `<mergeCells count="${this.merges.length}">${this.merges.map(m => `<mergeCell ref="${m}"/>`).join("")}</mergeCells>` : "";
    const p = this.page || {};
    const fit = p.fit ? '<sheetPr><pageSetUpPr fitToPage="1"/></sheetPr>' : "";
    const center = p.center ? '<printOptions horizontalCentered="1"/>' : "";
    const m = p.margins || { l: 0.7, r: 0.7, t: 0.75, b: 0.75 };
    const setup = `<pageSetup paperSize="9" orientation="${p.landscape ? "landscape" : "portrait"}"${p.fit ? ' fitToWidth="1" fitToHeight="1"' : ""}/>`;
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">${fit}<sheetViews><sheetView workbookViewId="0" showGridLines="${p.grid === false ? 0 : 1}"/></sheetViews><sheetFormatPr defaultRowHeight="13.5"/>${cols}<sheetData>${rows}</sheetData>${merges}${center}<pageMargins left="${m.l}" right="${m.r}" top="${m.t}" bottom="${m.b}" header="0.3" footer="0.3"/>${setup}</worksheet>`;
  }
}

/* ---------- ブック ---------- */
export class Workbook {
  constructor(){ this.sheets = []; this.styles = new Styles(); }
  style(spec){ return this.styles.id(spec); }
  addSheet(name){ const s = new Sheet(name); this.sheets.push(s); return s; }
  files(){
    const n = this.sheets.length;
    const quote = s => "'" + s.replace(/'/g, "''") + "'";
    const names = this.sheets.map((s, i) => s.printArea ? `<definedName name="_xlnm.Print_Area" localSheetId="${i}">${xmlEsc(quote(s.name))}!${s.printArea}</definedName>` : "").join("");
    const f = {};
    f["[Content_Types].xml"] = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>${this.sheets.map((_, i) => `<Override PartName="/xl/worksheets/sheet${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`).join("")}</Types>`;
    f["_rels/.rels"] = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`;
    f["xl/workbook.xml"] = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><bookViews><workbookView/></bookViews><sheets>${this.sheets.map((s, i) => `<sheet name="${xmlEsc(s.name)}" sheetId="${i + 1}" r:id="rId${i + 1}"/>`).join("")}</sheets>${names ? `<definedNames>${names}</definedNames>` : ""}<calcPr calcId="191029" fullCalcOnLoad="1"/></workbook>`;
    f["xl/_rels/workbook.xml.rels"] = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${this.sheets.map((_, i) => `<Relationship Id="rId${i + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${i + 1}.xml"/>`).join("")}<Relationship Id="rId${n + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>`;
    this.sheets.forEach((s, i) => f[`xl/worksheets/sheet${i + 1}.xml`] = s.xml());
    f["xl/styles.xml"] = this.styles.xml();
    return f;
  }
  toBytes(){ return zip(this.files()); }
}

/* ---------- ZIP（無圧縮） ---------- */
const CRC = (() => { const t = new Uint32Array(256); for (let n = 0; n < 256; n++){ let c = n; for (let k = 0; k < 8; k++) c = c & 1 ? 0xEDB88320 ^ (c >>> 1) : c >>> 1; t[n] = c >>> 0; } return t; })();
function crc32(b){ let c = 0xFFFFFFFF; for (let i = 0; i < b.length; i++) c = CRC[(c ^ b[i]) & 0xFF] ^ (c >>> 8); return (c ^ 0xFFFFFFFF) >>> 0; }
function zip(files){
  const parts = [], central = []; let offset = 0;
  const d = new Date(), dosT = (d.getHours() << 11) | (d.getMinutes() << 5) | (d.getSeconds() >> 1), dosD = ((d.getFullYear() - 1980) << 9) | ((d.getMonth() + 1) << 5) | d.getDate();
  for (const [name, text] of Object.entries(files)){
    const nb = enc.encode(name), data = enc.encode(text), crc = crc32(data);
    const lh = new DataView(new ArrayBuffer(30));
    lh.setUint32(0, 0x04034b50, true); lh.setUint16(4, 20, true); lh.setUint16(6, 0x0800, true); lh.setUint16(8, 0, true);
    lh.setUint16(10, dosT, true); lh.setUint16(12, dosD, true); lh.setUint32(14, crc, true);
    lh.setUint32(18, data.length, true); lh.setUint32(22, data.length, true); lh.setUint16(26, nb.length, true); lh.setUint16(28, 0, true);
    parts.push(new Uint8Array(lh.buffer), nb, data);
    const ch = new DataView(new ArrayBuffer(46));
    ch.setUint32(0, 0x02014b50, true); ch.setUint16(4, 20, true); ch.setUint16(6, 20, true); ch.setUint16(8, 0x0800, true); ch.setUint16(10, 0, true);
    ch.setUint16(12, dosT, true); ch.setUint16(14, dosD, true); ch.setUint32(16, crc, true);
    ch.setUint32(20, data.length, true); ch.setUint32(24, data.length, true); ch.setUint16(28, nb.length, true);
    ch.setUint32(42, offset, true);
    central.push(new Uint8Array(ch.buffer), nb);
    offset += 30 + nb.length + data.length;
  }
  const cdSize = central.reduce((a, b) => a + b.length, 0);
  const end = new DataView(new ArrayBuffer(22));
  end.setUint32(0, 0x06054b50, true); end.setUint16(8, Object.keys(files).length, true); end.setUint16(10, Object.keys(files).length, true);
  end.setUint32(12, cdSize, true); end.setUint32(16, offset, true);
  const all = [...parts, ...central, new Uint8Array(end.buffer)];
  const out = new Uint8Array(all.reduce((a, b) => a + b.length, 0)); let p = 0;
  for (const a of all){ out.set(a, p); p += a.length; }
  return out;
}
