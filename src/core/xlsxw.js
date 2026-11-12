// 极简带样式 xlsx 写出器（本地自写 OOXML，ZIP 存储方式）：字体、底色、边框、对齐、数字格式、列宽、行高、合并
import { zipStore } from './zip.js';

const xmlEsc = (s) => String(s).replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c])).replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, '');
export const colName = (i) => { let s = ''; i += 1; while (i) { const m = (i - 1) % 26; s = String.fromCharCode(65 + m) + s; i = Math.floor((i - 1) / 26); } return s; };
export const ref = (r, c) => colName(c) + (r + 1);

/** 样式表：style = { font:{name,sz,b,color}, fill:'FFRRGGBB', border:'thin', h:'center', v:'center', wrap, numFmt:'#,##0.00' } */
class Styles {
  constructor() { this.fonts = []; this.fills = ['<fill><patternFill patternType="none"/></fill>', '<fill><patternFill patternType="gray125"/></fill>']; this.borders = ['<border><left/><right/><top/><bottom/><diagonal/></border>']; this.fmts = []; this.xfs = []; this.map = new Map(); this.add({}); }
  idx(list, xml) { let i = list.indexOf(xml); if (i < 0) { list.push(xml); i = list.length - 1; } return i; }
  add(st) {
    const key = JSON.stringify(st); if (this.map.has(key)) return this.map.get(key);
    const f = { name: '宋体', sz: 11, ...(st.font ?? {}) };
    const font = this.idx(this.fonts, `<font>${f.b ? '<b/>' : ''}<sz val="${f.sz}"/>${f.color ? `<color rgb="${f.color}"/>` : '<color theme="1"/>'}<name val="${xmlEsc(f.name)}"/><charset val="134"/></font>`);
    const fill = st.fill ? this.idx(this.fills, `<fill><patternFill patternType="solid"><fgColor rgb="${st.fill}"/><bgColor indexed="64"/></patternFill></fill>`) : 0;
    const b = st.border ? `<left style="${st.border}"><color rgb="FF000000"/></left><right style="${st.border}"><color rgb="FF000000"/></right><top style="${st.border}"><color rgb="FF000000"/></top><bottom style="${st.border}"><color rgb="FF000000"/></bottom><diagonal/>` : null;
    const border = b ? this.idx(this.borders, `<border>${b}</border>`) : 0;
    let numFmtId = 0;
    if (st.numFmt) { let i = this.fmts.indexOf(st.numFmt); if (i < 0) { this.fmts.push(st.numFmt); i = this.fmts.length - 1; } numFmtId = 176 + i; }
    const al = `<alignment${st.h ? ` horizontal="${st.h}"` : ''} vertical="${st.v ?? 'center'}"${st.wrap ? ' wrapText="1"' : ''}/>`;
    this.xfs.push(`<xf numFmtId="${numFmtId}" fontId="${font}" fillId="${fill}" borderId="${border}" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1">${al}</xf>`);
    const id = this.xfs.length - 1; this.map.set(key, id); return id;
  }
  xml() {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">`
      + (this.fmts.length ? `<numFmts count="${this.fmts.length}">${this.fmts.map((f, i) => `<numFmt numFmtId="${176 + i}" formatCode="${xmlEsc(f)}"/>`).join('')}</numFmts>` : '')
      + `<fonts count="${this.fonts.length}">${this.fonts.join('')}</fonts><fills count="${this.fills.length}">${this.fills.join('')}</fills><borders count="${this.borders.length}">${this.borders.join('')}</borders>`
      + `<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="${this.xfs.length}">${this.xfs.join('')}</cellXfs>`
      + `<cellStyles count="1"><cellStyle name="常规" xfId="0" builtinId="0"/></cellStyles></styleSheet>`;
  }
}

/**
 * sheets: [{ name, rows: [[{ v, s }]]（s 为样式对象）, merges:['A1:M1'], cols:[{min,max,width}], heights:{行号(1起):ht}, defaultRowHeight, defaultColWidth }]
 * 日期请传 { v: Date, s:{numFmt:'yyyy-mm-dd'} }
 */
export function writeStyledXlsx(sheets) {
  const styles = new Styles();
  const enc = new TextEncoder();
  const files = [];
  const sheetXml = sheets.map((sh) => {
    const rows = sh.rows.map((row, r) => {
      const cells = row.map((cell, c) => {
        if (cell == null) return '';
        const s = styles.add(cell.s ?? {});
        const v = cell.v;
        const at = `r="${ref(r, c)}" s="${s}"`;
        if (v === '' || v == null) return `<c ${at}/>`;
        if (v instanceof Date) return `<c ${at}><v>${(Date.UTC(v.getUTCFullYear(), v.getUTCMonth(), v.getUTCDate()) / 864e5 + 25569)}</v></c>`;
        if (typeof v === 'number') return `<c ${at}><v>${v}</v></c>`;
        return `<c ${at} t="inlineStr"><is><t xml:space="preserve">${xmlEsc(v)}</t></is></c>`;
      }).join('');
      const ht = sh.heights?.[r + 1];
      return `<row r="${r + 1}"${ht ? ` ht="${ht}" customHeight="1"` : ''}>${cells}</row>`;
    }).join('');
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">`
      + (sh.fitWidth ? '<sheetPr><pageSetUpPr fitToPage="1"/></sheetPr>' : '') + `<sheetViews><sheetView workbookViewId="0"/></sheetViews><sheetFormatPr defaultColWidth="${sh.defaultColWidth ?? 9.23076923076923}" defaultRowHeight="${sh.defaultRowHeight ?? 16.8}"/>`
      + (sh.cols?.length ? `<cols>${sh.cols.map((c) => `<col min="${c.min}" max="${c.max}" width="${c.width}" customWidth="1"/>`).join('')}</cols>` : '')
      + `<sheetData>${rows}</sheetData>`
      + (sh.merges?.length ? `<mergeCells count="${sh.merges.length}">${sh.merges.map((m) => `<mergeCell ref="${m}"/>`).join('')}</mergeCells>` : '')
      + `<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>${sh.fitWidth ? '<pageSetup paperSize="9" orientation="landscape" fitToWidth="1" fitToHeight="0"/>' : ''}</worksheet>`;
  });
  files.push({ name: '[Content_Types].xml', data: enc.encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>${sheets.map((_, i) => `<Override PartName="/xl/worksheets/sheet${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`).join('')}</Types>`) });
  files.push({ name: '_rels/.rels', data: enc.encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`) });
  files.push({ name: 'xl/workbook.xml', data: enc.encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><bookViews><workbookView/></bookViews><sheets>${sheets.map((s, i) => `<sheet name="${xmlEsc(s.name)}" sheetId="${i + 1}" r:id="rId${i + 1}"/>`).join('')}</sheets></workbook>`) });
  files.push({ name: 'xl/_rels/workbook.xml.rels', data: enc.encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${sheets.map((_, i) => `<Relationship Id="rId${i + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${i + 1}.xml"/>`).join('')}<Relationship Id="rId${sheets.length + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>`) });
  sheetXml.forEach((x, i) => files.push({ name: `xl/worksheets/sheet${i + 1}.xml`, data: enc.encode(x) }));
  files.push({ name: 'xl/styles.xml', data: enc.encode(styles.xml()) });
  return zipStore(files);
}
