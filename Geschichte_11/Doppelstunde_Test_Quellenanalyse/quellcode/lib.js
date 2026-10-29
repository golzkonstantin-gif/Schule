// Gemeinsame docx-Helfer (Geschichte 11)
const { Paragraph, TextRun, Table, TableRow, TableCell, WidthType, BorderStyle, ShadingType, HeightRule } = require("docx");
const C = { dark: "2B2318", terra: "AD4426", sand: "F4ECDE", gold: "D9A441", muted: "5A5044", soft: "998C77", line: "BFB5A5", pale: "FBF8F2", green: "3C783C", blue: "3E5C76", white: "FFFFFF" };
const thin = { style: BorderStyle.SINGLE, size: 4, color: C.line }, none = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
function runs(text, o = {}) {
  return String(text).split(/(\*\*[^*]+\*\*)/).filter(Boolean).map(t => new TextRun({ text: t.startsWith("**") ? t.slice(2, -2) : t, bold: t.startsWith("**") || o.bold, size: o.size || 20, color: o.color || C.dark, italics: o.italics, font: o.font || "Calibri" }));
}
const p = (t, o = {}) => new Paragraph({ children: Array.isArray(t) ? t : runs(t, o), alignment: o.align, spacing: { after: o.after ?? 70, before: o.before ?? 0, line: o.line }, keepNext: o.keepNext, tabStops: o.tabs });
const h = (t, o = {}) => new Paragraph({ children: [new TextRun({ text: t, font: "Cambria", bold: true, size: o.size || 25, color: o.color || C.terra })], spacing: { before: o.before ?? 200, after: 90 }, keepNext: true });
function table(W, cw, head, rows, o = {}) {
  const cell = (c, w, fill, bold) => new TableCell({ width: { size: w, type: WidthType.DXA }, shading: fill ? { type: ShadingType.CLEAR, color: "auto", fill } : undefined,
    borders: { top: thin, bottom: thin, left: thin, right: thin }, margins: { top: 50, bottom: 50, left: 90, right: 90 },
    children: (Array.isArray(c) ? c : [c]).map(x => typeof x === "string" ? p(x, { size: o.size || 18, after: 20, bold, color: bold && fill === C.dark ? C.white : undefined }) : x) });
  const rs = [];
  if (head) rs.push(new TableRow({ tableHeader: true, children: head.map((t, i) => cell(t, cw[i], C.dark, true)) }));
  rows.forEach(r => rs.push(new TableRow({ cantSplit: true, height: o.h ? { value: o.h, rule: HeightRule.ATLEAST } : undefined, children: r.map((t, i) => cell(t, cw[i], o.firstFill && i === 0 ? C.sand : undefined, o.firstFill && i === 0)) })));
  return new Table({ columnWidths: cw, width: { size: W, type: WidthType.DXA }, rows: rs });
}
function box(W, children, fill = C.sand) {
  return new Table({ columnWidths: [W], width: { size: W, type: WidthType.DXA }, rows: [new TableRow({ children: [new TableCell({ children, width: { size: W, type: WidthType.DXA },
    shading: { type: ShadingType.CLEAR, color: "auto", fill }, borders: { top: none, bottom: none, left: none, right: none }, margins: { top: 110, bottom: 110, left: 170, right: 170 } })] })] });
}
function lines(W, n, h = 400) {
  return new Table({ columnWidths: [W], width: { size: W, type: WidthType.DXA }, rows: Array.from({ length: n }, () => new TableRow({ height: { value: h, rule: HeightRule.EXACT },
    children: [new TableCell({ width: { size: W, type: WidthType.DXA }, borders: { top: none, left: none, right: none, bottom: thin }, children: [new Paragraph({ children: [] })] })] })) });
}
// Notenpunkte (Oberstufe) nach Prozent; Schwelle = Mindestpunktzahl, auf halbe Punkte aufgerundet
const PROZ = [[15, 95], [14, 90], [13, 85], [12, 80], [11, 75], [10, 70], [9, 65], [8, 60], [7, 55], [6, 50], [5, 45], [4, 40], [3, 33], [2, 27], [1, 20], [0, 0]];
const schluessel = max => PROZ.map(([np, pr]) => [np, pr, Math.ceil(max * pr / 100 * 2) / 2]);
module.exports = { C, thin, none, runs, p, h, table, box, lines, schluessel };
