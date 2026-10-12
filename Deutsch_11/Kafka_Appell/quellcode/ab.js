// Arbeitsblatt „Kafkas Appell“: nur Kontext und Text (Drucke sparen), Worterklärungen als Fußnoten
const fs = require("fs");
const path = require("path");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, BorderStyle,
  ShadingType, AlignmentType, VerticalAlign, HeightRule, Header, Footer, TabStopType, PageNumber,
} = require("docx");

const NAVY = "1E2761", MUTED = "5B6B8C", LIGHT = "F4F7FD", LIGHT2 = "EAF0FA", RED = "D9534F";
const COL_A = "3F7CC4", COL_C = "D9534F";
const W = 10206;
const { TEXT, QUELLE, GLOSSAR, KONTEXT } = require("./daten");
const FS = 25;

const t = (text, o = {}) => new TextRun({ text, font: o.font || "Calibri", size: o.size || 22, bold: o.bold, italics: o.italics, color: o.color || "000000" });
const p = (runs, o = {}) => new Paragraph({ children: Array.isArray(runs) ? runs : [runs], alignment: o.align, spacing: { before: o.before ?? 0, after: o.after ?? 100, line: o.line } });
const none = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
const solid = (c, s = 6) => ({ style: BorderStyle.SINGLE, size: s, color: c });
const all = (b) => ({ top: b, bottom: b, left: b, right: b });
const cell = (children, o = {}) => new TableCell({
  children: Array.isArray(children) ? children : [children], width: { size: o.w, type: WidthType.DXA },
  shading: o.fill ? { fill: o.fill, type: ShadingType.CLEAR, color: "auto" } : undefined,
  borders: o.borders || all(solid("B9C6E8")), verticalAlign: o.valign || VerticalAlign.TOP,
  margins: { top: o.m ?? 50, bottom: o.m ?? 50, left: o.ml ?? 100, right: o.ml ?? 100 }, columnSpan: o.span,
});
const table = (widths, rows) => new Table({ width: { size: widths.reduce((a, b) => a + b, 0), type: WidthType.DXA }, columnWidths: widths, rows });
const row = (cells, h) => new TableRow({ children: cells, height: h ? { value: h, rule: HeightRule.ATLEAST } : undefined, cantSplit: true });
const box = (paras, fill = LIGHT2) => table([W], [row([cell(paras, { w: W, fill, borders: all(none), m: 120, ml: 160 })])]);

// Text mit Zeilennummern (jede Zeile nummeriert, feste Umbrüche); Worterklärungen als hochgestellte Fußnoten
// glossar: [Wort, Erklärung] – Fußnotenziffer hinter dem ersten Vorkommen (ganzes Wort); Einträge ohne Treffer als Hinweis ohne Ziffer
function textBlock(lines, glossar, col, start = 1) {
  const nW = 520, tW = W - nW;
  const marks = lines.map(() => []);
  const notes = [];
  // Fundstellen ermitteln und nach Position im Text sortieren, damit die Ziffern aufsteigend erscheinen
  const found = glossar.map(([w, e]) => {
    const li = /^\w/.test(w) ? lines.findIndex((l) => l.includes(w)) : -1;
    if (li < 0) return { w, e, li: Infinity };
    let end = lines[li].indexOf(w) + w.length;
    while (end < lines[li].length && /[\p{L}]/u.test(lines[li][end])) end++;
    return { w, e, li, end };
  }).sort((a, b) => a.li - b.li || a.end - b.end);
  let n = start;
  found.forEach(({ w, e, li, end }) => {
    if (li === Infinity) { notes.push({ w, e }); return; }
    marks[li].push({ end, n });
    notes.push({ w, e, n: n++ });
  });
  const lineRuns = (l, ms) => {
    const runs = []; let pos = 0;
    ms.sort((a, b) => a.end - b.end).forEach(({ end, n }) => {
      runs.push(t(l.slice(pos, end), { font: "Cambria", size: FS }));
      runs.push(new TextRun({ text: String(n), font: "Calibri", size: FS, bold: true, color: col, superScript: true }));
      pos = end;
    });
    runs.push(t(l.slice(pos), { font: "Cambria", size: FS }));
    return runs;
  };
  const rows = lines.map((l, i) => row([
    cell(p(t(String(i + 1), { size: 16, color: MUTED }), { align: AlignmentType.RIGHT, after: 0 }), { w: nW, borders: { top: none, bottom: none, left: none, right: solid(col, 10) }, m: 6, ml: 80, valign: VerticalAlign.CENTER }),
    cell(p(lineRuns(l, marks[i]), { after: 0, line: 360 }), { w: tW, borders: all(none), m: 6, ml: 140 }),
  ]));
  notes.forEach(({ w, e, n }, i) => rows.push(row([cell(p([
    n ? new TextRun({ text: String(n), font: "Calibri", size: 19, bold: true, color: col, superScript: true }) : t("", { size: 19 }),
    t((n ? " " : "") + w + " ", { size: 19, bold: true, color: col }), t("= " + e, { size: 19, color: MUTED }),
  ], { after: 0 }), { w: W, span: 2, borders: { top: i ? none : solid("D5DCEC", 6), bottom: none, left: none, right: none }, m: i ? 0 : 50, ml: 140 })])));
  return table([nW, tW], rows);
}
const header = new Header({ children: [new Paragraph({ tabStops: [{ type: TabStopType.RIGHT, position: W }], children: [t("Deutsch · Klasse 11 · Kafka, „Brief an den Vater“", { size: 17, color: MUTED }), t("\tName: ______________________   Datum: ___________", { size: 17, color: MUTED })] })] });
const footer = new Footer({ children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ font: "Calibri", size: 16, color: MUTED, children: [QUELLE + " · DigBib-Ausgabe"] })] })] });

const page = [
  p(t("FRANZ KAFKA: „BRIEF AN DEN VATER“ (1919)", { size: 17, bold: true, color: MUTED }), { after: 20 }),
  p(t("Worum bittet Kafka?", { font: "Cambria", size: 40, bold: true, color: NAVY }), { after: 140 }),
  box([
    p(t("Kontext", { bold: true, color: RED, size: 22 }), { after: 40 }),
    p(t(KONTEXT, { size: 22 }), { after: 0, line: 300 }),
  ]),
  p(t("Der Text steht in der Rechtschreibung vor 1996: daß = dass, mußte = musste, gewiß = gewiss.", { size: 17, italics: true, color: MUTED }), { before: 100, after: 160 }),
  textBlock(TEXT, GLOSSAR, COL_A),
  p(t(QUELLE, { size: 17, color: MUTED }), { before: 120, align: AlignmentType.RIGHT }),
];

const doc = new Document({
  styles: { default: { document: { run: { font: "Calibri", size: 22 } } } },
  sections: [{ properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 850, bottom: 750, left: 850, right: 850, header: 380, footer: 380 } } }, headers: { default: header }, footers: { default: footer }, children: page }],
});
Packer.toBuffer(doc).then((b) => fs.writeFileSync("Arbeitsblatt_Kafkas_Appell.docx", b));
