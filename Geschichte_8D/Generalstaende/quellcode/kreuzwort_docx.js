const fs = require("fs");
const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, BorderStyle, AlignmentType, VerticalAlign, PageBreak, HeightRule } = require("docx");
const D = JSON.parse(fs.readFileSync("kreuzwort_daten.json", "utf8"));
const NAVY = "1E2761", MUTED = "5B6B8C";
const t = (text, o = {}) => new TextRun({ text, font: o.font || "Calibri", size: o.size || 20, bold: o.bold, italics: o.italics, color: o.color || "000000" });
const p = (runs, o = {}) => new Paragraph({ children: Array.isArray(runs) ? runs : [runs], alignment: o.align, spacing: { before: o.before ?? 0, after: o.after ?? 40 } });
const none = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
const line = { style: BorderStyle.SINGLE, size: 8, color: "5B6B8C" };
const dash = { style: BorderStyle.DASHED, size: 6, color: "8A96B5" };
const CS = 400;
const numAt = new Map();
D.words.forEach((w) => { const k = w.r + "," + w.c; if (!numAt.has(k)) numAt.set(k, w.num); });

function gridTable(show) {
  const rows = D.grid.map((row, r) => new TableRow({ height: { value: CS, rule: HeightRule.EXACT }, children: row.map((ch, c) => {
    if (!ch) return new TableCell({ width: { size: CS, type: WidthType.DXA }, borders: { top: none, bottom: none, left: none, right: none }, children: [p(t(" ", { size: 8 }), { after: 0 })] });
    const num = numAt.get(r + "," + c);
    const kids = [];
    if (num) kids.push(p(t(String(num), { size: 11, color: MUTED }), { after: 0 }));
    kids.push(p(t(show ? ch : " ", { size: 24, bold: true, color: NAVY }), { after: 0, align: AlignmentType.CENTER }));
    return new TableCell({ width: { size: CS, type: WidthType.DXA }, borders: { top: line, bottom: line, left: line, right: line }, margins: { top: 0, bottom: 0, left: 20, right: 20 }, verticalAlign: VerticalAlign.CENTER, children: kids });
  }) }));
  return new Table({ width: { size: CS * D.cols, type: WidthType.DXA }, columnWidths: Array(D.cols).fill(CS), rows });
}
function clues() {
  const across = D.words.filter((w) => w.dir === "h"), down = D.words.filter((w) => w.dir === "v");
  const col = (title, list) => new TableCell({ width: { size: 5350, type: WidthType.DXA }, borders: { top: none, bottom: none, left: none, right: none }, margins: { top: 20, bottom: 20, left: 40, right: 100 }, children: [p(t(title, { bold: true, size: 20, color: NAVY }), { after: 20 }), ...list.map((w) => p([t(w.num + "  ", { bold: true, size: 18 }), t(w.clue, { size: 18 })], { after: 15 }))] });
  return new Table({ width: { size: 10700, type: WidthType.DXA }, columnWidths: [5350, 5350], rows: [new TableRow({ children: [col("Waagerecht →", across), col("Senkrecht ↓", down)] })] });
}
function blatt() {
  return [
    p([t("KREUZWORTRÄTSEL · GESCHICHTE 8D", { size: 15, bold: true, color: MUTED })], { after: 10 }),
    p(t("Absolutismus und Stände", { font: "Cambria", size: 30, bold: true, color: NAVY }), { after: 30 }),
    p(t("Löse das Rätsel. Ä, Ö und Ü sind je ein Feld. Schreibe danach die Lösungswörter ins Heft.", { size: 18, italics: true, color: MUTED }), { after: 50 }),
    gridTable(false),
    p(t(" ", { size: 8 }), { after: 40 }),
    clues(),
  ];
}
const sep = new Paragraph({ children: [new TextRun({ text: "✂ - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -", size: 14, color: "8A96B5" })], spacing: { before: 60, after: 60 }, alignment: AlignmentType.CENTER });
const sol = [
  p(t("Lösung für die Lehrkraft", { font: "Cambria", size: 30, bold: true, color: NAVY }), { after: 60 }),
  gridTable(true),
  p(t(" ", { size: 8 }), { after: 60 }),
  clues(),
  p(t("Die acht Lösungswörter (ABSOLUTISMUS, SONNENKÖNIG, VERSAILLES, GOTTESGNADENTUM, KLERUS, ADEL, STEUERN, VORRECHTE) werden als Begriffsspeicher in den nächsten Stunden wiederholt.", { size: 19, italics: true, color: MUTED }), { before: 80 }),
];
(async () => {
  const doc = new Document({ styles: { default: { document: { run: { font: "Calibri", size: 20 } } } }, sections: [
    { properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 500, bottom: 450, left: 600, right: 600 } } }, children: [...blatt(), sep, ...blatt()] },
    { properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 700, bottom: 600, left: 700, right: 700 } } }, children: sol },
  ] });
  fs.writeFileSync("../Kreuzwortraetsel_Absolutismus_Staende.docx", await Packer.toBuffer(doc));
  console.log("ok");
})();
