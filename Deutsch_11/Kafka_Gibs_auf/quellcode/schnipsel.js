// Schnipsel zum Einkleben: Merkkasten Stilmittel (bekannt + neu) und Schreibhilfe (inhaltsunabhängig), zwei Schnipsel pro A4-Seite
const fs = require("fs");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, BorderStyle,
  ShadingType, VerticalAlign, HeightRule,
} = require("docx");
const D = require("./daten");

const NAVY = "1E2761", MUTED = "5B6B8C", LIGHT = "F4F7FD", RED = "D9534F", GOLD = "E8A33D", BLUE = "3F7CC4";
const W = 10606, SN_H = 7900;

const t = (text, o = {}) => new TextRun({ text, font: o.font || "Calibri", size: o.size || 19, bold: o.bold, italics: o.italics, color: o.color || "111111" });
const p = (runs, o = {}) => new Paragraph({ children: Array.isArray(runs) ? runs : [runs], spacing: { before: o.before ?? 0, after: o.after ?? 0, line: o.line } });
const none = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
const thin = { style: BorderStyle.SINGLE, size: 4, color: "C9D2E6" };
const dash = { style: BorderStyle.DASHED, size: 6, color: "9AA6C4" };
const cell = (children, o = {}) => new TableCell({
  children: Array.isArray(children) ? children : [children], width: { size: o.w, type: WidthType.DXA },
  shading: o.fill ? { fill: o.fill, type: ShadingType.CLEAR, color: "auto" } : undefined,
  borders: o.borders || { top: thin, bottom: thin, left: thin, right: thin }, verticalAlign: o.valign || VerticalAlign.CENTER,
  margins: { top: o.m ?? 40, bottom: o.m ?? 40, left: 90, right: 90 }, columnSpan: o.span,
});
const row = (cells, h) => new TableRow({ children: cells, height: h ? { value: h, rule: HeightRule.ATLEAST } : undefined });

// Merkkasten: Stilmittel | So erkennst du es | Beispiel aus „Gibs auf!“
const IW = W - 320;
const mW = [1900, 4200, IW - 6100];
const kopf = (x, w) => cell(p(t(x, { bold: true, color: "FFFFFF", size: 18 })), { w, fill: NAVY });
const bekannt = [
  ["Parataxe", "Reihung von Hauptsätzen", "„… die Straßen rein und leer, ich ging zum Bahnhof.“"],
  ["Hypotaxe", "Satzgefüge mit Nebensätzen", "„Als ich eine Turmuhr … verglich, sah ich, dass …“"],
  ["Metapher", "bildhafter Ausdruck im übertragenen Sinn", "(z. B. „Rabenmutter“)"],
  ["Symbol", "konkreter Gegenstand für etwas Abstraktes", "die Uhr (Zeitdruck), der Weg (Lebensweg)"],
];
const merk = new Table({
  width: { size: IW, type: WidthType.DXA }, columnWidths: mW,
  rows: [
    row([kopf("Stilmittel", mW[0]), kopf("So erkennst du es", mW[1]), kopf("Beispiel aus „Gibs auf!“", mW[2])]),
    ...bekannt.map(([b, e, bsp]) => row([
      cell(p(t(b, { bold: true, color: NAVY })), { w: mW[0], fill: LIGHT }),
      cell(p(t(e, { size: 18 })), { w: mW[1] }),
      cell(p(t(bsp, { size: 17, italics: true, color: MUTED })), { w: mW[2] }),
    ])),
    row([cell(p(t("neu", { bold: true, size: 16, color: RED })), { w: IW, span: 3, fill: "FDECEA", m: 10 })]),
    ...D.STILMITTEL.map(([b, kern, bsp, test]) => row([
      cell(p(t(b, { bold: true, color: RED })), { w: mW[0], fill: LIGHT }),
      cell([p(t(kern, { size: 18 })), p(t("Test: " + test, { size: 16, color: MUTED }))], { w: mW[1] }),
      cell(p(t(bsp, { size: 17, italics: true, color: BLUE })), { w: mW[2] }),
    ])),
  ],
});

// Schreibhilfe: drei Spalten
const sW = Math.floor(IW / 3);
const schreib = new Table({
  width: { size: sW * 3, type: WidthType.DXA }, columnWidths: [sW, sW, sW],
  rows: [row(D.SCHREIBHILFE.map(([k, arr], i) => cell([
    p(t(k, { bold: true, color: i === 2 ? GOLD : NAVY, size: 18 }), { after: 30 }),
    ...arr.map((x) => p(t(x, { size: 17 }), { after: 20 })),
  ], { w: sW, fill: i === 2 ? "FFF8EC" : LIGHT, valign: VerticalAlign.TOP, borders: { top: none, bottom: none, left: none, right: none }, m: 60 })))],
});

function schnipsel() {
  return new TableRow({
    height: { value: SN_H, rule: HeightRule.EXACT }, cantSplit: true,
    children: [new TableCell({
      width: { size: W, type: WidthType.DXA }, verticalAlign: VerticalAlign.TOP,
      borders: { top: dash, bottom: dash, left: dash, right: dash },
      margins: { top: 120, bottom: 80, left: 160, right: 160 },
      children: [
        p([t("Kafka, „Gibs auf!“ · ", { bold: true, color: MUTED, size: 16 }), t("Stilmittel – Merkkasten", { bold: true, color: NAVY, size: 22, font: "Cambria" })], { after: 80 }),
        merk,
        p(t("Wofür? Stilmittel sind Bedeutungsträger: Sie zeigen, wie der Erzähler die Situation erlebt. So erklärst du die Wirkung: „Dadurch wirkt …“", { size: 16, italics: true, color: MUTED }), { before: 60, after: 100 }),
        p(t("Schreibhilfe für dein eigenes Gespräch", { bold: true, color: NAVY, size: 20, font: "Cambria" }), { after: 60 }),
        schreib,
      ],
    })],
  });
}

const doc = new Document({
  styles: { default: { document: { run: { font: "Calibri", size: 19 } } } },
  sections: [{
    properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 450, bottom: 300, left: 650, right: 650 } } },
    children: [new Table({ width: { size: W, type: WidthType.DXA }, columnWidths: [W], rows: [schnipsel(), schnipsel()] })],
  }],
});
Packer.toBuffer(doc).then((b) => fs.writeFileSync("Schnipsel_Stilmittel_Schreibhilfe.docx", b));
