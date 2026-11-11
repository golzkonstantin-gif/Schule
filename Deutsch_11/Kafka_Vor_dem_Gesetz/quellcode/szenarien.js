// Arbeitsblatt „Nervig oder kafkaesk?“ (Aufgabe 1): ein A4-Blatt, je Situation ankreuzen und mit einem Merkmal von „kafkaesk“ begründen
const fs = require("fs");
const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, BorderStyle, ShadingType, HeightRule, VerticalAlign, Header, TabStopType } = require("docx");
const D = require("./daten");

const NAVY = "1E2761", MUTED = "5B6B8C", LIGHT = "F4F7FD";
const W = 10206;
const t = (text, o = {}) => new TextRun({ text, font: o.font || "Calibri", size: o.size || 21, bold: o.bold, italics: o.italics, color: o.color || "000000" });
const p = (runs, o = {}) => new Paragraph({ children: Array.isArray(runs) ? runs : [runs], alignment: o.align, keepNext: o.keepNext, spacing: { before: o.before ?? 0, after: o.after ?? 60, line: o.line } });
const none = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
const all = (b) => ({ top: b, bottom: b, left: b, right: b });
const linie = { style: BorderStyle.SINGLE, size: 4, color: "9AA6C4" };
const nW = 520, tW = W - nW;

const zelle = (children, w, o = {}) => new TableCell({ width: { size: w, type: WidthType.DXA }, borders: o.borders || all(none), shading: o.fill ? { fill: o.fill, type: ShadingType.CLEAR, color: "auto" } : undefined, verticalAlign: o.valign || VerticalAlign.TOP, columnSpan: o.span, margins: { top: o.m ?? 40, bottom: o.m ?? 40, left: o.ml ?? 120, right: 80 }, children });

function situation([, text], i) {
  return [
    new TableRow({ cantSplit: true, children: [
      zelle([p(t(String(i + 1), { bold: true, color: NAVY, size: 24 }), { after: 0 })], nW, { fill: LIGHT, ml: 100 }),
      zelle([
        p(t(text, { size: 21 }), { after: 40, line: 270 }),
        p([t("☐ nervig        ☐ kafkaesk", { size: 22, bold: true, color: NAVY })], { after: 40 }),
        p(t("Begründung (Merkmal):", { size: 18, color: MUTED }), { after: 0 }),
      ], tW),
    ] }),
    ...[0, 1].map(() => new TableRow({ cantSplit: true, height: { value: 360, rule: HeightRule.ATLEAST }, children: [
      zelle([p(t(""), { after: 0 })], nW, { fill: LIGHT }),
      zelle([p(t(""), { after: 0 })], tW, { borders: { top: none, left: none, right: none, bottom: linie }, m: 0 }),
    ] })),
    new TableRow({ height: { value: 120, rule: HeightRule.EXACT }, children: [zelle([p(t(""), { after: 0 })], W, { span: 2, m: 0 })] }),
  ];
}

const header = new Header({ children: [new Paragraph({ tabStops: [{ type: TabStopType.RIGHT, position: W }], children: [t("Deutsch · Klasse 11 · Kafka", { size: 17, color: MUTED }), t("\tName: ______________________   Datum: ___________", { size: 17, color: MUTED })] })] });
const doc = new Document({
  styles: { default: { document: { run: { font: "Calibri", size: 21 } } } },
  sections: [{ properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 700, bottom: 500, left: 850, right: 850, header: 340 } } }, headers: { default: header }, children: [
    p(t(D.NK_AUFGABE.titel, { font: "Cambria", size: 30, bold: true, color: NAVY }), { after: 60 }),
    ...D.NK_AUFGABE.schritte.map((x, i) => p([t((i + 1) + ".  ", { bold: true, color: NAVY, size: 20 }), t(x, { size: 20 })], { after: 30 })),
    p(t(""), { after: 40 }),
    new Table({ width: { size: W, type: WidthType.DXA }, columnWidths: [W], rows: [new TableRow({ children: [zelle([
      p(t("Merkmale von „kafkaesk“", { bold: true, color: NAVY, size: 19 }), { after: 20 }),
      ...D.KAFKAESK.bausteine.map(([b, e]) => p([t("▸ " + b + ": ", { bold: true, size: 18 }), t(e, { size: 18 })], { after: 0 })),
      p([t("▸ Gefühle: ", { bold: true, size: 18 }), t(D.KAFKAESK.gefuehle, { size: 18 })], { after: 0 }),
    ], W, { fill: "EAF0FA", m: 60, ml: 160 })] })] }),
    p(t(""), { after: 60 }),
    new Table({ width: { size: W, type: WidthType.DXA }, columnWidths: [nW, tW], rows: D.SZENARIEN.flatMap(situation) }),
  ] }],
});
Packer.toBuffer(doc).then((b) => fs.writeFileSync("Szenarien_Nervig_oder_kafkaesk.docx", b));
