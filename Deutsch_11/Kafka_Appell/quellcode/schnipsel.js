// Schnipsel zum Austeilen: Übungssätze (Großschreibung, Kommas) + Regeln, drei Schnipsel pro A4-Seite (Schnittlinien gestrichelt)
const fs = require("fs");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, BorderStyle,
  ShadingType, VerticalAlign, HeightRule,
} = require("docx");
const { SAETZE, ZUSAMMEN, zaehle, REGELN } = require("./daten");

const NAVY = "1E2761", MUTED = "5B6B8C", RED = "D9534F", GREEN = "2E9E6B", LIGHT = "F4F7FD";
const W = 10606, SATZ_W = 7300, RULE_W = W - SATZ_W, SN_H = 5360;
const allein = zaehle(SAETZE.slice(ZUSAMMEN));

const t = (text, o = {}) => new TextRun({ text, font: o.font || "Calibri", size: o.size || 20, bold: o.bold, italics: o.italics, color: o.color || "111111" });
const p = (runs, o = {}) => new Paragraph({ children: Array.isArray(runs) ? runs : [runs], spacing: { before: o.before ?? 0, after: o.after ?? 0, line: o.line } });
const none = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
const dash = { style: BorderStyle.DASHED, size: 6, color: "9AA6C4" };
const cell = (children, o = {}) => new TableCell({
  children, width: { size: o.w, type: WidthType.DXA },
  shading: o.fill ? { fill: o.fill, type: ShadingType.CLEAR, color: "auto" } : undefined,
  borders: o.borders, verticalAlign: VerticalAlign.TOP,
  margins: { top: 110, bottom: 80, left: 160, right: 160 },
});

function schnipsel() {
  const links = [
    p(t("RECHTSCHREIBTRAINING", { bold: true, color: MUTED, size: 15 }), { after: 30 }),
    p(t("Groß oder klein? Setze die fehlenden Kommas.", { bold: true, color: NAVY, size: 21 }), { after: 80 }),
    ...SAETZE.map((s, i) => p([
      t((i + 1) + "  ", { bold: true, color: i < ZUSAMMEN ? GREEN : MUTED, size: 18 }),
      t(s.f, { font: "Cambria", size: 20 }),
    ], { line: 360, after: 0 })),
    p(t(`Sätze 1–${ZUSAMMEN} lösen wir zusammen. Ab Satz ${ZUSAMMEN + 1} allein: ${allein.gross + allein.komma} Fehler (${allein.gross} × groß, ${allein.komma} × Komma).`, { size: 16, italics: true, color: MUTED }), { before: 40 }),
  ];
  const rechts = [];
  REGELN.forEach((r, i) => {
    rechts.push(p(t("Regel " + (i + 1) + ": " + r.titel, { bold: true, color: i ? RED : GREEN, size: 19 }), { before: i ? 140 : 0, after: 30 }));
    rechts.push(p(t(r.kern, { bold: true, size: 18 }), { after: 30 }));
    r.punkte.forEach((pt) => rechts.push(p(t("• " + pt, { size: 17 }), { after: 25 })));
  });
  return new TableRow({
    height: { value: SN_H, rule: HeightRule.EXACT }, cantSplit: true,
    children: [
      cell(links, { w: SATZ_W, borders: { top: dash, bottom: dash, left: dash, right: none } }),
      cell(rechts, { w: RULE_W, fill: LIGHT, borders: { top: dash, bottom: dash, left: none, right: dash } }),
    ],
  });
}

const doc = new Document({
  styles: { default: { document: { run: { font: "Calibri", size: 20 } } } },
  sections: [{
    properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 380, bottom: 200, left: 650, right: 650 } } },
    children: [new Table({ width: { size: W, type: WidthType.DXA }, columnWidths: [SATZ_W, RULE_W], rows: [schnipsel(), schnipsel(), schnipsel()] })],
  }],
});
Packer.toBuffer(doc).then((b) => fs.writeFileSync("Schnipsel_Rechtschreibtraining.docx", b));
