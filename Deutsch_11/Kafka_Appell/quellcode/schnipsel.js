// Schnipsel zum Austeilen: WhatsApp-Fehlertext + Regeln, drei Schnipsel pro A4-Seite (Schnittlinien gestrichelt)
const fs = require("fs");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, BorderStyle,
  ShadingType, AlignmentType, VerticalAlign, HeightRule,
} = require("docx");
const { CHAT, zaehle, REGELN } = require("./daten");

const NAVY = "1E2761", MUTED = "5B6B8C", RED = "D9534F", GREEN = "2E9E6B";
const WA_BG = "ECE5DD", WA_ME = "DCF8C6", WA_HEAD = "075E54";
const W = 10606, CHAT_W = 7000, RULE_W = W - CHAT_W, SN_H = 5360;
const rest = zaehle(CHAT.slice(1));

const t = (text, o = {}) => new TextRun({ text, font: "Calibri", size: o.size || 20, bold: o.bold, italics: o.italics, color: o.color || "111111" });
const p = (runs, o = {}) => new Paragraph({ children: Array.isArray(runs) ? runs : [runs], alignment: o.align, spacing: { before: o.before ?? 0, after: o.after ?? 0, line: o.line } });
const none = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
const all = (b) => ({ top: b, bottom: b, left: b, right: b });
const dash = { style: BorderStyle.DASHED, size: 6, color: "9AA6C4" };
const cell = (children, o = {}) => new TableCell({
  children: Array.isArray(children) ? children : [children], width: { size: o.w, type: WidthType.DXA },
  shading: o.fill ? { fill: o.fill, type: ShadingType.CLEAR, color: "auto" } : undefined,
  borders: o.borders || all(none), verticalAlign: o.valign || VerticalAlign.TOP,
  margins: { top: o.m ?? 40, bottom: o.m ?? 40, left: o.ml ?? 100, right: o.ml ?? 100 }, columnSpan: o.span,
});
const table = (widths, rows) => new Table({ width: { size: widths.reduce((a, b) => a + b, 0), type: WidthType.DXA }, columnWidths: widths, rows });
const row = (cells, h) => new TableRow({ children: cells, height: h ? { value: h, rule: HeightRule.EXACT } : undefined, cantSplit: true });

// Eine Chatblase: links (Mia, weiß) oder rechts (Jonas, grün)
function bubble(m, i) {
  const bw = 5600, sw = CHAT_W - 240 - bw;
  const inner = cell([
    p([
      t((i + 1) + "  ", { bold: true, color: MUTED, size: 16 }),
      ...(m.seite === "l" ? [t(m.von + ":  ", { bold: true, color: GREEN, size: 19 })] : []),
      t(m.f, { size: 19 }),
      t("   " + m.zeit, { size: 14, color: "8A8A8A" }),
    ], { line: 310 }),
  ], { w: bw, fill: m.seite === "r" ? WA_ME : "FFFFFF", m: 30, ml: 110 });
  const spacer = cell(p(t("")), { w: sw });
  return table([m.seite === "r" ? sw : bw, m.seite === "r" ? bw : sw], [row(m.seite === "r" ? [spacer, inner] : [inner, spacer])]);
}

function regelnBlock() {
  const out = [];
  REGELN.forEach((r, i) => {
    out.push(p(t("Regel " + (i + 1) + ": " + r.titel, { bold: true, color: i ? RED : GREEN, size: 18 }), { before: i ? 90 : 0, after: 20 }));
    out.push(p(t(r.kern, { bold: true, size: 17 }), { after: 20 }));
    r.punkte.forEach((pt) => out.push(p(t("• " + pt, { size: 16 }), { after: 15 })));
  });
  out.push(p([t("Auftrag: ", { bold: true, color: NAVY, size: 17 }), t(`Korrigiere die Nachrichten 2–6 direkt im Chat. ${rest.gross + rest.komma} Fehler: ${rest.gross} × groß, ${rest.komma} × Komma.`, { size: 17 })], { before: 90 }));
  return out;
}

function schnipsel() {
  const chat = [
    p(t("Deutsch 11  ·  Mia, Jonas", { bold: true, color: "FFFFFF", size: 19 })),
  ];
  const chatCell = cell([
    table([CHAT_W - 240], [row([cell(chat, { w: CHAT_W - 240, fill: WA_HEAD, m: 40, ml: 120 })])]),
    ...CHAT.flatMap((m, i) => [p(t("", { size: 6 }), { line: 70 }), bubble(m, i)]),
    p(t("")),
  ], { w: CHAT_W, fill: WA_BG, m: 70, ml: 120, borders: { top: dash, bottom: dash, left: dash, right: none } });
  const ruleCell = cell([p(t("RECHTSCHREIB-CHECK", { bold: true, color: MUTED, size: 15 }), { after: 50 }), ...regelnBlock()], { w: RULE_W, m: 120, ml: 160, borders: { top: dash, bottom: dash, left: none, right: dash } });
  return row([chatCell, ruleCell], SN_H);
}

const doc = new Document({
  styles: { default: { document: { run: { font: "Calibri", size: 20 } } } },
  sections: [{
    properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 380, bottom: 200, left: 650, right: 650 } } },
    children: [new Table({
      width: { size: W, type: WidthType.DXA }, columnWidths: [CHAT_W, RULE_W],
      borders: { top: dash, bottom: dash, left: dash, right: dash, insideHorizontal: dash, insideVertical: none },
      rows: [schnipsel(), schnipsel(), schnipsel()],
    })],
  }],
});
Packer.toBuffer(doc).then((b) => fs.writeFileSync("Schnipsel_Rechtschreib_Check.docx", b));
