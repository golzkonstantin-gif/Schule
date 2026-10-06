const fs = require("fs");
const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, BorderStyle, ShadingType, AlignmentType, VerticalAlign, HeightRule } = require("docx");

const NAVY = "1E2761", MUTED = "5B6B8C", LIGHT = "F4F7FD", LIGHT2 = "EAF0FA", BAND = "5B6B8C", TOTAL = "D8E1F3";
const W = 10206;
const COLS = [1900, 4506, 900, 1450, 1450]; // Kriterium, Erwartung, Punkte, erreicht, Lehrkraft

const t = (text, o = {}) => new TextRun({ text, font: o.font || "Calibri", size: o.size || 18, bold: o.bold, color: o.color || "000000" });
const p = (runs, align) => new Paragraph({ children: [].concat(runs), alignment: align, spacing: { before: 0, after: 0 } });
const side = (c = "B9C6E8") => ({ style: BorderStyle.SINGLE, size: 6, color: c });
const borders = { top: side(), bottom: side(), left: side(), right: side() };

const cell = (content, w, o = {}) => new TableCell({
  children: [content], width: { size: w, type: WidthType.DXA }, columnSpan: o.span, borders,
  shading: o.fill ? { fill: o.fill, type: ShadingType.CLEAR, color: "auto" } : undefined,
  verticalAlign: VerticalAlign.CENTER, margins: { top: 50, bottom: 50, left: 90, right: 90 },
});
const row = (cells, h) => new TableRow({ children: cells, cantSplit: true, height: h ? { value: h, rule: HeightRule.ATLEAST } : undefined });

const head = row(["Kriterium", "Erwartung", "Punkte", "erreicht", "Lehrkraft"].map((x, i) =>
  cell(p(t(x, { bold: true, color: "FFFFFF", size: 19 }), i === 1 || i === 0 ? AlignmentType.LEFT : AlignmentType.CENTER), COLS[i], { fill: NAVY })));
const band = (text) => row([cell(p(t(text, { bold: true, color: "FFFFFF", size: 20 })), W, { span: 5, fill: BAND })]);
const crit = (name, text, pts) => row([
  cell(p(t(name, { bold: true, color: NAVY })), COLS[0], { fill: LIGHT }),
  cell(p(t(text, { size: 17 })), COLS[1]),
  cell(p(t(String(pts), { bold: true, size: 20 }), AlignmentType.CENTER), COLS[2]),
  cell(p(t("")), COLS[3]), cell(p(t("")), COLS[4]),
], 620);
const sum = (label, pts, fill = LIGHT2) => row([
  cell(p(t(label, { bold: true, color: NAVY })), COLS[0] + COLS[1], { span: 2, fill }),
  cell(p(t(String(pts), { bold: true, size: 20 }), AlignmentType.CENTER), COLS[2], { fill }),
  cell(p(t("")), COLS[3], { fill }), cell(p(t("")), COLS[4], { fill }),
], 480);

const rows = [
  head,
  band("I  Inhalt – Anwendung des Vier-Seiten-Modells und der Techniken"),
  crit("Sachinhalt klären", "Benennt die Kritik sachlich und konkret (z. B. Defensivarbeit, Einsatz im Training) – ohne Beleidigung und ohne „Flasche leer“.", 4),
  crit("Selbstoffenbarung ehrlich", "Zeigt Frust und Enttäuschung als Ich-Botschaft („Ich war enttäuscht, als …“) statt als Wutausbruch.", 4),
  crit("Beziehung reparieren", "Entschuldigt sich für die öffentliche Kritik, zeigt Respekt, geht auf die Sicht der Spieler ein (Empathie).", 4),
  crit("Appell an die Richtigen", "Richtet eine konkrete, erfüllbare Bitte direkt an die Spieler (wer, was, bis wann) – z. B. Gespräch unter vier Augen, Einsatz am Samstag.", 4),
  crit("Metakommunikation", "Spricht über die Pressekonferenz selbst: was schiefgelaufen ist und wie man künftig miteinander redet (intern statt über die Medien).", 4),
  sum("Summe Inhalt", 20),
  band("II  Sprache – Darstellung, Rechtschreibung, Grammatik, Zeichensetzung"),
  crit("Ausdruck und Ton", "Angemessener, respektvoller Ton; treffende Wortwahl; abwechslungsreicher Satzbau.", 2),
  crit("Briefform", "Ort, Datum, Anrede, Einleitung, Schluss, Grußformel.", 2),
  crit("Rechtschreibung", "Korrekte Schreibung, auch Groß- und Kleinschreibung; Anredepronomen (du/ihr bzw. Sie) einheitlich.", 2),
  crit("Grammatik", "Korrekter Satzbau, Kasus, Tempus und Bezüge.", 2),
  crit("Zeichensetzung", "Kommas (v. a. bei Nebensätzen und Anrede), Satzzeichen, Zeichen bei wörtlicher Rede.", 2),
  sum("Summe Sprache", 10),
  sum("Gesamt (Inhalt + Sprache)", 30, TOTAL),
];

const doc = new Document({
  sections: [{
    properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 850, bottom: 850, left: 850, right: 850 } } },
    children: [
      new Paragraph({ spacing: { after: 140 }, children: [t("✓  ", { font: "Cambria", size: 28, bold: true, color: "D9534F" }), t("Bewertungskriterien – Inhalt und Sprache ergeben die Gesamtnote", { font: "Cambria", size: 28, bold: true, color: NAVY })] }),
      new Table({ width: { size: W, type: WidthType.DXA }, columnWidths: COLS, rows }),
      new Paragraph({ spacing: { before: 120 }, children: [t("Note: ____________      Anmerkungen: ______________________________________________________________", { color: MUTED })] }),
    ],
  }],
});
Packer.toBuffer(doc).then((b) => { fs.writeFileSync("../Bewertungsraster_Brief.docx", b); });
