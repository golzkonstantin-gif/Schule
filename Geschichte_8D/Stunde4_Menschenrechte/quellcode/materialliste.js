const fs = require("fs");
const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, BorderStyle, ShadingType, AlignmentType, VerticalAlign, Footer, PageNumber } = require("docx");

const NAVY = "1E2761", MUTED = "5B6B8C", LIGHT = "F4F7FD", LIGHT2 = "EAF0FA", RED = "D9534F";
const W = 10206;
const t = (text, o = {}) => new TextRun({ text, font: o.font || "Calibri", size: o.size || 21, bold: o.bold, italics: o.italics, color: o.color || "000000" });
const p = (runs, o = {}) => new Paragraph({ children: Array.isArray(runs) ? runs : [runs], alignment: o.align, spacing: { before: o.before ?? 0, after: o.after ?? 60 }, keepNext: o.keepNext });
const solid = (c, s = 8) => ({ style: BorderStyle.SINGLE, size: s, color: c });
const all = (b) => ({ top: b, bottom: b, left: b, right: b });
const LINE = solid("5B6B8C", 8);
const cell = (children, o = {}) => new TableCell({ children: Array.isArray(children) ? children : [children], width: { size: o.w, type: WidthType.DXA }, shading: o.fill ? { fill: o.fill, type: ShadingType.CLEAR, color: "auto" } : undefined, borders: o.borders || all(LINE), verticalAlign: o.valign || VerticalAlign.CENTER, margins: { top: o.m ?? 70, bottom: o.m ?? 70, left: 110, right: 110 } });
const row = (cells, h) => new TableRow({ children: cells, height: h ? { value: h, rule: "atLeast" } : undefined, cantSplit: true });
const table = (widths, rows) => new Table({ width: { size: widths.reduce((a, b) => a + b, 0), type: WidthType.DXA }, columnWidths: widths, rows });
const h2 = (s) => p(t(s, { font: "Cambria", size: 26, bold: true, color: NAVY }), { before: 200, after: 80, keepNext: true });
const footer = new Footer({ children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ font: "Calibri", size: 16, color: MUTED, children: ["Geschichte 8D · Materialliste Menschenrechte · Mi 25.11.2026 · Seite ", PageNumber.CURRENT] })] })] });

const hdr = (cols, widths) => row(cols.map((x, i) => cell(p(t(x, { bold: true, color: "FFFFFF", size: 19 }), { after: 0 }), { w: widths[i], fill: NAVY, borders: all(solid(NAVY, 8)) })), 320);
const box = () => p(t("□", { size: 26 }), { after: 0, align: AlignmentType.CENTER });

// Druck
const DW = [550, 3500, 1700, 4456];
const druck = [
  ["Arbeitsblatt „Eine Quelle auswerten“ (Seite 1: Quelle und Tabelle, Seite 2: Zusatz, Urteil, Rückmeldung)", "26 Blatt (24 + 2 Reserve), doppelseitig, eine pro Person", "Arbeitsblatt_Menschenrechte_Quelle_8D.pdf, 26-mal drucken, doppelseitig (Wenden an der kurzen Seite)"],
  ["Lösung und Musterantworten (für dich)", "1 Blatt (zwei Seiten)", "Loesung_Arbeitsblatt_Menschenrechte_8D.pdf"],
];
const prep = [
  ["Tafelskript ausdrucken oder auf dem Tablet bereithalten", "Tafelskript_Stunde4_Menschenrechte_8D.pdf, 7 Seiten"],
  ["Folien öffnen und einmal durchklicken", "Praesentation_Stunde4_Menschenrechte_8D.pptx, 20 Folien (27 Punkt); Folie 1 (Karikatur mit Sprechblase und zwei Bildern an der Wand) steht beim Betreten da"],
  ["Arbeitsblätter auf die Tische legen (umgedreht)", "Pro Person ein Blatt, doppelseitig, nach unten. Erst bei Folie 10 umdrehen lassen"],
  ["Wortlaut der Artikel mit Lehrbuch und Kollegin abgleichen", "Gekürzte, vereinfachte Fassung (Artikel 1, 2, 3, 4, 6, 10, 11), Wortlaut für den Test einheitlich halten"],
  ["Hintergrund zum 4. August im Lehrbuch lesen", "abgeschaffte Vorrechte, Ablösung der Abgaben, Bauernaufstände"],
  ["Rückmeldung der Hospitation absprechen", "Schwerpunkt: Aktivierung mit Fokus auf Kompetenzerwerb (Quelle auswerten); Punkte folgen später"],
  ["Kreide, Grün zum Verbessern, Uhr oder Timer", "Die Übungsphase endet 5 Minuten vor Schluss, „Auswerten“ bleibt immer"],
];
const tafel = [
  ["Seitentafel", "Stundenfrage („Was will der Dritte Stand noch? Was steht dazu in der Erklärung der Menschenrechte?“), Ziele („Ich kann …“, 3 Sätze, Satz 3 mit *) und Ablauf: A Karikatur · V Rückblick · I Quelle auswerten · V Blatt, Urteil · A Antwort"],
  ["Mitte", "leer lassen"],
];

const d = [];
d.push(p(t("GESCHICHTE · KLASSE 8D · MITTWOCH 25.11.2026 · HOSPITATION", { size: 18, bold: true, color: MUTED }), { after: 20 }));
d.push(p(t("Materialliste: Menschenrechte", { font: "Cambria", size: 44, bold: true, color: NAVY }), { after: 40 }));
d.push(p(t("Zum Abhaken. Alle Dateien liegen in Geschichte_8D/Stunde4_Menschenrechte/. Ablauf (45 Minuten): 4 Karikatur · 5 Rückblick · 9 Quelle auswerten (Vormachen) · 20 Arbeitsblatt und Urteil · 7 Auswerten.", { size: 20, color: MUTED }), { after: 80 }));

d.push(h2("1 Drucken"));
d.push(table(DW, [hdr(["", "Was", "Menge", "Datei und Hinweis"], DW), ...druck.map(([a, b, c]) => row([cell(box(), { w: DW[0] }), cell(p(t(a, { bold: true, size: 20 }), { after: 0 }), { w: DW[1] }), cell(p(t(b, { size: 20 }), { after: 0 }), { w: DW[2] }), cell(p(t(c, { size: 19 }), { after: 0 }), { w: DW[3] })], 560))]));

d.push(h2("2 Vorbereiten"));
const PW = [550, 4100, 5556];
d.push(table(PW, [hdr(["", "Aufgabe", "Hinweis"], PW), ...prep.map(([a, b]) => row([cell(box(), { w: PW[0] }), cell(p(t(a, { bold: true, size: 20 }), { after: 0 }), { w: PW[1] }), cell(p(t(b, { size: 19 }), { after: 0 }), { w: PW[2] })], 520))]));

d.push(h2("3 Tafel vor der Stunde"));
const TW = [2000, 8206];
d.push(table(TW, [hdr(["Wo", "Was steht dort"], TW), ...tafel.map(([a, b]) => row([cell(p(t(a, { bold: true, size: 20 }), { after: 0 }), { w: TW[0], fill: LIGHT }), cell(p(t(b, { size: 20 }), { after: 0 }), { w: TW[1] })], 520))]));

d.push(h2("4 Kurz vor Beginn (5 Minuten)"));
["Folie 1 (Karikatur) steht, Beamer an", "Arbeitsblätter liegen umgedreht auf den Tischen", "Seitentafel beschrieben, Mitte leer", "Lösungsblatt griffbereit", "Ansagen im Kopf: „Stifte liegen. Ihr hört nur zu.“ · „Flüsterstufe. Nur dein Partner hört dich.“ · „Vergleicht und verbessert mit Grün.“"].forEach((x) => d.push(p([t("□  ", { size: 24 }), t(x, { size: 20 })], { after: 50 })));

(async () => {
  const doc = new Document({ styles: { default: { document: { run: { font: "Calibri", size: 21 } } } }, sections: [{ properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 800, bottom: 700, left: 850, right: 850 } } }, footers: { default: footer }, children: d }] });
  fs.writeFileSync("../Materialliste_Stunde4_Menschenrechte_8D.docx", await Packer.toBuffer(doc));
  console.log("ok");
})();
