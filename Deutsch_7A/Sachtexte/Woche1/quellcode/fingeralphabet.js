// Fingeralphabet DGS: Schülerblatt (2 pro A4 quer, zum Halbieren) + Beispielfolien (pptx)
// Bildvorlage: von der Lehrkraft bereitgestellt (bilder/DGS_Fingeralphabet.png), Einzelbuchstaben per ImageMagick ausgeschnitten (bilder/einzeln/).
// Aufruf: node fingeralphabet.js
const fs = require("fs");
const path = require("path");
const { Document, Packer, Paragraph, TextRun, ImageRun, Table, TableRow, TableCell, WidthType, BorderStyle, AlignmentType, VerticalAlign, HeightRule, PageOrientation } = require("docx");
const pptxgen = require("pptxgenjs");
const B = path.join(__dirname, "..", "bilder");
const NAVY = "1E2761", MUTED = "5B6B8C", RED = "D9534F", GREEN = "2E9E6B", GRAY = "9AA3B5", LIGHT = "F4F7FD", LIGHT2 = "EAF0FA", WHITE = "FFFFFF", ICE = "CADCFC", NAVY2 = "24306E", VIO = "7A5BA6";

// ---------- Schülerblatt ----------
const t = (text, o = {}) => new TextRun({ text, font: o.font || "Calibri", size: o.size || 20, bold: o.bold, italics: o.it, color: o.color || "000000" });
const p = (ch, o = {}) => new Paragraph({ children: Array.isArray(ch) ? ch : [ch], alignment: o.align, spacing: { before: o.before || 0, after: o.after ?? 50 }, indent: o.indent });
const NB = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
const NONE = { top: NB, bottom: NB, left: NB, right: NB };
const bs = (x) => ({ top: x, bottom: x, left: x, right: x });
const cell = (ch, w, o = {}) => new TableCell({ children: ch, width: { size: w, type: WidthType.DXA }, borders: o.borders || NONE, verticalAlign: o.valign || VerticalAlign.TOP, margins: { top: o.m ?? 0, bottom: o.m ?? 0, left: o.ml ?? 0, right: o.ml ?? 0 } });
const table = (cols, rows) => new Table({ columnWidths: cols, width: { size: cols.reduce((a, b) => a + b, 0), type: WidthType.DXA }, rows });
const IMG = fs.readFileSync(path.join(B, "DGS_Fingeralphabet.png"));
const PAGEW = 16838, MARG = 400, GAP = 500;
const CW = Math.floor((PAGEW - 2 * MARG - GAP) / 2);
function copy() {
  const imgH = 470, imgW = Math.round(imgH * 462 / 680); // px (≈ 12,4 cm hoch)
  const boxes = table(Array(10).fill(Math.floor((CW - 400) / 10)), [new TableRow({ height: { value: 620, rule: HeightRule.EXACT }, children: Array(10).fill(0).map(() => cell([p(t(""))], Math.floor((CW - 400) / 10), { borders: bs({ style: BorderStyle.SINGLE, size: 6, color: GRAY }) })) })]);
  return table([CW], [new TableRow({ children: [cell([
    p(t("Das Fingeralphabet der Deutschen Gebärdensprache (DGS)", { font: "Cambria", bold: true, size: 26, color: NAVY }), { after: 30 }),
    p(t("Mit dem Fingeralphabet buchstabiert man Namen und Wörter, für die es keine eigene Gebärde gibt. Man gebärdet mit einer Hand, die Handfläche zeigt meist zum Gegenüber.", { size: 18, color: MUTED }), { after: 80 }),
    p(new ImageRun({ type: "png", data: IMG, transformation: { width: imgW, height: imgH } }), { align: AlignmentType.CENTER, after: 60 }),
    p([t("Bewegung: ", { bold: true, size: 18, color: RED }), t("Z wird in die Luft gezeichnet. Bei Ä, Ö, Ü bewegt sich die Hand kurz nach unten. SCH ist ein eigenes Zeichen.", { size: 18 })], { after: 80 }),
    p([t("Auftrag: ", { bold: true, size: 20, color: NAVY }), t("Buchstabiere deinen Vornamen. Schreibe ihn zuerst in die Kästchen – ein Buchstabe pro Kästchen (SCH, Ä, Ö, Ü zählen als ein Zeichen).", { size: 19 })], { after: 60 }),
    boxes,
  ], CW, { borders: bs({ style: BorderStyle.DASHED, size: 6, color: GRAY }), m: 140, ml: 200 })] })]);
}
const cut = cell([p(t("✂", { color: GRAY, size: 18 }), { align: AlignmentType.CENTER })], GAP, { valign: VerticalAlign.CENTER });
const props = { page: { size: { width: 11906, height: 16838, orientation: PageOrientation.LANDSCAPE }, margin: { top: 350, bottom: 250, left: MARG, right: MARG } } };
Packer.toBuffer(new Document({ styles: { default: { document: { run: { font: "Calibri", size: 20 } } } }, sections: [{ properties: props, children: [table([CW, GAP, CW], [new TableRow({ children: [cell([copy()], CW), cut, cell([copy()], CW)] })])] }] }))
  .then((buf) => fs.writeFileSync(path.join(__dirname, "..", "AB_Fingeralphabet_DGS.docx"), buf));

// ---------- Beispielfolien ----------
const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE";
pres.title = "Fingeralphabet – Beispiel";
const HEAD = "Cambria", BODY = "Calibri";
let pageNo = 0;
const txt = (s, text, o) => s.addText(text, Object.assign({ isTextBox: true, fontFace: BODY, fontSize: 16, color: NAVY, margin: 0, valign: "top" }, o));
const box = (s, x, y, w, h, fill, o = {}) => s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, fill: { color: fill }, line: o.line ? { color: o.line, width: o.lw || 1.5 } : { type: "none" }, rectRadius: o.r ?? 0.08 });
function base(kicker, title, notes) {
  const s = pres.addSlide(); pageNo++;
  s.background = { color: WHITE };
  txt(s, kicker.toUpperCase(), { x: 0.6, y: 0.4, w: 9, h: 0.3, fontSize: 12, color: MUTED, charSpacing: 2 });
  txt(s, title, { x: 0.6, y: 0.72, w: 12.1, h: 0.7, fontFace: HEAD, fontSize: 32, bold: true });
  txt(s, "Deutsch 7A · Sachtexte · Gebärdensprache", { x: 0.6, y: 7.0, w: 7, h: 0.25, fontSize: 10, color: MUTED });
  if (notes) s.addNotes(notes);
  return s;
}
const L = (k) => path.join(B, "einzeln", k + ".png");
const CELLR = 92 / 113; // Seitenverhältnis eines Ausschnitts (B/H)

// Folie 1: das ganze Alphabet
{
  const s = base("Gebärdensprache", "Das Fingeralphabet (DGS)", "Kurz erklären: Das Fingeralphabet ist nur ein kleiner Teil der Gebärdensprache. Man nutzt es für Namen und Wörter ohne eigene Gebärde. Mit einer Hand, Handfläche zum Gegenüber. Z wird in die Luft gezeichnet, Ä/Ö/Ü mit kurzer Bewegung nach unten, SCH ist ein eigenes Zeichen.");
  const h = 5.6, w = h * 462 / 680;
  s.addImage({ path: path.join(B, "DGS_Fingeralphabet.png"), x: 0.6, y: 1.45, w, h });
  const x0 = 0.6 + w + 0.5;
  [["Wofür?", "Namen und Wörter buchstabieren, für die es keine eigene Gebärde gibt."], ["Wie?", "Mit einer Hand, Handfläche zum Gegenüber, Hand etwa auf Schulterhöhe."], ["Achtung", "Z wird in die Luft gezeichnet. Ä, Ö, Ü: Hand kurz nach unten bewegen. SCH ist ein eigenes Zeichen."]].forEach(([a, b], i) => {
    const y = 1.6 + i * 1.75;
    box(s, x0, y, 12.7 - x0, 1.5, i === 2 ? WHITE : LIGHT, i === 2 ? { line: RED, lw: 2 } : {});
    txt(s, a, { x: x0 + 0.25, y: y + 0.15, w: 12.2 - x0, h: 0.4, fontSize: 19, bold: true, color: i === 2 ? RED : NAVY });
    txt(s, b, { x: x0 + 0.25, y: y + 0.6, w: 12.2 - x0, h: 0.8, fontSize: 17 });
  });
}
// Beispielfolien
function beispiel(word, keys, labels, notes, title) {
  const s = base("Gebärdensprache · Beispiel", title, notes);
  const n = keys.length, gap = 0.35, maxW = (12.1 - gap * (n - 1)) / n;
  const h = Math.min(3.6, maxW / CELLR), w = h * CELLR;
  const total = n * w + (n - 1) * gap, x0 = 0.6 + (12.1 - total) / 2;
  keys.forEach((k, i) => {
    const x = x0 + i * (w + gap);
    box(s, x - 0.08, 1.6, w + 0.16, h + 0.16, WHITE, { line: "C9D3EA" });
    s.addImage({ path: L(k), x, y: 1.68, w, h });
    txt(s, labels[i], { x, y: 1.85 + h, w, h: 0.6, fontSize: 30, bold: true, color: RED, align: "center", fontFace: HEAD });
    if (i < n - 1) txt(s, "→", { x: x + w, y: 1.6 + h / 2 - 0.25, w: gap, h: 0.5, fontSize: 22, bold: true, color: MUTED, align: "center" });
  });
  box(s, 0.6, 6.05, 12.1, 0.65, LIGHT2);
  txt(s, [{ text: word + "  ", options: { bold: true, color: RED } }, { text: "Zeige die Zeichen nacheinander, langsam und mit einer kurzen Pause zwischen den Buchstaben." }], { x: 0.85, y: 6.17, w: 11.7, h: 0.45, fontSize: 17 });
}
beispiel("L – E – A", ["L", "E", "A"], ["L", "E", "A"], "Vormachen: Die Lehrkraft buchstabiert L-E-A langsam vor, die Klasse macht mit. Dann fragen: Wie viele Zeichen sind es? (3)", "So buchstabiert man einen Namen: L-E-A");
beispiel("SCH – U – L – E", ["SCH", "U", "L", "E"], ["SCH", "U", "L", "E"], "Zweites Beispiel mit SCH: SCH ist ein Zeichen, nicht drei. Die Klasse buchstabiert gemeinsam mit.", "Achtung: SCH ist ein Zeichen – S-C-H nicht einzeln");
// Auftrag
{
  const s = base("Gebärdensprache · Ü", "Ü  Buchstabiere deinen Namen", "Ansage: „Jetzt arbeitet ihr zu zweit. Erst allein üben, dann zeigt ihr euren Namen. Der Partner liest mit und schreibt den Namen auf. Stimmt er? Dann wechseln.“");
  [["1", "Schreibe deinen Vornamen in die Kästchen auf dem Blatt."], ["2", "Suche zu jedem Buchstaben das Zeichen und übe es langsam."], ["3", "Zeige deinem Partner deinen Namen. Er liest mit und schreibt ihn auf."], ["4", "Vergleicht: Stimmt der Name? Dann wechselt."]].forEach(([n, x], i) => {
    const y = 1.7 + i * 0.85;
    s.addShape(pres.shapes.OVAL, { x: 0.6, y, w: 0.5, h: 0.5, fill: { color: NAVY }, line: { type: "none" } });
    txt(s, n, { x: 0.6, y, w: 0.5, h: 0.5, fontSize: 16, bold: true, color: WHITE, align: "center", valign: "middle" });
    txt(s, x, { x: 1.3, y: y + 0.05, w: 11.3, h: 0.6, fontSize: 21, bold: true });
  });
  box(s, 0.6, 5.2, 12.1, 0.8, WHITE, { line: GREEN, lw: 2 });
  txt(s, [{ text: "* Für Schnelle: ", options: { bold: true, color: GREEN } }, { text: "Buchstabiere deine Sportart oder deinen Lieblingsverein – der Partner rät." }], { x: 0.85, y: 5.38, w: 11.7, h: 0.5, fontSize: 18 });
  box(s, 0.6, 6.25, 12.1, 0.55, LIGHT2);
  txt(s, [{ text: "Zeit: ", options: { bold: true } }, { text: "8 Minuten · erst allein, dann zu zweit      " }, { text: "Fertig? Hände runter.", options: { italic: true, color: MUTED } }], { x: 0.85, y: 6.33, w: 11.7, h: 0.4, fontSize: 16 });
}
pres.writeFile({ fileName: path.join(__dirname, "..", "Folien_Fingeralphabet_Beispiel.pptx") });
