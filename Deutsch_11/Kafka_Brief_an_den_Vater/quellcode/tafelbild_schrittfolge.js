// Tafelbild in Stichworten: Schrittfolge „Kommunikation analysieren“ am Modellabschnitt A (eine Seite, A4 quer)
const fs = require("fs");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, BorderStyle,
  ShadingType, AlignmentType, VerticalAlign, HeightRule, PageOrientation,
} = require("docx");

const NAVY = "1E2761", MUTED = "5B6B8C";
const BOARD = "2E4A3B", BOARD2 = "3A5A48", CHALK = "F3F1EA", DIM = "C9C3A8", YEL = "F5D76E";
const R = { B1: "8EC3F0", B2: "F5D76E", C: "F29A8E" }; // Kreidefarben der Reihen
const W = 15338;

const t = (text, o = {}) => new TextRun({ text, font: o.font || "Calibri", size: o.size || 20, bold: o.bold, italics: o.italics, color: o.color || CHALK });
const p = (runs, o = {}) => new Paragraph({ children: Array.isArray(runs) ? runs : [runs], alignment: o.align, spacing: { before: o.before ?? 0, after: o.after ?? 30 } });
const none = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
const line = { style: BorderStyle.SINGLE, size: 6, color: "6E8C7A" };
const cell = (children, o = {}) => new TableCell({
  children: Array.isArray(children) ? children : [children],
  width: { size: o.w, type: WidthType.DXA },
  shading: { fill: o.fill || BOARD, type: ShadingType.CLEAR, color: "auto" },
  borders: o.borders || { top: line, bottom: line, left: none, right: none },
  verticalAlign: o.valign || VerticalAlign.TOP,
  margins: { top: o.m ?? 50, bottom: o.m ?? 50, left: 140, right: 140 },
  columnSpan: o.span,
});
const row = (cells, h) => new TableRow({ children: cells, height: h ? { value: h, rule: HeightRule.ATLEAST } : undefined, cantSplit: true });

// Mehrere Stichwort-Zeilen in einer Zelle; Präfix „B1:“ usw. in Reihenfarbe
function lines(arr, size = 22) {
  return arr.map((s) => {
    const m = s.match(/^(B1|B2|C):\s*(.*)$/);
    if (m) return p([t(m[1] + "  ", { bold: true, size, color: R[m[1]] }), t(m[2], { size })]);
    return p(t(s, { size }));
  });
}

const steps = [
  ["1", "Einordnen", "Wer? An wen? Warum?",
    ["Brief 1919 · Sohn → Vater", "erklärt seine „Furcht“"],
    []],
  ["2", "Situation", "Worum geht es?",
    ["Redeverbot → Sohn verlernt das Sprechen"],
    ["B1: Drohen", "B2: Drohung als Prophezeiung", "C: Ironie"]],
  ["3", "Handeln markieren", "Was tut der Sender? (Verben)",
    ["V: verbietet · droht · hebt die Hand", "S: stottert · schweigt · verkriecht sich"],
    ["B1: droht · läuft schreiend um den Tisch", "C: spottet · spricht über die Mutter"]],
  ["4", "Modell zuordnen", "Seite? Axiom?",
    ["Appell „Schweig!“ + Beziehung „Ich bestimme“", "Hand = analoge Botschaft", "starr komplementär", "Interpunktion: „contra“ ↔ Gehorsam"],
    ["B1: inkongruent (droht, will nicht fassen)", "C: Ironie = inkongruent · Kritik über Dritte"]],
  ["5", "Wirkung belegen", "Zitat + Zeile!",
    ["stotternd (Z. 7) → schwieg (Z. 8) → verkroch mich (Z. 12)"],
    ["B2: „Vertrauen zu eigenem Tun“ verloren", "C: fragt nur noch die Mutter"]],
  ["6", "Deuten", "These?",
    ["✓ stützt: kein Widerspruch → Furcht", "+ ergänzt: „Stärke … Schwäche“ (Z. 16)"],
    ["B2: ⚠ „Du verstärktest nur, was war“"]],
  ["7", "Ausformulieren", "B → B → E",
    ["Behauptung → Beleg → Erklärung", "„Indem der Vater …, zeigt sich …“"],
    ["(Musterabsatz auf dem Whiteboard)"]],
];

const cw = [3500, 6738, 5100];
const head = row([
  cell([p(t("FAHRPLAN", { bold: true, size: 20, color: YEL })), p(t("Kommunikation analysieren", { size: 18, color: DIM }))], { w: cw[0], fill: BOARD2 }),
  cell([p(t("MODELL: Abschnitt A (Z. 1–16)", { bold: true, size: 20, color: YEL })), p(t("„ich verlernte das Reden“", { size: 18, italics: true, color: DIM }))], { w: cw[1], fill: BOARD2 }),
  cell([p(t("BELEGE DER REIHEN", { bold: true, size: 20, color: YEL })), p([t("B1 Wand  ", { size: 18, color: R.B1 }), t("B2 Mitte  ", { size: 18, color: R.B2 }), t("C Fenster", { size: 18, color: R.C })])], { w: cw[2], fill: BOARD2 }),
], 520);

const stepRows = steps.map(([n, name, frage, mod, bel]) => row([
  cell([p([t(n + "  ", { bold: true, size: 28, color: YEL }), t(name, { bold: true, size: 24 })]), p(t(frage, { size: 19, italics: true, color: DIM }))], { w: cw[0] }),
  cell(lines(mod), { w: cw[1] }),
  cell(bel.length ? lines(bel, 21) : [p(t(""))], { w: cw[2] }),
], 900));

const board = new Table({
  width: { size: W, type: WidthType.DXA }, columnWidths: cw,
  rows: [
    row([cell([
      p(t("Franz Kafka: „Brief an den Vater“ (1919)", { font: "Cambria", size: 32, bold: true }), { after: 40 }),
      p([t("These: ", { bold: true, size: 23, color: YEL }), t("Kafkas „Furcht“ gründet auf Kommunikationsproblemen.", { size: 23 })]),
    ], { w: W, span: 3, fill: BOARD, borders: { top: none, bottom: line, left: none, right: none }, m: 110 })]),
    head,
    ...stepRows,
    row([cell([
      p([t("These geprüft:  ", { bold: true, size: 22, color: YEL }), t("Kommunikation macht aus Unterschieden Furcht – das Machtgefälle gehört dazu.  → Der Brief = Metakommunikation.", { size: 21 })]),
    ], { w: W, span: 3, fill: BOARD2, borders: { top: line, bottom: none, left: none, right: none }, m: 100 })], 480),
  ],
});

const doc = new Document({
  styles: { default: { document: { run: { font: "Calibri", size: 20 } } } },
  sections: [{
    properties: { page: { size: { width: 11906, height: 16838, orientation: PageOrientation.LANDSCAPE }, margin: { top: 500, bottom: 400, left: 750, right: 750 } } },
    children: [
      new Paragraph({ spacing: { after: 60 }, children: [
        new TextRun({ text: "TAFELBILD IN STICHWORTEN · ", font: "Calibri", size: 17, bold: true, color: MUTED }),
        new TextRun({ text: "links: von der Klasse abschreiben lassen · Mitte: Ihr Modell an Abschnitt A · rechts: Ergebnisse der Reihen nach jedem Schritt", font: "Calibri", size: 17, color: MUTED }),
      ] }),
      board,
    ],
  }],
});
Packer.toBuffer(doc).then((b) => fs.writeFileSync("Tafelbild_Schrittfolge.docx", b));
