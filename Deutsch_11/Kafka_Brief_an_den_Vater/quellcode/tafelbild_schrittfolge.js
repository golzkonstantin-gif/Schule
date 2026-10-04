// Tafelbild in Stichworten: gemeinsame Matrix zu Abschnitt A, danach Abschnitt C allein (eine Seite, A4 quer)
const fs = require("fs");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, BorderStyle,
  ShadingType, AlignmentType, VerticalAlign, HeightRule, PageOrientation,
} = require("docx");

const MUTED = "5B6B8C";
const BOARD = "2E4A3B", BOARD2 = "3A5A48", CHALK = "F3F1EA", DIM = "C9C3A8", YEL = "F5D76E";
const COL_A = "8EC3F0", COL_C = "F29A8E";
const W = 15338;

const t = (text, o = {}) => new TextRun({ text, font: o.font || "Calibri", size: o.size || 21, bold: o.bold, italics: o.italics, color: o.color || CHALK });
const p = (runs, o = {}) => new Paragraph({ children: Array.isArray(runs) ? runs : [runs], alignment: o.align, spacing: { before: o.before ?? 0, after: o.after ?? 20 } });
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

const cw = [1300, 4300, 5000, 4738];
const colHead = row([
  cell(p(t("Z.", { bold: true, color: YEL })), { w: cw[0], fill: BOARD2 }),
  cell(p(t("Handlung des Vaters", { bold: true, color: YEL })), { w: cw[1], fill: BOARD2 }),
  cell(p(t("Modell (Seite / Axiom)", { bold: true, color: YEL })), { w: cw[2], fill: BOARD2 }),
  cell(p(t("Wirkung auf den Sohn (Zitat)", { bold: true, color: YEL })), { w: cw[3], fill: BOARD2 }),
], 400);
const band = (label, sub, col, phaseMark) => row([cell(p([
  t(phaseMark + "  ", { bold: true, color: DIM, size: 20 }), t(label, { bold: true, size: 23, color: col }), t("   " + sub, { italics: true, size: 19, color: DIM }),
]), { w: W, span: 4, fill: BOARD2 })], 380);
const r = (z, h, m, w, col) => row([
  cell(p(t(z, { bold: true, color: col, size: 20 })), { w: cw[0] }),
  cell(p(t(h, { size: 20, italics: h.startsWith("„") })), { w: cw[1] }),
  cell(p(t(m, { size: 20 })), { w: cw[2] }),
  cell(p(t(w, { size: 20, italics: true })), { w: cw[3] }),
], 520);

const A = [
  ["4–5", "„kein Wort der Widerrede!“", "Appell „Schweig!“ + Beziehung „Ich bestimme“", "„stockende, stotternde Art des Sprechens“ (Z. 7)"],
  ["5", "erhobene Hand", "analoge Botschaft verstärkt die Drohung", "„schließlich schwieg ich“ (Z. 8)"],
  ["6–7", "selbst „ausgezeichneter Redner“", "starr komplementär: oben / unten", "„weder denken noch reden“ (Z. 9)"],
  ["15", "deutet Schweigen als „contra“", "Interpunktion → Missverständnis (Trotz ↔ Gehorsam)", "„verkroch mich vor Dir“ (Z. 12–13)"],
];
const C = [
  ["3–4", "„Das ist Dir wohl schon zu viel?“", "Ironie: Frage (Sache) ≠ Vorwurf (Beziehung)", "„schon bestraft, ehe man noch wußte …“ (Z. 6)"],
  ["5", "„bösem Lachen und bösem Gesicht“", "inkongruent: analog widerspricht digital", ""],
  ["9–10", "spricht „zur Mutter“: „vom Herrn Sohn“", "Empfänger übergangen · Abwertung", "„nicht einmal des bösen Ansprechens gewürdigt“ (Z. 8)"],
  ["2", "„Deiner Überlegenheit über mich“", "komplementär – Sohn übernimmt das Muster", "fragt nur die Mutter: „Wie geht es dem Vater?“ (Z. 15)"],
];

const board = new Table({
  width: { size: W, type: WidthType.DXA }, columnWidths: cw,
  rows: [
    row([cell([
      p(t("Franz Kafka: „Brief an den Vater“ (1919)", { font: "Cambria", size: 32, bold: true }), { after: 40 }),
      p([t("These: ", { bold: true, size: 23, color: YEL }), t("Kafkas „Furcht“ gründet auf Kommunikationsproblemen.", { size: 23 })], { after: 50 }),
      p([t("So analysiert man: ", { bold: true, size: 20, color: YEL }),
        t("1 Verstehen", { bold: true, size: 20 }), t(" (Wer? An wen? Worum?)  →  ", { size: 19, color: DIM }),
        t("2 Untersuchen", { bold: true, size: 20 }), t(" (Handlung · Modell · Wirkung)  →  ", { size: 19, color: DIM }),
        t("3 Schreiben", { bold: true, size: 20 }), t(" (Behauptung → Beleg → Erklärung)", { size: 19, color: DIM })]),
    ], { w: W, span: 4, borders: { top: none, bottom: line, left: none, right: none }, m: 110 })]),
    band("A  Redeverbot – gemeinsam im Unterrichtsgespräch", "„ich verlernte das Reden“ (S. 5 f.) · Matrix in den Hefter übernehmen", COL_A, "①"),
    colHead,
    ...A.map(([z, h, m, w]) => r(z, h, m, w, COL_A)),
    band("C  Ironie – allein", "„Ein besonderes Vertrauen hattest Du zur Erziehung durch Ironie …“ (S. 6)", COL_C, "②"),
    ...C.map(([z, h, m, w]) => r(z, h, m, w, COL_C)),
    row([cell([
      p([t("③ These geprüft:  ", { bold: true, size: 22, color: YEL }), t("✓ stützt: Wo Widerspruch unmöglich ist, entsteht Furcht.   + ergänzt: „Stärke … Schwäche“ – „Du verstärktest nur, was war“", { size: 20 })], { after: 30 }),
      p([t("→ Der Brief selbst ist Metakommunikation – und erreicht den Vater nie.", { size: 20, bold: true })]),
    ], { w: W, span: 4, fill: BOARD2, borders: { top: line, bottom: line, left: none, right: none }, m: 100 })], 520),
    row([cell([
      p([t("④ Auftrag: Tafelbild → Fließtext", { bold: true, size: 23, color: YEL }), t("   Schreibe deine Analyse von A und C als zusammenhängenden Text.", { size: 20 })], { after: 40 }),
      p([t("Aufbau:  ", { bold: true, size: 20, color: YEL }), t("Einleitung (Wer? An wen? Worum?)  →  je Zeile der Matrix: Behauptung → Beleg (Zitat + Z.) → Erklärung  →  Schluss: These", { size: 20 })], { after: 40 }),
      p([t("Satzbausteine:  ", { bold: true, size: 20, color: YEL }), t("Indem der Vater …, sendet er …  ·  Dies zeigt sich in … (Z. …)  ·  Hinzu kommt, dass …  ·  Die Folge ist …  ·  Insgesamt stützt die Stelle die These, weil …", { size: 20, italics: true })]),
    ], { w: W, span: 4, fill: BOARD, borders: { top: none, bottom: none, left: none, right: none }, m: 110 })], 900),
  ],
});

const doc = new Document({
  styles: { default: { document: { run: { font: "Calibri", size: 20 } } } },
  sections: [{
    properties: { page: { size: { width: 11906, height: 16838, orientation: PageOrientation.LANDSCAPE }, margin: { top: 500, bottom: 400, left: 750, right: 750 } } },
    children: [
      new Paragraph({ spacing: { after: 60 }, children: [
        new TextRun({ text: "TAFELBILD IN STICHWORTEN · ", font: "Calibri", size: 17, bold: true, color: MUTED }),
        new TextRun({ text: "① Abschnitt A: im Unterrichtsgespräch, Matrix wächst an der Tafel und im Hefter · ② Abschnitt C: allein, dann an der Tafel zusammentragen · ③ These prüfen · ④ Matrix in Fließtext überführen", font: "Calibri", size: 17, color: MUTED }),
      ] }),
      board,
    ],
  }],
});
Packer.toBuffer(doc).then((b) => fs.writeFileSync("Tafelbild_Schrittfolge.docx", b));
