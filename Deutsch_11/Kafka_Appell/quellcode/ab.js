// Arbeitsblatt „Die Ebene des Appells“ nach dem Muster der Mentorin: Kontext, Textauszug, Information, Arbeitsaufträge 1–5 (2 Seiten, doppelseitig)
const fs = require("fs");
const path = require("path");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, BorderStyle,
  ShadingType, AlignmentType, VerticalAlign, HeightRule, Header, Footer, TabStopType, PageNumber,
} = require("docx");

const NAVY = "1E2761", MUTED = "5B6B8C", LIGHT = "F4F7FD", LIGHT2 = "EAF0FA", RED = "D9534F";
const COL_A = "3F7CC4", COL_C = "D9534F";
const W = 10206;
const { TEXT, QUELLE, GLOSSAR, KONTEXT, KONTEXT_ZITAT, APPELL, LEITFRAGEN, FORMULIERUNGSHILFEN, REDE_SITUATION, REDE_AUFTRAG, REDE_UMFANG, KRITERIEN } = require("./daten");
const FS = 24;

const t = (text, o = {}) => new TextRun({ text, font: o.font || "Calibri", size: o.size || 22, bold: o.bold, italics: o.italics, color: o.color || "000000" });
const p = (runs, o = {}) => new Paragraph({ children: Array.isArray(runs) ? runs : [runs], alignment: o.align, pageBreakBefore: o.pb, keepNext: o.keepNext, spacing: { before: o.before ?? 0, after: o.after ?? 100, line: o.line } });
const none = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
const solid = (c, s = 6) => ({ style: BorderStyle.SINGLE, size: s, color: c });
const all = (b) => ({ top: b, bottom: b, left: b, right: b });
const cell = (children, o = {}) => new TableCell({
  children: Array.isArray(children) ? children : [children], width: { size: o.w, type: WidthType.DXA },
  shading: o.fill ? { fill: o.fill, type: ShadingType.CLEAR, color: "auto" } : undefined,
  borders: o.borders || all(solid("B9C6E8")), verticalAlign: o.valign || VerticalAlign.TOP,
  margins: { top: o.m ?? 50, bottom: o.m ?? 50, left: o.ml ?? 100, right: o.ml ?? 100 }, columnSpan: o.span,
});
const table = (widths, rows) => new Table({ width: { size: widths.reduce((a, b) => a + b, 0), type: WidthType.DXA }, columnWidths: widths, rows });
const row = (cells, h) => new TableRow({ children: cells, height: h ? { value: h, rule: HeightRule.ATLEAST } : undefined, cantSplit: true });
const box = (paras, fill = LIGHT2) => table([W], [row([cell(paras, { w: W, fill, borders: all(none), m: 120, ml: 160 })])]);

// Text mit Zeilennummern (jede Zeile nummeriert, feste Umbrüche); Worterklärungen als hochgestellte Fußnoten
// glossar: [Wort, Erklärung] – Fußnotenziffer hinter dem ersten Vorkommen (ganzes Wort); Einträge ohne Treffer als Hinweis ohne Ziffer
function textBlock(lines, glossar, col, start = 1) {
  const nW = 520, tW = W - nW;
  const marks = lines.map(() => []);
  const notes = [];
  // Fundstellen ermitteln und nach Position im Text sortieren, damit die Ziffern aufsteigend erscheinen
  const found = glossar.map(([w, e]) => {
    const li = /^\w/.test(w) ? lines.findIndex((l) => l.includes(w)) : -1;
    if (li < 0) return { w, e, li: Infinity };
    let end = lines[li].indexOf(w) + w.length;
    while (end < lines[li].length && /[\p{L}]/u.test(lines[li][end])) end++;
    return { w, e, li, end };
  }).sort((a, b) => a.li - b.li || a.end - b.end);
  let n = start;
  found.forEach(({ w, e, li, end }) => {
    if (li === Infinity) { notes.push({ w, e }); return; }
    marks[li].push({ end, n });
    notes.push({ w, e, n: n++ });
  });
  const lineRuns = (l, ms) => {
    const runs = []; let pos = 0;
    ms.sort((a, b) => a.end - b.end).forEach(({ end, n }) => {
      runs.push(t(l.slice(pos, end), { font: "Cambria", size: FS }));
      runs.push(new TextRun({ text: String(n), font: "Calibri", size: FS, bold: true, color: col, superScript: true }));
      pos = end;
    });
    runs.push(t(l.slice(pos), { font: "Cambria", size: FS }));
    return runs;
  };
  const rows = lines.map((l, i) => row([
    cell(p(t(String(i + 1), { size: 16, color: MUTED }), { align: AlignmentType.RIGHT, after: 0 }), { w: nW, borders: { top: none, bottom: none, left: none, right: solid(col, 10) }, m: 6, ml: 80, valign: VerticalAlign.CENTER }),
    cell(p(lineRuns(l, marks[i]), { after: 0, line: 360 }), { w: tW, borders: all(none), m: 6, ml: 140 }),
  ]));
  notes.forEach(({ w, e, n }, i) => rows.push(row([cell(p([
    n ? new TextRun({ text: String(n), font: "Calibri", size: 19, bold: true, color: col, superScript: true }) : t("", { size: 19 }),
    t((n ? " " : "") + w + " ", { size: 19, bold: true, color: col }), t("= " + e, { size: 19, color: MUTED }),
  ], { after: 0 }), { w: W, span: 2, borders: { top: i ? none : solid("D5DCEC", 6), bottom: none, left: none, right: none }, m: i ? 0 : 50, ml: 140 })])));
  return table([nW, tW], rows);
}
const header = new Header({ children: [new Paragraph({ tabStops: [{ type: TabStopType.RIGHT, position: W }], children: [t("Deutsch · Klasse 11 · Kafka, „Brief an den Vater“", { size: 17, color: MUTED }), t("\tName: ______________________   Datum: ___________", { size: 17, color: MUTED })] })] });
const footer = new Footer({ children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ font: "Calibri", size: 16, color: MUTED, children: [QUELLE + " · DigBib-Ausgabe"] })] })] });

const h2 = (text, o = {}) => p(t(text, { font: "Cambria", size: 26, bold: true, color: NAVY }), { before: o.before ?? 200, after: 80, pb: o.pb, keepNext: true });
const aufgabe = (nr, titel, form) => p([t("Aufgabe " + nr + ": " + titel, { bold: true, color: NAVY, size: 22 }), t(form ? "   (" + form + ")" : "", { italics: true, color: MUTED, size: 19 })], { before: 160, after: 50, keepNext: true });
const bullet = (runs) => p([t("•  ", { size: 21, color: NAVY }), ...(Array.isArray(runs) ? runs : [runs])], { after: 40 });

// Kommunikationsquadrat zum Ausfüllen (Fokus Appell)
const qW = W / 2;
const quad = (label, frage, focus) => cell([
  p(t("[ " + label + " ]", { bold: true, size: 20, color: focus ? RED : NAVY }), { after: 10 }),
  p(t(frage, { italics: true, size: 17, color: MUTED }), { after: 0 }),
], { w: qW, fill: focus ? "FDECEA" : "FFFFFF", borders: all(solid(focus ? RED : "B9C6E8", focus ? 12 : 6)), m: 80, ml: 140 });
const quadrat = table([qW, qW], [
  row([quad("SACHINHALT", "Worüber spricht Kafka?"), quad("SELBSTKUNDGABE", "Was zeigt er von sich?")], 1250),
  row([quad("BEZIEHUNG", "Wie steht er zum Vater?"), quad("APPELL", "Was soll der Vater tun? offen / verdeckt", true)], 1650),
]);

const seite1 = [
  p(t("Arbeitsblatt: Die Ebene des Appells in Kafkas „Brief an den Vater“", { font: "Cambria", size: 32, bold: true, color: NAVY }), { after: 60 }),
  p([t("Thema: ", { bold: true, size: 20 }), t("Das Vier-Ohren-Modell nach Friedemann Schulz von Thun mit Fokus auf den ", { size: 20 }), t("Appell", { bold: true, size: 20 }), t(" (Wozu will ich dich bringen?)", { size: 20 })], { after: 30 }),
  p([t("Ziel: ", { bold: true, size: 20 }), t("Analyse der offenen und verdeckten Appelle, die Franz Kafka an seinen Vater richtet.", { size: 20 })], { after: 160 }),
  box([
    p(t("Kontext-Box (Brücke zur Biografie)", { bold: true, color: RED, size: 22 }), { after: 50 }),
    p(t(KONTEXT, { size: 21 }), { after: 80, line: 290 }),
    p(t("„" + KONTEXT_ZITAT[0] + "“", { font: "Cambria", italics: true, size: 21 }), { after: 10 }),
    p(t("– " + KONTEXT_ZITAT[1], { size: 18, color: MUTED }), { after: 0 }),
  ]),
  h2("1. Textauszug"),
  p(t("Rechtschreibung vor 1996: daß = dass, mußte = musste · […] = Auslassung", { size: 17, italics: true, color: MUTED }), { after: 100 }),
  textBlock(TEXT, GLOSSAR, COL_A),
  p(t(QUELLE, { size: 17, color: MUTED }), { before: 80, align: AlignmentType.RIGHT, after: 0 }),
  h2("2. Information: Das „Vier-Ohren-Modell“ (Fokus Appell)"),
  p(t("Jede Äußerung enthält vier Botschaften gleichzeitig: Sachinhalt, Selbstkundgabe, Beziehung und Appell. Das Appell-Ohr achtet darauf, wozu der Sprecher den anderen bringen will – was er tun, denken oder fühlen soll.", { size: 21 }), { after: 80, line: 290 }),
  table([2200, W - 2200], APPELL.slice(1).map(([k, e]) => row([
    cell(p(t(k === "offen" ? "Offener Appell" : "Verdeckter Appell", { bold: true, color: k === "offen" ? NAVY : RED, size: 21 }), { after: 0 }), { w: 2200, fill: LIGHT }),
    cell(p(t(e.charAt(0).toUpperCase() + e.slice(1), { size: 21 }), { after: 0 }), { w: W - 2200 }),
  ]))),
];

const seite2 = [
  h2("3. Arbeitsaufträge", { pb: true, before: 0 }),
  p(t("Bearbeite die Aufgaben in deinem Hefter.", { size: 19, italics: true, color: MUTED }), { after: 40 }),
  aufgabe(1, "Vorarbeit (Textverständnis)", "allein"),
  p(t("Fasse den Textauszug in einem Satz zusammen: Worum bittet Kafka seinen Vater?", { size: 21 }), { after: 40 }),
  aufgabe(2, "Analyse des Appells", "allein"),
  p(t("Untersuche den Auszug gezielt mit dem Appell-Ohr. Beantworte die Leitfragen und belege deine Aussagen mit Zitaten und Zeilenangaben:", { size: 21 }), { after: 50 }),
  ...LEITFRAGEN.map(([frage, tipp]) => bullet([t(frage + " ", { bold: true, size: 21 }), t("(" + tipp + ")", { size: 19, color: MUTED })])),
  box([p([t("Formulierungshilfen:  ", { bold: true, color: NAVY, size: 19 }), t(FORMULIERUNGSHILFEN.join("  ·  "), { size: 19, italics: true })], { after: 0 })], "FFF8EC"),
  aufgabe(3, "Visualisierung"),
  p(t("Trage deine Ergebnisse in das Kommunikationsquadrat ein. Konzentriere dich vor allem auf den Quadranten Appell.", { size: 21 }), { after: 80 }),
  quadrat,
  aufgabe(4, "Reflexion"),
  p(t("Kafka hat den Brief seinem Vater nie übergeben. Warum appelliert er so vorsichtig und versteckt? Hätte der Vater die Bitte verstanden?", { size: 21 }), { after: 40 }),
  aufgabe(5, "Transfer: Kafkas Rede an den Vater", "allein"),
  p([t("Situation: ", { bold: true, size: 21 }), t(REDE_SITUATION, { size: 21 })], { after: 30 }),
  p([t("Deine Aufgabe: ", { bold: true, size: 21 }), t(REDE_AUFTRAG + " Umfang: " + REDE_UMFANG.split(" · ")[0] + ".", { size: 21 })], { after: 50 }),
  ...KRITERIEN.map(([k, e]) => bullet([t(k + ": ", { bold: true, size: 20 }), t(e, { size: 20 })])),
];

const doc = new Document({
  styles: { default: { document: { run: { font: "Calibri", size: 22 } } } },
  sections: [{ properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 850, bottom: 750, left: 850, right: 850, header: 380, footer: 380 } } }, headers: { default: header }, footers: { default: footer }, children: [...seite1, ...seite2] }],
});
Packer.toBuffer(doc).then((b) => fs.writeFileSync("Arbeitsblatt_Kafkas_Appell.docx", b));
