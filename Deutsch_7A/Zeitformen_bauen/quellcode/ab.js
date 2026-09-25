const fs = require("fs");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, BorderStyle,
  ShadingType, AlignmentType, VerticalAlign, HeightRule, PageBreak, Header, Footer, TabStopType,
  PageNumber,
} = require("docx");

const NAVY = "1E2761", MUTED = "5B6B8C", LIGHT = "F4F7FD", LIGHT2 = "EAF0FA";
const C = { hv: "D9534F", mv: "2E9E6B", p2: "E8A33D", inf: "3F7CC4" };
const W = 10206; // content width (A4, 1.5 cm margins)

const t = (text, o = {}) => new TextRun({ text, font: o.font || "Calibri", size: o.size || 22, bold: o.bold, italics: o.italics, color: o.color || "000000" });
const p = (runs, o = {}) => new Paragraph({ children: Array.isArray(runs) ? runs : [runs], alignment: o.align, spacing: { before: o.before ?? 0, after: o.after ?? 100, line: o.line }, keepNext: o.keepNext });
const h1 = (text) => p(t(text, { font: "Cambria", size: 40, bold: true, color: NAVY }), { after: 60 });
const kicker = (text) => p(t(text.toUpperCase(), { size: 18, color: MUTED, bold: true }), { after: 20 });
const h2 = (label, text) => p([t(label + "  ", { font: "Cambria", size: 26, bold: true, color: C.hv }), t(text, { font: "Cambria", size: 26, bold: true, color: NAVY })], { before: 160, after: 80, keepNext: true });
const line = () => p(t(" "), { after: 0 });

const none = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
const noBorders = { top: none, bottom: none, left: none, right: none };
const solid = (c = NAVY, s = 8) => ({ style: BorderStyle.SINGLE, size: s, color: c });
const allBorders = (b) => ({ top: b, bottom: b, left: b, right: b });
const dashed = { style: BorderStyle.DASHED, size: 8, color: "7A869E" };

function cell(children, o = {}) {
  return new TableCell({
    children: Array.isArray(children) ? children : [children],
    width: { size: o.w, type: WidthType.DXA },
    shading: o.fill ? { fill: o.fill, type: ShadingType.CLEAR, color: "auto" } : undefined,
    borders: o.borders || allBorders(solid("B9C6E8", 6)),
    verticalAlign: o.valign || VerticalAlign.CENTER,
    margins: { top: o.m ?? 60, bottom: o.m ?? 60, left: 100, right: 100 },
    columnSpan: o.span,
  });
}
function table(widths, rows) {
  return new Table({ width: { size: widths.reduce((a, b) => a + b, 0), type: WidthType.DXA }, columnWidths: widths, rows });
}
function row(cells, h, rule = HeightRule.EXACT) {
  return new TableRow({ children: cells, height: h ? { value: h, rule } : undefined, cantSplit: true });
}

// ---------- Kopf / Fuß ----------
const header = new Header({
  children: [new Paragraph({
    tabStops: [{ type: TabStopType.RIGHT, position: W }],
    children: [t("Deutsch · Klasse 7A · Zeitformen-Baukasten", { size: 17, color: MUTED }), t("\tName: ______________________   Datum: ___________", { size: 17, color: MUTED })],
  })],
});
const footer = (label) => new Footer({
  children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ font: "Calibri", size: 16, color: MUTED, children: [label + " · Seite ", PageNumber.CURRENT] })] })],
});

// ---------- Karten ----------
function cardGrid(words, cols, cardW, cardH, size, tag) {
  const rows = [];
  for (let i = 0; i < words.length; i += cols) {
    const cells = [];
    for (let j = 0; j < cols; j++) {
      const w = words[i + j];
      if (w === undefined) { cells.push(cell(p(t("")), { w: cardW, borders: noBorders })); continue; }
      cells.push(cell([
        p(t(w, { font: "Cambria", size, bold: true, color: "111111" }), { align: AlignmentType.CENTER, after: 0 }),
      ], { w: cardW, borders: allBorders(dashed), m: 0 }));
    }
    rows.push(row(cells, cardH));
  }
  return table(Array(cols).fill(cardW), rows);
}

const wordsA = ["habe", "gespielt", "kann", "spielte", "wird", "gelaufen", "musst",
  "ist", "geschrieben", "darf", "laufen", "hatten", "gekauft", "wollen",
  "bist", "verstanden", "sollte", "ging", "werden", "gegessen", "könnt",
  "war", "aufgeräumt", "mochte", "gehst", "hast", "gefahren", "müssen"];
const wordsB = ["habe", "gebacken", "backte", "werde", "backen", "muss",
  "gebacken", "hatte", "haben", "backe", "gebacken", "werde", "backen"];

// ---------- Farbcode ----------
function legend() {
  const items = [["Hilfsverb", "haben · sein · werden", "rot", C.hv], ["Modalverb", "können · müssen · dürfen · wollen · sollen · mögen", "grün", C.mv], ["Partizip II", "gespielt · gegangen · verstanden", "gelb", C.p2], ["Infinitiv", "backen · spielen · gehen", "blau", C.inf]];
  const ws = [2551, 2551, 2552, 2552];
  return table(ws, [row(items.map(([n, ex, farbe, col], i) => cell([
    p(t(n + " = " + farbe, { bold: true, color: "FFFFFF", size: 21 }), { after: 20, align: AlignmentType.CENTER }),
    p(t(ex, { color: "FFFFFF", size: 16 }), { after: 0, align: AlignmentType.CENTER }),
  ], { w: ws[i], fill: col, borders: allBorders(solid("FFFFFF", 12)) })), 820)]);
}

function infoBox(runs, fill = LIGHT2) {
  return table([W], [row([cell(runs.map((r) => p(r, { after: 40 })), { w: W, fill, borders: allBorders(none), m: 120 })])]);
}

// ================= SEITE 1: Bastelbogen =================
const page1 = [
  kicker("Bastelbogen · bitte einseitig drucken"),
  h1("Zeitformen-Baukasten"),
  p([t("Schneide alle Karten an den gestrichelten Linien aus. Lege die ", {}), t("A-Karten", { bold: true }), t(" und die ", {}), t("B-Karten", { bold: true }), t(" auf zwei getrennte Stapel – noch nicht aufkleben!")], { after: 140 }),
  p([t("✂  A-Karten ", { font: "Cambria", bold: true, size: 26, color: NAVY }), t("für Aufgabe 1 (Sortieren)", { color: MUTED })], { after: 80 }),
  cardGrid(wordsA, 4, 2551, 880, 30),
  p(t(""), { after: 120 }),
  p([t("✂  B-Karten ", { font: "Cambria", bold: true, size: 26, color: NAVY }), t("für Aufgabe 2 (Zeitformen bauen)", { color: MUTED })], { after: 80 }),
  cardGrid(wordsB, 6, 1701, 760, 26),
  p(t(""), { after: 80 }),
  infoBox([[t("Tipp: ", { bold: true, color: NAVY }), t("Schreibe auf die Rückseite jeder A-Karte ein kleines „A“ und jeder B-Karte ein „B“. Dann geht nichts durcheinander.", { size: 20 })]]),
];

// ================= SEITE 2: Aufgabe 1 =================
const colW = [3402, 3402, 3402];
const sortRows = [row([
  cell(p(t("Hilfsverben", { bold: true, color: "FFFFFF", size: 24 }), { align: AlignmentType.CENTER, after: 0 }), { w: colW[0], fill: C.hv }),
  cell(p(t("Modalverben", { bold: true, color: "FFFFFF", size: 24 }), { align: AlignmentType.CENTER, after: 0 }), { w: colW[1], fill: C.mv }),
  cell(p(t("Partizip II", { bold: true, color: "FFFFFF", size: 24 }), { align: AlignmentType.CENTER, after: 0 }), { w: colW[2], fill: C.p2 }),
], 520)];
sortRows.push(row([
  cell(p(t("Formen von haben, sein, werden", { size: 17, italics: true, color: MUTED }), { align: AlignmentType.CENTER, after: 0 }), { w: colW[0], fill: LIGHT }),
  cell(p(t("können, müssen, dürfen, …", { size: 17, italics: true, color: MUTED }), { align: AlignmentType.CENTER, after: 0 }), { w: colW[1], fill: LIGHT }),
  cell(p(t("meist ge- … -t / ge- … -en", { size: 17, italics: true, color: MUTED }), { align: AlignmentType.CENTER, after: 0 }), { w: colW[2], fill: LIGHT }),
], 420));
for (let i = 0; i < 8; i++) sortRows.push(row(colW.map((w) => cell(p(t("")), { w })), 940));

const page2 = [
  new Paragraph({ children: [new PageBreak()] }),
  kicker("Aufgabe 1 · Sortieren"),
  h1("Welcher Baustein ist das?"),
  p([t("1. ", { bold: true }), t("Sortiere die "), t("A-Karten", { bold: true }), t(" in die richtige Spalte. Erst legen, dann vergleichen, dann aufkleben.")], { after: 40 }),
  p([t("2. ", { bold: true }), t("Vier Karten passen in keine Spalte – das sind die "), t("Stolperkarten", { bold: true }), t(". Klebe sie unten auf.")], { after: 40 }),
  p([t("3. ", { bold: true }), t("Male jede Karte in der Farbe ihres Bausteins an (Farbcode unten).")], { after: 120 }),
  legend(),
  p(t(""), { after: 100 }),
  table(colW, sortRows),
  p(t(""), { after: 100 }),
  p([t("Stolperkarten ", { font: "Cambria", bold: true, size: 24, color: NAVY }), t("– weder Hilfsverb, noch Modalverb, noch Partizip II", { color: MUTED, size: 20 })], { after: 60, keepNext: true }),
  table([2551, 2551, 2552, 2552], [row([2551, 2551, 2552, 2552].map((w) => cell(p(t("")), { w, fill: LIGHT })), 940)]),
  p(t(""), { after: 60 }),
  p([t("Was sind die Stolperkarten? ", { bold: true }), t("_______________________________________________________________")], { after: 60 }),
];

// ================= SEITE 3: Aufgabe 2 =================
const bw = [2000, 1100, 2000, 1450, 3656];
const hdr = (txt, sub, w, fill) => cell([
  p(t(txt, { bold: true, color: "FFFFFF", size: 20 }), { align: AlignmentType.CENTER, after: 0 }),
  sub ? p(t(sub, { color: "FFFFFF", size: 15 }), { align: AlignmentType.CENTER, after: 0 }) : null,
].filter(Boolean), { w, fill });
const bauRows = [row([
  hdr("Zeitform", null, bw[0], NAVY),
  hdr("Vorfeld", null, bw[1], NAVY),
  hdr("Position 2", "gebeugtes Verb", bw[2], "24306E"),
  hdr("Mittelfeld", null, bw[3], NAVY),
  hdr("Satzende", "Partizip II / Infinitiv", bw[4], "24306E"),
], 620)];
const zeiten = ["Präsens", "Präteritum", "Perfekt", "Plusquamperfekt", "Futur I", "Futur II", "mit Modalverb (Präsens)"];
zeiten.forEach((z, i) => {
  const isModal = i === 6;
  bauRows.push(row([
    cell(p(t(z, { bold: true, color: isModal ? C.mv : NAVY, size: 18 }), { after: 0 }), { w: bw[0], fill: LIGHT }),
    cell(p(t("Ich", { font: "Cambria", size: 24 }), { align: AlignmentType.CENTER, after: 0 }), { w: bw[1] }),
    cell(p(t("")), { w: bw[2], borders: allBorders(solid(NAVY, 12)) }),
    cell(p(t("einen Kuchen", { font: "Cambria", size: 22 }), { align: AlignmentType.CENTER, after: 0 }), { w: bw[3] }),
    cell(p(t("")), { w: bw[4], borders: allBorders(solid(NAVY, 12)) }),
  ], 900));
});

const page3 = [
  new Paragraph({ children: [new PageBreak()] }),
  kicker("Aufgabe 2 · Zeitformen bauen"),
  h1("Der Tempus-Bauplan"),
  p([t("Klebe die "), t("B-Karten", { bold: true }), t(" in die dicken Kästen. Jede Zeile wird ein Satz mit „Ich … einen Kuchen …“. Male die Karten danach im Farbcode an. ("), t("backe", { bold: true }), t(" und "), t("backte", { bold: true }), t(" sind einfache Verbformen – sie bleiben weiß.)")], { after: 60 }),
  p([t("Tipp: ", { bold: true, color: C.hv }), t("Manche Zeitformen brauchen am Satzende gar nichts. Welche?", { italics: true })], { after: 120 }),
  table(bw, bauRows),
  p(t(""), { after: 80 }),
  p(t("Schau genau hin – und kreuze an bzw. ergänze:", { bold: true, color: NAVY, size: 24, font: "Cambria" }), { after: 80, keepNext: true }),
  p([t("a) Nur "), t("einen", { bold: true }), t(" Verbteil haben:  ☐ Präsens  ☐ Präteritum  ☐ Perfekt  ☐ Futur I")], { after: 80 }),
  p([t("b) Wenn ein Hilfsverb auf Position 2 steht, wandert der zweite Verbteil ans ________________.")], { after: 80 }),
  p([t("c) Hilfsverb und Satzende bilden zusammen eine "), t("Verbklammer", { bold: true }), t(". Ziehe in jeder Zeile einen Bogen von Position 2 zum Satzende.")], { after: 80 }),
  p([t("d) Perfekt und Plusquamperfekt unterscheiden sich nur in ____________________________.")], { after: 40 }),
];

// ================= SEITE 4: Aufgabe 3 =================
const merkW = [3300, 6906];
const merk = [
  ["Perfekt", "haben / sein im __________________  +  __________________"],
  ["Plusquamperfekt", "haben / sein im __________________  +  __________________"],
  ["Futur I", "__________________ im Präsens  +  __________________"],
  ["Futur II", "werden im Präsens  +  __________________  +  haben / sein"],
  ["Satz mit Modalverb", "Modalverb (gebeugt)  +  __________________"],
];
const merkTable = table(merkW, merk.map(([a, b]) => row([
  cell(p(t(a, { bold: true, color: NAVY }), { after: 0 }), { w: merkW[0], fill: LIGHT2, borders: allBorders(solid("FFFFFF", 12)) }),
  cell(p(t(b, { size: 21 }), { after: 0 }), { w: merkW[1], fill: LIGHT, borders: allBorders(solid("FFFFFF", 12)) }),
], 520)));

const detW = [5106, 1700, 1700, 1700];
const saetze = [
  "Lena ist ins Kino gegangen.",
  "Wir hatten die Hausaufgaben vergessen.",
  "Du wirst die Prüfung schaffen.",
  "Die Kinder spielten im Garten.",
  "Ich darf heute länger aufbleiben.",
  "Bis morgen werdet ihr das Buch gelesen haben.",
  "Er war zu spät gekommen.",
];
const detRows = [row([
  hdr("Satz", null, detW[0], NAVY), hdr("Position 2", null, detW[1], NAVY), hdr("Satzende", null, detW[2], NAVY), hdr("Zeitform", null, detW[3], NAVY),
], 440)];
saetze.forEach((s, i) => detRows.push(row([
  cell(p([t(`${i + 1}  `, { bold: true, color: C.hv }), t(s, { font: "Cambria", size: 21 })], { after: 0 }), { w: detW[0] }),
  cell(p(t("")), { w: detW[1] }), cell(p(t("")), { w: detW[2] }), cell(p(t("")), { w: detW[3] }),
], 560, HeightRule.ATLEAST)));

const page4 = [
  new Paragraph({ children: [new PageBreak()] }),
  kicker("Aufgabe 3 · Sichern"),
  h1("So werden Zeitformen gebaut"),
  p([t("Ergänze die Bauformeln. Wörter zur Auswahl: "), t("Präsens · Präteritum · Partizip II (3×) · Infinitiv (2×) · werden", { bold: true, color: MUTED })], { after: 100 }),
  merkTable,
  p(t(""), { after: 60 }),
  p([t("Tempus-Detektiv  ", { font: "Cambria", bold: true, size: 26, color: NAVY }), t("Markiere die Verbteile im Farbcode, trage sie ein und bestimme die Zeitform.", { size: 20 })], { after: 80, keepNext: true }),
  table(detW, detRows),
  p(t(""), { after: 60 }),
  p([t("★ Profi-Aufgabe  ", { font: "Cambria", bold: true, size: 24, color: C.p2 }), t("Baue den Satz „Ich gehe nach Hause.“ im Perfekt, Plusquamperfekt und Futur I. Achtung: ", { size: 20 }), t("gehen", { bold: true, size: 20 }), t(" bildet das Perfekt mit ", { size: 20 }), t("sein", { bold: true, size: 20 }), t("!", { size: 20 })], { after: 100, keepNext: true }),
  p(t("Perfekt: ______________________________________________________________________"), { after: 100 }),
  p(t("Plusquamperfekt: ______________________________________________________________"), { after: 100 }),
  p(t("Futur I: ______________________________________________________________________"), { after: 0 }),
];

// ================= Lösung =================
function solTable(widths, headRow, rows, fills) {
  return table(widths, [
    row(headRow.map((h, i) => hdr(h, null, widths[i], NAVY)), 440),
    ...rows.map((r) => row(r.map((v, i) => cell(typeof v === "string" ? p(t(v, { size: 20, bold: fills && fills[i] ? true : false, color: fills && fills[i] ? fills[i] : "111111" }), { after: 0 }) : v, { w: widths[i] })), 400, HeightRule.ATLEAST)),
  ]);
}
const colored = (parts) => p(parts.map(([txt, col]) => t(txt, { size: 20, bold: !!col, color: col || "111111", font: "Cambria" })), { after: 0 });

const loesung = [
  kicker("Lösungsblatt für die Lehrkraft"),
  h1("Zeitformen-Baukasten – Lösungen"),
  h2("1", "Sortieren"),
  solTable([2551, 2551, 2552, 2552], ["Hilfsverben", "Modalverben", "Partizip II", "Stolperkarten"], [
    ["habe", "kann", "gespielt", "spielte (Präteritum)"],
    ["hast", "musst", "gelaufen", "ging (Präteritum)"],
    ["bist", "darf", "gegessen", "laufen (Infinitiv)"],
    ["ist", "könnt", "geschrieben", "gehst (Präsens)"],
    ["wird", "wollen", "gefahren", ""],
    ["werden", "müssen", "gekauft", ""],
    ["hatten", "sollte", "verstanden (ohne ge-!)", ""],
    ["war", "mochte", "aufgeräumt (ge- in der Mitte!)", ""],
  ], [C.hv, C.mv, C.p2, MUTED]),
  p([t("Stolperkarten: ", { bold: true, size: 20 }), t("Es sind einfache Verbformen (Vollverben im Präsens/Präteritum oder Infinitiv) – sie brauchen keinen Baustein-Partner.", { size: 20 })], { before: 80, after: 60 }),
  h2("2", "Tempus-Bauplan"),
  solTable([2000, 1150, 1900, 1500, 3656], ["Zeitform", "Vorfeld", "Position 2", "Mittelfeld", "Satzende"], [
    ["Präsens", "Ich", "backe", "einen Kuchen", "—"],
    ["Präteritum", "Ich", "backte", "einen Kuchen", "—"],
    ["Perfekt", "Ich", colored([["habe", C.hv]]), "einen Kuchen", colored([["gebacken", C.p2]])],
    ["Plusquamperfekt", "Ich", colored([["hatte", C.hv]]), "einen Kuchen", colored([["gebacken", C.p2]])],
    ["Futur I", "Ich", colored([["werde", C.hv]]), "einen Kuchen", colored([["backen", C.inf]])],
    ["Futur II", "Ich", colored([["werde", C.hv]]), "einen Kuchen", colored([["gebacken ", C.p2], ["haben", C.hv]])],
    ["mit Modalverb", "Ich", colored([["muss", C.mv]]), "einen Kuchen", colored([["backen", C.inf]])],
  ]),
  p([t("a) ", { bold: true, size: 20 }), t("Präsens, Präteritum   ", { size: 20 }), t("b) ", { bold: true, size: 20 }), t("Satzende   ", { size: 20 }), t("d) ", { bold: true, size: 20 }), t("der Zeitform des Hilfsverbs (habe = Präsens → Perfekt; hatte = Präteritum → Plusquamperfekt)", { size: 20 })], { before: 80, after: 60 }),
  h2("3", "Bauformeln & Tempus-Detektiv"),
  p(t("Perfekt = haben/sein im Präsens + Partizip II · Plusquamperfekt = haben/sein im Präteritum + Partizip II · Futur I = werden im Präsens + Infinitiv · Futur II = werden + Partizip II + haben/sein · Modalverb + Infinitiv", { size: 20 }), { after: 100 }),
  solTable([4206, 1500, 1800, 2700], ["Satz", "Position 2", "Satzende", "Zeitform"], [
    ["1  Lena ist ins Kino gegangen.", "ist", "gegangen", "Perfekt"],
    ["2  Wir hatten die Hausaufgaben vergessen.", "hatten", "vergessen", "Plusquamperfekt"],
    ["3  Du wirst die Prüfung schaffen.", "wirst", "schaffen", "Futur I"],
    ["4  Die Kinder spielten im Garten.", "spielten", "—", "Präteritum"],
    ["5  Ich darf heute länger aufbleiben.", "darf", "aufbleiben", "Präsens (+ Modalverb)"],
    ["6  Bis morgen werdet ihr das Buch gelesen haben.", "werdet", "gelesen haben", "Futur II"],
    ["7  Er war zu spät gekommen.", "war", "gekommen", "Plusquamperfekt"],
  ]),
  p([t("★ Profi: ", { bold: true, size: 20, color: C.p2 }), t("Ich bin nach Hause gegangen. · Ich war nach Hause gegangen. · Ich werde nach Hause gehen.", { size: 20 })], { before: 80, after: 0 }),
];

const pageProps = { page: { size: { width: 11906, height: 16838 }, margin: { top: 900, bottom: 800, left: 850, right: 850, header: 400, footer: 400 } } };
const styles = { default: { document: { run: { font: "Calibri", size: 22 } } } };

const ab = new Document({ styles, sections: [{ properties: pageProps, headers: { default: header }, footers: { default: footer("Zeitformen-Baukasten") }, children: [...page1, ...page2, ...page3, ...page4] }] });
const lo = new Document({ styles, sections: [{ properties: pageProps, footers: { default: footer("Lösungen") }, children: loesung }] });

Packer.toBuffer(ab).then((b) => fs.writeFileSync("Arbeitsblatt_Zeitformen_Baukasten.docx", b));
Packer.toBuffer(lo).then((b) => fs.writeFileSync("Loesung_Zeitformen_Baukasten.docx", b));
