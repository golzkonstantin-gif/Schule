const fs = require("fs");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, BorderStyle,
  ShadingType, AlignmentType, VerticalAlign, HeightRule, PageBreak, Header, Footer, TabStopType,
  PageNumber,
} = require("docx");

const NAVY = "1E2761", MUTED = "5B6B8C", LIGHT = "F4F7FD", LIGHT2 = "EAF0FA", BOARD = "F3F6F1";
const C = { hv: "D9534F", mv: "2E9E6B", p2: "E8A33D", inf: "3F7CC4" };
const W = 10206;

const t = (text, o = {}) => new TextRun({ text, font: o.font || "Calibri", size: o.size || 22, bold: o.bold, italics: o.italics, color: o.color || "000000", underline: o.u ? {} : undefined });
const p = (runs, o = {}) => new Paragraph({ children: Array.isArray(runs) ? runs : [runs], alignment: o.align, spacing: { before: o.before ?? 0, after: o.after ?? 100, line: o.line }, keepNext: o.keepNext, indent: o.indent });
const h1 = (text) => p(t(text, { font: "Cambria", size: 40, bold: true, color: NAVY }), { after: 60 });
const kicker = (text) => p(t(text.toUpperCase(), { size: 18, color: MUTED, bold: true }), { after: 20 });
const h2 = (label, text, after = 80) => p([t(label + "  ", { font: "Cambria", size: 26, bold: true, color: C.hv }), t(text, { font: "Cambria", size: 26, bold: true, color: NAVY })], { before: 200, after, keepNext: true });
const h3 = (text, col = NAVY) => p(t(text, { font: "Cambria", size: 23, bold: true, color: col }), { before: 120, after: 60, keepNext: true });
const gap = (n = 80) => p(t(""), { after: n });
const br = () => new Paragraph({ children: [new PageBreak()] });

const none = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
const noBorders = { top: none, bottom: none, left: none, right: none };
const solid = (c = NAVY, s = 8) => ({ style: BorderStyle.SINGLE, size: s, color: c });
const allBorders = (b) => ({ top: b, bottom: b, left: b, right: b });
const dashed = { style: BorderStyle.DASHED, size: 8, color: "7A869E" };
const thin = allBorders(solid("B9C6E8", 6));

function cell(children, o = {}) {
  return new TableCell({
    children: Array.isArray(children) ? children : [children],
    width: { size: o.w, type: WidthType.DXA },
    shading: o.fill ? { fill: o.fill, type: ShadingType.CLEAR, color: "auto" } : undefined,
    borders: o.borders || thin,
    verticalAlign: o.valign || VerticalAlign.CENTER,
    margins: { top: o.m ?? 60, bottom: o.m ?? 60, left: o.ml ?? 100, right: 100 },
    columnSpan: o.span,
  });
}
const table = (widths, rows) => new Table({ width: { size: widths.reduce((a, b) => a + b, 0), type: WidthType.DXA }, columnWidths: widths, rows });
const row = (cells, h, rule = HeightRule.ATLEAST) => new TableRow({ children: cells, height: h ? { value: h, rule } : undefined, cantSplit: true });
const hdr = (txt, w, fill = NAVY, size = 20) => cell(p(t(txt, { bold: true, color: "FFFFFF", size }), { align: AlignmentType.CENTER, after: 0 }), { w, fill });
const tc = (txt, w, o = {}) => cell(p(Array.isArray(txt) ? txt : t(txt, { size: o.size || 21, bold: o.bold, color: o.color, font: o.font, italics: o.it }), { after: 0, align: o.align }), { w, fill: o.fill });

// simple grid table: header + rows of strings/runs
function grid(widths, head, rows, o = {}) {
  return table(widths, [
    row(head.map((h, i) => hdr(h, widths[i], o.headFill ? o.headFill[i] : NAVY)), 420),
    ...rows.map((r) => row(r.map((v, i) => tc(v, widths[i], { font: o.fonts && o.fonts[i], align: o.align && o.align[i], bold: o.bold && o.bold[i], color: o.colors && o.colors[i] })), o.h || 460)),
  ]);
}

function box(children, fill = LIGHT2, border) {
  return table([W], [row([cell(children, { w: W, fill, borders: border || allBorders(none), m: 140, ml: 180 })])]);
}
// Merksatz box
function merk(nr, title, runs) {
  return box([
    p([t(`Merksatz ${nr} · `, { bold: true, color: C.hv, size: 20 }), t(title, { bold: true, color: NAVY, size: 22, font: "Cambria" })], { after: 60 }),
    ...runs.map((r) => p(r, { after: 40 })),
  ], LIGHT2);
}
// Tafel box
function tafel(title, children) {
  return table([W], [
    row([cell(p(t("TAFEL · " + title, { bold: true, color: "FFFFFF", size: 18 }), { after: 0 }), { w: W, fill: "3D4A3F", borders: allBorders(solid("3D4A3F", 12)) })]),
    row([cell(children, { w: W, fill: BOARD, borders: allBorders(solid("3D4A3F", 12)), m: 140, ml: 200 })]),
  ]);
}
const R = (s) => t(s, { bold: true, color: C.hv, font: "Cambria", size: 22 });
const Y = (s) => t(s, { bold: true, color: "C98A1E", font: "Cambria", size: 22 });
const B = (s) => t(s, { bold: true, color: C.inf, font: "Cambria", size: 22 });
const N = (s, o = {}) => t(s, Object.assign({ font: "Cambria", size: 22 }, o));
const S = (s, o = {}) => t(s, Object.assign({ size: 21 }, o));


// ================================================================
// Tafelskript „Verbformen verstehen“
// ================================================================

const header = (label) => new Header({
  children: [new Paragraph({
    tabStops: [{ type: TabStopType.RIGHT, position: W }],
    children: [t("Deutsch · Klasse 7A · " + label, { size: 17, color: MUTED }), t("\tfür die Lehrkraft", { size: 17, color: MUTED })],
  })],
});
const footer = (label) => new Footer({
  children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ font: "Calibri", size: 16, color: MUTED, children: [label + " · Seite ", PageNumber.CURRENT] })] })],
});

// ---------- Was die Klasse gerade tut (nur zur Planung) ----------
const LABELS = {
  zuhoeren: "hört zu", gemeinsam: "macht mündlich mit", abschreiben: "schreibt ab", luecken: "arbeitet allein",
  tafel: "kommt an die Tafel", kontrolle: "kontrolliert", partner: "arbeitet zu zweit", fertig: "gibt ab",
};
function labelCell(m, w) {
  return cell([
    p(t("Klasse", { size: 16, color: MUTED }), { after: 10 }),
    p(t(LABELS[m], { bold: true, color: NAVY, size: 20 }), { after: 0 }),
  ], { w, fill: LIGHT, borders: allBorders(solid("D5DDEE", 6)), m: 100, valign: VerticalAlign.TOP });
}

// Schritt: Modus + Inhalt
const BW = 1700, CW = W - BW;
function step(m, parts) {
  const children = [];
  if (parts.say) children.push(p([t("Ansage: ", { bold: true, size: 20, color: C.hv }), t("„" + parts.say + "“", { italics: true, size: 21 })], { after: 60 }));
  (parts.do || []).forEach((d) => children.push(p(Array.isArray(d) ? d : S(d), { after: 50 })));
  if (parts.board) { children.push(...[].concat(parts.board)); children.push(p(t(""), { after: 20 })); }
  if (parts.sol) children.push(p([t("Lösung: ", { bold: true, size: 19, color: MUTED }), t(parts.sol, { size: 19, color: MUTED })], { after: 0 }));
  if (children.length && !parts.sol && !parts.board) children[children.length - 1] = children[children.length - 1];
  return table([BW, CW], [row([labelCell(m, BW), cell(children, { w: CW, borders: allBorders(solid("D5DDEE", 6)), m: 100, ml: 160, valign: VerticalAlign.TOP })])]);
}
const steps = (arr) => arr.flatMap((s) => [step(s[0], s[1]), p(t(""), { after: 40 })]);

// Tafel-Kasten (innerhalb eines Schritts)
const TW = CW - 340;
function board(title, children) {
  return table([TW], [
    row([cell(p(t("TAFEL · " + title, { bold: true, color: "FFFFFF", size: 16 }), { after: 0 }), { w: TW, fill: "3D4A3F", borders: allBorders(solid("3D4A3F", 12)), m: 30 })]),
    row([cell(children, { w: TW, fill: BOARD, borders: allBorders(solid("3D4A3F", 12)), m: 100, ml: 160 })]),
  ]);
}
const bl = (runs, o = {}) => p(Array.isArray(runs) ? runs : [N(runs)], Object.assign({ after: 30 }, o));
const merkBoard = (nr, title, runs) => board(`Merksatz ${nr} (rot umrahmen)`, [
  p([t(title, { bold: true, color: C.hv, font: "Cambria", size: 22 })], { after: 40 }),
  p(runs, { after: 0 }),
]);
// Übung an der Tafel – immer im gleichen Format
function auftrag(title, todo, lines, meta) {
  return board("rechts · " + title, [
    ...todo.map((x, i) => p([t(`${i + 1}. `, { bold: true, color: NAVY, size: 21 }), t(x, { bold: true, color: NAVY, size: 21 })], { after: 30 })),
    p(t(""), { after: 30 }),
    ...lines,
    p(t(""), { after: 30 }),
    p([t("Zeit: ", { bold: true, size: 19, color: "3D4A3F" }), t(meta, { size: 19, color: "3D4A3F" }), t("     Fertig? Stift hinlegen und noch einmal durchlesen.", { size: 19, italics: true, color: "3D4A3F" })], { after: 0 }),
  ]);
}
function phase(nr, title, min) {
  return p([t(`${nr}  `, { font: "Cambria", size: 28, bold: true, color: C.hv }), t(title, { font: "Cambria", size: 28, bold: true, color: NAVY }), t(`   ${min} Min.`, { size: 20, bold: true, color: "C77C12" })], { before: 200, after: 100, keepNext: true });
}

// ================================================================
// SKRIPT
// ================================================================
const doc = [];
doc.push(kicker("Tafelskript für die Lehrkraft · Doppelstunde (90 Min.)"), h1("Verbformen verstehen"),
  p(t("Ohne Arbeitsblatt: Alles entsteht an der Tafel, die Klasse schreibt ins Heft. Klarheit entsteht durch zwei Dinge: eine eindeutige Ansage bei jedem Wechsel und Übungen, die an der Tafel immer im gleichen Format stehen.", { size: 21 }), { after: 100 }));

doc.push(h3("1  Eindeutige Ansagen"));
doc.push(p(t("Bei jedem Wechsel sagen Sie zuerst, was die Klasse tut – erst dann den Inhalt. Die wörtlichen Ansagen stehen im Skript. Es kommen immer dieselben Formulierungen vor:", { size: 21 }), { after: 80 }));
{
  const A = [
    ["„Stifte liegen. Ihr hört nur zu.“", "Sie erklären oder fragen, niemand schreibt."],
    ["„Meldet euch. Ihr schreibt noch nichts auf.“", "Gemeinsam an der Tafel erarbeiten."],
    ["„Jetzt schreibt ihr ab: …“", "Merksatz oder Tabelle wird ins Heft übernommen."],
    ["„Jetzt arbeitet ihr allein. Die Aufgabe steht rechts.“", "Übung im Heft, leise."],
    ["„Wer drankommt, schreibt an die Tafel. Alle prüfen mit.“", "Lösungen an der Tafel sammeln."],
    ["„Vergleicht und verbessert mit Grün.“", "Kontrolle im eigenen Heft."],
  ];
  doc.push(table([4700, 5506], A.map(([a, b]) => row([tc([t(a, { italics: true, bold: true, size: 20, color: NAVY })], 4700, { fill: LIGHT }), tc(b, 5506, { size: 20 })], 380))));
}
doc.push(p([S("Fertig-Signal: ", { bold: true }), S("Wer fertig ist, legt den Stift hin. So sehen Sie, wann alle so weit sind – erst dann wird gewischt oder weitergemacht.")], { before: 100, after: 40 }));
doc.push(p([S("Gleicher Ablauf in jeder Phase: ", { bold: true }), S("gemeinsam entdecken → Merksatz abschreiben → allein üben → an der Tafel vergleichen → verbessern.")], { after: 60 }));

doc.push(h3("2  Übungen immer im gleichen Format"));
doc.push(p(t("Jede Übung steht rechts an der Tafel mit Nummer und Titel, nummerierten Arbeitsschritten (Verb am Anfang), der Zeit und der Arbeitsform. Beispiel:", { size: 21 }), { after: 80 }));
doc.push(table([W], [row([cell([auftrag("Ü1 Wörterbuch-Detektiv", ["Schreibe Ü1 an den Rand.", "Schreibe jedes Wort ab und ergänze die Grundform."], [bl("konnte → ______     liest → ______     bin → ______", { after: 0 })], "3 Minuten · allein · leise")], { w: W, borders: allBorders(none), m: 0, ml: 0 })])]));

doc.push(br(), h3("Tafelaufteilung"));
doc.push(table([2600, 5006, 2600], [
  row([hdr("links", 2600, "3D4A3F", 18), hdr("Mitte", 5006, "3D4A3F", 18), hdr("rechts", 2600, "3D4A3F", 18)], 340),
  row([
    cell([p(S("Ablaufplan (vorher anschreiben, abhaken):", { bold: true, size: 19 }), { after: 40 }), ...["1 Grundform", "2 Familien haben & sein", "3 Früher: hatte & war", "4 Partizip II & Perfekt", "5 Exit-Ticket"].map((x) => p(S("☐ " + x, { size: 19 }), { after: 20 }))], { w: 2600, fill: BOARD, valign: VerticalAlign.TOP }),
    cell([p(S("Tafelbild und Merksätze der aktuellen Phase", { size: 19 }), { after: 40 }), p(S("wird nach jeder Phase gewischt – erst wenn alle „fertig“ sind", { size: 19, italics: true, color: MUTED }), { after: 0 })], { w: 5006, fill: BOARD, valign: VerticalAlign.TOP }),
    cell([p(S("Übungsfläche", { size: 19 }), { after: 40 }), p(S("Hier stehen die Übungen, hier schreiben die Schüler ihre Lösungen an", { size: 19, italics: true, color: MUTED }), { after: 0 })], { w: 2600, fill: BOARD, valign: VerticalAlign.TOP }),
  ]),
]));
doc.push(p([S("Material: ", { bold: true }), S("Kreide/Stifte in Rot, Gelb, Blau, Grün · Schüler: Heft, Lineal, Buntstifte Rot, Gelb, Blau, Grün, kleiner Zettel für das Exit-Ticket")], { before: 100, after: 40 }));
doc.push(p([S("Farben im Heft: ", { bold: true }), t("haben/sein rot", { bold: true, color: C.hv, size: 21 }), S(" · "), t("Partizip II gelb", { bold: true, color: "C98A1E", size: 21 }), S(" · "), t("Infinitiv blau", { bold: true, color: C.inf, size: 21 }), S(" · "), t("Verbesserungen grün", { bold: true, color: C.mv, size: 21 })], { after: 0 }));

// ---------- Hefteintrag ----------
doc.push(h3("So sieht das Heft am Ende der Stunde aus"));
doc.push(p(t("Tipp: Diese Übersicht zu Beginn kurz zeigen oder selbst im Kopf behalten – die Reihenfolge ist in jeder Phase gleich: Merksatz, dann Übung.", { size: 20, italics: true, color: MUTED }), { after: 80 }));
{
  const H = [
    ["Überschrift", "Verbformen verstehen (mit Lineal unterstrichen) · Datum rechts"],
    ["Merksatz 1 + 2", "Infinitiv · Personalform (rot umrahmt)"],
    ["Ü1", "konnte → können … (6 Formen)"],
    ["Merksatz 3", "Verbfamilie + Familien-Tabelle mit 5 Spalten (die „früher“-Spalten zunächst leer)"],
    ["Ü2", "zwei Spalten: Familie haben | Familie sein"],
    ["Merksatz 4", "haben und sein im Präteritum – „früher“-Spalten der Tabelle werden jetzt ausgefüllt"],
    ["Ü3", "nur Nummer + Lösungswort: 1 war, 2 hattest …"],
    ["Merksatz 5", "Partizip II: Bildung, wofür man es braucht, Test"],
    ["Ü4", "spielen → gespielt … (6 Verben)"],
    ["Merksatz 6", "Perfekt mit haben oder sein + je ein Beispielsatz"],
    ["Ü5", "nur Nummer + Lösungswort: 1 bin, 2 habe …"],
    ["Merksatz 7", "Hilfsverb oder nicht? + Beispielpaar"],
  ];
  doc.push(table([2300, 7906], H.map(([a, b]) => row([tc(a, 2300, { bold: true, color: a.startsWith("Merk") ? C.hv : a === "Überschrift" ? NAVY : "C77C12", fill: LIGHT, size: 20 }), tc(b, 7906, { size: 20 })], 360))));
}

// ---------- Phase 0 ----------
doc.push(br(), phase("0", "Start: Ablauf klären, Heft vorbereiten", 5));
doc.push(...steps([
  ["zuhoeren", { say: "Stifte liegen. Heute arbeiten wir ohne Arbeitsblatt. Ich sage euch immer genau, was ihr gerade tut: zuhören, abschreiben, allein üben oder vergleichen. Die Übungen stehen immer rechts an der Tafel – mit genauer Anleitung. Wer fertig ist, legt den Stift hin.", do: ["Ablaufplan links an der Tafel zeigen: Das machen wir heute."] }],
  ["abschreiben", { say: "Heft auf. Datum nach rechts. Überschrift: Verbformen verstehen. Mit Lineal unterstreichen. Wer fertig ist, legt den Stift hin.", board: board("Mitte", [bl([N("Verbformen verstehen", { bold: true, u: true })])]) }],
]));

// ---------- Phase 1 ----------
doc.push(phase("1", "Grundform und Personalform", 15));
doc.push(...steps([
  ["zuhoeren", { say: "Stifte liegen. Ihr hört nur zu. Tim sucht im Wörterbuch das Wort ‚ging‘. Er findet es nicht. Warum nicht?", do: ["Zwei, drei Antworten mündlich sammeln. Ziel: Im Wörterbuch steht die Grundform gehen."] }],
  ["gemeinsam", { say: "Meldet euch. Ihr schreibt noch nichts auf. Wie steht das Wort im Wörterbuch?", do: ["Die Grundformen nennen lassen und selbst blau an die Tafel schreiben.", "Danach „ich spiel_ / du spiel_ / er spiel_“ anschreiben, Endungen nennen lassen und unterstreichen."],
    board: board("Mitte", [bl([N("ging → "), B("gehen"), N("     isst → "), B("essen"), N("     hat → "), B("haben"), N("     war → "), B("sein")]), bl([N("ich spiel"), N("e", { bold: true, u: true }), N("   du spiel"), N("st", { bold: true, u: true }), N("   er spiel"), N("t", { bold: true, u: true })], { after: 0 })]) }],
  ["abschreiben", { say: "Jetzt schreibt ihr ab: Merksatz 1 und Merksatz 2. Rahmt jeden Merksatz mit Rot ein.", board: [merkBoard(1, "Infinitiv", [S("Der Infinitiv ist die "), S("Grundform", { bold: true }), S(" eines Verbs. So steht das Verb im "), S("Wörterbuch", { bold: true }), S(". Er endet auf -en oder -n: spielen, basteln.")]), p(t(""), { after: 40 }), merkBoard(2, "Personalform", [S("Die Personalform (gebeugte Form) passt sich der "), S("Person", { bold: true }), S(" an: ich spiele, du spielst, er spielt.")])] }],
  ["luecken", { say: "Jetzt arbeitet ihr allein. Die Aufgabe steht rechts: Ü1. Lest die Schritte mit mir: … Drei Minuten. Wer fertig ist: Stift hinlegen.", board: auftrag("Ü1 Wörterbuch-Detektiv", ["Schreibe Ü1 an den Rand.", "Schreibe jedes Wort ab und ergänze die Grundform."], [bl("konnte → ______     liest → ______     bin → ______"), bl("fuhr → ______     wusste → ______     schläft → ______", { after: 0 })], "3 Minuten · allein · leise"), sol: "können · lesen · sein · fahren · wissen · schlafen" }],
  ["tafel", { say: "Sechs von euch schreiben je eine Lösung an die Tafel. Alle anderen prüfen mit: Daumen hoch oder Daumen runter?", do: ["Schwächere Schüler zuerst drannehmen (leichtere Formen: liest, fuhr)."] }],
  ["kontrolle", { say: "Vergleicht euer Heft mit der Tafel. Fehler verbessert ihr mit Grün." }],
]));

// ---------- Phase 2 ----------
doc.push(phase("2", "Die Familien haben und sein", 15));
const famBoard = (past) => {
  const w = [1400, 1500, 1500, 1500, 1500];
  const P = ["ich", "du", "er/sie/es", "wir", "ihr", "sie/Sie"];
  const K = { hp: ["habe", "hast", "hat", "haben", "habt", "haben"], hv: ["hatte", "hattest", "hatte", "hatten", "hattet", "hatten"], sp: ["bin", "bist", "ist", "sind", "seid", "sind"], sv: ["war", "warst", "war", "waren", "wart", "waren"] };
  const nb = allBorders(solid("9AA59C", 4));
  const c = (x, i, o = {}) => cell(p(x, { after: 0, align: AlignmentType.CENTER }), { w: w[i], borders: nb, m: 20, fill: o.fill });
  return table(w, [
    row([c(N(""), 0), c(R("haben · jetzt"), 1), c(R("haben · früher"), 2), c(R("sein · jetzt"), 3), c(R("sein · früher"), 4)]),
    ...P.map((ps, i) => row([c(N(ps, { color: MUTED }), 0), c(N(K.hp[i]), 1), c(N(past ? K.hv[i] : ""), 2), c(N(K.sp[i]), 3), c(N(past ? K.sv[i] : ""), 4)])),
  ]);
};
doc.push(...steps([
  ["tafel", { say: "Wer gehört zur Familie haben? Wer drankommt, kommt nach vorne und schreibt ein Familienmitglied in die Tabelle.", do: ["Tabelle mit fünf Spalten vorbereitet anschreiben. Nacheinander Schüler die Präsensformen von haben und sein eintragen lassen. Die „früher“-Spalten bleiben leer."], board: board("Mitte", [famBoard(false)]) }],
  ["zuhoeren", { say: "Stifte liegen. Schaut auf ‚bin‘. Sieht das aus wie ‚sein‘? Woher weiß ich trotzdem, dass es dazugehört?", do: ["Ziel: nur über die Frage „Wie heißt die Grundform?“ – das ist der Familienname."] }],
  ["abschreiben", { say: "Schreibt Merksatz 3 ab. Dann zeichnet ihr die Tabelle mit Lineal ab – auch die leeren Spalten! Die füllen wir später.", board: merkBoard(3, "Verbfamilie", [S("Jede gebeugte Form gehört zu einer Verbfamilie. Den Familiennamen finde ich mit der Frage: "), S("Wie heißt die Grundform?", { bold: true }), S("  habe → haben, bist → sein")]) }],
  ["luecken", { say: "Jetzt arbeitet ihr allein: Ü2. Lest die drei Schritte mit. Achtung bei Schritt 3: Drei Wörter gehören zu keiner Familie. Vier Minuten.", board: auftrag("Ü2 Wer gehört zu welcher Familie?", ["Schreibe Ü2 an den Rand.", "Zeichne zwei Spalten: Familie haben | Familie sein.", "Sortiere die Wörter ein. Drei Wörter gehören zu keiner Familie – lass sie weg."], [bl("habe   Hand   ist   hat   seit   sind   bist   hart   habt   bin", { after: 0 })], "4 Minuten · allein · leise"), sol: "haben: habe, hat, habt · sein: ist, sind, bist, bin · keine Familie: Hand, seit, hart" }],
  ["tafel", { say: "Kommt nach vorne und schreibt je ein Wort in die richtige Spalte. Alle prüfen mit.", do: ["Bei Hand, seit, hart fragen: „Welche Grundform hat das?“ – es gibt keine, denn das sind gar keine Verben."] }],
  ["kontrolle", { say: "Vergleichen und mit Grün verbessern." }],
]));

// ---------- Phase 3 ----------
doc.push(phase("3", "Früher: hatte und war", 15));
doc.push(...steps([
  ["gemeinsam", { say: "Heute bin ich müde. Gestern …? Heute habe ich Zeit. Gestern …? Meldet euch.", do: ["Die Antworten war und hatte in die „früher“-Spalten schreiben."] }],
  ["tafel", { say: "Wir füllen die leeren Spalten. Wer drankommt, schreibt eine Form an die Tafel.", board: board("Mitte", [famBoard(true)]), do: ["Hervorheben: war sieht noch weniger nach sein aus als ist – trotzdem gleiche Familie."] }],
  ["abschreiben", { say: "Füllt jetzt in eurem Heft die leeren Spalten der Tabelle aus. Dann schreibt ihr Merksatz 4 ab.", board: merkBoard(4, "haben und sein im Präteritum", [S("Im Präteritum heißt es: ich "), S("hatte", { bold: true }), S(" (haben) und ich "), S("war", { bold: true }), S(" (sein). Auch „war“ gehört zur Familie sein!")]) }],
  ["partner", { say: "Jetzt arbeitet ihr zu zweit, ohne Heft. Ich mache es einmal vor.", board: auftrag("Fingerspiel zu zweit", ["Zeige mit den Fingern eine Person: 1 = ich … 6 = sie.", "Sage dazu: haben oder sein – jetzt oder früher.", "Dein Partner nennt die Form. Nach 5 Runden wechselt ihr."], [bl([N("Beispiel: 4 Finger + „sein, früher“ → "), N("wir waren", { bold: true })], { after: 0 })], "3 Minuten · zu zweit · Flüsterstimme"), do: ["Einmal mit einem Schüler vormachen, dann starten."] }],
  ["luecken", { say: "Jetzt arbeitet ihr allein: Ü3. Wichtig: Ihr schreibt die Sätze nicht ab – nur die Nummer und das fehlende Wort. Drei Minuten.", board: auftrag("Ü3 Gestern war alles anders", ["Schreibe Ü3 an den Rand.", "Schreibe untereinander nur die Nummer und das fehlende Wort."], [bl("1  Gestern ___ ich krank. (sein)"), bl("2  Du ___ hohes Fieber. (haben)"), bl("3  Wir ___ im Schwimmbad. (sein)"), bl("4  Ihr ___ keine Zeit. (haben)"), bl("5  ___ du schon einmal in Berlin? (sein)", { after: 0 })], "3 Minuten · allein · leise"), sol: "1 war · 2 hattest · 3 waren · 4 hattet · 5 Warst" }],
  ["kontrolle", { say: "Ich nehme fünf von euch dran, ihr sagt eure Lösung. Alle vergleichen und verbessern mit Grün.", do: ["Lösungswort jeweils in die Lücke an der Tafel schreiben."] }],
]));
doc.push(p([S("Bewegungspause (3 Min.): ", { bold: true }), S("Ansage: „Alle stehen auf. Ich rufe ein Wort – ihr zeigt die Familie.“ Sie rufen Formen, die Klasse zeigt: Hand links = haben, Hand rechts = sein. Wortliste: habe · war · bist · hatten · seid · hattest · ist · hat · waren · bin · habt · wart")], { before: 60, after: 0 }));

// ---------- Phase 4 ----------
doc.push(phase("4", "Partizip II, Perfekt mit haben oder sein, Hilfsverb", 28));
doc.push(h3("4a  Partizip II (10 Min.)", C.hv));
doc.push(...steps([
  ["gemeinsam", { say: "Was habt ihr am Wochenende gemacht? Erzählt in einem Satz. Ihr schreibt noch nichts.", do: ["Sechs Sätze anschreiben – schon heimlich sortiert: haben-Sätze links, sein-Sätze rechts, noch ohne Überschrift. Die Form von haben/sein rot, das Partizip II gelb markieren."],
    board: board("Mitte (Beispiel)", [table([3500, 3500], [
      ...[[["Ich ", "habe", " Fußball ", "gespielt", "."], ["Ich ", "bin", " ins Kino ", "gegangen", "."]], [["Wir ", "haben", " Pizza ", "gegessen", "."], ["Wir ", "sind", " zu Oma ", "gefahren", "."]], [["Sie ", "hat", " ein Buch ", "gelesen", "."], ["Er ", "ist", " spät ", "eingeschlafen", "."]]].map(([a, b]) => row([a, b].map((s) => cell(p([N(s[0]), R(s[1]), N(s[2]), Y(s[3]), N(s[4])], { after: 0 }), { w: 3500, borders: noBorders, m: 15 })))),
    ])]) }],
  ["zuhoeren", { say: "Schaut auf die gelben Wörter. Was haben sie gemeinsam?", do: ["ge-…-t und ge-…-en herausarbeiten. Test zeigen: Passt „ich habe …“ oder „ich bin …“ davor?", "Dann fragen: „Wofür brauchen wir diese Form?“ → Damit haben wir gerade vom Wochenende erzählt, also von etwas, das schon passiert ist."] }],
  ["abschreiben", { say: "Merksatz 5 abschreiben, rot umrahmen.", board: merkBoard(5, "Partizip II", [S("Das Partizip II verändert sich "), S("nie", { bold: true }), S(". Man bildet es meist mit ge-…-t (gespielt) oder ge-…-en (gelaufen). "), S("Wofür? ", { bold: true }), S("Mit haben oder sein erzähle ich, was schon passiert ist (Perfekt): Ich habe gespielt. Ich bin gelaufen. "), S("Test: ", { bold: true }), S("Passt „ich habe …“ oder „ich bin …“ davor?")]) }],
  ["gemeinsam", { say: "Die ersten zwei machen wir zusammen.", do: ["spielen → gespielt, gehen → gegangen gemeinsam an der Tafel lösen."] }],
  ["luecken", { say: "Jetzt arbeitet ihr allein: Ü4. Die ersten zwei schreibt ihr einfach ab, dann macht ihr allein weiter. Drei Minuten.", board: auftrag("Ü4 Das Partizip II", ["Schreibe Ü4 an den Rand.", "Schreibe ab: Verb → Partizip II. Die ersten zwei stehen schon da.", "Markiere ge- und die Endung gelb."], [bl("spielen → gespielt     gehen → gegangen"), bl("kaufen → ______     schreiben → ______     essen → ______     fahren → ______", { after: 0 })], "3 Minuten · allein · leise"), sol: "gekauft · geschrieben · gegessen · gefahren" }],
  ["tafel", { say: "Vier von euch schreiben die Lösung an. Markiert ge- und die Endung gelb.", do: ["Anschließend: „Vergleicht und verbessert mit Grün.“"] }],
]));
doc.push(h3("4b  Perfekt mit haben oder sein (12 Min.)", C.hv));
doc.push(...steps([
  ["gemeinsam", { say: "Schaut noch einmal auf die Wochenend-Sätze. Warum stehen die rechten Sätze zusammen? Was haben sie gemeinsam?", do: ["Kurz vorführen: Ein Kind geht von der Tür zum Fenster („Sie ist gegangen.“), ein anderes bleibt sitzen und liest („Er hat gelesen.“).", "Dann Überschriften „mit haben“ / „mit sein“ über die Spalten schreiben."] }],
  ["abschreiben", { say: "Merksatz 6 abschreiben. Darunter je einen Beispielsatz: einen mit haben, einen mit sein – in Farbe.", board: merkBoard(6, "Perfekt mit haben oder sein", [S("Perfekt = "), S("haben oder sein im Präsens + Partizip II", { bold: true }), S(". Mit sein: Verben der "), S("Bewegung", { bold: true }), S(" von A nach B (gehen, fahren) und der "), S("Veränderung", { bold: true }), S(" (einschlafen, aufwachen). Die meisten anderen Verben: "), S("haben", { bold: true }), S(".")]) }],
  ["luecken", { say: "Jetzt arbeitet ihr allein: Ü5. Wieder nur Nummer und fehlendes Wort. Denkt an den Merksatz: Bewegung oder Veränderung? Vier Minuten.", board: auftrag("Ü5 haben oder sein?", ["Schreibe Ü5 an den Rand.", "Schreibe nur Nummer + Wort: habe, hat, bin, ist oder sind."], [bl("1  Am Samstag ___ ich früh aufgewacht."), bl("2  Dann ___ ich Pfannkuchen gebacken."), bl("3  Am Nachmittag ___ wir zum See gefahren."), bl("4  Mein Bruder ___ Fußball gespielt."), bl("5  Abends ___ wir nach Hause gelaufen."), bl("6  Um neun ___ ich eingeschlafen.", { after: 0 })], "4 Minuten · allein · leise"), sol: "1 bin · 2 habe · 3 sind · 4 hat · 5 sind · 6 bin" }],
  ["tafel", { say: "Sechs von euch schreiben das Wort in die Lücke an der Tafel. Wer drankommt, sagt dazu: Bewegung, Veränderung oder keins von beidem." }],
  ["kontrolle", { say: "Vergleichen und mit Grün verbessern." }],
]));
doc.push(h3("4c  Hilfsverb oder nicht? (6 Min.)", C.hv));
doc.push(...steps([
  ["zuhoeren", { say: "Stifte liegen. Zwei Sätze. Was ist der Unterschied?", board: board("Mitte", [bl([N("Ich "), N("habe", { bold: true }), N(" Hunger.          Ich "), R("habe"), N(" gegessen", { color: "C98A1E", bold: true }), N(".")], { after: 0 })]), do: ["Ziel: Im zweiten Satz steht am Ende ein Partizip II – nur dann ist habe ein Hilfsverb. Im ersten Satz ist habe ein ganz normales Verb (besitzen/fühlen)."] }],
  ["abschreiben", { say: "Merksatz 7 und die beiden Beispielsätze abschreiben.", board: merkBoard(7, "Hilfsverb oder nicht?", [S("haben und sein sind nur dann "), S("Hilfsverben", { bold: true }), S(", wenn am Satzende ein "), S("Partizip II", { bold: true }), S(" steht.")]) }],
  ["gemeinsam", { say: "Ich lese einen Satz vor. Daumen hoch heißt Hilfsverb, Daumen runter heißt kein Hilfsverb. Ihr schreibt nichts.", do: ["Lisa ist sehr müde. (runter) · Lisa ist eingeschlafen. (hoch) · Wir hatten viel Spaß. (runter) · Ihr seid mit dem Bus gekommen. (hoch)"] }],
]));

// ---------- Phase 5 ----------
doc.push(phase("5", "Abschluss und Exit-Ticket", 7));
doc.push(...steps([
  ["zuhoeren", { say: "Wir schauen auf den Ablaufplan. Was haben wir heute gelernt?", do: ["Punkte im Ablaufplan abhaken; zu jedem Punkt einen Schüler den Merksatz in eigenen Worten sagen lassen."] }],
  ["luecken", { say: "Heft zu. Jetzt arbeitet ihr allein auf einem Zettel. Die Aufgabe steht rechts. Vier Minuten.", board: auftrag("Exit-Ticket", ["Nimm einen Zettel und schreibe deinen Namen oben hin.", "Schreibe zu jedem Satz: Familie (haben/sein) · Präsens oder Präteritum · Hilfsverb ja oder nein."], [bl("1  Wir waren im Zoo."), bl("2  Du hast gewonnen."), bl("3  Sie hatten keine Zeit."), bl("4  Ich bin nach Hause gerannt.", { after: 0 })], "4 Minuten · allein · ohne Heft"), sol: "1 sein, Präteritum, nein · 2 haben, Präsens, ja · 3 haben, Präteritum, nein · 4 sein, Präsens, ja" }],
  ["fertig", { say: "Stift hinlegen. Ich sammle ein.", do: ["Auswertung: „ja“ bei 1 oder 3 → Hilfsverb und Vollverb noch verwechselt (Merksatz 7 wiederholen). „Präteritum“ bei 2 oder 4 → Form des Hilfsverbs mit der Zeit des Geschehens verwechselt."] }],
]));
doc.push(p([S("Ausblick nächste Stunde: ", { bold: true }), S("Modalverben + Infinitiv, werden als dritte Hilfsverb-Familie (Futur). Der Ablauf (Ansage – Merksatz – Übung – Vergleich) bleibt gleich und ist dann schon vertraut.")], { before: 120, after: 0 }));

// ================================================================
const pageProps = { page: { size: { width: 11906, height: 16838 }, margin: { top: 900, bottom: 800, left: 850, right: 850, header: 400, footer: 400 } } };
const styles = { default: { document: { run: { font: "Calibri", size: 22 } } } };
const dScript = new Document({ styles, sections: [{ properties: pageProps, headers: { default: header("Verbformen verstehen") }, footers: { default: footer("Tafelskript") }, children: doc }] });
Packer.toBuffer(dScript).then((b) => fs.writeFileSync("Tafelskript_Verbformen_verstehen.docx", b));
