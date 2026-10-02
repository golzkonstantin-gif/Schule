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
// Gemeinsame Bausteine für Tafelskripte
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
// Merksatz steht auf der Folie (nicht an der Tafel)
function folie(title, children) {
  return table([TW], [
    row([cell(p(t("FOLIE · " + title, { bold: true, color: "FFFFFF", size: 16 }), { after: 0 }), { w: TW, fill: NAVY, borders: allBorders(solid(NAVY, 12)), m: 30 })]),
    row([cell(children, { w: TW, fill: LIGHT, borders: allBorders(solid(NAVY, 12)), m: 100, ml: 160 })]),
  ]);
}
const merkBoard = (nr, title, runs) => folie(`Merksatz ${nr} – abschreiben, rot umrahmen`, [
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
// Tafelskript „Das Futur I“ (45 Min.)
// ================================================================
const merkTafel = (nr, title, children) => board(`Merksatz ${nr} – abschreiben, rot umrahmen`, [
  p([t(title, { bold: true, color: C.hv, font: "Cambria", size: 22 })], { after: 40 }),
  ...children,
]);
// Beispielsatz mit gezeichneter Satzklammer: [Vorfeld, finit, Mittelfeld, Satzende]
function klammer(vf, fin, mf, end, endCol = C.inf) {
  const w = [1750, 1100, 3550, 1350];
  const K = solid(NAVY, 12);
  const nb = { top: none, bottom: none, left: none, right: none };
  const c = (runs, i, borders) => cell(p(runs, { after: 0, align: AlignmentType.CENTER }), { w: w[i], borders, m: 10, ml: 40 });
  const endRun = t(end, { bold: true, color: endCol, font: "Cambria", size: 22 });
  return table(w, [
    row([c([N(vf)], 0, nb), c([R(fin)], 1, nb), c([N(mf, { size: 21 })], 2, nb), c([endRun, N(".")], 3, nb)]),
    row([c([t("")], 0, nb), c([t("")], 1, { top: none, left: K, bottom: K, right: none }), c([t("")], 2, { top: none, left: none, bottom: K, right: none }), c([t("")], 3, { top: none, left: none, bottom: K, right: K })], 160, HeightRule.EXACT),
    row([c([t("")], 0, nb), c([t("Position 2", { size: 16, italics: true, color: MUTED })], 1, nb), c([t("Satzklammer", { size: 17, bold: true, color: NAVY })], 2, nb), c([t("Satzende", { size: 16, italics: true, color: MUTED })], 3, nb)]),
  ]);
}

const doc = [];
doc.push(kicker("Tafelskript für die Lehrkraft · Einzelstunde (45 Min.)"), h1("Das Futur I"));
doc.push(box([
  p([S("Anknüpfung an die letzte Stunde: ", { bold: true }), S("Die Kreide-Kette mit Mitschreiben lief sehr gut. Merksatz 6 (Das Hilfsverb: haben, sein, werden) steht im Heft. Dort wurde angekündigt: werden kommt beim Futur. Genau hier setzt die Stunde an. Das Futur I ist die erste Zeitform, die die Klasse mit einem Hilfsverb bildet – nur steht am Satzende diesmal kein Partizip II, sondern der Infinitiv.")], { after: 0 }),
], LIGHT2));
doc.push(p([S("Ziel: ", { bold: true }), S("Die Klasse bildet das Futur I (werden + Infinitiv), konjugiert werden sicher, schreibt eigene Sätze über ihr Wochenende im Futur I und baut sie ins Perfekt um. Dabei wird sichtbar: Die Klammer bleibt, nur beide Teile werden ausgetauscht.")], { before: 120, after: 100 }));

doc.push(h3("Ablauf"));
doc.push(grid([900, 3300, 6006], ["Zeit", "Phase", "Kern"], [
  ["7′", "0 Reaktivierung", "Kreide-Kette mit Mitschreiben: Infinitiv, finite Form, Partizip II"],
  ["10′", "1 Wie bildet man das Futur I?", "vier Beispielsätze, werden konjugieren, Klammer zeichnen"],
  ["4′", "2 Merksatz 7", "Das Futur I abschreiben"],
  ["6′", "3 Ü7 a Mein Wochenende", "drei eigene Sätze im Futur I schreiben"],
  ["7′", "4 Sätze an die Tafel", "Lehrkraft schreibt Schülersätze an, Kreide-Kette markiert werden und Infinitiv"],
  ["10′", "5 Ü7 b Zurück ins Perfekt", "„Ich frage euch nächste Woche …“ – Sätze per Kreide-Kette ins Perfekt umbauen"],
], { bold: [true, true, false], colors: [C.hv, NAVY, null], h: 360 }));

doc.push(h3("So sieht das Heft am Ende aus"));
doc.push(table([2300, 7906], [
  ["Überschrift", "Das Futur I (mit Lineal unterstrichen) · Datum rechts"],
  ["Tabelle", "werden im Präsens (ich werde, du wirst …)"],
  ["Merksatz 7", "Das Futur I + Beispielsatz mit Klammer"],
  ["Ü7 a", "drei eigene Sätze zum Wochenende im Futur I, werden rot, Infinitiv blau, Klammer"],
  ["Ü7 b", "die Tafelsätze in zwei Spalten: Futur I | Perfekt (mitgeschrieben)"],
].map(([a, b]) => row([tc(a, 2300, { bold: true, color: a.startsWith("Merk") ? C.hv : a === "Überschrift" ? NAVY : "C77C12", fill: LIGHT, size: 20 }), tc(b, 7906, { size: 20 })], 360))));

doc.push(p([S("Tipp: ", { bold: true }), S("Merksatz 7 und den Wortspeicher für Ü7 vor der Stunde auf die Seitentafel schreiben und zuklappen.")], { before: 120, after: 40 }));
doc.push(p([S("Kreide-Kette: ", { bold: true }), S("Wer die Kreide hat, löst eine Aufgabe an der Tafel und gibt die Kreide weiter. Die übrigen schreiben im Heft mit. Reihenfolge vorher festlegen.")], { after: 40 }));
doc.push(p([S("Material: ", { bold: true }), S("Kreide Rot, Blau, Gelb · Schüler: Heft, Lineal, Buntstifte Rot, Blau, Gelb, Grün")], { after: 40 }));
doc.push(p([S("Farben: ", { bold: true }), t("Hilfsverb rot", { bold: true, color: C.hv, size: 21 }), S(" · "), t("Infinitiv blau", { bold: true, color: C.inf, size: 21 }), S(" · "), t("Partizip II gelb", { bold: true, color: "C98A1E", size: 21 }), S(" · "), t("Verbesserungen grün", { bold: true, color: C.mv, size: 21 })], { after: 0 }));

// ---------- Phase 0 ----------
const WORDS = ["spielen", "gespielt", "spielst", "werde", "gegangen", "gehen", "hat", "wirst", "gekauft", "schlafen"];
doc.push(br(), phase("0", "Reaktivierung: Kreide-Kette mit Mitschreiben", 7));
doc.push(...steps([
  ["abschreiben", { say: "Heft auf. Datum nach rechts. Zeichnet drei Spalten: Infinitiv, finite Form, Partizip II. Wer die Kreide hat, schreibt ein Wort an der Tafel in die richtige Spalte. Alle anderen schreiben im Heft mit.",
    board: auftrag("Kreide-Kette: Welche Verbform?", ["Nimm ein Wort von der Liste.", "Schreibe es in die richtige Spalte.", "Gib die Kreide weiter."], [bl(WORDS.join("  ·  ")), p(t(""), { after: 20 }), table([2400, 2400, 2400], [row([hdr("Infinitiv", 2400, C.inf, 18), hdr("finite Form", 2400, NAVY, 18), hdr("Partizip II", 2400, C.p2, 18)], 320), row([2400, 2400, 2400].map((x) => tc("", x)), 700)])], "7 Minuten · Kreide-Kette, alle schreiben mit"),
    sol: "Infinitiv: spielen, gehen, schlafen · finit: spielst, werde, hat, wirst · Partizip II: gespielt, gegangen, gekauft",
    do: ["Bewusst dabei: werde und wirst. Nachfragen: „Zu welchem Infinitiv gehören werde und wirst?“ → werden. „Was wisst ihr schon über werden?“ → ein Hilfsverb (Merksatz 6)."] }],
]));

// ---------- Phase 1 ----------
const FS = [["Am Samstag ", "werde", " ich lange ", "schlafen", "."], ["Wir ", "werden", " ins Kino ", "gehen", "."], ["Mein Bruder ", "wird", " Fußball ", "spielen", "."], ["", "Wirst", " du mit uns ", "kommen", "?"]];
doc.push(phase("1", "Wie bildet man das Futur I?", 10));
doc.push(...steps([
  ["zuhoeren", { say: "Stifte liegen. Ihr hört nur zu. Ich schreibe vier Sätze an. Wann passiert das – gestern, jetzt oder später?", board: board("Mitte", FS.map((s, i) => bl([N(s[0]), N(s[1]), N(s[2]), N(s[3]), N(s[4])], { after: i === 3 ? 0 : 30 }))), do: ["Antwort: später, in der Zukunft. Begriff nennen: „Diese Zeitform heißt Futur I.“"] }],
  ["tafel", { say: "Kreide-Kette: Wer die Kreide hat, unterstreicht in einem Satz die finite Form von werden rot und das Verb am Satzende blau. Dann Kreide weitergeben. Alle prüfen mit.", board: board("Mitte (danach)", FS.map((s, i) => bl([N(s[0]), R(s[1]), N(s[2]), B(s[3]), N(s[4])], { after: i === 3 ? 0 : 30 }))) }],
  ["gemeinsam", { say: "Meldet euch. Ihr schreibt noch nichts auf. Schaut auf die roten Wörter: Was ist das für ein Verb? Und was steht am Satzende – ein Partizip II wie beim Hilfsverb haben?", do: [
    "Ziel: rot = finite Form von werden, ein Hilfsverb (Merksatz 6). Am Satzende steht kein Partizip II, sondern der Infinitiv.",
    "Bogen vom roten Wort zum blauen Infinitiv zeichnen: Wieder umschließen die beiden den Rest des Satzes.",
    "Satz 4 ist eine Frage: Dort steht werden vorne – die Klammer bleibt trotzdem.",
  ] }],
  ["tafel", { say: "Kreide-Kette: Wir konjugieren werden. Wer die Kreide hat, trägt eine Form ein. Alle anderen schreiben die Tabelle im Heft mit.", board: board("Mitte", [table([1500, 1800, 1500, 1800], [
    ["ich", "werde", "wir", "werden"], ["du", "wirst", "ihr", "werdet"], ["er/sie/es", "wird", "sie/Sie", "werden"],
  ].map((r) => row(r.map((v, i) => cell(p(i % 2 ? R(v) : N(v, { color: MUTED }), { after: 0, align: AlignmentType.CENTER }), { w: [1500, 1800, 1500, 1800][i], borders: allBorders(solid("9AA59C", 4)), m: 20 })))))]),
    sol: "werde · wirst · wird · werden · werdet · werden", do: ["Stolperstellen: du wirst und er wird (kein -d- in wirst, kein -t in wird)."] }],
]));

// ---------- Phase 2 ----------
doc.push(phase("2", "Merksatz 7: Das Futur I", 4));
doc.push(...steps([
  ["abschreiben", { say: "Jetzt schreibt ihr von der Tafel ab: Merksatz 7 mit dem Beispielsatz. werden rot, Infinitiv blau, darunter die Klammer. Rot umrahmen.", board: merkTafel(7, "Das Futur I", [
    p([S("Das "), S("Futur I", { bold: true }), S(" bildet man mit dem Hilfsverb "), S("werden", { bold: true }), S(" (finit, an Position 2) und dem "), S("Infinitiv", { bold: true }), S(" am Satzende.")], { after: 80 }),
    klammer("Am Samstag", "werde", "ich mit meinen Freunden", "zocken"),
    p(t(""), { after: 60 }),
    p([S("Wofür? ", { bold: true }), S("Mit dem Futur I sagt man, was "), S("in der Zukunft", { bold: true }), S(" passieren wird – Pläne und Vorhersagen. Es ist das erste "), S("Tempus", { bold: true }), S(", das wir mit einem Hilfsverb bilden. "), S("Test: ", { bold: true }), S("Steht eine Form von werden an Position 2 und am Satzende ein Infinitiv? Dann ist es Futur I.")], { after: 0 }),
  ]), do: ["Hinweis für die Klasse, falls sie fragt: Im Alltag sagt man oft auch „Morgen gehe ich ins Kino“ (Präsens mit Zeitangabe). Das Futur I macht die Zukunft eindeutig."] }],
]));

// ---------- Phase 3 ----------
doc.push(phase("3", "Ü7 a Mein Wochenende", 6));
doc.push(...steps([
  ["luecken", { say: "Jetzt arbeitet ihr allein. Die Aufgabe steht rechts an der Tafel: Ü7 a. Schreibt drei Sätze darüber, was ihr am Wochenende machen werdet. Fünf Minuten. Wer fertig ist: Stift hinlegen.", board: auftrag("Ü7 a Mein Wochenende", ["Schreibe Ü7 a an den Rand.", "Schreibe drei Sätze: Was wirst du am Wochenende machen? Benutze in jedem Satz das Futur I.", "Markiere werden rot und den Infinitiv blau.", "★ Beginne nicht jeden Satz mit „Ich“ – z. B. „Am Samstag …“, „Danach …“."], [
    p([S("Wortspeicher: ", { bold: true, size: 20 }), N("ausschlafen · zocken · trainieren · Freunde treffen · ins Kino gehen · eine Serie schauen · Oma besuchen · Hausaufgaben machen · Rad fahren", { size: 20 })], { after: 0 }),
  ], "5 Minuten · allein · leise"), do: [
    "Herumgehen und auf typische Fehler achten: „Ich werde spiele“ (finite Form statt Infinitiv), „du wirdst“ / „er werdet“, Infinitiv nicht am Satzende.",
    "Dabei schon passende Sätze für Phase 4 auswählen: möglichst einige mit Bewegungsverben (gehen, fahren) und einige ohne (zocken, schauen) – das macht den Perfekt-Umbau später ergiebig.",
  ] }],
]));

// ---------- Phase 4 ----------
doc.push(phase("4", "Schülersätze an die Tafel", 7));
const SB = [["Am Samstag ", "werde", " ich ", "ausschlafen", "."], ["Ich ", "werde", " mit Freunden ", "zocken", "."], ["Wir ", "werden", " ins Kino ", "gehen", "."], ["Mein Bruder ", "wird", " Fußball ", "spielen", "."], ["Am Sonntag ", "werde", " ich zu Oma ", "fahren", "."]];
doc.push(...steps([
  ["gemeinsam", { say: "Meldet euch. Lest einen eurer Sätze vor. Ich schreibe fünf davon an die Tafel.", do: ["Fünf Sätze in die linke Spalte schreiben (Überschrift „Futur I“), noch ohne Farben. Rechts eine leere Spalte „Perfekt“ lassen.", "Falls passende Sätze fehlen: Beispiele unten ergänzen."],
    board: board("Mitte (Beispiel)", SB.map((x, i) => bl([N(x[0]), N(x[1]), N(x[2]), N(x[3]), N(x[4])], { after: i === SB.length - 1 ? 0 : 30 }))) }],
  ["tafel", { say: "Kreide-Kette: Wer die Kreide hat, markiert in einem Satz werden rot und den Infinitiv blau und zeichnet die Klammer. Alle anderen schreiben die Sätze als Ü7 b in die linke Spalte im Heft mit.", do: ["Heft-Ansage vorher: „Zeichnet zwei Spalten: links Futur I, rechts Perfekt. Die rechte bleibt erst einmal leer.“"] }],
]));

// ---------- Phase 5 ----------
const PERF = [["Am Samstag ", "habe", " ich ", "ausgeschlafen", "."], ["Ich ", "habe", " mit Freunden ", "gezockt", "."], ["Wir ", "sind", " ins Kino ", "gegangen", "."], ["Mein Bruder ", "hat", " Fußball ", "gespielt", "."], ["Am Sonntag ", "bin", " ich zu Oma ", "gefahren", "."]];
doc.push(phase("5", "Ü7 b Zurück ins Perfekt", 10));
doc.push(...steps([
  ["zuhoeren", { say: "Stifte liegen. Ihr hört nur zu. Stellt euch vor, es ist Montag, und ich frage euch: Was habt ihr am Wochenende gemacht? Was müsst ihr an den Sätzen ändern?", do: ["Erwartung: Das Wochenende ist vorbei → Vergangenheit → Perfekt. Statt werde kommt habe oder bin, statt des Infinitivs das Partizip II."] }],
  ["tafel", { say: "Kreide-Kette: Wer die Kreide hat, schreibt einen Satz rechts ins Perfekt – Hilfsverb rot, Partizip II gelb, Klammer drunter. Alle anderen schreiben im Heft in der rechten Spalte mit. Alle prüfen mit.",
    board: board("Mitte · zwei Spalten", [table([3900, 3900], [
      row([cell(p(N("Futur I", { bold: true, color: NAVY }), { after: 0 }), { w: 3900, borders: noBorders }), cell(p(N("Perfekt", { bold: true, color: NAVY }), { after: 0 }), { w: 3900, borders: noBorders })]),
      ...SB.map((f, i) => row([f, PERF[i]].map((x, j) => cell(p([N(x[0], { size: 19 }), t(x[1], { bold: true, color: C.hv, font: "Cambria", size: 19 }), N(x[2], { size: 19 }), t(x[3], { bold: true, color: j ? "C98A1E" : C.inf, font: "Cambria", size: 19 }), N(x[4], { size: 19 })], { after: 0 }), { w: 3900, borders: noBorders, m: 15 })))),
    ])]),
    sol: "habe ausgeschlafen · habe gezockt · sind gegangen · hat gespielt · bin gefahren",
    do: [
      "Bei gehen und fahren hakt es vermutlich (haben oder sein?). Kurz klären, ohne neue Regel: „Hört sich ‚wir haben ins Kino gegangen‘ richtig an? Bei Bewegung von A nach B nimmt man sein.“ Die Regel kommt ausführlich in einer späteren Stunde.",
      "ausschlafen → ausgeschlafen: ge- steht in der Mitte (trennbares Verb).",
    ] }],
  ["gemeinsam", { say: "Meldet euch. Vergleicht die beiden Spalten: Was ist gleich geblieben, was hat sich verändert?", do: [
    "Erkenntnis: Die Klammer ist geblieben – links ein Hilfsverb an Position 2, rechts der Partner am Satzende. Ausgetauscht wurden beide Teile: werde → habe/bin, Infinitiv → Partizip II.",
    "Merkhilfe für die Zeitform: „Schau auf beide Enden der Klammer: werde + Infinitiv = Futur I, habe/bin + Partizip II = Perfekt.“",
  ] }],
  ["kontrolle", { say: "Vergleicht eure rechte Spalte mit der Tafel und verbessert mit Grün." }],
]));
doc.push(p([S("Ausblick: ", { bold: true }), S("Offen sind der Merksatz zur Satzklammer, das Perfekt als eigener Merksatz (haben oder sein: Bewegung, Veränderung) und „Hilfsverb oder Vollverb“. Die zwei Spalten von heute sind dafür die ideale Grundlage: Den Merksatz zur Satzklammer kann die Klasse in der nächsten Stunde direkt aus ihnen ableiten.")], { before: 120, after: 0 }));

const pageProps = { page: { size: { width: 11906, height: 16838 }, margin: { top: 900, bottom: 800, left: 850, right: 850, header: 400, footer: 400 } } };
const styles = { default: { document: { run: { font: "Calibri", size: 22 } } } };
const dScript = new Document({ styles, sections: [{ properties: pageProps, headers: { default: header("Das Futur I") }, footers: { default: footer("Tafelskript") }, children: doc }] });
Packer.toBuffer(dScript).then((b) => fs.writeFileSync("Tafelskript_Futur_I.docx", b));
