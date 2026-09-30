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
// Tafelskript „Hilfsverben erkennen“ (45 Min.)
// ================================================================

// Beispielsatz mit gezeichneter Satzklammer (Bogen über Tabellenränder)
function klammerBeispiel() {
  const w = [600, 1150, 4550, 1450];
  const K = solid(NAVY, 12);
  const nb = { top: none, bottom: none, left: none, right: none };
  const c = (runs, i, borders, align = AlignmentType.CENTER) => cell(p(runs, { after: 0, align }), { w: w[i], borders, m: 10, ml: 40 });
  return table(w, [
    row([c([N("Ich")], 0, nb), c([R("habe")], 1, nb), c([N("am Samstag mit meinem Bruder Fußball", { size: 20 })], 2, nb), c([Y("gespielt"), N(".")], 3, nb)]),
    row([c([t("")], 0, nb), c([t("")], 1, { top: none, left: K, bottom: K, right: none }), c([t("")], 2, { top: none, left: none, bottom: K, right: none }), c([t("")], 3, { top: none, left: none, bottom: K, right: K })], 160, HeightRule.EXACT),
    row([c([t("")], 0, nb), c([t("Position 2", { size: 16, italics: true, color: MUTED })], 1, nb), c([t("Satzklammer", { size: 17, bold: true, color: NAVY })], 2, nb), c([t("Satzende", { size: 16, italics: true, color: MUTED })], 3, nb)]),
  ]);
}
// In dieser Stunde ohne Präsentation: Merksätze stehen an der Tafel
const merkTafel = (nr, title, runs) => board(`Merksatz ${nr} – abschreiben, rot umrahmen`, [
  p([t(title, { bold: true, color: C.hv, font: "Cambria", size: 22 })], { after: 40 }),
  p(runs, { after: 0 }),
]);
const doc = [];
doc.push(kicker("Tafelskript für die Lehrkraft · Einzelstunde (45 Min.)"), h1("Hilfsverben erkennen"));
doc.push(box([
  p([S("Anknüpfung an die letzte Stunde: ", { bold: true }), S("Den Unterschied zwischen Infinitiv und finiter Form hat die Klasse verstanden. Das Partizip II wurde bestimmt, der Merksatz dazu fehlt noch. Die Kreide-Kette hat für eine konzentrierte Atmosphäre gesorgt. Deshalb: Reaktivierung per Kreide-Kette, Merksatz 5 nachholen, dann das Hilfsverb aus dem Partizip II heraus erarbeiten. Die Tempora folgen erst, wenn das Hilfsverb sitzt.")], { after: 0 }),
], LIGHT2));
doc.push(p([S("Ziel: ", { bold: true }), S("Die Klasse weiß, was ein Hilfsverb ist (haben, sein, werden; finit an Position 2; bildet mit dem Partizip II am Satzende eine Klammer) und unterscheidet Hilfsverb und Vollverb.")], { before: 120, after: 100 }));

doc.push(h3("Ablauf"));
doc.push(grid([900, 3300, 6006], ["Zeit", "Phase", "Kern"], [
  ["5′", "0 Reaktivierung", "Kreide-Kette: Wörter nach Infinitiv, finit, Partizip II sortieren"],
  ["5′", "1 Merksatz Partizip II", "Merksatz 5 von der Tafel nachholen"],
  ["18′", "2 Hilfsverb und Satzklammer", "„Ich Fußball gespielt.“ – Was fehlt? Hilfsverb erarbeiten, Merksatz 6 (Hilfsverb), Merksatz 7 (Satzklammer), Ü6"],
  ["12′", "3 Hilfsverb oder Vollverb", "„Ich habe Hunger“ / „Ich habe gegessen“, Merksatz 8, Ü7"],
  ["5′", "4 Exit-Ticket", "drei Sätze auf einem Zettel, dazu eine Erklärung in eigenen Worten"],
], { bold: [true, true, false], colors: [C.hv, NAVY, null], h: 380 }));

doc.push(h3("So sieht das Heft am Ende aus"));
doc.push(table([2300, 7906], [
  ["Überschrift", "Hilfsverben erkennen (mit Lineal unterstrichen) · Datum rechts"],
  ["Merksatz 5", "Partizip II"],
  ["Merksatz 6", "Das Hilfsverb"],
  ["Merksatz 7", "Die Satzklammer + farbiger Beispielsatz mit Bogen"],
  ["Ü6", "nur Nummer + Hilfsverb + Partizip II: 1 hat – gezockt …"],
  ["Merksatz 8", "Hilfsverb oder Vollverb + Beispielpaar"],
  ["Ü7", "nur Nummer + H oder V, dazu ein Begründungssatz"],
].map(([a, b]) => row([tc(a, 2300, { bold: true, color: a.startsWith("Merk") ? C.hv : a === "Überschrift" ? NAVY : "C77C12", fill: LIGHT, size: 20 }), tc(b, 7906, { size: 20 })], 360))));

doc.push(p([S("Kreide-Kette: ", { bold: true }), S("Wer die Kreide hat, löst eine Aufgabe an der Tafel und gibt die Kreide an den Nächsten weiter. Reihenfolge vorher festlegen (z. B. nach Sitzreihe), dann läuft es ohne Pausen. Alle anderen prüfen mit.")], { before: 120, after: 40 }));
doc.push(p([S("Tipp: ", { bold: true }), S("Die vier Merksätze vor der Stunde auf die Seitentafel schreiben und zuklappen. Dann kostet das Anschreiben in der Stunde keine Zeit, und Sie klappen jeweils nur den nächsten Merksatz auf.")], { after: 40 }));
doc.push(p([S("Material: ", { bold: true }), S("Kreide Rot und Gelb · Schüler: Heft, Lineal, Buntstifte Rot und Gelb, kleiner Zettel für das Exit-Ticket")], { after: 40 }));
doc.push(p([S("Farben: ", { bold: true }), t("Hilfsverb rot", { bold: true, color: C.hv, size: 21 }), S(" · "), t("Partizip II gelb", { bold: true, color: "C98A1E", size: 21 }), S(" · "), t("Infinitiv blau", { bold: true, color: C.inf, size: 21 }), S(" · "), t("Verbesserungen grün", { bold: true, color: C.mv, size: 21 })], { after: 0 }));

// ---------- Phase 0 ----------
const WORDS = ["spielen", "gespielt", "spielst", "war", "gegangen", "gehen", "hat", "bist", "gekauft"];
doc.push(br(), phase("0", "Reaktivierung: Kreide-Kette", 5));
doc.push(...steps([
  ["zuhoeren", { say: "Stifte liegen. Ihr hört nur zu. Letzte Stunde habt ihr drei Verbformen kennengelernt: Infinitiv, finite Form und Partizip II. Das holen wir jetzt an die Tafel – mit der Kreide-Kette.", do: ["Reihenfolge der Kreide-Kette festlegen (z. B. vorderste Reihe von links nach rechts)."] }],
  ["tafel", { say: "Wer die Kreide hat, nimmt ein Wort von der Liste, schreibt es in die richtige Spalte und gibt die Kreide weiter. Alle anderen prüfen mit.",
    board: [auftrag("Kreide-Kette: Welche Verbform?", ["Nimm ein Wort von der Liste.", "Schreibe es in die richtige Spalte.", "Gib die Kreide weiter."], [bl(WORDS.join("  ·  ")), p(t(""), { after: 20 }), table([2400, 2400, 2400], [row([hdr("Infinitiv", 2400, C.inf, 18), hdr("finite Form", 2400, NAVY, 18), hdr("Partizip II", 2400, C.p2, 18)], 320), row([2400, 2400, 2400].map((x) => tc("", x)), 700)])], "5 Minuten · Kreide-Kette")],
    sol: "Infinitiv: spielen, gehen · finit: spielst, war, hat, bist · Partizip II: gespielt, gegangen, gekauft" }],
  ["gemeinsam", { say: "Meldet euch. Woran habt ihr das Partizip II erkannt? Und woran die finite Form?", do: ["Partizip II: ge-…-t / ge-…-en, verändert sich nicht. Finite Form: verändert sich mit der Person (Test: ich → wir)."] }],
]));

// ---------- Phase 1 ----------
doc.push(phase("1", "Merksatz Partizip II nachholen", 5));
doc.push(...steps([
  ["abschreiben", { say: "Heft auf. Datum nach rechts. Neue Überschrift: Hilfsverben erkennen. Mit Lineal unterstreichen. Dann schreibt ihr von der Tafel ab: Merksatz 5. Rot umrahmen.", board: merkTafel(5, "Partizip II", [S("Das Partizip II ist infinit und verändert sich "), S("nie", { bold: true }), S(". Man bildet es meist mit ge-…-t (gespielt) oder ge-…-en (gelaufen). Verben auf be-, ver-, -ieren bekommen kein ge- (bestellt, verstanden, telefoniert). "), S("Wofür? ", { bold: true }), S("Das Partizip II braucht man für die "), S("Tempora", { bold: true }), S(" (Zeitformen) – aber nie allein, sondern immer zusammen mit einem anderen Verb: Ich habe gespielt. "), S("Test: ", { bold: true }), S("Passt „ich habe …“ oder „ich bin …“ davor?")]), do: ["„Für die Tempora – aber nie allein“ bewusst stehen lassen: Das ist die Brücke zu Phase 2. Welches Verb hilft, erarbeitet die Klasse gleich selbst."] }],
]));

// ---------- Phase 2 ----------
const HSATZ = [["Ich ", "habe", " Fußball ", "gespielt", "."], ["Wir ", "sind", " ins Kino ", "gegangen", "."], ["Lena ", "hat", " eine Pizza ", "bestellt", "."], ["Du ", "bist", " zu spät ", "gekommen", "."]];
doc.push(phase("2", "Hilfsverb und Satzklammer", 18));
doc.push(...steps([
  ["zuhoeren", { say: "Stifte liegen. Ihr hört nur zu. Ich schreibe vier Sätze an. Stimmt da etwas nicht?", board: board("Mitte", HSATZ.map((s, i) => bl([N(s[0]), N("_____  ", { color: MUTED }), N(s[2]), Y(s[3]), N(s[4])], { after: i === 3 ? 0 : 30 }))), do: ["Die Sätze ohne das Hilfsverb anschreiben (Lücke lassen), Partizip II schon gelb. Antwort sammeln: Es fehlt ein Wort – der Satz sagt nicht, wer es gemacht hat bzw. er klingt falsch."] }],
  ["tafel", { say: "Wer die Kreide hat, setzt das fehlende Wort ein – in Rot – und gibt die Kreide weiter. Alle prüfen mit.", board: board("Mitte (danach)", HSATZ.map((s, i) => bl([N(s[0]), R(s[1]), N(s[2]), Y(s[3]), N(s[4])], { after: i === 3 ? 0 : 30 }))), sol: "habe · sind · hat · bist" }],
  ["gemeinsam", { say: "Meldet euch. Ihr schreibt noch nichts auf. Schaut auf die roten Wörter: Sind sie finit oder infinit? Wo stehen sie im Satz? Und was steht am Satzende?", do: [
    "Ziel: Die roten Wörter sind finit (sie ändern sich mit der Person: ich habe – wir haben) und stehen an Position 2. Am Satzende steht das Partizip II.",
    "Einen Bogen von jedem roten Wort zum gelben Partizip zeichnen: Die beiden umschließen den Rest des Satzes wie eine Klammer. Den Begriff Satzklammer schon nennen – der Merksatz dazu folgt gleich nach dem Hilfsverb.",
    "Begriff einführen: „Das Partizip II kann nicht zeigen, wer etwas tut. Das rote Wort hilft ihm – deshalb heißt es Hilfsverb.“",
    "Kurz erwähnen: Es gibt drei Hilfsverben – haben, sein und werden. werden kommt später beim Futur.",
  ] }],
  ["abschreiben", { say: "Jetzt schreibt ihr von der Tafel ab: Merksatz 6. Rot umrahmen.", board: merkTafel(6, "Das Hilfsverb", [S("Hilfsverben sind "), S("haben, sein und werden", { bold: true }), S(". Das Hilfsverb ist "), S("finit", { bold: true }), S(" und steht an Position 2. "), S("Wofür? ", { bold: true }), S("Hilfsverben braucht man für die "), S("Tempora", { bold: true }), S(" (Zeitformen). Das Partizip II kann nicht zeigen, wer etwas tut – das übernimmt das Hilfsverb. Zusammen bilden sie die Zeitform. "), S("Test: ", { bold: true }), S("Steht am Satzende ein Partizip II? Dann ist das finite haben oder sein davor ein Hilfsverb.")]) }],
  ["abschreiben", { say: "Jetzt schreibt ihr von der Tafel ab: Merksatz 7 mit dem Beispielsatz. Hilfsverb rot, Partizip gelb, darunter den Bogen. Rot umrahmen.", board: board("Merksatz 7 – abschreiben, rot umrahmen", [
    p([t("Die Satzklammer", { bold: true, color: C.hv, font: "Cambria", size: 22 })], { after: 40 }),
    p([S("Die Satzklammer", { bold: true }), S(" besteht aus dem "), S("finiten Verb an Position 2", { bold: true }), S(" und dem "), S("infiniten Teil am Satzende", { bold: true }), S(". Zusammen umschließen sie den Rest des Satzes.")], { after: 80 }),
    klammerBeispiel(),
    p(t(""), { after: 60 }),
    p([S("Wofür? ", { bold: true }), S("Die Satzklammer braucht man für die "), S("Tempora", { bold: true }), S(": Am linken Teil erkennt man die Zeitform, am rechten den Partner. So findet man auch in langen Sätzen das ganze Prädikat. "), S("Test: ", { bold: true }), S("Suche das finite Verb an Position 2. Suche dann am Satzende seinen Partner. Beide zusammen bilden die Satzklammer.")], { after: 0 }),
  ]), do: ["Das Beispiel hat bewusst ein langes Mittelfeld: So sieht man, dass die Klammer auch weit auseinanderliegende Teile verbindet.", "Auf der Tafel den Bogen unter dem Satz von habe bis gespielt zeichnen, darüber klein „Position 2“ und „Satzende“."] }],
  ["luecken", { say: "Jetzt arbeitet ihr allein. Die Aufgabe steht rechts an der Tafel: Ü6. Ihr schreibt die Sätze nicht ab – nur Nummer, Hilfsverb und Partizip II. Vier Minuten.", board: auftrag("Ü6 Hilfsverb gesucht", ["Schreibe Ü6 an den Rand.", "Finde in jedem Satz das Hilfsverb und das Partizip II.", "Schreibe nur: Nummer, Hilfsverb – Partizip II (Hilfsverb rot, Partizip gelb)."], [
    bl("1  Mein Bruder hat gestern bis Mitternacht gezockt."), bl("2  Wir sind mit dem Bus zur Schule gefahren."), bl("3  Ihr habt die Hausaufgaben vergessen."), bl("4  Lea ist beim Training hingefallen."), bl("5  Ich habe mir ein neues Handy gekauft."), bl("6  Die Party hat um acht begonnen.", { after: 0 })], "4 Minuten · allein · leise"),
    sol: "1 hat – gezockt · 2 sind – gefahren · 3 habt – vergessen · 4 ist – hingefallen · 5 habe – gekauft · 6 hat – begonnen (kein ge-, Verb beginnt mit be-)" }],
  ["tafel", { say: "Kreide-Kette: Wer die Kreide hat, unterstreicht in einem Satz das Hilfsverb rot und das Partizip gelb und zeichnet die Klammer. Dann Kreide weitergeben.", do: ["Die Sätze von Ü6 stehen rechts an der Tafel."] }],
  ["kontrolle", { say: "Vergleicht und verbessert mit Grün." }],
]));

// ---------- Phase 3 ----------
doc.push(phase("3", "Hilfsverb oder Vollverb?", 12));
doc.push(...steps([
  ["zuhoeren", { say: "Stifte liegen. Ihr hört nur zu. Zwei Sätze, zweimal ‚habe‘. Ist es beide Male ein Hilfsverb?", board: board("Mitte", [bl([N("Ich "), N("habe", { bold: true }), N(" Hunger.          Ich "), R("habe"), N(" "), Y("gegessen"), N(".")], { after: 0 })]), do: ["Test aus Merksatz 6 anwenden lassen: Nur im zweiten Satz steht am Ende ein Partizip II → Hilfsverb.", "Im ersten Satz trägt habe selbst die Bedeutung (Hunger haben = spüren) → Vollverb. Genau diese Verwechslung war der Stolperstein in der Bastelstunde."] }],
  ["abschreiben", { say: "Jetzt schreibt ihr von der Tafel ab: Merksatz 8 mit den beiden Beispielsätzen.", board: merkTafel(8, "Hilfsverb oder Vollverb?", [S("haben und sein sind nur dann "), S("Hilfsverben", { bold: true }), S(", wenn am Satzende ein "), S("Partizip II", { bold: true }), S(" steht. Sonst sind sie "), S("Vollverben", { bold: true }), S(" und tragen selbst die Bedeutung. "), S("Wofür? ", { bold: true }), S("Wer Hilfsverben erkennt, kann später die Zeitform bestimmen. "), S("Test: ", { bold: true }), S("Steht am Satzende ein Partizip II?")]) }],
  ["luecken", { say: "Jetzt arbeitet ihr allein: Ü7. Nur Nummer und H oder V – und eine Begründung. Drei Minuten.", board: auftrag("Ü7 Hilfsverb oder Vollverb?", ["Schreibe Ü7 an den Rand.", "Schreibe zu jeder Nummer H (Hilfsverb) oder V (Vollverb).", "Begründe bei Nr. 2 in einem Satz."], [
    bl("1  Mein Handy ist kaputt."), bl("2  Mein Handy ist runtergefallen."), bl("3  Wir hatten keine Hausaufgaben."), bl("4  Ihr habt die Hausaufgaben vergessen."), bl("5  Ich bin total müde."), bl("6  Ich bin um elf eingeschlafen.", { after: 0 })], "3 Minuten · allein · leise"),
    sol: "1 V · 2 H · 3 V · 4 H · 5 V · 6 H. Begründung 2: Am Satzende steht das Partizip II runtergefallen, also ist ist ein Hilfsverb." }],
  ["tafel", { say: "Kreide-Kette: Wer die Kreide hat, schreibt H oder V hinter einen Satz und sagt den Grund. Dann Kreide weitergeben." }],
  ["kontrolle", { say: "Vergleicht und verbessert mit Grün." }],
]));

// ---------- Phase 4 ----------
doc.push(phase("4", "Exit-Ticket", 5));
doc.push(...steps([
  ["luecken", { say: "Heft zu. Jetzt arbeitet ihr allein auf einem Zettel. Die Aufgabe steht rechts an der Tafel. Vier Minuten.", board: auftrag("Exit-Ticket", ["Nimm einen Zettel und schreibe deinen Namen oben hin.", "Schreibe zu jedem Satz: Hilfsverb ja oder nein – und wenn ja, das Partizip II.", "Erkläre in einem Satz: Was ist ein Hilfsverb?"], [bl("1  Wir waren gestern im Kino."), bl("2  Wir sind gestern ins Kino gegangen."), bl("3  Sie hat einen Hund.", { after: 0 })], "4 Minuten · allein · ohne Heft"),
    sol: "1 nein (Vollverb) · 2 ja – gegangen · 3 nein (Vollverb). Erklärung z. B.: „Ein Hilfsverb ist haben, sein oder werden; es ist finit und bildet mit dem Partizip II eine Zeitform.“" }],
  ["fertig", { say: "Stift hinlegen. Ich sammle ein.", do: ["Auswertung: „ja“ bei 1 oder 3 → Hilfsverb und Vollverb noch verwechselt (Merksatz 8 in der nächsten Stunde kurz wiederholen). Die Erklärung zeigt, ob der Begriff sitzt – Voraussetzung für die Tempora."] }],
]));
doc.push(p([S("Ausblick nächste Stunde (Tempora): ", { bold: true }), S("Das Perfekt ist schon da: Hilfsverb im Präsens + Partizip II. Darauf aufbauend: haben oder sein (Bewegung, Veränderung), das Plusquamperfekt über hatte/war + Partizip II, später Futur mit werden. Der Chat aus der letzten Stunde („ich habe nach Hause gegangen“) passt als Einstieg – die Klasse hat die Fehler gefunden, aber noch nicht erklärt.")], { before: 120, after: 0 }));

const pageProps = { page: { size: { width: 11906, height: 16838 }, margin: { top: 900, bottom: 800, left: 850, right: 850, header: 400, footer: 400 } } };
const styles = { default: { document: { run: { font: "Calibri", size: 22 } } } };
const dScript = new Document({ styles, sections: [{ properties: pageProps, headers: { default: header("Hilfsverben erkennen") }, footers: { default: footer("Tafelskript") }, children: doc }] });
Packer.toBuffer(dScript).then((b) => fs.writeFileSync("Tafelskript_Hilfsverben_erkennen.docx", b));
