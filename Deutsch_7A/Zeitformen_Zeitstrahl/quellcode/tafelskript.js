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
// Tafelskript „Welche Zeitform wann?“ (Orientierungshospitation, 45 Min.)
// ================================================================
const merkTafel = (nr, title, children) => board(`Seitentafel · Merksatz ${nr} – abschreiben, rot umrahmen`, [
  p([t(title, { bold: true, color: C.hv, font: "Cambria", size: 22 })], { after: 40 }),
  ...children,
]);

const doc = [];
doc.push(kicker("Tafelskript für die Lehrkraft · Einzelstunde (45 Min.) · Unterrichtsbesuch"), h1("Welche Zeitform wann?"));
doc.push(box([
  p([S("Anknüpfung: ", { bold: true }), S("Letzte Stunde lief gut; die Satzstreifen mit Magneten kamen sehr gut an, waren aber zu klein – heute im Großformat (2 pro A4 quer). Auf den Schnipseln war zu wenig Platz – heute große Felder und viel Zeilenabstand. Alle Schüler hatten am Wochenende Wettkämpfe: Die Stunde startet dort. Ü11 wird zur Hausaufgabe über genau dieses Wochenende.")], { after: 0 }),
], LIGHT2));
doc.push(p([S("Ziel: ", { bold: true }), S("Die Klasse ordnet Sätze über ihr Wettkampfwochenende begründet den fünf Zeitformen am Zeitstrahl zu und erkennt, welche Zeitform in welchem Text passt (geschrieben: Präteritum, gesprochen: Perfekt, Vorgeschichte: Plusquamperfekt).")], { before: 120, after: 60 }));
doc.push(p([S("Markieren wie bisher: ", { bold: true }), t("Hilfsverb rot", { bold: true, color: C.hv, size: 21 }), S(" · "), t("Partizip II gelb", { bold: true, color: "C98A1E", size: 21 }), S(" · "), t("Infinitiv blau", { bold: true, color: C.inf, size: 21 }), S(" · Präteritum/Präsens: finites Verb unterstreichen · keine Klammer-Bögen")], { after: 100 }));

doc.push(h3("Ablauf"));
doc.push(grid([900, 3700, 5606], ["Zeit", "Phase", "Kern"], [
  ["4′", "0 Einstieg: unser Wochenende", "2–3 Schüler erzählen einen Satz – meist im Perfekt"],
  ["13′", "1 Magnet-Kette am Zeitstrahl", "8 Streifen hängen, markieren, begründen; Schnipsel mitschreiben"],
  ["7′", "2 Welche Zeitform wann? + Merksatz 10", "Etiketten auswerten, Merksatz abschreiben"],
  ["10′", "3 Ü12 Der Bericht für die Vereinszeitung", "4 Fehler finden, verbessern, einen begründen"],
  ["8′", "4 Kreide-Kette: Fehlerstreifen", "Verbesserung unter die Streifen, mit Grün verbessern"],
  ["3′", "5 Ausblick und Hausaufgabe", "Bericht über dieses Wochenende (Schreibrahmen Ü11)"],
], { bold: [true, true, false], colors: [C.hv, NAVY, null], h: 360 }));
doc.push(p([S("Puffer: ", { bold: true, color: C.hv }), S("Speed-Duell zu zweit (einer liest einen Satz vom Zeitstrahl in einer anderen Zeitform, der andere bestimmt sie). Kürzung: In Phase 4 nur zwei Fehlerstreifen besprechen, den Rest nächste Stunde.")], { before: 80, after: 60 }));

doc.push(h3("So sieht das Heft am Ende aus"));
doc.push(table([2300, 7906], [
  ["Überschrift", "Welche Zeitform wann? · Datum rechts"],
  ["Schnipsel", "Die Zeitformen am Zeitstrahl (Zeitform, Bauplan, Beispielsatz je Abschnitt)"],
  ["Merksatz 10", "Welche Zeitform wann?"],
  ["Ü12", "Schnipsel eingeklebt und verbessert · darunter: ein Fehler begründet (+ Zusatz)"],
  ["Hausaufgabe", "Bericht über das Wettkampfwochenende (Schreibrahmen Ü11)"],
].map(([a, b2]) => row([tc(a, 2300, { bold: true, color: a.startsWith("Merk") ? C.hv : a === "Überschrift" ? NAVY : "C77C12", fill: LIGHT, size: 20 }), tc(b2, 7906, { size: 20 })], 340))));
doc.push(p([S("Material: ", { bold: true }), S("Druckvorlage: Seite 1 Zeitstrahl-Schnipsel (2 pro A4), Seite 2 Ü12 (2 pro A4), Seiten 3–8 zwölf Streifen für die Tafel (einmal drucken, schneiden, Magnete): 8 Satzstreifen und 4 Fehlerstreifen „Bericht · Satz …“. Kreide rot, gelb, blau · Schüler: Heft, Kleber, Buntstifte rot, gelb, blau, grün.")], { before: 100, after: 40 }));
doc.push(p([S("Vorbereitung: ", { bold: true }), S("Zeitstrahl vorab über die ganze Tafel zeichnen (Pfeil, Abschnitte „vorher – Vergangenheit – jetzt – Zukunft“), darüber die Stundenfrage „Welche Zeitform wann?“. Die 8 Satzstreifen gemischt an den Rand hängen, die 4 Fehlerstreifen noch zurückhalten. Merksatz 10 auf die Seitentafel schreiben und zuklappen.")], { after: 0 }));

// ---------- Phase 0 ----------
doc.push(phase("0", "Einstieg: unser Wochenende", 4));
doc.push(...steps([
  ["gemeinsam", { say: "Stifte liegen. Ihr hattet am Wochenende alle Wettkämpfe. Erzählt in einem Satz: Wie lief es?", do: [
    "Zwei bis drei Schüler drannehmen, aus jeder Sportart einen. Die Verbform im Kopf mitnotieren – erwartbar ist Perfekt („Ich habe gewonnen“, „Wir sind Zweiter geworden“).",
  ] }],
  ["zuhoeren", { say: "Ihr habt gerade alle im Perfekt erzählt. In der Vereinszeitung klingt das anders. An der Tafel hängen acht Sätze aus drei Texten über so ein Wochenende: aus einer Sprachnachricht, aus der Vereinszeitung und aus einem Post. Heute klären wir: Welche Zeitform wann?", do: [
    "Auf die Stundenfrage über dem Zeitstrahl zeigen. Heft auf, Datum, Überschrift: Welche Zeitform wann?",
  ] }],
]));

// ---------- Phase 1 ----------
const r = (x) => t(x, { bold: true, color: C.hv, size: 20, font: "Cambria" });
const y = (x) => t(x, { bold: true, color: "C98A1E", size: 20, font: "Cambria" });
const b = (x) => t(x, { bold: true, color: C.inf, size: 20, font: "Cambria" });
const n = (x, o = {}) => t(x, Object.assign({ size: 20, font: "Cambria" }, o));
const u = (x) => n(x, { u: true, bold: true });
const lab = (x) => t(x + "   ", { size: 16, color: MUTED, bold: true });
doc.push(phase("1", "Magnet-Kette am Zeitstrahl", 13));
doc.push(...steps([
  ["tafel", { say: "Ihr bekommt einen Schnipsel. Klebt ihn ein. Wer drankommt, hängt einen Streifen an die richtige Stelle am Zeitstrahl, markiert die Verbformen, nennt die Zeitform und begründet. Alle anderen prüfen mit und tragen die Zeitform und einen Beispielsatz in den Schnipsel ein.", board: board("Mitte · Zeitstrahl mit den Streifen (Lösung)", [
    bl([lab("vorher"), N("Plusquamperfekt", { bold: true, color: NAVY })], { after: 10 }),
    bl([lab("ZEITUNG"), n("Die Judoka "), r("hatten"), n(" wochenlang für das Turnier "), y("trainiert"), n(".")], { after: 10 }),
    bl([lab("ZEITUNG"), n("Die Handballer "), r("waren"), n(" schon am Freitag "), y("angereist"), n(".")], { after: 60 }),
    bl([lab("Vergangenheit"), N("Präteritum", { bold: true, color: NAVY })], { after: 10 }),
    bl([lab("ZEITUNG"), n("Im Stoßen "), u("hob"), n(" unsere Gewichtheberin 75 Kilo.")], { after: 10 }),
    bl([lab("ZEITUNG"), n("Unsere Handballer "), u("kämpften"), n(" bis zur letzten Sekunde.")], { after: 60 }),
    bl([lab("Vergangenheit"), N("Perfekt", { bold: true, color: NAVY })], { after: 10 }),
    bl([lab("SPRACHNACHRICHT"), n("Ich "), r("habe"), n(" im Finale Bronze "), y("geholt"), n("!")], { after: 10 }),
    bl([lab("SPRACHNACHRICHT"), n("Wir "), r("sind"), n(" erst um zehn Uhr nach Hause "), y("gekommen"), n(".")], { after: 60 }),
    bl([lab("jetzt"), N("Präsens", { bold: true, color: NAVY }), n("   ·   "), lab("POST"), n("Heute "), u("analysieren"), n(" wir mit dem Trainer das Video.")], { after: 60 }),
    bl([lab("Zukunft"), N("Futur I", { bold: true, color: NAVY }), n("   ·   "), lab("POST"), n("Beim nächsten Turnier "), r("werde"), n(" ich den Haltegriff besser "), b("verteidigen"), n(".")], { after: 0 }),
  ]), do: [
    "Begründung immer im gleichen Muster einfordern: „Das ist Plusquamperfekt, weil hatte (Hilfsverb im Präteritum) + Partizip II.“",
    "Die Zeitform schreibt der Schüler mit Kreide über seinen Streifen. Präteritum und Perfekt landen beide bei „Vergangenheit“ – das stehen lassen, es ist der Anlass für Phase 2.",
    "Nach jedem Streifen kurz warten, bis alle eingetragen haben. Ein Beispielsatz pro Zeitform im Schnipsel reicht.",
  ] }],
  ["gemeinsam", { say: "Stifte liegen. Jetzt der Bauplan: Woraus besteht jede Zeitform? Meldet euch.", do: [
    "Bauplan unter den Zeitformnamen an die Tafel schreiben, danach eintragen lassen („Jetzt schreibt ihr ab: die Zeile Bauplan.“).",
  ], sol: "Plusquamperfekt: hatte/war + Partizip II · Präteritum: nur finites Verb (Präteritumform) · Perfekt: habe/bin + Partizip II · Präsens: nur finites Verb · Futur I: werde + Infinitiv" }],
]));

// ---------- Phase 2 ----------
doc.push(phase("2", "Welche Zeitform wann? + Merksatz 10", 7));
doc.push(...steps([
  ["gemeinsam", { say: "Stifte liegen. Präteritum und Perfekt hängen am selben Platz – beide erzählen Vergangenes. Schaut auf die Etiketten: Aus welchem Text kommen die Sätze?", do: [
    "Erwartung: Präteritum steht in der Vereinszeitung (geschrieben), Perfekt in der Sprachnachricht (gesprochen). Plusquamperfekt: auch Vereinszeitung – die Vorgeschichte. Präsens und Futur I: Post über heute und morgen.",
    "Rückbezug zum Einstieg: „Deshalb habt ihr vorhin im Perfekt erzählt – ihr habt gesprochen.“",
  ] }],
  ["abschreiben", { say: "Jetzt schreibt ihr von der Seitentafel ab: Merksatz 10. Rot umrahmen.", board: merkTafel(10, "Welche Zeitform wann?", [
    p([S("Was "), S("vorher", { bold: true }), S(" war, steht im "), S("Plusquamperfekt", { bold: true }), S(". Vergangenes "), S("schreibt", { bold: true }), S(" man im "), S("Präteritum", { bold: true }), S(", "), S("mündlich", { bold: true }), S(" erzählt man meist im "), S("Perfekt", { bold: true }), S(". Was jetzt ist: "), S("Präsens", { bold: true }), S(". Was kommt: "), S("Futur I", { bold: true }), S(".")], { after: 60 }),
    p([S("Wofür? ", { bold: true }), S("Damit jeder Text die passende Zeitform hat: der Bericht für die Vereinszeitung im Präteritum, die Sprachnachricht im Perfekt. "), S("Test: ", { bold: true }), S("Wann ist es passiert? Und: Schreibe ich oder spreche ich?")], { after: 0 }),
  ]) }],
]));

// ---------- Phase 3 ----------
doc.push(phase("3", "Ü12 Der Bericht für die Vereinszeitung", 10));
doc.push(...steps([
  ["luecken", { say: "Jetzt arbeitet ihr allein. Die Aufgabe steht rechts. Ein Teamkollege hat seinen Bericht so geschrieben, wie er spricht. Vier Verbformen sind falsch. Sieben Minuten.", board: auftrag("Ü12 Der Bericht für die Vereinszeitung", ["Klebe den Schnipsel Ü12 ins Heft.", "Unterstreiche die vier Verbformen, die falsch sind.", "Schreibe die richtige Form darüber.", "Begründe einen Fehler im Heft: „… ist falsch, weil …“"], [p(S("Zusatz: Was hattest du vor deinem Wettkampf gemacht? Ein Satz im Plusquamperfekt.", { size: 20, italics: true }), { after: 0 })], "7 Minuten · allein · leise"),
    do: ["Herumgehen. Wer bei Satz 4 hängt: „Wie heißt es im Perfekt – habe losgefahren oder bin losgefahren?“"],
    sol: "Satz 1 sind … gefahren → fuhren (Bericht = Präteritum) · Satz 3 hat … gewonnen → gewann · Satz 4 hatten … losgefahren → waren … losgefahren (losfahren bildet mit sein) · Satz 6 werften → warfen (werfen ist unregelmäßig). Sätze 2, 5, 7 sind richtig." }],
]));

// ---------- Phase 4 ----------
doc.push(phase("4", "Kreide-Kette: Fehlerstreifen", 8));
doc.push(...steps([
  ["tafel", { say: "Wer drankommt, schreibt die richtige Verbform unter den Streifen und begründet. Alle prüfen mit. Vergleicht und verbessert mit Grün.", board: board("unter dem Zeitstrahl · vier Fehlerstreifen", [
    bl([lab("SATZ 1"), n("Am Samstag sind unsere Judoka nach Leipzig gefahren."), n("   →  fuhren", { bold: true, color: "2E9E6B" })]),
    bl([lab("SATZ 3"), n("Im Halbfinale hat eine Judoka mit einem Haltegriff gewonnen."), n("   →  gewann", { bold: true, color: "2E9E6B" })]),
    bl([lab("SATZ 4"), n("Unsere Gewichtheber hatten schon um sechs Uhr losgefahren."), n("   →  waren … losgefahren", { bold: true, color: "2E9E6B" })]),
    bl([lab("SATZ 6"), n("Die Handballer werften am Sonntag 28 Tore."), n("   →  warfen", { bold: true, color: "2E9E6B" })], { after: 0 }),
  ]), do: [
    "Begründungen mit Fachbegriffen und Bezug auf Merksatz 10: „Im Bericht steht das Präteritum – sind gefahren ist Perfekt, das sagt man mündlich.“",
    "Wird die Zeit knapp: nur Satz 1 und Satz 4 besprechen (je ein Fehlertyp: Zeitform gewählt / Hilfsverb), den Rest nächste Stunde.",
  ] }],
]));

// ---------- Phase 5 ----------
doc.push(phase("5", "Ausblick und Hausaufgabe", 3));
doc.push(...steps([
  ["zuhoeren", { say: "Stifte liegen. Beantwortet die Stundenfrage in einem Satz: Welche Zeitform wann?", do: [
    "Ein Schüler antwortet. Danach Hausaufgabe nennen: „Schreibt euren Bericht über dieses Wochenende für die Vereinszeitung – im Präteritum, die Vorgeschichte im Plusquamperfekt. Nehmt den Schreibrahmen Ü11.“",
  ] }],
]));
doc.push(p([S("Ausblick: ", { bold: true }), S("Nächste Stunde: Partner-Check der Berichte. Danach: Futur II, Hilfsverb oder Vollverb, Modalverben – und zum Schluss die Satzklammer.")], { before: 120, after: 0 }));

const pageProps = { page: { size: { width: 11906, height: 16838 }, margin: { top: 900, bottom: 800, left: 850, right: 850, header: 400, footer: 400 } } };
const styles = { default: { document: { run: { font: "Calibri", size: 22 } } } };
Packer.toBuffer(new Document({ styles, sections: [{ properties: pageProps, headers: { default: header("Welche Zeitform wann?") }, footers: { default: footer("Tafelskript") }, children: doc }] })).then((buf) => fs.writeFileSync("Tafelskript_Zeitformen_Zeitstrahl.docx", buf));
