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
// Tafelskript „Paddy Fingal – Lesetechniken“ (45 Min., LB S. 136 f.) – AVIVA
// ================================================================
const avivaPhase = (k, title, min) => p([t(`${k}  `, { font: "Cambria", size: 28, bold: true, color: C.hv }), t(title, { font: "Cambria", size: 28, bold: true, color: NAVY }), t(`   ${min} Min.`, { size: 20, bold: true, color: "C77C12" })], { before: 200, after: 100, keepNext: true });
const folieBox = (title, children) => folie(title, children);

const doc = [];
doc.push(kicker("Tafelskript für die Lehrkraft · Einzelstunde (45 Min.) · Planungsmodell AVIVA · mit Präsentation"), h1("Paddy Fingal – Lesetechniken"));
doc.push(box([
  p([S("Bezug: ", { bold: true }), S("Einstieg in die Balladen-Einheit mit „Paddy Fingal“ (Wilhelm Brandes, Lehrbuch S. 136 f.). Es geht um einen ersten Leseeindruck, den ungefähren Inhalt und die drei Lesetechniken. LB-Aufgabe 1 ist kein Muss (mit * markiert). Aus der Rückmeldung zum Unterrichtsbesuch: Ziel und Ablauf transparent, jeder hat in jeder Phase einen eigenen Auftrag.")], { after: 0 }),
], LIGHT2));
doc.push(p([S("Lernziel: ", { bold: true }), S("Ich kann je nach Leseziel überfliegend, orientierend oder genau lesen und meinen ersten Eindruck von einer Ballade begründen.", { italics: true })], { before: 120, after: 40 }));
doc.push(p([S("Antwort auf „Was habt ihr heute gelernt?“: ", { bold: true }), S("„Es gibt drei Lesetechniken: überfliegen, um zu wissen, worum es geht – orientierend lesen, um eine Information zu finden – genau lesen, um alles zu verstehen. Paddy Fingal ist eine lustige Ballade über einen Riesen, der sich mit einer List rettet.“")], { after: 100 }));

doc.push(h3("Ablauf"));
doc.push(grid([900, 1300, 3400, 4606], ["Zeit", "AVIVA", "Phase", "Alle Schüler …"], [
  ["4′", "A", "Ankommen: Wer ist Paddy Fingal?", "schreiben eine Vermutung zum Titel und Bild"],
  ["4′", "V", "Vorwissen: Wie lest ihr?", "ordnen drei Situationen eine Leseweise zu"],
  ["6′", "I", "Merksatz 1: Lesetechniken", "schreiben ab"],
  ["22′", "V", "Ü1 Paddy Fingal lesen", "überfliegen – zuhören und Kehrreim mitsprechen – orientierend – genau"],
  ["5′", "A", "Auswerten: Ziel erreicht?", "lösen drei Situationen, schreiben „Heute habe ich gelernt, …“"],
], { bold: [true, true, true, false], colors: [C.hv, C.hv, NAVY, null], h: 360 }));
doc.push(p([S("Puffer: ", { bold: true, color: C.hv }), S("4 Minuten. Mit * markiert: Frage 5 in Ü1 b) und LB-Aufgabe 1 (Fragen beantworten, mündlich nacherzählen) – nur wenn Zeit bleibt, sonst nächste Stunde.")], { before: 80, after: 60 }));

doc.push(h3("So sieht das Heft am Ende aus"));
doc.push(table([2300, 7906], [
  ["Einheit", "Leute, höret die Geschicht’ – Balladen (neue Seite)"],
  ["Überschrift", "Paddy Fingal – Lesetechniken · Datum rechts"],
  ["Heft", "Vermutung zum Titel · Wie lest ihr? (A, B, C)"],
  ["Merksatz 1", "Lesetechniken (rot umrahmt)"],
  ["Ü1", "Paddy Fingal lesen: a) Satz und erster Eindruck · b) 4–5 Antworten mit Strophe · c) Stichpunkte zu Str. 7–9 und die List"],
  ["Abschluss", "„Heute habe ich gelernt, …“"],
].map(([a, b2]) => row([tc(a, 2300, { bold: true, color: a.startsWith("Merk") ? C.hv : a === "Überschrift" || a === "Einheit" ? NAVY : "C77C12", fill: LIGHT, size: 20 }), tc(b2, 7906, { size: 20 })], 340))));
doc.push(p([S("Material: ", { bold: true }), S("Lehrbuch S. 136 f. · Präsentation „Lesetechniken“ (Merksatz und Aufgaben, Lösungen in den Notizen) · Methodenkarte „Chorisches Lesen“.")], { before: 100, after: 40 }));

// ---------- A ----------
doc.push(avivaPhase("A", "Ankommen und einstimmen: Wer ist Paddy Fingal?", 4));
doc.push(...steps([
  ["luecken", { say: "Ab heute geht es um Balladen. Auf der Folie seht ihr unser Ziel und den Ablauf. Neue Heftseite: Leute, höret die Geschicht’ – Balladen. Darunter: Paddy Fingal – Lesetechniken. Lehrbuch Seite 136. Schaut nur auf den Titel und das erste Bild. Schreibt einen Satz: Wer ist Paddy Fingal?", board: folieBox("Folie 2 · Ziel und Ablauf", [bl("Ziel: Ich kann je nach Leseziel überfliegend, orientierend oder genau lesen und meinen ersten Eindruck begründen."), bl("Ablauf: 1 Wer ist Paddy Fingal? · 2 Wie lest ihr? · 3 Merksatz 1 · 4 Ü1 Paddy Fingal lesen (5 *) · 5 Was habe ich gelernt? · * LB-Aufgabe 1", { after: 0 })]), do: ["Zwei Vermutungen vorlesen lassen (Riese, Seemann, Raufbold …)."] }],
]));

// ---------- V ----------
doc.push(avivaPhase("V", "Vorwissen aktivieren: Wie lest ihr?", 4));
doc.push(...steps([
  ["luecken", { say: "Jetzt arbeitet ihr allein. Auf der Folie stehen drei Situationen. Schreibt ins Heft: Wie lest ihr – schnell drüber, gezielt suchen oder langsam Wort für Wort? Eine Minute. Dann vergleicht ihr zu zweit.", board: folieBox("Folie 4", [bl("A  Ausschreibung für ein Turnier: Du suchst nur deine Startzeit."), bl("B  Du blätterst in einem Buch und willst wissen, worum es in einem Kapitel geht."), bl("C  Eine neue Regel für deinen Sport, die du genau verstehen musst.", { after: 0 })]), sol: "A gezielt suchen · B schnell drüber · C langsam, Wort für Wort" }],
]));

// ---------- I ----------
doc.push(avivaPhase("I", "Informieren: Merksatz 1 Lesetechniken", 6));
doc.push(...steps([
  ["abschreiben", { say: "Jetzt schreibt ihr ab: Merksatz 1. Rot umrahmen.", board: folieBox("Folie 5 · Merksatz 1 – abschreiben, rot umrahmen", [
    p([S("Je nach "), S("Leseziel", { bold: true }), S(" lese ich unterschiedlich:")], { after: 30 }),
    p([S("überfliegend: ", { bold: true }), S("schnell über den Text (Überschrift, Anfang, Ende, auffällige Wörter) – ich will wissen, worum es geht.")], { after: 20 }),
    p([S("orientierend: ", { bold: true }), S("gezielt suchen, den Rest überspringen – ich will eine bestimmte Information finden.")], { after: 20 }),
    p([S("genau: ", { bold: true }), S("Satz für Satz, Unbekanntes klären, Randnotizen machen – ich will alles verstehen.")], { after: 40 }),
    p([S("Wofür? ", { bold: true }), S("Ich spare Zeit und finde, was ich brauche – die Startzeit in der Ausschreibung genauso wie die wichtige Stelle in einer Ballade. "), S("Test: ", { bold: true }), S("Frag dich vor dem Lesen: Will ich einen Überblick, eine Information oder alles verstehen?")], { after: 0 }),
  ]), do: ["Rückbezug: A = orientierend, B = überfliegend, C = genau."] }],
]));

// ---------- V ----------
doc.push(avivaPhase("V", "Verarbeiten: Ü1 Paddy Fingal lesen", 22));
doc.push(...steps([
  ["luecken", { say: "Schreibt ins Heft: Ü1 Paddy Fingal lesen, darunter a). Teil a: überfliegen. Ihr habt genau eine Minute für beide Seiten. Start! … Stopp. Buch zu. Schreibt einen Satz: In der Ballade geht es um …", do: ["Zeit laut stoppen – Überfliegen funktioniert nur mit echter Zeitgrenze."], sol: "z. B.: In der Ballade geht es um den Riesen Paddy Fingal und einen noch größeren Schotten, vor dem er sich versteckt." }],
  ["gemeinsam", { say: "Stifte liegen. Ich lese euch die Ballade vor. Jede Strophe endet mit Paddy Fingal. Diese letzte Zeile sprecht ihr alle zusammen mit.", do: ["Methodenkarte „Chorisches Lesen“: Lehrkraft liest die Strophen, die Klasse spricht den Kehrreim (letzte Zeile jeder Strophe) im Chor. Vorher Tempo und Lautstärke kurz festlegen.", "Danach schreiben alle ins Heft: Wie wirkt die Ballade auf dich – lustig, spannend oder unheimlich? Begründe in einem Satz. Zwei Schüler lesen vor."] }],
  ["luecken", { board: folieBox("Folie 8 · Ü1 b", [bl("1 Wie lang ist Paddys Bart? · 2 Woher kommt der Fremde? · 3 Wie heißt Paddys Frau? · 4 Was guckt aus dem Bett heraus? · 5* Was ist eine Furt?", { after: 0 })]), say: "Teil b: orientierend lesen. Schreibt b) ins Heft und nummeriert 1 bis 5. Ihr sucht nur die Antworten – nicht alles lesen! Notiert auch die Strophe. Sechs Minuten. Wer fertig ist, macht Frage 5.", sol: "1 sechs Ellen (Str. 1) · 2 aus Schottland (Str. 2) · 3 Schaya (Str. 4) · 4 nur seine Nase (Str. 5) · *5 eine flache Stelle in einem Fluss oder Bach (Worterklärung 7 – Fußnoten nutzen!)" }],
  ["luecken", { say: "Teil c: genau lesen. Schreibt c) ins Heft. Strophe 7 bis 9, Satz für Satz. Notiert zu jeder Strophe einen Stichpunkt. Dann beantwortet: Warum läuft der Schotte weg? Fünf Minuten.", sol: "Str. 7: Schaya warnt: Paddys Kind schläft. · Str. 8: Der Schotte schaut ins Bett und sieht die riesige Nase. · Str. 9: Er denkt: Wenn schon das Baby so groß ist, wie groß ist dann der Vater? · Warum? Schayas List: Sie gibt Paddy als Baby aus – der Schotte bekommt Angst vor dem „Vater“." }],
  ["partner", { say: "Vergleicht zu zweit a, b und c. Verbessert mit Grün.", do: ["Danach Frage c) im Plenum: ein Schüler erklärt die List."] }],
]));

// ---------- A ----------
doc.push(avivaPhase("A", "Auswerten: Ziel erreicht?", 5));
doc.push(...steps([
  ["luecken", { say: "Drei Situationen auf der Folie. Schreibt ins Heft: Welche Lesetechnik passt? Eine Minute.", board: folieBox("Folie 11", [bl("1  Du willst wissen, ob ein Zeitungsartikel über deinen Verein berichtet."), bl("2  Du suchst, wie lang Paddys Bart ist."), bl("3  Du sollst erklären, warum der Schotte wegläuft.", { after: 0 })]), sol: "1 überfliegend · 2 orientierend · 3 genau" }],
  ["gemeinsam", { say: "Schaut auf unser Ziel. Habt ihr es erreicht? Schreibt: Heute habe ich gelernt, …", do: ["Zwei Schüler lesen vor.", "* Wenn Zeit bleibt: LB-Aufgabe 1 b – ein Schüler erzählt die Ballade mündlich nach (Titel, Autor, Reihenfolge, nichts dazuerfinden)."] }],
]));

const pageProps = { page: { size: { width: 11906, height: 16838 }, margin: { top: 900, bottom: 800, left: 850, right: 850, header: 400, footer: 400 } } };
const styles = { default: { document: { run: { font: "Calibri", size: 22 } } } };
Packer.toBuffer(new Document({ styles, sections: [{ properties: pageProps, headers: { default: header("Balladen · Paddy Fingal") }, footers: { default: footer("Tafelskript") }, children: doc }] })).then((buf) => fs.writeFileSync("Tafelskript_Paddy_Fingal_Lesetechniken.docx", buf));
