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


// Tafelskript Fr 16.10. „Paddy Fingal – Inhalt sichern, Merkmale der Ballade (Ur-Ei)“ – AVIVA, ohne Präsentation
const avivaPhase = (k, title, min) => p([t(`${k}  `, { font: "Cambria", size: 28, bold: true, color: C.hv }), t(title, { font: "Cambria", size: 28, bold: true, color: NAVY }), t(`   ${min} Min.`, { size: 20, bold: true, color: "C77C12" })], { before: 200, after: 100, keepNext: true });
const zielTafel = (ziel, ablauf) => board("Seitentafel links · vor der Stunde anschreiben, bleibt die ganze Stunde stehen", [
  p([t("Ziel: ", { bold: true, size: 21, color: NAVY }), t(ziel, { size: 21, italics: true })], { after: 60 }),
  p(t("Ablauf", { bold: true, size: 20, color: NAVY }), { after: 20 }),
  ...ablauf.map((a, i) => p(t(`${i + 1}  ${a}`, { size: 20 }), { after: 10 })),
]);
const GC = { ly: "2A8C8C", ep: "7A5BA6", dr: "C77C12" };
const ZIEL = "Ich kann an einer Ballade lyrische, epische und dramatische Merkmale nachweisen.";
const doc = [];
doc.push(kicker("Tafelskript für die Lehrkraft · Fr 16.10. · Einzelstunde (45 Min.) · AVIVA · ohne Präsentation"), h1("Paddy Fingal – Inhalt sichern und Merkmale der Ballade"));
doc.push(box([p([S("Bezug: ", { bold: true }), S("Am Donnerstag wurde nach der Testrückgabe nur die erste Aufgabe (Ü1, Wiederholung Lesetechniken) geschafft. LB-Aufgabe 1 a ist deshalb heute die Wiederholung. Sie sichert den Inhalt, den die Klasse für die Analyse braucht. Danach wird die Ballade zum ersten Mal formal untersucht. Der Analyse-Schnipsel bündelt das „Ur-Ei“ und die Belegstellen, abgeschrieben wird nichts. Es ist die letzte Stunde vor den Herbstferien, nach den Ferien wird das Ur-Ei an „John Maynard“ wiederholt.")], { after: 0 })], LIGHT2));
doc.push(p([S("Lernziel: ", { bold: true }), S(ZIEL, { italics: true })], { before: 120, after: 40 }));
doc.push(p([S("Antwort auf „Was habt ihr heute gelernt?“: ", { bold: true }), S("„Eine Ballade ist ein Gedicht mit Strophen, Reim und Kehrreim. Sie erzählt eine Geschichte mit Figuren und Handlung, und die Figuren sprechen wie im Theater. Das habe ich an Paddy Fingal mit Strophen belegt.“")], { after: 100 }));
doc.push(h3("Ablauf"));
doc.push(grid([900, 1300, 3400, 4606], ["Zeit", "AVIVA", "Phase", "Alle Schüler …"], [
  ["3′", "A", "Ankommen: Ziel und Ablauf", "lesen Ziel und Ablauf, schreiben die Überschrift"],
  ["11′", "V", "Wiederholung: LB-Aufgabe 1 a", "lesen genau im Tandem, beantworten drei Fragen im Heft, prüfen in der Kreide-Kette mit"],
  ["8′", "I", "Die Ballade als „Ur-Ei“", "hören zu, lesen Strophe 1 im Chor, finden Verse und Reime gemeinsam"],
  ["15′", "V", "Schnipsel: Merkmale nachweisen", "tragen allein Belege ein, vergleichen zu zweit"],
  ["6′", "A", "Auswerten: Ziel erreicht?", "tragen Belege in der Kreide-Kette ins Tafel-Ei ein, alle prüfen; „Heute habe ich gelernt, …“"],
], { bold: [true, true, true, false], colors: [C.hv, C.hv, NAVY, null], h: 360 }));
doc.push(p([S("Puffer: ", { bold: true, color: C.hv }), S("2 Minuten. Mit * markiert: Nr. 8 auf dem Schnipsel (Spannung) und der Ausblick auf „John Maynard“. Was offen bleibt, ist nach den Ferien der Einstieg (Wiederholung).")], { before: 80, after: 60 }));
doc.push(p([S("Heft: ", { bold: true }), S("Überschrift „Paddy Fingal – Merkmale der Ballade“ · LB 1 a (drei Antworten in Sätzen, mit Strophe) · Schnipsel „Ur-Ei“ eingeklebt und ausgefüllt, mit Grün verbessert · „Heute habe ich gelernt, …“")], { after: 40 }));
doc.push(p([S("Material: ", { bold: true }), S("Lehrbuch S. 136 f. · Schnipsel „Die Ballade – ein Ur-Ei“ (2 pro A4, eine Hälfte pro Schüler) · Kleber · Kreide in drei Farben, wenn vorhanden (Lyrik, Epik, Dramatik).")], { after: 40 }));
doc.push(p([S("Tipp: ", { bold: true, color: C.hv }), S("Ziel und Ablauf vor der Stunde an die Seitentafel schreiben, die Fragen zu 1 a an die Mitteltafel. Das Ei erst in Phase I zeichnen.")], { after: 60 }));
doc.push(zielTafel(ZIEL, ["Wiederholung: Paddy Fingal, LB-Aufgabe 1 a", "Die Ballade als „Ur-Ei“", "Schnipsel: Merkmale nachweisen", "Ziel erreicht? Heute habe ich gelernt, …", "* Ausblick: John Maynard"]));

doc.push(avivaPhase("A", "Ankommen: Ziel und Ablauf", 3));
doc.push(...steps([["zuhoeren", { say: "Stifte liegen. Ihr hört nur zu. Links steht unser Ziel und der Ablauf. Zuerst wiederholen wir, worum es in Paddy Fingal geht. Danach schauen wir, wie die Ballade gebaut ist. Jetzt schreibt ihr die Überschrift ab: Paddy Fingal – Merkmale der Ballade." }]]));

doc.push(avivaPhase("V", "Vorwissen aktivieren: LB-Aufgabe 1 a als Wiederholung", 11));
doc.push(board("Mitteltafel · LB-Aufgabe 1 a", [
  bl([N("1 a  ", { bold: true }), N("Lest genau. Beantwortet im Heft in ganzen Sätzen, mit Strophe:")]),
  bl("    1  Wie sieht Paddy Fingal aus?"),
  bl("    2  Was möchte der Riese aus Schottland von Paddy?"),
  bl("    3  Welche List wendet Paddys Frau an?"),
]));
doc.push(p(t(""), { after: 60 }));
doc.push(...steps([
  ["partner", { say: "Jetzt arbeitet ihr im Tandem. Lest die Ballade genau, Strophe für Strophe abwechselnd. Dann schreibt jeder die drei Antworten in sein Heft. Sieben Minuten. Fertig? Stift hinlegen.", sol: "1 Ein Riese mit einem schwarzen Bart, sechs Ellen lang, Fäuste so groß wie Tonnen; er wirft mit einem Mastbaum (Str. 1). · 2 Der Schotte will herausfinden, wer von beiden stärker ist (Str. 2). · 3 Schaya steckt Paddy ins Bett, türmt die Kissen hoch und gibt ihn als ihr Baby aus. Der Schotte sieht nur die riesige Nase, denkt an einen riesigen Vater und flieht (Str. 5–10).", do: ["Methodenkarte „Tandemlesen“ (Tandems wie festgelegt)."] }],
  ["tafel", { say: "Wer drankommt, schreibt die Antwort in Stichworten mit Strophe an die Tafel. Alle prüfen mit und verbessern mit Grün.", do: ["Kreide-Kette: drei Schüler, je eine Frage. Bei Frage 3 nachfragen: „Welche Lesetechnik hast du gebraucht?“ (genau)."] }],
]));

doc.push(avivaPhase("I", "Informieren: Die Ballade als „Ur-Ei“", 8));
doc.push(board("Mitteltafel · Tafelbild (Ei in drei Teile geteilt, bleibt bis zur Auswertung stehen)", [
  p([N("Die Ballade – ein „Ur-Ei“ (Goethe)", { bold: true, color: NAVY })], { after: 60, align: AlignmentType.CENTER }),
  bl([N("Lyrik", { bold: true, color: GC.ly }), N(" – wie ein Gedicht: Strophen, Verse, Reim, Kehrreim")]),
  bl([N("Epik", { bold: true, color: GC.ep }), N(" – wie eine Erzählung: Erzähler, Figuren, Handlung (Anfang – Höhepunkt – Ende)")]),
  bl([N("Dramatik", { bold: true, color: GC.dr }), N(" – wie ein Theaterstück: wörtliche Rede, Spannung")]),
  bl([N("Belege: ", { bold: true }), N("… (füllt sich in der Auswertung)")], { before: 60 }),
]));
doc.push(p(t(""), { after: 60 }));
doc.push(...steps([
  ["zuhoeren", { say: "Stifte liegen. Ihr hört nur zu. Goethe hat die Ballade ein Ur-Ei genannt. In einem Ei steckt schon alles, was später daraus wird. In der Ballade stecken alle drei großen Gattungen: Lyrik, also Gedichte, Epik, also Erzählungen, und Dramatik, also Theater. Heute weisen wir das an Paddy Fingal nach.", do: ["Beim Sprechen das Ei zeichnen und die drei Teile beschriften. Relevanz nennen: Die Merkmale braucht ihr für die erste Bewertung (Analyse)."] }],
  ["gemeinsam", { say: "Lehrbuch S. 136. Wir lesen Strophe 1 zusammen im Chor. … Zählt die Verse. Welche Wörter reimen sich?", sol: "5 Verse · Riesenart – Bart (a a), groß – Wurfgeschoss (b b) = Paarreim · 5. Vers eingerückt: „so ein Kerl war Paddy Fingal!“ = Kehrreim", do: ["Methodenkarte „Chorisches Lesen“.", "Reimwörter an der Tafel mit a a b b beschriften. Nachfragen: „Kommt der letzte Vers noch einmal vor?“ – Ja, jede Strophe endet mit „Paddy Fingal“."] }],
  ["zuhoeren", { say: "Ich teile jetzt den Schnipsel aus. Klebt ihn unter die Antworten von 1 a. Nummer 1 tragen wir gemeinsam ein.", sol: "1  11 Strophen mit je 5 Versen" }],
]));

doc.push(avivaPhase("V", "Verarbeiten: Schnipsel – Merkmale nachweisen", 15));
doc.push(...steps([
  ["luecken", { say: "Jetzt arbeitet ihr allein. Füllt Nummer 2 bis 7 aus und notiert immer die Strophe. Zehn Minuten. Wer fertig ist, macht Nummer 8 mit Stern. Fertig? Stift hinlegen.", sol: "2 Paarreim (aabb), z. B. Riesenart / Bart (Str. 1) · 3 „… Paddy Fingal“ (5. Vers, wird abgewandelt) · 4 Paddy Fingal, Schaya, der Schotte (*die Nachbarn) · 5 Anfang Str. 1–3, Höhepunkt Str. 8–9 (der Schotte sieht die Nase), Ende Str. 10–11 (Flucht, Siegesgeschrei); andere begründete Grenzen gelten · 6 Präteritum, z. B. „Paddy Fingal war von Riesenart“ (Str. 1), „Ins Bette Paddy Fingal kroch“ (Str. 5) · 7 Paddy (Str. 4), der Schotte (Str. 6, 9), Schaya (Str. 7) · *8 Paddy gegen den Schotten; Frage: Findet der Schotte Paddy? / Wer ist stärker?", do: ["Herumgehen. Typische Hürde bei Nr. 5: Höhepunkt = die Stelle, an der sich alles entscheidet (der Schotte steht vor dem Bett).", "Bei Nr. 6 Bezug zur Verb-Einheit: Woran erkennst du das Präteritum? (war, kroch, lief)"] }],
  ["kontrolle", { say: "Vergleicht mit eurem Nachbarn und verbessert mit Grün. Drei Minuten." }],
]));

doc.push(avivaPhase("A", "Auswerten: Ziel erreicht?", 6));
doc.push(...steps([
  ["tafel", { say: "Wer drankommt, schreibt einen Beleg mit Strophe in den passenden Teil des Eis. Alle prüfen mit und verbessern mit Grün.", do: ["Kreide-Kette: je ein Schüler für Lyrik, Epik und Dramatik (z. B. Paarreim Str. 1 · Präteritum „kroch“ Str. 5 · wörtliche Rede Schaya Str. 7)."] }],
  ["luecken", { say: "Schaut auf unser Ziel links. Schreibt ins Heft: Heute habe ich gelernt, …", do: ["Zwei Schüler lesen vor. Auf den Ablauf zurückgreifen: Ziel erreicht? Was mit * offen ist, kommt nach den Ferien als Wiederholung.", "* Ausblick: „Nach den Ferien hört ihr John Maynard. Dann prüft ihr: Ist das auch ein Ur-Ei?“"] }],
]));
const pageProps = { page: { size: { width: 11906, height: 16838 }, margin: { top: 900, bottom: 800, left: 850, right: 850, header: 400, footer: 400 } } };
const styles = { default: { document: { run: { font: "Calibri", size: 22 } } } };
Packer.toBuffer(new Document({ styles, sections: [{ properties: pageProps, headers: { default: header("Balladen · Fr 16.10.") }, footers: { default: footer("Tafelskript") }, children: doc }] })).then((buf) => fs.writeFileSync("Tafelskript_Fr_16-10_Paddy_Fingal_Ur-Ei.docx", buf));
