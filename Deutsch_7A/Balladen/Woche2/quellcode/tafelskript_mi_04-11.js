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



// Tafelskript Mi 4.11. (Doppelstunde): Teil 1 Paddy Fingal – Inhalt · Teil 2 Nebensätze mit Subjunktion – AVIVA, mit Präsentation
const avivaPhase = (k, title, min) => p([t(`${k}  `, { font: "Cambria", size: 28, bold: true, color: C.hv }), t(title, { font: "Cambria", size: 28, bold: true, color: NAVY }), t(`   ${min} Min.`, { size: 20, bold: true, color: "C77C12" })], { before: 200, after: 100, keepNext: true });
const teil = (nr, title, col) => p([t(`TEIL ${nr}  `, { size: 20, bold: true, color: col }), t(title, { font: "Cambria", size: 30, bold: true, color: col })], { before: 280, after: 80, keepNext: true, border: { bottom: { style: BorderStyle.SINGLE, size: 8, color: col, space: 4 } } });
const VIO = "7A5BA6", GRN = "2E9E6B";
const Z1 = "Ich kann Fragen zum Inhalt der Ballade mit Textbeleg beantworten.";
const Z2 = "Ich kann Nebensätze mit Subjunktion erkennen und richtig bilden: Das finite Verb steht am Ende.";
const doc = [];
doc.push(kicker("Tafelskript für die Lehrkraft · Mi 4.11. · Doppelstunde (90 Min.) · AVIVA · mit Präsentation (16 Folien)"), h1("Paddy Fingal: Inhalt und Sprachtraining an der Ballade"));
doc.push(box([p([S("Bezug: ", { bold: true }), S("Die Stunde am 16.10. ist ausgefallen. Am Montag wurden in 15 Minuten die Lesetechniken wiederholt. Offen sind LB-Aufgabe 1 a zu Paddy Fingal, die Reime und die Merkmale der Ballade. Teil 1 holt den Inhalt nach und beginnt mit der Wiederholung vom Montag. Teil 2 integriert Grammatik nach dem neuen Prinzip, gegenüber der Klasse heißt sie „Sprachtraining“: Die Klasse entdeckt ein Phänomen in der Ballade (Nebensätze mit „dass“ und „als“), leitet daraus die Regel ab, übt sie an Sätzen aus dem Sportalltag und wendet sie wieder auf die Ballade an. Das knüpft an die Verb-Einheit an (finite Form, Hilfsverb). Reime und Ur-Ei folgen am Do/Fr (Material in Woche1/).")], { after: 0 })], LIGHT2));
doc.push(p([S("Lernziele: ", { bold: true }), S("Teil 1: " + Z1 + "  Teil 2: " + Z2, { italics: true })], { before: 120, after: 40 }));
doc.push(p([S("Antwort auf „Was habt ihr heute gelernt?“: ", { bold: true }), S("„Ich kann mit Strophen belegen, wie Paddy aussieht, was der Schotte will und welche List Schaya anwendet. Nach dass, als, weil, wenn steht das finite Verb am Ende, und davor kommt ein Komma.“ Überprüft wird das im Zielcheck am Ende.")], { after: 80 }));
doc.push(board("Seitentafel links · vor der Stunde anschreiben, bleibt stehen", [
  p([t("Teil 1 (ca. 40 Min.): ", { bold: true, size: 20, color: GRN }), t(Z1, { size: 20, italics: true })], { after: 30 }),
  p(t("1 Wiederholung Lesetechniken  2 Ü1 Tandemlesen, LB 1 a  3 * LB 1 b nacherzählen", { size: 19 }), { after: 60 }),
  p([t("Teil 2 (ca. 45 Min.): ", { bold: true, size: 20, color: VIO }), t(Z2, { size: 20, italics: true })], { after: 30 }),
  p(t("4 Wo steht das finite Verb?  5 Ü2 Entdecken, Merksatz  6 Ü3, Ü4  7 * Ü5  8 Zielcheck", { size: 19 }), { after: 0 }),
]));
doc.push(p(t(""), { after: 40 }));
doc.push(h3("Ablauf"));
doc.push(grid([800, 900, 1200, 3300, 4006], ["Zeit", "AVIVA", "Folie", "Phase", "Alle Schüler …"], [
  ["4′", "A", "1–2", "Ankommen: Überblick, zwei Ziele", "lesen Ziele und Ablauf"],
  ["6′", "V", "4", "T1 · Wiederholung Lesetechniken", "ordnen A–C zu, begründen mit „weil“"],
  ["2′", "I", "5", "T1 · Auftrag Ü1, Tandemlesen", "hören zu"],
  ["16′", "V", "5", "T1 · Tandemlesen, LB 1 a", "lesen im Tandem, beantworten drei Fragen im Heft"],
  ["5′", "A", "6", "T1 · Prüf-Kette", "an der Tafel: prüfen (✓/✗ + Begründung) und anschreiben; alle anderen korrigieren ihr Heft mit Grün"],
  ["* 8′", "V", "7", "T1 · * LB 1 b nacherzählen", "erzählen sich im Tandem nach, haken die Checkliste ab"],
  ["4′", "V", "9", "T2 · Wo steht das finite Verb?", "schreiben zwei Sätze ab, unterstreichen"],
  ["12′", "I", "10–11", "T2 · Ü2 entdecken, Merksatz", "untersuchen drei Stellen, kleben den Merksatz ein"],
  ["19′", "V", "12–13", "T2 · Ü3 Sätze verbinden, Ü4 Fehler im Chat", "verbinden allein, Prüf-Kette mit Heftkorrektur, korrigieren zu zweit"],
  ["* 5′", "V", "14", "T2 · * Ü5 zurück zur Ballade", "schreiben Sätze mit Satzanfängen"],
  ["6′", "A", "15–16", "Auswerten: Zielcheck", "* Forscherfrage; lösen drei Zielcheck-Aufgaben, kontrollieren mit der Lösungsfolie"],
], { bold: [true, true, false, true, false], colors: [C.hv, C.hv, MUTED, NAVY, null], h: 320 }));
doc.push(p([S("Zeit: ", { bold: true, color: C.hv }), S("74 Min. Pflicht, 13 Min. mit * (1 b, Ü5, Forscherfrage), 3 Min. Puffer. Wenn 1 b entfällt, ist es der Einstieg am Donnerstag.")], { before: 80, after: 60 }));
doc.push(h3("So sieht das Heft am Ende aus"));
doc.push(table([2300, 7906], [
  ["Teil 1", "Überschrift „Paddy Fingal – den Inhalt erschließen“ · A–C mit Lesetechnik und „weil“ · Ü1: drei Antworten in Sätzen, mit Strophe, mit Grün verbessert"],
  ["Teil 2", "Überschrift „Nebensätze mit Subjunktion“ · zwei Sätze (finites Verb unterstrichen) · Ü2: drei Stellen aus der Ballade · Merksatz-Schnipsel eingeklebt · Ü3: sechs Sätze · Ü4: vier verbesserte Chat-Sätze · * Ü5"],
  ["Abschluss", "Zielcheck: Antwort zu Ziel 1, Satz zu Ziel 2 a, Begründung zu Ziel 2 b – mit Grün kontrolliert"],
].map(([a, b]) => row([tc(a, 2300, { bold: true, fill: LIGHT }), tc(b, 7906)]))));
doc.push(p([S("Material: ", { bold: true }), S("Präsentation `Mittwoch_04-11_Deutsch_7A.pptx` (Lösungen in den Notizen) · Lehrbuch S. 136 f. · Merksatz-Schnipsel „Nebensätze mit Subjunktion“ (8 pro A4) · Tandems wie festgelegt · Handreichung Polnisch für S-PL.")], { before: 80, after: 60 }));

doc.push(avivaPhase("A", "Ankommen: Überblick und Ziele", 4));
doc.push(...steps([["zuhoeren", { say: "Stifte liegen. Ihr hört nur zu. Heute haben wir zwei Teile. In Teil 1 klären wir, worum es in Paddy Fingal genau geht. In Teil 2 ist Sprachtraining: Wir schauen uns die Sprache der Ballade an und lernen daraus eine Regel, mit der ihr besser schreibt. Die Ziele stehen auf der Folie und links an der Tafel.", do: ["Folie 2 zeigen, auf die Seitentafel verweisen. Folie 2 am Ende noch einmal zeigen."] }]]));

doc.push(teil("1", "Paddy Fingal – den Inhalt erschließen", GRN));
doc.push(avivaPhase("V", "Vorwissen aktivieren: Wiederholung Lesetechniken", 6));
doc.push(...steps([
  ["luecken", { say: "Heft auf, Überschrift: Paddy Fingal – den Inhalt erschließen. Jetzt arbeitet ihr allein. Schreibt A, B, C, die Lesetechnik und eine Begründung mit weil. Drei Minuten. Fertig? Stift hinlegen.", sol: "A orientierend, weil ich nur eine Information suche · B überfliegend, weil ich nur wissen will, wie es ausgeht · C genau, weil ich etwas erklären soll" }],
  ["kontrolle", { say: "Vergleicht mit eurem Nachbarn und verbessert mit Grün.", do: ["Falls E (Überfliegen) am Montag nicht geschafft wurde: jetzt 1 Minute überfliegen, Satz „In der Ballade geht es um …“.", "Hinweis für später: „Merkt euch eure weil-Sätze – in Teil 2 schauen wir sie genauer an.“"] }],
]));
doc.push(avivaPhase("I", "Informieren: Auftrag Ü1 und Tandemlesen", 2));
doc.push(...steps([["zuhoeren", { say: "Stifte liegen. Ihr lest jetzt im Tandem: Einer liest eine Strophe vor, der andere liest leise mit und hilft bei Fehlern, dann Wechsel. Danach schreibt jeder die drei Antworten in sein Heft, mit Strophe.", do: ["Methodenkarte „Tandemlesen“ (Tandems lesestärker/leseschwächer)."] }]]));
doc.push(avivaPhase("V", "Verarbeiten: Ü1 – LB-Aufgabe 1 a", 16));
doc.push(...steps([["partner", { say: "Jetzt arbeitet ihr im Tandem. Die Aufgabe steht auf der Folie. Acht Minuten lesen, dann acht Minuten schreiben. Fertig? Stift hinlegen.", sol: "1 Ein Riese mit einem schwarzen Bart, sechs Ellen lang, Fäuste so groß wie Tonnen; er wirft mit einem Mastbaum (Str. 1). · 2 Der Schotte will herausfinden, wer von beiden stärker ist (Str. 2). · 3 Schaya steckt Paddy ins Bett, türmt die Kissen hoch und gibt ihn als ihr Baby aus. Der Schotte sieht nur die riesige Nase, denkt an einen riesigen Vater und flieht (Str. 5–10).", do: ["Herumgehen, beim Tandemlesen zuhören. Wer fertig ist, beginnt mit * 1 b (Folie 7)."] }]]));
doc.push(avivaPhase("A", "Auswerten Teil 1: Prüf-Kette", 5));
doc.push(...steps([["tafel", { say: "Prüf-Kette. Wer drankommt, prüft zuerst die Antwort davor: Haken oder Kreuz und ein Stichwort zur Begründung. Dann schreibt er die nächste Antwort mit Strophe an. Alle anderen vergleichen mit ihrem Heft und verbessern mit Grün.", do: ["S1 schreibt Antwort 1. S2 prüft 1 und schreibt 2. S3 prüft 2 und schreibt 3. S4 prüft 3. Bei Frage 3 nachfragen: Welche Lesetechnik hast du gebraucht? (genau)."] }]]));
doc.push(avivaPhase("V", "* Verarbeiten: LB-Aufgabe 1 b nacherzählen", 8));
doc.push(...steps([["partner", { say: "A erzählt Strophe 1 bis 5 nach, B hört zu und hakt die Checkliste ab. Dann Wechsel: B erzählt Strophe 6 bis 11. Eure Antworten aus Ü1 sind euer Spickzettel.", do: ["Checkliste: Titel und Autor · richtige Reihenfolge · nichts dazuerfinden · alle wichtigen Details · * spannend mit wörtlicher Rede.", "Fällt 1 b weg: Einstieg am Donnerstag."] }]]));

doc.push(teil("2", "Sprachtraining: Nebensätze mit Subjunktion", VIO));
doc.push(avivaPhase("V", "Vorwissen aktivieren: Wo steht das finite Verb?", 4));
doc.push(...steps([["luecken", { say: "Neue Überschrift: Nebensätze mit Subjunktion. Jetzt arbeitet ihr allein. Schreibt die beiden Sätze von der Folie ab und unterstreicht das finite Verb. Wo steht es?", sol: "Paddy Fingal war von Riesenart. · Da lief er heim in jähem Schreck. → war, lief: an Position 2 (nach dem ersten Satzglied).", do: ["Kurz fragen: Woran erkennst du das finite Verb? – Es ändert sich mit der Person (ich war, du warst)."] }]]));
doc.push(avivaPhase("I", "Informieren: Ü2 entdecken und Merksatz", 12));
doc.push(board("Mitteltafel · Tafelbild (entsteht im Gespräch, bleibt stehen)", [
  bl([N("Hauptsatz:  ", { bold: true }), N("Da "), N("lief", { bold: true, underline: {} }), N(" er heim in jähem Schreck.        → finites Verb an Position 2", { color: MUTED })]),
  bl([N("Nebensatz:  ", { bold: true }), N("…, "), N("dass", { bold: true, color: VIO }), N(" der Fremdling zehn Schuh größer "), N("war", { bold: true, underline: {} }), N(".   → finites Verb am Ende", { color: MUTED })]),
  bl([N("                    …, "), N("als", { bold: true, color: VIO }), N(" er die Nase ragen "), N("sah", { bold: true, underline: {} }), N(".")]),
  bl([N("dass, als = Subjunktion · Komma davor", { bold: true, color: VIO })], { before: 40 }),
]));
doc.push(p(t(""), { after: 40 }));
doc.push(...steps([
  ["luecken", { say: "Jetzt arbeitet ihr allein. Auf der Folie stehen drei Stellen aus der Ballade. Schreibt die fett gedruckten Teile ins Heft. Unterstreicht das finite Verb und kreist das erste Wort ein. Vier Minuten.", sol: "dass der Fremdling zehn Schuh größer war · dass Paddys Kindlein nicht erwacht · als er die Nase ragen sah" }],
  ["partner", { say: "Vergleicht zu zweit: Was fällt euch auf, wenn ihr diese Sätze mit den Sätzen von eben vergleicht? Zwei Minuten.", do: ["Erwartet: Das finite Verb steht am Ende; vorne steht dass oder als; davor ein Komma."] }],
  ["gemeinsam", { say: "Wer drankommt, sagt, was er entdeckt hat. Alle prüfen mit.", do: ["Tafelbild entwickeln (siehe oben): Hauptsatz – Nebensatz. Fachbegriffe einführen: Subjunktion, Nebensatz."] }],
  ["zuhoeren", { say: "Stifte liegen. Ihr hört nur zu. Was ihr entdeckt habt, ist eine feste Regel. Ich teile den Merksatz aus. Klebt ihn ein. Wir lesen ihn zusammen.", do: ["Folie 11. Auf das Perfekt-Beispiel hinweisen: Das Hilfsverb ist finit und wandert ans Ende (Bezug zur Verb-Einheit). Bezug Englisch: because he is afraid – dort bleibt das Verb vorne."] }],
]));
doc.push(avivaPhase("V", "Verarbeiten: Ü3 Sätze verbinden, Ü4 Fehler im Chat", 19));
doc.push(...steps([
  ["luecken", { say: "Jetzt arbeitet ihr allein. Ü3 steht auf der Folie. Verbindet die Sätze mit der Subjunktion in Klammern und schreibt den ganzen Satz ins Heft. Kreist die Subjunktion ein und unterstreicht das finite Verb. Acht Minuten. Nummer 6 hat einen Stern. Fertig? Stift hinlegen.", sol: "1 Ich trainiere heute länger, weil am Samstag ein Wettkampf ist. · 2 Der Trainer sagt, dass wir nach Frankfurt (Oder) fahren. · 3 Ich war nervös, als ich auf der Matte stand. · 4 Wir gewinnen, wenn wir als Team spielen. · 5 Ich bin stolz, weil ich 80 kg gestoßen habe. · *6 Der Trainer fragt, ob ich gut geschlafen habe." }],
  ["tafel", { say: "Prüf-Kette. Wer drankommt, prüft zuerst den Satz davor: Haken oder Kreuz, Begründung – Verb am Ende? Komma? Dann schreibt er den nächsten Satz an und markiert Subjunktion und finites Verb. Alle anderen vergleichen mit ihrem Heft und verbessern mit Grün.", do: ["Sechs Schüler: S1 schreibt Satz 1, S2 prüft 1 und schreibt 2 … S6 prüft 5. Bei Satz 5 nachfragen: Welches Verb ist hier finit? (habe – das Hilfsverb, rot)."] }],
  ["partner", { say: "Ü4: Im Chat stecken Fehler. Findet sie zu zweit und schreibt die Sätze richtig ins Heft. Notiert, ob es um die Verbstellung oder um das Komma geht. Fünf Minuten.", sol: "1 …, weil ich noch duschen muss. · 2 …, dass wir morgen frei haben. · 3 Weißt du, ob der Neue auch kommt? · 4 Ich freue mich, weil wir frei haben. (Komma fehlt)", do: ["Kurz besprechen: Beim Sprechen hört man oft „weil ich muss …“. Im Geschriebenen ist das falsch."] }],
]));
doc.push(avivaPhase("V", "* Verarbeiten: Ü5 zurück zur Ballade", 5));
doc.push(...steps([["luecken", { say: "Wer fertig ist: Ü5. Schreibt mindestens drei Sätze über die Ballade mit den Satzanfängen auf der Folie.", sol: "z. B. … weil er Angst vor dem Schotten hat · … dass Paddy ihr Baby ist · … als er die riesige Nase sieht · … weil er glaubt, dass der Vater noch größer ist" }]]));
doc.push(avivaPhase("A", "Auswerten: Zielcheck", 6));
doc.push(...steps([
  ["gemeinsam", { say: "* Forscherfrage: In Strophe 3 steht: Und als er stapfte durch den Sund. Was ist hier anders – und warum macht der Dichter das?", sol: "Das finite Verb steht nicht am Ende. Der Dichter weicht wegen Reim und Rhythmus ab (Sund – stund). Im Gedicht ist das erlaubt, im Aufsatz gilt die Regel." }],
  ["luecken", { say: "Zielcheck. Jetzt arbeitet ihr allein. Löst die drei Aufgaben auf der Folie im Heft. Sie zeigen, ob ihr unsere Ziele erreicht habt. Vier Minuten. Fertig? Stift hinlegen.", sol: "Ziel 1: Er sieht die riesige Nase des „Babys“ und glaubt, der Vater müsse noch viel größer sein (Str. 8–10). · Ziel 2 a: Der Schotte erschrickt, als er die riesige Nase sieht. · Ziel 2 b: Falsch – im Nebensatz steht das finite Verb am Ende: …, weil er Angst hat." }],
  ["kontrolle", { say: "Vergleicht mit der Lösung auf der Folie und verbessert mit Grün.", do: ["Folie 16. Herumgehen und schauen, wer bei 2 a oder 2 b noch Fehler hat – diese Fälle am Donnerstag in der Wiederholung aufgreifen. Folie 2 noch einmal zeigen: Was mit * offen ist, kommt am Donnerstag."] }],
]));

// ---------- Reflexion ----------
const bul = (head, txt) => p([S("• ", { bold: true, color: C.hv }), ...(head ? [S(head + " ", { bold: true })] : []), S(txt)], { after: 40, indent: { left: 240, hanging: 240 } });
const refHead = (txt, col) => p(t(txt, { font: "Cambria", size: 22, bold: true, color: col }), { before: 140, after: 50, keepNext: true });
doc.push(p(new PageBreak()));
doc.push(kicker("Für die Lehrkraft · Reflexion vor und nach der Stunde"), h3("Umsetzung der Erkenntnisse zur Unterrichtsverbesserung"));
doc.push(refHead("1  AVIVA – Struktur und Transparenz", C.hv));
doc.push(bul("Ankommen:", "Überblicksfolie mit zwei Teilen, Zeiten und je einem „Ich kann …“-Ziel (Rückmeldung 14.10.: „hilft sehr“), zusätzlich an der Seitentafel. Am Ende Rückgriff auf beide Ziele."));
doc.push(bul("Vorwissen aktivieren:", "In beiden Teilen echte Wiederholung: Lesetechniken vom Montag (Teil 1), finite Form aus der Verb-Einheit (Teil 2)."));
doc.push(bul("Informieren:", "In Teil 2 induktiv: erst entdecken, dann der Merksatz als Schnipsel, nicht abschreiben."));
doc.push(bul("Verarbeiten und Auswerten:", "Immer erst allein oder zu zweit. Statt der klassischen Kreide-Kette eine Prüf-Kette: An der Tafel wird geprüft (✓/✗ + Begründung) und angeschrieben, alle anderen korrigieren gleichzeitig ihr Heft. Sicherung als Zielcheck mit konkreten Aufgaben statt „Heute habe ich gelernt“ (Rückmeldung 31.10.)."));
doc.push(bul("Zeit:", "74 Min. Pflicht, alles Weitere mit *, kleiner Puffer – nach den bisherigen Erfahrungen bewusst knapp gehalten."));
doc.push(refHead("2  Direkte Instruktion (DI)", NAVY));
doc.push(bul("Ziel und Relevanz:", "Zwei überprüfbare Ziele; „Wofür?“ im Merksatz: Begründen in jeder Analyse, Bezug zum Englischen."));
doc.push(bul("Ich – Wir – Ihr:", "Lehrkraft entwickelt mit der Klasse das Tafelbild Hauptsatz/Nebensatz und modelliert am Perfekt-Beispiel (Wir/Ich), danach Ü3–Ü5 selbstständig (Ihr)."));
doc.push(bul("Rückmeldung:", "Prüf-Kette nach Ü1 und Ü3, Partnerkorrektur bei Ü4, Zielcheck mit Lösungsfolie; typische Hürden vorab benannt: Hilfsverb ans Ende (Ü3 Nr. 5), gesprochenes „weil ich muss …“ (Ü4)."));
doc.push(refHead("3  Grammatik integriert („Sprachtraining“) – aus der Ballade heraus", VIO));
doc.push(bul("", "Phänomen in der Ballade entdecken (dass, als) → Regel ableiten → an Sätzen aus dem Sportalltag üben → Fehleranalyse im Chat → zurück zur Ballade (Ü5)."));
doc.push(bul("", "Sprachreflexion: Forscherfrage zur dichterischen Abweichung (Str. 3) – die Klasse erfährt, warum ein Gedicht anders klingen darf."));
doc.push(bul("", "Anschluss an die Verb-Einheit (finite Form, Hilfsverb rot) und an Montag (weil-Sätze)."));
doc.push(refHead("4  Differenzierung", "2A8C8C"));
doc.push(bul("", "Tandems lesestärker/leseschwächer; * 1 b, * Ü3 Nr. 6 (ob mit Personwechsel), * Ü5 und Forscherfrage für Schnelle; Satzanfänge in Ü5."));
doc.push(bul("", "S-PL: Handreichung mit Zielen, Aufträgen und Fachbegriffen auf Polnisch. Hinweis zur Verbstellung, weil sie im Polnischen frei ist."));
doc.push(refHead("5  Beobachtungsschwerpunkte", MUTED));
doc.push(bul("", "Reicht die Zeit für Teil 1 ohne 1 b? Wo hakt das Tandemlesen?"));
doc.push(bul("", "Entdecken die Schüler die Verbstellung selbst, oder braucht es mehr Lenkung?"));
doc.push(bul("", "Gelingt Ü3 Nr. 5 (Hilfsverb ans Ende)? Bleibt das gesprochene „weil ich muss …“ hartnäckig?"));
doc.push(bul("", "Welches Grammatikthema bietet sich als Nächstes an (z. B. Präteritum der starken Verben in der Ballade, indirekte Rede)?"));

const pageProps = { page: { size: { width: 11906, height: 16838 }, margin: { top: 900, bottom: 800, left: 850, right: 850, header: 400, footer: 400 } } };
const styles = { default: { document: { run: { font: "Calibri", size: 22 } } } };
Packer.toBuffer(new Document({ styles, sections: [{ properties: pageProps, headers: { default: header("Balladen · Mi 4.11.") }, footers: { default: footer("Tafelskript") }, children: doc }] })).then((buf) => fs.writeFileSync("Tafelskript_Mi_04-11_Paddy_Fingal_Subjunktionen.docx", buf));
