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


// Tafelskript Do 15.10. „Testrückgabe · Paddy Fingal weiterlesen, LB-Aufgabe 1“ – AVIVA, ohne Präsentation
const avivaPhase = (k, title, min) => p([t(`${k}  `, { font: "Cambria", size: 28, bold: true, color: C.hv }), t(title, { font: "Cambria", size: 28, bold: true, color: NAVY }), t(`   ${min} Min.`, { size: 20, bold: true, color: "C77C12" })], { before: 200, after: 100, keepNext: true });
const zielTafel = (ziel, ablauf) => board("Seitentafel links · vor der Stunde anschreiben, bleibt die ganze Stunde stehen", [
  p([t("Ziel: ", { bold: true, size: 21, color: NAVY }), t(ziel, { size: 21, italics: true })], { after: 60 }),
  p(t("Ablauf", { bold: true, size: 20, color: NAVY }), { after: 20 }),
  ...ablauf.map((a, i) => p(t(`${i + 1}  ${a}`, { size: 20 }), { after: 10 })),
]);
const doc = [];
doc.push(kicker("Tafelskript für die Lehrkraft · Do 15.10. · Einzelstunde (45 Min.) · AVIVA · ohne Präsentation"), h1("Testrückgabe · Paddy Fingal – Inhalt erschließen"));
doc.push(box([p([S("Bezug: ", { bold: true }), S("Am Mittwoch kam die Klasse in Teil 3 nur bis zum Lesetraining (Vorlesen, Kehrreim im Chor); Ü1 wurde nicht geschafft. Weil Wiederholung extrem wichtig ist, beginnt der Ballade-Teil mit den Lesetechniken und Ü1 als Wiederholung, danach wird weitergelesen und LB-Aufgabe 1 gelöst. Vorher wird die Leistungskontrolle zurückgegeben (ca. 10 Min.). Im Raum gibt es keine Präsentation: Alles steht an der Tafel. Merksätze werden nicht abgeschrieben.")], { after: 0 })], LIGHT2));
doc.push(p([S("Lernziel: ", { bold: true }), S("Ich kann mit den Lesetechniken Informationen in der Ballade finden und Fragen zum Inhalt mit Textbeleg beantworten.", { italics: true })], { before: 120, after: 40 }));
doc.push(p([S("Antwort auf „Was habt ihr heute gelernt?“: ", { bold: true }), S("„Wenn ich etwas Bestimmtes suche, lese ich orientierend; wenn ich etwas erklären soll, lese ich genau. Meine Antworten belege ich mit der Strophe.“")], { after: 100 }));
doc.push(h3("Ablauf"));
doc.push(grid([900, 1300, 3400, 4606], ["Zeit", "AVIVA", "Phase", "Alle Schüler …"], [
  ["10′", "A", "Ankommen: Ziel, Ablauf, Testrückgabe", "lesen Ziel und Ablauf; prüfen ihren Test, verbessern einen Fehler mit Grün"],
  ["9′", "V", "Wiederholung: Lesetechniken und Ü1", "ordnen drei Situationen zu, überfliegen (1 Min.), suchen orientierend Antworten"],
  ["2′", "I", "Auftrag LB-Aufgabe 1", "hören zu, schreiben die Überschrift"],
  ["15′", "V", "Weiterlesen und LB-Aufgabe 1 a", "lesen im Tandem, beantworten drei Fragen im Heft"],
  ["5′", "A", "Auswerten: Ziel erreicht?", "vergleichen in der Kreide-Kette; „Heute habe ich gelernt, …“"],
], { bold: [true, true, true, false], colors: [C.hv, C.hv, NAVY, null], h: 360 }));
doc.push(p([S("Puffer: ", { bold: true, color: C.hv }), S("4 Minuten. Mit * markiert: Frage 5 in Ü1 b und LB-Aufgabe 1 b (Nacherzählen im Tandem). Was nicht geschafft wird, ist der Einstieg (Wiederholung) am Freitag.")], { before: 80, after: 60 }));
doc.push(p([S("Heft: ", { bold: true }), S("Überschrift „Paddy Fingal – Inhalt erschließen“ · A, B, C mit Lesetechnik · Ü1 a (ein Satz) · Ü1 b (1–5 mit Strophe) · LB 1 a (drei Antworten in Sätzen, mit Strophe) · „Heute habe ich gelernt, …“")], { after: 40 }));
doc.push(p([S("Material: ", { bold: true }), S("korrigierte Leistungskontrollen · Notenschlüssel (für die Tafel) · Lehrbuch S. 136 f. · Tandems wie festgelegt · Schnipsel „Merksatz Lesetechniken“ nur bei Bedarf.")], { after: 40 }));
doc.push(p([S("Tipp: ", { bold: true, color: C.hv }), S("Ziel und Ablauf vor der Stunde an die Seitentafel schreiben, die Situationen A–C, Ü1 und die Fragen zu 1 a an die Mitteltafel (Ü1 und 1 a ggf. zunächst zugeklappt).")], { after: 60 }));
doc.push(zielTafel("Ich kann mit den Lesetechniken Informationen in der Ballade finden und Fragen zum Inhalt mit Textbeleg beantworten.", [
  "Leistungskontrolle zurück", "Wiederholung: Lesetechniken, Ü1", "Paddy Fingal weiterlesen: LB-Aufgabe 1 a", "* LB-Aufgabe 1 b: nacherzählen", "Ziel erreicht? Heute habe ich gelernt, …",
]));

doc.push(avivaPhase("A", "Ankommen: Ziel, Ablauf und Testrückgabe", 10));
doc.push(...steps([
  ["zuhoeren", { say: "Stifte liegen. Ihr hört nur zu. Links steht unser Ziel und der Ablauf für heute. Zuerst bekommt ihr die Leistungskontrolle zurück. Dann wiederholen wir die Lesetechniken und lesen Paddy Fingal weiter." }],
  ["luecken", { say: "Jetzt arbeitet ihr allein. Schaut euren Test durch. Sucht euch einen Fehler aus und verbessert ihn unter dem Test mit Grün.", board: board("Seitentafel rechts · Notenschlüssel", [bl([N("Notenschlüssel (25 P.):  ", { bold: true }), N("1 = 25–23,5 · 2 = 23–20 · 3 = 19,5–16,5 · 4 = 16–12,5 · 5 = 12–6 · 6 = 5,5–0")])]), do: ["Austeilen, während alle schon arbeiten. Wer keinen Fehler hat, prüft die Zusatzaufgabe (Futur II).", "Häufige Fehler kurz an der Tafel nennen (z. B. Hilfsverb bei Perfekt/Plusquamperfekt, Partizip II falsch gebildet).", "Tests einsammeln oder in die Mappe – wie gewohnt. Ansage: „Tests weg. Lehrbuch S. 136 auf.“"] }],
]));

doc.push(avivaPhase("V", "Vorwissen aktivieren: Lesetechniken und Ü1", 9));
doc.push(board("Mitteltafel · Wiederholung", [
  bl([N("Welche Lesetechnik passt? ", { bold: true }), N("überfliegend · orientierend · genau")]),
  bl("A  Du suchst im Lehrbuch, was das Wort „Hüne“ bedeutet."),
  bl("B  Du willst schnell wissen, ob die Ballade lustig oder traurig endet."),
  bl("C  Du sollst erklären, welche List Paddys Frau anwendet.", { after: 80 }),
  bl([N("Ü1 a  ", { bold: true }), N("Überfliege die Ballade (1 Min.). Schreibe: In der Ballade geht es um …")]),
  bl([N("Ü1 b  ", { bold: true }), N("Lies orientierend. Notiere die Antwort und die Strophe.")]),
  ...["1  Wie lang ist Paddys Bart?", "2  Woher kommt der Fremde?", "3  Wie heißt Paddys Frau?", "4  Was guckt aus dem Bett heraus?", "* 5  Was ist eine Furt? (Tipp: Worterklärungen)"].map((x) => bl("    " + x)),
]));
doc.push(p(t(""), { after: 60 }));
doc.push(...steps([
  ["luecken", { say: "Heft auf, Überschrift: Paddy Fingal – Inhalt erschließen. Schreibt A, B, C untereinander und die passende Lesetechnik dahinter. Eine Minute.", sol: "A orientierend · B überfliegend · C genau", do: ["Zu jeder Situation einen Schüler kurz begründen lassen („weil ich nur … suche“). Wer den Merksatz nicht mehr weiß, bekommt den Schnipsel."] }],
  ["luecken", { say: "Ü1 a: Überfliegen. Genau eine Minute. Start! … Stopp. Buch zu. Schreibt den Satz: In der Ballade geht es um …", sol: "z. B.: … um den Riesen Paddy Fingal und einen noch größeren Schotten, vor dem er sich mit einer List seiner Frau rettet." }],
  ["luecken", { say: "Ü1 b: Lest orientierend. Sucht nur die Antworten und notiert die Strophe. Vier Minuten. Fertig? Stift hinlegen.", sol: "1 sechs Ellen (Str. 1) · 2 aus Schottland (Str. 2) · 3 Schaya (Str. 4) · 4 nur seine Nase (Str. 5) · *5 flache Stelle in einem Fluss oder Bach (Worterklärung 7)" }],
  ["kontrolle", { say: "Vergleicht mit eurem Nachbarn und verbessert mit Grün.", do: ["Nur strittige Antworten im Plenum klären."] }],
]));

doc.push(avivaPhase("I", "Informieren: Auftrag LB-Aufgabe 1", 2));
doc.push(board("Mitteltafel · LB-Aufgabe 1", [
  bl([N("1 a  ", { bold: true }), N("Lest die Ballade im Tandem. Beantwortet im Heft in ganzen Sätzen, mit Strophe:")]),
  bl("    1  Wie sieht Paddy Fingal aus?"),
  bl("    2  Was möchte der Riese aus Schottland von Paddy?"),
  bl("    3  Welche List wendet Paddys Frau an?", { after: 80 }),
  bl([N("* 1 b  ", { bold: true }), N("Erzählt euch die Ballade nach: A Strophe 1–5, B Strophe 6–11.")]),
  bl([N("Checkliste: "), N("Titel und Autor · richtige Reihenfolge · nichts dazuerfinden · wichtige Details", { bold: true })]),
]));
doc.push(p(t(""), { after: 60 }));
doc.push(...steps([["zuhoeren", { say: "Stifte liegen. Ihr lest jetzt im Tandem weiter: Einer liest eine Strophe vor, der andere liest leise mit und hilft bei Fehlern, dann Wechsel. Danach beantwortet ihr die drei Fragen. Für Frage 3 müsst ihr genau lesen.", do: ["Methodenkarte „Tandemlesen“ (Tandems lesestärker/leseschwächer wie festgelegt)."] }]]));

doc.push(avivaPhase("V", "Verarbeiten: weiterlesen und LB-Aufgabe 1 a", 15));
doc.push(...steps([
  ["partner", { say: "Jetzt arbeitet ihr im Tandem. Die Aufgabe steht an der Tafel. Acht Minuten lesen, dann schreibt jeder die Antworten in sein Heft. Fertig? Stift hinlegen.", sol: "1 Ein Riese mit einem schwarzen Bart, sechs Ellen lang, Fäuste so groß wie Tonnen; er wirft mit einem Mastbaum (Str. 1). · 2 Der Schotte will herausfinden, wer von beiden stärker ist (Str. 2). · 3 Schaya steckt Paddy ins Bett, türmt die Kissen hoch und gibt ihn als ihr Baby aus. Der Schotte sieht nur die riesige Nase, denkt an einen riesigen Vater und flieht (Str. 5–10).", do: ["Herumgehen, beim Tandemlesen zuhören.", "Wer fertig ist: * 1 b nach Checkliste."] }],
]));

doc.push(avivaPhase("A", "Auswerten: Ziel erreicht?", 5));
doc.push(...steps([
  ["tafel", { say: "Wer drankommt, schreibt die Strophe zu einer Frage an die Tafel und nennt die Antwort. Alle prüfen mit und verbessern mit Grün.", do: ["Kreide-Kette: drei Schüler, je eine Frage aus 1 a. Bei Frage 3 nachfragen: „Welche Lesetechnik hast du gebraucht?“ (genau)."] }],
  ["luecken", { say: "Schaut auf unser Ziel links. Schreibt ins Heft: Heute habe ich gelernt, …", do: ["Zwei Schüler lesen vor. Rückgriff auf den Ablauf: Was mit * nicht geschafft wurde, kommt am Freitag als Wiederholung."] }],
]));
const pageProps = { page: { size: { width: 11906, height: 16838 }, margin: { top: 900, bottom: 800, left: 850, right: 850, header: 400, footer: 400 } } };
const styles = { default: { document: { run: { font: "Calibri", size: 22 } } } };
Packer.toBuffer(new Document({ styles, sections: [{ properties: pageProps, headers: { default: header("Balladen · Do 15.10.") }, footers: { default: footer("Tafelskript") }, children: doc }] })).then((buf) => fs.writeFileSync("Tafelskript_Do_15-10_Testrueckgabe_Paddy_Fingal.docx", buf));
