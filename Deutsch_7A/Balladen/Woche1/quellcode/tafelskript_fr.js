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
// Tafelskript Fr 16.10. (Einzelstunde): Rückgabe Mitarbeit · John Maynard gliedern – AVIVA
// ================================================================
const merkTafel = (nr, title, children) => board(`Seitentafel · Merksatz ${nr} – abschreiben, rot umrahmen`, [
  p([t(title, { bold: true, color: C.hv, font: "Cambria", size: 22 })], { after: 40 }),
  ...children,
]);
const avivaPhase = (k, title, min) => p([t(`${k}  `, { font: "Cambria", size: 28, bold: true, color: C.hv }), t(title, { font: "Cambria", size: 28, bold: true, color: NAVY }), t(`   ${min} Min.`, { size: 20, bold: true, color: "C77C12" })], { before: 200, after: 100, keepNext: true });
const zielTafel = (ziel, ablauf) => board("Seitentafel links · bleibt die ganze Stunde stehen", [
  p([t("Ziel: ", { bold: true, size: 21, color: NAVY }), t(ziel, { size: 21, italics: true })], { after: 60 }),
  p(t("Ablauf", { bold: true, size: 20, color: NAVY }), { after: 20 }),
  ...ablauf.map((a, i) => p(t(`${i + 1}  ${a}`, { size: 20 }), { after: 10 })),
]);

const doc = [];
doc.push(kicker("Tafelskript für die Lehrkraft · Fr 16.10. · Einzelstunde (45 Min.) · Planungsmodell AVIVA"), h1("John Maynard – eine Ballade gliedern"));
doc.push(box([
  p([S("Bezug: ", { bold: true }), S("Am Mittwoch hat die Klasse die Ballade gehört und die drei Lesetechniken kennengelernt; Strophe 5 und 6 wurden schon genau gelesen. Heute wird das genaue Lesen auf die ganze Ballade angewendet. Gestern war KA 1 – die Stunde ist kurz getaktet und endet vor den Ferien mit einem klaren Ergebnis.")], { after: 0 }),
], LIGHT2));
doc.push(p([S("Lernziel: ", { bold: true }), S("Ich kann eine Ballade in Abschnitte gliedern und passende Zwischenüberschriften formulieren.", { italics: true })], { before: 120, after: 40 }));
doc.push(p([S("Antwort auf „Was habt ihr heute gelernt?“: ", { bold: true }), S("„Ein neuer Abschnitt beginnt, wenn sich Ort, Zeit, Figuren oder die Handlung ändern. Eine Zwischenüberschrift nennt den Kern in wenigen Wörtern. Mit den Überschriften kann ich die Geschichte nacherzählen.“")], { after: 100 }));

doc.push(h3("Ablauf"));
doc.push(grid([900, 1300, 3400, 4606], ["Zeit", "AVIVA", "Phase", "Alle Schüler …"], [
  ["5′", "A", "Ankommen: Rückgabe Mitarbeit", "lesen die Rückmeldung still"],
  ["5′", "V", "Vorwissen: Lesetechniken", "beantworten drei Fragen im Heft"],
  ["7′", "I", "Merksatz 2: Einen Text gliedern", "schreiben ab; erster Abschnitt gemeinsam"],
  ["20′", "V", "Ü2 John Maynard gliedern", "lesen im Tandem, gliedern zu zweit, vergleichen in der Kreide-Kette"],
  ["5′", "A", "Auswerten: Ziel erreicht?", "erzählen mit den Überschriften nach, schreiben „Heute habe ich gelernt, …“"],
], { bold: [true, true, true, false], colors: [C.hv, C.hv, NAVY, null], h: 360 }));
doc.push(p([S("Puffer: ", { bold: true, color: C.hv }), S("3 Minuten. Mit * markiert: die Nacherzählung in der Auswertung (kann durch den Satz „Heute habe ich gelernt“ allein ersetzt werden).")], { before: 80, after: 60 }));

doc.push(h3("So sieht das Heft am Ende aus"));
doc.push(table([2300, 7906], [
  ["Überschrift", "John Maynard – eine Ballade gliedern · Datum rechts"],
  ["Heft", "Drei Antworten zu den Lesetechniken"],
  ["Merksatz 2", "Einen Text gliedern (rot umrahmt)"],
  ["Ü2", "Schnipsel „John Maynard gliedern“ (eingeklebt, mit Grün verbessert)"],
  ["Abschluss", "„Heute habe ich gelernt, …“"],
].map(([a, b2]) => row([tc(a, 2300, { bold: true, color: a.startsWith("Merk") ? C.hv : a === "Überschrift" ? NAVY : "C77C12", fill: LIGHT, size: 20 }), tc(b2, 7906, { size: 20 })], 340))));
doc.push(p([S("Material: ", { bold: true }), S("Lehrbuch · Mitarbeitsbögen mit Einschätzung und Note · Methodenkarte Tandemlesen · Druckvorlage: Ü2 (2 pro A4).")], { before: 100, after: 40 }));
doc.push(p([S("Vorbereitung: ", { bold: true }), S("Ziel und Ablauf links an die Seitentafel, Merksatz 2 rechts (zugeklappt). In der Tafelmitte untereinander die Strophennummern 1–9 anschreiben, rechts daneben Platz für Überschriften.")], { after: 60 }));
doc.push(zielTafel("Ich kann eine Ballade in Abschnitte gliedern und passende Zwischenüberschriften formulieren.", [
  "Rückgabe Mitarbeit", "Lesetechniken wiederholen", "Merksatz 2: Einen Text gliedern", "Ü2 John Maynard gliedern", "Was habe ich heute gelernt? (Nacherzählen *)",
]));

// ---------- A ----------
doc.push(avivaPhase("A", "Ankommen und einstimmen: Rückgabe Mitarbeit", 5));
doc.push(...steps([
  ["zuhoeren", { say: "Ihr bekommt eure Bögen zurück. Rechts steht meine Einschätzung und eure Note. Lest still. Wer darüber sprechen möchte, kommt in der Arbeitsphase kurz zu mir.", do: ["Bögen austeilen. Kurze Einzelgespräche während Ü2 anbieten, nicht im Plenum."] }],
]));

// ---------- V ----------
doc.push(avivaPhase("V", "Vorwissen aktivieren: Lesetechniken", 5));
doc.push(...steps([
  ["luecken", { say: "Heft auf, Überschrift: John Maynard – eine Ballade gliedern. Dann beantwortet ihr allein drei Fragen an der Tafel. Zwei Minuten.", board: board("Mitte", [
    bl("1  Welche drei Lesetechniken kennt ihr?"),
    bl("2  Wie habt ihr am Mittwoch die Strophe mit dem Feuer schnell gefunden?"),
    bl("3  Welche Lesetechnik braucht ihr, um eine Ballade zu gliedern? Begründe.", { after: 0 }),
  ]), sol: "1 überfliegend, orientierend, genau · 2 orientierend: nach Schlüsselwörtern wie „Feuer“ gesucht · 3 genau – man muss jeden Abschnitt verstehen, um zu merken, wo sich etwas ändert" }],
  ["gemeinsam", { say: "Wer liest Antwort 3 vor?", do: ["Überleitung: „Heute lesen wir die ganze Ballade genau und gliedern sie.“"] }],
]));

// ---------- I ----------
doc.push(avivaPhase("I", "Informieren: Merksatz 2 Einen Text gliedern", 7));
doc.push(...steps([
  ["abschreiben", { say: "Jetzt schreibt ihr von der Seitentafel ab: Merksatz 2. Rot umrahmen.", board: merkTafel(2, "Einen Text gliedern", [
    p([S("Einen Text teile ich in "), S("Abschnitte", { bold: true }), S(". Ein neuer Abschnitt beginnt, wenn sich "), S("Ort, Zeit, Figuren oder die Handlung", { bold: true }), S(" ändern.")], { after: 30 }),
    p([S("Jeder Abschnitt bekommt eine "), S("Zwischenüberschrift", { bold: true }), S(": kurz, ohne ganzen Satz, sie nennt den Kern (z. B. „Feuer an Bord“).")], { after: 60 }),
    p([S("Wofür? ", { bold: true }), S("So behalte ich den Überblick über die Handlung – das brauche ich später für Inhaltsangabe und Analyse. "), S("Test: ", { bold: true }), S("Kann ich mit meinen Zwischenüberschriften allein die Geschichte nacherzählen?")], { after: 0 }),
  ]) }],
  ["gemeinsam", { say: "Stifte liegen. Den ersten Abschnitt machen wir zusammen. Lest Strophe 2. Was ist der Kern? Wo ändert sich etwas?", do: ["Ergebnis an die Tafel neben Strophe 2: „Fröhliche Fahrt nach Buffalo“. Kurz klären: Strophe 1 ist ein Rahmen – die Frage „Wer ist John Maynard?“ und die Antwort, die am Ende auf dem Grabstein wiederkehrt."] }],
]));

// ---------- V ----------
doc.push(avivaPhase("V", "Verarbeiten: Ü2 John Maynard gliedern", 20));
doc.push(...steps([
  ["partner", { say: "Jetzt arbeitet ihr im Tandem. Klebt den Schnipsel Ü2 ein. Lest die Ballade zuerst zusammen halblaut – wer stockt, bekommt sofort Hilfe. Dann entscheidet ihr, wo ein neuer Abschnitt beginnt, und schreibt für jeden Abschnitt eine Zwischenüberschrift. Vierzehn Minuten.", board: auftrag("Ü2 John Maynard gliedern", ["Lest die Ballade im Tandem halblaut.", "Entscheidet, wo ein neuer Abschnitt beginnt (Ort, Zeit, Figuren, Handlung).", "Tragt Strophen und Zwischenüberschrift in den Schnipsel ein."], [p(S("Hilfe: Wo ändert sich etwas? Was ist der Kern?", { size: 20, italics: true }), { after: 0 })], "14 Minuten · Tandem · halblaut, dann leise"),
    sol: "Möglichkeit: Str. 1 Die Frage nach John Maynard · Str. 2 Fröhliche Fahrt nach Buffalo · Str. 3 Feuer an Bord! · Str. 4 Panik der Passagiere · Str. 5–6 John Maynard hält aus · Str. 7 Alle gerettet – einer fehlt · Str. 8–9 Die Stadt ehrt ihren Helden. Andere Abgrenzungen sind möglich, wenn sie begründet werden (z. B. Str. 5 und 6 getrennt).",
    do: ["Methodenkarte „Tandemlesen“: dieselben Tandems wie am Mittwoch.", "Herumgehen; Einzelgespräche zu den Mitarbeitsbögen hier führen.", "Typische Schwierigkeit: ganze Sätze statt Überschriften – auf den Merksatz verweisen."] }],
  ["tafel", { say: "Kreide-Kette: Wer drankommt, schreibt eure Überschrift neben die passende Strophe an die Tafel. Alle anderen vergleichen mit ihrem Schnipsel und schreiben eine bessere Idee mit Grün daneben.", do: ["Ein Paar pro Abschnitt, sieben Schüler. Bei Abweichungen begründen lassen: „Warum beginnt hier ein neuer Abschnitt?“"] }],
]));

// ---------- A ----------
doc.push(avivaPhase("A", "Auswerten: Ziel erreicht?", 5));
doc.push(...steps([
  ["gemeinsam", { say: "* Test aus dem Merksatz: Einer erzählt die Ballade nur mit den Überschriften an der Tafel nach. Alle anderen prüfen: Fehlt etwas?", do: ["Ein Schüler erzählt in fünf, sechs Sätzen nach."] }],
  ["luecken", { say: "Schaut auf unser Ziel. Schreibt ins Heft: Heute habe ich gelernt, …", do: ["Zwei Schüler lesen vor. Ausblick nach den Ferien: „Warum ist eine Ballade Erzählung, Gedicht und Theater zugleich?“ – schöne Ferien."] }],
]));

const pageProps = { page: { size: { width: 11906, height: 16838 }, margin: { top: 900, bottom: 800, left: 850, right: 850, header: 400, footer: 400 } } };
const styles = { default: { document: { run: { font: "Calibri", size: 22 } } } };
Packer.toBuffer(new Document({ styles, sections: [{ properties: pageProps, headers: { default: header("Balladen · Fr 16.10.") }, footers: { default: footer("Tafelskript") }, children: doc }] })).then((buf) => fs.writeFileSync("Tafelskript_Fr_16-10_Gliedern.docx", buf));
