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

const none = { style: BorderStyle.NONE, size: 0, color: NAVY };
const noBorders = { top: none, bottom: none, left: none, right: none };
const solid = (c = NAVY, s = 8) => ({ style: BorderStyle.SINGLE, size: s, color: c });
const allBorders = (b) => ({ top: b, bottom: b, left: b, right: b });
const dashed = { style: BorderStyle.DASHED, size: 8, color: "7A869E" };
const thin = allBorders(solid("B9C6E8", 6));

function cell(children, o = {}) {
  return new TableCell({
    children: Array.isArray(children) ? children : [children],
    width: { size: o.w, type: WidthType.DXA },
    shading: undefined, // Tinte sparen: keine Farbflächen im Druck
    borders: o.borders || thin,
    verticalAlign: o.valign || VerticalAlign.CENTER,
    margins: { top: o.m ?? 60, bottom: o.m ?? 60, left: o.ml ?? 100, right: 100 },
    columnSpan: o.span,
  });
}
const table = (widths, rows) => new Table({ width: { size: widths.reduce((a, b) => a + b, 0), type: WidthType.DXA }, columnWidths: widths, rows });
const row = (cells, h, rule = HeightRule.ATLEAST) => new TableRow({ children: cells, height: h ? { value: h, rule } : undefined, cantSplit: true });
const hdr = (txt, w, fill = NAVY, size = 20) => cell(p(t(txt, { bold: true, color: NAVY, size }), { align: AlignmentType.CENTER, after: 0 }), { w, fill });
const tc = (txt, w, o = {}) => cell(p(Array.isArray(txt) ? txt : t(txt, { size: o.size || 21, bold: o.bold, color: o.color, font: o.font, italics: o.it }), { after: 0, align: o.align }), { w, fill: o.fill });

// simple grid table: header + rows of strings/runs
function grid(widths, head, rows, o = {}) {
  return table(widths, [
    row(head.map((h, i) => hdr(h, widths[i], o.headFill ? o.headFill[i] : NAVY)), 420),
    ...rows.map((r) => row(r.map((v, i) => tc(v, widths[i], { font: o.fonts && o.fonts[i], align: o.align && o.align[i], bold: o.bold && o.bold[i], color: o.colors && o.colors[i] })), o.h || 460)),
  ]);
}

function box(children, fill = LIGHT2, border) {
  return table([W], [row([cell(children, { w: W, fill, borders: border || allBorders(solid("B9C6E8", 8)), m: 140, ml: 180 })])]);
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
    row([cell(p(t("TAFEL · " + title, { bold: true, color: NAVY, size: 18 }), { after: 0 }), { w: W, fill: "3D4A3F", borders: allBorders(solid("3D4A3F", 12)) })]),
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
  tafel: "kommt an die Tafel", kontrolle: "kontrolliert", partner: "arbeitet zu zweit", fertig: "gibt ab", gruppe: "arbeitet in der Gruppe",
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
    row([cell(p(t("TAFEL · " + title, { bold: true, color: NAVY, size: 16 }), { after: 0 }), { w: TW, fill: "3D4A3F", borders: allBorders(solid("3D4A3F", 12)), m: 30 })]),
    row([cell(children, { w: TW, fill: BOARD, borders: allBorders(solid("3D4A3F", 12)), m: 100, ml: 160 })]),
  ]);
}
const bl = (runs, o = {}) => p(Array.isArray(runs) ? runs : [N(runs)], Object.assign({ after: 30 }, o));
// Merksatz steht auf der Folie (nicht an der Tafel)
function folie(title, children) {
  return table([TW], [
    row([cell(p(t("FOLIE · " + title, { bold: true, color: NAVY, size: 16 }), { after: 0 }), { w: TW, fill: NAVY, borders: allBorders(solid(NAVY, 12)), m: 30 })]),
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



// Tafelskript Mi 18.11. (Doppelstunde, ca. 75 Min.): Einführung Sachtexte (LB S. 178–179) + Sprachtraining Zeitformen – AVIVA, mit Präsentation. Ohne Farbflächen.
const avivaPhase = (k, title, min) => p([t(`${k}  `, { font: "Cambria", size: 28, bold: true, color: C.hv }), t(title, { font: "Cambria", size: 28, bold: true, color: NAVY }), t(`   ${min} Min.`, { size: 20, bold: true, color: "C77C12" })], { before: 200, after: 100, keepNext: true });
const Z1 = "Ich kann an einem Text zeigen, woran ich einen Sachtext erkenne, und Merkmale mit Zeilenangabe belegen.";
const Z2 = "Ich kann Zeitformen in einem Sachtext bestimmen und erklären, warum sie verwendet werden.";
const doc = [];
doc.push(kicker("Tafelskript für die Lehrkraft · Mi 18.11. · Doppelstunde (ca. 75 Min., 15 Min. fallen weg) · AVIVA · mit Präsentation (18 Folien)"), h1("Sachtexte verstehen – Einführung"));
doc.push(box([p([S("Bezug: ", { bold: true }), S("Die Balladen-Einheit endet mit John Maynard (Test Paddy Fingal, Stilmittel-Blatt in Moodle). Neue Einheit laut SchiC Thema 2 „Sachtexte verstehen“ (ca. 4 Wochen) mit KA 3: Informationen aus einem Sachtext entnehmen und zusammenfassen, Lesestrategien des Leselotsen. Heute: Merkmale eines Sachtexts am LB-Text „Die Rauchzeichen der Native Americans“ (LB S. 178–179, 3 Teilüberschriften, Z. 1–59), LB-Aufgaben 1 (vor dem Lesen) und 2 (beim Lesen). Zum Vorwissen Beispiele zeigen und nach dem Fokus fragen (Wunsch der Lehrkraft). Merkmale nur am Sachtext (kein Kontrasttext). Mittwochs Sprachwiederholung: Zeitformen im Sachtext (Präsens, Perfekt, Präteritum, Plusquamperfekt – alle im Text) und Wiederholung aus der Verb-Einheit. LB 3 und 4 (genau lesen, Notizen, zusammenfassen) folgen in der nächsten Stunde.")], { after: 0 })]));
doc.push(p([S("Ziel Teil 1: ", { bold: true }), S(Z1, { italics: true })], { before: 120, after: 20 }));
doc.push(p([S("Ziel Teil 2: ", { bold: true }), S(Z2, { italics: true })], { after: 40 }));
doc.push(p([S("Antwort auf „Was habt ihr heute gelernt?“: ", { bold: true }), S("„Ein Sachtext informiert sachlich über Tatsachen. Ich erkenne ihn an Überschrift, Teilüberschriften, Fakten, erklärten Fachbegriffen, Beispielen und sachlicher Sprache – und kann das mit Zeilen belegen. Über früher wird im Präteritum berichtet, was heute gilt, steht im Präsens.“")], { after: 80 }));
doc.push(board("Seitentafel links · vor der Stunde anschreiben, bleibt stehen", [
  p([t("Teil 1 – Ziel: ", { bold: true, size: 20, color: NAVY }), t(Z1, { size: 20, italics: true })], { after: 40 }),
  p([t("Teil 2 – Ziel: ", { bold: true, size: 20, color: NAVY }), t(Z2, { size: 20, italics: true })], { after: 40 }),
  ...["Beispiele: Worauf liegt der Fokus?", "LB 1 Vor dem Lesen · LB 2 Beim Lesen", "Merkmale eines Sachtexts · Ü1 Merkmale belegen", "Sprachtraining: Ü2 Zeitformen · * Ü3 Wiederholung", "Zielcheck"].map((a, i) => p(t(`${i + 1}  ${a}`, { size: 20 }), { after: 10 })),
  p([t("Vermutungen (LB 1): ", { bold: true, size: 20, color: NAVY }), t("… (während der Stunde ergänzen, nach dem Lesen prüfen)", { size: 20, italics: true })], { before: 40, after: 0 }),
]));
doc.push(p(t(""), { after: 40 }));
doc.push(h3("Ablauf"));
doc.push(grid([800, 900, 1200, 3300, 4006], ["Zeit", "AVIVA", "Folie", "Phase", "Alle Schüler …"], [
  ["5′", "A", "1–3", "Ziele, Ablauf, Weg der Einheit (Leselotse → KA 3)", "lesen Ziele und Ablauf, Überschrift ins Heft"],
  ["7′", "V", "4–5", "Beispiele: Worauf liegt der Fokus?", "lesen fünf Kurztexte, besprechen zu zweit, nennen den Fokus"],
  ["6′", "I", "6", "LB 1 Vor dem Lesen", "verschaffen sich einen Überblick, vermuten"],
  ["10′", "V", "7", "LB 2 Beim Lesen (Tandem)", "lesen zügig, erklären sich gegenseitig"],
  ["8′", "I", "8–9", "Merkmale entdecken → Merksatz-Schnipsel", "notieren zu zweit Merkmale mit Zeile, kleben ein"],
  ["12′", "V", "10–11", "Ü1 Merkmale belegen", "Tabelle im Heft, vergleichen mit Grün"],
  ["7′", "V", "12–14", "Teil 2: Ü2 Zeitformen im Sachtext", "bestimmen und begründen, vergleichen mit Grün"],
  ["*8′", "V", "15–16", "* Ü3 Wiederholung Zeitformen", "setzen Sätze in die angegebene Zeitform"],
  ["8′", "A", "17–18", "Zielcheck", "lösen allein, verbessern mit Grün"],
], { bold: [true, true, false, true, false], colors: [C.hv, C.hv, MUTED, NAVY, null], h: 320 }));
doc.push(p([S("Zeit: ", { bold: true, color: C.hv }), S("71 Min. mit * Ü3. Wenn es eng wird: * Ü3 als Hausaufgabe oder in die Sprachwiederholung der nächsten Woche.")], { before: 80, after: 60 }));
doc.push(h3("So sieht das Heft am Ende aus"));
doc.push(table([2300, 7906], [
  ["Überschrift", "Neues Thema: Sachtexte verstehen · „Die Rauchzeichen der Native Americans“ (LB S. 178–179) · Datum"],
  ["Vorwissen", "Fokus von Sachtexten: informieren (Stichwort)"],
  ["Merkmale", "Notizen zu zweit · Merksatz-Schnipsel eingeklebt"],
  ["Ü1", "Tabelle: Merkmal | Beispiel | Zeile – sechs Zeilen, mit Grün verbessert"],
  ["Sprachtraining", "Ü2 vier Zeitformen mit Begründung · * Ü3 vier Sätze · Zielcheck"],
].map(([a, b]) => row([tc(a, 2300, { bold: true }), tc(b, 7906)]))));
doc.push(p([S("Material: ", { bold: true }), S("Präsentation `Mittwoch_18-11_Deutsch_7A.pptx` · Lehrbuch S. 178–179 · Schnipsel `Merksatz_Sachtext_Zeitformen_Schnipsel` (2 pro A4) · Handreichung Polnisch für S-PL (`Handreichung_Polnisch_18-11`).")], { before: 80, after: 60 }));

doc.push(avivaPhase("A", "Ankommen: Ziele, Ablauf, Weg der Einheit", 5));
doc.push(...steps([["zuhoeren", { say: "Stifte liegen. Ihr hört nur zu. Mit John Maynard sind wir bei den Balladen fertig. Heute beginnt ein neues Thema: Sachtexte. In vier Wochen schreibt ihr dazu eine Klassenarbeit: Ihr entnehmt Informationen aus einem Sachtext und fasst sie zusammen. Dafür lernt ihr, wie ein Lotse zu lesen: vor dem Lesen, beim Lesen, nach dem Lesen. Jetzt schreibt ihr die Überschrift ab: Sachtexte verstehen.", do: ["Folie 1–3."] }]]));

doc.push(avivaPhase("V", "Vorwissen aktivieren: Beispiele – worauf liegt der Fokus?", 7));
doc.push(...steps([
  ["partner", { say: "Lest die fünf Texte auf der Folie. Überlegt zu zweit: Was wollen alle diese Texte? Worauf liegt der Fokus? Zwei Minuten.", do: ["Folie 4: Nachricht (Handball-WM 2027), Lexikon (Judo), Fachtext (Reißen), Trainingsplan, Tabelle (Schlafdauer)."] }],
  ["gemeinsam", { say: "Wer drankommt, nennt den Fokus. Und: Wo lest ihr solche Texte?", sol: "Fokus: informieren – über Tatsachen, sachlich, ohne Geschichte und ohne Meinung · Orte: Lehrbuch, Nachrichten-App, Trainingsplan, Beipackzettel, Internet", do: ["Folie 5 aufdecken. Kurz die Lesetechniken erinnern: überfliegend, orientierend, genau."] }],
]));

doc.push(avivaPhase("I", "LB 1 Vor dem Lesen", 6));
doc.push(...steps([["luecken", { say: "Jetzt arbeitet ihr erst allein: Lest nur die Überschrift und die einführende Information. Seht euch das Bild an. Überfliegt die drei Teilüberschriften und die ersten drei Zeilen. Dann besprecht zu zweit: Worum könnte es gehen? Was wisst ihr schon?", do: ["Folie 6. Methode: überfliegendes Lesen. 2–3 Vermutungen an der Seitentafel notieren."] }]]));

doc.push(avivaPhase("V", "LB 2 Beim Lesen (Tandem)", 10));
doc.push(...steps([["partner", { say: "Lest den Text zügig. Dann erklärt euch gegenseitig, was ihr über die Verständigung der Native Americans erfahren habt. Wer zuhört, ergänzt.", do: ["Folie 7. Methodenkarte: Tandem. Danach: Stimmen unsere Vermutungen? (Seitentafel abhaken)", "Worterklärungen: Fußnoten 1 und 2 im LB."] }]]));

doc.push(avivaPhase("I", "Merkmale entdecken – Merksatz", 8));
doc.push(...steps([
  ["partner", { say: "Woran merkt ihr, dass das ein Sachtext ist? Seht euch Aufbau, Inhalt und Sprache an. Notiert zu zweit mindestens vier Merkmale mit einer Zeile aus dem Text. Sechs Minuten.", do: ["Folie 8. Hilfsfragen: Wie ist der Text gegliedert? Was steht in den Fußnoten? Gibt es eine Meinung oder ein „ich“?"] }],
  ["gemeinsam", { say: "Wer drankommt, nennt ein Merkmal und die Zeile. Dann vergleichen wir mit dem Merksatz. Ich teile den Schnipsel aus – klebt ihn ein.", do: ["Folie 9: Kernsatz im roten Rahmen, sechs Merkmale, Wofür? und Test."] }],
]));

doc.push(avivaPhase("V", "Ü1 Merkmale mit Zeilenangabe belegen", 12));
doc.push(...steps([
  ["luecken", { say: "Jetzt arbeitet ihr allein. Zeichnet die Tabelle ins Heft: Merkmal, Beispiel aus dem Text, Zeile. Findet zu jedem der sechs Merkmale ein Beispiel. Zehn Minuten. Fertig? Stift hinlegen.", sol: "1 Überschrift + einführende Information · 2 Teilüberschriften Z. 1–2, 30–31, 46 · 3 Fakten Z. 3–5 (400 Jahre; Apachen, Hopis, Navajos, Sioux), Z. 49–53 (eine, zwei, drei Wolken) · 4 Fachbegriff Z. 5 „indigene Stämme“ (Fußnote 1), Z. 56–57 „allgemeingültiger Code“ (Fußnote 2) · 5 Beispiel Z. 20–22 · 6 sachlich z. B. Z. 15–17", do: ["Folie 10. Zeilenzählung: Z. 1 = erste Teilüberschrift."] }],
  ["kontrolle", { say: "Vergleicht mit der Lösung und verbessert mit Grün.", do: ["Folie 11. Andere passende Stellen gelten."] }],
]));

doc.push(avivaPhase("V", "Teil 2 · Sprachtraining: Zeitformen im Sachtext", 7));
doc.push(...steps([
  ["luecken", { say: "Jetzt arbeitet ihr allein. Bestimmt die Zeitform der Verben in den vier Zeilen und begründet in Stichworten, warum sie hier steht. Fünf Minuten.", sol: "Z. 13 weiß – Präsens (gilt heute) · Z. 13–14 kommuniziert haben – Perfekt · Z. 20 warnten – Präteritum (Bericht über früher) · Z. 23–24 aufgespürt hatten – Plusquamperfekt (Vorvergangenheit)", do: ["Folie 12–13. Farbcode: Hilfsverb rot, Partizip II gelb."] }],
  ["kontrolle", { say: "Vergleicht mit der Lösung und verbessert mit Grün.", do: ["Folie 14 mit Merke (steht auch auf dem Schnipsel)."] }],
]));
doc.push(avivaPhase("V", "* Ü3 Sprachwiederholung: Zeitformen bilden", 8));
doc.push(...steps([["luecken", { say: "Setzt die Sätze in die angegebene Zeitform. Wer fertig ist, kontrolliert mit der nächsten Folie.", sol: "a benutzten · b hat … gemeldet · c deckten … zu · d hatte … erklärt", do: ["Folie 15–16. Wenn die Zeit fehlt: Hausaufgabe."] }]]));

doc.push(avivaPhase("A", "Auswerten: Zielcheck", 8));
doc.push(...steps([
  ["luecken", { say: "Zielcheck. Jetzt arbeitet ihr allein im Heft. Vier Minuten.", sol: "1 Sachtext – Fakt mit Zahl und Quelle, sachlich · 2 kein Sachtext – Meinung („ich finde“), will überzeugen · 3 Präteritum (entfachten)", do: ["Folie 17."] }],
  ["kontrolle", { say: "Vergleicht und verbessert mit Grün. Haben wir die Ziele erreicht?", do: ["Folie 18. Notieren, wer noch Fehler hat. Ausblick: LB 3 und 4 in der nächsten Stunde."] }],
]));

// ---------- Reflexion ----------
const bul = (head, txt) => p([S("• ", { bold: true, color: C.hv }), ...(head ? [S(head + " ", { bold: true })] : []), S(txt)], { after: 40, indent: { left: 240, hanging: 240 } });
const refHead = (txt, col) => p(t(txt, { font: "Cambria", size: 22, bold: true, color: col }), { before: 140, after: 50, keepNext: true });
doc.push(new Paragraph({ pageBreakBefore: true, children: [t("FÜR DIE LEHRKRAFT · REFLEXION VOR UND NACH DER STUNDE", { size: 18, color: MUTED, bold: true })], spacing: { after: 20 } }), h3("Umsetzung der Erkenntnisse zur Unterrichtsverbesserung"));
doc.push(refHead("1  AVIVA – Struktur und Transparenz", C.hv));
doc.push(bul("Ankommen:", "Überblicksfolie mit zwei Ich-Zielen und Zeiten (Doppelstunde); zusätzlich der Weg der Einheit bis zur KA 3, damit die Relevanz von Anfang an klar ist."));
doc.push(bul("Vorwissen aktivieren:", "Fünf kurze Sachtexte aus der Lebenswelt (Handball, Judo, Gewichtheben, Trainingsplan, Schlaf) – die Klasse entdeckt den Fokus selbst (Wunsch der Lehrkraft)."));
doc.push(bul("Informieren:", "Merkmale werden am LB-Text entdeckt und erst danach im Merksatz gebündelt (Schnipsel, nicht abschreiben)."));
doc.push(bul("Verarbeiten:", "Erst das Werkzeug (Merksatz), dann die Anwendung mit Zeilenangabe – wie das Belegen bei den Balladen."));
doc.push(bul("Auswerten:", "Zielcheck mit drei kurzen Aufgaben genau zu beiden Zielen."));
doc.push(refHead("2  Direkte Instruktion (DI)", NAVY));
doc.push(bul("Ziel und Relevanz:", "Je ein überprüfbares Ziel; Relevanz: KA 3 und Referate."));
doc.push(bul("Ich – Wir – Ihr:", "Lehrkraft zeigt den Weg (Leselotse), Merkmale werden gemeinsam gesammelt, Ü1 und Ü2 allein."));
doc.push(bul("Rückmeldung:", "Lösungsfolien nach Fokus-Frage, Ü1, Ü2, Ü3 und Zielcheck."));
doc.push(bul("Klarheit:", "Feste Ansagen, gleiches Übungsformat, Zeilenzählung ausdrücklich erklärt."));
doc.push(refHead("3  Differenzierung", "2A8C8C"));
doc.push(bul("", "Tandem beim Lesen; Hilfsfragen beim Merkmale-Entdecken; * Ü3 für Schnelle bzw. als Hausaufgabe."));
doc.push(bul("", "S-PL: Handreichung mit Zielen, Ablauf, Fachbegriffen und Worterklärungen zum Rauchzeichen-Text auf Polnisch; Fachbegriffe deutsch, keine Lösungen."));
doc.push(refHead("4  Beobachtungsschwerpunkte", MUTED));
doc.push(bul("", "Erkennen die Schüler den Fokus „informieren“ selbst? Unterscheiden sie Sachtext und Meinung (Zielcheck 2)?"));
doc.push(bul("", "Gelingt das Belegen mit Zeilenangabe im Fließtext (neu gegenüber Strophe/Vers)?"));
doc.push(bul("", "Sitzen die Zeitformen aus der Verb-Einheit noch (besonders Plusquamperfekt)?"));

const pageProps = { page: { size: { width: 11906, height: 16838 }, margin: { top: 900, bottom: 800, left: 850, right: 850, header: 400, footer: 400 } } };
const styles = { default: { document: { run: { font: "Calibri", size: 22 } } } };
Packer.toBuffer(new Document({ styles, sections: [{ properties: pageProps, headers: { default: header("Sachtexte · Mi 18.11.") }, footers: { default: footer("Tafelskript") }, children: doc }] })).then((buf) => fs.writeFileSync(__dirname + "/../Tafelskript_Mi_18-11_Sachtexte_Einfuehrung.docx", buf));
