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



// Tafelskript Mo 9.11. (Einzelstunde, Hospitation): Leseband (Ich-Du-Wir-Würfeln) · Kreuzworträtsel Balladen · Sprachtraining: Fehler im Chat – AVIVA, mit Präsentation
// Druckfassung ohne Farbflächen (Tinte sparen, Stand 6.11.)
const avivaPhase = (k, title, min) => p([t(`${k}  `, { font: "Cambria", size: 28, bold: true, color: C.hv }), t(title, { font: "Cambria", size: 28, bold: true, color: NAVY }), t(`   ${min} Min.`, { size: 20, bold: true, color: "C77C12" })], { before: 200, after: 100, keepNext: true });
const ZIEL = "Ich kann die Merkmale einer Ballade benennen und Nebensätze mit Subjunktion richtig schreiben.";
const doc = [];
doc.push(kicker("Tafelskript für die Lehrkraft · Mo 9.11. · Einzelstunde (45 Min.) · Hospitation · AVIVA · mit Präsentation (14 Folien)"), h1("Balladen – alles im Griff?"));
doc.push(box([p([S("Bezug: ", { bold: true }), S("Paddy Fingal ist inhaltlich und formal erarbeitet: Inhalt und Belegen (Mi 4.11., „sehr gut geklappt“), Reime (Do 5.11.), Ballade als Ur-Ei mit den drei Gattungen (Fr 6.11., „ganz gut“; die Fragen auf dem Ur-Ei-Schnipsel waren etwas uneindeutig). Heute wird alles gesichert, bevor die nächste Ballade kommt. Das Leseband (Ich-Du-Wir-Würfeln, eigener Text) steht am Anfang. Danach holt ein Kreuzworträtsel alle Fachbegriffe zurück, hierarchisch vom Allgemeinen ins Detail (Ballade → Gattungen → Merkmale). Der gemeinsame nächste Schritt (Begründung mit „weil“) ist die Brücke zum Sprachtraining: Die Subjunktionen vom Mittwoch werden an einem Chat mit echten Fehlern wiederholt (Fehleranalyse statt Lückenfüllen). Alle Aufgaben sind eindeutig gestellt: Nummer, Zahl der Antworten, Antwortformat.")], { after: 0 })]));
doc.push(p([S("Lernziel: ", { bold: true }), S(ZIEL, { italics: true })], { before: 120, after: 40 }));
doc.push(p([S("Antwort auf „Was habt ihr heute gelernt?“: ", { bold: true }), S("„Eine Ballade hat Merkmale von Lyrik (Strophen, Verse, Reim), Epik (Erzähler) und Dramatik (Dialog ohne Erzähler). Im Nebensatz mit weil oder dass steht das finite Verb am Ende, und vor der Subjunktion steht ein Komma.“ Überprüft wird das im Zielcheck.")], { after: 80 }));
doc.push(board("Seitentafel links · vor der Stunde anschreiben, bleibt stehen", [
  p([t("Ziel: ", { bold: true, size: 21, color: NAVY }), t(ZIEL, { size: 21, italics: true })], { after: 60 }),
  ...["Leseband: Ich-Du-Wir-Würfeln", "Ü1 Kreuzworträtsel Balladen + nächster Schritt gemeinsam", "Ü2 Sprachtraining: Fehler im Chat (* Nachricht 5)", "Zielcheck"].map((a, i) => p(t(`${i + 1}  ${a}`, { size: 20 }), { after: 10 })),
]));
doc.push(p(t(""), { after: 40 }));
doc.push(h3("Ablauf"));
doc.push(grid([800, 900, 1200, 3300, 4006], ["Zeit", "AVIVA", "Folie", "Phase", "Alle Schüler …"], [
  ["2′", "A", "1–2", "Ankommen: Ziel und Ablauf", "lesen Ziel und Ablauf"],
  ["10′", "V", "3", "Leseband: Ich-Du-Wir-Würfeln (eigener Text)", "lesen laut oder still mit, fassen jeden Abschnitt in einem Satz zusammen"],
  ["8′", "V", "4–5", "Ü1 Kreuzworträtsel Balladen", "lösen allein, verbessern mit der Lösungsfolie"],
  ["5′", "V", "6", "Ü1 Nr. 2: nächster Schritt gemeinsam", "begründen gemeinsam mit „weil“, alle schreiben mit"],
  ["3′", "I", "7–8", "Was fällt am weil-Satz auf? Merksatz (Wiederholung)", "nennen Subjunktion, finites Verb, Komma"],
  ["8′", "V", "9", "Ü2 Fehler im Chat", "verbessern zu zweit, jede/r ins eigene Heft"],
  ["3′", "V", "10–12", "Ü2 Lösung aufdecken", "lesen vor, verbessern mit Grün"],
  ["6′", "A", "13–14", "Zielcheck", "lösen allein, verbessern mit Grün"],
], { bold: [true, true, false, true, false], colors: [C.hv, C.hv, MUTED, NAVY, null], h: 320 }));
doc.push(p([S("Zeit: ", { bold: true, color: C.hv }), S("45 Min. ohne Puffer – deshalb Kürzungsoptionen: * Nachricht 5 im Chat entfällt; das Leseband nach 8 Min. beenden; Folie 8 (Merksatz) nur zeigen, wenn bei Folie 7 Unsicherheit da ist.")], { before: 80, after: 60 }));
doc.push(h3("So sieht das Heft am Ende aus"));
doc.push(table([2300, 7906], [
  ["Überschrift", "Balladen – Wiederholung · Datum rechts"],
  ["Ü1", "Kreuzworträtsel eingeklebt und ausgefüllt, mit Grün kontrolliert · Nr. 2: „Paddy Fingal“ ist eine Ballade, weil …"],
  ["Ü2", "Vier (*fünf) verbesserte Sätze aus dem Chat, jeweils mit K oder V, Subjunktion eingekreist, finites Verb unterstrichen"],
  ["Zielcheck", "Zwei Fachbegriffe · zwei verbundene Sätze mit unterstrichenem finitem Verb, mit Grün kontrolliert"],
].map(([a, b]) => row([tc(a, 2300, { bold: true }), tc(b, 7906)]))));
doc.push(p([S("Material: ", { bold: true }), S("Präsentation `Montag_09-11_Deutsch_7A.pptx` (Lösungen in den Notizen und auf Aufdeck-Folien) · Leseband-Text und ein Würfel pro Tischgruppe · Schnipsel `Schnipsel_Kreuzwortraetsel_Balladen` (A4 quer, 2 pro Blatt, Lösung auf Seite 2) · Lehrbuch S. 136 f. bereithalten · Handreichung Polnisch für S-PL (`Handreichung_Polnisch_09-11`).")], { before: 80, after: 60 }));

doc.push(avivaPhase("A", "Ankommen: Ziel und Ablauf", 2));
doc.push(...steps([["zuhoeren", { say: "Stifte liegen. Ihr hört nur zu. Heute holen wir alles zu den Balladen zurück ins Gedächtnis. Am Ende schreibt ihr Sätze mit Subjunktion ohne Fehler. Ziel und Ablauf stehen auf der Folie und links an der Tafel. Im Zielcheck prüft ihr, ob ihr das Ziel erreicht habt.", do: ["Folie 1–2."] }]]));

doc.push(avivaPhase("V", "Vorwissen aktivieren: Leseband · Ich-Du-Wir-Würfeln", 10));
doc.push(...steps([
  ["gruppe", { say: "Jetzt arbeitet ihr in der Tischgruppe. Würfelt reihum. Bei 1 oder 2 liest, wer gewürfelt hat. Bei 3 oder 4 liest die Person links. Bei 5 oder 6 lest ihr alle gemeinsam. Alle anderen lesen still mit. Nach jedem Abschnitt sagt die Person, die gewürfelt hat, in einem Satz, worum es ging. Zehn Minuten.", do: ["Folie 3. Methode nach Methodenkarte: Ich-Du-Wir-Würfeln. Eigener Text aus dem Leseband, ein Abschnitt = ein Absatz. Ein Würfel pro Tischgruppe (3–4 Schüler), kein Umräumen.", "Herumgehen: Tempo, Betonung, Zusammenfassung in einem Satz einfordern."] }],
  ["fertig", { say: "Stopp. Würfel in die Mitte. Stifte liegen.", do: ["Festes Schlusssignal, Würfel einsammeln lassen (eine Person pro Gruppe)."] }],
]));

doc.push(avivaPhase("V", "Vorwissen aktivieren: Ü1 Kreuzworträtsel Balladen", 8));
doc.push(...steps([
  ["luecken", { say: "Ich teile den Schnipsel aus. Klebt ihn ein. Jetzt arbeitet ihr allein. Tragt die Wörter der Reihe nach ein, von 1 bis 13. Die Nummern gehen vom Allgemeinen ins Detail. Bei Nummer 2 unten schreibt ihr noch nichts – das machen wir gleich gemeinsam. Acht Minuten. Fertig? Stift hinlegen.", sol: "1 BALLADE · 2 LYRIK · 3 STROPHE · 4 VERS · 5 REIM · 6 PAARREIM · 7 KREUZREIM · 8 UMARMENDER · 9 KEHRREIM · 10 EPIK · 11 ERZÄHLER · 12 DRAMATIK · 13 DIALOG", do: ["Folie 4. Herumgehen. Typische Hürden: Nr. 8 nur das Adjektiv (UMARMENDER), Nr. 11 Ä in einem Kästchen, Nr. 13 DIALOG (Dramatik = ohne Erzähler)."] }],
  ["kontrolle", { say: "Vergleicht mit der Lösung auf der Folie und verbessert mit Grün.", do: ["Folie 5: Gitter und Begriffsbaum. Am Baum kurz nachfragen: Was ist der Oberbegriff? Welche drei Gattungen stecken in der Ballade? Was unterscheidet Epik und Dramatik? (Erzähler – ja oder nein)."] }],
]));

doc.push(avivaPhase("V", "Vorwissen aktivieren: Ü1 Nr. 2 – nächster Schritt gemeinsam", 5));
doc.push(...steps([["gemeinsam", { say: "Jetzt machen wir den nächsten Schritt gemeinsam. Warum ist Paddy Fingal eine Ballade? Nehmt Lösungswörter aus dem Rätsel – je eins zu Lyrik, Epik und Dramatik. Ich fange an, ihr ergänzt. Alle schreiben auf den Schnipsel mit.", sol: "„Paddy Fingal“ ist eine Ballade, weil sie Strophen mit Reimen hat, weil ein Erzähler die Geschichte erzählt und weil in Str. 7 Schaya ohne Erzähler spricht.", do: ["Folie 6 (Ich – Wir): Lehrkraft macht Lyrik vor, Schüler nennen Epik und Dramatik, Lehrkraft fasst zusammen. Andere Belege gelten, wenn das Merkmal stimmt (z. B. Kehrreim; Str. 9 Rede des Schotten)."] }]]));

doc.push(avivaPhase("I", "Informieren: Was fällt am weil-Satz auf?", 3));
doc.push(...steps([
  ["gemeinsam", { say: "Schaut auf unseren Satz. Was fällt euch an den weil-Sätzen auf?", sol: "weil = Subjunktion, sie leitet den Nebensatz ein · das finite Verb steht am Ende (hat, erzählt, spricht) · vor weil steht ein Komma", do: ["Folie 7: Satz mit Farben. Verschiedene Schüler drannehmen."] }],
  ["zuhoeren", { say: "Stifte liegen. Ihr hört nur zu. Das ist unser Merksatz vom Mittwoch. Er klebt schon im Heft. Achtet gleich im Chat genau auf diese drei Punkte.", do: ["Folie 8 nur, wenn bei Folie 7 Unsicherheit da ist: Kernsatz im roten Rahmen, Beispiel, Wofür? und Test (nicht abschreiben)."] }],
]));

doc.push(avivaPhase("V", "Verarbeiten: Ü2 Sprachtraining – Fehler im Chat", 11));
doc.push(...steps([
  ["partner", { say: "Jetzt arbeitet ihr zu zweit. Der Schotte schreibt nach seiner Flucht mit seinem Bruder. In jeder Nachricht mit Nummer steckt genau ein Fehler: Komma oder Verbstellung. Schreibt die Sätze richtig ins eigene Heft und notiert K oder V. Kreist die Subjunktion ein und unterstreicht das finite Verb. Acht Minuten. Nachricht 5 mit Stern macht, wer fertig ist.", sol: "1 …weggerannt, weil ich das Baby gesehen habe. (K) · 2 …, dass Paddy einen schwarzen Bart hat. (V) · 3 …, dass ich leise sein muss. (V) · 4 …, weil ich neugierig war. (V) · *5 Ich glaube, dass Paddy selbst im Bett lag. (K)", do: ["Folie 9. Herumgehen. Typische Hürde: Nr. 1 und 5 klingen richtig – der Fehler ist das fehlende Komma. Bei Nr. 1 und 3: zwei Verben am Ende, das finite steht ganz hinten (gesehen habe, sein muss)."] }],
  ["kontrolle", { say: "Wer drankommt, liest seinen Satz vor und nennt K oder V. Dann decke ich auf. Alle verbessern mit Grün.", do: ["Folie 10–12 nacheinander. Niemand geht an die Tafel.", "Sprachreflexion (Folie 12, Notizen): Warum fällt der Fehler beim Sprechen nicht auf, beim Schreiben aber schon? Beim Schreiben fehlen Pausen und Betonung – Komma und Verbstellung zeigen, wo der Nebensatz beginnt und endet."] }],
]));

doc.push(avivaPhase("A", "Auswerten: Zielcheck", 6));
doc.push(...steps([
  ["luecken", { say: "Zielcheck. Jetzt arbeitet ihr allein im Heft. Drei Minuten. Fertig? Stift hinlegen.", sol: "1 Kreuzreim · Dialog · 2 Schaya versteckt Paddy im Bett, weil der Schotte ihn nicht finden soll. · 3 Der Schotte rennt sofort weg, als er das riesige Baby sieht.", do: ["Folie 13."] }],
  ["kontrolle", { say: "Vergleicht mit der Lösung auf der Folie und verbessert mit Grün. Haben wir das Ziel erreicht?", do: ["Folie 14. Herumgehen und notieren, wer noch Fehler hat – in der Wiederholung der nächsten Stunde gezielt aufgreifen."] }],
]));

// ---------- Reflexion ----------
const bul = (head, txt) => p([S("• ", { bold: true, color: C.hv }), ...(head ? [S(head + " ", { bold: true })] : []), S(txt)], { after: 40, indent: { left: 240, hanging: 240 } });
const refHead = (txt, col) => p(t(txt, { font: "Cambria", size: 22, bold: true, color: col }), { before: 140, after: 50, keepNext: true });
doc.push(new Paragraph({ pageBreakBefore: true, children: [t("FÜR DIE LEHRKRAFT · REFLEXION VOR UND NACH DER STUNDE", { size: 18, color: MUTED, bold: true })], spacing: { after: 20 } }), h3("Umsetzung der Erkenntnisse zur Unterrichtsverbesserung"));
doc.push(refHead("1  AVIVA – Struktur und Transparenz", C.hv));
doc.push(bul("Ankommen:", "Ziel als „Ich kann …“ und Ablauf auf Folie und Seitentafel, * Nachricht 5 markiert; Rückgriff im Zielcheck („Ziel erreicht?“)."));
doc.push(bul("Vorwissen aktivieren:", "Echte Wiederholung der ganzen Woche: Kreuzworträtsel hierarchisch vom Allgemeinen ins Detail (Rückmeldung 5.11.), die Anschlussaufgabe ist als gemeinsamer nächster Schritt mit Schreibzeilen gekennzeichnet (Rückmeldung 5.11.)."));
doc.push(bul("Informieren:", "Kein neuer Input, sondern die Regel wird am eigenen weil-Satz der Klasse wieder entdeckt; Merksatz nur bei Bedarf (nicht abschreiben)."));
doc.push(bul("Verarbeiten:", "Fehleranalyse statt Lückenfüllen (Abwechslung, Rückmeldung 08.10.); Aufdecken Schritt für Schritt, niemand an der Tafel (Rückmeldung 31.10.)."));
doc.push(bul("Auswerten:", "Zielcheck mit drei kurzen Aufgaben genau zum Ziel, Kontrolle mit Grün, Fehler werden in der nächsten Stunde aufgegriffen."));
doc.push(bul("Eindeutige Aufgaben:", "Nach der Rückmeldung zum Ur-Ei-Schnipsel nennt jede Aufgabe Nummer, Zahl der Antworten und Antwortformat (13 Wörter der Reihe nach; genau ein Fehler pro Nachricht, K oder V)."));
doc.push(refHead("2  Direkte Instruktion (DI)", NAVY));
doc.push(bul("Ziel und Relevanz:", "Ein überprüfbares Ziel; Relevanz: Merkmale braucht man für die Analyse (Bewertung 1), Nebensätze für jede schriftliche Begründung."));
doc.push(bul("Ich – Wir – Ihr:", "Lehrkraft macht die Begründung vor (Ich), die Klasse ergänzt gemeinsam (Wir), Paare verbessern den Chat, im Zielcheck arbeitet jede/r allein (Ihr)."));
doc.push(bul("Rückmeldung:", "Lösungsfolie nach Ü1, Aufdeck-Folien nach Ü2, Zielcheck mit Lösung; typische Hürden vorab benannt (UMARMENDER, fehlendes Komma, zwei Verben am Ende)."));
doc.push(bul("Klarheit:", "Feste wörtliche Ansagen bei jedem Wechsel, gleiches Übungsformat, festes Stopp-Signal beim Würfeln."));
doc.push(refHead("3  Differenzierung", "2A8C8C"));
doc.push(bul("", "Leseband: Wer unsicher liest, wird bei „Wir“ vom Chor getragen; die Ein-Satz-Zusammenfassung sichert das Verstehen."));
doc.push(bul("", "Ü2: Partnerarbeit als Stütze, * Nachricht 5 für Schnelle; im Zielcheck allein."));
doc.push(bul("", "S-PL: Handreichung mit Ziel, Ablauf und Aufträgen auf Polnisch, Hinweise zu den Rätsel-Hinweisen und zu schwierigen Wörtern im Chat; Fachbegriffe bleiben deutsch, keine Lösungen."));
doc.push(refHead("4  Beobachtungsschwerpunkte", MUTED));
doc.push(bul("", "Ist beim Würfeln die ganze Gruppe aktiv (Mitlesen, Zusammenfassen)?"));
doc.push(bul("", "Sitzen die Fachbegriffe, besonders Epik vs. Dramatik (Erzähler ja/nein)?"));
doc.push(bul("", "Erkennen die Schüler das fehlende Komma, wenn der Satz „richtig klingt“ (Nr. 1, 5)? Gelingt die Verbstellung mit zwei Verben (Nr. 1, 3)?"));
doc.push(bul("", "Hospitation: Rückmeldung der Beobachtenden zu Transparenz, Aktivierung aller und Zeitmanagement im Reflexionslog festhalten."));

const pageProps = { page: { size: { width: 11906, height: 16838 }, margin: { top: 900, bottom: 800, left: 850, right: 850, header: 400, footer: 400 } } };
const styles = { default: { document: { run: { font: "Calibri", size: 22 } } } };
Packer.toBuffer(new Document({ styles, sections: [{ properties: pageProps, headers: { default: header("Balladen · Mo 9.11.") }, footers: { default: footer("Tafelskript") }, children: doc }] })).then((buf) => fs.writeFileSync(__dirname + "/../Tafelskript_Mo_09-11_Leseband_Kreuzwortraetsel_Subjunktionen.docx", buf));
