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



// Tafelskript Do 5.11. (Einzelstunde): Wiederholung · Reime erkennen – AVIVA, mit Präsentation
const avivaPhase = (k, title, min) => p([t(`${k}  `, { font: "Cambria", size: 28, bold: true, color: C.hv }), t(title, { font: "Cambria", size: 28, bold: true, color: NAVY }), t(`   ${min} Min.`, { size: 20, bold: true, color: "C77C12" })], { before: 200, after: 100, keepNext: true });
const ZIEL = "Ich kann Reime erkennen und das Reimschema einer Strophe bestimmen.";
const RA = "2A8C8C", RB = "C77C12";
const doc = [];
doc.push(kicker("Tafelskript für die Lehrkraft · Do 5.11. · Einzelstunde (45 Min.) · AVIVA · mit Präsentation (17 Folien)"), h1("Paddy Fingal – Reime erkennen"));
doc.push(box([p([S("Bezug: ", { bold: true }), S("Am Mittwoch hat die Klasse den Aufbau der Ballade gelernt (Strophe, Vers, Belegen), Aussagen geprüft und Nebensätze mit Subjunktion geübt. Die Chat-Übung wurde dafür auf heute verschoben: Sie ist jetzt die Wiederholung des Sprachtrainings, dazu eine kurze Belegübung. Danach geht es um die Lyrik in der Ballade: Reim und Reimschema, kleinschrittig an Strophe 1 entdeckt. Der Reim-Schnipsel hat jetzt für jede Sportart der Klasse einen Text (Gewichtheben, Judo, Handball). Das Ur-Ei folgt am Freitag. Das Nacherzählen vom Mittwoch dient heute als Puffer.")], { after: 0 })], LIGHT2));
doc.push(p([S("Lernziel: ", { bold: true }), S(ZIEL, { italics: true })], { before: 120, after: 40 }));
doc.push(p([S("Antwort auf „Was habt ihr heute gelernt?“: ", { bold: true }), S("„Wörter reimen sich, wenn sie ab dem letzten betonten Vokal gleich klingen. Mit Buchstaben bestimme ich das Reimschema: Paarreim aabb, Kreuzreim abab, umarmender Reim abba. Paddy Fingal hat Paarreime und einen Kehrreim.“ Überprüft wird das im Zielcheck.")], { after: 80 }));
doc.push(board("Seitentafel links · vor der Stunde anschreiben, bleibt stehen", [
  p([t("Ziel: ", { bold: true, size: 21, color: NAVY }), t(ZIEL, { size: 21, italics: true })], { after: 60 }),
  ...["Wiederholung: Ü1 Fehler im Chat · Ü2 Belegen", "Reime in Strophe 1 entdecken", "Ü3 Reime erkennen (Schnipsel)", "Zielcheck", "* Nacherzählen mit Stichwörtern"].map((a, i) => p(t(`${i + 1}  ${a}`, { size: 20 }), { after: 10 })),
]));
doc.push(p(t(""), { after: 40 }));
doc.push(h3("Ablauf"));
doc.push(grid([800, 900, 1200, 3300, 4006], ["Zeit", "AVIVA", "Folie", "Phase", "Alle Schüler …"], [
  ["2′", "A", "1–2", "Ankommen: Ziel und Ablauf", "lesen Ziel und Ablauf, schreiben die Überschrift"],
  ["5′", "V", "3–4", "Ü1 Fehler im Chat (Wiederholung Sprachtraining)", "korrigieren allein, verbessern mit der Lösungsfolie"],
  ["4′", "V", "5–6", "Ü2 Belegen (Wiederholung)", "suchen zwei Stellen mit Str., V., Zitat; Lösungsfolie"],
  ["9′", "I", "7–9", "Reime in Strophe 1 entdecken, Merkkasten", "lesen im Chor, finden Reimwörter, kleben den Schnipsel ein"],
  ["12′", "V", "10–14", "Ü3 Reime erkennen (Schnipsel)", "bestimmen allein vier Reimschemata, prüfen gemeinsam (Aufdeck-Folien)"],
  ["6′", "A", "15–16", "Auswerten: Zielcheck", "bestimmen das Reimschema von Str. 6 und eine Reimart, kontrollieren mit der Lösungsfolie"],
  ["* 7′", "V", "17", "Puffer: Nacherzählen mit Stichwörtern", "erzählen sich im Tandem nach"],
], { bold: [true, true, false, true, false], colors: [C.hv, C.hv, MUTED, NAVY, null], h: 320 }));
doc.push(p([S("Zeit: ", { bold: true, color: C.hv }), S("38 Min. Pflicht, 7 Min. Puffer (Nacherzählen). Falls das Nacherzählen am Mittwoch schon geschafft wurde: Puffer für eine gründlichere Besprechung von Ü3.")], { before: 80, after: 60 }));
doc.push(h3("So sieht das Heft am Ende aus"));
doc.push(table([2300, 7906], [
  ["Überschrift", "Reime erkennen · Datum rechts"],
  ["Wiederholung", "Ü1: vier verbesserte Chat-Sätze · Ü2: zwei Belege (Str., V., Zitat)"],
  ["Reime", "Strophe 1: Versenden mit a a b b x · Schnipsel „Reime erkennen“ eingeklebt und ausgefüllt, mit Grün kontrolliert"],
  ["Zielcheck", "Str. 6: Reimschema, Reimart, Kehrreim · Reimart zu „Ball – Tor – Fall – vor“"],
].map(([a, b]) => row([tc(a, 2300, { bold: true, fill: LIGHT }), tc(b, 7906)]))));
doc.push(p([S("Material: ", { bold: true }), S("Präsentation `Donnerstag_05-11_Deutsch_7A.pptx` (Lösungen in den Notizen und auf Aufdeck-Folien) · Lehrbuch S. 136 f. · Schnipsel „Reime erkennen“ (2 pro A4) · Handreichung Polnisch für S-PL.")], { before: 80, after: 60 }));

doc.push(avivaPhase("A", "Ankommen: Ziel und Ablauf", 2));
doc.push(...steps([["zuhoeren", { say: "Stifte liegen. Ihr hört nur zu. Am Mittwoch habt ihr geprüft, was in der Ballade passiert. Heute geht es darum, wie sie klingt: um Reime. Ziel und Ablauf stehen auf der Folie und links an der Tafel. Jetzt schreibt ihr die Überschrift ab: Reime erkennen." }]]));

doc.push(avivaPhase("V", "Vorwissen aktivieren: Ü1 Fehler im Chat", 5));
doc.push(...steps([
  ["luecken", { say: "Wiederholung vom Mittwoch. Jetzt arbeitet ihr allein. Im Chat stecken Fehler. Schreibt die Sätze richtig ins Heft und notiert: Verbstellung oder Komma? Vier Minuten. Fertig? Stift hinlegen.", sol: "1 Komme später, weil ich noch duschen muss. · 2 Trainer hat gesagt, dass wir morgen frei haben! · 3 Weißt du, ob der Neue auch kommt? · 4 Ich freue mich, weil wir frei haben. (Komma)" }],
  ["kontrolle", { say: "Vergleicht mit der Lösung auf der Folie und verbessert mit Grün.", do: ["Folie 4. Kurz fragen: Was ist die Regel? – Nach dass, weil, ob steht das finite Verb am Ende, davor ein Komma. Beim Sprechen hört man oft „weil ich muss …“ – geschrieben ist das falsch.", "Gezielt die Schüler drannehmen, die im Zielcheck am Mittwoch noch Fehler hatten."] }],
]));
doc.push(avivaPhase("V", "Vorwissen aktivieren: Ü2 Belegen", 4));
doc.push(...steps([
  ["luecken", { say: "Ü2: Wo steht das? Gebt Strophe, Vers und Zitat an. Drei Minuten.", sol: "1 Str. 5, V. 21: „Ins Bette Paddy Fingal kroch“ · 2 Str. 7, V. 31: „Tut leise, Fremder, und tretet sacht“" }],
  ["kontrolle", { say: "Vergleicht mit der Lösung und verbessert mit Grün.", do: ["Folie 6. Auf Anführungszeichen und Strophe UND Vers achten."] }],
]));

doc.push(avivaPhase("I", "Informieren: Reime in Strophe 1 entdecken", 9));
doc.push(board("Mitteltafel · optional zusätzlich zur Folie (Versenden mit Buchstaben)", [
  bl([N("… von Riesen"), N("art", { bold: true, color: RA }), N("     a", { bold: true, color: RA })]),
  bl([N("… schwarzer B"), N("art", { bold: true, color: RA }), N("     a", { bold: true, color: RA })]),
  bl([N("… wie Tonnen gr"), N("oß", { bold: true, color: RB }), N("     b", { bold: true, color: RB })]),
  bl([N("… als Wurfgesch"), N("oss", { bold: true, color: RB }), N("     b", { bold: true, color: RB })]),
  bl([N("… so ein Kerl war Paddy Fingal!     "), N("x → Kehrreim", { bold: true })]),
  bl([N("aabb = Paarreim", { bold: true, color: NAVY })], { before: 40 }),
]));
doc.push(p(t(""), { after: 40 }));
doc.push(...steps([
  ["gemeinsam", { say: "Lehrbuch S. 136. Wir lesen Strophe 1 zusammen im Chor. … Hört auf die Enden der Verse. Welche Wörter klingen gleich – und ab welchem Laut?", do: ["Folie 7. Methodenkarte „Chorisches Lesen“. Nachfragen: Ab wo klingen Bart und Riesenart gleich? – ab dem betonten Vokal: -art."] }],
  ["gemeinsam", { say: "Was sich reimt, bekommt denselben Buchstaben. Welchen Buchstaben bekommt Vers 3? Und Vers 5?", sol: "aabb = Paarreim · Vers 5 reimt sich nicht (x), kehrt aber in jeder Strophe wieder: Kehrreim (Refrain).", do: ["Folie 8 erst nach dem Gespräch zeigen."] }],
  ["zuhoeren", { say: "Stifte liegen. Ihr hört nur zu. Ich teile den Schnipsel aus. Klebt ihn ein. Wir lesen den Kasten oben zusammen.", do: ["Folie 9. Kontrollfragen: Wie heißt das Muster, wenn sich Vers 1 und 3 reimen? (abab, Kreuzreim) – und Vers 1 und 4? (abba, umarmender Reim)."] }],
]));

doc.push(avivaPhase("V", "Verarbeiten: Ü3 Reime erkennen (Schnipsel)", 12));
doc.push(...steps([
  ["luecken", { say: "Jetzt arbeitet ihr allein. Lest die Versenden halblaut, schreibt hinter jeden Vers den Buchstaben und nennt die Reimart. Sechs Minuten. Fertig? Stift hinlegen.", sol: "1 Schreck a, Versteck a, daher b, sehr b, *Paddy Fingal x → Paarreim + Kehrreim · 2 leer a, da b, schwer a, Ja b → Kreuzreim · 3 zu a, Kragen b, wagen b, Ruh a → umarmender Reim · 4 frei a, vorbei a, rein b, schrein b → Paarreim", do: ["Herumgehen. Typische Hürde: umarmender Reim (Nr. 3) – Tipp: Vers 1 und 4 vergleichen."] }],
  ["gemeinsam", { say: "Wir prüfen gemeinsam, Text für Text. Wer drankommt, nennt die Buchstaben und die Reimart. Dann decke ich auf. Alle verbessern mit Grün.", do: ["Folie 11 bis 14. Jeweils einen anderen Schüler drannehmen, möglichst einen aus der jeweiligen Sportart. Niemand geht an die Tafel."] }],
]));

doc.push(avivaPhase("A", "Auswerten: Zielcheck", 6));
doc.push(...steps([
  ["luecken", { say: "Zielcheck. Jetzt arbeitet ihr allein im Heft. a: Bestimmt das Reimschema von Strophe 6 und nennt die Reimart. Welcher Vers ist der Kehrreim? b: Welche Reimart ist Ball – Tor – Fall – vor? Vier Minuten.", sol: "a herein – Scheitelbein (a a), Schweiß – Kreis (b b), „Wo steckt der Paddy Fingal?“ (x) → Paarreim, V. 30 ist der Kehrreim · b abab → Kreuzreim" }],
  ["kontrolle", { say: "Vergleicht mit der Lösung auf der Folie und verbessert mit Grün.", do: ["Folie 16. Herumgehen und schauen, wer noch Fehler hat – diese Fälle am Freitag in der Wiederholung aufgreifen. Ausblick: Am Freitag schauen wir, warum die Ballade mehr ist als ein Gedicht (Ur-Ei)."] }],
]));
doc.push(avivaPhase("V", "* Puffer: Nacherzählen mit Stichwörtern", 7));
doc.push(...steps([["partner", { say: "Erzählt die Ballade im Tandem nach. Auf der Folie stehen der Einstieg und acht Abschnitte mit Stichwörtern. A erzählt Abschnitt 1 bis 4, B Abschnitt 5 bis 8. Verknüpft mit zuerst, dann, als, weil, schließlich. Wer zuhört, hakt die Checkliste ab.", do: ["Folie 17. Nur, wenn Zeit bleibt oder falls Ü3 am Mittwoch nicht geschafft wurde."] }]]));

// ---------- Reflexion ----------
const bul = (head, txt) => p([S("• ", { bold: true, color: C.hv }), ...(head ? [S(head + " ", { bold: true })] : []), S(txt)], { after: 40, indent: { left: 240, hanging: 240 } });
const refHead = (txt, col) => p(t(txt, { font: "Cambria", size: 22, bold: true, color: col }), { before: 140, after: 50, keepNext: true });
doc.push(p(new PageBreak()));
doc.push(kicker("Für die Lehrkraft · Reflexion vor und nach der Stunde"), h3("Umsetzung der Erkenntnisse zur Unterrichtsverbesserung"));
doc.push(refHead("1  AVIVA – Struktur und Transparenz", C.hv));
doc.push(bul("Ankommen:", "Ziel und Ablauf auf Folie und Seitentafel, Rückgriff im Zielcheck."));
doc.push(bul("Vorwissen aktivieren:", "Echte Wiederholung beider Teile vom Mittwoch: Sprachtraining (Chat-Fehler, die verschobene Übung) und Belegen."));
doc.push(bul("Informieren:", "Kleinschrittig: Reime erst an Strophe 1 hören und entdecken, dann der Merkkasten auf dem Schnipsel (nicht abschreiben)."));
doc.push(bul("Verarbeiten und Auswerten:", "Alle arbeiten allein, Auswertung über Lösungs- und Aufdeck-Folien, niemand muss an die Tafel. Sicherung als Zielcheck mit neuer Strophe statt „Heute habe ich gelernt“."));
doc.push(bul("Zeit:", "Bewusst weniger Inhalt (Ur-Ei auf Freitag verlegt), 7 Min. Puffer mit dem Nacherzählen."));
doc.push(refHead("2  Direkte Instruktion (DI)", NAVY));
doc.push(bul("Ziel und Relevanz:", "Ein überprüfbares Ziel; „Wofür?“: Reim und Reimschema braucht man für die Analyse (Bewertung 1)."));
doc.push(bul("Ich – Wir – Ihr:", "Gemeinsam an Strophe 1 (Wir), Kontrollfragen zu abab/abba, dann vier Texte allein (Ihr)."));
doc.push(bul("Rückmeldung:", "Lösungsfolien nach Ü1/Ü2, Aufdeck-Folien nach Ü3, Zielcheck mit Lösung; typische Hürde benannt: umarmender Reim."));
doc.push(refHead("3  Differenzierung", "2A8C8C"));
doc.push(bul("", "Übungstexte für jede Sportart der Klasse (Gewichtheben, Judo, Handball); * Kehrreim in Nr. 1; Puffer Nacherzählen mit Stichwörtern."));
doc.push(bul("", "S-PL: Handreichung mit Ziel, Ablauf, Aufträgen und den Reimbegriffen auf Polnisch, dazu die Übungstexte sinngemäß übersetzt."));
doc.push(refHead("4  Beobachtungsschwerpunkte", MUTED));
doc.push(bul("", "Sitzt die Regel zu den Nebensätzen nach einem Tag (Ü1)? Sitzt das Belegen (Ü2)?"));
doc.push(bul("", "Hören die Schüler den Reim ab dem betonten Vokal, oder schauen sie nur auf die letzten Buchstaben?"));
doc.push(bul("", "Wird der umarmende Reim erkannt? Wie lange dauert Ü3 wirklich?"));

const pageProps = { page: { size: { width: 11906, height: 16838 }, margin: { top: 900, bottom: 800, left: 850, right: 850, header: 400, footer: 400 } } };
const styles = { default: { document: { run: { font: "Calibri", size: 22 } } } };
Packer.toBuffer(new Document({ styles, sections: [{ properties: pageProps, headers: { default: header("Balladen · Do 5.11.") }, footers: { default: footer("Tafelskript") }, children: doc }] })).then((buf) => fs.writeFileSync("Tafelskript_Do_05-11_Reime_erkennen.docx", buf));
