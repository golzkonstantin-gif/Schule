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
// Tafelskript „Das Plusquamperfekt“ (45 Min.)
// ================================================================
const merkTafel = (nr, title, children) => board(`Merksatz ${nr} – abschreiben, rot umrahmen`, [
  p([t(title, { bold: true, color: C.hv, font: "Cambria", size: 22 })], { after: 40 }),
  ...children,
]);
const sl = (runs, o = {}) => p(runs, Object.assign({ after: 30 }, o));
const n19 = (s, o = {}) => N(s, Object.assign({ size: 20 }, o));
const r19 = (s) => t(s, { bold: true, color: C.hv, font: "Cambria", size: 20 });
const y19 = (s) => t(s, { bold: true, color: "C98A1E", font: "Cambria", size: 20 });

const doc = [];
doc.push(kicker("Tafelskript für die Lehrkraft · Einzelstunde (45 Min.)"), h1("Das Plusquamperfekt"));
doc.push(box([
  p([S("Anknüpfung an die letzte Stunde: ", { bold: true }), S("Das Präteritum wurde geübt (Konjugation kämpfen/werfen, Sportverben, Vereinszeitung). Die Schüler sind inzwischen genervt davon, Tabellen an der Tafel auszufüllen. Deshalb heute: keine Tafel-Tabelle. Das Plusquamperfekt wird mit Satzstreifen am Zeitstrahl entdeckt, die Konjugation von hatte/war kommt als fertiger Schnipsel, bei der Kreide-Kette schreibt jeder ganze Sätze. Am Ende beginnt der eigene Wettkampfbericht (Ü11, Schreibrahmen ist gedruckt).")], { after: 0 }),
], LIGHT2));
doc.push(p([S("Ziel: ", { bold: true }), S("Die Klasse versteht das Plusquamperfekt als Vorgeschichte („was vorher passiert war“), bildet es mit hatte/war + Partizip II und wendet es im eigenen Wettkampfbericht an.")], { before: 120, after: 60 }));
doc.push(p([S("Markieren wie bisher: ", { bold: true }), t("Hilfsverb rot", { bold: true, color: C.hv, size: 21 }), S(" · "), t("Partizip II gelb", { bold: true, color: "C98A1E", size: 21 }), S(" · Präteritum-Verb unterstreichen · keine Klammer-Bögen")], { after: 100 }));

doc.push(h3("Ablauf"));
doc.push(grid([900, 3500, 5806], ["Zeit", "Phase", "Kern"], [
  ["5′", "0 Reaktivierung: Fehler finden", "drei falsche Präteritum-Sätze – Kreide-Kette verbessert ganze Sätze"],
  ["8′", "1 Satzstreifen am Zeitstrahl", "Wettkampftag ordnen – die hatte-Sätze gehören ganz nach links"],
  ["5′", "2 hatte und war (Schnipsel)", "fertige Tabelle einkleben, Stolperstellen markieren"],
  ["4′", "3 Merksatz 9", "Das Plusquamperfekt abschreiben"],
  ["8′", "4 Ü10 Was war vorher? (Schnipsel)", "eintragen, Vergleich: ganze Sätze an die Tafel"],
  ["13′", "5 Ü11 Mein erster Wettkampf", "Bericht beginnen (Schreibrahmen), Rest als Hausaufgabe"],
], { bold: [true, true, false], colors: [C.hv, NAVY, null], h: 360 }));
doc.push(p([S("Puffer: ", { bold: true, color: C.hv }), S("Phase 5 ist bewusst offen: Was nicht fertig wird, ist Hausaufgabe. Bleibt Zeit, lesen zwei Schüler ihren Anfang vor.")], { before: 80, after: 60 }));

doc.push(h3("So sieht das Heft am Ende aus"));
doc.push(table([2300, 7906], [
  ["Überschrift", "Das Plusquamperfekt · Datum rechts"],
  ["Schnipsel", "haben und sein im Präteritum (eingeklebt, Stolperstellen markiert)"],
  ["Merksatz 9", "Das Plusquamperfekt + kleiner Zeitstrahl"],
  ["Ü10", "Schnipsel eingeklebt: Was war vorher?"],
  ["Ü11", "Anfang des eigenen Wettkampfberichts (Rest: Hausaufgabe)"],
].map(([a, b]) => row([tc(a, 2300, { bold: true, color: a.startsWith("Merk") ? C.hv : a === "Überschrift" ? NAVY : "C77C12", fill: LIGHT, size: 20 }), tc(b, 7906, { size: 20 })], 340))));
doc.push(p([S("Material: ", { bold: true }), S("Druckvorlage (3 Teile): Seite 1 hatte/war-Schnipsel (8 pro A4), Seite 2 Ü10 (4 pro A4 – falls von der letzten Stunde noch vorhanden, diese nehmen), Seiten 3–4 sechs große Satzstreifen für die Tafel (einmal drucken, schneiden, Magnete). Schreibrahmen Ü11 (bereits gedruckt). Kreide Rot, Gelb · Schüler: Heft, Kleber, Buntstifte Rot, Gelb, Grün")], { before: 100, after: 40 }));
doc.push(p([S("Tipp: ", { bold: true }), S("Merksatz 9 vorab auf die Seitentafel schreiben und zuklappen. Den Zeitstrahl (Pfeil mit „vorher – Vergangenheit – heute“) vorab an die Tafel zeichnen, die Satzstreifen gemischt daneben hängen.")], { after: 0 }));

// ---------- Phase 0 ----------
doc.push(phase("0", "Reaktivierung: Fehler finden", 5));
doc.push(...steps([
  ["tafel", { say: "Heft auf, Datum, Überschrift: Das Plusquamperfekt. Dann schaut an die Tafel: In jedem Satz steckt ein Fehler im Präteritum. Kreide-Kette: Wer die Kreide hat, schreibt den ganzen Satz richtig darunter. Alle prüfen mit.", board: board("Mitte", [
    bl("1  Ich werfte den Ball ins Tor."), bl("2  Er kämpftet sehr gut."), bl("3  Wir gewinnten das Spiel knapp.", { after: 0 }),
  ]), sol: "1 Ich warf den Ball ins Tor. · 2 Er kämpfte sehr gut. · 3 Wir gewannen das Spiel knapp.", do: ["Nach jedem Satz kurz begründen lassen: „werfen ist unregelmäßig – neuer Stamm, kein -te“ · „er hat keine Personalendung: kämpfte“ · „gewinnen ist unregelmäßig: gewann, wir gewannen“."] }],
]));

// ---------- Phase 1 ----------
const STR = [
  ["F", "Ich hatte wochenlang hart trainiert.", true],
  ["B", "Ich hatte die Nacht davor kaum geschlafen.", true],
  ["A", "Am Morgen war ich sehr nervös.", false],
  ["C", "In der Halle wärmte ich mich auf.", false],
  ["D", "Dann begann mein erster Kampf.", false],
  ["E", "Am Ende gewann ich knapp.", false],
];
doc.push(phase("1", "Satzstreifen am Zeitstrahl", 8));
doc.push(...steps([
  ["tafel", { say: "Stifte liegen. An der Tafel hängen sechs Sätze aus einem Wettkampftag – durcheinander. Wer drankommt, hängt einen Streifen an die richtige Stelle am Zeitstrahl. Alle anderen sagen: Stimmt das?", board: board("Mitte · Satzstreifen (gemischt aufhängen)", STR.map(([k, s2, pq], i) => bl([N(k + "  ", { bold: true, color: MUTED }), N(s2)], { after: i === STR.length - 1 ? 0 : 30 }))),
    sol: "Richtige Reihenfolge: F, B (vorher) → A, C, D, E (der Wettkampftag). F und B können auch getauscht werden – beides war vorher.",
    do: ["Zeitstrahl an der Tafel: links „vorher“, Mitte „Wettkampftag (Vergangenheit)“, rechts „heute“.", "Erst ordnen lassen, dann fragen: „Welche zwei Streifen stehen ganz links? Was ist an ihren Verben anders?“"] }],
  ["gemeinsam", { say: "Meldet euch. Ihr schreibt noch nichts auf. Die beiden linken Sätze: Welche Verben stehen dort – und was davon kennt ihr schon?", do: [
    "Ziel: hatte (Hilfsverb haben im Präteritum) + Partizip II (trainiert, geschlafen). Am Streifen hatte rot und das Partizip gelb markieren.",
    "Begriff: Plusquamperfekt – die Vorgeschichte, also das, was vor der Vergangenheit passiert war.",
    "Kontrast mündlich: Ich habe trainiert (Perfekt) – Ich hatte trainiert (Plusquamperfekt). Der Unterschied steckt nur im Hilfsverb.",
  ] }],
]));

// ---------- Phase 2 ----------
doc.push(phase("2", "hatte und war (Schnipsel)", 5));
doc.push(...steps([
  ["luecken", { say: "Ihr bekommt einen Schnipsel mit haben und sein im Präteritum. Klebt ihn ein und markiert die drei Stolperstellen, die auf dem Schnipsel stehen. Drei Minuten.", do: ["Kurze mündliche Kontrolle: „Welche Formen sind gleich?“ (ich hatte – er hatte; ich war – er war) · „Wo fehlt das e?“ (ihr wart).", "Keine Tafel-Tabelle – die Formen stehen fertig auf dem Schnipsel."] }],
]));

// ---------- Phase 3 ----------
doc.push(phase("3", "Merksatz 9: Das Plusquamperfekt", 4));
doc.push(...steps([
  ["abschreiben", { say: "Jetzt schreibt ihr von der Tafel ab: Merksatz 9 mit dem Beispiel und dem kleinen Zeitstrahl. Rot umrahmen.", board: merkTafel(9, "Das Plusquamperfekt", [
    p([S("Das "), S("Plusquamperfekt", { bold: true }), S(" bildet man mit "), S("haben oder sein im Präteritum", { bold: true }), S(" (hatte, war) und dem "), S("Partizip II", { bold: true }), S(".")], { after: 60 }),
    p([S("Beispiel: ", { bold: true, color: MUTED }), n19("Ich "), r19("hatte"), n19(" wochenlang "), y19("trainiert"), n19(".  ·  Wir "), r19("waren"), n19(" früh "), y19("losgefahren"), n19(".")], { after: 60 }),
    p([S("Zeitstrahl: ", { bold: true, color: MUTED }), S("hatte trainiert (vorher)  →  war nervös (Vergangenheit)  →  heute")], { after: 60 }),
    p([S("Wofür? ", { bold: true }), S("Damit erzählt man, was "), S("vorher", { bold: true }), S(" passiert war – die Vorgeschichte zu einem Ereignis in der Vergangenheit. "), S("Test: ", { bold: true }), S("Steht hatte oder war mit einem Partizip II? Dann ist es Plusquamperfekt. (Steht habe oder bin, ist es Perfekt.)")], { after: 0 }),
  ]) }],
]));

// ---------- Phase 4 ----------
doc.push(phase("4", "Ü10 Was war vorher? (Schnipsel)", 8));
doc.push(...steps([
  ["luecken", { say: "Jetzt arbeitet ihr allein: Schnipsel Ü10. Klebt ihn ein und ergänzt im zweiten Satz das Plusquamperfekt. Vier Minuten.", board: auftrag("Ü10 Was war vorher?", ["Klebe den Schnipsel Ü10 ins Heft.", "Ergänze das Plusquamperfekt mit dem Verb in Klammern.", "Markiere hatte/war rot und das Partizip II gelb."], [p(S("Die Sätze stehen auf dem Schnipsel.", { size: 20, italics: true }), { after: 0 })], "4 Minuten · allein · leise"),
    sol: "1 hatte trainiert · 2 hatten gemacht · 3 hatte aufgewärmt · 4 waren losgefahren · 5 hatte geschafft" }],
  ["tafel", { say: "Kreide-Kette: Wer die Kreide hat, schreibt den ganzen zweiten Satz an die Tafel – nicht nur die Lücke – und markiert hatte/war und das Partizip. Alle verbessern mit Grün.", do: ["Ganze Sätze statt einzelner Wörter: weniger „Ausfüllen“, mehr Schreiben.", "Bei Nr. 4 nachfragen: Warum war und nicht hatte? → losfahren ist Bewegung, wie beim Perfekt (sind losgefahren)."] }],
]));

// ---------- Phase 5 ----------
doc.push(phase("5", "Ü11 Mein erster Wettkampf", 13));
doc.push(...steps([
  ["luecken", { say: "Jetzt schreibt ihr euren eigenen Bericht für die Vereinszeitung – ins Heft, als Ü11. Nehmt den Schreibrahmen dazu. Beginnt mit der Vorgeschichte im Plusquamperfekt. Was ihr heute nicht schafft, schreibt ihr zu Hause fertig.", board: auftrag("Ü11 Mein erster Wettkampf", ["Schreibe 5–6 Sätze über deinen ersten Wettkampf (oder einen, an den du dich gut erinnerst) ins Heft.", "Beginne mit der Vorgeschichte im Plusquamperfekt: „Vorher hatte ich …“", "Erzähle den Wettkampf im Präteritum."], [p(S("Hilfe: Satzanfänge, Bausteine und Wörter für deine Sportart auf dem Schreibrahmen.", { size: 20, italics: true }), { after: 0 })], "10 Minuten · allein · leise · Rest: Hausaufgabe"),
    do: ["Herumgehen: Typische Fehler sind Wechsel ins Perfekt („Dann habe ich gewonnen“), regelmäßig gebildete Präteritumformen („werfte“), habe statt hatte in der Vorgeschichte.", "Wenn Zeit bleibt: zwei Schüler lesen ihren Anfang vor, die Klasse hört auf die hatte-Sätze."] }],
]));
doc.push(p([S("Ausblick: ", { bold: true }), S("Nächste Stunde: Partner-Check der fertigen Berichte und Reflexion „Wann nehme ich welche Zeitform?“. Danach: Futur II, Hilfsverb oder Vollverb, Modalverben – und zum Schluss der Tempus-Einheit die Satzklammer.")], { before: 120, after: 0 }));

const pageProps = { page: { size: { width: 11906, height: 16838 }, margin: { top: 900, bottom: 800, left: 850, right: 850, header: 400, footer: 400 } } };
const styles = { default: { document: { run: { font: "Calibri", size: 22 } } } };
const dScript = new Document({ styles, sections: [{ properties: pageProps, headers: { default: header("Das Plusquamperfekt") }, footers: { default: footer("Tafelskript") }, children: doc }] });
Packer.toBuffer(dScript).then((b) => fs.writeFileSync("Tafelskript_Plusquamperfekt.docx", b));
