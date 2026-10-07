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
// Tafelskript „Präteritum üben, Plusquamperfekt“ (45 Min.)
// ================================================================
const merkTafel = (nr, title, children) => board(`Merksatz ${nr} – abschreiben, rot umrahmen`, [
  p([t(title, { bold: true, color: C.hv, font: "Cambria", size: 22 })], { after: 40 }),
  ...children,
]);
const sl = (runs, o = {}) => p(runs, Object.assign({ after: 30 }, o));
const n19 = (s, o = {}) => N(s, Object.assign({ size: 20 }, o));
const r19 = (s) => t(s, { bold: true, color: C.hv, font: "Cambria", size: 20 });
const y19 = (s) => t(s, { bold: true, color: "C98A1E", font: "Cambria", size: 20 });
const u19 = (s) => t(s, { bold: true, font: "Cambria", size: 20, u: true });

const doc = [];
doc.push(kicker("Tafelskript für die Lehrkraft · Einzelstunde (45 Min.)"), h1("Präteritum üben, Plusquamperfekt"));
doc.push(box([
  p([S("Anknüpfung an die letzte Stunde: ", { bold: true }), S("Die Klasse kam bis zum Präteritum (Merksatz 8). Das war für die Schüler ungewohnt – Präteritum hat kein Hilfsverb, und die unregelmäßigen Formen (warf, hob, traf) sind neu. Deshalb zuerst Übung auf Schnipseln (spart das Abschreiben), danach erst das Plusquamperfekt. Der Schreibauftrag Ü11 „Mein erster Wettkampf“ (Schreibrahmen ist schon gedruckt) folgt in der nächsten Stunde.")], { after: 0 }),
], LIGHT2));
doc.push(p([S("Ziel: ", { bold: true }), S("Die Klasse bildet Präteritumformen sicherer – auch unregelmäßige – und kann das Plusquamperfekt (hatte/war + Partizip II) als „Vorgeschichte“ bilden und erkennen.")], { before: 120, after: 60 }));
doc.push(p([S("Markieren wie bisher: ", { bold: true }), t("Hilfsverb rot", { bold: true, color: C.hv, size: 21 }), S(" · "), t("Partizip II gelb", { bold: true, color: "C98A1E", size: 21 }), S(" · Präteritum-Verb unterstreichen · keine Klammer-Bögen")], { after: 100 }));

doc.push(h3("Ablauf"));
doc.push(grid([900, 3500, 5806], ["Zeit", "Phase", "Kern"], [
  ["3′", "0 Reaktivierung", "Merksatz 8 mündlich: Was ist besonders am Präteritum?"],
  ["10′", "1 Ü9a Präteritum der Sportverben", "Schnipsel eintragen, Vergleich per Kreide-Kette"],
  ["10′", "2 Ü9b Bericht für die Vereinszeitung", "Lückentext im Präteritum, Vergleich per Kreide-Kette"],
  ["12′", "3 Das Plusquamperfekt", "Zeitstrahl, haben/sein im Präteritum (Heft), Merksatz 9"],
  ["8′", "4 Ü10 Was war vorher?", "Schnipsel eintragen, Vergleich per Kreide-Kette"],
  ["(2′)", "5 optional", "Welche Zeitform? – drei Sätze mündlich"],
], { bold: [true, true, false], colors: [C.hv, NAVY, null], h: 360 }));
doc.push(p([S("Puffer: ", { bold: true, color: C.hv }), S("Wird es knapp, Ü10 nur bis Nr. 3 vergleichen und den Rest als Hausaufgabe geben. Phase 3 (Merksatz 9) darf nicht wegfallen.")], { before: 80, after: 60 }));

doc.push(h3("So sieht das Heft am Ende aus"));
doc.push(table([2300, 7906], [
  ["Überschrift", "Präteritum üben, Plusquamperfekt · Datum rechts"],
  ["Ü9a", "Schnipsel eingeklebt: Präteritum der Sportverben"],
  ["Ü9b", "Schnipsel eingeklebt: Bericht für die Vereinszeitung"],
  ["Tabelle", "haben und sein im Präteritum (ich hatte, ich war …)"],
  ["Merksatz 9", "Das Plusquamperfekt + kleiner Zeitstrahl"],
  ["Ü10", "Schnipsel eingeklebt: Was war vorher?"],
].map(([a, b]) => row([tc(a, 2300, { bold: true, color: a.startsWith("Merk") ? C.hv : (a === "Überschrift" || a === "Tabelle") ? NAVY : "C77C12", fill: LIGHT, size: 20 }), tc(b, 7906, { size: 20 })], 340))));
doc.push(p([S("Material: ", { bold: true }), S("Schnipsel-Vorlage (2 Seiten): Seite 1 (Ü9a + Ü9b, 2 Schüler pro A4) halbe Klassenstärke, Seite 2 (Ü10, 4 Schüler pro A4) ein Viertel der Klassenstärke kopieren und zerschneiden · Kreide Rot, Gelb · Schüler: Heft, Kleber, Buntstifte Rot, Gelb, Grün")], { before: 100, after: 40 }));
doc.push(p([S("Tipp: ", { bold: true }), S("Merksatz 9 und die leere Tabelle haben/sein vorab auf die Seitentafel schreiben und zuklappen. Schnipsel vor der Stunde schneiden.")], { after: 0 }));

// ---------- Phase 0 ----------
doc.push(phase("0", "Reaktivierung: Was ist besonders am Präteritum?", 3));
doc.push(...steps([
  ["gemeinsam", { say: "Meldet euch. Ihr schreibt nichts auf. Was ist besonders am Präteritum? Wo benutzt man es?", do: ["Erwartung (Merksatz 8): nur ein Verb, kein Hilfsverb; regelmäßig mit -te, unregelmäßig mit verändertem Stamm; man benutzt es schriftlich, z. B. in Berichten.", "Heft auf, Datum, Überschrift: Präteritum üben, Plusquamperfekt."] }],
]));

// ---------- Phase 1 ----------
doc.push(phase("1", "Ü9a Präteritum der Sportverben", 10));
doc.push(...steps([
  ["luecken", { say: "Jetzt arbeitet ihr allein. Ihr bekommt den Schnipsel Ü9a. Klebt ihn ein, tragt die Präteritumformen ein und markiert die unregelmäßigen mit einem Stern. Fünf Minuten.", board: auftrag("Ü9a Präteritum der Sportverben", ["Klebe den Schnipsel Ü9a ins Heft.", "Trage die Präteritumform ein (ich …).", "Markiere unregelmäßige Verben mit einem Stern."], [p(S("werfen, fallen, greifen, halten, fangen, treffen, laufen, heben, reißen, stoßen, gewinnen, verlieren, kämpfen, trainieren", { size: 20 }), { after: 0 })], "5 Minuten · allein · leise"),
    sol: "warf* · fiel* · griff* · hielt* · fing* · traf* · lief* · hob* · riss* · stieß* · gewann* · verlor* · kämpfte · trainierte" }],
  ["tafel", { say: "Kreide-Kette: Wer die Kreide hat, schreibt ein Verb mit Präteritumform an die Tafel. Alle prüfen mit und verbessern mit Grün.", do: ["Mit werfen beginnen (Judo und Handball), dann gezielt nach Sportart fragen.", "Merkhilfe für unregelmäßige Verben: „Der Vokal ändert sich, und es gibt kein -te: werfen → warf, heben → hob.“"] }],
]));

// ---------- Phase 2 ----------
doc.push(phase("2", "Ü9b Bericht für die Vereinszeitung", 10));
doc.push(...steps([
  ["luecken", { say: "Jetzt arbeitet ihr allein: Schnipsel Ü9b. Das ist ein Bericht für die Vereinszeitung. Setzt die Verben im Präteritum ein. Fünf Minuten.", board: auftrag("Ü9b Bericht für die Vereinszeitung", ["Klebe den Schnipsel Ü9b ins Heft.", "Setze das Verb in Klammern im Präteritum ein.", "Unterstreiche alle Präteritum-Verben."], [p(S("Der Text steht auf dem Schnipsel.", { size: 20, italics: true }), { after: 0 })], "5 Minuten · allein · leise"),
    sol: "fuhr · kämpften · warf · spielten · traf · waren · hob · gewann" }],
  ["tafel", { say: "Kreide-Kette: Ich lese den Bericht Satz für Satz vor. Wer die Kreide hat, schreibt das fehlende Verb an die Tafel. Alle prüfen mit und verbessern mit Grün.", do: ["Nach dem Vergleich kurz fragen: „Warum steht im Bericht Präteritum und nicht Perfekt?“ → schriftlicher Bericht (Merksatz 8)."] }],
]));

// ---------- Phase 3 ----------
doc.push(phase("3", "Das Plusquamperfekt", 12));
doc.push(...steps([
  ["zuhoeren", { say: "Stifte liegen. Ihr hört nur zu. Zwei Sätze aus einem Wettkampfbericht. Was ist zuerst passiert?", board: board("Mitte", [
    sl([n19("Ich "), u19("war"), n19(" vor dem Kampf sehr nervös. Ich "), r19("hatte"), n19(" die Nacht davor kaum "), y19("geschlafen"), n19(".")]),
    p(t(""), { after: 30 }),
    table([2600, 2600, 2600], [row([
      cell(p([t("hatte kaum geschlafen", { bold: true, size: 18, color: NAVY })], { after: 0, align: AlignmentType.CENTER }), { w: 2600, fill: "E4E9F5", borders: allBorders(solid(NAVY, 6)) }),
      cell(p([t("war nervös", { bold: true, size: 18, color: NAVY })], { after: 0, align: AlignmentType.CENTER }), { w: 2600, fill: "E4E9F5", borders: allBorders(solid(NAVY, 6)) }),
      cell(p([t("heute", { bold: true, size: 18, color: NAVY })], { after: 0, align: AlignmentType.CENTER }), { w: 2600, fill: "E4E9F5", borders: allBorders(solid(NAVY, 6)) }),
    ]), row([
      cell(p(t("Vorvergangenheit", { size: 16, italics: true, color: MUTED }), { after: 0, align: AlignmentType.CENTER }), { w: 2600, borders: noBorders }),
      cell(p(t("Vergangenheit", { size: 16, italics: true, color: MUTED }), { after: 0, align: AlignmentType.CENTER }), { w: 2600, borders: noBorders }),
      cell(p(t("Gegenwart  →", { size: 16, italics: true, color: MUTED }), { after: 0, align: AlignmentType.CENTER }), { w: 2600, borders: noBorders }),
    ])]),
  ]), do: ["Zeitstrahl darunter zeichnen. Antwort: Das Schlafen war vorher – vor der Vergangenheit.", "Kontrast: Ich habe geschlafen (Perfekt) – Ich hatte geschlafen (Plusquamperfekt). Der Unterschied steckt im Hilfsverb."] }],
  ["tafel", { say: "Kreide-Kette: Wir konjugieren haben und sein im Präteritum. Wer die Kreide hat, trägt eine Form ein. Alle anderen schreiben die Tabelle ins Heft.", board: board("Mitte · haben und sein im Präteritum", [table([1500, 1900, 1900], [
    row([cell(p(N(""), { after: 0 }), { w: 1500, borders: allBorders(solid("9AA59C", 4)) }), cell(p(R("haben"), { after: 0, align: AlignmentType.CENTER }), { w: 1900, borders: allBorders(solid("9AA59C", 4)) }), cell(p(R("sein"), { after: 0, align: AlignmentType.CENTER }), { w: 1900, borders: allBorders(solid("9AA59C", 4)) })]),
    ...[["ich", "hatte", "war"], ["du", "hattest", "warst"], ["er/sie/es", "hatte", "war"], ["wir", "hatten", "waren"], ["ihr", "hattet", "wart"], ["sie/Sie", "hatten", "waren"]].map(([ps, h, w]) => row([cell(p(N(ps, { color: MUTED }), { after: 0 }), { w: 1500, borders: allBorders(solid("9AA59C", 4)), m: 15 }), cell(p(N(h), { after: 0, align: AlignmentType.CENTER }), { w: 1900, borders: allBorders(solid("9AA59C", 4)), m: 15 }), cell(p(N(w), { after: 0, align: AlignmentType.CENTER }), { w: 1900, borders: allBorders(solid("9AA59C", 4)), m: 15 })])),
  ])]), sol: "hatte · hattest · hatte · hatten · hattet · hatten | war · warst · war · waren · wart · waren", do: ["Stolperstellen: ihr wart (ohne e), du warst."] }],
  ["abschreiben", { say: "Jetzt schreibt ihr von der Tafel ab: Merksatz 9 mit dem Beispiel und dem kleinen Zeitstrahl. Rot umrahmen.", board: merkTafel(9, "Das Plusquamperfekt", [
    p([S("Das "), S("Plusquamperfekt", { bold: true }), S(" bildet man mit "), S("haben oder sein im Präteritum", { bold: true }), S(" (hatte, war) und dem "), S("Partizip II", { bold: true }), S(".")], { after: 60 }),
    p([S("Beispiel: ", { bold: true, color: MUTED }), n19("Ich "), r19("hatte"), n19(" wochenlang "), y19("trainiert"), n19(".  ·  Wir "), r19("waren"), n19(" früh "), y19("losgefahren"), n19(".")], { after: 60 }),
    p([S("Wofür? ", { bold: true }), S("Damit erzählt man, was "), S("vorher", { bold: true }), S(" passiert war – die Vorgeschichte zu einem Ereignis in der Vergangenheit. "), S("Test: ", { bold: true }), S("Steht hatte oder war mit einem Partizip II? Dann ist es Plusquamperfekt. (Steht habe oder bin, ist es Perfekt.)")], { after: 0 }),
  ]) }],
]));

// ---------- Phase 4 ----------
doc.push(phase("4", "Ü10 Was war vorher?", 8));
doc.push(...steps([
  ["luecken", { say: "Jetzt arbeitet ihr allein: Schnipsel Ü10. Ergänzt im zweiten Satz das Plusquamperfekt. Vier Minuten.", board: auftrag("Ü10 Was war vorher?", ["Klebe den Schnipsel Ü10 ins Heft.", "Ergänze das Plusquamperfekt mit dem Verb in Klammern.", "Markiere hatte/war rot und das Partizip II gelb."], [p(S("Die Sätze stehen auf dem Schnipsel.", { size: 20, italics: true }), { after: 0 })], "4 Minuten · allein · leise"),
    sol: "1 hatte trainiert · 2 hatten gemacht · 3 hatte aufgewärmt · 4 waren losgefahren · 5 hatte geschafft" }],
  ["tafel", { say: "Kreide-Kette: Wer die Kreide hat, schreibt eine Lösung an die Tafel und sagt, was zuerst passiert ist. Alle verbessern mit Grün.", do: ["Bei Nr. 4 nachfragen: Warum war und nicht hatte? → losfahren ist Bewegung, wie beim Perfekt (sind losgefahren)."] }],
]));

// ---------- Phase 5 ----------
doc.push(phase("5", "Optional: Welche Zeitform?", 2));
doc.push(...steps([
  ["gemeinsam", { say: "Meldet euch. Welche Zeitform – und woran erkennt ihr sie?", board: board("Mitte", [bl("1  Ich habe gewonnen."), bl("2  Ich gewann."), bl("3  Ich hatte gewonnen.", { after: 0 })]), sol: "1 Perfekt (habe + Partizip II) · 2 Präteritum (nur ein Verb) · 3 Plusquamperfekt (hatte + Partizip II)" }],
]));
doc.push(p([S("Ausblick: ", { bold: true }), S("Nächste Stunde: Ü11 „Mein erster Wettkampf“ mit dem gedruckten Schreibrahmen, Partner-Check und Reflexion „Wann nehme ich welche Zeitform?“.")], { before: 120, after: 0 }));

const pageProps = { page: { size: { width: 11906, height: 16838 }, margin: { top: 900, bottom: 800, left: 850, right: 850, header: 400, footer: 400 } } };
const styles = { default: { document: { run: { font: "Calibri", size: 22 } } } };
const dScript = new Document({ styles, sections: [{ properties: pageProps, headers: { default: header("Präteritum üben, Plusquamperfekt") }, footers: { default: footer("Tafelskript") }, children: doc }] });
Packer.toBuffer(dScript).then((b) => fs.writeFileSync("Tafelskript_Praeteritum_ueben_Plusquamperfekt.docx", b));
