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
// Verlaufsplan „Die Zeitformen am Zeitstrahl“ (Orientierungshospitation, 45 Min.)
// ================================================================
const { PageOrientation } = require("docx");
const LW = 16838 - 1400; // nutzbare Breite quer
const s20 = (x, o = {}) => t(x, Object.assign({ size: 19 }, o));
const lp = (runs, o = {}) => p(Array.isArray(runs) ? runs : [s20(runs)], Object.assign({ after: 50 }, o));
const L = (x) => lp([s20("L: ", { bold: true, color: NAVY }), s20(x)]);
const SS = (x) => lp([s20("S: ", { bold: true, color: "C77C12" }), s20(x)]);
const plain = (x, o = {}) => lp(Array.isArray(x) ? x : [s20(x, o)]);

const VW = [800, 1650, 6300, 1550, 1900, 3238];
const vhdr = ["Zeit", "Phase", "Unterrichtsgeschehen (L = Lehrkraft, S = Schüler)", "Sozialform", "Medien / Material", "Didaktisch-methodischer Kommentar"];
const vcell = (children, w, o = {}) => cell(children, { w, fill: o.fill, borders: allBorders(solid("B9C6E8", 6)), m: 80, ml: 100, valign: VerticalAlign.TOP });
const vrow = (zeit, phase, geschehen, sf, medien, komm) => new TableRow({ cantSplit: false, children: [
  vcell([p(t(zeit, { bold: true, size: 19, color: C.hv }), { after: 0 })], VW[0], { fill: LIGHT }),
  vcell([p(t(phase, { bold: true, size: 19, color: NAVY, font: "Cambria" }), { after: 0 })], VW[1], { fill: LIGHT }),
  vcell(geschehen, VW[2]),
  vcell(sf.map((x) => plain(x)), VW[3]),
  vcell(medien.map((x) => plain(x)), VW[4]),
  vcell(komm.map((x) => plain(x)), VW[5]),
] });

const doc = [];
doc.push(kicker("Verlaufsplan · Orientierungshospitation · Deutsch · Klasse 7A · Einzelstunde (45 Min.)"));
doc.push(h1("Welche Zeitform wann? – Die Zeitformen am Zeitstrahl"));
doc.push(table([3300, LW - 3300], [
  ["Einheit", "Verben und Zeitformen (Präsens, Präteritum, Perfekt, Plusquamperfekt, Futur I; Futur II folgt später)"],
  ["Lerngruppe", "Klasse 7A, ausschließlich Leistungssportler (Judo, Handball, Gewichtheben). Alle hatten am vergangenen Wochenende Wettkämpfe."],
  ["Kernanliegen", "Die Schülerinnen und Schüler ordnen Sätze über ihr Wettkampfwochenende begründet den fünf bekannten Tempora auf einem Zeitstrahl zu und reflektieren, dass die Wahl der Zeitform vom Zeitbezug und von der Textsorte (geschrieben oder gesprochen) abhängt."],
  ["Teilziele", "Die Schülerinnen und Schüler …\nTZ 1  bestimmen die Zeitform eines Satzes anhand ihrer Bildung (Hilfsverb + Partizip II bzw. Infinitiv) und begründen sie mit Fachbegriffen,\nTZ 2  ordnen die Tempora zeitlich ein (vorher – Vergangenheit – jetzt – Zukunft),\nTZ 3  erkennen, dass man Vergangenes schriftlich im Präteritum, mündlich meist im Perfekt erzählt und die Vorgeschichte im Plusquamperfekt steht,\nTZ 4  wenden dieses Wissen an, indem sie falsche Verbformen in einem Bericht für die Vereinszeitung verbessern und begründen."],
].map(([a, b]) => row([
  tc(a, 3300, { bold: true, color: NAVY, fill: LIGHT, size: 19 }),
  cell(b.split("\n").map((x, i, arr) => p(s20(x), { after: i === arr.length - 1 ? 0 : 30 })), { w: LW - 3300 }),
]))));
doc.push(p(t(""), { after: 120 }));
doc.push(p(s20("UG = Unterrichtsgespräch · EA = Einzelarbeit · PA = Partnerarbeit", { color: MUTED }), { after: 60 }));
doc.push(table(VW, [
  new TableRow({ tableHeader: true, children: vhdr.map((h, i) => hdr(h, VW[i], NAVY, 18)) }),
  vrow("0′–4′", "Einstieg", [
    L("knüpft an das Wochenende an: „Ihr hattet am Wochenende alle Wettkämpfe. Erzählt in einem Satz: Wie lief es?“"),
    SS("zwei bis drei erzählen mündlich – erfahrungsgemäß im Perfekt."),
    L("greift das auf: „Ihr habt gerade im Perfekt erzählt. In der Vereinszeitung klingt das anders.“ Zeigt den Zeitstrahl und acht ungeordnete Satzstreifen aus drei Texten (Sprachnachricht, Vereinszeitung, Post). Stundenfrage an der Tafel: „Welche Zeitform wann?“"),
  ], ["UG"], ["Tafel: Zeitstrahl (vorgezeichnet)", "8 Satzstreifen mit Magneten (Großformat)"], [
    "Lebensweltbezug: reale Wettkämpfe der Klasse – hohe Motivation.",
    "Die spontane Verwendung des Perfekts wird zum Anlass für die spätere Reflexion der Sprachverwendung (integrativer Zugang).",
  ]),
  vrow("4′–17′", "Erarbeitung I: Zeitstrahl", [
    L("teilt den Schnipsel „Die Zeitformen am Zeitstrahl“ aus, gibt die feste Ansage: „Wer drankommt, hängt einen Streifen an den Zeitstrahl, markiert und begründet. Alle prüfen mit und tragen ein.“"),
    SS("Magnet-Kette: Nacheinander hängt jeweils eine Schülerin oder ein Schüler einen Streifen an den Zeitstrahl, markiert Hilfsverb rot, Partizip II gelb, Infinitiv blau bzw. unterstreicht das finite Verb, nennt die Zeitform und begründet sie („Plusquamperfekt, weil hatte + Partizip II“)."),
    SS("Die übrigen prüfen mit und tragen Zeitform und einen Beispielsatz in den Schnipsel ein."),
    L("lässt im Anschluss den Bauplan jeder Zeitform nennen; S ergänzen die Zeile „Bauplan“."),
  ], ["EA an der Tafel", "UG"], ["Satzstreifen, Kreide rot/gelb/blau", "Schnipsel 1 (Zeitstrahl)", "Heft, Kleber, Buntstifte"], [
    "Bekannte, bewährte Routine (Kreide-Kette mit Mitschreiben): Sicherheit und Aktivierung aller.",
    "Begründung über die Bildung sichert Fachbegriffe (TZ 1).",
    "Zeitstrahl statt Tafel-Tabelle – die Klasse ist tabellenmüde.",
    "Präteritum und Perfekt landen am selben Platz: bewusst erzeugter Widerspruch als Überleitung.",
  ]),
  vrow("17′–24′", "Erarbeitung II: Reflexion + Merksatz 10", [
    L("Impuls: „Präteritum und Perfekt hängen am selben Platz. Schaut auf die Etiketten: Aus welchem Text kommen die Sätze?“"),
    SS("erkennen: Vereinszeitung (geschrieben) → Präteritum, Vorgeschichte im Plusquamperfekt; Sprachnachricht (gesprochen) → Perfekt; Post über heute und morgen → Präsens, Futur I."),
    L("deckt Merksatz 10 an der Seitentafel auf."),
    SS("schreiben Merksatz 10 ins Heft ab und rahmen ihn rot ein."),
  ], ["UG", "EA"], ["Tafel (Etiketten der Streifen)", "Seitentafel: Merksatz 10", "Heft"], [
    "Integrativer Kern: Grammatik wird in ihrer Funktion für Textsorten reflektiert (TZ 3).",
    "Merksatz wie gewohnt mit „Wofür?“ und Test.",
  ]),
  vrow("24′–34′", "Anwendung: Fehlertext", [
    L("teilt Schnipsel Ü12 aus; Auftrag steht im gewohnten Format an der Tafel. Geht herum und unterstützt."),
    SS("lesen den Bericht eines Teamkollegen für die Vereinszeitung, unterstreichen die vier falschen Verbformen, schreiben die richtige Form darüber und begründen einen Fehler im Heft („… ist falsch, weil …“)."),
    SS("Zusatz für Schnelle: ein Satz im Plusquamperfekt zur Vorgeschichte des eigenen Wettkampfs."),
  ], ["EA"], ["Schnipsel 2 (Ü12)", "Heft"], [
    "Fehleranalyse statt Lückentext: altersgerecht, verlangt Begründung (TZ 4).",
    "Die Fehler spiegeln typische Schwierigkeiten: Perfekt statt Präteritum, unregelmäßiges Präteritum, falsches Hilfsverb.",
    "Zusatzaufgabe als Binnendifferenzierung.",
  ]),
  vrow("34′–42′", "Sicherung", [
    L("hängt die vier Fehlersätze als Streifen auf."),
    SS("Kreide-Kette: schreiben die richtige Verbform unter den Streifen und begründen mündlich mit Fachbegriffen."),
    SS("vergleichen und verbessern im Schnipsel mit Grün."),
  ], ["EA an der Tafel", "UG"], ["4 Fehlerstreifen", "Kreide, grüne Stifte"], [
    "Rückbezug auf Zeitstrahl und Merksatz 10.",
    "Verbessern mit Grün macht den Lernzuwachs im Heft sichtbar.",
  ]),
  vrow("42′–45′", "Ausblick / Hausaufgabe", [
    L("beantwortet mit der Klasse die Stundenfrage noch einmal in einem Satz."),
    L("Hausaufgabe: „Schreibt euren Bericht über dieses Wochenende für die Vereinszeitung – im Präteritum, die Vorgeschichte im Plusquamperfekt.“ (Schreibrahmen Ü11)"),
  ], ["UG"], ["Schreibrahmen Ü11 (bereits ausgeteilt)"], [
    "Transfer in die eigene Textproduktion zum realen Wettkampf.",
  ]),
  vrow("opt.", "Didaktische Reserve / Kürzung", [
    plain([s20("Reserve: ", { bold: true }), s20("Speed-Duell zu zweit – einer liest einen Satz vom Zeitstrahl in einer anderen Zeitform vor, der andere bestimmt die Zeitform; ein Punkt pro richtige Antwort.")]),
    plain([s20("Kürzung: ", { bold: true }), s20("Wird es knapp, in der Sicherung nur zwei Fehler an der Tafel besprechen; die übrigen zu Beginn der nächsten Stunde.")]),
  ], ["PA"], ["–"], ["Zeitpuffer – die Klasse braucht für Kreide-Kette und Mitschreiben erfahrungsgemäß mehr Zeit."]),
]));

// ---------- Tafelbild ----------
doc.push(br());
doc.push(h3("Geplantes Tafelbild"));
const TBW = [2700, 2900, 2900, 2600, 2700];
const sec = (x, span) => cell(p(t(x, { bold: true, color: "FFFFFF", size: 20 }), { after: 0, align: AlignmentType.CENTER }), { w: span ? TBW[1] + TBW[2] : 0, fill: "3D4A3F", span });
const tb = (children, w) => cell(children, { w, fill: BOARD, borders: allBorders(solid("3D4A3F", 8)), m: 80, valign: VerticalAlign.TOP });
const zf = (name, bau) => [p(t(name, { bold: true, font: "Cambria", size: 21, color: NAVY }), { after: 10, align: AlignmentType.CENTER }), p(t(bau, { size: 18, color: MUTED }), { after: 60, align: AlignmentType.CENTER })];
const ex = (runs) => p(runs, { after: 40 });
const r = (x) => t(x, { bold: true, color: C.hv, size: 19, font: "Cambria" });
const y = (x) => t(x, { bold: true, color: "C98A1E", size: 19, font: "Cambria" });
const b = (x) => t(x, { bold: true, color: C.inf, size: 19, font: "Cambria" });
const n = (x, o = {}) => t(x, Object.assign({ size: 19, font: "Cambria" }, o));
const tag = (x) => t(x + "  ", { size: 15, color: MUTED, bold: true });
doc.push(p([s20("Stundenfrage oben: ", { bold: true }), s20("Welche Zeitform wann?   ·   Links und rechts: Seitentafel mit Merksatz 10 und Auftrag Ü12")], { after: 80 }));
doc.push(table(TBW, [
  row([
    cell(p(t("vorher", { bold: true, color: "FFFFFF", size: 20 }), { after: 0, align: AlignmentType.CENTER }), { w: TBW[0], fill: "3D4A3F" }),
    cell(p(t("Vergangenheit", { bold: true, color: "FFFFFF", size: 20 }), { after: 0, align: AlignmentType.CENTER }), { w: TBW[1] + TBW[2], fill: "3D4A3F", span: 2 }),
    cell(p(t("jetzt", { bold: true, color: "FFFFFF", size: 20 }), { after: 0, align: AlignmentType.CENTER }), { w: TBW[3], fill: "3D4A3F" }),
    cell(p(t("Zukunft  ▶", { bold: true, color: "FFFFFF", size: 20 }), { after: 0, align: AlignmentType.CENTER }), { w: TBW[4], fill: "3D4A3F" }),
  ]),
  row([
    tb([...zf("Plusquamperfekt", "hatte/war + Partizip II"), ex([tag("ZEITUNG"), n("Die Judoka "), r("hatten"), n(" wochenlang für das Turnier "), y("trainiert"), n(".")]), ex([tag("ZEITUNG"), n("Die Handballer "), r("waren"), n(" schon am Freitag "), y("angereist"), n(".")])], TBW[0]),
    tb([...zf("Präteritum", "nur finites Verb · geschrieben"), ex([tag("ZEITUNG"), n("Im Stoßen "), n("hob", { u: true, bold: true }), n(" unsere Gewichtheberin 75 Kilo.")]), ex([tag("ZEITUNG"), n("Unsere Handballer "), n("kämpften", { u: true, bold: true }), n(" bis zur letzten Sekunde.")])], TBW[1]),
    tb([...zf("Perfekt", "habe/bin + Partizip II · gesprochen"), ex([tag("SPRACHNACHRICHT"), n("Ich "), r("habe"), n(" im Finale Bronze "), y("geholt"), n("!")]), ex([tag("SPRACHNACHRICHT"), n("Wir "), r("sind"), n(" erst um zehn Uhr nach Hause "), y("gekommen"), n(".")])], TBW[2]),
    tb([...zf("Präsens", "nur finites Verb"), ex([tag("POST"), n("Heute "), n("analysieren", { u: true, bold: true }), n(" wir mit dem Trainer das Video.")])], TBW[3]),
    tb([...zf("Futur I", "werde + Infinitiv"), ex([tag("POST"), n("Beim nächsten Turnier "), r("werde"), n(" ich den Haltegriff besser "), b("verteidigen"), n(".")])], TBW[4]),
  ]),
]));
doc.push(p([s20("Unter dem Zeitstrahl (Sicherung): ", { bold: true }), s20("die vier Fehlerstreifen aus Ü12, darunter jeweils die Verbesserung: fuhren · gewann · waren … losgefahren · warfen")], { before: 100, after: 60 }));
doc.push(h3("Merksatz 10 (Seitentafel)"));
doc.push(box([
  p([t("Welche Zeitform wann?", { bold: true, color: C.hv, font: "Cambria", size: 22 })], { after: 40 }),
  p([s20("Was "), s20("vorher", { bold: true }), s20(" war, steht im "), s20("Plusquamperfekt", { bold: true }), s20(". Vergangenes "), s20("schreibt", { bold: true }), s20(" man im "), s20("Präteritum", { bold: true }), s20(", "), s20("mündlich", { bold: true }), s20(" erzählt man meist im "), s20("Perfekt", { bold: true }), s20(". Was jetzt ist: "), s20("Präsens", { bold: true }), s20(". Was kommt: "), s20("Futur I", { bold: true }), s20(".")], { after: 40 }),
  p([s20("Wofür? ", { bold: true }), s20("Damit jeder Text die passende Zeitform hat: der Bericht für die Vereinszeitung im Präteritum, die Sprachnachricht im Perfekt.   "), s20("Test: ", { bold: true }), s20("Wann ist es passiert? Und: Schreibe ich oder spreche ich?")], { after: 0 }),
], LIGHT2, allBorders(solid(C.hv, 12))));

const landProps = { page: { size: { width: 11906, height: 16838, orientation: PageOrientation.LANDSCAPE }, margin: { top: 700, bottom: 600, left: 700, right: 700, header: 350, footer: 350 } } };
const styles = { default: { document: { run: { font: "Calibri", size: 20 } } } };
const vhead = new Header({ children: [new Paragraph({ tabStops: [{ type: TabStopType.RIGHT, position: LW }], children: [t("Deutsch · Klasse 7A · Welche Zeitform wann?", { size: 17, color: MUTED }), t("\tVerlaufsplan", { size: 17, color: MUTED })] })] });
Packer.toBuffer(new Document({ styles, sections: [{ properties: landProps, headers: { default: vhead }, footers: { default: footer("Verlaufsplan") }, children: doc }] })).then((buf) => fs.writeFileSync("Verlaufsplan_Zeitformen_Zeitstrahl.docx", buf));
