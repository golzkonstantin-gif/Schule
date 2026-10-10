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
// Verlaufsplan „Welche Zeitform wann?“ (Orientierungshospitation, 45 Min.)
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
  ["Lerngruppe", "Klasse 7A, ausschließlich Leistungssportler (Judo, Handball, Gewichtheben). Alle hatten am vergangenen Wochenende Wettkämpfe. Satzstreifen am Zeitstrahl und die Kreide-Kette kennt die Klasse und hat sie in der letzten Stunde gut angenommen."],
  ["Kernanliegen", "Die Schülerinnen und Schüler erarbeiten an einem Mustertext über einen Wettkampf die zeitliche Abfolge der fünf bekannten Tempora und deren Bildung als Übersicht und wenden sie in einem eigenen zusammenhängenden Text über ihr Wettkampfwochenende an."],
  ["Teilziele", "Die Schülerinnen und Schüler …\nTZ 1  ordnen die Sätze eines Mustertextes am Zeitstrahl (vorher – Vergangenheit – jetzt – Zukunft),\nTZ 2  ordnen jedem Abschnitt Zeitform und Bauplan zu (Hilfsverb + Partizip II bzw. Infinitiv) und begründen dies am Satz,\nTZ 3  sichern die Übersicht in einer Tabelle im Heft,\nTZ 4  schreiben mithilfe von Formulierungshilfen einen eigenen Text, in dem alle fünf Zeitformen passend vorkommen (Exit-Ticket)."],
].map(([a, b]) => row([
  tc(a, 3300, { bold: true, color: NAVY, fill: LIGHT, size: 19 }),
  cell(b.split("\n").map((x, i, arr) => p(s20(x), { after: i === arr.length - 1 ? 0 : 30 })), { w: LW - 3300 }),
]))));
doc.push(p(t(""), { after: 120 }));
doc.push(p(s20("UG = Unterrichtsgespräch · EA = Einzelarbeit · PA = Partnerarbeit", { color: MUTED }), { after: 60 }));
doc.push(table(VW, [
  new TableRow({ tableHeader: true, children: vhdr.map((h, i) => hdr(h, VW[i], NAVY, 18)) }),
  vrow("0′–3′", "Einstieg", [
    L("knüpft an das Wochenende an: „Ihr hattet am Wochenende alle Wettkämpfe. Erzählt in einem Satz: Wie lief es?“"),
    SS("zwei bis drei erzählen mündlich."),
    L("leitet über: „Heute schreibt ihr euer Wochenende auf – mit allen Zeitformen. Zuerst schauen wir uns an, wie das bei einem Gewichtheber aussieht.“ Stundenfrage an der Tafel: „Welche Zeitform wann?“"),
  ], ["UG"], ["Tafel: Zeitstrahl (vorgezeichnet), Stundenfrage"], [
    "Lebensweltbezug: reale Wettkämpfe der Klasse – hohe Motivation.",
    "Transparenz: Das Ziel der Stunde (eigener Text) wird zu Beginn genannt.",
  ]),
  vrow("3′–12′", "Erarbeitung I: Mustertext am Zeitstrahl", [
    L("hängt acht Satzstreifen eines Mustertextes gemischt an den Rand; Ansage: „Wer drankommt, hängt einen Streifen an die richtige Stelle am Zeitstrahl und markiert die Verbformen. Alle prüfen mit.“"),
    SS("Magnet-Kette: ordnen die Sätze am Zeitstrahl (vorher – Vergangenheit – jetzt – Zukunft), markieren Hilfsverb rot, Partizip II gelb, Infinitiv blau bzw. unterstreichen das finite Verb und begründen mit den Zeitangaben („Vor dem Wettkampf …“)."),
    L("gibt bei den beiden Trainer-Streifen („Danach sagte mein Trainer:“ / „Du hast super gehoben!“) vor: „Wer erzählt, schreibt im Präteritum. Wer spricht – wie der Trainer –, nimmt das Perfekt.“"),
  ], ["EA an der Tafel", "UG"], ["8 Satzstreifen (Großformat, Magnete)", "Kreide rot/gelb/blau"], [
    "Vorgegebener Mustertext wie in der letzten Stunde: vertraute Methode, Fokus auf der zeitlichen Struktur (TZ 1).",
    "Die Zeitangaben am Satzanfang stützen das Ordnen und sind später Formulierungshilfe.",
    "Den Unterschied Präteritum/Perfekt gibt die Lehrkraft vor – kein eigener Entdeckungsschritt nötig.",
  ]),
  vrow("12′–25′", "Erarbeitung II: Übersicht", [
    L("teilt den Schnipsel „Die Zeitformen im Überblick“ aus; hängt fünf Zeitform-Karten und fünf Bauplan-Karten gemischt an die Seite."),
    SS("Kreide-Kette: Nacheinander hängt jeweils eine Schülerin oder ein Schüler eine Karte über den passenden Abschnitt des Zeitstrahls (zuerst die Zeitformen, dann die Baupläne) und begründet am Satz darunter („hatte + trainiert: Hilfsverb im Präteritum + Partizip II“). Falsch Gehängtes stellt der Nächste um."),
    SS("Die übrigen ergänzen gleichzeitig die Tabelle: Zeitform, Bauplan, Beispielsatz aus dem Text."),
  ], ["EA an der Tafel", "EA"], ["5 Zeitform-Karten, 5 Bauplan-Karten (Großformat, Magnete)", "Schnipsel 1 (Tabelle)", "Heft, Kleber, Buntstifte"], [
    "Die Übersicht entsteht schrittweise an der Tafel und zugleich im Heft (TZ 2, TZ 3).",
    "Kreide-Kette mit Mitschreiben hat sich bewährt: Alle sind beschäftigt.",
    "Umstellen statt Anschreiben spart Zeit und macht Fehler sofort korrigierbar.",
    "Tabelle auf dem Schnipsel statt an der Tafel – an der Tafel entsteht sie aus Karten.",
  ]),
  vrow("25′–40′", "Anwendung: eigener Text (Exit-Ticket)", [
    L("teilt das Exit-Ticket aus; Auftrag: „Schreibt 5–7 Sätze über euren Wettkampf am Wochenende – nach dem Muster an der Tafel. Jede Zeitform kommt mindestens einmal vor.“ Geht herum und unterstützt."),
    SS("schreiben einen zusammenhängenden Text, gestützt durch Satzanfänge je Zeitform und den Mustertext an der Tafel; markieren die Verbformen und haken die Checkliste ab."),
  ], ["EA"], ["Exit-Ticket mit Satzanfängen je Zeitform (Bildung farbig hervorgehoben) und Schreiblinien", "Tafel (Mustertext und Übersicht)"], [
    "Transfer: Die Schüler machen den Mustertext mit eigenem Inhalt nach (TZ 4).",
    "Satzanfänge statt fertiger Sätze: kein Abschreiben, Lenkung auf Inhalte (Vorbereitung, Ablauf, wörtliche Rede, Befinden, Ziele), Bildung farbig sichtbar.",
    "Das Exit-Ticket zeigt der Lehrkraft den Lernstand jedes Einzelnen – Grundlage für die nächste Stunde.",
  ]),
  vrow("40′–45′", "Sicherung", [
    SS("zwei lesen ihren Text vor; die Klasse zeigt am Zeitstrahl, in welchem Abschnitt der Text gerade ist, und nennt die Zeitform."),
    L("beantwortet mit der Klasse die Stundenfrage in einem Satz und sammelt die Exit-Tickets ein."),
  ], ["UG"], ["Tafel", "Exit-Tickets"], [
    "Die zeitliche Struktur wird am eigenen Text hörbar.",
    "Rückbezug auf die Stundenfrage rundet die Stunde ab.",
  ]),
  vrow("opt.", "Didaktische Reserve / Kürzung", [
    plain([s20("Reserve: ", { bold: true }), s20("Speed-Duell zu zweit – einer liest einen Satz vom Zeitstrahl in einer anderen Zeitform vor, der andere bestimmt die Zeitform; ein Punkt pro richtige Antwort.")]),
    plain([s20("Kürzung: ", { bold: true }), s20("Wird es knapp, schreiben die Schüler mindestens fünf Sätze (einen pro Zeitform); in der Sicherung liest nur eine Schülerin oder ein Schüler vor.")]),
  ], ["PA"], ["–"], ["Zeitpuffer – die Klasse braucht für Kreide-Kette und Mitschreiben erfahrungsgemäß mehr Zeit."]),
]));

// ---------- Tafelbild ----------
doc.push(br());
doc.push(h3("Geplantes Tafelbild (am Ende der Erarbeitung II)"));
const TBW = [2900, 2800, 3000, 2600, 2700];
const tbc = (children, w) => cell(children, { w, fill: BOARD, borders: allBorders(solid("3D4A3F", 8)), m: 80, valign: VerticalAlign.TOP });
const card = (name, bau) => [p(t(name, { bold: true, font: "Cambria", size: 21, color: NAVY }), { after: 10, align: AlignmentType.CENTER }), p(bau, { after: 70, align: AlignmentType.CENTER })];
const ex = (runs) => p(runs, { after: 40 });
const r = (x) => t(x, { bold: true, color: C.hv, size: 19, font: "Cambria" });
const y = (x) => t(x, { bold: true, color: "C98A1E", size: 19, font: "Cambria" });
const b = (x) => t(x, { bold: true, color: C.inf, size: 19, font: "Cambria" });
const n = (x, o = {}) => t(x, Object.assign({ size: 19, font: "Cambria" }, o));
const u = (x) => n(x, { u: true, bold: true });
const bs = (x, c = MUTED) => t(x, { size: 18, color: c, bold: true });
doc.push(p([s20("Stundenfrage oben: ", { bold: true }), s20("Welche Zeitform wann?   ·   Karten (Zeitform, Bauplan) über den Abschnitten, darunter die Streifen des Mustertextes")], { after: 80 }));
doc.push(table(TBW, [
  row([
    cell(p(t("vorher", { bold: true, color: "FFFFFF", size: 20 }), { after: 0, align: AlignmentType.CENTER }), { w: TBW[0], fill: "3D4A3F" }),
    cell(p(t("Vergangenheit", { bold: true, color: "FFFFFF", size: 20 }), { after: 0, align: AlignmentType.CENTER }), { w: TBW[1] + TBW[2], fill: "3D4A3F", span: 2 }),
    cell(p(t("jetzt", { bold: true, color: "FFFFFF", size: 20 }), { after: 0, align: AlignmentType.CENTER }), { w: TBW[3], fill: "3D4A3F" }),
    cell(p(t("Zukunft  ▶", { bold: true, color: "FFFFFF", size: 20 }), { after: 0, align: AlignmentType.CENTER }), { w: TBW[4], fill: "3D4A3F" }),
  ]),
  row([
    tbc([...card("Plusquamperfekt", [bs("hatte/war", C.hv), bs(" + "), bs("Partizip II", "C98A1E")]), ex([n("Vor dem Wettkampf "), r("hatte"), n(" ich wochenlang hart "), y("trainiert"), n(".")]), ex([n("Am Morgen "), r("waren"), n(" wir früh zur Halle in Frankfurt (Oder) "), y("gefahren"), n(".")])], TBW[0]),
    tbc([...card("Präteritum (erzählt)", [bs("finites Verb im Präteritum")]), ex([n("Im Reißen "), u("schaffte"), n(" ich 50 Kilo.")]), ex([n("Im Stoßen "), u("hob"), n(" ich 62 Kilo.")]), ex([n("Danach "), u("sagte"), n(" mein Trainer:")])], TBW[1]),
    tbc([...card("Perfekt (gesprochen)", [bs("habe/bin", C.hv), bs(" + "), bs("Partizip II", "C98A1E")]), ex([n("„Du "), r("hast"), n(" super "), y("gehoben"), n("!“")])], TBW[2]),
    tbc([...card("Präsens", [bs("finites Verb im Präsens")]), ex([n("Heute "), u("analysiere"), n(" ich mit ihm das Video.")])], TBW[3]),
    tbc([...card("Futur I", [bs("werde", C.hv), bs(" + "), bs("Infinitiv", C.inf)]), ex([n("Beim nächsten Wettkampf "), r("werde"), n(" ich 65 Kilo "), b("stoßen"), n(".")])], TBW[4]),
  ]),
]));
doc.push(p([s20("Hinweis: ", { bold: true }), s20("Der Satz des Trainers ist auf zwei Streifen geteilt: „Danach sagte mein Trainer:“ (erzählt, Präteritum) und „Du hast super gehoben!“ (gesprochen, Perfekt). Sie hängen nebeneinander, so ist jede Zeitform eindeutig zuzuordnen.")], { before: 100, after: 0 }));

const landProps = { page: { size: { width: 11906, height: 16838, orientation: PageOrientation.LANDSCAPE }, margin: { top: 700, bottom: 600, left: 700, right: 700, header: 350, footer: 350 } } };
const styles = { default: { document: { run: { font: "Calibri", size: 20 } } } };
const vhead = new Header({ children: [new Paragraph({ tabStops: [{ type: TabStopType.RIGHT, position: LW }], children: [t("Deutsch · Klasse 7A · Welche Zeitform wann?", { size: 17, color: MUTED }), t("\tVerlaufsplan", { size: 17, color: MUTED })] })] });
Packer.toBuffer(new Document({ styles, sections: [{ properties: landProps, headers: { default: vhead }, footers: { default: footer("Verlaufsplan") }, children: doc }] })).then((buf) => fs.writeFileSync("Verlaufsplan_Zeitformen_Zeitstrahl.docx", buf));
