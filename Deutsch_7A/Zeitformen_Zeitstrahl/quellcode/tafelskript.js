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
// Tafelskript „Welche Zeitform wann?“ (Orientierungshospitation, 45 Min.)
// ================================================================
const doc = [];
doc.push(kicker("Tafelskript für die Lehrkraft · Einzelstunde (45 Min.) · Unterrichtsbesuch"), h1("Welche Zeitform wann?"));
doc.push(box([
  p([S("Anknüpfung: ", { bold: true }), S("Letzte Stunde hat die Klasse einen Wettkampftag mit Satzstreifen am Zeitstrahl geordnet – das hat sie gut verstanden. Heute dasselbe Vorgehen mit einem Mustertext, der alle fünf Zeitformen enthält. Die Streifen sind diesmal im Großformat (2 pro A4 quer), die Schnipsel haben große Felder. Alle Schüler hatten am Wochenende Wettkämpfe – darüber schreiben sie am Ende selbst.")], { after: 0 }),
], LIGHT2));
doc.push(p([S("Ziel: ", { bold: true }), S("Die Klasse erarbeitet an einem Mustertext die zeitliche Abfolge und die Bildung der fünf Zeitformen als Übersicht und schreibt danach einen eigenen Text über ihr Wettkampfwochenende (Exit-Ticket).")], { before: 120, after: 60 }));
doc.push(p([S("Markieren wie bisher: ", { bold: true }), t("Hilfsverb rot", { bold: true, color: C.hv, size: 21 }), S(" · "), t("Partizip II gelb", { bold: true, color: "C98A1E", size: 21 }), S(" · "), t("Infinitiv blau", { bold: true, color: C.inf, size: 21 }), S(" · Präteritum/Präsens: finites Verb unterstreichen · keine Klammer-Bögen")], { after: 100 }));

doc.push(h3("Ablauf"));
doc.push(grid([900, 3700, 5606], ["Zeit", "Phase", "Kern"], [
  ["3′", "0 Einstieg: unser Wochenende", "2–3 Schüler erzählen einen Satz; Ziel der Stunde nennen"],
  ["9′", "1 Mustertext am Zeitstrahl", "Magnet-Kette: 7 Streifen ordnen und markieren"],
  ["13′", "2 Die Übersicht entsteht", "Kreide-Kette mit Karten; alle ergänzen die Tabelle"],
  ["15′", "3 Exit-Ticket: mein Wochenende", "eigener Text, jede Zeitform mindestens einmal"],
  ["5′", "4 Vorlesen und einsammeln", "zwei Texte, Klasse zeigt am Zeitstrahl mit"],
], { bold: [true, true, false], colors: [C.hv, NAVY, null], h: 360 }));
doc.push(p([S("Puffer: ", { bold: true, color: C.hv }), S("Speed-Duell zu zweit (einer liest einen Satz vom Zeitstrahl in einer anderen Zeitform, der andere bestimmt sie). Kürzung: Beim Exit-Ticket reichen fünf Sätze, in Phase 4 liest nur einer vor.")], { before: 80, after: 60 }));

doc.push(h3("So sieht das Heft am Ende aus"));
doc.push(table([2300, 7906], [
  ["Überschrift", "Welche Zeitform wann? · Datum rechts"],
  ["Schnipsel", "Die Zeitformen im Überblick (Tabelle: Zeitform, Bauplan, Beispielsatz – mit „Wofür?“ und „Test“)"],
  ["Exit-Ticket", "wird eingesammelt (zurückgeben und einkleben in der nächsten Stunde)"],
].map(([a, b2]) => row([tc(a, 2300, { bold: true, color: a === "Überschrift" ? NAVY : "C77C12", fill: LIGHT, size: 20 }), tc(b2, 7906, { size: 20 })], 340))));
doc.push(p([S("Material: ", { bold: true }), S("Druckvorlage: Seite 1 Tabelle (2 pro A4), Seite 2 Exit-Ticket mit Satzanfängen (mind. 4 je Zeitform, geordnet nach Inhalt: Vorbereitung – Ablauf und Ergebnis – was jemand gesagt hat – wie es dir jetzt geht – Ziele) und Schreiblinien (2 pro A4), Seiten 3–11 für die Tafel (einmal drucken, schneiden, Magnete): 7 Streifen Mustertext (kleine graue Zahl oben links = richtige Reihenfolge), 5 Zeitform-Karten, 5 Bauplan-Karten. Kreide rot, gelb, blau · Schüler: Heft, Kleber, Buntstifte rot, gelb, blau.")], { before: 100, after: 40 }));
doc.push(p([S("Vorbereitung: ", { bold: true }), S("Zeitstrahl über die ganze Tafel zeichnen (Pfeil, Abschnitte „vorher – Vergangenheit – jetzt – Zukunft“; Vergangenheit doppelt so breit), darüber die Stundenfrage „Welche Zeitform wann?“. Mustertext-Streifen gemischt links an den Rand hängen, Zeitform- und Bauplan-Karten gemischt rechts – die Karten erst in Phase 2 zeigen.")], { after: 0 }));

// ---------- Phase 0 ----------
doc.push(phase("0", "Einstieg: unser Wochenende", 3));
doc.push(...steps([
  ["gemeinsam", { say: "Stifte liegen. Ihr hattet am Wochenende alle Wettkämpfe. Erzählt in einem Satz: Wie lief es?", do: ["Zwei bis drei Schüler drannehmen, möglichst aus jeder Sportart einen."] }],
  ["zuhoeren", { say: "Am Ende der Stunde schreibt ihr euer Wochenende auf – mit allen Zeitformen, die ihr kennt. Vorher schauen wir uns an, wie das bei einem Gewichtheber aussieht. Heft auf, Datum, Überschrift: Welche Zeitform wann?" }],
]));

// ---------- Phase 1 ----------
const r = (x) => t(x, { bold: true, color: C.hv, size: 20, font: "Cambria" });
const y = (x) => t(x, { bold: true, color: "C98A1E", size: 20, font: "Cambria" });
const b = (x) => t(x, { bold: true, color: C.inf, size: 20, font: "Cambria" });
const n = (x, o = {}) => t(x, Object.assign({ size: 20, font: "Cambria" }, o));
const u = (x) => n(x, { u: true, bold: true });
const lab = (x) => t(x + "   ", { size: 16, color: MUTED, bold: true });
doc.push(phase("1", "Mustertext am Zeitstrahl", 9));
doc.push(...steps([
  ["tafel", { say: "Stifte liegen. Links hängen sieben Sätze eines Gewichthebers – durcheinander. Wer drankommt, hängt einen Streifen an die richtige Stelle am Zeitstrahl, markiert die Verbformen und sagt, woran er es erkennt. Alle prüfen mit.", board: board("Mitte · Zeitstrahl mit dem Mustertext (Lösung)", [
    bl([lab("vorher"), n("Vor dem Wettkampf "), r("hatte"), n(" ich wochenlang hart "), y("trainiert"), n(".")], { after: 10 }),
    bl([lab("vorher"), n("Am Morgen "), r("waren"), n(" wir früh zur Halle in Frankfurt (Oder) "), y("gefahren"), n(".")], { after: 50 }),
    bl([lab("Vergangenheit"), n("Im Reißen "), u("schaffte"), n(" ich 50 Kilo.")], { after: 10 }),
    bl([lab("Vergangenheit"), n("Im Stoßen "), u("hob"), n(" ich 62 Kilo.")], { after: 10 }),
    bl([lab("Vergangenheit"), n("Danach "), u("sagte"), n(" mein Trainer: „Du "), r("hast"), n(" super "), y("gehoben"), n("!“")], { after: 50 }),
    bl([lab("jetzt"), n("Heute "), u("analysiere"), n(" ich mit ihm das Video.")], { after: 50 }),
    bl([lab("Zukunft"), n("Beim nächsten Wettkampf "), r("werde"), n(" ich 65 Kilo "), b("stoßen"), n(".")], { after: 0 }),
  ]), do: [
    "Begründung über die Zeitangabe und die Verbform einfordern: „Vor dem Wettkampf – und hatte + trainiert, das war vorher.“",
    "Die beiden vorher-Sätze dürfen getauscht hängen. Reißen kommt im Wettkampf vor dem Stoßen.",
  ] }],
  ["zuhoeren", { say: "Stifte liegen. Schaut auf den Trainer-Satz: Da stehen zwei Zeitformen in einem Satz. Wer erzählt, schreibt im Präteritum – sagte. Wer spricht, wie der Trainer, nimmt das Perfekt – hast gehoben.", do: ["Der Trainer-Satz hängt im Abschnitt Vergangenheit an der Grenze zwischen Präteritum und Perfekt (die Karten kommen in Phase 2 darüber)."] }],
]));

// ---------- Phase 2 ----------
doc.push(phase("2", "Die Übersicht entsteht", 13));
doc.push(...steps([
  ["tafel", { say: "Ihr bekommt einen Schnipsel mit einer Tabelle. Klebt ihn ein. Rechts hängen Karten mit den Zeitformen und den Bauplänen. Kreide-Kette: Wer drankommt, hängt eine Karte über den richtigen Abschnitt und begründet am Satz darunter. Hängt etwas falsch, stellt der Nächste es um. Alle anderen ergänzen die Tabelle: Zeitform, Bauplan und einen Beispielsatz.", board: board("über den Abschnitten · Karten (Lösung)", [
    bl([lab("vorher"), N("Plusquamperfekt", { bold: true, color: NAVY }), n("   "), r("hatte/war"), n(" + "), y("Partizip II")], { after: 20 }),
    bl([lab("Vergangenheit"), N("Präteritum", { bold: true, color: NAVY }), n("   finites Verb im Präteritum   ·   "), N("Perfekt", { bold: true, color: NAVY }), n("   "), r("habe/bin"), n(" + "), y("Partizip II")], { after: 20 }),
    bl([lab("jetzt"), N("Präsens", { bold: true, color: NAVY }), n("   finites Verb im Präsens")], { after: 20 }),
    bl([lab("Zukunft"), N("Futur I", { bold: true, color: NAVY }), n("   "), r("werde"), n(" + "), b("Infinitiv")], { after: 0 }),
  ]), do: [
    "Reihenfolge: erst die fünf Zeitform-Karten, dann die fünf Bauplan-Karten. Jede Karte ein anderer Schüler – zehn Schüler sind an der Tafel.",
    "Begründungsmuster: „Präteritum, weil hob nur ein finites Verb ist – ohne Hilfsverb.“ · „hatte/war + Partizip II, weil hatte … trainiert.“",
    "Nach jeder Karte kurz warten, bis alle eingetragen haben. In der Tabelle reicht ein Beispielsatz pro Zeile.",
  ] }],
  ["kontrolle", { say: "Stift hinlegen. Lest unten auf dem Schnipsel „Wofür?“ und „Test“. Wann ist es passiert? Welches Hilfsverb steht da? Damit findet ihr jede Zeitform.", do: ["Ein Schüler liest den Test laut vor und wendet ihn auf einen Satz an der Tafel an."] }],
]));

// ---------- Phase 3 ----------
doc.push(phase("3", "Exit-Ticket: mein Wochenende", 15));
doc.push(...steps([
  ["luecken", { say: "Jetzt arbeitet ihr allein. Ihr schreibt über euren Wettkampf am Wochenende – nach dem Muster an der Tafel. Fünf bis sieben Sätze, jede Zeitform mindestens einmal. Auf dem Exit-Ticket stehen zu jeder Zeitform Satzanfänge. Die farbige Linie zeigt, wohin das Partizip II oder der Infinitiv kommt – den Rest formuliert ihr selbst. Ihr gebt das Ticket am Ende ab.", board: auftrag("Exit-Ticket: Mein Wochenende in fünf Zeitformen", ["Schreibe 5–7 Sätze über deinen Wettkampf am Wochenende.", "Benutze jede Zeitform mindestens einmal.", "Markiere Hilfsverb rot, Partizip II gelb, Infinitiv blau.", "Hake die Checkliste ab."], [p(S("Hilfe: Mustertext an der Tafel, Satzanfänge je Zeitform auf dem Ticket.", { size: 20, italics: true }), { after: 0 })], "12 Minuten · allein · leise"),
    do: [
      "Herumgehen. Typische Stolperstellen: Wechsel ins Perfekt beim Erzählen („Dann habe ich gewonnen“), habe statt hatte in der Vorgeschichte, regelmäßig gebildetes Präteritum („werfte“).",
      "Wer schnell fertig ist: einen zweiten Satz zur Vorgeschichte schreiben, diesmal mit war.",
      "Zwei Schüler mit gelungenem Text für das Vorlesen ansprechen.",
    ] }],
]));

// ---------- Phase 4 ----------
doc.push(phase("4", "Vorlesen und einsammeln", 5));
doc.push(...steps([
  ["gemeinsam", { say: "Stifte liegen. Zwei lesen ihren Text vor. Alle anderen zeigen am Zeitstrahl mit: Wo sind wir gerade – vorher, Vergangenheit, jetzt oder Zukunft?", do: ["Die Lehrkraft zeigt beim Vorlesen am Zeitstrahl mit, die Klasse nennt die Zeitform."] }],
  ["fertig", { say: "Beantwortet die Stundenfrage in einem Satz: Welche Zeitform wann? … Gebt jetzt eure Exit-Tickets ab.", do: ["Erwartete Antwort: Was vorher war: Plusquamperfekt; erzählte Vergangenheit: Präteritum, gesprochen: Perfekt; jetzt: Präsens; Zukunft: Futur I."] }],
]));
doc.push(p([S("Ausblick: ", { bold: true }), S("Die Exit-Tickets zeigen, wer die Zeitformen schon sicher im Text einsetzt. Nächste Stunde: Rückgabe mit kurzem Feedback, Überarbeitung, danach Futur II, Hilfsverb oder Vollverb, Modalverben – und zum Schluss die Satzklammer.")], { before: 120, after: 0 }));

const pageProps = { page: { size: { width: 11906, height: 16838 }, margin: { top: 900, bottom: 800, left: 850, right: 850, header: 400, footer: 400 } } };
const styles = { default: { document: { run: { font: "Calibri", size: 22 } } } };
Packer.toBuffer(new Document({ styles, sections: [{ properties: pageProps, headers: { default: header("Welche Zeitform wann?") }, footers: { default: footer("Tafelskript") }, children: doc }] })).then((buf) => fs.writeFileSync("Tafelskript_Zeitformen_Zeitstrahl.docx", buf));
