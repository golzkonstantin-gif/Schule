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
// Druckvorlage „Welche Zeitform wann?“
// Seite 1: Tabelle „Die Zeitformen im Überblick“ (2 pro A4)
// Seite 2: Exit-Ticket „Mein Wochenende in fünf Zeitformen“ mit kurzen Formulierungshilfen und Schreiblinien (2 pro A4)
// ab Seite 3: Mustertext (7 Satzstreifen), 5 Zeitform-Karten, 5 Bauplan-Karten (Großformat, 2 pro A4 quer)
// ================================================================
const { PageOrientation } = require("docx");
const PW = 11906 - 1200;
const IW = PW - 300; // Innenbreite im Schnipsel
const gray = "9AA3B5";
const s = (x, o = {}) => t(x, Object.assign({ size: 21 }, o));
const sb = (x, o = {}) => s(x, Object.assign({ bold: true }, o));
const scissors = () => p(t("✂ - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -", { color: gray, size: 16 }), { before: 20, after: 20, align: AlignmentType.CENTER });
const slip = (children) => table([PW], [row([cell(children, { w: PW, borders: allBorders(dashed), m: 100, ml: 150, valign: VerticalAlign.TOP })])]);
const slipHead = (title) => p([t(title, { font: "Cambria", bold: true, size: 26, color: NAVY }), t("          Name: ______________________", { size: 18, color: MUTED })], { after: 70 });
const red = (x, o = {}) => sb(x, Object.assign({ color: C.hv }, o));
const yel = (x, o = {}) => sb(x, Object.assign({ color: "C98A1E" }, o));
const blu = (x, o = {}) => sb(x, Object.assign({ color: C.inf }, o));

// ---------- Schnipsel 1: Tabelle ----------
const TW2 = [2050, 2050, 2650, IW - 6750];
const tb = allBorders(solid("7A869E", 8));
const th = (x, w) => cell(p(t(x, { bold: true, color: "FFFFFF", size: 20 }), { after: 0, align: AlignmentType.CENTER }), { w, fill: NAVY, borders: tb });
const WANN = ["vorher", "Vergangenheit\n(erzählt)", "Vergangenheit\n(gesprochen)", "jetzt", "Zukunft"];
function slipTab() {
  return slip([
    slipHead("Die Zeitformen im Überblick"),
    p([s("Ergänze, während die Tafel entsteht: die "), sb("Zeitform"), s(", ihren "), sb("Bauplan"), s(" und einen "), sb("Beispielsatz"), s(" aus dem Text. Markiere Hilfsverb "), red("rot"), s(", Partizip II "), yel("gelb"), s(", Infinitiv "), blu("blau"), s(".")], { after: 80 }),
    table(TW2, [
      row([th("Wann?", TW2[0]), th("Zeitform", TW2[1]), th("Bauplan", TW2[2]), th("Beispielsatz aus dem Text", TW2[3])], 380),
      ...WANN.map((w) => row([
        cell(w.split("\n").map((x, i) => p(t(x, { bold: i === 0, size: i === 0 ? 20 : 17, color: i === 0 ? NAVY : MUTED }), { after: 0 })), { w: TW2[0], fill: LIGHT, borders: tb, valign: VerticalAlign.CENTER }),
        cell(p(t(""), { after: 0 }), { w: TW2[1], borders: tb }),
        cell(p(t(""), { after: 0 }), { w: TW2[2], borders: tb }),
        cell(p(t(""), { after: 0 }), { w: TW2[3], borders: tb }),
      ], 1060, HeightRule.EXACT)),
    ]),
    p([sb("Wofür? ", { size: 19 }), s("Mit den Zeitformen ordnest du, wann etwas passiert. Wer erzählt, schreibt im Präteritum; wer spricht, nimmt das Perfekt.  ", { size: 19 }), sb("Test: ", { size: 19 }), s("Wann ist es passiert? Welches Hilfsverb steht da?", { size: 19 })], { before: 80, after: 0 }),
  ]);
}

// ---------- Schnipsel 2: Exit-Ticket ----------
const lineTable = (n) => table([IW], Array.from({ length: n }, () => row([cell(p(t(""), { after: 0 }), { w: IW, borders: { top: none, left: none, right: none, bottom: solid("9AA3B5", 6) } })], 480, HeightRule.EXACT)));
function slipExit() {
  const f = 18;
  const h = (x) => s(x, { size: f });
  const R = (x) => s(x, { size: f, bold: true, color: C.hv });
  const P2 = () => s("_______", { size: f, bold: true, color: "E8A33D" });
  const INF = () => s("_______", { size: f, bold: true, color: C.inf });
  const dot = () => s("  ·  ", { size: f, color: gray });
  const EW = [2500, IW - 2500];
  const eb = allBorders(solid("B9C6E8", 6));
  const bt = (x, c) => t(x, { size: 15, bold: true, color: c || MUTED });
  const join = (items) => items.flatMap((it, i) => (i ? [dot(), ...it] : it));
  const hrow = (name, bau, inhalt, items) => row([
    cell([p([t(name, { bold: true, size: 19, color: NAVY, font: "Cambria" })], { after: 0 }), p(bau, { after: 0 }), p(t(inhalt, { size: 15, italics: true, color: MUTED }), { after: 0 })], { w: EW[0], fill: LIGHT, borders: eb, m: 40, valign: VerticalAlign.TOP }),
    cell(p(join(items), { after: 0 }), { w: EW[1], borders: eb, m: 50 }),
  ]);
  return slip([
    p([t("Exit-Ticket: Mein Wochenende", { font: "Cambria", bold: true, size: 26, color: NAVY }), t("          Name: ____________________", { size: 18, color: MUTED })], { after: 40 }),
    p([s("Schreibe "), sb("5–7 Sätze"), s(" über deinen Wettkampf – jede Zeitform "), sb("mindestens einmal"), s(". Farbige Linie = hier kommt Partizip II bzw. Infinitiv hin. Platz reicht nicht? Rückseite!", { size: 19 })], { after: 60 }),
    table(EW, [
      hrow("Plusquamperfekt", [bt("hatte/war", C.hv), bt(" + "), bt("Partizip II", "C98A1E")], "Vorbereitung", [
        [h("Vor dem Wettkampf "), R("hatte"), h(" ich … "), P2()],
        [h("Wochenlang "), R("hatte"), h(" ich … "), P2()],
        [h("Am Abend davor "), R("hatte"), h(" ich … "), P2()],
        [h("Am Morgen "), R("waren"), h(" wir … "), P2()],
      ]),
      hrow("Präteritum", [bt("finites Verb im Präteritum")], "Ablauf und Ergebnis", [
        [h("In der Halle …")], [h("Im ersten Kampf / In der ersten Halbzeit / Im ersten Versuch …")], [h("Danach …")], [h("Am Ende …")],
      ]),
      hrow("Perfekt", [bt("habe/bin", C.hv), bt(" + "), bt("Partizip II", "C98A1E")], "Was jemand gesagt hat", [
        [h("Mein Trainer sagte: „Du "), R("hast"), h(" … "), P2(), h("“")],
        [h("Mein Team rief: „Wir "), R("haben"), h(" … "), P2(), h("“")],
        [h("Meine Eltern fragten: „"), R("Hast"), h(" du … "), P2(), h("?“")],
        [h("Ich antwortete: „Ich "), R("bin"), h(" … "), P2(), h("“")],
      ]),
      hrow("Präsens", [bt("finites Verb im Präsens")], "Wie es dir jetzt geht", [
        [h("Heute …")], [h("Jetzt fühle ich mich …")], [h("Mein Körper …")], [h("Im Moment …")],
      ]),
      hrow("Futur I", [bt("werde", C.hv), bt(" + "), bt("Infinitiv", C.inf)], "Deine Ziele", [
        [h("Beim nächsten Wettkampf "), R("werde"), h(" ich … "), INF()],
        [h("Im Training "), R("werde"), h(" ich … "), INF()],
        [h("Ab morgen "), R("werde"), h(" ich … "), INF()],
        [h("Wir "), R("werden"), h(" … "), INF()],
      ]),
    ]),
    lineTable(6),
    p([s("☐ alle fünf Zeitformen drin   ☐ Hilfsverb rot, Partizip II gelb, Infinitiv blau markiert", { size: 18 }), s("     Fertig? Stift hinlegen.", { size: 18, italics: true, color: MUTED })], { before: 80, after: 0 }),
  ]);
}

// ---------- Großformat für die Tafel ----------
const LWS = 16838 - 1200;
const big = (x, o = {}) => t(x, Object.assign({ font: "Cambria", size: 96, bold: true, color: "111111" }, o));
const STRIPS = [
  ["#3", [big("Im Reißen schaffte ich 50 Kilo.")]],
  ["#7", [big("Beim nächsten Wettkampf werde ich 65 Kilo stoßen.")]],
  ["#1", [big("Vor dem Wettkampf hatte ich wochenlang hart trainiert.")]],
  ["#5", [big("Danach sagte mein Trainer: „Du hast super gehoben!“")]],
  ["#6", [big("Heute analysiere ich mit ihm das Video.")]],
  ["#2", [big("Am Freitag waren wir mit dem Team nach Leipzig gefahren.")]],
  ["#4", [big("Im Stoßen hob ich 62 Kilo.")]],
  ["ZEITFORM", [big("Plusquamperfekt", { size: 120, color: NAVY })]],
  ["ZEITFORM", [big("Präteritum", { size: 120, color: NAVY })]],
  ["ZEITFORM", [big("Perfekt", { size: 120, color: NAVY })]],
  ["ZEITFORM", [big("Präsens", { size: 120, color: NAVY })]],
  ["ZEITFORM", [big("Futur I", { size: 120, color: NAVY })]],
  ["BAUPLAN", [big("hatte / war", { color: C.hv }), big(" + "), big("Partizip II", { color: "C98A1E" })]],
  ["BAUPLAN", [big("finites Verb im Präteritum")]],
  ["BAUPLAN", [big("habe / bin", { color: C.hv }), big(" + "), big("Partizip II", { color: "C98A1E" })]],
  ["BAUPLAN", [big("finites Verb im Präsens")]],
  ["BAUPLAN", [big("werde", { color: C.hv }), big(" + "), big("Infinitiv", { color: C.inf })]],
];
const strip = ([lab, runs]) => table([LWS], [row([cell([
  lab.startsWith("#") ? p([t(lab.slice(1), { size: 18, color: "A9B0BE" })], { after: 120 }) : p([t(lab, { size: 32, color: MUTED, bold: true })], { after: 120 }),
  p(runs, { after: 0, align: AlignmentType.CENTER }),
], { w: LWS, borders: allBorders(dashed), m: 200, ml: 260, valign: VerticalAlign.CENTER })], 5050, HeightRule.EXACT)]);
const strips = [];
STRIPS.forEach((x, i) => {
  strips.push(strip(x));
  if (i % 2 === 0 && i < STRIPS.length - 1) strips.push(p(t(""), { after: 0 }));
  else if (i < STRIPS.length - 1) strips.push(new Paragraph({ children: [new PageBreak()] }));
});

const doc = [
  slipTab(), scissors(), slipTab(),
  new Paragraph({ children: [new PageBreak()] }),
  slipExit(), scissors(), slipExit(),
];
const props = { page: { size: { width: 11906, height: 16838 }, margin: { top: 500, bottom: 400, left: 600, right: 600 } } };
const landProps = { page: { size: { width: 11906, height: 16838, orientation: PageOrientation.LANDSCAPE }, margin: { top: 500, bottom: 400, left: 600, right: 600 } } };
const styles = { default: { document: { run: { font: "Calibri", size: 20 } } } };
Packer.toBuffer(new Document({ styles, sections: [{ properties: props, children: doc }, { properties: landProps, children: strips }] })).then((buf) => fs.writeFileSync("Druckvorlage_Zeitformen_Zeitstrahl.docx", buf));
