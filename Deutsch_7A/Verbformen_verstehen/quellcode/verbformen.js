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

const header = (label, name = true) => new Header({
  children: [new Paragraph({
    tabStops: [{ type: TabStopType.RIGHT, position: W }],
    children: [t("Deutsch · Klasse 7A · " + label, { size: 17, color: MUTED }), t(name ? "\tName: ______________________   Datum: ___________" : "\tfür die Lehrkraft", { size: 17, color: MUTED })],
  })],
});
const footer = (label) => new Footer({
  children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ font: "Calibri", size: 16, color: MUTED, children: [label + " · Seite ", PageNumber.CURRENT] })] })],
});
const task = (nr, title, sub) => [
  p([t(`Übung ${nr}  `, { font: "Cambria", size: 26, bold: true, color: C.hv }), t(title, { font: "Cambria", size: 26, bold: true, color: NAVY })], { before: 160, after: 40, keepNext: true }),
  sub ? p(t(sub, { size: 21, color: "333333" }), { after: 100, keepNext: true }) : null,
].filter(Boolean);

// ================================================================
// Inhalte
// ================================================================
const u1 = [["ging", "gehen"], ["isst", "essen"], ["hat", "haben"], ["war", "sein"], ["konnte", "können"], ["liest", "lesen"], ["bin", "sein"], ["fuhr", "fahren"], ["wusste", "wissen"], ["schläft", "schlafen"]];
const u2words = ["habe", "Hand", "ist", "hat", "seit", "sind", "hast", "Hafen", "seid", "habt", "bin", "hart", "bist", "haben"];
const u2sent = [["Wir sind heute sehr müde.", "sind", "sein"], ["Hast du einen Stift für mich?", "Hast", "haben"], ["Ich bin zwölf Jahre alt.", "bin", "sein"], ["Ihr habt ein tolles Plakat.", "habt", "haben"], ["Das ist mein Hund.", "ist", "sein"]];
const persons = ["ich", "du", "er / sie / es", "wir", "ihr", "sie / Sie"];
const konj = { hp: ["habe", "hast", "hat", "haben", "habt", "haben"], hv: ["hatte", "hattest", "hatte", "hatten", "hattet", "hatten"], sp: ["bin", "bist", "ist", "sind", "seid", "sind"], sv: ["war", "warst", "war", "waren", "wart", "waren"] };
const u4 = [["Gestern ______ ich krank.", "sein", "war"], ["Du ______ hohes Fieber.", "haben", "hattest"], ["Am Wochenende ______ wir im Schwimmbad.", "sein", "waren"], ["Ihr ______ keine Zeit für uns.", "haben", "hattet"], ["Mein Opa ______ früher Bäcker.", "sein", "war"], ["Die Kinder ______ großen Hunger.", "haben", "hatten"], ["______ du schon einmal in Berlin?", "sein", "Warst"]];
const u5 = [["spielen", "gespielt"], ["machen", "gemacht"], ["kaufen", "gekauft"], ["lachen", "gelacht"], ["gehen", "gegangen"], ["fahren", "gefahren"], ["schreiben", "geschrieben"], ["essen", "gegessen"], ["schlafen", "geschlafen"], ["trinken", "getrunken"]];
const u6 = [
  ["Am Samstag ", "bin", " ich früh aufgewacht."],
  ["Dann ", "habe", " ich mit meiner Schwester Pfannkuchen gebacken."],
  ["Am Nachmittag ", "sind", " wir mit dem Rad zum See gefahren."],
  ["Mein Bruder ", "hat", " dort Fußball gespielt."],
  ["Er ", "hat", " drei Tore geschossen."],
  ["Abends ", "sind", " wir nach Hause gelaufen."],
  ["Ich ", "habe", " noch ein Buch gelesen."],
  ["Um neun Uhr ", "bin", " ich eingeschlafen."],
];
const u7 = [["Ich habe Hunger.", "Vollverb"], ["Ich habe Nudeln gekocht.", "Hilfsverb"], ["Lisa ist sehr müde.", "Vollverb"], ["Lisa ist spät eingeschlafen.", "Hilfsverb"], ["Wir hatten viel Spaß.", "Vollverb"], ["Ihr seid mit dem Bus gekommen.", "Hilfsverb"]];
const exit = [["Wir waren im Zoo.", "waren", "sein", "Präteritum", "nein"], ["Du hast gewonnen.", "hast", "haben", "Präsens", "ja"], ["Sie hatten keine Zeit.", "hatten", "haben", "Präteritum", "nein"], ["Ich bin nach Hause gerannt.", "bin", "sein", "Präsens", "ja"]];

// ================================================================
// ARBEITSBLATT
// ================================================================
const ab = [];
ab.push(kicker("Wiederholung · Doppelstunde"), h1("Verbformen verstehen"),
  p(t("Heute klären wir Schritt für Schritt: Was ist die Grundform? Was ist eine gebeugte Form? Und wie helfen haben und sein beim Erzählen von früher?", { size: 21, color: "333333" }), { after: 60 }));

// Übung 1
ab.push(...task(1, "Wörterbuch-Detektiv", "Diese Formen findest du nicht im Wörterbuch. Schreibe die Grundform (den Infinitiv) daneben."));
{
  const w = [1500, 3603, 1500, 3603];
  const rows = [];
  for (let i = 0; i < 5; i++) rows.push([u1[i][0], "", u1[i + 5][0], ""]);
  ab.push(grid(w, ["gebeugt", "Infinitiv", "gebeugt", "Infinitiv"], rows, { fonts: ["Cambria", null, "Cambria", null], bold: [true, false, true, false], h: 520 }));
}

// Übung 2
ab.push(...task(2, "Wer gehört zu welcher Familie?", "a) Sortiere die Wörter in die Tabelle. Vorsicht: Einige Wörter gehören zu keiner der beiden Familien!"));
ab.push(box([p(u2words.flatMap((wd, i) => [t(wd, { font: "Cambria", size: 26, bold: true, color: NAVY }), t(i < u2words.length - 1 ? "     " : "", { size: 26 })]), { align: AlignmentType.CENTER, after: 0 })], LIGHT));
ab.push(gap(80));
{
  const w = [3402, 3402, 3402];
  ab.push(table(w, [
    row([hdr("Familie haben", w[0], C.hv, 22), hdr("Familie sein", w[1], C.hv, 22), hdr("keine von beiden", w[2], MUTED, 22)], 440),
    ...[0, 1].map(() => row(w.map((x) => tc("", x)), 1050)),
  ]));
}
ab.push(p(t("b) Unterstreiche die Form von haben oder sein. Schreibe den Familiennamen (Infinitiv) dahinter.", { size: 21 }), { before: 120, after: 80, keepNext: true }));
ab.push(grid([6406, 3800], ["Satz", "Familie"], u2sent.map(([s]) => [s, ""]), { fonts: ["Cambria", null], h: 480 }));

// Übung 3
ab.push(br(), ...task(3, "Die Familien-Tabelle", "Fülle die Tabelle aus. Präsens = so sagt man es jetzt. Präteritum = so erzählt man es von früher."));
ab.push((() => {
  const w = [1806, 2100, 2100, 2100, 2100];
  const rows = persons.map((ps, i) => [ps, i === 0 ? "habe" : "", i === 0 ? "hatte" : "", i === 0 ? "bin" : "", i === 0 ? "war" : ""]);
  const spanHdr = (txt, wsum) => cell(p(t(txt, { bold: true, color: "FFFFFF", size: 22 }), { align: AlignmentType.CENTER, after: 0 }), { w: wsum, fill: C.hv, span: 2 });
  return table(w, [
    row([hdr("", w[0]), spanHdr("Familie haben", w[1] + w[2]), spanHdr("Familie sein", w[3] + w[4])], 420),
    row([hdr("Person", w[0]), hdr("Präsens", w[1], "24306E"), hdr("Präteritum", w[2], "24306E"), hdr("Präsens", w[3], "24306E"), hdr("Präteritum", w[4], "24306E")], 400),
    ...rows.map((r) => row(r.map((v, i) => tc(v, w[i], { bold: i === 0 || v !== "", font: i === 0 ? null : "Cambria", color: i === 0 ? NAVY : MUTED, fill: i === 0 ? LIGHT : undefined, align: i === 0 ? undefined : AlignmentType.CENTER })), 520)),
  ]);
})());

ab.push(gap(60));
ab.push(box([
  p([t("Würfelspiel zu zweit  ", { font: "Cambria", bold: true, size: 24, color: NAVY }), t("(Material: 2 Würfel)", { size: 19, color: MUTED })], { after: 60 }),
  p([S("Würfel 1 = Person:  ", { bold: true }), S("1 ich · 2 du · 3 er/sie/es · 4 wir · 5 ihr · 6 sie")], { after: 30 }),
  p([S("Würfel 2 = Zeit:  ", { bold: true }), S("gerade Zahl → Präsens · ungerade Zahl → Präteritum")], { after: 30 }),
  p([S("Familie:  ", { bold: true }), S("abwechselnd haben und sein. Dein Partner sagt die Form, du prüfst mit der Tabelle. Richtig = 1 Punkt.")], { after: 30 }),
  p([S("Beispiel: ", { italics: true, color: MUTED }), S("4 und 3, Familie sein → „wir waren“", { italics: true, color: MUTED })], { after: 0 }),
], LIGHT));

// Übung 4
ab.push(...task(4, "Gestern war alles anders", "Setze die richtige Form im Präteritum ein. In Klammern steht die Familie."));
ab.push(grid([7406, 1300, 1500], ["Satz", "Familie", "Form"], u4.map(([s, f]) => [s, f, ""]), { fonts: ["Cambria", null, null], align: [null, AlignmentType.CENTER, null], colors: [null, MUTED, null], h: 440 }));

// Übung 5
ab.push(br(), ...task(5, "Das Partizip II", "Bilde das Partizip II. Test: Passt „ich habe …“ oder „ich bin …“ davor? Markiere ge- und die Endung gelb."));
{
  const w = [1700, 3403, 1700, 3403];
  const rows = [];
  for (let i = 0; i < 5; i++) rows.push([u5[i][0], "", u5[i + 5][0], ""]);
  ab.push(grid(w, ["Infinitiv", "Partizip II", "Infinitiv", "Partizip II"], rows, { fonts: ["Cambria", null, "Cambria", null], bold: [true, false, true, false], h: 500, headFill: [C.inf, C.p2, C.inf, C.p2] }));
  ab.push(p([S("Was fällt dir auf? Die linke Spalte endet auf ", {}), S("ge-…-____", { bold: true }), S(", die rechte Spalte auf ", {}), S("ge-…-____", { bold: true }), S(".")], { before: 80, after: 40 }));
}

// Übung 6
ab.push(...task(6, "Mein Wochenende: haben oder sein?", "Setze habe, hat, bin, ist oder sind ein. Kreise danach das Partizip II gelb ein. Bewegung von A nach B oder Veränderung? → sein!"));
ab.push(box(u6.map(([a, , c], i) => p([S(`${i + 1}  `, { bold: true, color: C.hv }), N(a), N("________"), N(c)], { after: 90 })), LIGHT));

// Übung 7
ab.push(...task(7, "Helfer oder nicht?", "Ist haben / sein hier ein Hilfsverb? Tipp: Steht am Satzende ein Partizip II, dann ist es ein Hilfsverb."));
ab.push(grid([6206, 2000, 2000], ["Satz", "Hilfsverb", "kein Hilfsverb"], u7.map(([s]) => [s, "☐", "☐"]), { fonts: ["Cambria", null, null], align: [null, AlignmentType.CENTER, AlignmentType.CENTER], h: 420 }));

// Merksatz-Hilfe
ab.push(br(), kicker("Hilfe für dein Heft"), h1("Merksätze zum Ergänzen"),
  p(t("Ergänze die Lücken mit den Wörtern von der Tafel. Klebe die Kästen danach in dein Heft.", { size: 21 }), { after: 120 }));
const mh = [
  ["Infinitiv", "Der Infinitiv ist die ______________ eines Verbs. So steht das Verb im ______________. Er endet auf -en oder -____ (spielen, basteln)."],
  ["Personalform", "Die Personalform (gebeugte Form) passt sich der ______________ an: ich spiel__, du spiel__, er spiel__."],
  ["Verbfamilie", "Jede gebeugte Form gehört zu einer Verbfamilie. Den Familiennamen finde ich mit der Frage: Wie heißt die ______________? habe → ______________, bist → ______________"],
  ["haben und sein im Präteritum", "Im Präteritum heißt es: ich ______________ (haben) und ich ______________ (sein). Auch „war“ gehört zur Familie ______________!"],
  ["Partizip II", "Das Partizip II verändert sich ______________. Man bildet es meist mit ge-…-____ (gespielt) oder ge-…-____ (gelaufen). Test: Passt „ich habe …“ oder „ich bin …“ davor?"],
  ["Perfekt mit haben oder sein", "Perfekt = haben oder sein im ______________ + ______________. Mit sein: Verben der ______________ von A nach B und der ______________ (einschlafen). Die meisten anderen Verben bilden das Perfekt mit ______________."],
  ["Hilfsverb oder nicht?", "haben und sein sind nur dann Hilfsverben, wenn am Satzende ein ______________ steht. Ich habe Hunger. → kein Hilfsverb. Ich habe gegessen. → ______________"],
];
mh.forEach(([title, text], i) => {
  ab.push(table([W], [row([cell([
    p([t(`Merksatz ${i + 1} · `, { bold: true, color: C.hv, size: 19 }), t(title, { bold: true, color: NAVY, size: 21, font: "Cambria" })], { after: 50 }),
    p(t(text, { size: 21 }), { after: 0, line: 320 }),
  ], { w: W, borders: allBorders(dashed), m: 110, ml: 160 })])]), gap(90));
});

// Exit-Tickets
ab.push(br(), kicker("Exit-Ticket · ausschneiden"));
function ticket2() {
  const w = [3900, 1700, 2300, 2006];
  return table([W], [row([cell([
    p([t("Exit-Ticket  ", { font: "Cambria", bold: true, size: 24, color: NAVY }), t("Name: ___________________", { size: 19, color: MUTED })], { after: 40 }),
    p(t("Unterstreiche die Form von haben oder sein und fülle die Tabelle aus.", { size: 19 }), { after: 60 }),
    table(w, [
      row(["Satz", "Familie", "Präsens / Präteritum?", "Hilfsverb?"].map((h, i) => hdr(h, w[i], NAVY, 17)), 340),
      ...exit.map(([s]) => row([tc(s, w[0], { font: "Cambria", size: 20 }), tc("", w[1]), tc("", w[2]), tc("ja ☐    nein ☐", w[3], { size: 19, align: AlignmentType.CENTER })], 380)),
    ]),
  ], { w: W, borders: allBorders(dashed), m: 120, ml: 160 })])]);
}
for (let i = 0; i < 4; i++) ab.push(ticket2(), gap(i < 3 ? 120 : 0));

// ================================================================
// HANDREICHUNG
// ================================================================
const hr = [];
hr.push(kicker("Handreichung für die Lehrkraft · Doppelstunde (90 Min.)"), h1("Verbformen verstehen"),
  p(t("Ziel: Die Schülerinnen und Schüler erkennen gebeugte Formen von haben und sein (Präsens und Präteritum) und führen sie auf den Infinitiv zurück. Sie bilden das Partizip II, wählen im Perfekt haben oder sein und unterscheiden Hilfsverb von Vollverb. Modalverben und werden/Futur folgen in der nächsten Stunde.", { size: 21 }), { after: 120 }));

hr.push(h2("", "Stundenverlauf", 80));
{
  const w = [700, 2400, 5106, 2000];
  const rows = [
    ["5′", "Einstieg", "Wörterbuch-Rätsel: „Tim sucht das Wort ging im Wörterbuch und findet es nicht. Warum?“", "Wörterbuch"],
    ["15′", "Infinitiv & Personalform", "Tafelbild 1 entwickeln, M1 + M2 ins Heft. Übung 1 (EA), kurze Kontrolle.", "AB Übung 1"],
    ["15′", "Verbfamilien haben/sein (Präsens)", "Tafelbild 2 (Familienbilder), M3 ins Heft. Übung 2 (EA/PA).", "AB Übung 2"],
    ["15′", "haben/sein im Präteritum", "Tafelbild 2 um Spalte „früher“ erweitern, M4. Übung 3 + Würfelspiel (PA). Übung 4 als Zusatz/HA.", "AB Übung 3–4, Würfel"],
    ["3′", "Bewegungspause", "Formen rufen – Hand links = haben, Hand rechts = sein.", "Wortliste unten"],
    ["30′", "Partizip II & Perfekt mit haben/sein; Hilfsverb vs. Vollverb", "Wochenend-Erzählungen → Tafelbild 3, M5. Übung 5. Sätze nach haben/sein sortieren, Vorführung, M6. Kontrastpaar „Ich habe Hunger / Ich habe gegessen“, M7. Übung 6 + 7.", "AB Übung 5–7"],
    ["7′", "Abschluss", "Gesamttafelbild ergänzen (Paare einzeichnen). Exit-Ticket.", "Exit-Tickets"],
  ];
  hr.push(grid(w, ["Zeit", "Phase", "Ablauf / Impulse", "Material"], rows, { bold: [true, true, false, false], colors: [C.hv, NAVY, null, MUTED], h: 400 }));
}
hr.push(p([S("Farbcode an der Tafel (wie in der Bastelstunde): ", { bold: true }), t("haben/sein rot", { bold: true, color: C.hv, size: 21 }), S(" · "), t("Partizip II gelb", { bold: true, color: "C98A1E", size: 21 }), S(" · "), t("Infinitiv blau", { bold: true, color: C.inf, size: 21 })], { before: 100, after: 0 }));

// Phase 1
hr.push(br(), h2("1", "Infinitiv und Personalform"));
hr.push(p([S("Impulse: ", { bold: true }), S("„Wie heißt das Verb, wenn es ganz ‚nackt‘ ist – ohne ich, du, er?“ · „Was verändert sich, wenn ich ich, du, er davorsetze?“")], { after: 100 }));
hr.push(tafel("Tafelbild 1", [
  table([4700, 800, 4300], [
    row([cell(p(N("gebeugte Form", { bold: true, color: NAVY }), { after: 0 }), { w: 4700, borders: noBorders }), cell(p(t("")), { w: 800, borders: noBorders }), cell(p(N("Grundform = Infinitiv", { bold: true, color: C.inf }), { after: 0 }), { w: 4300, borders: noBorders })]),
    ...[["ging", "gehen"], ["isst", "essen"], ["hat", "haben"], ["war", "sein"]].map(([a, b]) => row([cell(p(N(a), { after: 0 }), { w: 4700, borders: noBorders, m: 20 }), cell(p(N("→"), { after: 0, align: AlignmentType.CENTER }), { w: 800, borders: noBorders, m: 20 }), cell(p(B(b), { after: 0 }), { w: 4300, borders: noBorders, m: 20 })])),
  ]),
  p(t(""), { after: 60 }),
  p([N("ich spiel", {}), N("e", { bold: true, u: true }), N("   ·   du spiel"), N("st", { bold: true, u: true }), N("   ·   er spiel"), N("t", { bold: true, u: true }), N("   →  die Personalform passt sich der Person an", { italics: true, color: MUTED })], { after: 0 }),
]));
hr.push(gap(100), merk(1, "Infinitiv", [[S("Der Infinitiv ist die "), S("Grundform", { bold: true }), S(" eines Verbs. So steht das Verb im "), S("Wörterbuch", { bold: true }), S(". Er endet auf -en oder -n (spielen, basteln).")]]));
hr.push(gap(80), merk(2, "Personalform", [[S("Die Personalform (gebeugte Form) passt sich der "), S("Person", { bold: true }), S(" an: ich spiele, du spielst, er spielt.")]]));
hr.push(p([S("Hinweis: ", { bold: true, color: C.hv }), S("Die Frage „Wie heißt die Grundform?“ ab jetzt bei jeder markierten Verbform stellen – genau diese Routine fehlte, als „habe“ nicht erkannt wurde.")], { before: 100 }));

// Phase 2
hr.push(h2("2", "Die Familien haben und sein"));
hr.push(p([S("Impulse: ", { bold: true }), S("„Wer gehört alles zur Familie haben?“ (Klasse konjugiert gemeinsam) · „Sieht bin aus wie sein? Woher weiß ich trotzdem, dass es dazugehört?“ → nur über die Frage nach der Grundform.")], { after: 100 }));
const famTable = (withPast) => {
  const w = withPast ? [1500, 2100, 2100, 2100, 2100] : [1500, 2100, 2100];
  const headRow = withPast
    ? [cell(p(t("")), { w: w[0], borders: noBorders }), cell(p(R("Familie haben"), { after: 0, align: AlignmentType.CENTER }), { w: w[1] + w[2], span: 2, borders: noBorders }), cell(p(R("Familie sein"), { after: 0, align: AlignmentType.CENTER }), { w: w[3] + w[4], span: 2, borders: noBorders })]
    : [cell(p(t("")), { w: w[0], borders: noBorders }), cell(p(R("Familie haben"), { after: 0, align: AlignmentType.CENTER }), { w: w[1], borders: noBorders }), cell(p(R("Familie sein"), { after: 0, align: AlignmentType.CENTER }), { w: w[2], borders: noBorders })];
  const sub = withPast ? [["", ""], ["jetzt (Präsens)", ""], ["früher (Präteritum)", ""], ["jetzt (Präsens)", ""], ["früher (Präteritum)", ""]] : null;
  const rows = [row(headRow)];
  if (withPast) rows.push(row(sub.map(([s], i) => cell(p(t(s, { size: 17, italics: true, color: MUTED }), { after: 0, align: AlignmentType.CENTER }), { w: w[i], borders: noBorders, m: 10 }))));
  persons.forEach((ps, i) => {
    const vals = withPast ? [konj.hp[i], konj.hv[i], konj.sp[i], konj.sv[i]] : [konj.hp[i], konj.sp[i]];
    rows.push(row([cell(p(N(ps, { color: MUTED }), { after: 0 }), { w: w[0], borders: noBorders, m: 15 }), ...vals.map((v, j) => cell(p(N(v), { after: 0, align: AlignmentType.CENTER }), { w: w[j + 1], borders: noBorders, m: 15 }))]));
  });
  return table(w, rows);
};
hr.push(tafel("Tafelbild 2 · Familienbilder (erst nur Präsens)", [famTable(false)]));
hr.push(gap(100), merk(3, "Verbfamilie", [[S("Jede gebeugte Form gehört zu einer Verbfamilie. Den Familiennamen finde ich mit der Frage: "), S("Wie heißt die Grundform?", { bold: true }), S("  habe → haben, bist → sein")]]));

// Phase 3
hr.push(br(), h2("3", "haben und sein im Präteritum"));
hr.push(p([S("Impuls: ", { bold: true }), S("„Heute bin ich müde. Gestern … ?“ – „Heute habe ich Zeit. Gestern … ?“ Die Formen werden als neue Spalte „früher“ ins Familienbild geschrieben. Betonen: war sieht noch weniger nach sein aus als ist.")], { after: 100 }));
hr.push(tafel("Tafelbild 2 · erweitert", [famTable(true)]));
hr.push(gap(100), merk(4, "haben und sein im Präteritum", [[S("Im Präteritum heißt es: ich "), S("hatte", { bold: true }), S(" (haben) und ich "), S("war", { bold: true }), S(" (sein). Auch „war“ gehört zur Familie sein!")]]));
hr.push(p([S("Bewegungspause – Wortliste: ", { bold: true }), S("habe · war · bist · hatten · seid · hattest · ist · hat · waren · bin · habt · wart · hatte · sind")], { before: 120 }));
hr.push(p([S("Hinweis: ", { bold: true, color: C.hv }), S("Den Begriff Plusquamperfekt hier bewusst nicht einführen. Die Präteritumformen werden nur als Formen gesichert – so ist später der Unterschied habe/hatte gebacken vorbereitet.")], { before: 80 }));

// Phase 4
hr.push(h2("4", "Partizip II, Perfekt mit haben oder sein, Hilfsverb"));
hr.push(p([S("Ablauf: ", { bold: true }), S("Die Klasse erzählt vom Wochenende. Sätze an der Tafel sammeln, Partizip II gelb, Form von haben/sein rot markieren. Dann nach haben / sein sortieren und fragen: „Was haben die sein-Sätze gemeinsam?“ Vorführung: Ein Kind geht von der Tür zum Fenster („Sie ist gegangen.“), ein anderes sitzt und liest („Er hat gelesen.“).")], { after: 100 }));
hr.push(tafel("Tafelbild 3", [
  table([4900, 4900], [
    row([cell(p(N("mit haben", { bold: true, color: NAVY }), { after: 0 }), { w: 4900, borders: noBorders }), cell(p(N("mit sein", { bold: true, color: NAVY }), { after: 0 }), { w: 4900, borders: noBorders })]),
    ...[[["Ich ", "habe", " Fußball ", "gespielt", "."], ["Ich ", "bin", " ins Kino ", "gegangen", "."]], [["Wir ", "haben", " Pizza ", "gegessen", "."], ["Wir ", "sind", " nach Hamburg ", "gefahren", "."]], [["Sie ", "hat", " ein Buch ", "gelesen", "."], ["Er ", "ist", " spät ", "eingeschlafen", "."]]].map(([a, b]) => row([a, b].map((s) => cell(p([N(s[0]), R(s[1]), N(s[2]), Y(s[3]), N(s[4])], { after: 0 }), { w: 4900, borders: noBorders, m: 20 })))),
    row([cell(p(t(""), { after: 0 }), { w: 4900, borders: noBorders }), cell(p(t("→ Bewegung von A nach B / Veränderung", { italics: true, color: MUTED, size: 20 }), { after: 0 }), { w: 4900, borders: noBorders })]),
  ]),
  p(t(""), { after: 60 }),
  p([N("Ich "), N("habe", { bold: true }), N(" Hunger.  → kein Hilfsverb      ·      Ich "), R("habe"), N(" gegessen.  → "), R("Hilfsverb")], { after: 0 }),
]));
hr.push(gap(100), merk(5, "Partizip II", [[S("Das Partizip II verändert sich "), S("nie", { bold: true }), S(". Man bildet es meist mit ge-…-t (gespielt) oder ge-…-en (gelaufen). Test: Passt „ich habe …“ oder „ich bin …“ davor?")]]));
hr.push(gap(80), merk(6, "Perfekt mit haben oder sein", [[S("Perfekt = "), S("haben oder sein im Präsens + Partizip II", { bold: true }), S(". Mit sein: Verben der "), S("Bewegung", { bold: true }), S(" von A nach B (gehen, fahren) und der "), S("Veränderung", { bold: true }), S(" (einschlafen, aufwachen). Die meisten anderen Verben bilden das Perfekt mit "), S("haben", { bold: true }), S(".")]]));
hr.push(gap(80), merk(7, "Hilfsverb oder nicht?", [[S("haben und sein sind nur dann "), S("Hilfsverben", { bold: true }), S(", wenn am Satzende ein "), S("Partizip II", { bold: true }), S(" steht. Ich habe Hunger. → kein Hilfsverb. Ich habe gegessen. → Hilfsverb.")]]));
hr.push(p([S("Stolperstein: ", { bold: true, color: C.hv }), S("Im Perfekt steht das Hilfsverb im Präsens (habe), trotzdem erzählt der Satz von früher. Deshalb im Exit-Ticket nach der Form des Hilfsverbs fragen („Präsens oder Präteritum?“), nicht nach der Zeit des Geschehens. Falls „ich bin gesessen/gestanden“ auftaucht: als süddeutsche/österreichische Variante einordnen.")], { before: 100 }));

// Phase 5
hr.push(h2("5", "Abschluss: Gesamttafelbild"));
hr.push(p(t("Die Tafel wird in zwei Spalten zusammengefasst. Zum Schluss werden die Paare mit Pfeilen verbunden.", { size: 21 }), { after: 100 }));
hr.push(tafel("Gesamttafelbild", [
  table([4900, 4900], [
    row([cell(p(N("passt sich der Person an", { bold: true, color: NAVY }), { after: 0, align: AlignmentType.CENTER }), { w: 4900, borders: noBorders }), cell(p(N("bleibt immer gleich", { bold: true, color: NAVY }), { after: 0, align: AlignmentType.CENTER }), { w: 4900, borders: noBorders })]),
    row([cell([p(R("Familie haben"), { after: 0, align: AlignmentType.CENTER }), p(N("habe, hast, hat … / hatte, hattest …"), { after: 80, align: AlignmentType.CENTER }), p(R("Familie sein"), { after: 0, align: AlignmentType.CENTER }), p(N("bin, bist, ist … / war, warst …"), { after: 0, align: AlignmentType.CENTER })], { w: 4900, borders: noBorders }),
      cell([p(B("Infinitiv"), { after: 0, align: AlignmentType.CENTER }), p(N("spielen, gehen, haben, sein"), { after: 80, align: AlignmentType.CENTER }), p(Y("Partizip II"), { after: 0, align: AlignmentType.CENTER }), p(N("gespielt, gegangen"), { after: 0, align: AlignmentType.CENTER })], { w: 4900, borders: noBorders })]),
  ]),
  p(t(""), { after: 40 }),
  p([R("habe / bin"), N("  + "), Y("Partizip II"), N("  =  Perfekt:  Ich "), R("bin"), N(" nach Hause "), Y("gegangen"), N(".")], { after: 0, align: AlignmentType.CENTER }),
]));
hr.push(p([S("Ausblick nächste Stunde: ", { bold: true }), S("Modalverben + Infinitiv, werden als dritte Hilfsverb-Familie (Futur). Danach kann der Zeitformen-Baukasten als Anwendung eingesetzt werden.")], { before: 120 }));

// Lösungen
hr.push(br(), kicker("Lösungen"), h1("Lösungen zum Arbeitsblatt"));
hr.push(h3("Übung 1"));
hr.push(p(t(u1.map(([a, b]) => `${a} → ${b}`).join("   ·   "), { size: 21 }), { after: 60 }));
hr.push(h3("Übung 2"));
hr.push(p([S("a) haben: ", { bold: true, color: C.hv }), S("habe, hat, hast, habt, haben   "), S("sein: ", { bold: true, color: C.hv }), S("ist, sind, seid, bin, bist   "), S("keine: ", { bold: true, color: MUTED }), S("Hand, seit, Hafen, hart")], { after: 40 }));
hr.push(p([S("b) "), S(u2sent.map(([, f, fam]) => `${f} → ${fam}`).join("  ·  "))], { after: 60 }));
hr.push(h3("Übung 3"));
hr.push(grid([1806, 2100, 2100, 2100, 2100], ["Person", "haben Präsens", "haben Präteritum", "sein Präsens", "sein Präteritum"], persons.map((ps, i) => [ps, konj.hp[i], konj.hv[i], konj.sp[i], konj.sv[i]]), { h: 340, fonts: [null, "Cambria", "Cambria", "Cambria", "Cambria"] }));
hr.push(h3("Übung 4"));
hr.push(p(t(u4.map(([, , l], i) => `${i + 1} ${l}`).join("   ·   "), { size: 21 }), { after: 60 }));
hr.push(h3("Übung 5"));
hr.push(p(t(u5.map(([a, b]) => `${a} → ${b}`).join("  ·  "), { size: 21 }), { after: 40 }));
hr.push(p(t("Linke Spalte (regelmäßig): ge-…-t · rechte Spalte (unregelmäßig): ge-…-en", { size: 21 }), { after: 60 }));
hr.push(h3("Übung 6"));
hr.push(p(t(u6.map(([, h], i) => `${i + 1} ${h}`).join("   ·   "), { size: 21 }), { after: 40 }));
hr.push(p(t("Partizipien: aufgewacht, gebacken, gefahren, gespielt, geschossen, gelaufen, gelesen, eingeschlafen", { size: 21 }), { after: 60 }));
hr.push(h3("Übung 7"));
hr.push(p(t(u7.map(([s, l], i) => `${i + 1} ${l}`).join("   ·   "), { size: 21 }), { after: 60 }));
hr.push(h3("Exit-Ticket"));
hr.push(grid([3900, 1700, 2300, 2306], ["Satz", "Familie", "Form", "Hilfsverb?"], exit.map(([s, , fam, z, hv]) => [s, fam, z, hv]), { h: 340, fonts: ["Cambria", null, null, null] }));
hr.push(p(t("Auswertung: Wer bei Satz 1 und 3 „ja“ ankreuzt, verwechselt noch Hilfsverb und Vollverb (M7). Wer bei Satz 2 und 4 „Präteritum“ schreibt, verwechselt die Form des Hilfsverbs mit der Zeit des Geschehens.", { size: 20, italics: true, color: MUTED }), { before: 80 }));

// ================================================================
const pageProps = { page: { size: { width: 11906, height: 16838 }, margin: { top: 900, bottom: 800, left: 850, right: 850, header: 400, footer: 400 } } };
const styles = { default: { document: { run: { font: "Calibri", size: 22 } } } };
const docAB = new Document({ styles, sections: [{ properties: pageProps, headers: { default: header("Verbformen verstehen") }, footers: { default: footer("Verbformen verstehen") }, children: ab }] });
const docHR = new Document({ styles, sections: [{ properties: pageProps, headers: { default: header("Verbformen verstehen", false) }, footers: { default: footer("Handreichung") }, children: hr }] });
Packer.toBuffer(docAB).then((b) => fs.writeFileSync("Arbeitsblatt_Verbformen_verstehen.docx", b));
Packer.toBuffer(docHR).then((b) => fs.writeFileSync("Handreichung_Verbformen_verstehen.docx", b));
