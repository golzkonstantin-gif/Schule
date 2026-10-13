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
// Leistungskontrolle Verben und Zeitformen (30 Min.) – Gruppe A und B + Erwartungshorizont
// ================================================================
const PW = 11906 - 1600;
const gray = "9AA3B5";
const s = (x, o = {}) => t(x, Object.assign({ size: 22 }, o));
const sb = (x, o = {}) => s(x, Object.assign({ bold: true }, o));
const bd = allBorders(solid("7A869E", 8));
const th = (x, w) => cell(p(t(x, { bold: true, size: 20, color: "FFFFFF" }), { after: 0, align: AlignmentType.CENTER }), { w, fill: NAVY, borders: bd, m: 50 });
const td = (x, w, o = {}) => cell(p(Array.isArray(x) ? x : t(x, { size: 22, font: o.font || "Calibri", italics: o.it, color: o.color, bold: o.bold }), { after: 0, align: o.align }), { w, borders: bd, fill: o.fill, m: 50, ml: 90 });
const aufgabe = (nr, title, pts, intro) => [
  p([t(`Aufgabe ${nr}  `, { font: "Cambria", bold: true, size: 26, color: C.hv }), t(title, { font: "Cambria", bold: true, size: 26, color: NAVY }), t(`     ___ / ${pts} P.`, { size: 20, color: MUTED })], { before: 280, after: 60, keepNext: true }),
  p(intro, { after: 100, keepNext: true }),
];
const SOLC = "1E7B4A";
const lineRow = (label, ans = "") => table([2400, PW - 2400], [row([cell(p(t(label, { bold: true, size: 21, color: NAVY }), { after: 0 }), { w: 2400, borders: noBorders, m: 40, valign: VerticalAlign.BOTTOM }), cell(p(t(ans, { size: 22, color: SOLC, italics: true }), { after: 0 }), { valign: VerticalAlign.BOTTOM,  w: PW - 2400, borders: { top: none, left: none, right: none, bottom: solid("7A869E", 6) } })], 620, HeightRule.EXACT)]);

const GROUPS = {
  A: {
    t1: [["heben", "", ""], ["", "er läuft", ""], ["", "", "gewonnen"], ["trainieren", "", ""], ["", "er fällt", ""], ["", "", "gekämpft"]],
    t1sol: [["heben", "er hebt", "gehoben"], ["laufen", "er läuft", "gelaufen"], ["gewinnen", "er gewinnt", "gewonnen"], ["trainieren", "er trainiert", "trainiert"], ["fallen", "er fällt", "gefallen"], ["kämpfen", "er kämpft", "gekämpft"]],
    t2: [[["Du ", ["hast"], " den Kampf ", ["gewonnen"], "."], ["2.", "Singular", "Perfekt"]],
      [["Die Trainerin ", ["lobte"], " das Team."], ["3.", "Singular", "Präteritum"]],
      [["Ihr ", ["werdet"], " morgen früh ", ["starten"], "."], ["2.", "Plural", "Futur I"]],
      [["Ich ", ["hatte"], " mich lange ", ["aufgewärmt"], "."], ["1.", "Singular", "Plusquamperfekt"]],
      [["Die Judoka ", ["kämpfen"], " heute im Finale."], ["3.", "Plural", "Präsens"]],
      [["Wir ", ["warfen"], " viele Tore."], ["1.", "Plural", "Präteritum"]]],
    t3: [["a", "Ich gewinne den Kampf.", ["Ich gewann den Kampf.", "Ich habe den Kampf gewonnen.", "Ich hatte den Kampf gewonnen.", "Ich werde den Kampf gewinnen."]],
      ["b", "Wir fahren zum Turnier.", ["Wir fuhren zum Turnier.", "Wir sind zum Turnier gefahren.", "Wir waren zum Turnier gefahren.", "Wir werden zum Turnier fahren."]]],
    t4: ["Am Samstag bin ich früh aufgestanden.", "Mein Trainer hatte alles vorbereitet.", "Im ersten Kampf habe ich gut gekämpft.", "Danach trank ich viel Wasser.", "Nächste Woche werde ich wieder antreten.", "Meine Eltern haben mich laut angefeuert."],
    t4sol: "bin – aufgestanden · hatte – vorbereitet · habe – gekämpft · (trank: kein Hilfsverb) · werde – antreten (Infinitiv) · haben – angefeuert",
    zusatz: ["Aufgabe 3a", "Ich werde den Kampf gewonnen haben."],
    t4mark: [{ bin: "h", aufgestanden: "p" }, { hatte: "h", vorbereitet: "p" }, { habe: "h", "gekämpft": "p" }, {}, { werde: "h", antreten: "i" }, { haben: "h", angefeuert: "p" }],
  },
  B: {
    t1: [["halten", "", ""], ["", "er trifft", ""], ["", "", "gestoßen"], ["spielen", "", ""], ["", "er verliert", ""], ["", "", "gesprungen"]],
    t1sol: [["halten", "er hält", "gehalten"], ["treffen", "er trifft", "getroffen"], ["stoßen", "er stößt", "gestoßen"], ["spielen", "er spielt", "gespielt"], ["verlieren", "er verliert", "verloren"], ["springen", "er springt", "gesprungen"]],
    t2: [[["Ihr ", ["habt"], " das Spiel ", ["verloren"], "."], ["2.", "Plural", "Perfekt"]],
      [["Der Trainer ", ["rief"], " sehr laut."], ["3.", "Singular", "Präteritum"]],
      [["Du ", ["wirst"], " bald ", ["gewinnen"], "."], ["2.", "Singular", "Futur I"]],
      [["Wir ", ["hatten"], " wochenlang ", ["geübt"], "."], ["1.", "Plural", "Plusquamperfekt"]],
      [["Die Gewichtheberin ", ["hebt"], " heute 70 Kilo."], ["3.", "Singular", "Präsens"]],
      [["Ich ", ["lief"], " sehr schnell."], ["1.", "Singular", "Präteritum"]]],
    t3: [["a", "Du wirfst den Ball.", ["Du warfst den Ball.", "Du hast den Ball geworfen.", "Du hattest den Ball geworfen.", "Du wirst den Ball werfen."]],
      ["b", "Sie läuft ins Ziel.", ["Sie lief ins Ziel.", "Sie ist ins Ziel gelaufen.", "Sie war ins Ziel gelaufen.", "Sie wird ins Ziel laufen."]]],
    t4: ["Am Sonntag sind wir früh losgefahren.", "Die Trainerin hatte uns gut vorbereitet.", "In der ersten Halbzeit haben wir drei Tore geworfen.", "In der Pause aß ich eine Banane.", "Im Rückspiel werden wir noch stärker spielen.", "Die Fans haben laut gesungen."],
    t4sol: "sind – losgefahren · hatte – vorbereitet · haben – geworfen · (aß: kein Hilfsverb) · werden – spielen (Infinitiv) · haben – gesungen",
    zusatz: ["Aufgabe 3a", "Du wirst den Ball geworfen haben."],
    t4mark: [{ sind: "h", losgefahren: "p" }, { hatte: "h", vorbereitet: "p" }, { haben: "h", geworfen: "p" }, {}, { werden: "h", spielen: "i" }, { haben: "h", gesungen: "p" }],
  },
};
const TENSES = ["Präteritum", "Perfekt", "Plusquamperfekt", "Futur I"];
const MC = { h: C.hv, p: "C98A1E", i: C.inf };
const markRuns = (sent, m) => sent.split(/(\s+|[.,!?])/).filter((x) => x !== "").map((w) => (m[w] ? t(w, { font: "Cambria", size: 26, bold: true, u: true, color: MC[m[w]] }) : t(w, { font: "Cambria", size: 26 })));
const verbRuns = (parts) => parts.map((x) => (Array.isArray(x) ? t(x[0], { size: 22, bold: true, font: "Cambria" }) : t(x, { size: 22, font: "Cambria" })));

function test(g, sol = false) {
  const G = GROUPS[g];
  const SV = (x) => ({ it: true, color: SOLC });
  const d = [];
  d.push(table([PW - 2600, 2600], [row([
    cell([p([t("Leistungskontrolle · Deutsch · Klasse 7A", { size: 20, color: MUTED, bold: true })], { after: 20 }), p(t("Verben und Zeitformen", { font: "Cambria", bold: true, size: 34, color: NAVY }), { after: 60 }), p([s("Name: ______________________________     Datum: 14.10.")], { after: 0 })], { w: PW - 2600, borders: noBorders, m: 0 }),
    cell([p(t(`Gruppe ${g}`, { font: "Cambria", bold: true, size: 36, color: C.hv }), { after: 60, align: AlignmentType.CENTER }), ...(sol ? [p(t("LÖSUNG", { bold: true, size: 26, color: SOLC }), { after: 40, align: AlignmentType.CENTER })] : []), p(s("Punkte: ____ / 25"), { after: 40, align: AlignmentType.CENTER }), p(s("Zusatz: ____ / 2"), { after: 40, align: AlignmentType.CENTER }), p(s("Note: ______"), { after: 0, align: AlignmentType.CENTER })], { w: 2600, borders: allBorders(solid(NAVY, 10)), m: 100 }),
  ])]));
  d.push(p([s("Lies jede Aufgabe genau. Schreibe sauber. Du hast 30 Minuten Zeit.", { italics: true, color: MUTED, size: 20 })], { before: 120, after: 0 }));

  // Aufgabe 1
  const W1 = [Math.floor(PW / 3), Math.floor(PW / 3), PW - 2 * Math.floor(PW / 3)];
  d.push(...aufgabe(1, "Infinitiv, finite Verbform, Partizip II", 6, [s("Ergänze die Tabelle. Die finite Verbform steht in der 3. Person Singular Präsens (er …). Die erste Zeile ist ein Beispiel.")]));
  d.push(table(W1, [
    row([th("Infinitiv", W1[0]), th("finite Verbform (er …)", W1[1]), th("Partizip II", W1[2])], 420),
    row([td("werfen", W1[0], { it: true, color: MUTED, fill: LIGHT }), td("er wirft", W1[1], { it: true, color: MUTED, fill: LIGHT }), td("geworfen", W1[2], { it: true, color: MUTED, fill: LIGHT })], 520, HeightRule.EXACT),
    ...G.t1.map((r, ri) => row(r.map((x, i) => (x || !sol ? td(x, W1[i], { font: "Cambria", bold: !!x }) : td(G.t1sol[ri][i], W1[i], SV()))), 600, HeightRule.EXACT)),
  ]));

  // Aufgabe 2
  const W2 = [PW - 5700, 1500, 1700, 2500];
  d.push(...aufgabe(2, "Person, Numerus und Tempus bestimmen", 6, [s("Bestimme für die "), sb("fett gedruckte"), s(" Verbform Person, Numerus und Tempus. Die erste Zeile ist ein Beispiel.")]));
  d.push(table(W2, [
    row([th("Satz", W2[0]), th("Person", W2[1]), th("Numerus", W2[2]), th("Tempus", W2[3])], 420),
    row([td(verbRuns(["Wir ", ["trainieren"], " jeden Tag."]), W2[0], { fill: LIGHT }), td("1. Person", W2[1], { it: true, color: MUTED, fill: LIGHT }), td("Plural", W2[2], { it: true, color: MUTED, fill: LIGHT }), td("Präsens", W2[3], { it: true, color: MUTED, fill: LIGHT })], 520, HeightRule.EXACT),
    ...G.t2.slice(0, 4).map(([parts, so]) => row([td(verbRuns(parts), W2[0]), td(sol ? so[0] + " Person" : "", W2[1], SV()), td(sol ? so[1] : "", W2[2], SV()), td(sol ? so[2] : "", W2[3], SV())], 620, HeightRule.EXACT)),
  ]));

  // Aufgabe 3 (auf der Rückseite)
  d.push(new Paragraph({ children: [new PageBreak()] }));
  d.push(...aufgabe(3, "Zeitformen bilden", 8, [s("Schreibe den Satz in den angegebenen Zeitformen. Achte auf das richtige Hilfsverb ("), sb("haben"), s(" oder "), sb("sein"), s(").")]));
  G.t3.forEach(([k, satz, sols]) => {
    d.push(p([sb(`${k})  `, { color: C.hv }), t("Präsens: ", { size: 22, color: MUTED }), t(satz, { font: "Cambria", size: 24, bold: true })], { before: 120, after: 20, keepNext: true }));
    TENSES.forEach((tn, i) => d.push(lineRow(tn + ":", sol ? sols[i] : "")));
  });

  // Aufgabe 4
  d.push(...aufgabe(4, "Hilfsverb und Partizip II bzw. Infinitiv unterstreichen", 5, [s("Unterstreiche in jedem Satz das "), sb("Hilfsverb rot", { color: C.hv }), s(" und das "), sb("Partizip II gelb", { color: "C98A1E" }), s(" bzw. den "), sb("Infinitiv blau", { color: C.inf }), s(". Achtung: Nicht jeder Satz hat ein Hilfsverb. (Ohne Farbstifte: Hilfsverb einfach, Partizip II/Infinitiv doppelt unterstreichen.)")]));
  G.t4.forEach((x, i) => d.push(p([t(`${i + 1}  `, { size: 20, bold: true, color: MUTED }), ...(sol ? markRuns(x, G.t4mark[i]) : [t(x, { font: "Cambria", size: 26 })])], { after: 0, line: 480 })));

  // Zusatz
  d.push(p([t("Zusatz  ", { font: "Cambria", bold: true, size: 26, color: C.hv }), t("Futur II", { font: "Cambria", bold: true, size: 26, color: NAVY }), t("     ___ / 2 P.", { size: 20, color: MUTED })], { before: 280, after: 60, keepNext: true }));
  d.push(p([s(`Bilde zum Satz aus ${G.zusatz[0]} das Futur II. Tipp: `), sb("werden + Partizip II + haben/sein"), s(".")], { after: 60, keepNext: true }));
  d.push(lineRow("Futur II:", sol ? G.zusatz[1] : ""));
  d.push(p(s("Fertig? Lies alles noch einmal durch und prüfe die Hilfsverben.", { italics: true, color: MUTED, size: 20 }), { before: 240, after: 0, align: AlignmentType.RIGHT }));
  return d;
}

function solution() {
  const d = [];
  d.push(kicker("Für die Lehrkraft · Leistungskontrolle · Verben und Zeitformen · 14.10. · 30 Minuten"), h1("Erwartungshorizont Gruppe A und B"));
  d.push(p([S("Gesamt: ", { bold: true }), S("25 Punkte + 2 Zusatzpunkte. Aufgabe 1: je richtige Form 0,5 P. (12 Formen) · Aufgabe 2: je richtige Angabe 0,5 P. (12 Angaben) · Aufgabe 3: je richtiger Satz 1 P. (bei richtigem Hilfsverb und falscher Präteritumform/Partizip 0,5 P.) · Aufgabe 4: je richtig markiertes Paar 1 P. (5 Paare; ein fälschlich markiertes Präteritum-Verb: −0,5 P.) · Zusatz: 2 P. (1 P. bei kleinem Fehler).")], { after: 100 }));
  ["A", "B"].forEach((g) => {
    const G = GROUPS[g];
    d.push(h3(`Gruppe ${g}`, C.hv));
    d.push(p([S("Aufgabe 1: ", { bold: true }), S(G.t1sol.map((r) => r.join(" – ")).join("  ·  "))], { after: 60 }));
    d.push(p([S("Aufgabe 2: ", { bold: true }), S(G.t2.slice(0, 4).map(([parts, sol], i) => `${i + 1} ${parts.map((x) => (Array.isArray(x) ? x[0] : "")).filter(Boolean).join(" … ")}: ${sol[0]} Person, ${sol[1]}, ${sol[2]}`).join("  ·  "))], { after: 60 }));
    d.push(p([S("Aufgabe 3: ", { bold: true }), S(G.t3.map(([k, , sols]) => `${k}) ` + sols.map((x, i) => `${TENSES[i]}: ${x}`).join(" / ")).join("   "))], { after: 60 }));
    d.push(p([S("Aufgabe 4: ", { bold: true }), S(G.t4sol)], { after: 60 }));
    d.push(p([S("Zusatz: ", { bold: true }), S(G.zusatz[1])], { after: 120 }));
  });
  d.push(h3("Notenschlüssel (Vorschlag – an den Schulschlüssel anpassen)"));
  const NK = [["1", "25 – 23,5"], ["2", "23 – 20"], ["3", "19,5 – 16,5"], ["4", "16 – 12,5"], ["5", "12 – 6"], ["6", "5,5 – 0"]];
  d.push(table([1800, 3000], [row([hdr("Note", 1800), hdr("Punkte (von 25)", 3000)]), ...NK.map(([n, pt]) => row([tc(n, 1800, { bold: true, align: AlignmentType.CENTER }), tc(pt, 3000, { align: AlignmentType.CENTER })], 340))]));
  d.push(p([S("Zusatzpunkte werden zur Gesamtpunktzahl addiert (höchstens bis 25).", { color: MUTED })], { before: 80, after: 0 }));
  return d;
}

const props = { page: { size: { width: 11906, height: 16838 }, margin: { top: 800, bottom: 700, left: 800, right: 800 } } };
const styles = { default: { document: { run: { font: "Calibri", size: 22 } } } };
const foot = (g) => new Footer({ children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ font: "Calibri", size: 16, color: MUTED, children: [`Leistungskontrolle Verben · Gruppe ${g} · Seite `, PageNumber.CURRENT] })] })] });
["A", "B"].forEach((g) => Packer.toBuffer(new Document({ styles, sections: [{ properties: props, footers: { default: foot(g) }, children: test(g) }] })).then((buf) => fs.writeFileSync(`LK_Verben_Gruppe_${g}.docx`, buf)));
["A", "B"].forEach((g) => Packer.toBuffer(new Document({ styles, sections: [{ properties: props, footers: { default: foot(g) }, children: test(g, true) }] })).then((buf) => fs.writeFileSync(`LK_Verben_Gruppe_${g}_Loesung.docx`, buf)));
Packer.toBuffer(new Document({ styles, sections: [{ properties: props, children: solution() }] })).then((buf) => fs.writeFileSync("LK_Verben_Erwartungshorizont.docx", buf));
