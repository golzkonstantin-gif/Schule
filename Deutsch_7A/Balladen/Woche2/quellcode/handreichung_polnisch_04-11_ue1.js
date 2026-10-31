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



// Arbeitsblatt Ü1 „Paddy Fingal – Aussagen prüfen“ (Mi 4.11.) – zweisprachige Fassung für S-PL
const GRN = "2E9E6B";
const PW = 11906 - 1100;
const gray = "9AA3B5";
const s = (x, o = {}) => t(x, Object.assign({ size: 21 }, o));
const gapL = (n) => s("_".repeat(n), { color: gray });
const bd = allBorders(solid("C9D3EA", 6));
const PLC = "5B4A8A";
const pls = (x, o = {}) => t(x, Object.assign({ size: 17, italics: true, color: PLC }, o));
const PL = [
  "Paddy jest tak silny, że rzuca masztem statku jak włócznią.",
  "Szkot przychodzi do Paddy’ego, bo chce go pozyskać jako sojusznika przeciw wrogowi.",
  "Kiedy Paddy widzi obcego z daleka, zauważa, że obaj są mniej więcej tego samego wzrostu.",
  "Paddy chce stanąć do walki i prosi Schayę, żeby pomogła mu się przygotować.",
  "Schaya prosi obcego, żeby był cicho, aby dziecko się nie obudziło.",
  "Szkot słucha ostrzeżenia Schai i trzyma się z daleka od łóżka.",
  "Szkot rozpoznaje Paddy’ego w łóżku i ucieka przed nim.",
  "Podczas ucieczki Szkot prawie stracił życie.",
  "Na końcu Paddy jest słusznie świętowany jako dzielny bohater, bo pokonał Szkota.",
];
const AUSSAGEN = [
  "Paddy ist so stark, dass er einen Schiffsmast wie einen Speer wirft.",
  "Der Schotte kommt zu Paddy, weil er ihn als Verbündeten gegen einen Feind gewinnen will.",
  "Als Paddy den Fremden von Weitem sieht, merkt er, dass beide ungefähr gleich groß sind.",
  "Paddy will sich dem Kampf stellen und bittet Schaya, ihm beim Vorbereiten zu helfen.",
  "Schaya bittet den Fremden, leise zu sein, damit das Kind nicht aufwacht.",
  "Der Schotte hält sich an Schayas Warnung und bleibt vom Bett weg.",
  "Der Schotte erkennt Paddy im Bett und flieht vor ihm.",
  "Auf der Flucht hätte der Schotte beinahe sein Leben verloren.",
  "Paddy wird am Ende zu Recht als tapferer Held gefeiert, weil er den Schotten besiegt hat.",
];
const NW = 650, KW = 1900, AW = PW - NW - KW;
const item = (i, a) => {
  const star = i === 8;
  return [
    row([
      cell(p(s((star ? "* " : "") + (i + 1), { bold: true, color: GRN, size: 24 }), { after: 0, align: AlignmentType.CENTER }), { w: NW, borders: bd, fill: LIGHT, valign: VerticalAlign.CENTER, rowSpan: undefined }),
      cell([p(s(a, { bold: true }), { after: 0 }), p(pls(PL[i]), { after: 0 })], { w: AW, borders: bd, m: 50, ml: 100 }),
      cell([p([s("☐ richtig", { size: 20 }), s("   ", {}), s("☐ falsch", { size: 20 })], { after: 0, align: AlignmentType.CENTER }), p(pls("prawda · fałsz", { size: 15 }), { after: 0, align: AlignmentType.CENTER })], { w: KW, borders: bd, valign: VerticalAlign.CENTER }),
    ]),
    row([cell([
      p([s("Beleg ", { bold: true, size: 19, color: MUTED }), pls("(dowód)", { size: 15 }), s(": Str. ", { size: 19 }), gapL(4), s("  V. ", { size: 19 }), gapL(5), s("  „", { size: 19 }), gapL(50), s("“", { size: 19 })], { after: 0, line: 470 }),
      p([s("Richtig ist ", { bold: true, size: 19, color: MUTED }), pls("(poprawnie – po niemiecku)", { size: 15 }), s(": ", { size: 19 }), gapL(66)], { after: 0, line: 470 }),
    ], { w: PW, borders: bd, m: 50, ml: 130, span: 3 })]),
  ];
};
const doc = [
  p([t("Ü1  Paddy Fingal – Aussagen prüfen", { font: "Cambria", bold: true, size: 30, color: NAVY }), pls("   Sprawdź stwierdzenia", { size: 20 }), t("    Name: ________________", { size: 18, color: MUTED })], { after: 60 }),
  p([s("1  ", { bold: true, color: GRN }), s("Lies die Ballade im Tandem (LB S. 136 f.).   "), s("2  ", { bold: true, color: GRN }), s("Prüfe jede Aussage allein: Kreuze an.")], { after: 20 }),
  p([s("3  ", { bold: true, color: GRN }), s("Belege mit Strophe, Vers und Zitat.   "), s("4  ", { bold: true, color: GRN }), s("Ist die Aussage falsch? Schreibe auf, was richtig ist.")], { after: 20 }),
  p([pls("PL: ", { bold: true, italics: false }), pls("1 Czytajcie balladę w parze (podręcznik s. 136 n.). 2 Sama sprawdź każde zdanie: zaznacz prawda (richtig) albo fałsz (falsch). 3 Podaj dowód: strofa (Str.), wers (V.) i cytat. 4 Jeśli zdanie jest fałszywe, napisz po niemiecku, jak jest naprawdę. Uwaga: zdania są napisane innymi słowami niż w balladzie.")], { after: 80 }),
  box([p([s("Beispiel: ", { bold: true, color: GRN, size: 19 }), s("„Paddy hat einen roten Bart.“  ", { size: 19, italics: true }), s("☒ falsch", { size: 19, bold: true }), s("  ·  Beleg: Str. 1, V. 2, „Sechs Ellen flog sein schwarzer Bart“  ·  Richtig ist: Paddys Bart ist schwarz.", { size: 19 })], { after: 0 }), p(pls("Przykład: „Paddy ma rudą brodę.“ – fałsz · dowód: strofa 1, wers 2 · poprawnie: Broda Paddy’ego jest czarna."), { after: 0 })], LIGHT2),
  p(t(""), { after: 60 }),
  table([NW, AW, KW], AUSSAGEN.slice(0, 5).flatMap((a, i) => item(i, a))),
  p(new PageBreak()),
  p([t("Ü1  Paddy Fingal – Aussagen prüfen (Rückseite)", { font: "Cambria", bold: true, size: 26, color: NAVY }), pls("   ciąg dalszy", { size: 18 })], { after: 100 }),
  table([NW, AW, KW], AUSSAGEN.slice(5).flatMap((a, i) => item(i + 5, a))),
  p([t("Wichtige Wörter", { bold: true, size: 22, color: NAVY, font: "Cambria" }), pls("   ważne słowa", { size: 18 })], { before: 200, after: 40 }),
  table([2600, 2850, 2600, PW - 8050], [
    ["der Schiffsmast · der Speer", "maszt statku · włócznia", "der Verbündete · der Feind", "sojusznik · wróg"],
    ["sich dem Kampf stellen", "stanąć do walki", "von Weitem", "z daleka"],
    ["die Warnung · sich halten an", "ostrzeżenie · stosować się do", "erkennen · fliehen", "rozpoznać · uciekać"],
    ["beinahe · das Leben verlieren", "prawie · stracić życie", "zu Recht · feiern", "słusznie · świętować"],
    ["der Beleg · das Zitat", "dowód · cytat", "die Aussage · richtig · falsch", "stwierdzenie · prawda · fałsz"],
  ].map((r) => row(r.map((x, i) => cell(p(i % 2 ? pls(x, { italics: false, color: "333333", size: 18 }) : s(x, { bold: true, size: 18, color: NAVY }), { after: 0 }), { w: [2600, 2850, 2600, PW - 8050][i], borders: bd, m: 30, ml: 70, fill: i % 2 ? "FBFCFF" : undefined }))))),
  p([s("Zeit: ", { bold: true, size: 18 }), s("8 Min. lesen (Tandem) + 10 Min. prüfen (allein)      ", { size: 18 }), s("Fertig? Stift hinlegen.", { size: 18, italics: true, color: MUTED })], { before: 80, after: 0 }),
];
const props = { page: { size: { width: 11906, height: 16838 }, margin: { top: 400, bottom: 300, left: 550, right: 550 } } };
const styles = { default: { document: { run: { font: "Calibri", size: 21 } } } };
Packer.toBuffer(new Document({ styles, sections: [{ properties: props, children: doc }] })).then((buf) => fs.writeFileSync("Handreichung_Polnisch_04-11_Ue1_Aussagen_pruefen.docx", buf));
