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

const none = { style: BorderStyle.NONE, size: 0, color: NAVY };
const noBorders = { top: none, bottom: none, left: none, right: none };
const solid = (c = NAVY, s = 8) => ({ style: BorderStyle.SINGLE, size: s, color: c });
const allBorders = (b) => ({ top: b, bottom: b, left: b, right: b });
const dashed = { style: BorderStyle.DASHED, size: 8, color: "7A869E" };
const thin = allBorders(solid("B9C6E8", 6));

function cell(children, o = {}) {
  return new TableCell({
    children: Array.isArray(children) ? children : [children],
    width: { size: o.w, type: WidthType.DXA },
    shading: undefined, // Tinte sparen: keine Farbflächen im Druck
    borders: o.borders || thin,
    verticalAlign: o.valign || VerticalAlign.CENTER,
    margins: { top: o.m ?? 60, bottom: o.m ?? 60, left: o.ml ?? 100, right: 100 },
    columnSpan: o.span,
  });
}
const table = (widths, rows) => new Table({ width: { size: widths.reduce((a, b) => a + b, 0), type: WidthType.DXA }, columnWidths: widths, rows });
const row = (cells, h, rule = HeightRule.ATLEAST) => new TableRow({ children: cells, height: h ? { value: h, rule } : undefined, cantSplit: true });
const hdr = (txt, w, fill = NAVY, size = 20) => cell(p(t(txt, { bold: true, color: NAVY, size }), { align: AlignmentType.CENTER, after: 0 }), { w, fill });
const tc = (txt, w, o = {}) => cell(p(Array.isArray(txt) ? txt : t(txt, { size: o.size || 21, bold: o.bold, color: o.color, font: o.font, italics: o.it }), { after: 0, align: o.align }), { w, fill: o.fill });

// simple grid table: header + rows of strings/runs
function grid(widths, head, rows, o = {}) {
  return table(widths, [
    row(head.map((h, i) => hdr(h, widths[i], o.headFill ? o.headFill[i] : NAVY)), 420),
    ...rows.map((r) => row(r.map((v, i) => tc(v, widths[i], { font: o.fonts && o.fonts[i], align: o.align && o.align[i], bold: o.bold && o.bold[i], color: o.colors && o.colors[i] })), o.h || 460)),
  ]);
}

function box(children, fill = LIGHT2, border) {
  return table([W], [row([cell(children, { w: W, fill, borders: border || allBorders(solid("B9C6E8", 8)), m: 140, ml: 180 })])]);
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
    row([cell(p(t("TAFEL · " + title, { bold: true, color: NAVY, size: 18 }), { after: 0 }), { w: W, fill: "3D4A3F", borders: allBorders(solid("3D4A3F", 12)) })]),
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
    row([cell(p(t("TAFEL · " + title, { bold: true, color: NAVY, size: 16 }), { after: 0 }), { w: TW, fill: "3D4A3F", borders: allBorders(solid("3D4A3F", 12)), m: 30 })]),
    row([cell(children, { w: TW, fill: BOARD, borders: allBorders(solid("3D4A3F", 12)), m: 100, ml: 160 })]),
  ]);
}
const bl = (runs, o = {}) => p(Array.isArray(runs) ? runs : [N(runs)], Object.assign({ after: 30 }, o));
// Merksatz steht auf der Folie (nicht an der Tafel)
function folie(title, children) {
  return table([TW], [
    row([cell(p(t("FOLIE · " + title, { bold: true, color: NAVY, size: 16 }), { after: 0 }), { w: TW, fill: NAVY, borders: allBorders(solid(NAVY, 12)), m: 30 })]),
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



const VIO = "7A5BA6", GRN = "2E9E6B", PLC = "5B4A8A";
const PW = 11906 - 800;
const f = 19;
const s = (x, o = {}) => t(x, Object.assign({ size: f + 1 }, o));
const pls = (x, o = {}) => t(x, Object.assign({ size: f, italics: true, color: PLC }, o));
const de2 = (runs, o = {}) => p(runs, Object.assign({ after: 0 }, o));
const pl2 = (x, o = {}) => p([pls("PL: ", { bold: true, italics: false }), pls(x)], Object.assign({ after: 40 }, o));
const hd = (x, plx, col = NAVY) => p([t(x, { bold: true, size: f + 3, color: col, font: "Cambria" }), pls("   " + plx)], { before: 70, after: 10, keepNext: true });
const bd = allBorders(solid("C9D3EA", 6));
const wtab = (pairs, w1 = 2300) => {
  const half = Math.floor((PW - 40) / 2), w2 = half - w1;
  const rws = [];
  for (let i = 0; i < pairs.length; i += 2) {
    const c = (pr) => pr ? [cell(p(s(pr[0], { bold: true, color: NAVY, size: f }), { after: 0 }), { w: w1, borders: bd, m: 15, ml: 60 }), cell(p(pls(pr[1], { italics: false, color: "333333" }), { after: 0 }), { w: w2, borders: bd, m: 15, ml: 60 })] : [cell(p(t("")), { w: w1, borders: bd }), cell(p(t("")), { w: w2, borders: bd })];
    rws.push(row([...c(pairs[i]), ...c(pairs[i + 1])]));
  }
  return table([w1, half - w1, w1, half - w1], rws);
};

// Handreichung Polnisch: Reportage „Sehen mit den Fingerkuppen“ – Aussagen prüfen. Eine A4-Seite, ohne Farbflächen, keine Lösungen.
const doc = [
  p([t("Sehen mit den Fingerkuppen – Aussagen prüfen", { font: "Cambria", bold: true, size: 26, color: NAVY }), pls("   pomoc do lekcji", { size: 18 })], { after: 30 }),
  pl2("Tekst: reportaż „Sehen mit den Fingerkuppen – einen Tag blind erleben“ (podręcznik s. 183–184). Uczniowie spędzają 90 minut w ciemnej wystawie z niewidomym przewodnikiem i dowiadują się, jak żyją osoby niewidome (sygnalizacja świetlna, pismo Braille’a)."),
  hd("Aufgabe", "zadanie", GRN),
  de2([s("1 Prüfe jede Aussage allein: Kreuze an.  2 Belege mit Zeile und Zitat.  3 Ist die Aussage falsch? Schreibe auf, was richtig ist.")]),
  pl2("1 Sprawdź każde zdanie sama: zaznacz richtig (prawda) albo falsch (fałsz). 2 Udowodnij: numer linijki (Z.) i dokładny cytat z tekstu w cudzysłowie „…“. 3 Jeśli zdanie jest fałszywe, napisz po niemiecku, jak jest naprawdę („Richtig ist: …“). Zdanie 10 z gwiazdką (*) – jeśli skończysz wcześniej."),
  hd("Die Aussagen", "zdania do sprawdzenia – po polsku, bez rozwiązań", VIO),
  wtab([
    ["1", "Zwiedzanie ciemnych sal trwa dziewięćdziesiąt minut."], ["2", "Przewodnik, który prowadzi grupę, dobrze widzi."],
    ["3", "Reporterka cieszy się bez żadnych obaw na dzień w ciemności."], ["4", "Niewidomi znajdują słup sygnalizacji dzięki stukaniu."],
    ["5", "Szybkie pikanie oznacza, że zaraz zapali się czerwone światło."], ["6", "Pismo dla niewidomych czyta się opuszkami palców."],
    ["7", "Pismo Braille’a wynaleziono ok. 50 lat temu w Niemczech."], ["8", "Znak pisma Braille’a składa się najwyżej z sześciu punktów."],
    ["9", "Przewodnik uważa, że wszędzie myśli się już o potrzebach niewidomych."], ["* 10", "Grupa widzi swojego przewodnika dopiero na końcu."],
  ], 700),
  hd("Wörter im Text", "trudne słowa (Z. = linijka)", NAVY),
  wtab([
    ["die Finsternis (Z. 1)", "ciemność"], ["blind · sehbehindert (Z. 3–4)", "niewidomy · niedowidzący"],
    ["der Guide (Z. 4)", "przewodnik"], ["bewältigen (Z. 7)", "poradzić sobie z czymś"],
    ["ängstlich (Z. 19)", "trochę przestraszony"], ["der Blindenstock (Z. 32)", "biała laska"],
    ["die Ampel · der Ampelmast", "sygnalizacja świetlna · słup"], ["das Auffindesignal (Z. 52)", "sygnał do odnalezienia"],
    ["das Freigabesignal (Z. 58)", "sygnał: można przejść"], ["die Fingerkuppe (Z. 86–87)", "opuszek palca"],
    ["ertasten (Z. 90)", "wyczuć dotykiem"], ["die Punktschrift (Z. 96)", "pismo punktowe (Braille)"],
    ["die Bedürfnisse (Z. 116)", "potrzeby"], ["die Öffentlichkeit (Z. 114–115)", "przestrzeń publiczna"],
  ], 2900),
];
const props = { page: { size: { width: 11906, height: 16838 }, margin: { top: 450, bottom: 350, left: 400, right: 400 } } };
const styles = { default: { document: { run: { font: "Calibri", size: f } } } };
Packer.toBuffer(new Document({ styles, sections: [{ properties: props, children: doc }] })).then((buf) => fs.writeFileSync(__dirname + "/../Handreichung_Polnisch_Reportage_Aussagen.docx", buf));
