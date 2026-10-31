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



// Handreichung Polnisch: Lesetechniken wiederholen (15 Min.) – halbe A4-Seite, 2 pro Blatt
const PW = 11906 - 900;
const gray = "9AA3B5", PLC = "5B4A8A";
const s = (x, o = {}) => t(x, Object.assign({ size: 19 }, o));
const pls = (x, o = {}) => t(x, Object.assign({ size: 17, italics: true, color: PLC }, o));
const de2 = (runs, o = {}) => p(runs, Object.assign({ after: 0 }, o));
const pl2 = (x, o = {}) => p([pls("PL: ", { bold: true, italics: false }), pls(x)], Object.assign({ after: 50 }, o));
const hd = (x, plx) => p([t(x, { bold: true, size: 20, color: NAVY }), pls("   " + plx)], { before: 50, after: 10 });
const bd = allBorders(solid("C9D3EA", 6));
const WW = [2700, PW - 260 - 2700];
const wrow = (d, pl) => row([cell(p(s(d, { bold: true, color: NAVY, size: 18 }), { after: 0 }), { w: WW[0], borders: bd, m: 20, ml: 70 }), cell(p(pls(pl, { italics: false, color: "333333" }), { after: 0 }), { w: WW[1], borders: bd, m: 20, ml: 70, fill: "FBFCFF" })]);
function blatt() {
  return table([PW], [row([cell([
    p([t("Lesetechniken wiederholen", { font: "Cambria", bold: true, size: 25, color: NAVY }), pls("   Powtórka: techniki czytania", { size: 18 })], { after: 30 }),
    de2([s("Ziel: ", { bold: true }), s("Ich kann die drei Lesetechniken unterscheiden und begründen, welche zu meinem Leseziel passt.", { italics: true })]),
    pl2("Cel: Potrafię rozróżnić trzy techniki czytania i uzasadnić, która pasuje do mojego celu czytania."),
    hd("Ablauf", "przebieg lekcji"),
    pl2("1 Techniki czytania z pamięci · 2 Która technika pasuje? (A–D) · 3 * Przejrzyj „Paddy Fingal“ (tylko jeśli starczy czasu) · 4 Czy osiągnęliśmy cel?"),
    hd("Wichtige Wörter", "ważne słowa"),
    table(WW, [
      wrow("das Leseziel", "cel czytania – co chcę wyczytać z tekstu"),
      wrow("überfliegend lesen", "czytać pobieżnie – szybko, żeby wiedzieć, o czym jest tekst"),
      wrow("orientierend lesen", "czytać wybiórczo – szukam jednej konkretnej informacji"),
      wrow("genau lesen", "czytać dokładnie – zdanie po zdaniu, chcę wszystko zrozumieć"),
      wrow("der Trainingsplan · durchblättern", "plan treningów · przeglądać, kartkować"),
      wrow("die Wettkampfregeln (Pl.)", "przepisy zawodów (disqualifiziert werden = zostać zdyskwalifikowanym)"),
    ]),
    hd("1  Lesetechniken aus dem Gedächtnis", "z pamięci"),
    de2(s("Schreibe die drei Lesetechniken untereinander. Ergänze zu jeder: Ich lese so, wenn …")),
    pl2("Zapisz w zeszycie trzy techniki czytania, jedna pod drugą. Do każdej dopisz zdanie „Ich lese so, wenn …“ (Czytam tak, gdy …). Potem porównaj z sąsiadem i popraw na zielono."),
    hd("2  Welche Lesetechnik passt? Begründe mit „weil“.", "która technika pasuje? Uzasadnij"),
    pl2("Zadanie A–D jest na tablicy. Zapisz literę, technikę i uzasadnienie. Uwaga: w zdaniu z „weil“ czasownik stoi na końcu, np. „…, weil ich nur eine Uhrzeit suche.“", { after: 20 }),
    pl2("A – sprawdzasz w planie treningów, o której w środę zaczyna się trening.", { after: 0 }),
    pl2("B – przeglądasz gazetę sportową i chcesz wiedzieć, czy jakiś artykuł cię interesuje.", { after: 0 }),
    pl2("C – przed zawodami czytasz nowe przepisy, żeby nie zostać zdyskwalifikowaną.", { after: 0 }),
    pl2("D – masz wyjaśnić, dlaczego Szkot w „Paddy Fingal“ na końcu ucieka.", { after: 40 }),
    hd("* E  Paddy Fingal überfliegen (1 Min.)", "przejrzyj balladę"),
    pl2("Przejrzyj balladę w podręczniku (s. 136) przez 1 minutę. Napisz: „In der Ballade geht es um …“ (Ballada opowiada o …)."),
    de2([s("Zum Schluss: ", { bold: true }), s("Heute habe ich gelernt, …")]),
    pl2("Na koniec zapisz w zeszycie: „Heute habe ich gelernt, …“ (Dziś nauczyłam się, …).", { after: 0 }),
  ], { w: PW, borders: allBorders(dashed), m: 80, ml: 130, valign: VerticalAlign.TOP })])]);
}
const props = { page: { size: { width: 11906, height: 16838 }, margin: { top: 450, bottom: 300, left: 450, right: 450 } } };
const styles = { default: { document: { run: { font: "Calibri", size: 19 } } } };
Packer.toBuffer(new Document({ styles, sections: [{ properties: props, children: [blatt()] }] })).then((buf) => fs.writeFileSync("Handreichung_Polnisch_Lesetechniken_wiederholen.docx", buf));
