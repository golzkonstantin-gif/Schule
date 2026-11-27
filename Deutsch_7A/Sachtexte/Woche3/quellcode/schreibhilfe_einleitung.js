// Schnipsel „Einleitung eines Sachtexts – Schreibhilfe“ (27.11.): Merksatz, Satzanfänge, Schreibzeilen, Checkliste.
// Klasse: 2 pro A4 quer (zum Halbieren). PL=1: Handreichung Polnisch = eine A4-Seite hoch, zweisprachig. Ohne Farbflächen, ohne Zeitangaben.
const fs = require("fs");
const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, BorderStyle, AlignmentType, VerticalAlign, PageOrientation } = require("docx");
const NAVY = "1E2761", MUTED = "5B6B8C", RED = "D9534F", GREEN = "2E9E6B", GRAY = "9AA3B5", PLC = "5B4A8A";
const PLM = process.env.PL === "1";
const t = (text, o = {}) => new TextRun({ text, font: "Calibri", size: o.size || 21, bold: o.bold, italics: o.it, color: o.color || "000000" });
const p = (ch, o = {}) => new Paragraph({ children: Array.isArray(ch) ? ch : [ch], alignment: o.align, spacing: { before: o.before || 0, after: o.after ?? 30, line: o.line }, indent: o.indent });
const pl = (x, o = {}) => PLM ? [p(t(x, { size: 19, it: true, color: PLC }), Object.assign({ after: 30 }, o))] : [];
const b = (style, size, color) => ({ style, size, color });
const bs = (x) => ({ top: x, bottom: x, left: x, right: x });
const NB = bs(b(BorderStyle.NONE, 0, "FFFFFF"));
const cell = (ch, w, o = {}) => new TableCell({ children: ch, width: { size: w, type: WidthType.DXA }, borders: o.borders || NB, verticalAlign: o.valign || VerticalAlign.TOP, margins: { top: o.m ?? 40, bottom: o.m ?? 40, left: o.ml ?? 100, right: o.ml ?? 100 } });
const table = (cols, rows) => new Table({ columnWidths: cols, width: { size: cols.reduce((a, c) => a + c, 0), type: WidthType.DXA }, rows });
const HILF = [
  ["Thema nennen", NAVY, ["… ist eine Sportart, bei der …", "In … geht es darum, …", "Bei … treten … gegeneinander an.", "… gehört zu den …-Sportarten."], "podaj temat: o jakim sporcie piszesz"],
  ["Interesse wecken", RED, ["Wusstest du, dass …?", "Kaum jemand weiß, dass …", "Stell dir vor, …", "Was haben … und … gemeinsam?"], "zainteresuj czytelnika: pytanie, zaskakujący fakt, zwrot do czytelnika"],
  ["Überblick geben", GREEN, ["In diesem Text erfährst du, …", "Hier liest du, wie …", "Außerdem geht es darum, …", "Zum Schluss …"], "powiedz, czego czytelnik dowie się z tekstu"],
];
function blatt(W) {
  const inner = W - 300, cw = Math.floor(inner / 3);
  const lines = (n) => Array.from({ length: n }, () => p(t("_".repeat(PLM ? 80 : 56), { color: GRAY, size: 21 }), { before: 0, after: 0, line: 540 }));
  return table([W], [new TableRow({ children: [cell([
    p([t("Die Einleitung eines Sachtexts", { bold: true, size: 28, color: NAVY }), t("   Schreibhilfe", { size: 19, it: true, color: MUTED })], { after: 40 }),
    ...(PLM ? [p(t("Cel: Potrafię napisać wstęp do tekstu informacyjnego o moim sporcie: podaje temat, budzi ciekawość i daje przegląd. Przebieg: Leseband · sprawdzenie ćwiczenia z reportażem · cechy wstępu · piszesz wstęp o swoim sporcie (po niemiecku) · * partner sprawdza z listą kontrolną.", { size: 17, it: true, color: PLC }), { after: 60 })] : []),
    table([inner], [new TableRow({ children: [cell([
      p(t("Die Einleitung ist der Anfang eines Sachtexts. Sie nennt das Thema, weckt Interesse und gibt einen Überblick. Sie ist kurz (3–5 Sätze), sachlich und im Präsens.", { bold: true, size: 22 }), { after: PLM ? 20 : 0 }),
      ...pl("Wstęp (die Einleitung) to początek tekstu informacyjnego. Podaje temat, budzi ciekawość i mówi, o czym będzie tekst. Jest krótki (3–5 zdań), rzeczowy i w czasie teraźniejszym (Präsens).", { after: 0 }),
    ], inner, { borders: bs(b(BorderStyle.SINGLE, 12, RED)), m: 60 })] })]),
    p(t("Satzanfänge", { bold: true, size: 23, color: NAVY }), { before: 80, after: 30 }),
    table([cw, cw, inner - 2 * cw], [new TableRow({ children: HILF.map(([h, col, xs, plx], i) => cell([
      p(t(h, { bold: true, size: 22, color: col }), { after: 20 }),
      ...pl(plx, { after: 20 }),
      ...xs.map((x) => p(t("• " + x, { size: 20 }), { after: 30 })),
    ], i < 2 ? cw : inner - 2 * cw, { borders: bs(b(BorderStyle.SINGLE, 8, col)), m: 50, ml: 80 })) })]),
    p([t("Überschrift: ", { bold: true, size: 19 }), t("_".repeat(PLM ? 70 : 46), { color: GRAY })], { before: 120, after: 0, line: 430 }),
    ...lines(PLM ? 7 : 9),
    p(t("Checkliste (Partner prüft: ✓ / ✗)", { bold: true, size: 22, color: NAVY }), { before: 100, after: 20 }),
    ...pl("Lista kontrolna – partner sprawdza (✓ tak / ✗ nie)", { after: 20 }),
    ...[["Das Thema steht gleich am Anfang.", "Temat jest na samym początku."], ["Ein Satz weckt Interesse (Frage, Tatsache, Ansprache).", "Jedno zdanie budzi ciekawość."], ["Der Leser weiß, was im Text kommt.", "Czytelnik wie, o czym będzie tekst."], ["3–5 Sätze, sachlich, im Präsens.", "3–5 zdań, rzeczowo, w czasie teraźniejszym."]].flatMap(([x, y]) => [p([t("☐  ", { size: 22 }), t(x, { size: 21 })], { after: PLM ? 0 : 10 }), ...pl(y, { indent: { left: 360 }, after: 10 })]),
  ], W, { borders: bs(b(BorderStyle.DASHED, 6, GRAY)), m: 120, ml: 150 })] })]);
}
let doc, props;
if (PLM) {
  const W = 11906 - 1000;
  doc = [blatt(W)];
  props = { page: { size: { width: 11906, height: 16838 }, margin: { top: 450, bottom: 350, left: 500, right: 500 } } };
} else {
  const PAGEW = 16838, MARG = 350, GAP = 400, CW = Math.floor((PAGEW - 2 * MARG - GAP) / 2);
  const cut = cell([p(t("✂", { color: GRAY, size: 18 }), { align: AlignmentType.CENTER })], GAP, { valign: VerticalAlign.CENTER, m: 0, ml: 0 });
  doc = [table([CW, GAP, CW], [new TableRow({ children: [cell([blatt(CW)], CW, { m: 0, ml: 0 }), cut, cell([blatt(CW)], CW, { m: 0, ml: 0 })] })])];
  props = { page: { size: { width: 11906, height: 16838, orientation: PageOrientation.LANDSCAPE }, margin: { top: 350, bottom: 250, left: MARG, right: MARG } } };
}
Packer.toBuffer(new Document({ styles: { default: { document: { run: { font: "Calibri", size: 19 } } } }, sections: [{ properties: props, children: doc }] }))
  .then((buf) => fs.writeFileSync(__dirname + "/../" + (PLM ? "Handreichung_Polnisch_27-11_Einleitung" : "Schnipsel_Schreibhilfe_Einleitung_Sachtext") + ".docx", buf));
