// „Paddy Fingal“ (Wilhelm Brandes, 1854–1928; Text gemeinfrei) – zweimal auf einem A4-Blatt (quer, zum Halbieren)
// Anlage zum Test „Paddy Fingal unter der Lupe“. Vers- und Strophenzählung wie im LB S. 136 f. Worterklärungen eigene Formulierung.
// Ohne Farbflächen. Aufruf: node text_paddy_fingal_2x.js
const fs = require("fs");
const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, BorderStyle, AlignmentType, VerticalAlign, HeightRule, PageOrientation } = require("docx");
const NAVY = "1E2761", MUTED = "5B6B8C", GRAY = "9AA3B5";
const FS = 17; // 8,5 pt
const STR = [
  ["Paddy Fingal war von Riesenart:", "Sechs Ellen flog sein schwarzer Bart,", "seine Fäuste waren wie Tonnen groß,", "einen Mastbaum führt’ er als Wurfgeschoss –", "so ein Kerl war Paddy Fingal!"],
  ["Nun wuchs ein Hüne im Schottenland,", "der hörte von Fingals starker Hand;", "da wollt’ er erproben alsogleich,", "wer fester sei auf Stoß und Streich,", "er oder Paddy Fingal."],
  ["Und als er stapfte durch den Sund,", "Paddy Fingal just am Ufer stund", "und maß von ferne klipp und klar,", "dass der Fremdling zehn Schuh größer war,", "noch größer als Paddy Fingal."],
  ["Da lief er heim in jähem Schreck:", "„O Schaya, birg mich im Versteck!", "Von Schottland kommt ein Kerl daher,", "wie ein Berg so groß – ich fürchte sehr,", "der sucht den Paddy Fingal.“"],
  ["Ins Bette Paddy Fingal kroch,", "Frau Schaya türmte die Kissen hoch;", "wie aus dem Hedesack die Maus,", "so guckte die Nase nur heraus,", "die Nase von Paddy Fingal."],
  ["Indem so schob der Schotte herein,", "an den Balken rührte sein Scheitelbein,", "und er schnob und wischte sich den Schweiß", "und rollte die Augen wild im Kreis:", "„Wo steckt der Paddy Fingal?“"],
  ["„Tut leise, Fremder, und tretet sacht,", "dass Paddys Kindlein nicht erwacht!", "Denn wenn er schrie und Fingal käm,", "für euch kein gutes End es nähm:", "Nicht spaßen tut Paddy Fingal.“"],
  ["Doch wie sie warnte mit Wort und Wink,", "der Schotte neugierig ans Lager ging:", "O heiliger Patrick, wie ward ihm da,", "als er die Nase ragen sah,", "die Nase von Paddy Fingal!"],
  ["„Beim Pfeifer, der vor Moses blies:", "welch heidenhafter Nasenspieß!", "Ist das ein Baby, wie Ihr sagt,", "ein Narr, wer’s mit ihm selber wagt!", "Nicht wart ich auf Paddy Fingal.“"],
  ["Und er trollte davon mit scheuem Blick", "und stolperte durch den Sund zurück;", "fast wär ertrunken der gute Held,", "dieweil in der Eil er die Furt verfehlt,", "so lief er vor Paddy Fingal."],
  ["Der aber erhob ein Siegesgeschrei.", "Da kamen die Nachbarn rings herbei;", "die staunten den großen Fingal an,", "der den langen Schotten gejagt von dann,", "den tapfern Paddy Fingal."],
];
const WORT = "Elle: altes Längenmaß (etwa ein Unterarm) · Hüne: Riese · alsogleich: sofort · Sund: Meerenge · stund: stand · birg mich: versteck mich · Hedesack: Sack mit Flachsabfällen · Scheitelbein: Knochen oben am Kopf · schnob: schnaubte · sacht: leise · Lager: Bett · heiliger Patrick: Schutzheiliger Irlands · Beim Pfeifer, der vor Moses blies: irischer Ausruf des Staunens · Nasenspieß: riesige Nase · trollte davon: ging weg · dieweil: weil · Furt: flache Stelle zum Durchqueren eines Flusses";

const t = (text, o = {}) => new TextRun({ text, font: o.font || "Calibri", size: o.size || FS, bold: o.bold, italics: o.it, color: o.color || "000000" });
const p = (ch, o = {}) => new Paragraph({ children: Array.isArray(ch) ? ch : [ch], alignment: o.align, spacing: { before: o.before || 0, after: o.after ?? 0, line: o.line }, indent: o.indent });
const NB = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
const NONE = { top: NB, bottom: NB, left: NB, right: NB };
const cell = (children, w, o = {}) => new TableCell({ children: Array.isArray(children) ? children : [children], width: { size: w, type: WidthType.DXA }, borders: o.borders || NONE, verticalAlign: o.valign || VerticalAlign.TOP, margins: { top: 0, bottom: 0, left: o.ml ?? 0, right: o.mr ?? 0 } });
const table = (cols, rows) => new Table({ columnWidths: cols, width: { size: cols.reduce((a, b) => a + b, 0), type: WidthType.DXA }, rows });
const LH = 300; // Zeilenhöhe (twips)
const NW = 270, SW = 230; // Spalten: Versnummer, Strophennummer

function column(from, to, extra = []) {
  const cw = COLW - NW - SW;
  const rows = [];
  for (let si = from; si < to; si++) {
    STR[si].forEach((v, k) => {
      const vn = si * 5 + k + 1;
      const kehr = k === 4;
      rows.push(new TableRow({ height: { value: LH, rule: HeightRule.ATLEAST }, children: [
        cell(p(t(vn % 5 === 0 ? String(vn) : "", { size: 14, color: MUTED }), { align: AlignmentType.RIGHT }), NW, { mr: 60 }),
        cell(p(t(k === 0 ? String(si + 1) : "", { size: 15, bold: true, color: NAVY })), SW),
        cell(p(t(v), { indent: kehr ? { left: 500 } : undefined }), cw),
      ] }));
    });
    if (si < to - 1) rows.push(new TableRow({ height: { value: 140, rule: HeightRule.EXACT }, children: [cell(p(t("")), NW), cell(p(t("")), SW), cell(p(t("")), cw)] }));
  }
  return [table([NW, SW, cw], rows), ...extra];
}
const PAGEW = 16838, MARG = 300, GAP = 420;
const COPYW = Math.floor((PAGEW - 2 * MARG - GAP) / 2);
const CG = 120, COLW = Math.floor((COPYW - CG) / 2);
function copy() {
  const head = [
    p([t("Wilhelm Brandes", { size: 17, color: MUTED })]),
    p([t("Paddy Fingal", { font: "Cambria", size: 26, bold: true, color: NAVY })], { after: 80 }),
  ];
  const wort = [p([t("Worterklärungen: ", { size: 14, bold: true, color: MUTED }), t(WORT, { size: 14, color: MUTED })], { before: 160 })];
  return table([COLW, CG, COLW], [new TableRow({ children: [
    cell([...head, ...column(0, 6)], COLW),
    cell(p(t("")), CG),
    cell([p(t(""), { after: 330 }), ...column(6, 11, wort)], COLW),
  ] })]);
}
const cut = cell(p(t("✂", { color: GRAY, size: 18 }), { align: AlignmentType.CENTER }), GAP, { valign: VerticalAlign.CENTER, borders: { top: NB, bottom: NB, right: NB, left: { style: BorderStyle.DASHED, size: 6, color: GRAY } } });
const doc = [table([COPYW, GAP, COPYW], [new TableRow({ children: [cell([copy()], COPYW), cut, cell([copy()], COPYW)] })])];
const props = { page: { size: { width: 11906, height: 16838, orientation: PageOrientation.LANDSCAPE }, margin: { top: 450, bottom: 300, left: MARG, right: MARG } } };
Packer.toBuffer(new Document({ styles: { default: { document: { run: { font: "Calibri", size: FS } } } }, sections: [{ properties: props, children: doc }] }))
  .then((buf) => fs.writeFileSync(__dirname + "/../Text_Paddy_Fingal_2x_A4.docx", buf));
