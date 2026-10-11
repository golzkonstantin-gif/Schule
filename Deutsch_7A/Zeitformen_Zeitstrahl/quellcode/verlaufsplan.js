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
// Verlaufsplan „Welche Zeitform wann?“ (Orientierungshospitation, 45 Min.) – Kurzfassung
// ================================================================
const { PageOrientation } = require("docx");
const LW = 16838 - 1400; // nutzbare Breite quer
const s20 = (x, o = {}) => t(x, Object.assign({ size: 20 }, o));
const lp = (runs, o = {}) => p(Array.isArray(runs) ? runs : [s20(runs)], Object.assign({ after: 40 }, o));
const L = (x) => lp([s20("L: ", { bold: true, color: NAVY }), s20(x)]);
const SS = (x) => lp([s20("S: ", { bold: true, color: "C77C12" }), s20(x)]);
const plain = (x) => lp([s20(x)]);

const VW = [1000, 2600, 7600, 1600, LW - 12800];
const vhdr = ["Zeit", "Phase", "Unterrichtsgeschehen", "Sozialform", "Medien"];
const vcell = (children, w, o = {}) => cell(children, { w, fill: o.fill, borders: allBorders(solid("B9C6E8", 6)), m: 70, ml: 100, valign: VerticalAlign.TOP });
const vrow = (zeit, phase, geschehen, sf, medien) => row([
  vcell([p(t(zeit, { bold: true, size: 20, color: C.hv }), { after: 0 })], VW[0], { fill: LIGHT }),
  vcell([p(t(phase, { bold: true, size: 20, color: NAVY, font: "Cambria" }), { after: 0 })], VW[1], { fill: LIGHT }),
  vcell(geschehen, VW[2]),
  vcell([plain(sf)], VW[3]),
  vcell(medien.map(plain), VW[4]),
]);

const doc = [];
doc.push(kicker("Verlaufsplan · Deutsch · Klasse 7A · Einzelstunde (45 Min.)"));
doc.push(h1("Welche Zeitform wann? – Die Zeitformen am Zeitstrahl"));
doc.push(p([s20("Stundenziel: ", { bold: true, color: NAVY }), s20("Die Schülerinnen und Schüler ordnen die fünf bekannten Tempora an einem Mustertext zeitlich ein, sichern ihre Bildung in einer Übersicht und wenden sie in einem eigenen Text über ihr Wettkampfwochenende an.")], { after: 140 }));

doc.push(table(VW, [
  row(vhdr.map((h, i) => hdr(h, VW[i], NAVY, 19)), 380),
  vrow("0′–3′", "Einstieg", [
    SS("erzählen kurz von ihren Wettkämpfen am Wochenende."),
    L("nennt Stundenfrage und Ziel: eigener Text in allen Zeitformen."),
  ], "UG", ["Tafel (Zeitstrahl)"]),
  vrow("3′–12′", "Erarbeitung I: Mustertext am Zeitstrahl", [
    SS("ordnen in einer Magnet-Kette acht Satzstreifen (Wettkampf eines Gewichthebers) am Zeitstrahl und markieren die Verbformen."),
    L("gibt vor: Wer erzählt, schreibt im Präteritum; wer spricht, nimmt das Perfekt."),
  ], "EA an der Tafel, UG", ["Satzstreifen (Magnete)"]),
  vrow("12′–25′", "Erarbeitung II: Übersicht", [
    SS("hängen in einer Kreide-Kette Zeitform- und Bauplan-Karten über die Abschnitte und begründen am Satz; die übrigen ergänzen die Tabelle im Heft."),
  ], "EA an der Tafel, EA", ["Zeitform- und Bauplan-Karten", "Tabelle (Schnipsel)"]),
  vrow("25′–40′", "Anwendung: Exit-Ticket", [
    SS("schreiben 5–7 Sätze über ihr Wettkampfwochenende, jede Zeitform mindestens einmal, gestützt durch Satzanfänge."),
  ], "EA", ["Exit-Ticket"]),
  vrow("40′–45′", "Sicherung", [
    SS("zwei lesen vor, die Klasse ordnet am Zeitstrahl zu."),
    L("beantwortet mit der Klasse die Stundenfrage, sammelt die Exit-Tickets ein."),
  ], "UG", ["Tafel, Exit-Tickets"]),
  vrow("opt.", "Reserve", [plain("Speed-Duell zu zweit: Zeitformen bestimmen.")], "PA", ["–"]),
]));
doc.push(p(s20("UG = Unterrichtsgespräch · EA = Einzelarbeit · PA = Partnerarbeit", { color: MUTED, size: 17 }), { before: 80, after: 0 }));

const landProps = { page: { size: { width: 11906, height: 16838, orientation: PageOrientation.LANDSCAPE }, margin: { top: 700, bottom: 600, left: 700, right: 700, header: 350, footer: 350 } } };
const styles = { default: { document: { run: { font: "Calibri", size: 20 } } } };
const vhead = new Header({ children: [new Paragraph({ tabStops: [{ type: TabStopType.RIGHT, position: LW }], children: [t("Deutsch · Klasse 7A · Welche Zeitform wann?", { size: 17, color: MUTED }), t("\tVerlaufsplan", { size: 17, color: MUTED })] })] });
Packer.toBuffer(new Document({ styles, sections: [{ properties: landProps, headers: { default: vhead }, children: doc }] })).then((buf) => fs.writeFileSync("Verlaufsplan_Zeitformen_Zeitstrahl.docx", buf));
