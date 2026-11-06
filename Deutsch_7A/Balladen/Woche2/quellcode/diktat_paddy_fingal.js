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



// Diktat „Paddy Fingal“ (Puffer, Woche 2): Lehrerblatt mit Sinnschritten · Lösung zum Projizieren · Lückendiktat (Differenzierung, S-PL)
const VIO = "7A5BA6";
const s = (x, o = {}) => t(x, Object.assign({ size: 22 }, o));
const gray = "9AA3B5";
// Text in Sinnschritten; Markup: {s:…} Subjunktion, {k:,} Komma vor Subjunktion, {d:…} das/dass, {v:…} starkes Verb Präteritum, {r:…} Rechtschreibstelle (ss/ß, Großschreibung)
const SAETZE = [
  "Paddy Fingal war ein {r:Riese}{k:,} | der in Irland lebte.",
  "Alle wussten{k:,} | {d:dass} er sehr stark war.",
  "Eines Tages hörte er{k:,} | {d:dass} ein Hüne aus Schottland kam.",
  "{s:Als} Paddy ihn aus der Ferne sah{k:,} | {v:erschrak} er{k:,} | {s:weil} der Fremde viel {r:größer} war.",
  "Schnell {v:lief} er nach Hause und {v:rief}: | „Schaya, versteck mich!“",
  "Seine Frau steckte Paddy ins Bett | und häufte die Kissen hoch{k:,} | {s:sodass} nur noch seine Nase herausschaute.",
  "Zum Schotten sagte sie{k:,} | {d:dass} er leise sein sollte{k:,} | {s:damit} {d:das} Baby nicht aufwachte.",
  "{s:Als} der Schotte die riesige Nase sah{k:,} | bekam er {r:große} Angst.",
  "{s:Wenn} schon {d:das} Baby so groß ist{k:,} | wie riesig {r:muss} dann erst der Vater sein?",
  "Ohne ein Wort {v:floh} er zurück nach Schottland.",
  "Paddy aber {v:ließ} sich von den Nachbarn | als Held feiern.",
];
const WORDS = SAETZE.join(" ").replace(/\{[a-z]:([^}]*)\}/g, "$1").replace(/\|/g, "").split(/\s+/).filter((w) => /[A-Za-zÄÖÜäöüß]/.test(w)).length;
const COL = { s: VIO, k: C.hv, d: "2E9E6B", v: "C77C12", r: C.inf };
function parse(str, mode, size = 22) {
  // mode: "teacher" (Sinnschritte sichtbar, Markierung), "solution" (Markierung, keine Striche), "gap" (Lücken)
  const ss = (x, o = {}) => t(x, Object.assign({ size }, o));
  const src = mode === "teacher" ? str : str.replace(/ \|/g, "");
  const out = [];
  src.split(/(\{[a-z]:[^}]*\}|\|)/).filter((x) => x !== "").forEach((x) => {
    if (x === "|") { out.push(t("|", { size, bold: true, color: gray })); return; }
    const m = x.match(/^\{([a-z]):([^}]*)\}$/);
    if (!m) { out.push(ss(x)); return; }
    const [, k, v] = m;
    if (mode === "gap") {
      if (k === "k") out.push(ss("☐", { color: gray }));
      else if (k === "d") out.push(ss("______", { color: gray }));
      else if (k === "v") out.push(ss("________", { color: gray }));
      else out.push(ss(v));
      return;
    }
    out.push(ss(v, { bold: true, color: COL[k], ...(k === "k" ? { size: size + 4 } : {}) }));
  });
  return out;
}
const fixSpaces = (runs) => runs;
const doc = [];
// Seite 1: Lehrerblatt
doc.push(kicker("Für die Lehrkraft · Deutsch 7A · Puffer, falls Zeit bleibt (ca. 15 Min.)"), h1("Diktat „Paddy Fingal“"));
doc.push(box([p([S("Wozu: ", { bold: true }), S(`Das Diktat erzählt die Ballade im Präteritum nach (${WORDS} Wörter) und wiederholt dabei, was diese Woche dran war: Kommas zwischen Haupt- und Nebensatz (Sprachtraining Mittwoch), das/dass, das Präteritum starker Verben (Erzähltempus der Epik) und die Zeichen der wörtlichen Rede mit Redebegleitsatz.`)], { after: 0 })], LIGHT2));
doc.push(h3("Ablauf"));
[["1", "Text einmal ganz vorlesen. Ansage: „Stifte liegen. Ihr hört nur zu.“"],
 ["2", "In Sinnschritten (|) diktieren, jeden Schritt zweimal. Satzzeichen am Satzende und die Zeichen der wörtlichen Rede ansagen, die Kommas NICHT – sie sind der Übungsschwerpunkt."],
 ["3", "Text noch einmal ganz vorlesen. Ansage: „Lest mit und prüft jedes Komma.“"],
 ["4", "Selbstkorrektur mit der Lösung (Seite 2 projizieren): mit Grün verbessern, jede Subjunktion einkreisen, das finite Verb im Nebensatz unterstreichen."],
 ["5", "Wer will, zählt die Fehler je Schwerpunkt (Komma · das/dass · Verb · Rechtschreibung) – ohne Note."]].forEach(([n, x]) => doc.push(p([S(n + "  ", { bold: true, color: C.hv }), S(x)], { after: 50 })));
doc.push(h3("Diktattext mit Sinnschritten"));
doc.push(box(SAETZE.map((z) => p(parse(z, "teacher"), { after: 80, line: 340 })), LIGHT));
doc.push(p([S("Farben im Text: ", { bold: true }), S("Subjunktion", { bold: true, color: VIO }), S(" · "), S("Komma zwischen Haupt- und Nebensatz", { bold: true, color: C.hv }), S(" · "), S("das/dass", { bold: true, color: "2E9E6B" }), S(" · "), S("starkes Verb im Präteritum", { bold: true, color: "C77C12" }), S(" · "), S("ss/ß, Großschreibung", { bold: true, color: C.inf })], { before: 80, after: 60 }));
doc.push(h3("Differenzierung"));
doc.push(p(S("Lückendiktat (Seite 3, 2 pro A4) für Schwächere und S-PL: Der Text steht da, ausgefüllt werden nur die Kommas (☐), das/dass und die starken Verben. Auf der Fassung steht eine Zeile Polnisch für S-PL."), { after: 40 }));
// Seite 2: Lösung zum Projizieren
doc.push(new Paragraph({ pageBreakBefore: true, children: [t("LÖSUNG ZUM PROJIZIEREN", { size: 18, bold: true, color: MUTED })], spacing: { after: 40 } }));
doc.push(p(t("Paddy Fingal", { font: "Cambria", size: 36, bold: true, color: NAVY }), { after: 120 }));
doc.push(p(SAETZE.flatMap((z) => [...parse(z, "solution", 30), t(" ", { size: 30 })]), { after: 120, line: 420 }));
doc.push(p([t("Verbessere mit Grün. ", { bold: true, size: 26 }), t("Kreise jede Subjunktion ein (", { size: 26 }), t("violett", { bold: true, color: VIO, size: 26 }), t(") und unterstreiche das finite Verb im Nebensatz. Zwischen Haupt- und Nebensatz steht ein Komma (", { size: 26 }), t("rot", { bold: true, color: C.hv, size: 26 }), t(").", { size: 26 })], { before: 160 }));
// Seite 3: Lückendiktat
const lueck = () => table([11906 - 1100], [row([cell([
  p([t("Lückendiktat „Paddy Fingal“", { font: "Cambria", bold: true, size: 26, color: NAVY })], { after: 40 }),
  p([s("Setze die Kommas (☐), schreibe ", { size: 19 }), s("das", { size: 19, bold: true }), s(" oder ", { size: 19 }), s("dass", { size: 19, bold: true }), s(" und die fehlenden Verben im Präteritum ein. Hör genau zu!", { size: 19 })], { after: 20 }),
  p([t("PL: ", { size: 16, bold: true, color: "5B4A8A" }), t("Słuchaj dyktanda. Wstaw przecinki (☐), wpisz „das“ albo „dass“ i brakujące czasowniki w czasie przeszłym (Präteritum).", { size: 16, italics: true, color: "5B4A8A" })], { after: 80 }),
  p(SAETZE.flatMap((z) => [...parse(z, "gap"), s(" ")]), { after: 0, line: 480 }),
], { w: 11906 - 1100, borders: allBorders(dashed), m: 120, ml: 160 })])]);
doc.push(new Paragraph({ pageBreakBefore: true, children: [] }), lueck(), p(t("✂ - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -", { color: gray, size: 16 }), { before: 60, after: 60, align: AlignmentType.CENTER }), lueck());
const pageProps = { page: { size: { width: 11906, height: 16838 }, margin: { top: 700, bottom: 600, left: 550, right: 550, header: 300, footer: 300 } } };
const styles = { default: { document: { run: { font: "Calibri", size: 22 } } } };
Packer.toBuffer(new Document({ styles, sections: [{ properties: pageProps, footers: { default: footer("Diktat Paddy Fingal") }, children: doc }] })).then((buf) => fs.writeFileSync("Diktat_Paddy_Fingal.docx", buf));
