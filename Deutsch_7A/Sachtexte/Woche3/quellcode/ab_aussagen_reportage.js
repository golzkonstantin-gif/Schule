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
  tafel: "kommt an die Tafel", kontrolle: "kontrolliert", partner: "arbeitet zu zweit", fertig: "gibt ab", gruppe: "arbeitet in der Gruppe",
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



// Arbeitsblatt „Sehen mit den Fingerkuppen – Aussagen prüfen“ (LB S. 183–184, Reportage von Hilda Bermech, Z. 1–135)
// Schülerfassung + Lösung. Ohne Farbflächen. Nr. 10 mit *.
const AUS = JSON.parse(fs.readFileSync(__dirname + "/aussagen_reportage.json", "utf8"));
const GRN = "2E9E6B", SOLC = "2E7D4F";
const PW = 11906 - 1100;
const gray = "9AA3B5";
const s = (x, o = {}) => t(x, Object.assign({ size: 21 }, o));
const gapL = (n) => s("_".repeat(n), { color: gray });
const bd = allBorders(solid("C9D3EA", 6));
const NW = 650, KW = 1900, AW = PW - NW - KW;
const PLC = "5B4A8A";
const WORT = [["die Finsternis (Z. 1)", "ciemność"], ["blind · sehbehindert (Z. 3–4)", "niewidomy · niedowidzący"], ["der Guide (Z. 4)", "przewodnik"], ["bewältigen (Z. 7)", "poradzić sobie z czymś"], ["ängstlich (Z. 19)", "trochę przestraszony"], ["der Blindenstock (Z. 32)", "biała laska"], ["die Ampel · der Ampelmast", "sygnalizacja świetlna · słup"], ["das Auffindesignal (Z. 52)", "sygnał do odnalezienia słupa"], ["das Freigabesignal (Z. 58)", "sygnał: można przejść"], ["die Fingerkuppe (Z. 86–87)", "opuszek palca"], ["ertasten (Z. 90)", "wyczuć dotykiem"], ["die Punktschrift (Z. 96)", "pismo punktowe (Braille)"], ["die Bedürfnisse (Z. 116)", "potrzeby"], ["die Öffentlichkeit (Z. 114–115)", "przestrzeń publiczna"]];
function blatt(sol, PLm) {
  const plp = (x) => p(t(x, { size: 18, italics: true, color: PLC }), { after: 60 });
  const item = (i, [a, ok, z, zit, r]) => {
    const star = i === 9;
    const sol0 = sol; sol = sol0 || (PLm && i === 0); // S-PL: Nr. 1 als gelöstes Beispiel
    try {
    const kreuz = (lab, on) => s(`${sol && on ? "☒" : "☐"} ${lab}`, { size: 20, bold: sol && on, color: sol && on ? SOLC : "000000" });
    return [
      row([
        cell(p(s((star ? "* " : "") + (i + 1), { bold: true, color: GRN, size: 24 }), { after: 0, align: AlignmentType.CENTER }), { w: NW, borders: bd, valign: VerticalAlign.CENTER }),
        cell(p([s("„" + a + "“", { bold: true }), ...(PLm && i === 0 ? [t("   ← Beispiel / przykład", { size: 17, italics: true, color: PLC, bold: true })] : [])], { after: 0 }), { w: AW, borders: bd, m: 60, ml: 100 }),
        cell(p([kreuz("richtig", ok), s("   "), kreuz("falsch", !ok)], { after: 0, align: AlignmentType.CENTER }), { w: KW, borders: bd, valign: VerticalAlign.CENTER }),
      ]),
      row([cell(sol ? [
        p([s("Beleg: ", { bold: true, size: 19, color: MUTED }), s(z + "  " + zit, { size: 19, italics: true, color: SOLC, bold: true })], { after: 20 }),
        ...(r ? [p([s("Richtig ist: ", { bold: true, size: 19, color: MUTED }), s(r, { size: 19, italics: true, color: SOLC, bold: true })], { after: 0 })] : []),
      ] : [
        p([s("Beleg: ", { bold: true, size: 19, color: MUTED }), s("Z. ", { size: 19 }), gapL(7), s("  „", { size: 19 }), gapL(62), s("“", { size: 19 })], { after: 0, line: 380 }),
        p([s("Richtig ist: ", { bold: true, size: 19, color: MUTED }), gapL(86)], { after: 0, line: 380 }),
      ], { w: PW, borders: bd, m: 50, ml: 130, span: 3 })]),
    ];
    } finally { sol = sol0; }
  };
  return [
    p([t("Sehen mit den Fingerkuppen – Aussagen prüfen", { font: "Cambria", bold: true, size: 30, color: NAVY }), t(sol ? "      LÖSUNG" : "      Name: ____________________", { size: 18, color: sol ? SOLC : MUTED, bold: sol })], { after: 30 }),
    p(s("Reportage von Hilda Bermech, LB S. 183–184", { size: 19, italics: true, color: MUTED }), { after: 60 }),
    p([s("1  ", { bold: true, color: GRN }), s("Prüfe jede Aussage allein: Kreuze an.   "), s("2  ", { bold: true, color: GRN }), s("Belege mit Zeile und Zitat.   "), s("3  ", { bold: true, color: GRN }), s("Ist die Aussage falsch? Schreibe auf, was richtig ist.")], { after: PLm ? 20 : 80 }),
    ...(PLm ? [plp("PL: 1 Sprawdź każde zdanie sama: zaznacz richtig (prawda) albo falsch (fałsz). 2 Udowodnij: numer linijki (Z.) i dokładny cytat z tekstu w cudzysłowie „…“. 3 Jeśli zdanie jest fałszywe, napisz po niemiecku, jak jest naprawdę (Richtig ist: …). Zdanie 10 z gwiazdką (*) – jeśli skończysz wcześniej.")] : []),
    box([p([s("So belegst du: ", { bold: true, color: C.hv, size: 19 }), s("Zeile (Z.)", { bold: true, size: 19 }), s(" = am Rand gezählt (5, 10, 15 …) · ", { size: 19 }), s("Beleg", { bold: true, size: 19 }), s(" = Z. + Zitat in Anführungszeichen „…“", { size: 19 })], { after: 30 }),
         p([s("Beispiel: ", { bold: true, color: GRN, size: 19 }), s("„Die Ausstellung ist hell beleuchtet.“  ", { size: 19, italics: true }), s("☒ falsch", { size: 19, bold: true }), s("  ·  Beleg: Z. 1, „Herzlich willkommen in der Finsternis!“  ·  Richtig ist: Die Räume sind völlig dunkel.", { size: 19 })], { after: 0 })]),
    p(t(""), { after: 60 }),
    table([NW, AW, KW], AUS.flatMap((x, i) => item(i, x))),
    p([s("Zeit: ", { bold: true, size: 18 }), s("12 Minuten · allein · Nr. 10 mit *      ", { size: 18 }), s("Fertig? Stift hinlegen.", { size: 18, italics: true, color: MUTED })], { before: 80, after: 0 }),
    ...(PLm ? [
      p([t("Wörter im Text", { font: "Cambria", bold: true, size: 22, color: NAVY }), t("   trudne słowa (Z. = linijka)", { size: 18, italics: true, color: PLC })], { before: 160, after: 40 }),
      table([2600, PW / 2 - 2600, 2600, PW / 2 - 2600], Array.from({ length: Math.ceil(WORT.length / 2) }, (_, k) => row([WORT[2 * k], WORT[2 * k + 1]].flatMap((w) => w ? [cell(p(s(w[0], { bold: true, size: 18, color: NAVY }), { after: 0 }), { w: 2600, borders: bd, m: 20, ml: 60 }), cell(p(t(w[1], { size: 18, color: "333333" }), { after: 0 }), { w: PW / 2 - 2600, borders: bd, m: 20, ml: 60 })] : [cell(p(t("")), { w: 2600, borders: bd }), cell(p(t("")), { w: PW / 2 - 2600, borders: bd })])))),
    ] : []),
  ];
}
const props = { page: { size: { width: 11906, height: 16838 }, margin: { top: 420, bottom: 320, left: 550, right: 550 } } };
const styles = { default: { document: { run: { font: "Calibri", size: 21 } } } };
const save = (name, sol, PLm) => Packer.toBuffer(new Document({ styles, sections: [{ properties: props, children: blatt(sol, PLm) }] })).then((buf) => fs.writeFileSync(__dirname + "/../" + name + ".docx", buf));
save("AB_Reportage_Aussagen_pruefen", false);
save("AB_Reportage_Aussagen_pruefen_Loesung", true);
save("AB_Reportage_Aussagen_pruefen_Polnisch", false, true);
