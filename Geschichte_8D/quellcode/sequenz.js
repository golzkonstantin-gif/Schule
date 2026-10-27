const fs = require("fs");
const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, BorderStyle, ShadingType, AlignmentType, VerticalAlign, PageOrientation, PageBreak, Footer, PageNumber } = require("docx");

const NAVY = "1E2761", MUTED = "5B6B8C", LIGHT = "F4F7FD", LIGHT2 = "EAF0FA", RED = "D9534F", GOLD = "E8A33D", GREY = "E4E7EE";
const W = 15438;
const t = (text, o = {}) => new TextRun({ text, font: o.font || "Calibri", size: o.size || 20, bold: o.bold, italics: o.italics, color: o.color || "000000" });
const p = (runs, o = {}) => new Paragraph({ children: Array.isArray(runs) ? runs : [runs], alignment: o.align, spacing: { before: o.before ?? 0, after: o.after ?? 40 }, keepNext: o.keepNext });
const solid = (c, s = 6) => ({ style: BorderStyle.SINGLE, size: s, color: c });
const all = (b) => ({ top: b, bottom: b, left: b, right: b });
const WHITE = "FFFFFF";
function cell(children, o = {}) {
  return new TableCell({ children: Array.isArray(children) ? children : [children], width: { size: o.w, type: WidthType.DXA }, shading: o.fill ? { fill: o.fill, type: ShadingType.CLEAR, color: "auto" } : undefined, borders: o.borders || all(solid(WHITE, 12)), verticalAlign: o.valign || VerticalAlign.TOP, margins: { top: o.m ?? 70, bottom: o.m ?? 70, left: o.ml ?? 90, right: o.ml ?? 90 }, columnSpan: o.span });
}
const row = (cells, h) => new TableRow({ children: cells, height: h ? { value: h, rule: "atLeast" } : undefined, cantSplit: true });
const table = (widths, rows) => new Table({ width: { size: widths.reduce((a, b) => a + b, 0), type: WidthType.DXA }, columnWidths: widths, rows });
const h2 = (s) => p(t(s, { font: "Cambria", size: 30, bold: true, color: NAVY }), { before: 80, after: 80, keepNext: true });
const footer = new Footer({ children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ font: "Calibri", size: 16, color: MUTED, children: ["Geschichte 8D · Sequenz Französische Revolution · Stand 27.10.2026 · Seite ", PageNumber.CURRENT] })] })] });

// type: i = Inhalt, m = Methode im Mittelpunkt, t = Test, u = Urteil
const S = [
  { n: 1, d8: "Mi 04.11.", d8c: "Fr 16.10.", title: "Generalstände", sub: "Weg zu 1789, Abstimmung nach Ständen", ty: "i",
    ich: "… den Weg bis zu den Generalständen am Zeitstrahl erklären. Ich kann erklären, wie die Stände den Dritten Stand überstimmen konnten. Ich kann begründen, warum die Abstimmung nach Ständen unfair war.", pruef: "Zeitstrahl; Satz „unfair, weil …“", vor: "Text selbst erarbeiten, Ständespiel", komp: "Ursachen; Rekonstruktion" },
  { n: 2, d8: "Mi 11.11.", d8c: "Fr 06.11.", title: "Dritter Stand handelt", sub: "Nationalversammlung, Ballhausschwur", ty: "m", tag: "Karikatur",
    ich: "… erklären, warum sich der Dritte Stand 1789 zur Nationalversammlung erklärte und was er im Ballhausschwur versprach. Ich kann eine Karikatur analysieren: beschreiben, deuten, beurteilen.", pruef: "Kurzanalyse einer Karikatur", vor: "Methode direkt (Modell, dann angeleitet); Eingangsdiagnose", komp: "Epochenjahr 1789; Karikatur" },
  { n: 3, d8: "Mi 18.11.", d8c: "Fr 13.11.", title: "Bastille, Vorrechte, Menschenrechte", sub: "14. Juli, 4. August, 26. August", ty: "i",
    ich: "… erklären, warum der Sturm auf die Bastille zum Symbol wurde. Ich kann zwei Artikel der Menschenrechtserklärung mit eigenen Worten wiedergeben.", pruef: "Zwei Artikel erklären und einem Stand zuordnen", vor: "Textauszüge, Rollenkarten (Standeskarten)", komp: "Menschenrechtserklärung" },
  { n: 4, d8: "Mi 25.11.", d8c: "Fr 20.11.", title: "Verfassung 1791", sub: "konstitutionelle Monarchie", ty: "m", tag: "Schaubild",
    ich: "… am Schaubild erklären, wer in der konstitutionellen Monarchie welche Macht hatte. Ich kann erklären, wer 1791 wählen durfte und wer nicht.", pruef: "Fehler im Schaubild finden", vor: "Methode direkt; kurze Wahl mit Steuerkarten", komp: "Verfassungsschaubild" },
  { n: 5, d8: "Mi 02.12.", d8c: "Fr 27.11.", title: "Test 1", sub: "Stunden 1–4", ty: "t",
    ich: "… die Ursachen, die Ereignisse von 1789, die Menschenrechte und die Verfassung 1791 erklären.", pruef: "Test (max. 2 Seiten, 35 Min.)", vor: "Wahr/falsch mit Begründung, Zuordnung, Karikatur kurz", komp: "alle Bereiche" },
  { n: 6, d8: "Mi 09.12.", d8c: "Fr 04.12.", title: "Vom König zur Republik", sub: "Flucht, Krieg, Hinrichtung", ty: "i",
    ich: "… erklären, warum die konstitutionelle Monarchie 1792 scheiterte. Ich kann die Ereignisse von 1791 bis 1793 in der richtigen Reihenfolge nennen.", pruef: "Ursache-Folge-Kette in eigenen Sätzen", vor: "Ereignisse erst zu zweit ordnen, dann vergleichen; Berichtigung Test 1", komp: "Zusammenhänge rekonstruieren" },
  { n: 7, d8: "Mi 16.12.", d8c: "Fr 11.12.", title: "Jakobinerherrschaft, Terror", sub: "Wohlfahrtsausschuss, Guillotine", ty: "u", tag: "Urteil",
    ich: "… ein Sachurteil zum Terror der Jakobiner mit Beleg formulieren. Ich kann ein Werturteil dazu mit offengelegtem Maßstab begründen.", pruef: "Zwei Sätze: Sachurteil mit Beleg, Werturteil mit Maßstab", vor: "Urteil direkt (Modell, Satzanfänge in 3 Stufen)", komp: "Sach- und Werturteil" },
  { n: 8, d8: "Mi 06.01.", d8c: "Fr 18.12.", title: "Ende der Revolution, Napoleon", sub: "1794–1799, Staatsstreich", ty: "i",
    ich: "… den Weg vom Ende des Terrors 1794 bis zum Staatsstreich 1799 am Zeitstrahl darstellen. Ich kann zwei Gründe nennen, warum viele Menschen 1799 einen starken Mann akzeptierten.", pruef: "Zwei begründete Ursachen", vor: "Text und Zeitstrahl", komp: "Zusammenhänge rekonstruieren" },
  { n: 9, d8: "Mi 13.01.", d8c: "Fr 08.01.", title: "Napoleon: Reformen und Fall", sub: "Erbe der Revolution oder Verrat?", ty: "u", tag: "Urteil",
    ich: "… zwei Reformen Napoleons nennen. Ich kann zwei Gründe für sein Ende erklären. Ich kann mit Hilfen ein Urteil zu „Erbe der Revolution oder Verrat?“ schreiben.", pruef: "Zeitstrahl; Urteil mit Kriterien", vor: "Text; angeleitetes Urteil (Kriterien für Test 2)", komp: "Urteil (angeleitet)" },
  { n: 10, d8: "Mi 20.01.", d8c: "Fr 15.01.", title: "Test 2", sub: "Stunden 6–9, Urteilsaufgabe", ty: "t",
    ich: "… die Ereignisse von 1792 bis 1815 erklären. Ich kann ein begründetes Urteil schreiben.", pruef: "Test mit Urteilsaufgabe (Kriterien bekannt)", vor: "Wie Test 1, plus Verfassungsschaubild als Wiederholung", komp: "Urteil (selbstständig)" },
  { n: 11, d8: "Mi 27.01.", d8c: "Fr 22.01.", title: "Rückgabe, Gesamturteil", sub: "Wiener Kongress*", ty: "u", tag: "Urteil",
    ich: "… ein begründetes Urteil zu „War die Revolution ein Erfolg?“ schreiben. Ich kann es mit meiner Eingangsseite aus Stunde 2 vergleichen.", pruef: "Gesamturteil; Vorher-Nachher-Vergleich", vor: "Berichtigung mit Grün, angeleitetes Schreiben", komp: "Urteil (Reihenabschluss)" },
];
const FILL = { i: LIGHT2, m: "FDF1D8", t: "F8D9D7", u: "E3F3EC" };
const EDGE = { i: NAVY, m: GOLD, t: RED, u: "2E9E6B" };
const KIND = { i: "Inhalt", m: "Methode", t: "Test", u: "Urteil" };

const d = [];
// ---- Seite 1: Titel, Leitfrage, Zeitleiste ----
d.push(p(t("GESCHICHTE · KLASSE 8D · SCHIC KLASSE 8, MODUL 2", { size: 18, bold: true, color: MUTED }), { after: 20 }));
d.push(p(t("Sequenz: Die Französische Revolution", { font: "Cambria", size: 52, bold: true, color: NAVY }), { after: 60 }));
d.push(table([5600, 5000, 4838], [row([
  cell([p(t("Leitfrage der Reihe", { bold: true, size: 18, color: MUTED })), p(t("Wie kam es zur Französischen Revolution, und war sie ein Erfolg?", { font: "Cambria", size: 26, bold: true, color: NAVY }), { after: 0 })], { w: 5600, fill: LIGHT, borders: all(solid(NAVY, 12)), m: 110, ml: 160 }),
  cell([p(t("Rahmen", { bold: true, size: 18, color: MUTED })), p(t("1 Stunde pro Woche, mittwochs · 11 Stunden, 04.11.2026 bis 27.01.2027, keine Reserve · 2 Tests", { size: 21 }), { after: 0 })], { w: 5000, fill: LIGHT, borders: all(solid(NAVY, 12)), m: 110, ml: 160 }),
  cell([p(t("8C (freitags)", { bold: true, size: 18, color: MUTED })), p(t("Immer eine Stunde vorn: 8C-Stunde n liegt vor 8D-Stunde n. Rückmeldungen der 8C fließen in die 8D-Planung ein.", { size: 21 }), { after: 0 })], { w: 4838, fill: LIGHT, borders: all(solid(NAVY, 12)), m: 110, ml: 160 }),
])]));
d.push(p(t(""), { after: 100 }));
d.push(h2("Zeitleiste: 11 Stunden bis zum Halbjahresende"));

const order = [1, 2, 3, 4, 5, 6, 7, "F", 8, 9, 10, 11];
const CWID = Math.floor(W / 12);
const widths = order.map((o) => (o === "F" ? W - CWID * 11 : CWID));
const cellsRow = order.map((o, i) => {
  if (o === "F") return cell([p(t("Weihnachts-ferien", { bold: true, size: 18, color: MUTED }), { align: AlignmentType.CENTER, after: 20 }), p(t("23.12. bis 02.01.", { size: 16, color: MUTED }), { align: AlignmentType.CENTER, after: 0 })], { w: widths[i], fill: GREY, valign: VerticalAlign.CENTER, borders: all(solid("5B6B8C", 8)) });
  const s = S.find((x) => x.n === o);
  return cell([
    p([t(String(s.n), { font: "Cambria", size: 40, bold: true, color: EDGE[s.ty] })], { after: 0 }),
    p(t(s.d8, { bold: true, size: 18 }), { after: 0 }),
    p(t("8C " + s.d8c.replace("Fr ", ""), { size: 15, color: MUTED }), { after: 50 }),
    p(t(s.title.replace("Generalstände","General\u00ADstände").replace("Menschenrechte","Menschen\u00ADrechte").replace("Jakobinerherrschaft","Jakobiner\u00ADherrschaft").replace("Verfassung","Verfas\u00ADsung"), { font: "Cambria", bold: true, size: 17, color: NAVY }), { after: 30 }),
    p(t(s.sub.replace("Nationalversammlung","National\u00ADversammlung").replace("Wohlfahrtsausschuss","Wohlfahrts\u00ADausschuss").replace("konstitutionelle","konstitu\u00ADtionelle"), { size: 15 }), { after: 40 }),
    p(t(s.tag ? KIND[s.ty] + ": " + s.tag : KIND[s.ty], { bold: true, size: 15, color: EDGE[s.ty] }), { after: 0 }),
  ], { w: widths[i], fill: FILL[s.ty], borders: { top: solid(EDGE[s.ty], 36), bottom: solid("5B6B8C", 8), left: solid("5B6B8C", 8), right: solid("5B6B8C", 8) }, m: 90, ml: 60 });
});
d.push(table(widths, [row(cellsRow, 4300)]));
d.push(p(t(""), { after: 40 }));
d.push(table([2500, 2500, 2500, 2500, 5438], [row([
  ...["i", "m", "t", "u"].map((k, i) => cell(p([t("■ ", { color: EDGE[k], size: 22 }), t(KIND[k] + (k === "i" ? " (Inhalt und Ereignisse)" : k === "m" ? " (neue Methode)" : k === "t" ? " (Leistungskontrolle)" : " (Sach- und Werturteil)"), { size: 17 })], { after: 0 }), { w: 2500, fill: WHITE, borders: all(solid(WHITE, 4)), m: 30 })),
  cell(p(t("* mit Stern: wird nur bei Zeit behandelt (Wiener Kongress)", { size: 17, italics: true, color: MUTED }), { after: 0 }), { w: 5438, fill: WHITE, borders: all(solid(WHITE, 4)), m: 30 }),
])]));
d.push(p(t(""), { after: 60 }));
d.push(table([7600, 7838], [row([
  cell([p(t("Eingangsdiagnose und Vorher-Nachher", { bold: true, size: 20, color: NAVY })), p(t("Stunde 2, Ankommen: zwei Sätze im Heft zu „Was ist eine Revolution? Ist sie gut oder schlecht? Warum?“ Stunde 11: Vergleich mit dem begründeten Gesamturteil.", { size: 19 }), { after: 0 })], { w: 7600, fill: LIGHT, borders: all(solid("B9C6E8", 8)), m: 90, ml: 140 }),
  cell([p(t("Rückfall bei Ausfall (Streichkandidaten)", { bold: true, size: 20, color: NAVY })), p(t("1. Wiener Kongress nach Modul 3 verschieben · 2. Napoleons Fall nur kurz · 3. Bastille als Zeitstrahl-Eintrag statt eigener Erarbeitung.", { size: 19 }), { after: 0 })], { w: 7838, fill: LIGHT, borders: all(solid("B9C6E8", 8)), m: 90, ml: 140 }),
])]));

// ---- Seite 2: Detailtabelle ----
d.push(new Paragraph({ children: [new PageBreak()] }));
d.push(h2("Die Stunden im Überblick"));
const DW = [700, 1400, 5900, 2400, 2800, 2238];
const hdr = ["Nr.", "Termin 8D", "Ich kann …", "Woran erkenne ich es?", "Vorgehen", "SchiC / Kompetenz"];
const dRows = [row(hdr.map((x, i) => cell(p(t(x, { bold: true, color: WHITE, size: 18 }), { after: 0 }), { w: DW[i], fill: NAVY, valign: VerticalAlign.CENTER, m: 60, borders: all(solid(NAVY, 8)) })), 330)];
S.forEach((s) => {
  const f = s.ty === "t" ? "FCEFEE" : s.n % 2 ? WHITE : LIGHT; const mm = 40;
  dRows.push(row([
    cell(p(t(String(s.n), { font: "Cambria", size: 26, bold: true, color: EDGE[s.ty] }), { after: 0 }), { w: DW[0], fill: f, borders: { top: solid("5B6B8C", 8), bottom: solid("5B6B8C", 8), left: solid(EDGE[s.ty], 36), right: solid("5B6B8C", 8) }, m: mm }),
    cell([p(t(s.d8, { bold: true, size: 18 }), { after: 0 }), p(t("8C: " + s.d8c, { size: 15, color: MUTED }), { after: 0 })], { w: DW[1], fill: f, m: mm, borders: all(solid("5B6B8C", 8)) }),
    cell([p(t(s.title, { font: "Cambria", bold: true, size: 19, color: NAVY }), { after: 20 }), p(t("Ich kann " + s.ich.replace(/^… /, ""), { size: 16 }), { after: 0 })], { w: DW[2], fill: f, m: mm, borders: all(solid("5B6B8C", 8)) }),
    cell(p(t(s.pruef, { size: 16 }), { after: 0 }), { w: DW[3], fill: f, m: mm, borders: all(solid("5B6B8C", 8)) }),
    cell(p(t(s.vor, { size: 16 }), { after: 0 }), { w: DW[4], fill: f, m: mm, borders: all(solid("5B6B8C", 8)) }),
    cell(p(t(s.komp, { size: 16 }), { after: 0 }), { w: DW[5], fill: f, m: mm, borders: all(solid("5B6B8C", 8)) }),
  ]));
});
d.push(table(DW, dRows));

// ---- Seite 3: Kompetenzmatrix ----
d.push(new Paragraph({ children: [new PageBreak()] }));
d.push(h2("Kompetenzaufbau: Wann wird was gezeigt, geübt, selbstständig gekonnt?"));
const LW = 3000, MW = Math.floor((W - LW) / 11), MLAST = W - LW - MW * 10;
const mw = [LW, ...Array(10).fill(MW), MLAST];
const MARK = { M: ["M", GOLD, "000000"], a: ["a", "CADCFC", "000000"], s: ["s", NAVY, WHITE] };
const MAT = [
  ["Ursachen / Zusammenhänge rekonstruieren", { 1: "a", 6: "a", 8: "s", 10: "s" }],
  ["Karikatur analysieren", { 2: "M", 5: "s" }],
  ["Quelle lesen: Menschenrechte", { 3: "a" }],
  ["Verfassungsschaubild lesen", { 4: "M", 10: "s" }],
  ["Sachurteil und Werturteil", { 1: "a", 7: "M", 9: "a", 10: "s", 11: "s" }],
];
const mRows = [row([cell(p(t("Kompetenz", { bold: true, color: WHITE, size: 18 }), { after: 0 }), { w: LW, fill: NAVY, valign: VerticalAlign.CENTER, m: 60, borders: all(solid(NAVY, 8)) }), ...S.map((s, i) => cell(p(t(String(s.n), { bold: true, color: WHITE, size: 20 }), { align: AlignmentType.CENTER, after: 0 }), { w: mw[i + 1], fill: s.ty === "t" ? RED : NAVY, valign: VerticalAlign.CENTER, m: 60, borders: all(solid(WHITE, 8)) }))], 340)];
MAT.forEach(([lab, m], r) => mRows.push(row([
  cell(p(t(lab, { size: 19, bold: true }), { after: 0 }), { w: LW, fill: r % 2 ? WHITE : LIGHT, valign: VerticalAlign.CENTER, m: 80, borders: all(solid("5B6B8C", 8)) }),
  ...S.map((s, i) => {
    const mk = m[s.n];
    if (!mk) return cell(p(t(" ", { size: 18 }), { after: 0 }), { w: mw[i + 1], fill: s.ty === "t" ? "FCEFEE" : r % 2 ? WHITE : LIGHT, m: 80, borders: all(solid("5B6B8C", 8)) });
    const [lbl, fill, col] = MARK[mk];
    return cell(p(t(lbl, { bold: true, size: 24, color: col }), { align: AlignmentType.CENTER, after: 0 }), { w: mw[i + 1], fill, valign: VerticalAlign.CENTER, m: 80, borders: all(solid("5B6B8C", 8)) });
  }),
], 480)));
d.push(table(mw, mRows));
d.push(p([t("M", { bold: true, size: 20 }), t(" = Lehrkraft zeigt das Modell (laut denken)    ", { size: 19 }), t("a", { bold: true, size: 20 }), t(" = angeleitet mit Hilfen    ", { size: 19 }), t("s", { bold: true, size: 20 }), t(" = selbstständig    Rot: Test", { size: 19 })], { before: 60, after: 160 }));

d.push(table([5000, 5200, 5238], [row([
  cell([p(t("Für jede Stunde gilt (Planungsstandard)", { bold: true, size: 21, color: NAVY }), { after: 40 }), ...["Ziel „Ich kann …“, Prüfaufgabe, Mindestziel, Beitrag zur Reihe", "AVIVA, „Auswerten“ nie gestrichen", "Echte Wiederholung der Vorstunde am Anfang", "Warum-Zeile und Gelenkstelle je Phase"].map((x) => p([t("• ", { size: 19 }), t(x, { size: 19 })], { after: 30 }))], { w: 5000, fill: LIGHT, borders: all(solid("B9C6E8", 8)), m: 100, ml: 150 }),
  cell([p(t("Für die 8D (laut, 25 Schülerinnen und Schüler)", { bold: true, size: 21, color: NAVY }), { after: 40 }), ...["Wenig Plenumsgespräch, Aufträge schriftlich", "Gezielte Abfrage (3–6 Personen) im Wechsel mit Heft-Stichprobe, unbenotet", "Wörtliche Lautstärke-Ansagen, feste Zeiten", "Spiele mit Rollen, Standeskarten wiederverwenden"].map((x) => p([t("• ", { size: 19 }), t(x, { size: 19 })], { after: 30 }))], { w: 5200, fill: LIGHT, borders: all(solid("B9C6E8", 8)), m: 100, ml: 150 }),
  cell([p(t("Kopplung mit der 8C", { bold: true, size: 21, color: NAVY }), { after: 40 }), ...["Dieselben Stundennummern, 8C freitags eine Stunde vorn", "Vor jeder 8D-Planung: Reflexionslog der 8C lesen", "Status in der Sequenztabelle (Markdown-Datei)", "Reserve nur bei der 8C (29.01.)"].map((x) => p([t("• ", { size: 19 }), t(x, { size: 19 })], { after: 30 }))], { w: 5238, fill: LIGHT, borders: all(solid("B9C6E8", 8)), m: 100, ml: 150 }),
])]));

(async () => {
  const doc = new Document({ styles: { default: { document: { run: { font: "Calibri", size: 20 } } } }, sections: [{ properties: { page: { size: { width: 11906, height: 16838, orientation: PageOrientation.LANDSCAPE }, margin: { top: 650, bottom: 600, left: 700, right: 700 } } }, footers: { default: footer }, children: d }] });
  fs.writeFileSync("../Sequenz_Franzoesische_Revolution_8D.docx", await Packer.toBuffer(doc));
  console.log("ok");
})();
