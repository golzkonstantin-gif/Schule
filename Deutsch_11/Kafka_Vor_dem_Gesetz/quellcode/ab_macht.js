// Arbeitsblatt „Vor dem Gesetz“ – Macht und Ohnmacht (Neustart, 45 Min.) im Muster der Mentorin.
// Aufruf: node ab_macht.js        → leeres Arbeitsblatt
//         node ab_macht.js loesung → ausgefülltes Arbeitsblatt (Lösung für die Lehrkraft)
const fs = require("fs");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, BorderStyle,
  ShadingType, AlignmentType, VerticalAlign, HeightRule, Header, Footer, TabStopType,
} = require("docx");
const D = require("./daten_macht");
const LOES = process.argv[2] === "loesung";

const NAVY = "1E2761", MUTED = "5B6B8C", LIGHT = "F4F7FD", LIGHT2 = "EAF0FA", RED = "D9534F", BLUE = "3F7CC4", GREEN = "2E9E6B";
const T_FILL = "FBD9D7", M_FILL = "D6E6F8";
const W = 10206;

const t = (text, o = {}) => new TextRun(Object.assign({ text, font: o.font || "Calibri", size: o.size || 21, bold: o.bold, italics: o.italics, color: o.color || "000000" }, o.run || {}));
const g = (text, o = {}) => t(text, Object.assign({ color: GREEN, size: 20 }, o)); // Lösungstext
const p = (runs, o = {}) => new Paragraph({ children: Array.isArray(runs) ? runs : [runs], alignment: o.align, pageBreakBefore: o.pb, keepNext: o.keepNext, spacing: { before: o.before ?? 0, after: o.after ?? 60, line: o.line } });
const none = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
const solid = (c, s = 6) => ({ style: BorderStyle.SINGLE, size: s, color: c });
const all = (b) => ({ top: b, bottom: b, left: b, right: b });
const cell = (children, o = {}) => new TableCell({
  children: Array.isArray(children) ? children : [children], width: { size: o.w, type: WidthType.DXA },
  shading: o.fill ? { fill: o.fill, type: ShadingType.CLEAR, color: "auto" } : undefined,
  borders: o.borders || all(solid("B9C6E8")), verticalAlign: o.valign || VerticalAlign.TOP,
  margins: { top: o.m ?? 50, bottom: o.m ?? 50, left: o.ml ?? 100, right: o.ml ?? 100 }, columnSpan: o.span,
});
const table = (widths, rows) => new Table({ width: { size: widths.reduce((a, b) => a + b, 0), type: WidthType.DXA }, columnWidths: widths, rows });
const row = (cells, h) => new TableRow({ children: cells, height: h ? { value: h, rule: HeightRule.ATLEAST } : undefined, cantSplit: true });
const kopf = (x, w) => cell(p(t(x, { bold: true, color: "FFFFFF", size: 18 }), { after: 0 }), { w, fill: NAVY });
const h2 = (text, o = {}) => p(t(text, { font: "Cambria", size: 26, bold: true, color: NAVY }), { before: o.before ?? 160, after: 60, pb: o.pb, keepNext: true });
const aufgabe = (a) => [
  p(t(a.titel, { bold: true, color: NAVY, size: 22 }), { before: 160, after: 40, keepNext: true }),
  ...a.schritte.map((x, i) => p([t((i + 1) + ".  ", { bold: true, color: NAVY, size: 20 }), t(x, { size: 20 })], { after: 30, keepNext: true })),
];
const linien = (n, w = W, h = 400) => table([w], Array.from({ length: n }, () => row([cell(p(t(""), { after: 0 }), { w, borders: { top: none, left: none, right: none, bottom: solid("9AA6C4", 4) }, m: 0, ml: 0 })], h)));
// Schreibfeld: leer = Linien, Lösung = grüner Text über einer Linie
const feld = (n, text, w = W) => LOES
  ? table([w], [row([cell(p(g(text), { after: 0, line: 300 }), { w, borders: { top: none, left: none, right: none, bottom: solid("9AA6C4", 4) }, m: 40, ml: 0 })], n * 400)])
  : linien(n, w);

// ---------- Text mit Markierungen ----------
function textBlock() {
  const full = D.ZEILEN.join("\n");
  const cls = new Array(full.length).fill("");
  const esc = (s) => s.replace(/[.*+?^${}()|[\]\\]/g, "\\$&").replace(/ /g, "\\s+");
  let pos = 0;
  if (LOES) D.MARKIERUNG.forEach(([s, c]) => {
    const re = new RegExp(esc(s), "g");
    re.lastIndex = pos; let m = re.exec(full);
    if (!m) { re.lastIndex = 0; m = re.exec(full); }
    for (let i = m.index; i < m.index + m[0].length; i++) cls[i] = c;
    pos = m.index + m[0].length;
  });
  const nW = 480, tW = W - nW;
  let off = 0;
  const rows = D.ZEILEN.map((l, i) => {
    const runs = [];
    let k = 0;
    while (k < l.length) {
      const c = cls[off + k]; let e = k;
      while (e < l.length && cls[off + e] === c) e++;
      const o = { font: "Cambria", size: 21 };
      if (c === "T") o.run = { shading: { type: ShadingType.CLEAR, fill: T_FILL, color: "auto" } };
      if (c === "M") o.run = { shading: { type: ShadingType.CLEAR, fill: M_FILL, color: "auto" } };
      if (c === "TW" || c === "MW") { o.bold = true; o.run = { border: { style: BorderStyle.SINGLE, size: 8, color: c === "TW" ? RED : BLUE, space: 1 } }; }
      runs.push(t(l.slice(k, e), o));
      k = e;
    }
    off += l.length + 1;
    return row([
      cell(p(t((i + 1) % 5 === 0 || i === 0 ? String(i + 1) : "", { size: 16, color: MUTED }), { align: AlignmentType.RIGHT, after: 0 }), { w: nW, borders: { top: none, bottom: none, left: none, right: solid(BLUE, 10) }, m: 4, ml: 80, valign: VerticalAlign.CENTER }),
      cell(p(runs, { after: 0, line: 290 }), { w: tW, borders: all(none), m: 4, ml: 140 }),
    ]);
  });
  return table([nW, tW], rows);
}
const legende = p([
  t("Lösung: ", { bold: true, size: 18, color: GREEN }),
  t(" Türhüter: was ", { size: 18, run: { shading: { type: ShadingType.CLEAR, fill: T_FILL, color: "auto" } } }), t("  ", { size: 18 }),
  t("wie", { size: 18, bold: true, run: { border: { style: BorderStyle.SINGLE, size: 8, color: RED, space: 1 } } }), t("     ", { size: 18 }),
  t(" Mann: was ", { size: 18, run: { shading: { type: ShadingType.CLEAR, fill: M_FILL, color: "auto" } } }), t("  ", { size: 18 }),
  t("wie", { size: 18, bold: true, run: { border: { style: BorderStyle.SINGLE, size: 8, color: BLUE, space: 1 } } }),
  t("   → Der Türhüter spricht (wörtlich), der Mann bittet (meist indirekt).", { size: 18, color: GREEN }),
], { after: 60 });

// ---------- Aufgabe 2: Einleitung ----------
const tW2 = [2400, W - 2400];
const tatte = table(tW2, D.TATTE.map(([k, v]) => row([
  cell(p([t(k[0], { bold: true, color: RED, size: 20 }), t(k.slice(1), { bold: true, color: NAVY, size: 20 })], { after: 0 }), { w: tW2[0], fill: LIGHT, valign: VerticalAlign.CENTER }),
  cell(p(LOES ? g(v) : t(""), { after: 0 }), { w: tW2[1], valign: VerticalAlign.CENTER }),
], 430)));
const label = (x) => p(t(x, { bold: true, size: 19, color: NAVY }), { before: 100, after: 20, keepNext: true });

// ---------- Aufgabe 3: Beobachtungsbogen ----------
const bW = [2300, (W - 2300) / 2, (W - 2300) / 2];
const beob = table(bW, [
  row([kopf("Merkmal", bW[0]), kopf("Der Mann vom Lande", bW[1]), kopf("Der Türhüter", bW[2])]),
  ...D.BEOBACHTUNG.map(([m, mann, th]) => row([
    cell(p(t(m, { bold: true, size: 19, color: NAVY }), { after: 0 }), { w: bW[0], fill: LIGHT }),
    cell(p(LOES ? g(mann, { size: 19 }) : t(""), { after: 0 }), { w: bW[1] }),
    cell(p(LOES ? g(th, { size: 19 }) : t(""), { after: 0 }), { w: bW[2] }),
  ], 900)),
]);

// ---------- Aufgabe 4: Gründe ----------
const gruende = D.GRUENDE.flatMap(([k, v, anker], i) => [
  p([t((i + 1) + "  ", { bold: true, color: RED, size: 20 }), t(k, { bold: true, size: 20 })], { before: 80, after: 20, keepNext: true }),
  feld(2, v + (anker.length ? " (vgl. " + anker.map(D.zeile).join("; ") + ")" : "")),
]);

// ---------- Aufgabe 5*: Stilmittel zuordnen ----------
const sW = [3500, 1000, W - 3500 - 1000 - 1500, 1500];
const buchst = Object.keys(D.DEUTUNGEN);
const stil = table(sW, [
  row([kopf("Stilmittel und Textstelle", sW[0]), kopf("Buchstabe", sW[1]), kopf("Deutung", sW[2]), kopf("Macht oder Ohnmacht?", sW[3])]),
  ...D.STILMITTEL.map(([sm, stelle, anker, lsg, mo], i) => row([
    cell([p(t((i + 1) + ". " + sm, { bold: true, color: RED, size: 19 }), { after: 10 }), p([t(stelle, { size: 18, italics: true }), t("  (vgl. " + D.zeile(anker) + ")", { size: 17, color: MUTED })], { after: 0 })], { w: sW[0], fill: LIGHT }),
    cell(p(LOES ? g(lsg, { bold: true, size: 24 }) : t(""), { after: 0, align: AlignmentType.CENTER }), { w: sW[1], valign: VerticalAlign.CENTER }),
    cell(p([t(buchst[i] + "  ", { bold: true, color: NAVY, size: 19 }), t(D.DEUTUNGEN[buchst[i]], { size: 18 })], { after: 0 }), { w: sW[2] }),
    cell(p(LOES ? g(mo, { bold: true, size: 19 }) : t(""), { after: 0, align: AlignmentType.CENTER }), { w: sW[3], valign: VerticalAlign.CENTER }),
  ], 900)),
]);

const header = new Header({ children: [new Paragraph({ tabStops: [{ type: TabStopType.RIGHT, position: W }], children: [t("Deutsch · Klasse 11 · Kafka, „Vor dem Gesetz“", { size: 17, color: MUTED }), LOES ? t("\tLÖSUNG (für die Lehrkraft)", { size: 17, bold: true, color: GREEN }) : t("\tName: ______________________   Datum: ___________", { size: 17, color: MUTED })] })] });
const footer = new Footer({ children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [t("Franz Kafka: „Vor dem Gesetz“ (1915), in heutiger Rechtschreibung" + (LOES ? " · Zeilenangaben ± 1 gelten als richtig" : ""), { size: 16, color: MUTED })] })] });

const seite1 = [
  p(t("Arbeitsblatt: Macht und Ohnmacht in Kafkas „Vor dem Gesetz“", { font: "Cambria", size: 32, bold: true, color: NAVY }), { after: 50 }),
  p([t("Leitfrage: ", { bold: true, size: 20 }), t(D.LEITFRAGE, { size: 20 })], { after: 20 }),
  p([t("Ziel: ", { bold: true, size: 20 }), t(D.ZIEL, { size: 20 })], { after: 40 }),
  ...aufgabe(D.AUFGABE1),
  ...(LOES ? [legende] : [p(t(""), { after: 40 })]),
  textBlock(),
];
const seite2 = [
  h2("Arbeitsaufträge", { pb: true, before: 0 }),
  ...aufgabe(D.AUFGABE2),
  tatte,
  label("Inhalt in drei Sätzen:"),
  feld(4, D.EINLEITUNG.inhalt),
  label("Deutungshypothese:"),
  ...(LOES ? [] : [p(t(D.EINLEITUNG.geruest, { size: 19, italics: true, color: BLUE }), { after: 0 })]),
  feld(3, D.EINLEITUNG.hypothese),
  ...aufgabe(D.AUFGABE5),
  stil,
  ...(LOES ? [p(g(D.STIL_ERGEBNIS, { size: 19 }), { before: 80 })] : []),
];
// Das Gespräch im Spiel (Beobachtungsbogen) wird mündlich besprochen – nicht auf dem Arbeitsblatt (Entscheidung 18.11.)
const seite3 = [
  p(t(""), { pb: true, after: 0 }),
  ...aufgabe(D.AUFGABE4),
  ...gruende,
  label("Ergebnis: Stimmt unsere Deutungshypothese?"),
  feld(2, D.ERGEBNIS),
];

const doc = new Document({
  styles: { default: { document: { run: { font: "Calibri", size: 21 } } } },
  sections: [{ properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 760, bottom: 620, left: 850, right: 850, header: 340, footer: 340 } } }, headers: { default: header }, footers: { default: footer }, children: [...seite1, ...seite2, ...seite3] }],
});
const name = "../Macht_Ohnmacht/Arbeitsblatt_Macht_Ohnmacht" + (LOES ? "_Loesung" : "") + ".docx";
Packer.toBuffer(doc).then((b) => fs.writeFileSync(name, b));
