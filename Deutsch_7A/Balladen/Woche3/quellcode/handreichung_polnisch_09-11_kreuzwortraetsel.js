// Kreuzworträtsel Balladen für S-PL (Mo 9.11.), A4 hoch, ein Schnipsel: Hinweise deutsch + polnisch, Wörter bleiben deutsch, ohne Lösung
// Hierarchische Nummerierung (vom Allgemeinen ins Detail), keine Farbflächen (Tinte sparen)
// Aufruf: node schnipsel_kreuzwortraetsel_balladen.js   (liest kw.json, erzeugt docx im Elternordner)
const fs = require("fs");
const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, BorderStyle, AlignmentType, VerticalAlign, HeightRule, PageOrientation } = require("docx");
const NAVY = "1E2761", MUTED = "5B6B8C", RED = "D9534F", VIO = "7A5BA6", GRAY = "9AA3B5";
const KW = JSON.parse(fs.readFileSync(__dirname + "/kw.json", "utf8"));
const t = (text, o = {}) => new TextRun({ text, font: o.font || "Calibri", size: o.size || 19, bold: o.bold, italics: o.it, color: o.color || "000000" });
const p = (ch, o = {}) => new Paragraph({ children: Array.isArray(ch) ? ch : [ch], alignment: o.align, spacing: { before: o.before || 0, after: o.after ?? 60, line: o.line }, indent: o.indent });
const B = (style, size, color) => ({ style, size, color });
const borders = (b) => ({ top: b, bottom: b, left: b, right: b });
const NONE = borders(B(BorderStyle.NONE, 0, "FFFFFF"));
const cell = (children, o = {}) => new TableCell({ children: Array.isArray(children) ? children : [children], width: { size: o.w, type: WidthType.DXA }, borders: o.borders || NONE, verticalAlign: o.valign || VerticalAlign.TOP, margins: { top: o.m ?? 0, bottom: o.m ?? 0, left: o.ml ?? o.m ?? 0, right: o.ml ?? o.m ?? 0 } });
const table = (cols, rows) => new Table({ columnWidths: cols, width: { size: cols.reduce((a, b) => a + b, 0), type: WidthType.DXA }, rows });

// Hierarchie: [Wort, Ebene, Hinweis]
const PL = {
  BALLADE: "wiersz, który opowiada napiętą historię",
  LYRIK: "rodzaj literacki: wiersze (jak tekst piosenki)",
  STROPHE: "część wiersza; między nimi jest pusta linijka",
  VERS: "jedna linijka wiersza",
  REIM: "taki sam dźwięk na końcu wersu",
  PAARREIM: "układ rymów aabb",
  KREUZREIM: "układ rymów abab",
  UMARMENDER: "układ abba – wpisz tylko przymiotnik",
  KEHRREIM: "wers, który powtarza się na końcu każdej strofy",
  EPIK: "rodzaj literacki: opowiadania (jak relacja z meczu)",
  "ERZÄHLER": "kto w epice opowiada (forma „er“ albo „ich“)",
  DRAMATIK: "rodzaj literacki: sztuki teatralne (jak scena z serialu)",
  DIALOG: "rozmowa postaci, zupełnie bez narratora",
};
const PLC = "5B4A8A";
const H = [
  ["BALLADE", 0, "Gedicht, das eine spannende Geschichte erzählt, z. B. „Paddy Fingal“"],
  ["LYRIK", 1, "Gattung der Gedichte – wie ein Song-Text"],
  ["STROPHE", 2, "Abschnitt eines Gedichts; dazwischen steht eine Leerzeile"],
  ["VERS", 2, "eine Zeile in einem Gedicht"],
  ["REIM", 2, "gleicher Klang am Versende, z. B. Bart – Riesenart"],
  ["PAARREIM", 3, "Reimschema aabb"],
  ["KREUZREIM", 3, "Reimschema abab"],
  ["UMARMENDER", 3, "Reimschema abba: der ___ Reim"],
  ["KEHRREIM", 2, "Vers, der am Ende jeder Strophe wiederkehrt"],
  ["EPIK", 1, "Gattung der Erzählungen – wie ein Spielbericht"],
  ["ERZÄHLER", 2, "vermittelt in der Epik das Geschehen (Er- oder Ich-Form)"],
  ["DRAMATIK", 1, "Gattung der Theaterstücke – wie eine Serienszene"],
  ["DIALOG", 2, "Gespräch zwischen Figuren, ganz ohne Erzähler"],
];
const NUM = {}; H.forEach(([w], i) => (NUM[w] = i + 1));
const DIR = {}; const startNum = {}; const letters = {};
KW.placed.forEach(([w, r, c, d]) => {
  DIR[w] = d; startNum[`${r},${c}`] = NUM[w];
  [...w].forEach((ch, i) => (letters[d === "H" ? `${r},${c + i}` : `${r + i},${c}`] = ch));
});

const CS = 520; // Kästchen ca. 0,9 cm
const kwB = borders(B(BorderStyle.SINGLE, 8, NAVY));
function grid(withSol) {
  const rows = [];
  for (let r = 0; r < KW.h; r++) {
    const cells = [];
    for (let c = 0; c < KW.w; c++) {
      const k = `${r},${c}`;
      if (!letters[k]) { cells.push(cell(p(t(""), { after: 0 }), { w: CS })); continue; }
      cells.push(cell([
        p(t(startNum[k] ? String(startNum[k]) : " ", { size: 13, bold: true, color: VIO }), { after: 0, line: 150 }),
        p(t(withSol ? letters[k] : " ", { size: 21, bold: true, color: withSol ? RED : NAVY }), { after: 0, align: AlignmentType.CENTER, line: 210 }),
      ], { w: CS, borders: kwB, m: 0, ml: 25 }));
    }
    rows.push(new TableRow({ children: cells, height: { value: CS, rule: HeightRule.EXACT } }));
  }
  return table(Array(KW.w).fill(CS), rows);
}
const SW = 10906; // A4 hoch
const line = () => p(t("_".repeat(96), { color: GRAY, size: 19 }), { before: 170, after: 0 });
function schnipsel(withSol) {
  const kids = [
    p([t("Kreuzworträtsel: Balladen", { font: "Cambria", bold: true, size: 26, color: NAVY }), t("     Krzyżówka: ballady – powtórka", { size: 17, it: true, color: PLC })], { after: 30 }),
    p([t("1  Trage die Wörter ein. Die Nummern gehen vom Allgemeinen ins Detail. ", { size: 18 }), t("→ waagerecht, ↓ senkrecht, Ä = ein Kästchen.", { size: 18, color: MUTED })], { after: 20 }),
    p(t("PL: Wpisz słowa po niemiecku, po kolei od 1 do 13. Numery idą od ogólnego do szczegółu: Ballade → rodzaje literackie → cechy. → = poziomo, ↓ = pionowo, Ä to jedna kratka.", { size: 17, it: true, color: PLC }), { after: 80 }),
    grid(withSol),
    p(t(""), { after: 60 }),
    ...H.map(([w, lvl, clue]) => p([
      t(`${NUM[w]} ${DIR[w] === "H" ? "→" : "↓"}  `, { bold: true, color: VIO, size: 18 }),
      t(clue, { size: 18, bold: lvl === 0 || lvl === 1 }),
      t("  –  " + PL[w], { size: 17, it: true, color: PLC }),
    ], { after: 30, indent: { left: lvl * 300 } })),
    p([t("2  Nächster Schritt – das machen wir gemeinsam: ", { bold: true, size: 19, color: NAVY }), t("Begründe mit Lösungswörtern zu Lyrik, Epik und Dramatik, warum „Paddy Fingal“ eine Ballade ist.", { size: 18 })], { before: 140, after: 10 }),
    p(t("PL: Ten krok robimy razem z klasą: uzasadniamy, dlaczego „Paddy Fingal“ to ballada – po jednym słowie z krzyżówki do Lyrik, Epik i Dramatik. Zapisuj razem z klasą.", { size: 17, it: true, color: PLC }), { after: 0 }),
    withSol
      ? p(t("Beispiel: „Paddy Fingal“ ist eine Ballade, weil sie Strophen mit Reimen hat, weil ein Erzähler die Geschichte erzählt und weil in Str. 7 Schaya ohne Erzähler spricht.", { size: 18, it: true, color: RED }), { before: 80, after: 0 })
      : p(t("„Paddy Fingal“ ist eine Ballade, weil ____________________________", { size: 19 }), { before: 170, after: 0 }),
    ...(withSol ? [] : [line(), line()]),
  ];
  return cell(kids, { w: SW, m: 120, ml: 160, borders: borders(B(BorderStyle.DASHED, 6, GRAY)) });
}
const doc = [table([SW], [new TableRow({ children: [schnipsel(false)] })])];
const props = { page: { size: { width: 11906, height: 16838 }, margin: { top: 450, bottom: 350, left: 500, right: 500 } } };
Packer.toBuffer(new Document({ styles: { default: { document: { run: { font: "Calibri", size: 19 } } } }, sections: [{ properties: props, children: doc }] }))
  .then((buf) => fs.writeFileSync(__dirname + "/../Handreichung_Polnisch_09-11_Kreuzwortraetsel.docx", buf));
