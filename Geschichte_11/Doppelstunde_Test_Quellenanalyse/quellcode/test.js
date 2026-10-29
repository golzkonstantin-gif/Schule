// Test Geschichte 11: Athen um 600 bis Peisistratos (45 Min., 37 Punkte + 1 Zusatzpunkt)
const fs = require("fs");
const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, BorderStyle, ShadingType, AlignmentType,
  TabStopType, ImageRun, HeightRule, PageBreak, Footer, PageNumber } = require("docx");
const OUT = process.argv[2];
const C = { dark: "2B2318", terra: "AD4426", sand: "F4ECDE", muted: "5A5044", line: "BFB5A5" };
const W = 10206;
const R = (t, o = {}) => new TextRun({ text: t, font: o.font || "Calibri", size: o.size || 21, bold: o.bold, italics: o.italics, color: o.color || C.dark });
const P = (runs, o = {}) => new Paragraph({ alignment: o.align, spacing: { after: o.after ?? 40, before: o.before ?? 0 }, keepNext: o.keepNext, tabStops: o.tabs, children: Array.isArray(runs) ? runs : [R(runs, o)] });
const thin = { style: BorderStyle.SINGLE, size: 4, color: C.line }, none = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
const lines = n => new Table({ columnWidths: [W], width: { size: W, type: WidthType.DXA }, rows: Array.from({ length: n }, () => new TableRow({ height: { value: 400, rule: HeightRule.EXACT },
  children: [new TableCell({ width: { size: W, type: WidthType.DXA }, borders: { top: none, left: none, right: none, bottom: thin }, children: [new Paragraph({ children: [] })] })] })) });
const task = (nr, title, pts) => P([R(nr + "  ", { bold: true, color: C.terra, size: 23, font: "Cambria" }), R(title, { bold: true, size: 22, font: "Cambria" }), R("\t" + pts + " P.", { size: 19, color: C.muted, bold: true })],
  { tabs: [{ type: TabStopType.RIGHT, position: W }], before: 160, after: 60, keepNext: true });
function grid(cw, rows, o = {}) {
  return new Table({ columnWidths: cw, width: { size: W, type: WidthType.DXA }, rows: rows.map((r, ri) => new TableRow({ cantSplit: true, height: o.h && !(ri === 0 && o.head) ? { value: o.h, rule: HeightRule.ATLEAST } : undefined,
    children: r.map((c, i) => new TableCell({ width: { size: cw[i], type: WidthType.DXA }, borders: { top: thin, bottom: thin, left: thin, right: thin },
      shading: ri === 0 && o.head ? { type: ShadingType.CLEAR, color: "auto", fill: C.sand } : undefined, margins: { top: 50, bottom: 50, left: 90, right: 90 },
      children: [P([R(c, { size: 19, bold: ri === 0 && o.head })], { after: 0, align: i > 0 && o.center ? AlignmentType.CENTER : undefined })] })) })) });
}

const children = [
  P([R("GESCHICHTE · KLASSE 11 · TEST", { bold: true, size: 17, color: C.terra }), R("\tName: ______________________   Datum: __________", { size: 19, color: C.muted })], { tabs: [{ type: TabStopType.RIGHT, position: W }], after: 60 }),
  P([R("Athen um 600 v. Chr. – Drakon, Solon, Peisistratos", { font: "Cambria", bold: true, size: 30 })], { after: 40 }),
  P([R("Bearbeitungszeit 45 Minuten · 37 Punkte + 1 Zusatzpunkt · Hilfsmittel: keine", { size: 18, italics: true, color: C.muted })], { after: 60 }),

  task("1", "Die Krise Athens um 600 v. Chr.", 6),
  P("a) Nenne vier Probleme, die Athen um 600 v. Chr. in eine Krise stürzten. (4 P.)", { size: 20 }),
  lines(2),
  P("b) Erkläre, wie ein Kleinbauer in die Schuldknechtschaft geraten konnte. (2 P.)", { size: 20, before: 80 }),
  lines(2),

  task("2", "Drakon, Solon, Peisistratos: richtig oder falsch?", 9),
  P("Kreuze an. Berichtige jede falsche Aussage in einem Satz in der rechten Spalte. (je 1 P. fürs Kreuz, je 1 P. für die Berichtigung)", { size: 20 }),
  grid([4300, 750, 750, 4406], [
    ["Aussage", "richtig", "falsch", "Berichtigung (nur bei falschen Aussagen)"],
    ["a) Drakon schrieb das Recht in Athen erstmals auf.", "☐", "☐", ""],
    ["b) Drakon schaffte die Schuldknechtschaft ab.", "☐", "☐", ""],
    ["c) Solon strich die Schulden und kaufte versklavte Bürger zurück.", "☐", "☐", ""],
    ["d) Solon verteilte das Land der Großgrundbesitzer an die Bauern.", "☐", "☐", ""],
    ["e) Peisistratos wurde von der Volksversammlung zum Herrscher gewählt.", "☐", "☐", ""],
    ["f) Peisistratos ließ Solons Verfassung bestehen, besetzte die Ämter aber mit seinen Anhängern.", "☐", "☐", ""],
  ], { head: true, center: true, h: 840 }),

  task("3", "Sachurteil oder Werturteil?", 6),
  P("a) Ordne zu: S = Sachurteil, W = Werturteil. (je 1 P.)", { size: 20 }),
  grid([8906, 1300], [
    ["Aussage", "S / W"],
    ["1) Gemessen am Schutz der Schwachen halte ich Drakons Gesetze für ungerecht, weil schon kleine Vergehen sehr hart bestraft wurden.", ""],
    ["2) Aus Sicht der damaligen Zeit waren Drakons Gesetze ein Fortschritt, weil nun jeder wissen konnte, welche Strafe drohte.", ""],
    ["3) Aus Sicht der damaligen Zeit schwächte Peisistratos den Adel, weil er die Ämter mit seinen eigenen Anhängern besetzte.", ""],
    ["4) Die Tyrannis des Peisistratos war ein Fehler.", ""],
  ], { head: true, center: true }),
  P("b) Aussage 4 ist ein schwaches Werturteil. Verbessere es: Nenne einen Maßstab und begründe. (2 P.)", { size: 20, before: 80 }),
  lines(2),

  task("4", "Solons Verfassung (594 v. Chr.)", 10),
  P("Die vier Klassen von links: Fünfhundertscheffler, Ritter, Hopliten, Theten. Die Zeilen in den Klassen-Kästen musst du nicht ausfüllen.", { size: 18, italics: true, color: C.muted, keepNext: true }),
  P("a) Trage die vier Ämter (oben) und die Volksversammlung (unten) ein. (je 1 P.)", { size: 20, keepNext: true }),
  P("b) Kreuze bei jeder Klasse an, was sie darf. (je vollständig richtige Klasse 1 P.)", { size: 20, keepNext: true }),
  P("c) Zeichne Pfeile ein: Wer kann in welches Amt gewählt werden? (1 P.)", { size: 20, keepNext: true }),
  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 40 }, children: [new ImageRun({ type: "png", data: fs.readFileSync(__dirname + "/stumm.png"), transformation: { width: 470, height: 317 } })] }),

  task("5", "Selbst urteilen", 6),
  P("a) Sachurteil: Inwieweit gelang es Solon, die Krise Athens zu lösen? Beziehe Stellung und nenne zwei Belege. (3 P.)", { size: 20 }),
  lines(5),
  P("b) Werturteil: Ist es gerecht, dass die politischen Rechte vom Besitz abhingen? Nenne deinen Maßstab, beziehe Stellung und begründe mit einem Beispiel. (3 P.)", { size: 20, before: 80 }),
  lines(5),

  task("Zusatz", "Erkläre den Begriff „Timokratie“ in einem Satz.", "+1"),
  lines(1),
];
const doc = new Document({ sections: [{ properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 560, bottom: 500, left: 850, right: 850 } } },
  footers: { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ children: ["Seite ", PageNumber.CURRENT, " von 2"], size: 16, color: "998C77" })] })] }) },
  children }] });
Packer.toBuffer(doc).then(b => { fs.writeFileSync(OUT, b); console.log("done"); });
