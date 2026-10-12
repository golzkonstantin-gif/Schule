const fs = require("fs");
const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, BorderStyle, ShadingType, AlignmentType, VerticalAlign, PageBreak, Footer, PageNumber } = require("docx");

const NAVY = "1E2761", MUTED = "5B6B8C", LIGHT = "F4F7FD", LIGHT2 = "EAF0FA", RED = "D9534F";
const W = 10206;
const t = (text, o = {}) => new TextRun({ text, font: o.font || "Calibri", size: o.size || 26, bold: o.bold, italics: o.italics, color: o.color || "000000" });
const p = (runs, o = {}) => new Paragraph({ children: Array.isArray(runs) ? runs : [runs], alignment: o.align, spacing: { before: o.before ?? 0, after: o.after ?? 120, line: o.line }, keepNext: o.keepNext });
const solid = (c, s = 8) => ({ style: BorderStyle.SINGLE, size: s, color: c });
const all = (b) => ({ top: b, bottom: b, left: b, right: b });
function box(children, fill, border) {
  return new Table({ width: { size: W, type: WidthType.DXA }, columnWidths: [W], rows: [new TableRow({ cantSplit: true, children: [new TableCell({ children, width: { size: W, type: WidthType.DXA }, shading: { fill, type: ShadingType.CLEAR, color: "auto" }, borders: border || all({ style: BorderStyle.NONE, size: 0, color: "FFFFFF" }), margins: { top: 140, bottom: 140, left: 200, right: 200 }, verticalAlign: VerticalAlign.TOP })] })] });
}
const footer = new Footer({ children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ font: "Calibri", size: 16, color: MUTED, children: ["Geschichte 8 · Französische Revolution · Expertentexte · Seite ", PageNumber.CURRENT] })] })] });

const TEXTE = [
  {
    gruppe: "A",
    thema: "Von der leeren Kasse zur Nationalversammlung",
    zeit: "Mai bis Juni 1789",
    absaetze: [
      "1789 ist Frankreich fast zahlungsunfähig. Hof, Heer und teure Kriege haben hohe Schulden gemacht. König Ludwig XVI. braucht neue Steuern. Doch Adel und Klerus weigern sich, selbst Steuern zu zahlen. Deshalb ruft der König am 5. Mai 1789 die Generalstände nach Versailles zusammen. Das ist eine Versammlung von Vertretern aller drei Stände. Zuletzt war sie 1614 zusammengetreten, vor 175 Jahren.",
      "Es kommen etwa 300 Vertreter des Klerus, 300 des Adels und 600 des Dritten Standes (Bauern und Bürger). Doch beim Abstimmen soll jeder Stand nur eine Stimme haben. Klerus und Adel stimmen meist gleich, der Dritte Stand verliert immer. Seine Vertreter verlangen deshalb: Wir stimmen nach Köpfen ab.",
      "Der König lehnt das ab. Am 17. Juni 1789 erklären sich die Vertreter des Dritten Standes zur Nationalversammlung. Sie wollen das ganze Volk vertreten. Der König lässt den Sitzungssaal sperren. Am 20. Juni treffen sich die Abgeordneten in einem Ballhaus und schwören: Wir gehen nicht auseinander, bis Frankreich eine Verfassung hat.",
    ],
    woerter: [["Generalstände", "Versammlung von Vertretern der drei Stände"], ["Nationalversammlung", "Versammlung, die das ganze Volk vertreten will"], ["Verfassung", "Gesetze, an die auch der König gebunden ist"]],
  },
  {
    gruppe: "B",
    thema: "Der Sturm auf die Bastille",
    zeit: "14. Juli 1789",
    absaetze: [
      "Im Sommer 1789 hungern viele Menschen in Paris. Die Ernte war schlecht, ein Brot kostet so viel wie nie. Wer wenig verdient, gibt den größten Teil seines Lohns für Brot aus. Gleichzeitig zieht der König viele Soldaten rund um Paris zusammen. Gerüchte gehen um: Er wolle die Nationalversammlung mit Gewalt auflösen.",
      "Am 14. Juli 1789 ziehen tausende Pariser durch die Stadt. Sie suchen Waffen und Pulver, um sich zu verteidigen. Ihr Ziel ist die Bastille, eine alte Festung, die auch als Gefängnis dient. Sie gilt als Zeichen der Willkür des Königs, denn hier können Menschen ohne Gerichtsverfahren eingesperrt werden. Nach mehreren Stunden Kampf geben die Wachen auf, die Menge dringt ein. Gefunden werden nur sieben Gefangene.",
      "Die Nachricht verbreitet sich schnell in ganz Frankreich. Der König hat sich nicht gegen das Volk durchgesetzt. Heute ist der 14. Juli der Nationalfeiertag Frankreichs.",
    ],
    woerter: [["Bastille", "Festung und Gefängnis in Paris"], ["Willkür", "Der Herrscher entscheidet nach Belieben, ohne Gesetz"], ["Nationalfeiertag", "wichtigster Feiertag eines Landes"]],
  },
  {
    gruppe: "C",
    thema: "Das Ende der Vorrechte und die Menschenrechte",
    zeit: "August 1789",
    absaetze: [
      "Nach dem Sturm auf die Bastille gibt es auch auf dem Land Unruhen. Bauern greifen Schlösser an und verbrennen Urkunden, in denen ihre Abgaben stehen. Die Nationalversammlung merkt: So kann es nicht weitergehen.",
      "In der Nacht vom 4. auf den 5. August 1789 beschließt sie, viele Vorrechte abzuschaffen. Adel und Klerus müssen nun auch Steuern zahlen, das alleinige Jagdrecht des Adels fällt, und auch der Kirchenzehnt wird abgeschafft. Alle Menschen sollen vor dem Gesetz gleich sein.",
      "Am 26. August 1789 folgt die Erklärung der Menschen- und Bürgerrechte. Dort steht: Alle Menschen werden frei und gleich an Rechten geboren. Jeder hat Rechte, zum Beispiel auf Freiheit und Sicherheit, und niemand darf ohne Grund verhaftet werden. Damit ist die Ständegesellschaft zumindest auf dem Papier zu Ende.",
    ],
    woerter: [["Vorrechte", "Rechte, die nur Adel und Klerus hatten"], ["Menschenrechte", "Rechte, die jedem Menschen zustehen"], ["Urkunde", "schriftliches Dokument"]],
  },
];

const c = [];
TEXTE.forEach((x, i) => {
  if (i) c.push(new Paragraph({ children: [new PageBreak()] }));
  c.push(p(t(`EXPERTENGRUPPE ${x.gruppe} · GESCHICHTE · KLASSE 8`, { size: 18, bold: true, color: MUTED }), { after: 20 }));
  c.push(p(t(x.thema, { font: "Cambria", size: 38, bold: true, color: NAVY }), { after: 40 }));
  c.push(p(t(x.zeit, { size: 24, color: RED, bold: true }), { after: 160 }));
  x.absaetze.forEach((a) => c.push(p(t(a, { font: "Cambria", size: 26 }), { after: 140, line: 330 })));
  c.push(p(t(""), { after: 60 }));
  c.push(box([p(t("Wörter erklärt", { bold: true, color: NAVY, size: 22 }), { after: 60 }), ...x.woerter.map(([w, e]) => p([t(w + ": ", { bold: true, size: 22 }), t(e, { size: 22 })], { after: 30 }))], LIGHT2));
  c.push(p(t(""), { after: 60 }));
  c.push(box([p(t("Deine Aufgabe", { bold: true, color: RED, size: 22 }), { after: 60 }), p(t("1  Lies den Text still und unterstreiche die wichtigsten Stichworte.", { size: 22 }), { after: 30 }), p(t("2  Ergänze das Raster im Heft: Wann? Was geschah? Welches Problem steckt dahinter?", { size: 22 }), { after: 30 }), p(t("3  Erkläre dein Thema in der Stammgruppe in eigenen Worten (3 Minuten).", { size: 22 }), { after: 0 })], LIGHT, all(solid("B9C6E8", 6))));
});

(async () => {
  const doc = new Document({ styles: { default: { document: { run: { font: "Calibri", size: 26 } } } }, sections: [{ properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 900, bottom: 800, left: 850, right: 850 } } }, footers: { default: footer }, children: c }] });
  fs.writeFileSync("../Expertentexte_Franzoesische_Revolution.docx", await Packer.toBuffer(doc));
  console.log("ok");
})();
