// Arbeitsblatt: Quelle (S. 1) und Schreibplan (S. 2), A4 doppelseitig
const fs = require("fs");
const { Document, Packer, Paragraph, TextRun, AlignmentType, TabStopType, PageBreak, Footer } = require("docx");
const { C, p, h, table, box, lines } = require("./lib.js");
const D = require("./data.js");
const OUT = process.argv[2];
const W = 9922;
const kicker = t => p(t, { size: 16, bold: true, color: C.terra, after: 30 });
const title = t => new Paragraph({ children: [new TextRun({ text: t, font: "Cambria", bold: true, size: 32, color: C.dark })], spacing: { after: 80 } });
const zeile = (n, t) => new Paragraph({ tabStops: [{ type: TabStopType.RIGHT, position: 400 }, { type: TabStopType.LEFT, position: 640 }], spacing: { after: 0, line: 360 },
  children: [new TextRun({ text: "\t" + (n === 1 || n % 5 === 0 ? n : ""), size: 17, color: C.soft, font: "Calibri" }), new TextRun({ text: "\t" + t, size: 23, font: "Calibri", color: C.dark })] });

const s1 = [
  kicker("GESCHICHTE · KLASSE 11 · QUELLENANALYSE"),
  title("Aristoteles über die Reformen des Kleisthenes"),
  p("Lies die Quelle. Die Zeilennummern brauchst du für den Schreibplan auf der Rückseite.", { size: 19, italics: true, color: C.muted, after: 140 }),
  ...D.zeilen.map((t, i) => zeile(i + 1, t)),
  p(D.nachweis, { size: 16, color: C.muted, before: 120, after: 60 }),
  p("**Worterklärungen:** Alkmeoniden = mächtige Adelsfamilie · Isagoras = adliger Gegenspieler des Kleisthenes · Bünde der Adligen = Zusammenschlüsse adliger Freunde · Demos, Trittys, Phyle: siehe Montag", { size: 18, after: 200 }),
  box(W, [p("**Zum Autor und zur Schrift**", { size: 20, after: 60 }), ...D.info.map(t => p(t, { size: 20, after: 60 }))]),
  p("", { after: 120 }),
  box(W, [
    p("**Wörter, die beim Schreiben helfen**", { size: 20, after: 60 }),
    p("Aristoteles **berichtet / erklärt / betont / beschreibt**, dass …", { size: 20, after: 40 }),
    p("Indirekte Rede: Kleisthenes **habe** … eingeteilt · die Bürger **seien** … · er **wolle** …", { size: 20, after: 40 }),
    p("Zeilen angeben: (Z. 6) · (Z. 7f.) · (Z. 9–12)", { size: 20, after: 0 }),
  ], C.pale),
];

const step = (nr, t, sub) => p([new TextRun({ text: nr + "  ", font: "Cambria", bold: true, size: 25, color: C.terra }), new TextRun({ text: t, font: "Cambria", bold: true, size: 24, color: C.dark }),
  new TextRun({ text: sub ? "   " + sub : "", size: 18, color: C.muted, italics: true })], { before: 160, after: 70, keepNext: true });

const s2 = [
  new Paragraph({ children: [new PageBreak()] }),
  kicker("SCHREIBPLAN · INTERPRETATION EINER TEXTQUELLE"),
  title("Schritt für Schritt"),
  step("1", "Themenfrage", "Welche Frage bestimmt die Untersuchung? (heute vorgegeben)"),
  box(W, [p("**" + D.themenfrage + "**", { size: 20, after: 0 })]),

  step("2a", "Analyse: formale Aspekte", "TATTE(A): Titel – Autor – Textsorte – Thema – Entstehung – Adressat"),
  table(W, [2700, 7222], null, [["T  Titel", ""], ["A  Autor", ""], ["T  Textsorte", ""], ["T  Thema", ""], ["E  Entstehungszeit und -ort", ""], ["(A)  Adressat", ""]], { firstFill: true, h: 340, size: 19 }),
  p("**Satzbaustein:** Bei der Quelle handelt es sich um [Textsorte] „…“ [Titel] von … [Autor], die/das … in … [Entstehung] entstand. Sie richtet sich an … [Adressat] und behandelt … [Thema].", { size: 18, before: 60, after: 40 }),
  lines(W, 2, 360),

  step("2b", "Analyse: inhaltliche Aspekte", "eigene Worte · indirekte Rede · mit Zeilen · noch keine Meinung"),
  table(W, [1150, 8772], ["Zeilen", "Wesentliche Aussagen"], D.abschnitte.map(([z, , vorgabe], i) => [z, i === 0 ? "✓ Beispiel: " + vorgabe : [p(vorgabe || " ", { size: 19, after: 20, italics: !!vorgabe, color: C.muted }), p(" ", { size: 19, after: 20 }), p(" ", { size: 19, after: 20 })]]),
    { firstFill: true, size: 19, h: 1150 }),
  p("**Schlüsselbegriffe:** ____________________________________________________________________________", { size: 19, before: 100, after: 60 }),
  p("**Textsprache** (ankreuzen): " + D.sprachen.map(s => "☐ " + s).join("   ") + "   – Beleg: Z. ____", { size: 19, after: 40 }),

  step("3", "Historischer Kontext*", "In welchen Zusammenhang gehört die Quelle? Was geschah vorher?"),
  lines(W, 2, 360),
  step("4", "Urteilen*", "Sachurteil: Absicht? Glaubwürdig? (Wann schrieb er, woher wusste er es?) · Werturteil: aus heutiger Sicht"),
  lines(W, 3, 360),
];

const doc = new Document({ styles: { default: { document: { run: { font: "Calibri", size: 20, color: C.dark } } } },
  sections: [{ properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 680, bottom: 600, left: 992, right: 992 } } },
    footers: { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "* nur wenn die Zeit reicht", size: 15, color: C.soft })] })] }) },
    children: [...s1, ...s2] }] });
Packer.toBuffer(doc).then(b => { fs.writeFileSync(OUT, b); console.log("done"); });
