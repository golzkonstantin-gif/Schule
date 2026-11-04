// Lehrkraft: Quellentext (Aristoteles) mit Zeilennummern und Erwartungshorizont der Interpretation in Stichpunkten,
// gegliedert nach den Arbeitsschritten der Fachschaft (Themenfrage – Analyse – Kontext – Urteil). Tintensparend: keine Farbflächen.
const fs = require("fs");
const { Document, Packer, Paragraph, TextRun, TabStopType, PageBreak, BorderStyle } = require("docx");
const { C, p } = require("./lib.js");
const D = require("./data.js");
const OUT = process.argv[2];
const S = 20;
const kicker = t => p(t, { size: 16, bold: true, color: C.terra, after: 30 });
const title = t => new Paragraph({ children: [new TextRun({ text: t, font: "Cambria", bold: true, size: 32, color: C.dark })], spacing: { after: 80 } });
const zeile = (n, t) => new Paragraph({ tabStops: [{ type: TabStopType.RIGHT, position: 400 }, { type: TabStopType.LEFT, position: 640 }], spacing: { after: 0, line: 360 },
  children: [new TextRun({ text: "\t" + (n === 1 || n % 5 === 0 ? n : ""), size: 17, color: C.soft, font: "Calibri" }), new TextRun({ text: "\t" + t, size: 23, font: "Calibri", color: C.dark })] });
// Schritt-Überschrift mit farbigem Strich links (tintensparend)
const step = (nr, t, color, sub) => new Paragraph({ keepNext: true, spacing: { before: 200, after: 70 }, indent: { left: 120 },
  border: { left: { style: BorderStyle.SINGLE, size: 24, color, space: 6 } },
  children: [new TextRun({ text: nr + "  ", font: "Cambria", bold: true, size: 26, color }), new TextRun({ text: t, font: "Cambria", bold: true, size: 24, color: C.dark }),
    ...(sub ? [new TextRun({ text: "   " + sub, size: 17, italics: true, color: C.muted })] : [])] });
const li = (t, o = {}) => p("–  " + t, { size: S, after: 40, ...o });
const sub = t => p("     ·  " + t, { size: 19, after: 30, color: C.muted });

const quelle = [
  kicker("GESCHICHTE · KLASSE 11 · LEHRKRAFT"),
  title("Aristoteles über die Reformen des Kleisthenes"),
  p("Quellentext der Doppelstunde (identisch mit dem Schülerblatt) und Erwartungshorizont der Interpretation.", { size: 18, italics: true, color: C.muted, after: 140 }),
  ...D.zeilen.map((t, i) => zeile(i + 1, t)),
  p(D.nachweis, { size: 16, color: C.muted, before: 120, after: 60 }),
  p("**Worterklärungen:** Alkmeoniden = mächtige Adelsfamilie · Isagoras = adliger Gegenspieler des Kleisthenes · Bünde der Adligen = Zusammenschlüsse adliger Freunde · Demos, Trittys, Phyle: siehe Montag", { size: 18, after: 160 }),
  p("**Zum Autor und zur Schrift**", { size: 19, after: 40 }),
  ...D.info.map(t => p(t, { size: 19, after: 40 })),
];

const T = Object.fromEntries(D.einleitung.felder);
const erwartung = [
  new Paragraph({ children: [new PageBreak()] }),
  kicker("ERWARTUNGSHORIZONT · STICHPUNKTE"),
  title("So sollte die Interpretation aussehen"),
  p("Mindestziel der Stunde: Schritte 1, 2a und 2b. Schritte 3 und 4 nur bei Zeit bzw. in der nächsten Doppelstunde. Andere Formulierungen sind richtig, wenn der Kern stimmt.", { size: 18, italics: true, color: C.muted, after: 60 }),

  step("1", "Themenfrage", C.dark, "heute vorgegeben"),
  li(D.themenfrage),

  step("2a", "Analyse: formal (TATTEA)", C.terra),
  li("**T – Titel:** " + T.Titel),
  li("**A – Autor:** " + T.Autor),
  li("**T – Textsorte:** " + T.Textsorte),
  li("**T – Thema:** " + T.Thema),
  li("**E – Entstehung:** " + T.Entstehung),
  li("**A – Adressat:** " + T.Adressat),
  li("**Einleitungssatz:** „" + D.einleitung.satz + "“", { color: C.muted }),
  li("**Achtung:** kein Augenzeuge – rund 180 Jahre Abstand.", { color: C.terra }),

  step("2b", "Analyse: inhaltlich", C.terra, "eigene Worte · indirekte Rede · Zeilen · keine Meinung"),
  li("**Z. 1–4 – Machtkampf:** Nach dem Sturz der Tyrannen kämpfen Isagoras und Kleisthenes um die Macht; Kleisthenes unterliegt bei den Adligen und gewinnt das Volk, indem er ihm Macht im Staat überträgt."),
  li("**Z. 5–8 – Phylen und Rat:** zehn statt vier Phylen, um die Bürger zu „vermischen“ und mehr Menschen zu beteiligen (Z. 6f.); Rat der 500 statt 400, 50 je Phyle (Z. 7f.)."),
  li("**Z. 9–12 – Trittyen:** Land nach Demen in 30 Teile (je 10 Stadt, Küste, Binnenland); je drei Trittyen per Los einer Phyle zugeteilt, damit jede Phyle Anteil an allen Gebieten hat (Z. 11f.)."),
  li("**Z. 12–14 – Demenname:** Benennung nach dem Demos statt nach dem Vater, damit Neubürger nicht erkennbar sind."),
  li("**Schlüsselbegriffe:** " + D.begriffe),
  li("**Textsprache:** " + D.sprache),

  step("3", "Historischer Kontext", C.blue, "nur bei Zeit"),
  li("Vorher Tyrannis der Peisistratiden; 510 Vertreibung des Hippias, danach Machtkampf der Adelsfamilien."),
  li("Isagoras holt Sparta zu Hilfe, das Volk vertreibt die Spartaner; Kleisthenes setzt 508/07 seine Reform durch."),
  li("Ziel der Reform (Montag): die Macht der Adelsfamilien über ihre Gegend brechen – kein Adliger hat mehr einen geschlossenen Block hinter sich."),
  li("Einordnung in die Reihe: Solon (Rechte nach Besitz) → Peisistratos → **Kleisthenes (Mischung der Bürger)** → Perikles."),

  step("4a", "Sachurteil", C.green, "aus damaliger Sicht"),
  li("**Absicht des Autors:** erklärt die Entstehung der Verfassung; deutet die Reform als gezielte Beteiligung des Volkes („damit …“, Z. 7, 11, 13)."),
  li("**Glaubwürdigkeit:**"),
  sub("dagegen: kein Augenzeuge, rund 180 Jahre später; stützt sich auf ältere Schriften; Absichten des Kleisthenes sind Deutung des Autors"),
  sub("dafür: sachliche Sprache; Angaben zu Phylen, Rat und Trittyen passen zu anderen Überlieferungen (Herodot)"),
  li("**Ergebnis:** als Quelle für die neue Ordnung gut brauchbar, für die Absichten des Kleisthenes nur eingeschränkt."),

  step("4b", "Werturteil", C.green, "aus heutiger Sicht, Maßstab offenlegen"),
  li("**Gleichheit:** großer Schritt, weil viel mehr Bürger beteiligt wurden und Herkunft weniger zählte – aber unvollständig: Frauen, Sklaven und Fremde blieben ausgeschlossen."),
  li("**Gleiche Regeln:** positiv – Bürgerrecht über die Dorfliste, nicht über die Familie."),
  li("**Bezug zur Themenfrage:** Kleisthenes veränderte die Ordnung durch Mischung der Bürger, um das Volk zu beteiligen und die Macht der Adelsfamilien zu brechen."),
];

const doc = new Document({ styles: { default: { document: { run: { font: "Calibri", size: S, color: C.dark } } } },
  sections: [{ properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 680, bottom: 600, left: 992, right: 992 } } }, children: [...quelle, ...erwartung] }] });
Packer.toBuffer(doc).then(b => { fs.writeFileSync(OUT, b); console.log("done"); });
