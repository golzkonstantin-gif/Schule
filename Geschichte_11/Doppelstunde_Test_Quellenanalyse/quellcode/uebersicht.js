// Gesamtübersicht für den Hefter: Interpretation einer Textquelle (nach dem Schema der Fachschaft), 1 Seite A4
const fs = require("fs");
const { Document, Packer, Paragraph, TextRun } = require("docx");
const { C, p, table, box } = require("./lib.js");
const OUT = process.argv[2];
const W = 10206;
const S = 20;

const rows = [
  ["1  Themenfrage", ["Welche Frage soll die Untersuchung beantworten?", "Sie steht in der Aufgabe – oder du stellst sie nach dem ersten Lesen selbst auf."],
    ["„Die Untersuchung geht der Frage nach, ob / wie …“"]],
  ["2a  Analyse: formale Aspekte", ["**Autor** (Amt, Stellung, soziale Schicht)", "**Zeit und Ort** der Entstehung bzw. Veröffentlichung", "**Textart** (Brief, Rede, Gesetz, Bericht …)", "**Thema** · **Adressat** (Privatperson, Öffentlichkeit, Nachwelt …)"],
    ["„Bei der Quelle handelt es sich um … von …, die/das … in … entstand.", "Sie richtet sich an … und behandelt …“"]],
  ["2b  Analyse: inhaltliche Aspekte", ["**Wesentliche Aussagen** nach Sinnabschnitten – eigene Worte, indirekte Rede, Zeilen", "**Schlüsselbegriffe**", "**Textsprache** (sachlich, emotional, appellativ, argumentativ …)"],
    ["„Zunächst berichtet der Autor, … habe … (Z. …).“", "„Anschließend erläutert / betont er, dass …“", "„Die Sprache ist …, was sich an … (Z. …) zeigt.“"]],
  ["3  Historischer Kontext", ["In welches Ereignis, welche Epoche, welchen Konflikt gehört die Quelle?", "Was geschah vorher, was danach?"],
    ["„Die Quelle bezieht sich auf …, nachdem …“", "„Sie entstand zur Zeit …“"]],
  ["4a  Sachurteil (aus damaliger Sicht)", ["Welchen Standpunkt nimmt der Autor ein? Welche **Absicht** verfolgt er?", "Welche **Wirkung** sollte der Text haben?", "Wie **glaubwürdig** ist er? Widersprüche? Sachlich angemessen?", "Vergleich mit anderen Quellen"],
    ["„Der Autor verfolgt die Absicht, …“", "„Die Quelle ist nur eingeschränkt glaubwürdig, weil …“", "„Es ist nicht auszuschließen, dass …“"]],
  ["4b  Werturteil (aus heutiger Sicht)", ["Wie bewerte ich die Quelle im Hinblick auf die Themenfrage?", "**Maßstab offenlegen** (z. B. Gleichheit, Leistung, Schutz der Schwachen, gleiche Regeln)"],
    ["„Gemessen an … halte ich … für …, weil …“", "„Dem ist allerdings entgegenzuhalten, dass …“"]],
];
const cellParas = (arr, o = {}) => arr.map(t => p(t, { size: S, after: 30, italics: o.it, color: o.it ? C.muted : undefined }));

const content = [
  p("GESCHICHTE · KLASSE 11 · FÜR DEN HEFTER", { size: 16, bold: true, color: C.terra, after: 30 }),
  new Paragraph({ children: [new TextRun({ text: "So interpretierst du eine Textquelle", font: "Cambria", bold: true, size: 34, color: C.dark })], spacing: { after: 60 } }),
  p("In der Klausur heißt die Aufgabe oft: „Interpretieren Sie die Quelle.“ Gehe immer in dieser Reihenfolge vor – erst untersuchen, dann urteilen.", { size: 19, after: 120 }),
  table(W, [2300, 4253, 3653], ["Schritt", "Leitfragen", "So kannst du formulieren"],
    rows.map(([s, f, b]) => [s, cellParas(f), cellParas(b, { it: true })]), { firstFill: true, size: S }),
  p("", { after: 100 }),
  box(W, [
    p("**Achtung – typische Fehler**", { size: 19, after: 50 }),
    p("–  **Meinung zu früh:** Deine Meinung gehört erst in Schritt 4 – vorher gibst du nur wieder, was der Autor sagt.", { size: S, after: 30 }),
    p("–  **Nacherzählen:** Nicht Satz für Satz, sondern das Wesentliche jedes Abschnitts, in indirekter Rede mit Zeilen.", { size: S, after: 30 }),
    p("–  **Augenzeuge?** Vergleiche die Lebensdaten des Autors mit dem Ereignis.", { size: S, after: 0 }),
  ]),
  p("", { after: 80 }),
  box(W, [
    p("**Operatoren:** analysieren (2a, 2b) · einordnen, erläutern (3) · **beurteilen** = Sachurteil (4a) · **bewerten**, Stellung nehmen = Werturteil (4b)", { size: S, after: 40 }),
    p("**Bildquelle:** gleiche Schritte. In 2a zusätzlich Auftraggeber, Zweck und Titel; in 2b statt der Textaussagen die Gestaltungsmittel **beschreiben** (Figuren, Gegenstände, Symbole, Farben, Perspektive …) und **deuten**: Was ist die Botschaft?", { size: S, after: 0 }),
  ], C.pale),
];

const doc = new Document({ styles: { default: { document: { run: { font: "Calibri", size: 20, color: C.dark } } } },
  sections: [{ properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 700, bottom: 600, left: 850, right: 850 } } }, children: content }] });
Packer.toBuffer(doc).then(b => { fs.writeFileSync(OUT, b); console.log("done"); });
