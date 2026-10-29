// Gesamtübersicht für den Hefter: Interpretation einer Textquelle (nach dem Schema der Fachschaft), 1 Seite A4
// Formale Analyse mit der TATTE(A)-Formel, die die Klasse schon kennt.
const fs = require("fs");
const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, ShadingType, AlignmentType, VerticalAlign } = require("docx");
const { C, p, box, none } = require("./lib.js");
const OUT = process.argv[2];
const W = 10206, L = 2050, S = 20;
const nob = { top: none, bottom: none, left: none, right: none };
const white = { top: { style: "single", size: 18, color: "FFFFFF" }, bottom: { style: "single", size: 18, color: "FFFFFF" }, left: { style: "single", size: 18, color: "FFFFFF" }, right: { style: "single", size: 18, color: "FFFFFF" } };
const cell = (children, w, fill, o = {}) => new TableCell({ children, width: { size: w, type: WidthType.DXA }, borders: o.borders || nob, verticalAlign: o.v || VerticalAlign.TOP,
  shading: fill ? { type: ShadingType.CLEAR, color: "auto", fill } : undefined, margins: { top: o.m ?? 80, bottom: o.m ?? 80, left: 120, right: 120 }, columnSpan: o.span });
const run = (t, o = {}) => new TextRun({ text: t, font: o.font || "Calibri", size: o.size || S, bold: o.bold, italics: o.it, color: o.color || C.dark });
const para = (runs, o = {}) => new Paragraph({ children: runs, alignment: o.align, spacing: { after: o.after ?? 30, before: o.before ?? 0 } });
const label = (nr, name, sub, fill) => cell([
  para([run(nr, { font: "Cambria", bold: true, size: 40, color: C.white })], { after: 0 }),
  para([run(name, { font: "Cambria", bold: true, size: 21, color: C.white })], { after: 20 }),
  ...(sub ? [para([run(sub, { size: 15, color: "F4ECDE", it: true })], { after: 0 })] : []),
], L, fill, { v: VerticalAlign.CENTER });
const qs = arr => arr.map(t => p("▸ " + t, { size: S, after: 25 }));
const say = t => p("✎ " + t, { size: S, italics: true, color: C.muted, after: 0, before: 30 });
const row = (lab, right, w2 = W - L) => new Table({ columnWidths: [L, w2], width: { size: W, type: WidthType.DXA }, rows: [new TableRow({ cantSplit: true, children: [lab, cell(right, w2, C.pale)] })] });
const arrow = () => para([run("▼", { size: 16, color: C.gold })], { align: AlignmentType.CENTER, after: 0 });

// TATTE(A)-Leiste
const TA = [["T", "Titel"], ["A", "Autor"], ["T", "Textsorte"], ["T", "Thema"], ["E", "Entstehungs-zeit und -ort"], ["(A)", "Adressat"]];
const tw = Math.floor((W - L - 240) / 6);
const tatte = new Table({ columnWidths: TA.map(() => tw), width: { size: tw * 6, type: WidthType.DXA }, rows: [new TableRow({ children: TA.map(([b, w], i) => cell([
  para([run(b, { font: "Cambria", bold: true, size: 40, color: i === 5 ? C.muted : C.terra })], { align: AlignmentType.CENTER, after: 0 }),
  para([run(w, { size: 17, bold: true })], { align: AlignmentType.CENTER, after: 0 }),
], tw, i === 5 ? "EDE4D3" : C.sand, { borders: white, m: 40, v: VerticalAlign.CENTER })) })] });
const tag = (t, k) => [run(t, { size: S, it: true }), run(" [" + k + "]", { size: 14, bold: true, color: C.terra })];

const content = [
  p("GESCHICHTE · KLASSE 11 · FÜR DEN HEFTER", { size: 16, bold: true, color: C.terra, after: 20 }),
  new Paragraph({ children: [run("So interpretierst du eine Textquelle", { font: "Cambria", bold: true, size: 34 })], spacing: { after: 40 } }),
  p("Klausur: „Interpretieren Sie die Quelle.“ Gehe immer von oben nach unten vor – **erst untersuchen, dann urteilen.**", { size: 19, after: 100 }),

  row(label("1", "Themenfrage", null, C.dark), [...qs(["Welche Frage soll die Untersuchung beantworten? Sie steht in der Aufgabe – oder du stellst sie nach dem ersten Lesen selbst auf."]), say("„Die Untersuchung geht der Frage nach, ob / wie …“")]),
  arrow(),
  row(label("2a", "Analyse: formal", "Wer? Was? Wann?", C.terra), [
    p("Die **TATTE(A)-Formel** – damit beginnt jede Interpretation:", { size: S, after: 50 }),
    tatte,
    p("Beispiel:", { size: 15, bold: true, color: C.muted, before: 70, after: 10 }),
    para([run("„Bei der Quelle handelt es sich um ", { size: S, it: true }), ...tag("ein Gedicht", "Textsorte"), run(" ohne Titel ", { size: S, it: true }), run("[T]", { size: 14, bold: true, color: C.terra }),
      run(" des Gesetzgebers ", { size: S, it: true }), ...tag("Solon", "A"), run(", das ", { size: S, it: true }), ...tag("um 590 v. Chr. in Athen", "E"), run(" entstand. Er wendet sich ", { size: S, it: true }),
      ...tag("an die Athener", "(A)"), run(" und ", { size: S, it: true }), ...tag("rechtfertigt seine Reformen", "Thema"), run(".“", { size: S, it: true })], { after: 20 }),
    p("Kein Titel? Dann die Überschrift im Buch oder Material nennen.", { size: 15, color: C.muted, after: 0 }),
  ]),
  arrow(),
  row(label("2b", "Analyse: inhaltlich", "Was steht drin?", C.terra), [...qs(["**Wesentliche Aussagen** je Sinnabschnitt – eigene Worte, indirekte Rede, Zeilen", "**Schlüsselbegriffe** · **Textsprache** (sachlich, emotional, appellativ, argumentativ …)"]),
    say("„Zunächst berichtet der Autor, … habe … (Z. …). Anschließend betont er, dass …“")]),
  arrow(),
  row(label("3", "Historischer Kontext", "Was geschah damals?", C.blue), [...qs(["In welches Ereignis, welche Epoche, welchen Konflikt gehört die Quelle? Was geschah vorher, was danach?"]),
    say("„Die Quelle bezieht sich auf …, nachdem …“")]),
  arrow(),
  new Table({ columnWidths: [L, (W - L) / 2, (W - L) / 2], width: { size: W, type: WidthType.DXA }, rows: [new TableRow({ cantSplit: true, children: [
    label("4", "Urteilen", "erst Sach-, dann Werturteil", C.green),
    cell([p("**a) Sachurteil** – aus damaliger Sicht", { size: S, color: C.green, after: 40 }), ...qs(["Standpunkt und **Absicht** des Autors?", "**Wirkung** auf die Adressaten?", "**Glaubwürdig?** Widersprüche?"]), say("„Der Autor verfolgt die Absicht, …“")], (W - L) / 2, C.pale, { borders: { ...nob, right: { style: "single", size: 12, color: "FFFFFF" } } }),
    cell([p("**b) Werturteil** – aus heutiger Sicht", { size: S, color: C.green, after: 40 }), ...qs(["Wie bewerte ich die Quelle im Hinblick auf die Themenfrage?", "**Maßstab offenlegen:** Gleichheit, Leistung, Schutz der Schwachen, gleiche Regeln"]), say("„Gemessen an … halte ich … für …, weil …“")], (W - L) / 2, C.pale),
  ] })] }),

  p("", { after: 80 }),
  new Table({ columnWidths: [W / 2, W / 2], width: { size: W, type: WidthType.DXA }, rows: [new TableRow({ children: [
    cell([p("**⚠ Typische Fehler**", { size: S, color: C.gold, after: 40 }),
      p("Meinung zu früh – sie gehört erst in Schritt 4.", { size: 19, color: C.white, after: 25 }),
      p("Nacherzählen statt Wesentliches mit Zeilen.", { size: 19, color: C.white, after: 25 }),
      p("„Augenzeuge“? Lebensdaten mit dem Ereignis vergleichen.", { size: 19, color: C.white, after: 0 })], W / 2, C.dark, { borders: { ...nob, right: { style: "single", size: 24, color: "FFFFFF" } } }),
    cell([p("**Operatoren:** analysieren (2) · einordnen, erläutern (3) · **beurteilen** = Sachurteil · **bewerten**, Stellung nehmen = Werturteil", { size: 19, after: 50 }),
      p("**Bildquelle:** gleiche Schritte; in 2b die Gestaltungsmittel **beschreiben** (Figuren, Symbole, Farben, Perspektive …) und **deuten**: Was ist die Botschaft?", { size: 19, after: 0 })], W / 2, C.sand),
  ] })] }),
];

const doc = new Document({ styles: { default: { document: { run: { font: "Calibri", size: 20, color: C.dark } } } },
  sections: [{ properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 640, bottom: 500, left: 850, right: 850 } } }, children: content }] });
Packer.toBuffer(doc).then(b => { fs.writeFileSync(OUT, b); console.log("done"); });
