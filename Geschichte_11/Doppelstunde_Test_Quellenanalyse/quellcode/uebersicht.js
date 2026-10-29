// Übersicht für den Hefter (A4 doppelseitig): Interpretation einer Textquelle nach dem Schema der Fachschaft.
// Vorderseite: Arbeitsschritte als Tabelle (formale Analyse mit TATTEA). Rückseite: Formulierungshilfen je Schritt.
const fs = require("fs");
const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, ShadingType, VerticalAlign, PageBreak, BorderStyle } = require("docx");
const { C } = require("./lib.js");
const OUT = process.argv[2];
const W = 10206, S = 24;
const line = { style: BorderStyle.SINGLE, size: 6, color: "BFB5A5" };
const borders = { top: line, bottom: line, left: line, right: line };
const run = (t, o = {}) => new TextRun({ text: t, font: o.font || "Calibri", size: o.size || S, bold: o.bold, italics: o.it, color: o.color || C.dark });
// **fett** im Text
const runs = (t, o = {}) => String(t).split(/(\*\*[^*]+\*\*)/).filter(Boolean).map(x => run(x.startsWith("**") ? x.slice(2, -2) : x, { ...o, bold: x.startsWith("**") || o.bold }));
const para = (t, o = {}) => new Paragraph({ children: Array.isArray(t) ? t : runs(t, o), spacing: { after: o.after ?? 50, before: o.before ?? 0 } });
const cell = (children, w, fill, o = {}) => new TableCell({ children, width: { size: w, type: WidthType.DXA }, borders, verticalAlign: o.v || VerticalAlign.TOP, columnSpan: o.span,
  shading: fill ? { type: ShadingType.CLEAR, color: "auto", fill } : undefined, margins: { top: 130, bottom: 130, left: 130, right: 130 } });
const head = (cols, cw) => new TableRow({ tableHeader: true, children: cols.map((t, i) => cell([para(t, { bold: true, color: C.white, after: 0, size: 24 })], cw[i], C.dark)) });
const step = (nr, name, sub, fill, w) => cell([
  para([run(nr, { font: "Cambria", bold: true, size: 46, color: C.white })], { after: 0 }),
  para([run(name, { font: "Cambria", bold: true, size: 25, color: C.white })], { after: 20 }),
  ...(sub ? [para([run(sub, { size: 20, it: true, color: "F4ECDE" })], { after: 0 })] : []),
], w, fill, { v: VerticalAlign.CENTER });
const items = (arr, o = {}) => arr.map(t => para("▸ " + t, { after: 40, ...o }));
const title = (kick, t) => [para(kick, { size: 18, bold: true, color: C.terra, after: 20 }), para([run(t, { font: "Cambria", bold: true, size: 40 })], { after: 80 })];

const STEPS = [
  ["1", "Themenfrage", null, C.dark],
  ["2a", "Analyse: formal", "Wer? Was? Wann?", C.terra],
  ["2b", "Analyse: inhaltlich", "Was steht drin?", C.terra],
  ["3", "Historischer Kontext", "Was geschah damals?", C.blue],
  ["4a", "Sachurteil", "aus damaliger Sicht", C.green],
  ["4b", "Werturteil", "aus heutiger Sicht", C.green],
];

// ---------- Vorderseite: Arbeitsschritte ----------
const cw1 = [2250, 5056, 2900];
const tatte = [["T", "Titel"], ["A", "Autor"], ["T", "Textsorte"], ["T", "Thema"], ["E", "Entstehungszeit und -ort"], ["A", "Adressat"]]
  .map(([b, w]) => para([run(b + "  ", { font: "Cambria", bold: true, size: 34, color: C.terra }), run(w, { bold: true })], { after: 0 }));
const LEIT = [
  [items(["Welche Frage soll die Untersuchung beantworten?", "Sie steht in der Aufgabe – oder du stellst sie nach dem ersten Lesen selbst auf."]),
   "Die Frage am Ende im Urteil beantworten."],
  [[para("Die **TATTEA-Formel** – damit beginnt jede Interpretation:", { after: 60 }), ...tatte],
   "**Augenzeuge?** Lebensdaten mit dem Ereignis vergleichen. Kein Titel? Überschrift im Material nennen."],
  [items(["**Wesentliche Aussagen** je Sinnabschnitt", "**Schlüsselbegriffe**", "**Textsprache** (sachlich, emotional, appellativ, argumentativ …)"]),
   "Nicht nacherzählen: eigene Worte, indirekte Rede, Zeilen. **Noch keine Meinung.**"],
  [items(["In welches Ereignis, welche Epoche, welchen Konflikt gehört die Quelle?", "Was geschah vorher, was danach?"]),
   "Nur, was die Quelle erklärt – kein allgemeiner Aufsatz."],
  [items(["Standpunkt und **Absicht** des Autors?", "**Wirkung** auf die Adressaten?", "**Glaubwürdig?** Widersprüche? Sachlich angemessen?"]),
   "Immer mit Beleg aus Quelle oder Kontext."],
  [items(["Wie bewerte ich die Quelle im Hinblick auf die Themenfrage?", "**Maßstab offenlegen:** Gleichheit, Leistung, Schutz der Schwachen, gleiche Regeln"]),
   "„Meiner Meinung nach“ reicht nicht – Maßstab und Begründung nennen."],
];
const front = [
  ...title("GESCHICHTE · KLASSE 11 · FÜR DEN HEFTER", "So interpretierst du eine Textquelle"),
  para("Klausur: „Interpretieren Sie die Quelle.“ Gehe die Tabelle von oben nach unten durch – **erst untersuchen, dann urteilen.**", { size: 24, after: 140 }),
  new Table({ columnWidths: cw1, width: { size: W, type: WidthType.DXA }, rows: [
    head(["Schritt", "Was mache ich?", "⚠ Achtung"], cw1),
    ...STEPS.map(([nr, name, sub, fill], i) => new TableRow({ cantSplit: true, children: [
      step(nr, name, sub, fill, cw1[0]), cell(LEIT[i][0], cw1[1], i % 2 ? C.white : C.pale), cell([para(LEIT[i][1], { after: 0, size: 23 })], cw1[2], C.sand)] })),
    new TableRow({ children: [cell([
      para("**Operatoren:** analysieren (2) · einordnen, erläutern (3) · **beurteilen** = Sachurteil (4a) · **bewerten**, Stellung nehmen = Werturteil (4b)", { size: 22, after: 50 }),
      para("**Bildquelle:** gleiche Schritte; in 2b die Gestaltungsmittel **beschreiben** (Figuren, Symbole, Farben, Perspektive …) und **deuten**: Was ist die Botschaft?", { size: 22, after: 0 }),
    ], W, "EDE4D3", { span: 3 })] }),
  ] }),
  para("Rückseite: Formulierungshilfen für jeden Schritt", { size: 18, it: true, color: C.muted, before: 80, after: 0 }),
];

// ---------- Rückseite: Formulierungshilfen ----------
const cw2 = [2250, 7956];
const F = [
  ["„Die Untersuchung geht der Frage nach, ob / wie …“", "„Im Folgenden soll geprüft werden, inwieweit …“", "„Leitend ist die Frage, welche Rolle … spielte.“", "„Es stellt sich die Frage, ob …“"],
  ["„Bei der vorliegenden Quelle handelt es sich um [Textsorte] „…“ [Titel] von … [Autor].“", "„Der Autor war zu dieser Zeit … (Amt, Stellung).“ – „Bei dem Verfasser handelt es sich um …, der als … tätig war.“", "„Der Text entstand … [Jahr] in … [Ort] und richtet sich an … [Adressat].“", "„Er behandelt / thematisiert … [Thema].“"],
  ["**Aufbau:** „Zunächst … (Z. …)“ – „Anschließend / Im zweiten Abschnitt …“ – „Schließlich …“", "**Verben:** berichtet, erläutert, betont, hebt hervor, stellt dar, begründet … mit, kritisiert, fordert, verweist auf", "**Indirekte Rede:** „Er behauptet, die Bürger seien … / er habe …“", "**Sprache:** „Die Sprache ist überwiegend sachlich / emotional, was sich an … (Z. …) zeigt.“ – „Auffällig ist der Begriff …“"],
  ["„Die Quelle steht im Zusammenhang mit …“", "„Sie entstand zur Zeit …, als …“", "„Vorausgegangen war …“ – „In der Folgezeit …“", "„Dies erklärt, warum der Autor …“", "„Der Autor knüpft an … an.“ – „Damit wird deutlich, dass die Quelle Teil von … ist.“"],
  ["**Absicht:** „Der Autor verfolgt die Absicht, … / zielt darauf ab, …“ – „Er will die Adressaten dazu bewegen, …“", "**Glaubwürdigkeit:** „Für die Glaubwürdigkeit spricht, dass …; dagegen spricht …“ – „Da der Autor …, ist die Quelle nur eingeschränkt glaubwürdig.“", "**Darstellung:** „Der Autor stellt … einseitig / sachlich angemessen dar, denn …“", "**Vorsicht:** „Es ist nicht auszuschließen, dass …“ – „Vermutlich …“", "**Ergebnis:** „Insgesamt lässt sich festhalten, dass …“"],
  ["„Gemessen am Maßstab … halte ich … für …, weil …“", "„Aus heutiger Sicht ist … zu begrüßen / abzulehnen, da …“", "„Diesem Urteil liegt der Maßstab … zugrunde.“", "**Abwägen:** „Zugegebenermaßen …, allerdings …“ – „Dem ist entgegenzuhalten, dass …“", "**Fazit:** „Insgesamt komme ich zu dem Ergebnis, dass …“ – damit die Themenfrage beantworten"],
];
const back = [
  new Paragraph({ children: [new PageBreak()] }),
  ...title("RÜCKSEITE · FORMULIERUNGSHILFEN", "So kannst du formulieren"),
  new Table({ columnWidths: cw2, width: { size: W, type: WidthType.DXA }, rows: [
    head(["Schritt", "Formulierungen"], cw2),
    ...STEPS.map(([nr, name, sub, fill], i) => new TableRow({ cantSplit: true, children: [step(nr, name, null, fill, cw2[0]), cell(items(F[i], { size: 23, after: 40 }), cw2[1], i % 2 ? C.white : C.pale)] })),
    new TableRow({ children: [cell([
      para("**Verknüpfungswörter:** weil · da · denn · deshalb · folglich · somit · außerdem · zudem · jedoch · allerdings · dagegen · trotzdem · obwohl", { size: 23, after: 0 }),
    ], W, "EDE4D3", { span: 2 })] }),
  ] }),
];

const doc = new Document({ styles: { default: { document: { run: { font: "Calibri", size: S, color: C.dark } } } },
  sections: [{ properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 640, bottom: 500, left: 850, right: 850 } } }, children: [...front, ...back] }] });
Packer.toBuffer(doc).then(b => { fs.writeFileSync(OUT, b); console.log("done"); });
