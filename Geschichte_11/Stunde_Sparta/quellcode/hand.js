// Handreichung Sparta: Kopf, Druckliste, Verlauf (Input 5 – Einzelarbeit 25 – Auswertung 15) nach AVIVA mit Warum-Zeilen, Erwartungshorizont, Tafelskizze, Argumente zur Abschlussaufgabe.
const fs = require("fs");
const { Document, Packer, Paragraph, TextRun, ImageRun } = require("docx");
const { C, p, h, table } = require("../../Probeklausur_PseudoXenophon/quellcode/lib.js");
const D = require("./data.js");
const OUT = process.argv[2];
const W = 9638;
const children = [
  p("GESCHICHTE · KLASSE 11 · HANDREICHUNG · EINZELSTUNDE (45 MIN.)", { size: 16, bold: true, color: C.terra }),
  new Paragraph({ children: [new TextRun({ text: D.titel, font: "Cambria", bold: true, size: 34, color: C.dark })], spacing: { after: 120 } }),
  table(W, [2200, 7438], null, [
    ["Ziel", D.ziel],
    ["Prüfaufgabe", D.abschluss + " – „Ich erkläre es dir: …, weil …“ (Folie 11)."],
    ["Mindestziel", "Fragen 1–6 beantwortet; das Schaubild gemeinsam ausgefüllt."],
    ["Beitrag zur Reihe", "Erste Annäherung an Sparta (BM 2: Attika und Sparta im Vergleich). Die Doppelstunde vergleicht dann Athen und Sparta quellengestützt – der Hinweis, dass fast alle Quellen von Nicht-Spartanern stammen (Text Z. 39f.), leitet dorthin über."],
  ], { firstFill: true, size: 19 }),
  h("Drucken"),
  p("–  **Text_Sparta** (A4, einseitig) – 1 pro Schüler.", { size: 19 }),
  p("–  **Fragen_Sparta** (A4 doppelseitig: vorne Fragentabelle, hinten stummes Schaubild quer) – 1 pro Schüler. Text und Fragen liegen nebeneinander.", { size: 19 }),
  p("–  **Hilfen_Lehrkraft_Sparta** und **Antworten_Sparta** – nur für dich.", { size: 19 }),
  h("Verlauf"),
  table(W, [800, 1500, 4338, 2200, 800], ["Min.", "Phase", "Ablauf und Ansage", "Warum?", "Folie"], [
    ["0–1", "A Ankommen", "Ziel und Ablauf. Mindestziel nennen.", "Transparenz (direkt)", "1–2"],
    ["1–2", "V Vorwissen", "„Das ist aber spartanisch!“ – Was meint man damit? 2 Schüler. „Heute: Woher kommt dieses Bild?“", "an Bekanntes anknüpfen, Neugier (direkt)", "3"],
    ["2–5", "I Informieren", "Sparta auf einen Blick (Lage, Gegenstück zu Athen, Eroberung Messeniens); Begriffe Spartiaten, Periöken, Heloten. Warnung: Spartaner ≠ Spartiaten.", "Orientierung und Begriffe vor dem Lesen (direkt)", "4–5"],
    ["5–30", "V Verarbeiten", "Einzelarbeit: Text lesen, Fragen 1–8 in Stichworten mit Zeilen; wer fertig ist, füllt das Schaubild. Du gehst mit dem Hilfeblatt herum.", "selbstständiges Erschließen (indirekt), gestufte Hilfen individuell", "6"],
    ["30–35", "A Auswerten", "Lösungen 1–8: vor jeder Lösung eine Stichprobe, dann vergleichen.", "Sicherung, Kontrolle über Lösungsfolie", "7–8"],
    ["35–40", "A Auswerten", "Schaubild an der Tafel gemeinsam ausfüllen (Schüler sagen, du trägst ein, alle schreiben mit), Lösungsfolie.", "Wissen ordnen und verknüpfen", "9–10"],
    ["40–45", "A Auswerten", "Abschlussaufgabe schriftlich 2 Min., 2–3 vorlesen; Beispiel; Ziel; Ausblick auf die Doppelstunde.", "prüft das Ziel; nie streichen", "11–12"],
  ], { size: 17 }),
  p("Bei Zeitnot: Frage 8 und das selbstständige Ausfüllen des Schaubilds entfallen (Schaubild nur gemeinsam). Die Abschlussaufgabe bleibt.", { size: 18, italics: true, color: C.muted, before: 80 }),
  h("Erwartungshorizont"),
  table(W, [3000, 900, 5738], ["Frage", "Zeilen", "Erwartung"], D.fragen.map(([q, z, e], i) => [(i + 1) + "  " + q, z, e]), { size: 17 }),
  h("Tafelskizze (Lösung)"),
  p("So einfach an die Tafel zeichnen: links drei Kästen (Spartiaten klein, Periöken mittel, Heloten groß) mit vier Spalten zum Ankreuzen, rechts vier Kästen (Könige – Ephoren – Gerusia – Volksversammlung) mit Pfeilen.", { size: 18 }),
  new Paragraph({ children: [new ImageRun({ data: fs.readFileSync(__dirname + "/../Schaubild_Sparta_Loesung.png"), transformation: { width: 640, height: 336 }, type: "png" })] }),
  h("Abschlussaufgabe: mögliche Argumente"),
  ...D.argumente.map(a => p("–  " + a, { size: 19 })),
  p("**Beispiel:** " + D.beispiel, { size: 19, before: 60 }),
];
const doc = new Document({ styles: { default: { document: { run: { font: "Calibri", size: 20, color: C.dark } } } },
  sections: [{ properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 800, bottom: 700, left: 1134, right: 1134 } } }, children }] });
Packer.toBuffer(doc).then(b => { fs.writeFileSync(OUT, b); console.log("done"); });
