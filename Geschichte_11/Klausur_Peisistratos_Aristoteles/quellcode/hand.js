// Handreichung Teil 2 (TATTEA und Einleitung zur Übungsklausur Peisistratos): Kopf, Druckliste, Verlauf nach AVIVA mit Warum-Zeilen, Lösung TATTEA, Muster-Einleitung, Stolpersteine.
const fs = require("fs");
const { Document, Packer, Paragraph, TextRun, BorderStyle } = require("docx");
const { C, p, h, table } = require("../../Probeklausur_PseudoXenophon/quellcode/lib.js");
const D = require("./data.js");
const OUT = process.argv[2];
const W = 9638;
const children = [
  p("GESCHICHTE · KLASSE 11 · HANDREICHUNG · TEIL 2 DER STUNDE (ca. 45 MIN.)", { size: 16, bold: true, color: C.terra }),
  new Paragraph({ children: [new TextRun({ text: "Übungsklausur Peisistratos: TATTEA und Einleitung", font: "Cambria", bold: true, size: 34, color: C.dark })], spacing: { after: 120 } }),
  table(W, [2200, 7438], null, [
    ["Ziel", "Ich kann zu einer neuen Quelle TATTEA bestimmen und die Einleitung schreiben."],
    ["Prüfaufgabe", "Fehler-Einleitung („Aristoteles berichtet als Augenzeuge …“) verbessern: „Ich rate dir, …, weil …“ (Folie 16)."],
    ["Mindestziel", "TATTEA vollständig in Stichworten; die Einleitung mindestens begonnen (Sätze 1 und 2)."],
    ["Beitrag zur Reihe", "Teil 1 übte das Ende einer Interpretation (Urteil), Teil 2 den Anfang (Einleitung) – an einer neuen Quelle, wie in der Klausur am 25.11."],
  ], { firstFill: true, size: 19 }),
  h("Drucken"),
  p("–  **Uebungsklausur_Peisistratos** (A4 doppelseitig: Aufgabenstellung, M1 mit Zeilennummern wie im Original) – 1 pro Schüler.", { size: 19 }),
  p("–  **TATTEA_Einleitung_A5** (2 × A5 auf A4 quer, schneiden) – 1 pro Schüler.", { size: 19 }),
  p("–  Übersicht „So interpretierst du eine Textquelle“ liegt im Hefter.", { size: 19 }),
  h("Verlauf"),
  table(W, [800, 1500, 4338, 2200, 800], ["Min.", "Phase", "Ablauf und Ansage", "Warum?", "Folie"], [
    ["0–1", "A Ankommen", "Ziel und Ablauf. „Eben das Ende einer Interpretation, jetzt der Anfang – an einer neuen Klausur.“", "Transparenz, Gelenk zu Teil 1 (direkt)", "1–2"],
    ["1–4", "V Vorwissen", "TATTEA allein aus dem Gedächtnis ins Heft, Stichprobe 4 Schüler, Lösung.", "Wissens-Check vor der Anwendung (direkt)", "3–4"],
    ["4–11", "I Informieren", "Begriffe Fürsorge/Machtsicherung; Fundstellen (Infokasten, Nachweis, Text, erschließen); Warnung: Autor umstritten, 200 Jahre später; Vormachen Titel und Autor laut denkend.", "Neue Quelle, zwei typische Fallen – erst zeigen, dann üben (direkt)", "5–8"],
    ["11–22", "V Verarbeiten", "TATTEA allein in die Tabelle (8 Min.), Infokasten an der Wand; Lösung vergleichen.", "Erste Anwendung, leicht: alle schaffen sie", "9–12"],
    ["22–35", "V Verarbeiten", "Einleitung ins Heft (10 Min.) mit Satzanfängen zum Abhaken; Muster vergleichen, in anderer Farbe ergänzen.", "Angeleitetes Üben mit Teilhilfen (Satzanfänge)", "13–15"],
    ["35–45", "A Auswerten", "Fehler-Einleitung zu zweit (2 Min.), Blitzlicht 3–4 Paare „Ich rate dir …, weil …“; Lösung; Ziel.", "Prüft genau das Ziel; nie streichen", "16–17"],
  ], { size: 17 }),
  p("Bei Zeitnot: Folien 5–6 kürzer (Begriffe nur vorlesen), Einleitung nur Sätze 1–2 schreiben lassen, Rest als Hausaufgabe. Das Auswerten bleibt.", { size: 18, italics: true, color: C.muted, before: 80 }),
  h("Lösung TATTEA"),
  table(W, [1800, 6038, 1800], ["", "Stichworte", "Fundstelle"], D.tattea, { firstFill: true, size: 19 }),
  h("Muster-Einleitung"),
  p(D.einleitung, { size: 19, line: 290 }),
  h("Stolpersteine und Hilfsimpulse"),
  table(W, [2600, 7038], ["Stolperstein", "Impuls (gestuft)"], [
    ["Titel = „M1: Aristoteles über …“", "„Das ist die Überschrift der Klausur. Wie heißt die Schrift? → Infokasten.“"],
    ["Autor = „Aristoteles“ ohne Einschränkung", "„Lies den Infokasten noch einmal: Was steht dort über den Autor?“ → „Als Verfasser gilt …“"],
    ["Entstehung = Zeit des Peisistratos", "„Wann herrschte Peisistratos? Wann wurde der Text geschrieben? Wie viele Jahre liegen dazwischen?“"],
    ["Thema mit Wertung („wie gut er war“)", "„Worüber schreibt der Text – nicht: was hältst du davon?“ → „Sie behandelt …“"],
    ["Themenfrage fehlt am Ende", "„Womit endet eine Einleitung?“ → „Die Untersuchung geht der Frage nach, ob …“"],
  ], { size: 18 }),
];
const doc = new Document({ styles: { default: { document: { run: { font: "Calibri", size: 20, color: C.dark } } } },
  sections: [{ properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 800, bottom: 700, left: 1134, right: 1134 } } }, children }] });
Packer.toBuffer(doc).then(b => { fs.writeFileSync(OUT, b); console.log("done"); });
