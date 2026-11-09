// Handreichung „Perikles – Herrschaft des Volkes oder eines Mannes?“ (Einzelstunde, klassische Textarbeit). Tintensparend.
const fs = require("fs");
const { Document, Packer, Paragraph, TextRun, AlignmentType, Header } = require("docx");
const { C, p, h, table, box } = require("./lib.js");
const D = require("./data.js");
const OUT = process.argv[2];
const W = 9638, s = 19;
const li = t => p("–  " + t, { size: s, after: 40 });

const VERLAUF = [
  ["0–2", "**A** Ankommen", "Ziel und Ablauf (Seitentafel + Folie). „Heute ist das letzte Stück auf dem Weg zur Demokratie dran – am Ende entscheidet ihr, ob 430 wirklich das Volk regiert hat.“", "Transparenz; Leitfrage öffnen (direkt).", "Plenum", "2"],
  ["2–8", "**V** Wiederholung Kleisthenes", "Buch und Hefter zu: fünf Fragen allein ins Heft (2), Stichproben (1), Lösung (1), Treppe Dorf → Bezirk → Phyle → Rat erklären lassen (2).", "Abruf aus dem Gedächtnis – die Reform war schwer (Rückmeldung 4.11.); Grundlage für den neuen Text.", "EA → Stichproben", "3–5"],
  ["8–10", "**I** Gelenkstelle + Daumen", "„Kleisthenes entmachtete den Adel als Block – aber die höchsten Ämter hatten noch die Reichen, und der Areopag kontrollierte alles.“ Daumenabfrage: Wer regiert 430 – das Volk oder ein Mann?", "Inhaltliche Brücke; Festlegen für den Aha-Moment.", "Plenum", "6"],
  ["10–12", "**I** Begriffe", "Stratege, Diäten. Prüffrage zu Diäten.", "Begriffs-Check: Beide Begriffe braucht der Text (direkt).", "Lehrkraft, Stichproben", "7–8"],
  ["12–14", "**I** Vormachen", "Frage 1 laut denkend: Signalwort „zur Strafe für“ (Z. 5) suchen, Stichwort in eigenen Worten. Warnung: nicht abschreiben.", "Modellieren der Textarbeit (direkt).", "Lehrkraft", "9"],
  ["14–30", "**V** Lesen + Fragen", "Einzelarbeit mit Textblatt und Fragentabelle: Fragen 2–8 in Stichworten in die Tabelle. Wer festhängt, meldet sich; Hilfe individuell nach dem Hilfeblatt. Letzte 3 Min.: Partnervergleich.", "Selbstständige Texterschließung (indirekt), klassische Textarbeit auf Wunsch der Lehrkraft.", "EA → PA", "10–13"],
  ["30–34", "**V** Lösung", "Je Frage ein Schüler vom Platz, aufdecken, grün ergänzen.", "Kontrolle über Lösungsfolie, wenig Unterrichtsgespräch.", "Stichproben", "14–15"],
  ["34–36", "**V** Zeitleiste", "Vier Stationen an der Tafel, alle ins Heft. Prüffrage: „Welche Maßnahme half den Ärmsten am meisten?“", "Sicherung: der Demokratisierungsprozess im Überblick (Klausur).", "Lehrkraft, Stichproben", "16"],
  ["36–37", "**A** Aha", "Thukydides: „dem Namen nach eine Volksherrschaft …“ – Vergleich mit der Daumenabfrage.", "Auflösung; leitet zum Urteil über.", "Plenum", "17"],
  ["37–43", "**A** Abschluss", "„Bei uns herrscht in Wahrheit nur einer – Perikles.“ Hat er recht? 2–3 Sätze (4), drei lesen vor, Beispiel (2).", "Prüft das Ziel: Position, Belege aus dem Text, „weil“. **Nie kürzen.**", "EA, Stichproben", "18–19"],
  ["43–45", "**A** Ziel", "Rückbezug aufs Ziel, Hinweis Klausur 25.11.", "Transparenz.", "Plenum", "20"],
];

const content = [
  p("HANDREICHUNG · GESCHICHTE KLASSE 11 · EINZELSTUNDE (45 MIN.)", { size: 16, bold: true, color: C.terra, after: 30 }),
  new Paragraph({ children: [new TextRun({ text: D.titel, font: "Cambria", bold: true, size: 34 })], spacing: { after: 140 } }),
  box(W, [
    p("**Ziel (Ich kann …):** " + D.ziel, { size: s }),
    p("**Prüfaufgabe:** " + D.abschluss + " Erfüllt, wenn die Antwort eine Position bezieht und sie mit mindestens einem Beleg aus dem Text begründet (z. B. jedes Jahr Stratege / Volksversammlung entscheidet, Ostrakismos).", { size: s }),
    p("**Mindestziel:** Fragen 1–7 (Maßnahmen von Themistokles, Ephialtes und Perikles) und die Abschlussaufgabe. Mit *: Frage 8 (Macht des Perikles).", { size: s }),
    p("**Beitrag zur Reihe:** Abschluss des Demokratisierungsprozesses (Solon → Peisistratos → Kleisthenes → Themistokles/Ephialtes/Perikles); Perikles ist Klausurthema am 25.11.", { size: s, after: 0 }),
  ], "FFFFFF"),
  p("**Bezug zur letzten Stunde:** Die Reform des Kleisthenes war für die Klasse schwer nachzuvollziehen – deshalb steht eine kurze Wiederholung am Anfang (Prüffragen und Treppe aus der Wiederholungspräsentation). Methode auf Wunsch der Lehrkraft: klassische Textarbeit (Lesen + Fragen) statt Mystery oder Fälle; Text als eigenes Blatt (ohne Buch), Fragen als Tabelle zum Ausfüllen daneben, Vormachen einer Frage, gestufte Hilfen individuell durch die Lehrkraft (Hilfeblatt), Lösung über Folie.", { size: s, before: 100 }),

  h("Druckliste"),
  table(W, [5000, 2600, 2038], ["Material", "Format", "Menge"], [
    ["Text_Perikles (Text mit Zeilennummern)", "A4 hoch, einseitig", "1 je Schüler"],
    ["Fragentabelle (Fragen und Antwortspalte)", "A4 hoch", "1 je Schüler"],
    ["Antworten_Fragentabelle (Lösungen mit Zeilen)", "A4 hoch", "1 × (Lehrkraft)"],
    ["Hilfen_Lehrkraft (gestufte Hilfen + Lösungen)", "A4 quer", "1 × (Lehrkraft)"],
    ["Folien_Perikles (20 Folien)", "Beamer", "–"],
  ], { size: 18 }),

  h("Begriffs-Check"),
  table(W, [3000, 6638], ["Begriff", "Status"], [
    ["Archonten, Areopag, Theten, Rat", "bekannt (Solon, Kleisthenes)"],
    ["Ostrakismos", "bekannt aus dem Mystery (Kleisthenes)"],
    ["Stratege, Diäten", "**neu** – Begriffsfolien 7–8 vor dem Lesen"],
    ["Losverfahren, demagogós", "im Text erklärt; Frage 3 und 5 fragen danach"],
    ["Perserkriege, Themistokles, Ephialtes", "neu, aber im Text und in der Randspalte erklärt"],
  ], { size: 18 }),

  h("Verlauf nach AVIVA"),
  table(W, [700, 1500, 3400, 2338, 1000, 700], ["Min.", "Phase", "Ablauf und Ansage", "Warum?", "Form", "Folie"], VERLAUF, { size: 16 }),
  p("**Bei Zeitnot kürzen:** Frage 8 weglassen; Zeitleiste (Folie 16) nur zeigen. **Nie kürzen:** Abschlussaufgabe (Folien 18–19).", { size: 18, before: 80 }),

  h("Erwartungshorizont: Fragen"),
  table(W, [600, 2700, 900, 5438], ["", "Frage", "Zeilen", "Erwartung"], D.fragen.map(([q, z, e], i) => [String(i + 1) + (i === 7 ? "*" : ""), q, z, e]), { size: 17, firstFill: false }),

  h("Stolpersteine und gestufte Hilfsimpulse"),
  table(W, [3200, 6438], ["Stolperstein", "Impuls"], [
    ["Ganze Sätze werden abgeschrieben.", "„Deck den Text ab und sag es mir in fünf Wörtern.“"],
    ["Frage 3: „Themistokles schaffte die Archonten ab.“", "„Gibt es die Archonten danach noch? Was ändert sich daran, wie man Archont wird?“ (gelost statt gewählt)"],
    ["Frage 3: Wer gewinnt? – keine Antwort", "„Welche Beamten wurden weiter gewählt?“ (Strategen)"],
    ["Frage 4: Verbannte verlieren alles.", "„Lies Zeile 19 und 20: Was behält der Verbannte?“"],
    ["Frage 5: „demagogós“ = Volksverhetzer (heutige Bedeutung).", "„Wie übersetzt der Text das Wort (Z. 23f.)? Damals war es kein Schimpfwort.“"],
    ["Frage 6: Der Areopag wird abgeschafft.", "„Wofür ist er danach noch zuständig?“ (Mordprozesse)"],
  ], { size: 18 }),
  p("**Gestufte Hilfen (Hilfeblatt, nur Lehrkraft):** individuell beim Umhergehen – Stufe 1 Fundstelle (Zeilen), Stufe 2 Impulsfrage, Stufe 3 Satzanfang. Keine Hilfekarten am Pult (Wunsch der Lehrkraft: sonst holen sich alle Karten statt zu lesen).", { size: s, before: 80 }),

  h("Abschlussaufgabe – mögliche Argumente"),
  p("**Ja, ein Mann herrschte:**", { size: s, after: 30 }),
  ...D.argumente.ja.map(li),
  p("**Nein, das Volk herrschte:**", { size: s, before: 60, after: 30 }),
  ...D.argumente.nein.map(li),
  p("**Beispielantwort (Folie 19):** " + D.beispiel, { size: s, before: 80 }),
  p("Alle Positionen sind richtig, wenn sie belegt und mit „weil“ begründet sind. Klausur-Hinweis: Das ist ein Sachurteil aus damaliger Sicht.", { size: 18, color: C.muted }),

  h("Tafelbild"),
  p("Zeitstrahl mit vier Stationen: **508/07 Kleisthenes** – Bürger gemischt · **487/86 Themistokles** – Archonten gelost (Strategen gewählt), Ostrakismos durch das Volk; Flotte → Theten wichtig · **462 Ephialtes und Perikles** – Areopag entmachtet · **ab 462 Perikles** – Diäten. Darunter: „Jetzt spricht man von ‚Demokratie‘.“ Links daneben das Ergebnis der Daumenabfrage.", { size: s }),

  h("Sachinformationen"),
  li("Grundlage: eigener, vereinfachter Text nach dem Lehrbuch S. 17 f. (Darstellung); das Thukydides-Zitat steht bewusst nicht im Text, es ist der Aha-Moment am Ende. Die im Buch genannten M7a, b und M8 lagen nicht vor."),
  li("Themistokles: Archon 493/92; Flottenbau mit dem Silber aus Laureion; Sieg bei Salamis 480. Die Theten als Ruderer gelten als Grund für die spätere politische Aufwertung der Unterschicht."),
  li("487/86: Archonten werden aus Vorgewählten gelost (klērōsis ek prokritōn); die zehn Strategen bleiben gewählt und werden die eigentlichen Führungsämter. Erster bekannter Ostrakismos 488/87."),
  li("462/61: Ephialtes (kurz danach ermordet) und Perikles entziehen dem Areopag seine politischen Befugnisse; ihm bleibt die Blutgerichtsbarkeit."),
  li("Diäten (misthos): zuerst für die Richter, dann für Rat und viele Ämter. Perikles war 443–430 (bis zu seinem Tod 429) jedes Jahr Stratege."),
  li("Thukydides (Historien 2,65): Athen sei „dem Namen nach eine Demokratie, in Wirklichkeit die Herrschaft des ersten Mannes“ gewesen. Thukydides war Zeitgenosse und Bewunderer des Perikles."),
  li("Weiterhin ausgeschlossen: Frauen, Sklaven, Fremde (Metöken). Für ein Werturteil in der Klausur bietet sich der Maßstab Gleichheit an."),
];

const doc = new Document({ styles: { default: { document: { run: { font: "Calibri", size: 20, color: C.dark } } } },
  sections: [{ properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 900, bottom: 800, left: 1134, right: 1134 } } },
    headers: { default: new Header({ children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ text: "Geschichte · Klasse 11 · Handreichung „Perikles“", size: 16, color: "998C77" })] })] }) },
    children: content }] });
Packer.toBuffer(doc).then(b => { fs.writeFileSync(OUT, b); console.log("done"); });
