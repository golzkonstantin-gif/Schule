const fs = require("fs");
const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, BorderStyle, ShadingType, AlignmentType,
  Header, Footer, PageNumber, ImageRun } = require("docx");
const { D, C, NR, card, nr } = require("./common.js");
const OUT = process.argv[2];
const W = 9638;
function runs(text, o = {}) {
  return text.split(/(\*\*[^*]+\*\*)/).filter(Boolean).map(t => new TextRun({ text: t.startsWith("**") ? t.slice(2, -2) : t, bold: t.startsWith("**") || o.bold, size: o.size, color: o.color, italics: o.italics, font: o.font }));
}
const p = (t, o = {}) => new Paragraph({ children: runs(t, o), spacing: { after: o.after ?? 80, before: o.before ?? 0 }, keepNext: o.keepNext });
const kicker = t => p(t.toUpperCase(), { size: 16, color: C.terra, bold: true, after: 40 });
const h1 = t => new Paragraph({ children: [new TextRun({ text: t, font: "Cambria", bold: true, size: 34, color: C.dark })], spacing: { after: 140 } });
const h2 = t => new Paragraph({ children: [new TextRun({ text: t, font: "Cambria", bold: true, size: 25, color: C.terra })], spacing: { before: 220, after: 100 }, keepNext: true });
const bullet = (t, o = {}) => p((o.box ? "☐  " : "–  ") + t, { size: o.size || 19, after: 50 });
const thin = { style: BorderStyle.SINGLE, size: 4, color: "BFB5A5" }, none = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
function cell(content, width, o = {}) {
  const children = (Array.isArray(content) ? content : [content]).map(c => typeof c === "string" ? p(c, { size: o.size ?? 17, after: 30 }) : c);
  return new TableCell({ children, width: { size: width, type: WidthType.DXA }, shading: o.fill ? { type: ShadingType.CLEAR, color: "auto", fill: o.fill } : undefined,
    borders: { top: thin, bottom: thin, left: thin, right: thin }, margins: { top: 60, bottom: 60, left: 90, right: 90 } });
}
function table(cw, head, rows, o = {}) {
  const hdr = new TableRow({ tableHeader: true, children: head.map((t, i) => cell(p(t, { bold: true, color: "FFFFFF", size: 17, after: 0 }), cw[i], { fill: C.dark })) });
  return new Table({ columnWidths: cw, width: { size: W, type: WidthType.DXA }, rows: [hdr, ...rows.map(r => new TableRow({ cantSplit: true,
    children: r.map((t, i) => cell(Array.isArray(t) ? t : t, cw[i], { fill: o.firstFill && i === 0 ? C.sand : undefined })) }))] });
}
function box(children, fill = C.sand) {
  return new Table({ columnWidths: [W], width: { size: W, type: WidthType.DXA }, rows: [new TableRow({ children: [new TableCell({ children, width: { size: W, type: WidthType.DXA },
    shading: { type: ShadingType.CLEAR, color: "auto", fill }, borders: { top: none, bottom: none, left: none, right: none }, margins: { top: 120, bottom: 120, left: 180, right: 180 } })] })] });
}
const img = (f, w, h) => new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 60 }, children: [new ImageRun({ type: "png", data: fs.readFileSync(f), transformation: { width: w, height: h } })] });

const VERLAUF = [
  ["0–2", "**A** Ankommen", "Ziel und Ablauf (Seitentafel + Folie). „Heute klären wir, wie Kleisthenes die Macht der Adelsfamilien brach – ohne ihnen Land wegzunehmen.“", "Transparenz: Die Klasse weiß, worauf alles hinausläuft (direkt).", "Plenum", "1–2"],
  ["2–7", "**V** Wiederholung", "Fehler-Übersicht zu Herrschaftsformen und Protagonisten (Aristokratie, Drakon, Solon, Peisistratos): allein 6 Fehler finden und im Heft berichtigen (2), Partnercheck (1), Stichproben bei 4–6 Schülern (2).", "Abruf aus dem Gedächtnis über die ganze Reihe (Klausur!) und Fehleranalyse statt Abfrage (direkt).", "EA → PA → Stichproben", "3"],
  ["7–9", "**V** Lösung + Leitfrage", "Lösungsfolie, Fehler grün verbessern. Dann die offene Leitfrage „Peisistratos – Wegbereiter der Demokratie?“: zwei Schüler antworten in einem Satz, danach der Satz unten auf der Folie.", "Offene Frage der letzten Stunde schließen; das „Aber“ (Adel bestimmt weiter über seine Anhänger) ist die Gelenkstelle zu heute.", "EA, Stichproben", "4"],
  ["9–10", "**V→I** Gelenkstelle", "Sturz der Tyrannis 514/510 kurz erzählen. „Jetzt kämpfen wieder die Adelsfamilien – wie kann man ihre Macht brechen, ohne Land zu nehmen?“", "Inhaltliche Brücke vom Sturz zur Reform (direkt).", "Lehrkraft", "5"],
  ["10–11", "**I** Rätsel", "Mystery-Frage vorlesen, jeder schreibt eine Vermutung ins Heft.", "Festlegen für den Aha-Moment; Leitfrage der Stunde.", "EA", "6"],
  ["11–13", "**I** Begriffe", "Demos und Trittys (Phyle ist bekannt). Prüffrage: „Was ist größer – Demos oder Trittys?“", "Begriffs-Check: Ohne die Begriffe sind die Schnipsel nicht lesbar (direkt).", "Plenum, Stichproben", "7"],
  ["13–15", "**I** Vormachen", "Zwei Schnipsel laut denkend verbinden („aber“), Warnung vor dem Fehler. Gruppen zu dritt bilden, Sets liegen bereit.", "Modellieren: Die Methode Mystery ist neu (direkt).", "Lehrkraft", "8"],
  ["15–29", "**V** Mystery", "Gruppen legen die Erklärungskette, jeder schreibt die Antwort ins Heft. Gestufte Hilfen am Pult. Nach 10 Min.: „Noch 4 Minuten – jetzt schreibt jeder.“", "Selbstständige Erschließung neuer Informationen (indirekt); Zusammenhänge selbst herstellen.", "GA (3)", "9"],
  ["29–33", "**V** Lösung", "Zwei Gruppen lesen vom Platz vor, du trägst die Kette an der Tafel ein. Vergleich mit der Vermutung.", "Puzzleteile zum Gesamtbild fügen; Aha auflösen.", "Plenum", "10"],
  ["33–38", "**V** Schaubild", "Phylenordnung erklären (2 Min.), Prüffragen an 3–4 Schüler. *Ostrakismos nur bei Zeit.", "Sicherung des Kerns als Schaubild (Klausur: Schaubild erklären) (direkt).", "Lehrkraft, Stichproben", "11"],
  ["38–45", "**A** Abschluss", "Abschlussaufgabe „Kleisthenes und Solon“ (4), zwei lesen vor, Rückbezug aufs Ziel, Ausblick Mittwoch.", "Prüft genau das Ziel; Verbindung zu Solon (Klausur). **Nie kürzen.**", "EA, Stichproben", "12–13"],
];

const content = [
  kicker("Handreichung · Geschichte Klasse 11 · Einzelstunde (45 Min., 10. Stunde)"),
  h1("Warum stimmt Lysias anders? – Die Reformen des Kleisthenes"),
  box([
    p("**Ziel (Ich kann …):** Ich kann erklären, wie Kleisthenes mit der neuen Phylenordnung die Macht der Adelsfamilien brach.", { size: 19 }),
    p("**Prüfaufgabe:** Abschlussaufgabe „Kleisthenes änderte an Solons Vermögensklassen nichts – trotzdem Wegbereiter der Demokratie. Erkläre.“ Erfüllt, wenn die Antwort den Mechanismus nennt: Mischung der Bürger in neuen Phylen → Gefolgschaft der Adligen zerschlagen → Stimmrecht frei nutzbar.", { size: 19 }),
    p("**Mindestziel:** Phylenordnung und ihre Wirkung (Mystery bis zur Lösungskette). Mit *: Schaubild ausführlich, Ostrakismos.", { size: 19 }),
    p("**Beitrag zur Reihe:** Solon gab Rechte (timokratisch mit demokratischen Zügen), Peisistratos schwächte den Adel teilweise, Kleisthenes bricht die Bindungen an die Adelsfamilien – Perikles vollendet die Demokratie. Klausur am 25.11.", { size: 19, after: 0 }),
  ]),
  p("**Bezug zur letzten Stunde:** Die Leitfrage „Ein Tyrann als Wegbereiter?“ wurde nicht mehr geschafft, und der Abruf aus dem Gedächtnis fehlte – beides steht am Anfang (Fehler-Übersicht zur ganzen Reihe, Wunsch der Lehrkraft). Vorstellen läuft vom Platz aus, du trägst ein (Rückmeldung Peisistratos).", { size: 19, before: 100 }),

  h2("Begriffs-Check"),
  table([2600, 7038], ["Begriff", "Status"], [
    ["Phyle", "bekannt (Lückentext Polis, Solon: Rat der 400 aus vier Phylen) – auf Folie 7 erinnert"],
    ["Demos (Demen), Trittys (Trittyen)", "**neu** – Begriffsfolie 7 vor dem Mystery"],
    ["Rat der 500, Ostrakismos", "neu, aber auf den Schnipseln erklärt"],
    ["Tyrannis, Alkmeoniden, Akropolis, Vermögensklassen", "bekannt aus den letzten Stunden"],
  ], { firstFill: true }),

  h2("1. Druckliste und Vorbereitung"),
  table([2900, 2100, 4638], ["Material", "Menge", "Hinweis"], [
    [["**Mystery-Schnipsel**", "Mystery_Schnipsel_und_Hilfen.docx, Seite 1"], "1 Blatt pro Dreiergruppe (ca. 6)", "Schnipsel sind gemischt nummeriert. Die Gruppen schneiden selbst (oder vorab schneiden und in Umschläge stecken, spart 2 Min.)."],
    [["**Gestufte Hilfen**", "dieselbe Datei, Seite 2"], "1–2 ×", "Schneiden, am Pult auslegen."],
    ["Scheren", "1 pro Gruppe", "Falls die Gruppen selbst schneiden."],
  ]),
  bullet("Seitentafel: Ziel und Ablauf (wie Folie 2).", { box: true }),
  bullet("Hauptfläche frei lassen für die Erklärungskette (Folie 10) und die Phylen-Skizze (Folie 11).", { box: true }),

  h2("2. Verlauf nach AVIVA"),
  table([700, 1400, 3300, 2338, 1100, 800], ["Zeit", "Phase", "Ablauf", "Warum", "Form", "Folie"], VERLAUF.map(r => r.map(x => p(x, { size: 16, after: 0 }))), { firstFill: true }),
  p("**Bei Zeitnot kürzen:** Schaubild (Folie 11) nur zeigen, nicht erklären; Ostrakismos weglassen. **Nie kürzen:** Abschlussaufgabe (Folien 12–13).", { size: 18, before: 80 }),

  h2("3. Vormachen (Folie 8) – laut denken"),
  box([p(`„Ich nehme zwei Schnipsel. Nummer ${NR.B} sagt: Seit Solon darf Lysias abstimmen – er hat also das Recht. Nummer ${NR.C} sagt: Sein Nachbar hat ihm Saatgut geliehen, seine Familie ist die mächtigste weit und breit. Wie passt das zusammen? Er hat das Recht, aber er ist abhängig. Ich lege die beiden nebeneinander und schreibe ‚aber‘ dazwischen. So baut ihr eure Kette: Was hängt womit zusammen – weil, deshalb, aber?“`, { size: 18, italics: true, after: 60 }),
    p("**Warnung vor dem Fehler:** „Kleisthenes nimmt niemandem Land weg und ändert Solons Vermögensklassen nicht. Die Lösung liegt woanders.“", { size: 18, after: 0 })]),

  h2("4. Erwartungshorizont"),
  p("**Fehler-Übersicht (Folie 3) – die 6 Fehler:**", { size: 18, keepNext: true }),
  table([2400, 3600, 3638], ["Protagonist", "Falsch", "Richtig"], [
    ["Aristokratie", "Recht sprechen gewählte Richter aus dem Volk.", "Recht spricht der Adel; geschriebene Gesetze gibt es nicht."],
    ["Drakon", "Er schafft die Schuldknechtschaft ab.", "Die Schuldknechtschaft bleibt bestehen."],
    ["Solon", "… und verteilt das Land der Großgrundbesitzer neu.", "Schulden gestrichen, aber keine Landverteilung."],
    ["Solon", "Rechte nach der Herkunft; Theten dürfen Archont werden.", "Rechte nach dem Besitz; Theten nur abstimmen und klagen."],
    ["Peisistratos", "Die Volksversammlung wählt ihn.", "Er erobert die Macht mit Söldnern, besetzt die Akropolis (546)."],
    ["Peisistratos", "Er schafft Solons Verfassung ab.", "Die Verfassung bleibt, die Ämter bekommen seine Anhänger."],
  ], { firstFill: true }),
  p("**Mystery – die Erklärungskette:**", { size: 18, before: 100, keepNext: true }),
  table([1500, 1500, 6638], ["Schritt", "Schnipsel", "Inhalt"], D.kette.map(([t, ids, d]) => [t, ids ? nr(ids) : "–", d]), { firstFill: true }),
  p(`**Ablenker:** Schnipsel ${nr("Q, R")} (Oliven, Großvater in Sikyon) sind wahr, helfen aber nicht. Sie dürfen aussortiert werden – das ist Teil der Methode.`, { size: 18, before: 80 }),
  p("**Musterantwort (Heft):** „Vorher war Lysias von seinem adligen Nachbarn abhängig und stimmte deshalb mit dessen Familie, obwohl er seit Solon abstimmen durfte. Nach dem Sturz des Hippias setzte Kleisthenes mit Hilfe des Volkes eine Reform durch: Er teilte Attika in drei Gebiete und bildete zehn neue Phylen, die Stadt, Küste und Binnenland mischten. Dadurch wurden die Anhänger des Adligen auf verschiedene Phylen verteilt, und Lysias saß im Rat und im Heer mit Fremden zusammen. Deshalb kann er sein Stimmrecht jetzt frei nutzen.“", { size: 18 }),

  h2("5. Beim Umhergehen: Stolpersteine und gestufte Hilfen"),
  table([3300, 6338], ["Stolperstein", "Impuls"], [
    ["Die Gruppe sucht eine Enteignung oder neue Rechte.", "„Steht auf einem Schnipsel, dass jemand Land verliert? Was ändert Kleisthenes stattdessen?“"],
    ["Die Ereignisse 510/508 (Sturz, Isagoras, Sparta) werden als Lösung genommen.", "„Das erklärt, warum Kleisthenes an die Macht kommt. Aber warum stimmt Lysias anders?“"],
    ["„Phyle“ wird mit „Demos“ verwechselt.", "„Was ist kleiner: dein Dorf oder deine Phyle?“ (Folie 7)"],
    ["Die Kette endet bei „Kleisthenes bildet zehn Phylen“.", `„Und was heißt das für Lysias und seinen Nachbarn? Lest ${NR.K} und ${NR.L}.“`],
    ["Ostrakismos wird zur Hauptursache.", "„Hilft das Scherbengericht Lysias bei der Abstimmung im Alltag – oder gegen wen ist es gedacht?“"],
  ], { firstFill: true }),
  p(`**Gestufte Hilfen am Pult:** Hilfe 1 – in drei Stapel sortieren (Vorher, Wende, Nachher). Hilfe 2 – Schnipsel ${NR.D} und ${NR.J} vergleichen (alte und neue Phylen). Hilfe 3 – Schnipsel ${NR.K} und ${NR.L}: Wo landen die Anhänger des Nachbarn, mit wem ist Lysias zusammen?`, { size: 18, before: 80 }),

  h2("6. Tafelskizze"),
  p("**Hauptfläche:** zuerst die Kette Ausgangslage → Wende → Reform → Folge → Ergebnis (Stichworte der Gruppen), dann die Phylen-Skizze:", { size: 18, keepNext: true }),
  img(__dirname + "/schaubild-11.png", 520, 293),
  p("Schnell zu zeichnen: drei Reihen à 10 Kästchen (Stadt, Küste, Binnenland, in drei Kreidefarben), Pfeil „per Los“, zehn Säulen aus je drei farbigen Kästchen, darunter „× 50 = Rat der 500“.", { size: 17, color: C.muted }),

  h2("7. Abschlussaufgabe – mögliche Argumente"),
  bullet("Solon gab allen Bürgern das Stimmrecht und das Klagerecht – aber die Bauern blieben an ihre Adelsherren gebunden."),
  bullet("Kleisthenes änderte nicht die Rechte (Vermögensklassen, Ämter), sondern die Ordnung der Bürgerschaft: neue, gemischte Phylen."),
  bullet("Die Gefolgschaft eines Adligen wurde zerteilt; im Rat der 500 und im Heer saßen Bürger aus allen Gebieten zusammen."),
  bullet("Bürger ist, wer zu seinem Demos gehört – nicht, aus welcher Familie er kommt (Isonomie: gleiche Ordnung für alle Bürger)."),
  bullet("Zusatz: Mit dem Ostrakismos kann das Volk einen neuen Tyrannen verhindern."),

  h2("8. Sachinformationen zu Kleisthenes"),
  p("Quellen: Herodot, Historien 5,66–78; Aristoteles, Staat der Athener 20–22; Lehrbuch „Die antike Welt“, S. 16 f. Lysias ist erfunden; Paiania (Binnenland) gehörte tatsächlich zur Phyle Pandionis.", { size: 17, color: C.muted }),
  bullet("Kleisthenes war **Alkmeonide**, unter Hippias selbst Archon (525/24). Nach 510 Machtkampf mit **Isagoras** (Archon 508/07). Kleisthenes gewann das Volk für sich; Isagoras holte den spartanischen König Kleomenes, rund 700 Familien wurden verbannt. Das **Volk** belagerte die Spartaner auf der Akropolis und erzwang ihren Abzug – danach Reformen 508/07."),
  bullet("**Demen** (rund 139, Zahl modern rekonstruiert) als Grundeinheit; Bürgerliste im Demos; Name nach dem Demos (z. B. „Perikles aus Cholargos“)."),
  bullet("**3 Gebiete** (Stadt, Küste, Binnenland) × **10 Trittyen** = 30 Trittyen; **10 Phylen** aus je einer Trittys pro Gebiet, durch das Los zugeteilt. Die vier alten Phylen verlieren ihre politische Bedeutung."),
  bullet("**Rat der 500:** 50 pro Phyle, nach Größe der Demen; Mindestalter 30; bereitet die Volksversammlung vor. Prytanien und Losverfahren für den Rat sind für Kleisthenes' Zeit umstritten."),
  bullet("**Ostrakismos:** Aristoteles schreibt ihn Kleisthenes zu; erst 488/87 angewendet. 6000 Stimmen (umstritten, ob insgesamt oder gegen eine Person), 10 Jahre Verbannung ohne Verlust von Vermögen und Bürgerrecht. Ab 501/500 zehn Strategen, je einer pro Phyle."),
  bullet("**Unverändert:** Solons Vermögensklassen, Archonten aus den oberen Klassen, Areopag (bis 462). Frauen, Sklaven und Fremde ausgeschlossen. Leitbegriff **Isonomie** (gleiche Ordnung); das Wort „Demokratie“ ist erst später belegt."),

  h2("9. Ausblick Mittwoch"),
  p("Test, danach Doppelstunde Quellenanalyse mit Schreibplan – vorgeschlagen an Aristoteles, Staat der Athener 20–21 (die Reform des Kleisthenes in seinen Worten, etwa 180 Jahre später geschrieben). Als Wiederholung vor dem Test eignet sich ein fehlerhafter Lexikonartikel zu Kleisthenes.", { size: 19 }),
];
const doc = new Document({
  styles: { default: { document: { run: { font: "Calibri", size: 20, color: C.dark } } } },
  sections: [{ properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 900, bottom: 800, left: 1134, right: 1134 } } },
    headers: { default: new Header({ children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ text: "Geschichte · Klasse 11 · Handreichung „Kleisthenes“", size: 16, color: "998C77" })] })] }) },
    footers: { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ children: [PageNumber.CURRENT], size: 16, color: "998C77" })] })] }) },
    children: content }],
});
Packer.toBuffer(doc).then(b => { fs.writeFileSync(OUT, b); console.log("done"); });
