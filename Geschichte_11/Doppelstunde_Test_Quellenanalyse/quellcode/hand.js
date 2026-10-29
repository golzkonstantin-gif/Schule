// Handreichung Doppelstunde: Test + Quellenanalyse mit Schreibplan
const fs = require("fs");
const { Document, Packer, Paragraph, TextRun, AlignmentType, Header } = require("docx");
const { C, p, h, table, box } = require("./lib.js");
const D = require("./data.js");
const OUT = process.argv[2];
const W = 9638;
const s = 19;
const li = t => p("–  " + t, { size: s, after: 40 });

const VERLAUF = [
  ["0–2", "**A** Ankommen", "Tests einsammeln, Blatt austeilen. Ziel und Ablauf: „Heute lernt ihr, wie man in der Klausur mit einer Quelle umgeht – Schritt für Schritt.“", "Transparenz; Bedeutung für die Klausur (direkt).", "Plenum", "2–3"],
  ["2–7", "**V** Richtig oder falsch?", "Vier Aussagen (Quelle/Darstellung, Kleisthenes): allein ins Heft (2), Partnercheck (1), Stichproben bei 4–6 Schülern (1), Lösung (1).", "Echte Wiederholung von Montag und Begriff „Quelle“ aktivieren, beides braucht die Analyse (direkt).", "EA → PA → Stichproben", "4–5"],
  ["7–8", "**I** Warum?", "„In der Klausur ist die erste Aufgabe meist: Analysiere die Quelle.“", "Bedeutung erklären (DI).", "Lehrkraft", "6"],
  ["8–10", "**I** Portion 1", "Der Schreibplan in vier Schritten. Warnung: nicht nacherzählen. Prüffragen an 4–6 Schüler.", "Überblick vor den Einzelschritten (direkt, neue Kompetenz).", "Lehrkraft, Stichproben", "7"],
  ["10–13", "**I** Portion 2 + Modell", "Fünf W-Fragen der Einleitung; an Solons Gedicht laut denkend vormachen, Einleitungssatz zeigen. Warnung: Lebensdaten prüfen.", "Modellieren an einer anderen Quelle, damit die Lösung für Aristoteles nicht vorweggenommen wird.", "Lehrkraft, Stichproben", "8–10"],
  ["13–19", "**V** Einleitung", "Allein: fünf Felder ausfüllen, Einleitungssatz (4). Zwei lesen vor, Lösung (2).", "Erste Anwendung: leicht, alle schaffen sie (Infokasten liefert alles).", "EA", "11–13"],
  ["19–23", "**I** Portion 3 + Modell", "Regeln für die Inhaltswiedergabe (Sinnabschnitte, eigene Worte, indirekte Rede, Zeilen). Warnung: keine Meinung. Abschnitt Z. 1–4 laut denkend vormachen.", "Neue Teilkompetenz modellieren; das Beispiel steht im Schreibplan (✓).", "Lehrkraft, Stichproben", "14–15"],
  ["23–36", "**V** Inhalt", "Paare: Z. 5–8 mit Satzanfang und Stichworten, Z. 9–12 mit Satzanfang, Z. 12–14 ohne Hilfe (10). Quelle an der Wand. Je Abschnitt ein Paar vom Platz, Lösung (3).", "Angeleitetes → selbstständiges Üben; Teillösungen nehmen ab (indirekt nach dem Modell).", "PA", "16–22"],
  ["*36–38", "**V** Schritte 3 und 4*", "Nur bei Zeit: Kontext (Sturz der Tyrannis) und Glaubwürdigkeit (180 Jahre später) kurz erklären, zwei Prüffragen.", "Ausblick auf die nächste Doppelstunde; wird dort vertieft.", "Lehrkraft", "23"],
  ["38–45", "**A** Abschluss", "Abschlussaufgabe „Ich rate dir …, weil …“ (4), drei Schüler lesen je einen Ratschlag vor, Lösung (2), Rückbezug aufs Ziel (1).", "Prüft genau das Ziel: Fehler in Einleitung und Inhalt erkennen und begründen. **Nie kürzen.**", "EA, Stichproben", "24–26"],
];

const content = [
  p("HANDREICHUNG · GESCHICHTE KLASSE 11 · DOPPELSTUNDE (90 MIN.)", { size: 16, bold: true, color: C.terra, after: 30 }),
  new Paragraph({ children: [new TextRun({ text: "Test und Quellenanalyse: Aristoteles über Kleisthenes", font: "Cambria", bold: true, size: 34 })], spacing: { after: 140 } }),
  box(W, [
    p("**Ziel (Ich kann …):** " + "Ich kann eine Textquelle mit dem Schreibplan analysieren: Einleitung schreiben und Inhalt wiedergeben.", { size: s }),
    p("**Prüfaufgabe:** Abschlussaufgabe – in einem fehlerhaften Klausur-Anfang drei Fehler erkennen und begründen („Ich rate dir …, weil …“). Erfüllt, wenn mindestens zwei Ratschläge Einleitung (Autor nicht dabei, Angaben fehlen) und Inhaltswiedergabe (indirekte Rede, Zeilen, keine Meinung) treffen.", { size: s }),
    p("**Mindestziel:** Schritt 1 (Einleitung) und Schritt 2 (Inhalt) am Aristoteles-Text. Mit *: Schritte 3 und 4.", { size: s }),
    p("**Beitrag zur Reihe:** Die Doppelstunde ist die Methodenstunde zu Montag (Kleisthenes): kein neuer Stoff, sondern derselbe Inhalt aus einer Quelle. Der Schreibplan ist die Vorlage für die Analyseaufgabe in der Klausur am 25.11.", { size: s, after: 0 }),
  ]),
  p("**Bezug zur letzten Stunde:** Am Montag wurde die Phylenreform als Mystery erarbeitet. Heute wird derselbe Inhalt in einer Quelle wiedererkannt – das entlastet das Lesen, die Klasse kann sich auf die Methode konzentrieren. Die Zeit ist knapp geplant, deshalb sind die Schritte 3 und 4 mit * markiert; der Abschluss bleibt geschützt (Rückmeldung Peisistratos). Die Lösungen werden vom Platz vorgelesen, nicht an der Tafel erarbeitet.", { size: s, before: 100 }),
  p("**Hinweis:** Das hochgeladene Material zur Quellenanalyse lag nicht vor. Der Schreibplan folgt deshalb dem üblichen Vierschritt (Einleitung – Inhalt – Einordnen – Urteilen). Wenn die Fachschaft ein eigenes Schema oder eigene Operatoren nutzt, Folien 6–7 und Seite 2 des Blatts daran anpassen.", { size: s, color: C.muted }),

  h("Druckliste"),
  table(W, [5200, 2400, 2038], ["Material", "Format", "Menge"], [
    ["Test_Athen_bis_Peisistratos", "A4 doppelseitig", "1 je Schüler"],
    ["Quelle_und_Schreibplan", "A4 doppelseitig", "1 je Schüler"],
    ["Test_Erwartungshorizont_Notenschluessel", "A4", "nur Lehrkraft"],
    ["Folien_Quellenanalyse (26 Folien, Teil 1 = Folie 1)", "Beamer", "–"],
  ], { size: 18 }),

  h("Teil 1: Test (45 Min.)"),
  li("Folie 1 zeigen. Tische auseinander, Test verdeckt austeilen, gemeinsam beginnen. Endzeit an die Tafel, 5 Minuten vor Schluss ansagen."),
  li("37 Punkte + 1 Zusatzpunkt; Erwartungshorizont und Notenschlüssel (Notenpunkte) im eigenen Dokument."),
  li("Wer früher fertig ist, dreht das Blatt um und bleibt sitzen; nach dem Einsammeln direkt Blatt „Quelle und Schreibplan“ austeilen."),

  h("Begriffs-Check Teil 2"),
  table(W, [3000, 6638], ["Begriff", "Status"], [
    ["Quelle / Darstellung", "aus der Sek. I bekannt – in der Wiederholung (Folie 4) aktiviert"],
    ["Phyle, Demos, Trittys", "Montag gesichert"],
    ["Sinnabschnitt, indirekte Rede", "aus Deutsch bekannt – Prüffrage auf Folie 14"],
    ["Alkmeoniden, Isagoras, Bünde der Adligen", "**neu**, aber nur Randinformation – Worterklärung unter der Quelle"],
  ], { size: 18 }),

  h("Verlauf Teil 2 (45 Min.) nach AVIVA"),
  table(W, [700, 1500, 3400, 2338, 1000, 700], ["Min.", "Phase", "Ablauf und Ansage", "Warum?", "Form", "Folie"], VERLAUF, { size: 16 }),

  h("Erwartungshorizont: Schreibplan"),
  p("**Schritt 1 – Einleitung**", { size: s, keepNext: true }),
  table(W, [2200, 7438], null, D.einleitung.felder, { firstFill: true, size: 18 }),
  p("Einleitungssatz: „" + D.einleitung.satz + "“", { size: s, before: 60, italics: true }),
  p("**Schritt 2 – Inhalt** (andere Formulierungen richtig, wenn Kern, indirekte Rede und Zeilen stimmen)", { size: s, before: 100, keepNext: true }),
  table(W, [1200, 8438], ["Zeilen", "Lösung"], D.abschnitte.map(([z, , , l]) => [z, l]), { firstFill: true, size: 18 }),
  p("**Schritt 3 – Einordnen (nur bei Zeit):** Nach dem Sturz des Tyrannen Hippias (510) kämpften die Adelsfamilien wieder um die Macht. Kleisthenes setzte seine Reform 508/07 mit Hilfe des Volkes durch. Laut Aristoteles wollte er die Bürger „vermischen“, damit mehr Menschen am Staat teilhatten (Z. 6f.) – die Bindung an die Adelsfamilien sollte gebrochen werden (Montag).", { size: s, before: 100 }),
  p("**Schritt 4 – Urteilen (nur bei Zeit):** Aristoteles war kein Augenzeuge; er schrieb rund 180 Jahre später und stützte sich auf ältere Schriften. Er nennt Absichten („damit …“), die er selbst so deutet. Die Angaben zu Phylen, Rat und Trittyen passen aber zu anderen Überlieferungen – als Quelle für die Ordnung gut brauchbar, für die Absichten des Kleisthenes nur eingeschränkt.", { size: s }),

  h("Stolpersteine (Warnung vor dem Fehler) und gestufte Hilfsimpulse"),
  table(W, [2600, 7038], ["Stolperstein", "Was du sagst"], [
    ["„Aristoteles war Augenzeuge.“", "„Wann lebte er – und wann war die Reform? Rechne nach.“"],
    ["Entstehungszeit = 508/07", "„Das ist die Zeit der Reform. Wann wurde der Text geschrieben?“"],
    ["Thema: „Aristoteles“ oder „Athen“", "„Worüber schreibt er in diesem Ausschnitt genau?“"],
    ["Abschreiben statt eigene Worte", "„Deck die Quelle ab und sag es mir in einem Satz.“"],
    ["Indikativ („Kleisthenes teilte …“)", "„Wer sagt das? Mach deutlich, dass es Aristoteles sagt: … habe …“"],
    ["Meinung im Inhalt („klug“, „gut“)", "„Steht das in der Quelle? Heb es dir für Schritt 4 auf.“"],
    ["Z. 12–14 unklar (Neubürger)", "„Kleisthenes nahm neue Bürger auf, z. B. Zugezogene. Woran hätte man sie bisher erkannt?“ (am Vatersnamen)"],
  ], { size: 18 }),
  p("**Gestufte Hilfen beim Inhalt** (mündlich, der Reihe nach): 1. „Unterstreiche den wichtigsten Satz des Abschnitts.“ – 2. „Wer tut was? Achte auf ‚damit‘ – das ist der Zweck.“ – 3. Stichworte aus der Lösung nennen (Lösung selbst bleibt am Pult).", { size: s, before: 80 }),

  h("Abschlussaufgabe: mögliche Ratschläge"),
  li("Ich rate dir, die Lebensdaten zu prüfen, weil Aristoteles rund 180 Jahre später schrieb und nicht dabei war."),
  li("Ich rate dir, Quellenart, Entstehungszeit und Adressat zu nennen, weil die Einleitung sonst unvollständig ist."),
  li("Ich rate dir, indirekte Rede zu nutzen, weil du wiedergibst, was Aristoteles sagt."),
  li("Ich rate dir, Zeilen anzugeben, weil man deine Aussagen sonst nicht überprüfen kann."),
  li("Ich rate dir, „sehr gute Idee“ wegzulassen, weil die Meinung erst in Schritt 4 gehört."),
  li("Ich rate dir, nichts hinzuzufügen, weil „die Demokratie eingeführt“ so nicht in der Quelle steht."),

  h("Sachinformation zur Quelle"),
  p("Die „Athenaion politeia“ wurde um 330 v. Chr. im Umfeld des Aristoteles verfasst (ob von ihm selbst, ist umstritten). Der Text auf dem Blatt ist eine eigene, stark vereinfachte und gekürzte Übertragung von Kap. 20,1 und 21,1–4. Isagoras war 508/07 Archon und wurde von Sparta unterstützt; Kleisthenes gehörte zu den Alkmeoniden. Die „Bünde der Adligen“ sind die Hetairien (Adelsclubs). Mit den „neuen Bürgern“ (Z. 13f.) sind Zugezogene und andere gemeint, die Kleisthenes ins Bürgerrecht aufnahm; der Demenname (z. B. „aus Paiania“) verdeckte ihre Herkunft.", { size: s }),
];

const doc = new Document({ styles: { default: { document: { run: { font: "Calibri", size: 20, color: C.dark } } } },
  sections: [{ properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 900, bottom: 800, left: 1134, right: 1134 } } },
    headers: { default: new Header({ children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ text: "Geschichte · Klasse 11 · Test und Quellenanalyse – Handreichung", size: 16, color: "998C77" })] })] }) },
    children: content }] });
Packer.toBuffer(doc).then(b => { fs.writeFileSync(OUT, b); console.log("done"); });
