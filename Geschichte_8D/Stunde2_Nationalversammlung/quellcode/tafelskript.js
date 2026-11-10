const fs = require("fs");
const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, BorderStyle, ShadingType, AlignmentType, VerticalAlign, Header, Footer, PageNumber } = require("docx");

const NAVY = "1E2761", MUTED = "5B6B8C", LIGHT = "F4F7FD", BOARD = "F3F6F1", RED = "D9534F", DARK = "3D4A3F", GOLD = "FFF6E0";
const W = 10206;
const t = (text, o = {}) => new TextRun({ text, font: o.font || "Calibri", size: o.size || 22, bold: o.bold, italics: o.italics, color: o.color || "000000" });
const p = (runs, o = {}) => new Paragraph({ children: Array.isArray(runs) ? runs : [runs], alignment: o.align, spacing: { before: o.before ?? 0, after: o.after ?? 100 }, keepNext: o.keepNext, indent: o.indent });
const solid = (c, s = 8) => ({ style: BorderStyle.SINGLE, size: s, color: c });
const all = (b) => ({ top: b, bottom: b, left: b, right: b });
function cell(children, o = {}) {
  return new TableCell({ children: Array.isArray(children) ? children : [children], width: { size: o.w, type: WidthType.DXA }, shading: o.fill ? { fill: o.fill, type: ShadingType.CLEAR, color: "auto" } : undefined, borders: o.borders || all(solid("B9C6E8", 6)), verticalAlign: o.valign || VerticalAlign.CENTER, margins: { top: o.m ?? 70, bottom: o.m ?? 70, left: 110, right: 110 } });
}
const table = (widths, rows) => new Table({ width: { size: widths.reduce((a, b) => a + b, 0), type: WidthType.DXA }, columnWidths: widths, rows });
const row = (cells, h) => new TableRow({ children: cells, height: h ? { value: h, rule: "atLeast" } : undefined, cantSplit: true });
const th = (txt, w, fill = DARK, size = 19) => cell(p(t(txt, { bold: true, color: "FFFFFF", size }), { after: 0, align: AlignmentType.CENTER }), { w, fill, borders: all(solid(DARK, 8)) });
const tc = (txt, w, o = {}) => cell(p(t(txt, { size: o.size || 20, bold: o.bold }), { after: 0 }), { w, fill: o.fill });
const kicker = (s) => p(t(s.toUpperCase(), { size: 18, color: MUTED, bold: true }), { after: 20 });
const h1 = (s) => p(t(s, { font: "Cambria", size: 38, bold: true, color: NAVY }), { after: 60 });
const h3 = (s) => p(t(s, { font: "Cambria", size: 24, bold: true, color: NAVY }), { before: 160, after: 60, keepNext: true });
const S = (s, o = {}) => t(s, Object.assign({ size: 21 }, o));
const header = new Header({ children: [new Paragraph({ tabStops: [{ type: "right", position: W }], children: [t("Geschichte · Klasse 8D · Nationalversammlung", { size: 17, color: MUTED }), t("\tfür die Lehrkraft", { size: 17, color: MUTED })] })] });
const footer = new Footer({ children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ font: "Calibri", size: 16, color: MUTED, children: ["Tafelskript 8D · Seite ", PageNumber.CURRENT] })] })] });

const BW = 1700, CW = W - BW, TW = CW - 340;
function step(klasse, parts) {
  const ch = [];
  if (parts.say) ch.push(p([t("Ansage: ", { bold: true, size: 20, color: RED }), t("„" + parts.say + "“", { italics: true, size: 21 })], { after: 60 }));
  (parts.do || []).forEach((x) => ch.push(p(S(x), { after: 50 })));
  if (parts.extra) ch.push(...parts.extra);
  if (parts.sol) ch.push(p([t("Erwartung: ", { bold: true, size: 19, color: MUTED }), t(parts.sol, { size: 19, color: MUTED })], { after: 0 }));
  return table([BW, CW], [row([cell([p(t("Klasse", { size: 16, color: MUTED }), { after: 10 }), p(t(klasse, { bold: true, color: NAVY, size: 20 }), { after: 0 })], { w: BW, fill: LIGHT, valign: VerticalAlign.TOP }), cell(ch, { w: CW, valign: VerticalAlign.TOP })])]);
}
const gapP = () => p(t(""), { after: 30 });
const phase = (nr, title, min, tag) => p([t(`${nr}  `, { font: "Cambria", size: 28, bold: true, color: RED }), t(title, { font: "Cambria", size: 28, bold: true, color: NAVY }), t(`   ${min} Min.`, { size: 20, bold: true, color: "C77C12" }), ...(tag ? [t(`   ${tag}`, { size: 18, color: MUTED, italics: true })] : [])], { before: 180, after: 90, keepNext: true });
const boardBox = (title, children) => table([TW], [row([cell(p(t("TAFEL · " + title, { bold: true, color: "FFFFFF", size: 16 }), { after: 0 }), { w: TW, fill: DARK, borders: all(solid(DARK, 12)), m: 40 })]), row([cell(children, { w: TW, fill: BOARD, borders: all(solid(DARK, 12)), valign: VerticalAlign.TOP, m: 100 })])]);
const folieBox = (title, children) => table([TW], [row([cell(p(t("FOLIE · " + title, { bold: true, color: "FFFFFF", size: 16 }), { after: 0 }), { w: TW, fill: NAVY, borders: all(solid(NAVY, 12)), m: 40 })]), row([cell(children, { w: TW, fill: LIGHT, borders: all(solid(NAVY, 12)), valign: VerticalAlign.TOP, m: 100 })])]);
const bl = (s, o = {}) => p(t(s, { font: "Cambria", size: 20 }), Object.assign({ after: 30 }, o));
const note = (label, txt, fill = GOLD, col = "B03A36", border = "E8A33D") => table([W - 20], [row([cell(p([t(label + "  ", { bold: true, size: 19, color: col }), t(txt, { size: 19 })], { after: 0 }), { w: W - 20, fill, borders: all(solid(border, 8)) })])]);
const bullets = (arr) => arr.map((x) => p([S("• "), S(x)], { after: 40, indent: { left: 240 } }));
const sp = () => p(t(""), { after: 20 });
const warum = (why, gelenk) => table([W - 20], [row([cell([p([t("Warum  ", { bold: true, size: 19, color: NAVY }), t(why, { size: 19 })], { after: gelenk ? 40 : 0 }), ...(gelenk ? [p([t("Gelenk  ", { bold: true, size: 19, color: NAVY }), t(gelenk, { size: 19, italics: true })], { after: 0 })] : [])], { w: W - 20, fill: "EAF0FA", borders: all(solid("B9C6E8", 6)) })])]);

const d = [];
d.push(kicker("Tafelskript für die Lehrkraft · Kurzstunde (25 Min.) · 25 Schülerinnen und Schüler · eine Lehrkraft · Mittwoch 11.11.2026"), h1("Abstimmung beobachten und der Dritte Stand handelt"));
d.push(p(t("Ziel: Die Abstimmung aus Stunde 1, die nicht mehr geschafft wurde, wird jetzt erlebt: Die Klasse bildet drei Stände, stimmt einen Antrag offen ab, zählt dabei nur die Ja-Stimmen je Stand und rechnet danach selbst aus, wie die Abstimmung nach Ständen und nach Köpfen ausgeht. Danach erarbeitet sie mit dem Text aus der 8C, wie der Dritte Stand (Bauern und Bürger) die Nationalversammlung gründet, bis zum Zusammenziehen der Truppen vor Paris. Gearbeitet wird mit nur drei Methoden (Beobachtungszettel, Text mit Zeitstrahl, Rat an den Freund), weil die 8C am 06.11. an zu vielen Methoden gelitten hat.", { size: 21, color: MUTED }), { after: 100 }));

d.push(h3("Kopf der Planung"));
d.push(table([2600, 7606], [
  ["Ziel (Ich kann …)", "… erklären, wie die Stände den Dritten Stand (Bauern und Bürger) überstimmen können. … die Abstimmung nach Ständen beurteilen und das Urteil begründen. … am Zeitstrahl erklären, wie sich der Dritte Stand zur Nationalversammlung erklärte.*"],
  ["Prüfaufgabe: Woran erkenne ich es?", "Beobachtungszettel ausgefüllt und ausgerechnet (Ziel 1), Heft: Ergebnisse und Beurteilung der Regel mit einer Zahl (Ziel 2). Zeitstrahl bis „Anfang Juli: Soldaten um Paris“ (Ziel 3). Offene Abschlussfrage mündlich: Rat an den Freund mit Position, Sachargument und „weil“."],
  ["Mindestziel", "Ziel 1 und 2. Ziel 3 hat einen Stern. „Auswerten“ bleibt immer."],
  ["Beitrag zur Reihe", "Stunde 2 von 11 (Sequenz Französische Revolution). Sie schließt die Abstimmung aus Stunde 1 ab (Ursache des Streits) und führt zur Nationalversammlung und zum Ballhausschwur. Die Karikatur-Analyse und der Merksatz rücken in die Folgestunden (siehe Sequenz)."],
  ["Vorgehen", "Phase 1 direkt (Regel und Anträge vorgegeben, Zettel strukturiert die Beobachtung), Phase 2 und 3 indirekt (Text und Zeitstrahl selbst erarbeiten). Direkt oder indirekt wird je Phase entschieden."],
].map(([a, b]) => row([tc(a, 2600, { fill: LIGHT, bold: true }), tc(b, 7606)], 340))));

d.push(h3("Was aus der 8C und aus Stunde 1 übernommen und geändert wurde"));
d.push(...bullets([
  "Rückmeldung 8C (06.11.): Für 45 Minuten waren es zu viele Methoden. Deshalb hier für 25 Minuten nur drei: Abstimmung mit Beobachtungszettel, Text mit Zeitstrahl, Abschlussfrage. Die Abschlussfrage läuft mündlich. Gestrichen oder verschoben: Buchstabenrätsel (Stunde 1 gelaufen, Begriffe bleiben im Heft), Merkbild, Merksatz zum Abschreiben, Karikatur-Dreischritt, Ballhausschwur-Bild.",
  "Rückmeldung 8D (04.11.): Die Abstimmung wurde nicht mehr geschafft, der Zeitstrahl lief gut. Deshalb steht die Abstimmung jetzt vorn, als Wiederholung des Unerledigten. Die Folie „Das ist mir wichtig“ kommt kurz wieder (hat motiviert).",
  "Entscheidung der Lehrkraft: Ein Antrag (A Steuern), Antrag B als Zusatz (*). Der Stimmzettel wird zum Beobachtungszettel: vorab ankreuzen, offen abstimmen, nur die Ja-Zahlen je Stand eintragen, danach selbst ausrechnen. Grund: Das gleichzeitige Zählen nach Ständen und nach Köpfen wäre zu kompliziert. Der Text bleibt wie in der 8C, mit Sieyès-Quelle und fünf Daten.",
  "Historische Vereinfachung: Klerus und Adel stimmen im Spiel immer gleich (Vorrechte behalten). Das ist verkürzt (siehe Prüfhinweise).",
]));

d.push(h3("Zeitplan (25 Minuten)"));
d.push(table([800, 2300, 4300, 2806], [
  row([th("Min.", 800), th("AVIVA", 2300), th("Phase", 4300), th("Arbeitsform", 2806)], 320),
  ...[["3", "Ankommen", "0  Ankommen, Ziele, „Das ist mir wichtig“", "still, allein"], ["9", "Vorwissen", "1  Abstimmung: vorab ankreuzen, beobachten, auswerten", "allein, Stände, Zettel"], ["6", "Informieren", "2  Text mit Sieyès-Quelle", "allein, leise"], ["3", "Verarbeiten", "3  Zeitstrahl bis Anfang Juli", "allein, Kontrolle per Folie"], ["4", "Auswerten", "4  Rat an den Freund, Ziele prüfen, Ausblick", "allein, gezielte Abfrage"]].map((r) => row(r.map((v, i) => tc(v, [800, 2300, 4300, 2806][i], { bold: i < 3 })), 320)),
]));
d.push(note("ZEIT", "Die Stunde ist sehr knapp: Kein Puffer. Läuft die Abstimmung länger als 4 Minuten, entfallen die Beurteilung (Phase 1c) und die beiden ersten Zeitstrahl-Einträge. Phase 4 wird nicht gestrichen."));

d.push(h3("Lautstärke und Ansagen (wörtlich)"));
d.push(table([4700, 5506], [
  ["„Stifte liegen. Ihr hört nur zu.“", "Sie erklären, niemand schreibt."],
  ["„Jetzt arbeitet ihr allein. Die Aufgabe steht rechts.“", "Einzelarbeit, leise."],
  ["„Jetzt stimmt ihr offen ab. Die Wahlleitung führt, alle anderen beobachten und schreiben auf.“", "Abstimmung (Phase 1b)."],
  ["„Vergleicht und verbessert mit Grün.“", "Kontrolle im eigenen Heft bzw. Zettel."],
  ["„Es ist zu laut. Ab jetzt arbeitet ihr allein und schreibt.“", "Wenn es laut wird, keine Diskussion."],
].map(([a, b]) => row([tc(a, 4700, { fill: LIGHT, bold: true }), tc(b, 5506)], 340))));

d.push(h3("Material und Vorbereitung (Details: Materialliste)"));
d.push(...bullets([
  "Druckvorlage „Karten, Beobachtungszettel und Rollenkarten“: Seite 1 und 2: 24 Standeskarten (3 Klerus, 4 Adel, 17 Dritter Stand, davon 1 Reserve). Seite 3: Beobachtungszettel, vier pro A4 (Schnipsel), 7-mal drucken und schneiden (28 Zettel), tintensparend. Die Auswertung (drei Fragen) schreibt die Klasse ins Heft. Seite 4: zwei Rollenkarten (Wahlleitung mit Zählhilfe, Schriftführung mit Tafelrahmen), je einmal.",
  "Text mit Zeitstrahl: „Der Dritte Stand handelt“ (Arbeitsblatt aus der 8C, ein Blatt pro Person, 27 Stück). Lösung für Sie: Loesung_Arbeitsblatt_Der_Dritte_Stand_handelt.pdf. Die Abschlussfrage läuft mündlich, es gibt kein Schnipsel.",
  "Vorab: Karten verdeckt mit dem Zettel auf die Plätze legen. Klerus und Adel gezielt verteilen (zuverlässige Personen, die ihre Rolle ernst nehmen). Wahlleitung und Schriftführung vorher ansprechen (zwei zuverlässige, rechensichere Personen, nicht stimmberechtigt), Rollenkarte geben.",
  "Wenn Schülerinnen oder Schüler fehlen: Klerus bleibt 3, Adel 4, die Reservekarten gleichen den Dritten Stand aus.",
  "Seitentafel: Ziele und Ablauf (siehe Phase 0). Mitte der Tafel: Rahmen (Tabelle) für die Ergebnisse (siehe Phase 1b), Platz für den Zeitstrahl.",
  "Folien 1 bis 15 (Schrift mindestens 27 Punkt), Kreide, Grün zum Verbessern, Uhr oder Timer.",
]));

// Phase 0
d.push(phase("0", "Ankommen", 3, "AVIVA: Ankommen"));
d.push(warum("Die Klasse wird ruhig, kennt Ziel und Ablauf und weiß, dass wir gemeinsam arbeiten (direkt, von der Lehrkraft geführt).", "„Auf dem Zeitstrahl endete der Streit um die Abstimmung. Heute erlebt ihr diesen Streit selbst.“"));
d.push(sp());
d.push(step("kommt an, arbeitet leise", {
  say: "Stifte liegen. Ihr hört nur zu. Lasst die Karte verdeckt liegen.",
  do: [
    "Folie 1 steht beim Betreten da (Heft raus, Datum und Überschrift, Karte verdeckt). Folie 2 „Das ist mir wichtig“ (20 Sekunden, Wortlaut wie in Stunde 1, ruhig, ohne Diskussion), Folie 3 (Lernziele vorlesen, 30 Sekunden), Folie 4 (Ablauf, 30 Sekunden). Zusammen etwa 2 Minuten.",
    "Hinweis zum Stern: „Was wir eventuell nicht schaffen, hat einen Stern. Am Ende prüfen wir, ob wir die Ziele erreicht haben.“",
  ],
  extra: [boardBox("Seitentafel · Ziel und Ablauf", [
    bl("Ich kann erklären, wie die Stände den Dritten Stand überstimmen können.", { after: 20 }),
    bl("Ich kann die Abstimmung nach Ständen beurteilen und mein Urteil begründen.", { after: 20 }),
    bl("Ich kann am Zeitstrahl erklären, wie sich der Dritte Stand zur Nationalversammlung erklärte.*", { after: 60 }),
    bl("A Ankommen · V Abstimmung · I Text · V Zeitstrahl · A Rat an den Freund", { after: 0 }),
  ]), gapP()],
}));

// Phase 1
d.push(phase("1", "Abstimmung: vorab ankreuzen, zählen, ausrechnen", 9, "AVIVA: Vorwissen aktivieren"));
d.push(warum("Der in Stunde 1 nicht geschaffte Stoff wird nachgeholt: Die Klasse erlebt die Regel „jeder Stand eine Stimme“ selbst. Gezählt werden nur die Ja-Stimmen je Stand (einfach), das Ergebnis nach Ständen und nach Köpfen rechnet die Klasse danach selbst aus. So erkennt sie selbst: dieselben Stimmen, zwei Regeln, zwei Ergebnisse (direkt: Regel und Antrag vorgegeben, Auswertung geführt).", "„Der Dritte Stand (Bauern und Bürger) macht in dieser Abstimmung eine Erfahrung. Gleich seht ihr, welche.“"));
d.push(sp());
d.push(step("1a  arbeitet allein", {
  say: "Stifte liegen. Ihr hört nur zu. Das ist der Antrag. Jeder Stand hat eine Stimme. … Jetzt arbeitet ihr allein. Die Aufgabe steht rechts.",
  do: [
    "Folie 5 (Antrag, 1 Minute): A „Auch Adel und Klerus zahlen künftig Steuern.“ Klerus 3, Adel 4, Dritter Stand (Bauern und Bürger) 16 Personen, Wahlleitung und Schriftführung stimmen nicht mit. Regel: Jeder Stand hat eine Stimme. Antrag B (Stände abschaffen) ist Zusatz (*), nur wenn mehr als 2 Minuten übrig sind.",
    "Folie 6 (Ü1, 1 Minute): Karte umdrehen, leise lesen, auf dem Zettel Stand und eigene Stimme vorab ankreuzen. Klerus und Adel stimmen nach ihrer Rolle (Vorrechte behalten), der Dritte Stand nach seiner Überzeugung. Keine Absprache, keine Bewegung.",
  ],
}));
d.push(sp());
d.push(step("1b  beobachtet und schreibt", {
  say: "Stifte liegen. Jetzt stimmt ihr offen ab. Die Wahlleitung führt, die Schriftführung schreibt an die Tafel, alle anderen beobachten und tragen die Ja-Zahlen auf dem Zettel ein.",
  do: [
    "Folie 7 (Ü2, 3 Minuten): Die Wahlleitung ruft „Klerus. Wer ist für Antrag A? Hand hoch.“, zählt nur die Ja-Stimmen und ruft „Klerus: Ja …“ aus, dann Adel und Dritter Stand. Die Schriftführung schreibt die Zahl an die Tafel, alle anderen auf den Zettel. Nein-Stimmen und Ergebnisse werden jetzt nicht gezählt. Sie greifen nur im Notfall ein. Notfallplan: Dauert es länger als 3 Minuten, helfen Sie, die Zahlen stehen dann von Ihnen an der Tafel.",
    "Erwartet: Klerus Ja 0 von 3, Adel Ja 0 von 4, Dritter Stand Ja 16 von 16. Echte Zahlen gelten.",
  ],
  extra: [boardBox("Mitte · Rahmen (vorab anschreiben)", [
    bl("Stand | Personen | Ja", { after: 20 }),
    bl("Klerus | 3 | ", { after: 20 }), bl("Adel | 4 | ", { after: 20 }), bl("Dritter Stand | 16 | ", { after: 0 }),
  ]), gapP()],
}));
d.push(sp());
d.push(step("1c  rechnet und wertet aus", {
  say: "Stifte liegen. Jetzt arbeitet ihr allein. Die Aufgabe steht rechts. Rechnet erst auf dem Zettel, dann schreibt ihr ins Heft.",
  do: [
    "Folie 8 (Ü3, 3 Minuten): (1) Auf dem Zettel ausrechnen, etwa 1 Minute: Nach Ständen: Wie viele der drei Stände sagen Ja? Mindestens 2 von 3 = angenommen. Nach Köpfen: Alle Ja zusammenzählen (von allen 23), mehr als die Hälfte = angenommen. (2) Ins Heft, etwa 2 Minuten: „Nach Ständen …, nach Köpfen …“ und die Beurteilung „Ich finde die Regel …, weil … (nenne eine Zahl)“. Die Beurteilung gibt kein Urteil vor, mehrere sind möglich. Wenn die Zeit knapp ist, genügen die Ergebnisse und ein Urteilssatz. Heft-Stichprobe: bei 4 bis 6 Personen unterschiedlichen Niveaus mitlesen, nicht kommentieren (unbenotet).",
    "Folie 9 (Kontrolle, 1 Minute): „Vergleicht und verbessert mit Grün.“ Nach Ständen Klerus Nein, Adel Nein, Dritter Stand Ja = 1 von 3, abgelehnt. Nach Köpfen 16 von 23 Ja, angenommen. Einordnung in einem Satz: „Im Spiel stimmen Klerus und Adel immer gleich. In Frankreich sind sie zusammen nur etwa 3 % der Menschen.“ Gezielte Abfrage nur bei Bedarf (2 bis 3 Personen, kurz): „Wie haben Klerus und Adel abgestimmt? Warum?“",
  ],
  sol: "Rechnen: Nach Ständen 1 von 3 Ständen Ja, abgelehnt. Nach Köpfen 16 von 23 Ja, angenommen. Heft: Nach Ständen wurde der Antrag abgelehnt, nach Köpfen angenommen, weil jeder Stand nur eine Stimme hat und Klerus und Adel gleich abstimmten. Urteil offen, mit Zahl: z. B. „Ich finde die Regel ungerecht, weil der Dritte Stand (Bauern und Bürger) ca. 97 % der Menschen stellt, aber nur 1 von 3 Stimmen hat“ oder „Ich finde sie nachvollziehbar, weil jeder Stand als Gruppe zählt, aber die Zahl 97 % zeigt ein Problem“.",
}));
d.push(sp());
d.push(note("SOLLBRUCHSTELLE", "Wird die Abstimmung länger als 3 Minuten, schreiben Sie die Ja-Zahlen selbst an. Wird das Ausrechnen und Schreiben länger als 3 Minuten, entfällt die Beurteilung (Satz „Ich finde die Regel …“) und die Kontrolle läuft in 30 Sekunden (nur Folie 9). Sie kann in Stunde 3 als Wiederholung wieder aufgegriffen werden."));

// Phase 2
d.push(phase("2", "Text mit Sieyès-Quelle", 6, "AVIVA: Informieren"));
d.push(warum("Wie der Dritte Stand die Macht des Königs infrage stellt, wird selbst aus dem bearbeiteten Text erarbeitet (indirekt: Lesen allein, Unterstreichen, Heft-Stichprobe). Kein Vorlesen.", "„Der Dritte Stand hat gegen die Regel verloren, aber er gibt nicht auf. Wie reagiert er? Das lest ihr jetzt.“"));
d.push(sp());
d.push(step("liest allein", {
  say: "Stifte liegen. Ihr hört nur zu. Das Blatt liegt vor euch. Jetzt arbeitet ihr allein. Die Aufgabe steht rechts.",
  do: [
    "Folie 10 (Ü4, 6 Minuten): Text lesen, alle Daten unterstreichen, die Quelle von Sieyès lesen (Was verlangt der Dritte Stand?). Der Text ist unverändert wie in der 8C: Eröffnung der Generalstände am 5. Mai 1789, Sieyès, Nationalversammlung am 17. Juni, Ballhausschwur am 20. Juni, Soldaten um Paris Anfang Juli.",
    "Warnung vor dem Fehler: Das Lehrbuch nennt den 20. Juni für die Nationalversammlung. Historisch ist es der 17. Juni (Nationalversammlung) und der 20. Juni (Ballhausschwur). Sagen Sie das vorab in einem Satz.",
    "Heft-Stichprobe beim Durchgehen (unterstrichene Daten). Wer schneller fertig ist, überlegt, wie der Dritte Stand die Macht des Königs infrage stellt.",
  ],
}));

// Phase 3
d.push(phase("3", "Zeitstrahl bis Anfang Juli", 3, "AVIVA: Verarbeiten"));
d.push(warum("Die Ereignisse werden auf dem Zeitstrahl in die richtige Reihenfolge gebracht (angeleitet: Daten vorgegeben, Ereignisse selbst eintragen). Der letzte Eintrag führt zur nächsten Stunde.", "„Am Ende des Zeitstrahls zieht der König Soldaten um Paris zusammen. Was meint ihr, wie die Menschen in Paris reagieren?“ (nicht beantworten lassen, nur denken)"));
d.push(sp());
d.push(step("arbeitet allein", {
  say: "Stifte liegen. Jetzt arbeitet ihr allein. Die Aufgabe steht rechts. … Vergleicht und verbessert mit Grün.",
  do: [
    "Folie 11 (Ü5, 2 Minuten): Zu jedem Datum das Ereignis in Stichworten eintragen, die Daten sind vorgegeben. Danach Folie 12 (Lösung, 1 Minute): „Vergleicht und verbessert mit Grün.“ Aufgabe 3 auf dem Blatt (Sieyès und die Zahl) entfällt oder ist Zusatz (*).",
    "Wenn die Zeit knapp ist: nur die drei letzten Einträge (17. Juni, 20. Juni, Anfang Juli) ausfüllen.",
  ],
  extra: [boardBox("Mitte · Zeitstrahl (Lösung, Folie 12)", [
    bl("Winter 1788/89    Sieyès: Der Dritte Stand ist ALLES"),
    bl("5. Mai 1789    Eröffnung der Generalstände, Streit um die Abstimmung"),
    bl("17. Juni 1789    Nationalversammlung"),
    bl("20. Juni 1789    Ballhausschwur"),
    bl("Anfang Juli 1789    Soldaten um Paris, Angst in Paris", { after: 0 }),
  ]), gapP()],
}));

// Phase 4
d.push(phase("4", "Auswerten", 4, "AVIVA: Auswerten"));
d.push(warum("Der Lernzuwachs wird überprüft: offene Abschlussfrage mit mehreren guten Antworten, die ein Urteil mit Sachargument verlangt (Vorstufe zu den Urteilsaufgaben der Reihe), dann Ziele prüfen. Diese Phase wird nie gestrichen.", "„Nächste Stunde: Der König zieht Soldaten zusammen, was passiert in Paris?“"));
d.push(sp());
d.push(step("arbeitet allein", {
  say: "Stifte liegen. Ihr hört nur zu. Das ist die letzte Aufgabe. Jetzt arbeitet ihr allein. Die Aufgabe steht rechts.",
  do: [
    "Folie 13 (Abschlussfrage, 2 Minuten): Zeichnung („Dein Freund“, Bauer, mit Sprechblase, „Du“ erklärst): „Dein Freund ist auch Bauer und sagt: Die Abgeordneten des Dritten Standes (Bauern und Bürger) wollen eine Verfassung. Soll ich mich ihnen anschließen? Ich habe doch genug mit der Ernte zu tun. Erkläre ihm, was du ihm rätst.“ Die Antwort kommt mündlich, kein Schnipsel, kein Schreiben: „Ich rate dir, …, weil …“. Mehrere gute Antworten (anschließen oder abwarten), geprüft werden Position, ein Sachargument aus der Stunde und ein „weil“. Erst zehn Sekunden denken, dann mündlich.",
    "Form nach Verlauf wählen: (a) Blitzlicht (empfohlen): 4 bis 6 Personen unterschiedlichen Niveaus gezielt aufrufen, je eine Stimme „anschließen“ und „abwarten“, kein Kommentar. (b) Rat-Duell zu zweit, leise („Flüsterstufe. Nur dein Partner hört dich.“), danach zwei Personen berichten (nur wenn Zeit ist). Unbenotet.",
    "Folie 14 (30 Sekunden, still): Haben wir die Ziele erreicht? Keine Antworten ins Plenum. Folie 15 (30 Sekunden, nur lesen): Ausblick, Soldaten um Paris. „Wie reagieren die Menschen in Paris? Das sehen wir in der nächsten Stunde.“",
  ],
  sol: "Mögliche Argumente (Auswahl, keine Musterlösung). Anschließen: Der Dritte Stand (Bauern und Bürger) ist fast die ganze Bevölkerung (Sieyès: ALLES); die Bauern tragen die Steuern und Abgaben, Klerus und Adel kaum; gemeinsam sind sie stark; mit einer Verfassung entscheidet der König nicht mehr allein; wer abseits steht, hat keine Stimme. Abwarten: Ernte und Familie brauchen Kraft und Zeit; unklar, ob der König nachgibt; Veränderungen dauern; erst beobachten. Zwischenlösung ist begründbar. Kriterien: Position, ein Sachargument aus der Stunde, ein „weil“.",
}));

d.push(h3("So sieht das Heft am Ende der Stunde aus"));
d.push(table([2300, 7906], [
  ["Überschrift", "Nationalversammlung 1789 · Datum"],
  ["Beobachtungszettel", "eingeklebt: eigener Stand und eigene Stimme, Ja-Zahlen je Stand, Ergebnis nach Ständen und nach Köpfen. Im Heft: Ergebnisse und Beurteilung der Regel, Verbesserungen grün"],
  ["Text und Zeitstrahl", "Blatt eingeklebt, Daten unterstrichen, Ereignisse eingetragen, Verbesserungen grün"],
].map(([a, b]) => row([tc(a, 2300, { bold: true, fill: LIGHT }), tc(b, 7906)], 340))));

d.push(h3("Nach der Stunde festhalten (Reflexionslog)"));
d.push(...bullets(["Passten Ziel, Inhalt und Methode zusammen? Wie lief die Zeit (vor allem Phase 1b)? Wie viel echte Lernzeit hatte die Klasse?", "Haben Wahlleitung und Schriftführung die Abstimmung getragen? Musste der Notfallplan greifen?", "Wo wurde es laut, und welche Ansage half?", "Wie viele nannten in der Abschlussfrage ein Sachargument aus der Stunde (Stichprobe im Heft)? Welche Form (Blitzlicht, schriftlich, Duell) hat funktioniert?", "Wurde der Zeitstrahl bis Anfang Juli erreicht? Was muss in Stunde 3 wiederholt werden?"]));
d.push(h3("Zum Prüfen"));
d.push(...bullets([
  "Das Spiel vereinfacht: Im Spiel stimmen Klerus und Adel immer gleich. Historisch waren die Stände nicht immer einheitlich (viele einfache Pfarrer unterstützten den Dritten Stand, einzelne Adlige schlossen sich später an), bei der Abstimmungsregel hielten die Privilegierten aber oft zusammen. Die Anträge A und B sind Spielanträge, historisch ging der Streit um die Abstimmungsregel (1789 wurde nicht nach Köpfen abgestimmt).",
  "Die Standeskarten verteilen die Rollen: 3 Klerus (13 %), 4 Adel (17 %), 16 Dritter Stand (70 %), dazu Wahlleitung und Schriftführung. Das ist verzerrt, in Frankreich sind Klerus und Adel zusammen etwa 3 % (Klerus ca. 150 000, Adel ca. 500 000, Dritter Stand ca. 20 000 000 Menschen, Zahlen aus Ihrem Gruppenpuzzle-Material).",
  "Zahlen und Daten stammen aus dem Text der 8C und sind gegen das Schulbuch zu lesen (Lehrbuch S. 112/113): 17. Juni Nationalversammlung, 20. Juni Ballhausschwur, Anfang Juli Soldaten und Entlassung des Finanzministers.",
  "Die Figuren („Dein Freund“, „Du“) sind selbst gezeichnet, keine historischen Porträts.",
]));

(async () => {
  const doc = new Document({ styles: { default: { document: { run: { font: "Calibri", size: 22 } } } }, sections: [{ properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 1000, bottom: 850, left: 850, right: 850 } } }, headers: { default: header }, footers: { default: footer }, children: d }] });
  fs.writeFileSync("../Tafelskript_Stunde2_Nationalversammlung_8D.docx", await Packer.toBuffer(doc));
  console.log("ok");
})();
