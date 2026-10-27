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
const header = new Header({ children: [new Paragraph({ tabStops: [{ type: "right", position: W }], children: [t("Geschichte · Klasse 8 · Generalstände, Stunde 1", { size: 17, color: MUTED }), t("\tfür die Lehrkraft", { size: 17, color: MUTED })] })] });
const footer = new Footer({ children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ font: "Calibri", size: 16, color: MUTED, children: ["Tafelskript · Seite ", PageNumber.CURRENT] })] })] });

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
const sollbruch = (txt) => table([TW], [row([cell(p([t("SOLLBRUCHSTELLE  ", { bold: true, size: 19, color: "B03A36" }), t(txt, { size: 19 })], { after: 0 }), { w: TW, fill: GOLD, borders: all(solid("E8A33D", 8)) })])]);

const d = [];
d.push(kicker("Tafelskript für die Lehrkraft · Einzelstunde (45 Min.) · 12 Schülerinnen und Schüler · zwei Lehrkräfte"), h1("Generalstände, Stunde 1: Wie unfair ist die Abstimmung?"));
d.push(p(t("Ziel: Die Klasse bekommt den Test zurück und reflektiert ihre Vorbereitung. Sie erarbeitet den Verlauf bis zu den Generalständen als Zeitleiste und rechnet zuerst in einem Gedankenspiel zu 1789 nach und erlebt dann in einer Mini-Abstimmung in der Schule, wie unfair die Abstimmung nach Ständen ist. Aufbau nach AVIVA. Es gibt keinen Hörtext, der Text wird gelesen. Rollen der beiden Lehrkräfte: L1 moderiert, L2 spielt in der Abstimmung den Adel und unterstützt an der Tafel.", { size: 21 }), { after: 80 }));

d.push(table([800, 2300, 4300, 2806], [
  row([th("Min.", 800), th("AVIVA", 2300), th("Phase", 4300), th("Arbeitsform", 2806)], 320),
  ...[["8", "Ankommen", "0  Test zurück, Reflexion", "gemeinsam, mündlich"], ["3", "Vorwissen", "1  Probleme Frankreichs sammeln", "gemeinsam, mündlich"], ["12", "Informieren", "2  Text lesen, Zeitstrahl ausfüllen, kontrollieren", "allein, gemeinsam"], ["14", "Verarbeiten", "3  Zahlen, Gedankenspiel 1789, Schulbeispiel, Rechnung, Merksatz", "gemeinsam, abschreiben"], ["5", "Auswerten", "4  Satz „unfair, weil …“, Ausblick", "allein, mündlich"], ["3", "—", "Puffer (Abschreiben der Zeitleiste entfällt)", "—"]].map((r) => row(r.map((v, i) => tc(v, [800, 2300, 4300, 2806][i], { bold: i < 3 })), 320)),
]));

d.push(h3("Eindeutige Ansagen (wörtlich)"));
d.push(table([4700, 5506], [
  ["„Stifte liegen. Ihr hört nur zu.“", "Sie erklären oder fragen, niemand schreibt."],
  ["„Meldet euch. Ihr schreibt noch nichts auf.“", "Gemeinsames Gespräch."],
  ["„Jetzt schreibt ihr ab: …“", "Zeitleiste oder Merksatz ins Heft."],
  ["„Jetzt arbeitet ihr allein. Die Aufgabe steht rechts.“", "Einzelarbeit."],
  ["„Wer drankommt, schreibt an die Tafel. Alle prüfen mit.“", "Zeitleiste oder Ergebnis an der Tafel."],
  ["„Vergleicht und verbessert mit Grün.“", "Kontrolle im eigenen Heft."],
].map(([a, b]) => row([tc(a, 4700, { fill: LIGHT, bold: true }), tc(b, 5506)], 340))));
d.push(p(S("Material: Text „Wie kam es zu den Generalständen?“ mit Zeitstrahl unten auf demselben Blatt (ein Blatt pro Person, Daten vorgegeben, Ereignisfelder leer), Tests (verdeckt), Folien 1 bis 16 (mit Lernzielen, Ablauf nach AVIVA und Visualisierungen), Grün zum Verbessern. Vorab an der Tafel vorbereiten: Strichliste der Noten (ohne Namen) und die Erfolgsquote je Test-Aufgabe in Prozent.", { italics: true, color: MUTED }), { before: 80, after: 40 }));

// Phase 0
d.push(phase("0", "Test zurück, Reflexion", 8, "AVIVA: Ankommen"));
d.push(step("hört zu, meldet sich", {
  say: "Stifte liegen. Ihr bekommt jetzt euren Test zurück. Schaut ihn noch nicht an.",
  do: [
    "1  Folien 1 bis 4 (Titel, Lernziele, Ablauf, Test): Lernziele vorlesen lassen, Ablauf nach AVIVA zeigen. Tests verdeckt austeilen, niemand nennt Noten laut.",
    "2  Ihre Rückmeldung (Vorschlag): „Die Ergebnisse zeigen zwei Hälften. Eine Hälfte hat erkennbar gelernt, und das zahlt sich aus. Die andere Hälfte hat offensichtlich nicht oder kaum gelernt. Der Stoff war derselbe, die Merksätze und der Lückentext lagen bei allen im Heft. Das ist keine Frage von Begabung, sondern von Vorbereitung.“",
    "3  Strichliste der Noten (ohne Namen) und die Erfolgsquote je Aufgabe an der Tafel zeigen. Fragen: Welche Aufgabe lief am besten, welche am schlechtesten?",
    "4  „Meldet euch. Ihr schreibt noch nichts auf. Was haben die gemacht, die gut abgeschnitten haben? Was macht ihr beim nächsten Mal anders?“ Drei Stichworte an die Tafel sammeln (z. B. Merksätze wiederholen, Lückentext ohne Wortspeicher üben, Begriffe abfragen).",
    "5  „Vergleicht und verbessert mit Grün. Die falschen Aufgaben berichtigt ihr bis zur nächsten Stunde.“",
  ],
  sol: "Hinweis: keine Namen, keine Noten laut. Die zweite Lehrkraft kann die Strichliste vorab auszählen. Die Berichtigung ist Hausaufgabe.",
}));

// Phase 1
d.push(phase("1", "Probleme Frankreichs sammeln", 3, "AVIVA: Vorwissen aktivieren"));
d.push(step("meldet sich", {
  say: "Meldet euch. Ihr schreibt noch nichts auf. Welche Probleme hatte Frankreich?",
  do: ["Folie 5 (Probleme Frankreichs als Kacheln). Stichworte aus der Münzgeschichte an die Tafel (bleiben stehen): leere Staatskasse · ungleiche Steuerlast · Hunger und Brotpreis · keine Mitsprache des Dritten Standes (Bauern und Bürger) · Entfremdung zwischen Adel und Dorf.", "Leitfragen an die Tafel (bleiben stehen): „Wie kommt es zu den Generalständen? Wie fair ist die Abstimmung?“"],
}));

// Phase 2
d.push(phase("2", "Text lesen, Zeitstrahl ausfüllen", 12, "AVIVA: Informieren"));
d.push(step("arbeitet allein, dann gemeinsam", {
  say: "Ich teile den Text aus. Jetzt arbeitet ihr allein. Die Aufgabe steht rechts.",
  do: ["Folie 1. Folie 6. Blatt austeilen (oben der Text, unten der Zeitstrahl mit den Daten, aber leeren Ereignisfeldern). 5 Minuten lesen, Jahreszahlen unterstreichen, zu jedem Datum das Ereignis als Stichwort eintragen."],
  extra: [folieBox("6 · Ü1 Text und Zeitstrahl", [p(t("Ü1 Wie kam es zu den Generalständen?", { font: "Cambria", size: 22, bold: true, color: NAVY }), { after: 40 }), bl("1. Lies den Text und unterstreiche die Jahreszahlen."), bl("2. Schreibe zu jedem Datum das Ereignis in den Zeitstrahl."), bl("3. Prüfe: Passt jedes Ereignis zum Datum?", { after: 40 }), p([t("Zeit: ", { bold: true, size: 19 }), t("5 Minuten · allein · leise     ", { size: 19 }), t("Fertig? Stift hinlegen.", { size: 19, italics: true })], { after: 0 })]), gapP()],
}));
d.push(p(t(""), { after: 20 }));
d.push(step("kommt an die Tafel", {
  say: "Wer drankommt, schreibt an die Tafel. Alle prüfen mit.",
  do: ["Folie 7 (Zeitstrahl als Bild) erst zeigen, wenn die Klasse die Ereignisse genannt hat. Kontrolle des Zeitstrahls, 5 Minuten. Pro Datum nennt ein Schüler oder eine Schülerin das Ereignis, ein Schüler oder eine Schülerin schreibt es an die Tafel (Datum vorgegeben). Leitfrage danach:"],
  extra: [boardBox("Mitte · Zeitstrahl (Lösung)", [
    bl("1614    letzte Generalstände (danach 175 Jahre nicht mehr)"),
    bl("1778–1783    Krieg in Amerika, hohe Schulden"),
    bl("1787    Plan „Steuern auch für Adel und Klerus“ scheitert"),
    bl("1788    Missernte, Brotpreis steigt, Staatsbankrott droht"),
    bl("August 1788    König kündigt Generalstände an"),
    bl("Frühjahr 1789    Wahlen, Beschwerdehefte"),
    bl("5. Mai 1789    Eröffnung in Versailles", { after: 0 }),
  ]), gapP()],
  sol: "Leitfrage: „Warum ruft der König die Generalstände ein, obwohl er sie 175 Jahre nicht brauchte?“ Antwort: Er braucht Geld, die Privilegierten blockieren Reformen, dazu kommen Hunger und drohender Bankrott. Danach „Vergleicht und verbessert mit Grün.“ Das Abschreiben entfällt, weil die Schüler die Ereignisse schon im Zeitstrahl notiert haben. Die gewonnene Zeit (2 Minuten) geht in den Puffer.",
}));

// Phase 3
d.push(phase("3", "Abstimmung: erst Gedankenspiel 1789, dann die Schule", 14, "AVIVA: Verarbeiten"));
d.push(step("meldet sich", {
  say: "Meldet euch. Ihr schreibt noch nichts auf. Was fällt euch an diesen Zahlen auf?",
  do: ["3a  Folie 8 zeigen (3 Minuten, Balken). Die Abstimmungsregel (früher eine Stimme pro Stand, der Dritte Stand fordert eine Stimme pro Abgeordnetem) steht im letzten Absatz des Textes. Zahlen vergleichen: Wie viele Menschen, wie viele Abgeordnete, wie viele Stimmen?"],
  extra: [folieBox("8 · Wer hat wie viel Gewicht?", [table([2900, 1500, 1500, 2326], [
    row([th("", 2900, NAVY), th("Klerus", 1500, NAVY), th("Adel", 1500, NAVY), th("Dritter Stand (Bauern und Bürger)", 2326, NAVY)], 340),
    ...[["Menschen", "150 000", "500 000", "20 000 000"], ["Anteil", "ca. 1 %", "ca. 2 %", "ca. 97 %"], ["Abgeordnete", "ca. 290", "ca. 270", "ca. 580"], ["Stimmen bei der Abstimmung nach Ständen", "1", "1", "1"]].map((r) => row(r.map((v, i) => tc(v, [2900, 1500, 1500, 2326][i], { bold: i === 0, fill: i === 0 ? LIGHT : undefined, size: 19 })), 360)),
  ])]), gapP()],
  sol: "Erwartet: Der Dritte Stand (Bauern und Bürger) ist fast die ganze Bevölkerung und stellt etwa die Hälfte der Abgeordneten, hat aber nur eine von drei Stimmen. Klerus und Adel zusammen (ca. 3 %) haben zwei von drei Stimmen und können ihn immer überstimmen.",
}));
d.push(p(t(""), { after: 20 }));
d.push(step("rechnet mit, meldet sich", {
  say: "Meldet euch. Ihr schreibt noch nichts auf. Stellt euch vor, 1789 stimmt man über diesen Antrag ab.",
  do: [
    "3b  Gedankenspiel zu 1789 (3 Minuten). Wichtig vorab: Nach Köpfen wurde 1789 nicht abgestimmt. Wir rechnen, was wäre, wenn. Folie 9: Antrag „Auch Adel und Klerus zahlen künftig Steuern.“ Ergebnis auf Folie 10. Modellannahme: Klerus und Adel stimmen geschlossen mit Nein, der Dritte Stand mit Ja.",
    "Schritt 1, Abstimmung nach Ständen: Die Klasse rechnet. Klerus Nein, Adel Nein, Dritter Stand Ja. Ergebnis 2:1, der Antrag ist abgelehnt. Ein Schüler schreibt das Ergebnis an die Tafel.",
    "Schritt 2, Abstimmung nach Köpfen: Ja 580, Nein 290 + 270 = 560. Der Antrag wäre knapp angenommen. Die Klasse bemerkt: Es entscheidet also der Abstimmungsmodus, nicht die Mehrheit der Menschen.",
    "Frage: „Wie wichtig wäre die Abstimmungsregel für den Dritten Stand?“ Dann klarstellen: Tatsächlich wurde nicht nach Köpfen abgestimmt. Adel und Klerus bestanden auf der Abstimmung nach Ständen, der Dritte Stand (Bauern und Bürger) forderte die nach Köpfen. Dieser Streit blockierte die Generalstände (Fortsetzung nach den Ferien).",
  ],
  extra: [boardBox("Mitte · Gedankenspiel 1789 (Modell)", [
    bl("Gedankenspiel (keine echte Abstimmung):", { after: 40 }),
    bl("nach Ständen:   Klerus Nein · Adel Nein · Dritter Stand Ja   →   2 : 1   →   abgelehnt"),
    bl("nach Köpfen:    Ja 580 · Nein 290 + 270 = 560   →   knapp angenommen", { after: 0 }),
  ]), gapP()],
  sol: "Hinweis: Die Annahme „geschlossen“ ist vereinfacht. Tatsächlich unterstützten manche Geistliche (vor allem Pfarrer) und einzelne Adlige den Dritten Stand, nach Köpfen wäre das Ergebnis also deutlicher. Wichtig bleibt: Dieselben Stimmen führen je nach Regel zu zwei verschiedenen Ergebnissen.",
}));
d.push(p(t(""), { after: 20 }));
d.push(step("spielt mit, meldet sich", {
  say: "Jetzt spielen wir dasselbe in der Schule. Das ist nur ein Spiel, es gilt nicht. Die Lehrkräfte sind Klerus und Adel, ihr seid der Dritte Stand (Bauern und Bürger).",
  do: [
    "3c  Schulbeispiel (4 Minuten). Folie 11: Antrag „Nur Schüler putzen künftig das Klassenzimmer und räumen die Tische auf. Lehrkräfte sind davon befreit.“",
    "Schritt 1, Abstimmung nach Ständen: Jeder Stand eine Stimme. Die Schüler beraten 30 Sekunden und stimmen als Gruppe ab. L1 (Klerus) Ja, L2 (Adel) Ja, Dritter Stand Nein. Ergebnis 2:1, angenommen. Ein Schüler schreibt es an die Tafel.",
    "Reaktion der Klasse abwarten: „Wir sind doch viel mehr!“ Dann Schritt 2, Abstimmung nach Köpfen: 2 Ja, 12 Nein, abgelehnt. Ergebnis auf Folie 12.",
    "Fragen: Wie fühlt sich das an? Was war unfair? Sind die Lehrkräfte das Problem oder die Regeln?",
  ],
  sol: "Die Regeln sind das Problem, nicht die Personen. Betonen Sie die Rollen („Wir spielen“), damit es nicht persönlich wird. Unterschied zu 1789: In der Schule ist das Ergebnis deutlich (12 gegen 2), 1789 knapp (580 gegen 560), weil der Dritte Stand mehr Abgeordnete bekam.",
}));
d.push(p(t(""), { after: 20 }));
d.push(sollbruch("Wenn die Zeit knapp wird, entfällt 3c (das Schulbeispiel) und beginnt Stunde 2 (4 Minuten). Das Gedankenspiel 3b und der Merksatz bleiben in Stunde 1. Phase 4 wird dann mündlich."));
d.push(p(t(""), { after: 20 }));
d.push(step("rechnet mit, schreibt ab", {
  say: "Meldet euch. Ihr schreibt noch nichts auf. Wie viel zählt eine Stimme?",
  do: [
    "3d  Rechnung und Merksatz (4 Minuten, Folie 13: Gewicht einer Stimme als maßstabsgetreue Balken): Im Klassenspiel zählt eine Lehrerstimme wie 1 Person, die Schülerstimme nur wie 1 von 12. Die Lehrerstimme wiegt zwölfmal so viel.",
    "Übertragung: In Frankreich stehen hinter einer Stimme des Klerus 150 000 Menschen, hinter einer Stimme des Dritten Standes (Bauern und Bürger) 20 000 000. 20 000 000 : 150 000 ≈ 133. Beim Adel: 20 000 000 : 500 000 = 40.",
    "Frage: „Was würdet ihr als Dritter Stand in Frankreich fordern?“ Erwartung: Abstimmung nach Köpfen. Das ist genau, was die Abgeordneten 1789 fordern (Brücke zu Stunde 2).",
    "„Jetzt schreibt ihr von der Folie ab: Merksatz 1. Rahmt ihn rot ein.“ (Folie 14)",
  ],
  extra: [folieBox("14 · Merksatz 1 – abschreiben, rot umrahmen", [
    p(t("Abstimmung nach Ständen", { bold: true, color: RED, font: "Cambria", size: 22 }), { after: 40 }),
    p(t("Bei den Generalständen hat jeder Stand eine Stimme. Klerus und Adel (zusammen ca. 3 % der Bevölkerung) haben damit zwei von drei Stimmen und können den Dritten Stand (Bauern und Bürger, ca. 97 %) immer überstimmen.", { font: "Cambria", size: 22 }), { after: 50 }),
    p(t("Beispiel: Eine Stimme des Klerus wiegt rund 133-mal so viel wie eine Stimme des Dritten Standes (Bauern und Bürger).", { size: 20, color: MUTED }), { after: 50 }),
    p(t("Test: Wie viele Menschen stehen hinter einer Stimme?", { size: 20, color: MUTED }), { after: 0 }),
  ]), gapP()],
}));

// Phase 4
d.push(phase("4", "Auswerten", 5, "AVIVA: Auswerten"));
d.push(step("arbeitet allein", {
  say: "Jetzt arbeitet ihr allein. Die Aufgabe steht rechts.",
  extra: [folieBox("15 · Ü2 Unfair?", [p(t("Ü2 Die Abstimmung nach Ständen", { font: "Cambria", size: 22, bold: true, color: NAVY }), { after: 40 }), bl("1. Schreibe Ü2 ins Heft."), bl("2. Beende den Satz: „Die Abstimmung nach Ständen ist unfair, weil …“"), bl("3. Begründe mit mindestens einer Zahl.", { after: 40 }), p([t("Zeit: ", { bold: true, size: 19 }), t("4 Minuten · allein · leise     ", { size: 19 }), t("Fertig? Stift hinlegen.", { size: 19, italics: true })], { after: 0 })]), gapP()],
  do: ["Folie 15. Danach 1 bis 2 Sätze vorlesen lassen. Folie 16: Lernziele prüfen (still), Ausblick: „Was passiert, wenn der Dritte Stand nach Köpfen fordert und der König ablehnt? Das sehen wir in der nächsten Stunde.“"],
  sol: "Beispiel: „Die Abstimmung nach Ständen ist unfair, weil der Dritte Stand (Bauern und Bürger) ca. 97 % der Menschen stellt, aber nur eine von drei Stimmen hat und immer überstimmt wird.“",
}));

d.push(h3("So sieht das Heft am Ende der Stunde aus"));
d.push(table([2300, 7906], [
  ["Überschrift", "Generalstände 1789 (mit Lineal unterstrichen) · Datum rechts"],
  ["Berichtigung", "Falsche Testaufgaben mit Grün verbessert (bis zur nächsten Stunde)"],
  ["Zeitstrahl", "Blatt eingeklebt, Ereignisse zu den Daten eingetragen (1614 bis 5. Mai 1789), Verbesserungen grün"],
  ["Merksatz 1", "Abstimmung nach Ständen – rot umrahmt"],
  ["Ü2", "Satz „Die Abstimmung nach Ständen ist unfair, weil …“"],
].map(([a, b]) => row([tc(a, 2300, { bold: true, fill: LIGHT }), tc(b, 7906)], 340))));

d.push(h3("Zum Prüfen"));
[
  "Zahlen und Daten (Abgeordnete 290/270/580, Krieg in Amerika, 1787, August 1788, Zinslast) stammen aus meinem Wissen, bitte gegen Ihr Buch lesen. Die Anzahl der Menschen (150 000 / 500 000 / 20 000 000) stammt aus Ihrem Gruppenpuzzle-Material.",
  "Die Analogie zeigt das Verhältnis der Bevölkerung (wenige Privilegierte, viele Menschen). Bei den Abgeordneten stellt der Dritte Stand ca. die Hälfte, nicht 97 %. Beides steht auf Folie 2 und lässt sich im Gespräch trennen.",
  "Die Test-Rückgabe hängt von Ihren Ergebnissen ab. Der Rückmeldungstext ist ein Vorschlag, bitte in Ihren Worten anpassen.",
].forEach((x) => d.push(p([S("• "), S(x)], { after: 40, indent: { left: 240 } })));

(async () => {
  const doc = new Document({ styles: { default: { document: { run: { font: "Calibri", size: 22 } } } }, sections: [{ properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 1000, bottom: 850, left: 850, right: 850 } } }, headers: { default: header }, footers: { default: footer }, children: d }] });
  fs.writeFileSync("../Tafelskript_Generalstaende_Stunde1.docx", await Packer.toBuffer(doc));
  console.log("ok");
})();
