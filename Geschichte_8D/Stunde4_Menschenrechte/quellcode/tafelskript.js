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
const header = new Header({ children: [new Paragraph({ tabStops: [{ type: "right", position: W }], children: [t("Geschichte · Klasse 8D · Menschenrechte", { size: 17, color: MUTED }), t("\tfür die Lehrkraft", { size: 17, color: MUTED })] })] });
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
d.push(kicker("Tafelskript für die Lehrkraft · Einzelstunde (45 Min.) · 24 Schülerinnen und Schüler · eine Lehrkraft · Mittwoch 25.11.2026 · Hospitation"), h1("Menschenrechte: Eine Quelle auswerten"));
d.push(p(t("Ziel: Die Stunde knüpft an die Wege des Widerstands (Stunde 3) an. Klerus und Adel behaupten nach dem Ende der Vorrechte („Jetzt sind alle gleich. Was wollt ihr noch?“). Die Klasse prüft das an einer Quelle, der Erklärung der Menschen- und Bürgerrechte vom 26. August 1789. Schwerpunkt der Hospitation ist die Aktivierung mit Fokus auf dem Kompetenzerwerb: Die Klasse lernt, eine Quelle in drei Schritten auszuwerten (Lesen in eigenen Worten, Einordnen im Vergleich zur Ständeordnung, Begriff). Auf das Vormachen folgt die Progression angeleitet, selbstständig, Urteil. Zum Schluss antwortet die Klasse als Dritter Stand mit einem Beleg aus der Quelle.", { size: 21 }), { after: 100 }));

d.push(h3("Kopf der Planung"));
d.push(table([2600, 7606], [
  ["Ziel (Ich kann …)", "… Artikel der Erklärung der Menschenrechte in eigenen Worten wiedergeben. … einen Artikel mit der Ständeordnung vergleichen: Was war vorher anders? … mit einem Beleg aus der Quelle begründen, welcher Artikel für den Dritten Stand am wichtigsten ist.*"],
  ["Prüfaufgabe: Woran erkenne ich es?", "Arbeitsblatt Seite 1 (Artikel 2, 4 und 6 selbst ausgefüllt), Urteil auf Seite 2 (Beleg, Wirkung, Urteil mit Maßstab, Gegenargument), mündliche Antwort als Dritter Stand („Ich antworte euch: …, weil Artikel … sagt …“). Freiwillige Abgabe, Rückmeldung nach drei Stufen. Punkte oder Noten gibt es heute nicht (später)."],
  ["Mindestziel", "Ziel 1 und 2. Ziel 3 hat einen Stern (Urteil auf zwei Sätze kürzbar). „Auswerten“ bleibt immer."],
  ["Beitrag zur Reihe", "Stunde 4 von 11 (Sequenz, Fassung vom 17.11.). Sie zeigt, wie sich die Rechte nach dem Ende der Vorrechte in einer Quelle niederschlagen, baut die Kompetenz Quellenauswertung auf (wiederkehrend, auch im Test) und führt mit „Für wen gelten die Rechte wirklich?“ zur Verfassung 1791 (Stunde 5)."],
  ["Vorgehen", "Rückblick indirekt (allein, Partnercheck, Plenum), Quellenauswertung direkt: Vormachen (Lehrkraft denkt laut), angeleitet (Artikel 3 mit Hilfe), selbstständig (Artikel 2, 4, 6). Direktes Vorgehen, weil das Vorwissen zur Quellenarbeit gering ist. Urteil indirekt (offenes Format mit Satzanfängen). Direkt oder indirekt wird je Phase entschieden."],
].map(([a, b]) => row([tc(a, 2600, { fill: LIGHT, bold: true }), tc(b, 7606)], 340))));

d.push(h3("Was aus den bisherigen Stunden übernommen wird"));
d.push(...bullets([
  "Stunde 3 (18.11.): Ballhausschwur und Bastille mit fünf Kriterien, Reihenfolge am Zeitstrahl, „wirksam heißt nicht gerechtfertigt“. Rückblick (Folie 6 und 7) wiederholt das; Artikel 3 und 2 knüpfen daran an (Macht geht vom Volk aus, Widerstand). Falls in Stunde 3 etwas nicht geschafft wurde, hier als Wiederholung aufgreifen.",
  "Stunde 2 (11.11.): Ständeordnung und Abstimmung nach Ständen (Vorrechte). Sie ist der Vergleichsmaßstab für Schritt 2 „Einordnen“.",
  "Karikatur als Einstieg und als Abschluss (Antwort in der Rolle des Dritten Stands): selbst gezeichnet, keine historische Quelle. Impuls, keine Analyse.",
  "Rückmeldung der Kollegin (Parallelklasse): Im ersten Teil wurden wichtige Ereignisse des Epochenjahrs 1789 wiederholt, im zweiten Teil gab es ein Arbeitsblatt zu Artikel 1, 2, 3, 4, 6, 10 und 11 mit Zuordnung der Begriffe Freiheit, Religionsfreiheit, Meinungsfreiheit, Volkssouveränität, Grundrechte und Gleichheit sowie Begründung eines wichtigen Artikels. Dieselben Artikel und Begriffe gelten hier (Vergleichbarkeit, Test).",
  "Wunsch der Lehrkraft (17.11.): Schwerpunkt der Hospitation ist die Aktivierung mit Fokus auf einem didaktischen Schwerpunkt (hier: Kompetenzerwerb, Quelle auswerten). Punkte folgen später, freiwillige Abgabe mit Rückmeldung ist möglich.",
  "Regel „höchstens drei Methoden“: Quellenauswertung mit Vormachen und Partnercheck, Urteil, mündliche Abschlussantwort. Karikatur und Rückblick sind Impuls beziehungsweise Routine.",
]));

d.push(h3("Zeitplan (45 Minuten)"));
d.push(table([800, 2300, 4300, 2806], [
  row([th("Min.", 800), th("AVIVA", 2300), th("Phase", 4300), th("Arbeitsform", 2806)], 320),
  ...[["4", "Ankommen", "0  Karikatur, Vermutung, Stundenfrage, Ziele, Ablauf", "still, allein"], ["5", "Vorwissen", "1  Rückblick 1789: richtig oder falsch", "allein, Partnercheck, Plenum"], ["9", "Informieren", "2  Zeitstrahl bis 26. August (3), Drei Schritte und Modell Artikel 1 (4), Prüffrage und Raster (2)", "hört zu, Stichprobe"], ["20", "Verarbeiten", "3  Artikel 3 mit Hilfe (3), Artikel 2, 4, 6 allein (6), Partnercheck (3), Kontrolle (3), Urteil (5)", "allein, zu zweit"], ["7", "Auswerten", "4  Antwort als Dritter Stand (4), Ziele, Abgabe, Ausblick (3)", "mündlich, still"]].map((r) => row(r.map((v, i) => tc(v, [800, 2300, 4300, 2806][i], { bold: i === 0 })), 340)),
]));
d.push(note("ZEIT", "Die Stunde ist voll. Sollbruchstellen: (1) Rückblick: Plenum nur 1 Minute. (2) Übung 2: Wer nach 6 Minuten nur zwei Artikel hat, ist richtig. (3) Partnercheck: auf 2 Minuten kürzen. (4) Urteil: auf zwei Sätze kürzen (Ziel 3 mit Stern). Phase 4 „Auswerten“ wird nicht gestrichen. Die Übungsphase endet 5 Minuten vor Schluss."));

d.push(h3("Worauf die Hospitation achten kann (Aktivierung)"));
d.push(...bullets([
  "Alle in jeder Phase aktiv: Ankommen (Vermutung im Heft), Rückblick (allein schreiben), Modell (Prüffrage mit Stichproben), Übungen (allein, Partner), Urteil (allein), Abschluss (alle denken, wenige sprechen).",
  "Kompetenzaufbau sichtbar: gleiche drei Schritte in Modell, angeleiteter und selbstständiger Phase; Hilfen werden abgebaut (Artikel 1 gelb ausgefüllt, Artikel 3 mit Hilfe, Artikel 2, 4, 6 ohne Hilfe).",
  "Differenzierung: gestufte Hilfen auf dem Blatt (Satzanfänge, Wortspeicher), Artikel 4 als leichter Einstieg, Zusatz* (Artikel 10 und 11) für Schnelle.",
  "Stichproben statt Gespräch (3 bis 6 Personen unterschiedlichen Niveaus, unbenotet), wenig Plenum, wörtliche Ansagen.",
]));

d.push(h3("Lautstärke und Ansagen (wörtlich)"));
d.push(table([4700, 5506], [
  ["„Stifte liegen. Ihr hört nur zu.“", "Sie erklären, niemand schreibt."],
  ["„Jetzt arbeitet ihr allein. Die Aufgabe steht rechts.“", "Einzelarbeit, leise."],
  ["„Flüsterstufe. Nur dein Partner hört dich.“", "Rückblick und Partnercheck."],
  ["„Vergleicht und verbessert mit Grün.“", "Kontrolle im eigenen Blatt."],
  ["„Es ist zu laut. Ab jetzt arbeitet ihr allein und schreibt.“", "Wenn es laut wird, keine Diskussion."],
].map(([a, b]) => row([tc(a, 4700, { fill: LIGHT, bold: true }), tc(b, 5506)], 340))));

d.push(h3("Material und Vorbereitung (Details: Materialliste)"));
d.push(...bullets([
  "Arbeitsblatt_Menschenrechte_Quelle_8D.pdf: eine A4-Seite doppelseitig (Seite 1: Quelle und Auswertungstabelle, Seite 2: Zusatz*, Urteil, Rückmeldung). 26-mal drucken, doppelseitig.",
  "Loesung_Arbeitsblatt_Menschenrechte_8D.pdf (für Sie): Musterantworten für beide Seiten (Stichworte genügen).",
  "Folien 1 bis 21 (Schrift mindestens 27 Punkt; Folie 1 und 19 mit der selbst gezeichneten Karikatur).",
  "Seitentafel: Stundenfrage, Ziele, Ablauf (siehe Phase 0). Mitte der Tafel leer.",
  "Kreide, Grün zum Verbessern, Uhr oder Timer. Kein Material für Schülerrollen, keine Karten.",
]));

// Phase 0
d.push(phase("0", "Ankommen: Karikatur und Vermutung", 4, "AVIVA: Ankommen"));
d.push(warum("Klerus und Adel stellen eine Behauptung auf („Jetzt sind alle gleich.“), die die Stundenfrage trägt. Die Vermutung im Heft wird am Ende mit dem Ergebnis verglichen. Alle haben beim Ankommen etwas zu tun (direkt: Aufträge auf der Folie).", "„Stimmt das? Das prüfen wir an einer Quelle.“"));
d.push(sp());
d.push(step("kommt an, schreibt leise", {
  say: "Stifte liegen. Ihr hört nur zu. Schaut auf die Folie und lest, was Klerus und Adel sagen. Schreibt ins Heft: Reicht das? Ich vermute …",
  do: [
    "Folie 1 steht beim Betreten da: selbst gezeichnete Karikatur (keine historische Quelle) mit Klerus, Adel und dem Dritten Stand (Bauern und Bürger), der zuhört. Klerus und Adel sagen: „Wir sind euch entgegengekommen. Jetzt sind alle gleich. Was wollt ihr noch?“ Auftrag: Heft raus, Datum und Überschrift, Vermutung (ein bis zwei Sätze). Keine Besprechung.",
    "Folie 2 „Das ist mir wichtig“ (20 Sekunden), Folie 3 (Stundenfrage und Begriff „Quelle“, 45 Sekunden), Folie 4 (Lernziele, 30 Sekunden), Folie 5 (Ablauf, 30 Sekunden). Alle Folien: Sie sprechen, die Klasse hört zu.",
  ],
  extra: [boardBox("Seitentafel · Stundenfrage, Ziel und Ablauf", [
    bl("Was steht in der Erklärung der Menschenrechte, und was ändert sie für den Dritten Stand?", { after: 40 }),
    bl("Ich kann Artikel in eigenen Worten wiedergeben.", { after: 20 }),
    bl("Ich kann einen Artikel mit der Ständeordnung vergleichen.", { after: 20 }),
    bl("Ich kann begründen, welcher Artikel am wichtigsten ist.*", { after: 60 }),
    bl("A Karikatur · V Rückblick · I Quelle auswerten · V Blatt, Urteil · A Antwort", { after: 0 }),
  ]), gapP()],
}));

// Phase 1
d.push(phase("1", "Rückblick: richtig oder falsch?", 5, "AVIVA: Vorwissen aktivieren"));
d.push(warum("Der Stoff der Stunde 3 wird echt wiederholt (allein, Partnercheck, Plenum), und zwar mit Falschaussagen statt reiner Abfrage. Die zwei falschen Aussagen sind typische Fehler (Bastille gewaltlos, falsches Datum). Wer in Stunde 3 etwas nicht geschafft hat, sieht es hier noch einmal.", "„1789 hat der Dritte Stand gehandelt. Was geschah danach? Das hört ihr jetzt.“"));
d.push(sp());
d.push(step("schreibt allein, vergleicht zu zweit", {
  say: "Stifte liegen. Ihr hört nur zu. Jetzt arbeitet ihr allein. Die Aufgabe steht rechts. … Flüsterstufe. Nur dein Partner hört dich. Vergleicht.",
  do: [
    "Folie 6 (2 Minuten allein): Schreibe W oder F, verbessere falsche Aussagen ins Heft. Dann 1 Minute Partnercheck, 1 Minute Plenum: Stichproben bei 3 bis 4 Personen unterschiedlichen Niveaus (kurze Antwort, unbenotet).",
    "Folie 7 (Auflösung): Aussage 1 und 2 richtig, 3 und 4 falsch. „Vergleicht und verbessert mit Grün.“",
  ],
  sol: "Erwartung: 1 richtig (Klerus und Adel zusammen zwei von drei Stimmen), 2 richtig (Eid im Ballhaus, 20. Juni), 3 falsch (gewaltsam, Schüsse, etwa 100 Tote), 4 falsch (20. Juni, nicht 14. Juli).",
}));

// Phase 2
d.push(phase("2", "Quelle auswerten: Zeitstrahl, drei Schritte, Modell", 9, "AVIVA: Informieren"));
d.push(warum("Direktes Vorgehen, weil das Vorwissen zur Quellenarbeit gering ist. Zwei Portionen: erst der historische Rahmen (4. und 26. August), dann die Methode mit Modell. Lehrkraft denkt laut an Artikel 1, die Klasse sieht, wie die Schritte zusammengehören. Prüffrage mit Stichprobe, dann Raster. Es wird nicht weitergegangen, wenn etwas nicht verstanden ist.", "„Jetzt seid ihr dran. Auf dem Blatt steht dasselbe in der Tabelle: gelb ist mein Beispiel.“"));
d.push(sp());
d.push(step("hört zu, antwortet in Stichproben", {
  say: "Stifte liegen. Ihr hört nur zu. … Ich mache es vor, ihr schaut zu.",
  do: [
    "Portion 1 (Folie 8, 3 Minuten): Nach der Bastille wächst die Angst vor Aufständen, im Land greifen Bauern Schlösser an. In der Nacht zum 4. August 1789 beschließt die Nationalversammlung, Vorrechte abzuschaffen. Am 26. August folgt die Erklärung der Menschen- und Bürgerrechte, Vorbild ist die amerikanische Unabhängigkeitserklärung von 1776. Warnung vor dem Fehler: Abgeschafft heißt nicht, dass sofort alle gleich leben (viele Abgaben müssen Bauern noch ablösen). Prüffrage (Stichprobe, 2 Personen): „Warum verzichten Klerus und Adel am 4. August?“ Erwartung: aus Angst vor den Aufständen.",
    "Portion 2, Teil 1 (Folie 9, 1 Minute): Drei Schritte, immer in dieser Reihenfolge: Lesen (in eigenen Worten), Einordnen (Was war vorher anders? Wer gewinnt?), Begriff (Wortspeicher). Warnung vor dem Fehler: nicht abschreiben, in eigene Worte fassen.",
    "Portion 2, Teil 2 (Folie 10, 3 Minuten): Sie denken laut an Artikel 1: „Was steht da? Alle Menschen haben von Geburt an dieselben Rechte. Was war vorher anders? Die Geburt entschied über den Stand. Wer gewinnt? Der Dritte Stand. Welcher Begriff? Gleichheit.“ Warnung vor dem Fehler: Gleich an Rechten heißt nicht gleich reich.",
    "Prüffrage (Folie 11, 1 Minute): „Was machst du im Schritt 2?“ und „Warum schreibst du nicht einfach ab?“ Stichprobe bei 4 bis 6 Personen unterschiedlichen Niveaus (unbenotet, getrennt von den mündlichen Noten). Bei Unsicherheit Schritt 2 am Artikel 1 wiederholen. Folie 12 (Raster, 45 Sekunden): drei Stufen, Abgabe freiwillig, keine Selbsteinschätzung.",
  ],
  sol: "Erwartung zur Prüffrage: Ich frage, was vorher anders war und wer gewinnt. / Ich muss es in eigene Worte fassen, so zeigt sich, ob ich es verstanden habe.",
}));

// Phase 3
d.push(phase("3", "Arbeitsblatt: angeleitet, selbstständig, Urteil", 20, "AVIVA: Verarbeiten"));
d.push(warum("Progression der direkten Instruktion: erste Anwendung (Artikel 3 mit Hilfe), selbstständiges Üben (Artikel 2, 4, 6 ohne Hilfe), Partnercheck, Kontrolle, dann das Urteil als Sachurteil mit Beleg und Werturteil mit Maßstab. Artikel 2 und 4 verknüpfen mit Stunde 3 (Widerstand, Freiheit endet, wo sie anderen schadet).", "„Welcher Artikel ist für den Dritten Stand am wichtigsten? Das entscheidet ihr jetzt.“"));
d.push(sp());
d.push(step("3a  Artikel 3 mit Hilfe (3 Min.)", {
  say: "Stifte liegen. Ihr hört nur zu. Jetzt nehmt ihr das Arbeitsblatt. Jetzt arbeitet ihr allein. Die Aufgabe steht rechts.",
  do: [
    "Folie 13: Arbeitsblatt Seite 1, Artikel 3 lesen, „Lesen“ und „Einordnen“ mit den Hilfen ausfüllen (Begriff steht schon). Nach 2 Minuten: „Flüsterstufe. Nur dein Partner hört dich. Vergleicht.“ Folie 14 (1 Minute): Lösung, „Vergleicht und verbessert mit Grün.“ Stichprobe bei 2 Personen.",
  ],
  sol: "Lesen: Die Herrschaft geht vom Volk aus, nicht vom König allein. Einordnen: Vorher bestimmte der König allein, passt zum Ballhausschwur. Begriff: Volkssouveränität.",
}));
d.push(sp());
d.push(step("3b  Artikel 2, 4, 6 selbst (6 Min.)", {
  say: "Stifte liegen. Jetzt arbeitet ihr allein. Die Aufgabe steht rechts.",
  do: [
    "Folie 15: Selbstständiges Üben ohne Hilfen. Durch die Reihen gehen (Heft-Stichprobe bei 4 bis 6 Personen unterschiedlichen Niveaus, mitlesen, nicht kommentieren). Differenzierung: Wer unsicher ist, beginnt mit Artikel 4. Wer früh fertig ist, bearbeitet das Zusatz* (Artikel 10 und 11).",
  ],
}));
d.push(sp());
d.push(step("3c  Partnercheck und Kontrolle (6 Min.)", {
  say: "Stifte liegen. Flüsterstufe. Nur dein Partner hört dich. Vergleicht Artikel 2, 4 und 6. … Vergleicht und verbessert mit Grün.",
  do: [
    "Folie 16 (3 Minuten): Jede Person liest vor, Unterschiede klären, verbessern. Sie hören bei 4 bis 6 Paaren zu, ohne zu helfen.",
    "Folie 17 (3 Minuten): Lösung der Tabelle (Stichworte, andere richtige Formulierungen zählen). Stichprobe bei 2 Personen. Haltung: Artikel 2 nennt das Recht auf Widerstand, aber nicht Gewalt. Das Recht auf Freiheit endet nach Artikel 4, wo man anderen schadet. Sagen Sie: „Wir beschreiben, was dort steht. Wir erklären nicht, dass Gewalt gut ist.“",
  ],
  sol: "Art. 2: Der Staat muss die Rechte schützen, auch das Recht auf Widerstand. Vorher: Ungehorsam war Verbrechen. Grundrechte. Art. 4: Jeder darf tun, was anderen nicht schadet. Vorher: Einsperren ohne Urteil. Freiheit. Art. 6: Gleiches Gesetz für alle. Vorher: Steuerfreiheit, Ämter nur für Adel. Gleichheit.",
}));
d.push(sp());
d.push(step("3d  urteilt allein (5 Min.)", {
  say: "Stifte liegen. Jetzt arbeitet ihr allein. Die Aufgabe steht rechts.",
  do: [
    "Folie 18 (Ü3): Wähle einen Artikel, der für den Dritten Stand am wichtigsten ist; auf Seite 2 mit den Satzanfängen: Beleg („Artikel … sagt: …“), Wirkung („Das ändert für den Dritten Stand, dass …“), Urteil mit Maßstab („Ich halte Artikel … für am wichtigsten, weil …“), Gegenargument („Dagegen spricht, dass …“). Kein richtiges Urteil. Stichprobe bei 4 bis 6 Personen (unbenotet). Wer das Blatt abgeben möchte, legt es nach der Stunde vorn ab.",
  ],
  sol: "Mögliche Urteile (Auswahl, keine Musterlösung). Artikel 6 (gleiches Gesetz für alle: keine Steuerfreiheit mehr). Artikel 3 (Volk bestimmt, Mitbestimmung). Artikel 4 (kein Einsperren ohne Urteil). Artikel 2 (Widerstand ist ein Recht, ohne Gewalt zu nennen). Gegenargument: Auf dem Papier alles gleich, aber wer darf wählen?",
}));

// Phase 4
d.push(phase("4", "Auswerten: Antwort als Dritter Stand", 7, "AVIVA: Auswerten"));
d.push(warum("Die Klasse greift die Aussage von Klerus und Adel vom Anfang auf und antwortet als Dritter Stand mit Beleg aus der Quelle. Das prüft offen, ob die Kompetenz sitzt (mehrere gute Antworten: Position, Beleg, „weil“). Der Vergleich mit der Vermutung zeigt den Lernzuwachs, ohne Floskel.", "„Gilt das für alle? Das sehen wir nächste Woche.“"));
d.push(sp());
d.push(step("schaut auf die Karikatur, denkt, antwortet als Dritter Stand", {
  say: "Stifte liegen. Ihr hört nur zu. Das ist die letzte Aufgabe. Schaut noch einmal auf die Karikatur vom Anfang. Ihr seid der Dritte Stand. Denkt zehn Sekunden nach. Ich rufe euch auf.",
  do: [
    "Folie 19 (Abschlussfrage, 4 Minuten): Mündlich, in der Ich-Form, ohne Schnipsel: „Ich antworte euch: …, weil Artikel … sagt …“. Form nach Verlauf: Blitzlicht (4 bis 6 Personen unterschiedlichen Niveaus gezielt aufrufen, kein Kommentar) oder Duell zu zweit (leise: eine Person Klerus und Adel, eine Dritter Stand). Wer hat umgedacht? Vergleich mit der Vermutung vom Anfang.",
    "Folie 20 (Ziele prüfen, 30 Sekunden, still): Haben wir die Ziele erreicht? Wer das Blatt abgeben möchte, legt es vorn ab. Folie 21 (Ausblick, 30 Sekunden, nur lesen): „Für wen gelten diese Rechte wirklich? Das sehen wir in der nächsten Stunde.“",
  ],
  sol: "Mögliche Antworten als Dritter Stand (Auswahl, keine Musterlösung). „Das reicht uns nicht, weil Artikel 3 sagt, dass die Herrschaft vom Volk ausgeht, und wir mitbestimmen wollen.“ „Das ist ein Anfang, weil Artikel 6 sagt, dass das Gesetz für alle gleich ist.“ „Wir geben uns nicht zufrieden, weil Bauern ihre Abgaben noch ablösen müssen, obwohl auf dem Papier alles gleich ist.“",
}));

d.push(h3("So sieht das Heft und das Blatt am Ende aus"));
d.push(table([2300, 7906], [
  ["Überschrift", "Datum · Überschrift zur Stundenfrage (z. B. „Die Menschenrechte“)"],
  ["Vermutung", "„Reicht das? Ich vermute …“ (Beginn der Stunde)"],
  ["Rückblick", "W oder F zu vier Aussagen, falsche verbessert (grün)"],
  ["Blatt Seite 1", "Tabelle: Artikel 1 (gelb), Artikel 3 mit Hilfe, Artikel 2, 4, 6 selbst; Verbesserungen grün"],
  ["Blatt Seite 2", "Zusatz* (Artikel 10 und 11), Urteil mit Beleg, Wirkung, Urteil und Gegenargument; bei Abgabe Rückmeldung nach Stufen"],
].map(([a, b]) => row([tc(a, 2300, { bold: true, fill: LIGHT }), tc(b, 7906)], 340))));

d.push(h3("Nach der Stunde festhalten (Reflexionslog)"));
d.push(...bullets(["Passten Ziel, Inhalt und Methode zusammen? Wie lief die Zeit? Wie viel echte Lernzeit hatte die Klasse?", "Haben die Schritte „Lesen, Einordnen, Begriff“ gesessen? Wo hakte es (Prüffrage, Artikel 3, selbstständige Phase)?", "Wie viele Blätter wurden abgegeben, auf welcher Stufe lagen sie?", "Wie lief die Aktivierung (Hospitationsrückmeldung zum Kompetenzerwerb)? Was nimmst du mit?", "Wie lief Stunde 3 (Rückmeldung nachtragen), und wie lief der Rückblick?"]));
d.push(h3("Zum Prüfen"));
d.push(...bullets([
  "Wortlaut der Artikel: gekürzte und vereinfachte Fassung, mit der Kollegin und dem Lehrbuch abgleichen (Artikel 3: im Original „Nation“, hier „Nation, also Volk“). Der Test sollte denselben Wortlaut nutzen.",
  "4. August: Aufzählung der abgeschafften Vorrechte und die Ablösepflicht für Bauern gegen das Lehrbuch lesen. Die Aussage „viele Abgaben müssen abgelöst werden“ ist allgemeines Wissen.",
  "Artikel 10 (Zusatz): „Vorher hatte der katholische Glaube Vorrang“ ist eine Vereinfachung (Protestanten waren benachteiligt). Artikel 11: Zensur durch den König. Beides steht nur in der Lösung und im Zusatz*.",
  "Die Karikatur (Klerus, Adel und Dritter Stand) ist selbst gezeichnet, keine historische Quelle. Die Aussage der Figuren („Wir sind euch entgegengekommen. Jetzt sind alle gleich.“) ist bewusst überspitzt, eine erfundene Selbstdarstellung. Historisch war der Verzicht am 4. August nicht freiwillig (Angst vor den Bauernaufständen), und viele Abgaben mussten abgelöst werden. Die Klasse widerlegt die Aussage am Ende selbst (Prüffrage Folie 8, Abschlussantwort).",
  "Die Notenfrage (Punkte, Schlüssel) ist offen: Die Tabelle liefert pro Artikel Felder, das Urteil gibt eine Stufe. Das können wir später ergänzen.",
]));

(async () => {
  const doc = new Document({ styles: { default: { document: { run: { font: "Calibri", size: 22 } } } }, sections: [{ properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 1000, bottom: 850, left: 850, right: 850 } } }, headers: { default: header }, footers: { default: footer }, children: d }] });
  fs.writeFileSync("../Tafelskript_Stunde4_Menschenrechte_8D.docx", await Packer.toBuffer(doc));
  console.log("ok");
})();
