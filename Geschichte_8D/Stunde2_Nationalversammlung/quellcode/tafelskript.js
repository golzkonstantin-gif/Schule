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


const MW = [600, 900, 900, 1250, 900, 1150, 850, 1400];
const mh = (txt, w, fill) => cell(p(t(txt, { bold: true, size: 15, color: "FFFFFF" }), { after: 0, align: AlignmentType.CENTER }), { w, fill: fill || DARK, borders: all(solid(DARK, 6)), m: 40 });
const mc = (txt, w, o = {}) => cell(p(t(txt, { size: o.size || 17, bold: o.bold }), { after: 0, align: AlignmentType.CENTER }), { w, fill: "FFFFFF", borders: all(solid(DARK, 6)), m: 50 });
const matrixT = (filled) => table(MW, [
  row([mh("", MW[0]), mh("nach Ständen", MW[1] + MW[2] + MW[3] + MW[4] + MW[5]), mh("nach Köpfen", MW[6] + MW[7])].map((c, i) => c), 260),
  row([mh("Antrag", MW[0]), mh("Klerus (3)", MW[1], "6B4A9E"), mh("Adel (4)", MW[2], "B9791A"), mh("Dritter Stand (18)", MW[3], "2F6AAE"), mh("Stände mit Ja", MW[4]), mh("Ergebnis", MW[5]), mh("alle Ja (von 25)", MW[6]), mh("Ergebnis", MW[7])], 300),
  ...["A", "B"].map((x) => row([mc(x, MW[0], { bold: true, size: 20 }), mc(filled ? "0" : "", MW[1], { size: 20 }), mc(filled ? "0" : "", MW[2], { size: 20 }), mc(filled ? "18" : "", MW[3], { size: 20 }), mc(filled ? "1 von 3" : "", MW[4]), mc(filled ? "abgelehnt" : "", MW[5], { bold: true, size: 16 }), mc(filled ? "18" : "", MW[6], { size: 20 }), mc(filled ? "angenommen" : "", MW[7], { bold: true, size: 16 })], 480)),
]);
const d = [];
d.push(kicker("Tafelskript für die Lehrkraft · Kurzstunde (25 Min.) · 25 Schülerinnen und Schüler · eine Lehrkraft · Mittwoch 11.11.2026"), h1("Kann der Dritte Stand die Politik einfach wegstimmen?"));
d.push(p(t("Ziel: Die Stunde beginnt mit einer selbst gezeichneten Karikatur: Klerus und Adel unterhalten sich, ein Adliger sagt „Wenn der Dritte Stand unzufrieden ist, soll er einfach gegen die Politik des Königs stimmen.“ Die Klasse schreibt eine Vermutung ins Heft. Das Ziel der Stunde ist, diese Behauptung zu untersuchen: Die Klasse bildet drei Stände und stimmt auf einem Zettel über zwei Anträge ab. Die Lehrkraft zählt während Text und Zeitstrahl aus und füllt eine Matrix an der Tafel aus. Die Klasse rechnet selbst aus: Nach Ständen verliert der Dritte Stand (Bauern und Bürger) immer, obwohl er fast alle Menschen vertritt. Der gekürzte Text zeigt, wie er sich daraufhin wehrt (Nationalversammlung, Ballhausschwur) bis zum Zusammenziehen der Truppen vor Paris. Die Karikatur ist nur der Einstieg, die Methoden sind Abstimmung mit Matrix und Text mit Zeitstrahl (dazu ein kurzer mündlicher Rat), weil die 8C am 06.11. an zu vielen Methoden gelitten hat.", { size: 21, color: MUTED }), { after: 100 }));

d.push(h3("Kopf der Planung"));
d.push(table([2600, 7606], [
  ["Ziel (Ich kann …)", "… erklären, wie die Stände den Dritten Stand (Bauern und Bürger) überstimmen können. … beurteilen, ob der Dritte Stand die Politik durch Abstimmen ändern kann, und das mit einer Zahl begründen. … am Zeitstrahl erklären, wie sich der Dritte Stand wehrt.*"],
  ["Prüfaufgabe: Woran erkenne ich es?", "Heft: Vermutung vom Anfang, Ergebnisse nach Ständen und nach Köpfen selbst ausgerechnet (Ziel 1), Antwort auf die Stundenfrage mit einer Zahl und Vergleich mit der Vermutung (Ziel 2). Zeitstrahl bis „Anfang Juli: Soldaten um Paris“ (Ziel 3). Abschlussfrage mündlich: Antwort an den Adligen (Bezug auf die Karikatur) mit Position, Sachargument und „weil“."],
  ["Mindestziel", "Ziel 1 und 2. Ziel 3 hat einen Stern. „Auswerten“ bleibt immer."],
  ["Beitrag zur Reihe", "Stunde 2 von 11 (Sequenz Französische Revolution). Sie schließt die Abstimmung aus Stunde 1 ab (Ursache des Streits) und führt zur Nationalversammlung und zum Ballhausschwur. Die Karikatur-Analyse und der Merksatz rücken in die Folgestunden (siehe Sequenz)."],
  ["Vorgehen", "Phase 1 direkt (Regel und Anträge vorgegeben), Phase 2 und 3 indirekt (Text und Zeitstrahl selbst erarbeiten), Phase 4 geführt (Matrix rechnen, Antwort auf die Stundenfrage, mündlicher Rat). Direkt oder indirekt wird je Phase entschieden."],
].map(([a, b]) => row([tc(a, 2600, { fill: LIGHT, bold: true }), tc(b, 7606)], 340))));

d.push(h3("Was aus der 8C und aus Stunde 1 übernommen und geändert wurde"));
d.push(...bullets([
  "Rückmeldung 8C (06.11.): Für 45 Minuten waren es zu viele Methoden. Deshalb hier nur drei. Gestrichen oder verschoben: Buchstabenrätsel, Merkbild, Merksatz zum Abschreiben, Karikatur-Dreischritt, Ballhausschwur-Bild.",
  "Rückmeldung 8D (04.11.): Die Abstimmung wurde nicht mehr geschafft, der Zeitstrahl lief gut. Deshalb kommt die Abstimmung zuerst. Die Folie „Das ist mir wichtig“ kommt kurz wieder.",
  "Entscheidung der Lehrkraft (12.11.): Karikatur als Einstieg (Klerus und Adel: „einfach abstimmen“), Stundenziel ist die Untersuchung dieser Behauptung, der Text wird gekürzt (vier Absätze auf einen kurzen Block, vier Daten statt fünf, ohne Quellenbox). Entscheidungen der Lehrkraft (11.11.): Zwei Anträge mit Vorab-Stimme. Keine Wahlhelfer. Gezählt wird nicht live: Die Schüler stimmen auf einem Zettel ab und geben ihn ab, die Lehrkraft zählt während Text und Zeitstrahl aus und füllt die Matrix an der Tafel aus. Gezählt werden nur die Ja-Stimmen je Stand, das Ergebnis nach Ständen und nach Köpfen rechnet die Klasse danach selbst aus.",
  "Historische Vereinfachung: Klerus und Adel stimmen im Spiel immer gleich (Vorrechte behalten). Das ist verkürzt (siehe Prüfhinweise).",
]));

d.push(h3("Zeitplan (25 Minuten)"));
d.push(table([800, 2300, 4300, 2806], [
  row([th("Min.", 800), th("AVIVA", 2300), th("Phase", 4300), th("Arbeitsform", 2806)], 320),
  ...[["4", "Ankommen", "0  Karikatur, Vermutung, Ziele, „Das ist mir wichtig“", "still, allein"], ["4", "Vorwissen", "1  Abstimmung: Stimmzettel ausfüllen und abgeben", "allein, leise"], ["4", "Informieren", "2  Text (gekürzt, Sie zählen aus)", "allein, leise"], ["3", "Verarbeiten", "3  Zeitstrahl bis Anfang Juli (Sie schreiben die Matrix an)", "allein, Kontrolle per Folie"], ["10", "Auswerten", "4  Matrix rechnen, Antwort auf die Stundenfrage, Antwort an den Adligen, Ziele, Ausblick", "allein, mündlich"]].map((r) => row(r.map((v, i) => tc(v, [800, 2300, 4300, 2806][i], { bold: i < 3 })), 320)),
]));
d.push(note("ZEIT", "Die Stunde ist sehr knapp: Kein Puffer. Das Auszählen der Zettel (25 Zettel, 2 Anträge) dauert etwa 4 bis 6 Minuten und läuft in Phase 2 und 3. Wird es nicht fertig, schreiben Sie vor Phase 4 nur Antrag A in die Matrix (Antrag B entfällt). Phase 4 wird nicht gestrichen."));

d.push(h3("Lautstärke und Ansagen (wörtlich)"));
d.push(table([4700, 5506], [
  ["„Stifte liegen. Ihr hört nur zu.“", "Sie erklären, niemand schreibt."],
  ["„Jetzt arbeitet ihr allein. Die Aufgabe steht rechts.“", "Einzelarbeit, leise."],
  ["„Stifte hinlegen. Ich sammle die Zettel ein.“", "Einsammeln der Stimmzettel (Phase 1)."],
  ["„Vergleicht und verbessert mit Grün.“", "Kontrolle im Heft bzw. Zeitstrahl."],
  ["„Es ist zu laut. Ab jetzt arbeitet ihr allein und schreibt.“", "Wenn es laut wird, keine Diskussion."],
].map(([a, b]) => row([tc(a, 4700, { fill: LIGHT, bold: true }), tc(b, 5506)], 340))));

d.push(h3("Material und Vorbereitung (Details: Materialliste)"));
d.push(...bullets([
  "Druckvorlage „Karten und Stimmzettel“: Seite 1 und 2: 28 Standeskarten (3 Klerus, 4 Adel, 18 Dritter Stand, 3 Reserve). Seite 3: Stimmzettel, acht pro A4, 4-mal drucken und schneiden (32 Zettel).",
  "Tafelmatrix_Auswertung_8D.pdf: Seite 1 Matrix leer (zum Abzeichnen an die Tafel), Seite 2 Matrix ausgefüllt (erwartet), Seite 3 Strichliste zum Auszählen.",
  "Text mit Zeitstrahl: „Der Dritte Stand wehrt sich“ (gekürzte Fassung, Arbeitsblatt_Der_Dritte_Stand_wehrt_sich.pdf, 27 Stück; ein Drittel A4 Text, vier Daten). Lösung für Sie: Loesung_Arbeitsblatt_Der_Dritte_Stand_wehrt_sich.pdf. Kein Schnipsel für die Abschlussfrage (mündlich).",
  "Vorab: Karten verdeckt mit dem Stimmzettel auf die Plätze legen. Klerus (3) und Adel (4) gezielt verteilen (zuverlässige Personen), alle übrigen sind Dritter Stand (18). Keine Wahlhelfer. Matrix vorab an die Tafel zeichnen (Mitte), Platz für den Zeitstrahl lassen.",
  "Folien 1 bis 15 (Schrift mindestens 27 Punkt; Folie 1 mit der selbst gezeichneten Karikatur), Kreide, Grün zum Verbessern, Uhr oder Timer.",
]));

// Phase 0
d.push(phase("0", "Ankommen: Karikatur und Vermutung", 4, "AVIVA: Ankommen"));
d.push(warum("Die Behauptung „einfach abstimmen“ weckt die Stundenfrage und macht neugierig; die Vermutung im Heft wird am Ende mit dem Ergebnis verglichen (Vorher-Nachher, Eingangsdiagnose ohne Fachbegriff). Alle haben beim Ankommen etwas zu tun (direkt, von der Lehrkraft geführt).", "„Klerus und Adel sagen: Einfach abstimmen. Ob das stimmt, prüfen wir jetzt.“"));
d.push(sp());
d.push(step("kommt an, schreibt leise", {
  say: "Stifte liegen. Ihr hört nur zu. Lasst die Karte verdeckt liegen. Schaut auf die Folie und lest, was der Adlige sagt. Schreibt ins Heft: Stimmt das? Ich vermute …, weil …",
  do: [
    "Folie 1 steht beim Betreten da: selbst gezeichnete Karikatur (keine historische Quelle). Klerus und Adel unterhalten sich, der Adlige sagt: „Wenn der Dritte Stand unzufrieden ist, soll er einfach gegen die Politik des Königs stimmen.“ Auftrag: Heft raus, Datum und Überschrift, „Stimmt das? Ich vermute …, weil …“ (etwa 2 Minuten, keine Besprechung, keine Bewertung). Sie gehen durch die Reihen und lesen bei 3 bis 4 Personen mit (nicht kommentieren).",
    "Folie 2 „Das ist mir wichtig“ (20 Sekunden, ruhig, ohne Diskussion), Folie 3 (Stundenfrage, 30 Sekunden: „Kann der Dritte Stand die Politik des Königs einfach wegstimmen? Das prüfen wir mit einer Abstimmung.“), Folie 4 (Lernziele, 30 Sekunden), Folie 5 (Ablauf, 30 Sekunden). Hinweis zum Stern: „Was wir eventuell nicht schaffen, hat einen Stern.“",
  ],
  extra: [boardBox("Seitentafel · Stundenfrage, Ziel und Ablauf", [
    bl("Kann der Dritte Stand die Politik des Königs einfach wegstimmen?", { after: 40 }),
    bl("Ich kann erklären, wie die Stände den Dritten Stand überstimmen können.", { after: 20 }),
    bl("Ich kann beurteilen, ob der Dritte Stand die Politik durch Abstimmen ändern kann (mit einer Zahl).", { after: 20 }),
    bl("Ich kann am Zeitstrahl erklären, wie sich der Dritte Stand wehrt.*", { after: 60 }),
    bl("A Karikatur · V Abstimmung · I Text · V Zeitstrahl · A Ergebnis und Rat", { after: 0 }),
  ]), gapP()],
}));

// Phase 1
d.push(phase("1", "Abstimmung: Stimmzettel ausfüllen und abgeben", 4, "AVIVA: Vorwissen aktivieren"));
d.push(warum("Die Behauptung wird getestet, und der in Stunde 1 nicht geschaffte Stoff wird nachgeholt: Die Klasse erlebt die Regel „jeder Stand eine Stimme“, indem sie in ihrer Rolle abstimmt (Klerus und Adel nach ihren Vorrechten, der Dritte Stand (Bauern und Bürger) nach Überzeugung). Das Zählen läuft ohne Klasse (einfach, ruhig), die Auswertung kommt am Stundenende (direkt: Regel und Anträge vorgegeben).", "„Klerus und Adel sagen: Einfach abstimmen. Gleich seht ihr, wie die Abstimmung ausgeht. Vorher lest ihr, wie es 1789 weiterging.“"));
d.push(sp());
d.push(step("arbeitet allein", {
  say: "Stifte liegen. Ihr hört nur zu. Das sind die beiden Anträge. Jeder Stand hat eine Stimme. … Jetzt arbeitet ihr allein. Die Aufgabe steht rechts.",
  do: [
    "Folie 6 (Anträge, 1 Minute): A „Auch Adel und Klerus zahlen künftig Steuern.“ B „Die Stände werden abgeschafft. Alle haben gleiche Rechte.“ Klerus 3, Adel 4, Dritter Stand (Bauern und Bürger) 18 Personen. Regel: Jeder Stand hat eine Stimme.",
    "Folie 7 (Ü1, 2 Minuten): Karte umdrehen, leise lesen, auf dem Stimmzettel Stand und Stimme bei A und B ankreuzen. Klerus und Adel stimmen nach ihrer Rolle (Vorrechte behalten), der Dritte Stand nach seiner Überzeugung. Keine Absprache, keine Bewegung.",
    "Einsammeln (1 Minute): „Stifte hinlegen. Ich sammle die Zettel ein.“ Sie gehen durch die Reihen und sammeln ein. Danach sofort Folie 8.",
  ],
}));

// Phase 2
d.push(phase("2", "Text (gekürzt)", 4, "AVIVA: Informieren"));
d.push(warum("Wie der Dritte Stand die Macht des Königs infrage stellt, wird selbst aus dem gekürzten Text erarbeitet (indirekt: Lesen allein, Unterstreichen). Der Text ist bewusst kurz (ein Drittel A4, vier Daten), damit die Stunde in 25 Minuten gelingt. Kein Vorlesen. Während die Klasse arbeitet, zählen Sie die Zettel aus.", "„Der Dritte Stand hat gegen die Regel verloren, aber er gibt nicht auf. Wie reagiert er? Das lest ihr jetzt.“"));
d.push(sp());
d.push(step("liest allein", {
  say: "Stifte liegen. Ihr hört nur zu. Das Blatt liegt vor euch. Jetzt arbeitet ihr allein. Die Aufgabe steht rechts.",
  do: [
    "Folie 8 (Ü2, 4 Minuten): Text lesen, alle Daten unterstreichen, die Wörter unter dem Text lesen (Flugschrift, Verfassung, Nationalversammlung, Ballhalle). Gekürzte Fassung gegenüber der 8C: vier kurze Absätze, Sieyès nur in einem Satz (ALLES, NICHTS, ETWAS ZU SEIN), keine Quellenbox.",
    "Warnung vor dem Fehler: Das Lehrbuch nennt den 20. Juni für die Nationalversammlung. Historisch ist es der 17. Juni (Nationalversammlung) und der 20. Juni (Ballhausschwur). Sagen Sie das vorab in einem Satz.",
    "Auszählen (Sie, am Pult, leise): Zettel nach Stand sortieren (Klerus 3, Adel 4, Dritter Stand 18), Ja und Nein je Antrag in der Strichliste (Tafelmatrix-Datei, Seite 3). Kontrolle: Ja plus Nein je Stand ist die Personenzahl.",
  ],
}));

// Phase 3
d.push(phase("3", "Zeitstrahl bis Anfang Juli", 3, "AVIVA: Verarbeiten"));
d.push(warum("Die Ereignisse werden auf dem Zeitstrahl in die richtige Reihenfolge gebracht (angeleitet: vier Daten vorgegeben, Ereignisse selbst eintragen). Der letzte Eintrag führt zur nächsten Stunde.", "„Am Ende des Zeitstrahls zieht der König Soldaten um Paris zusammen. Gleich seht ihr erst das Ergebnis eurer Abstimmung.“"));
d.push(sp());
d.push(step("arbeitet allein", {
  say: "Stifte liegen. Jetzt arbeitet ihr allein. Die Aufgabe steht rechts. … Vergleicht und verbessert mit Grün.",
  do: [
    "Folie 9 (Ü3, 2 Minuten): Zu jedem der vier Daten das Ereignis in Stichworten eintragen. Danach Folie 10 (Lösung, 1 Minute): „Vergleicht und verbessert mit Grün.“",
    "Während der Zeitstrahl läuft, schreiben Sie die Ja-Zahlen in die Matrix an der Tafel (nur die Spalten Klerus, Adel, Dritter Stand). Die Spalten „Stände mit Ja“, „Ergebnis“ und „alle Ja“ bleiben leer, sie rechnet die Klasse in Phase 4 aus. Bei Zeitnot: nur Antrag A.",
    "Wenn die Zeit knapp ist: nur die drei letzten Einträge (17. Juni, 20. Juni, Anfang Juli) ausfüllen.",
  ],
  extra: [boardBox("Mitte · Auswertungsmatrix (vorab zeichnen; unten erwartetes Ergebnis nach Phase 4)", [matrixT(false), p(t(" "), { after: 40 }), matrixT(true)]), gapP(),
  boardBox("Mitte · Zeitstrahl (Lösung, Folie 10)", [
    bl("5. Mai 1789    Eröffnung der Generalstände, Streit um die Abstimmung"),
    bl("17. Juni 1789    Nationalversammlung"),
    bl("20. Juni 1789    Ballhausschwur"),
    bl("Anfang Juli 1789    Soldaten um Paris, Angst in Paris", { after: 0 }),
  ]), gapP()],
}));

// Phase 4
d.push(phase("4", "Auswerten: Ergebnis, Stundenfrage, Antwort an den Adligen", 10, "AVIVA: Auswerten"));
d.push(warum("Die Klasse rechnet selbst aus und beantwortet die Stundenfrage: Nach Ständen verliert der Dritte Stand (Bauern und Bürger) immer, nach Köpfen hätte er gewonnen, „einfach abstimmen“ funktioniert so nicht. Sie vergleicht mit ihrer Vermutung vom Anfang. Die Klasse nimmt am Ende Bezug auf die Aussage des Adligen vom Anfang (Karikatur) und antwortet ihm mündlich: Das überprüft das Gelernte offen mit mehreren guten Antworten (Vorstufe zu den Urteilsaufgaben der Reihe). Diese Phase wird nie gestrichen.", "„Der Adlige vom Anfang sagt: Einfach abstimmen. Was zeigt eure Abstimmung? Was antwortet ihr ihm?“"));
d.push(sp());
d.push(step("rechnet und schreibt", {
  say: "Stifte liegen. Ihr hört nur zu. Hier ist das Ergebnis eurer Abstimmung. … Jetzt arbeitet ihr allein. Die Aufgabe steht rechts.",
  do: [
    "Folie 11 (Ü4, 4 Minuten): Die Matrix mit den Ja-Zahlen steht an der Tafel. (1) Ablesen und rechnen, etwa 1 Minute: Nach Ständen: Wie viele der drei Stände sagen Ja? Mindestens 2 von 3 = angenommen. Nach Köpfen: Alle Ja zusammenzählen (von 25), mehr als die Hälfte = angenommen. (2) Ins Heft, etwa 3 Minuten: „Nach Ständen …, nach Köpfen … Das zeigt: … (nenne Zahlen)“ und Vergleich mit der Vermutung vom Anfang. Die Antwort wird nicht vorgegeben. Die Antwort an den Adligen kommt danach mündlich (Folie 13). Wenn die Zeit knapp ist, genügen die Ergebnisse und ein Antwortsatz. Heft-Stichprobe: bei 4 bis 6 Personen unterschiedlichen Niveaus mitlesen, nicht kommentieren (unbenotet).",
    "Folie 12 (Kontrolle, 1 Minute): „Vergleicht und verbessert mit Grün.“ Sie tragen die Ergebnisse in die Matrix an der Tafel ein. Einordnung in einem Satz: „Im Spiel stimmen Klerus und Adel immer gleich. In Frankreich sind sie zusammen nur etwa 3 % der Menschen.“ Gezielte Abfrage nur bei Bedarf (2 bis 3 Personen, kurz): „Wie haben Klerus und Adel abgestimmt? Warum?“",
  ],
  sol: "Rechnen: Nach Ständen 1 von 3 Ständen Ja, abgelehnt. Nach Köpfen 18 von 25 Ja, angenommen (bei A und B gleich, echte Zahlen gelten). Antwort auf die Stundenfrage (mit Zahl, mehrere Formulierungen möglich): „Nein, das stimmt nicht, weil Klerus und Adel zusammen 2 von 3 Stimmen haben, obwohl sie nur etwa 3 % der Menschen sind“ oder „Nur nach Köpfen hätte der Dritte Stand gewonnen: 18 von 25“. Vergleich: Wer vermutet hat, einfach abstimmen reiche nicht, sieht sich bestätigt; wer das Gegenteil vermutet hat, korrigiert.",
}));
d.push(sp());
d.push(step("schaut auf die Karikatur, denkt, dann mündlich", {
  say: "Stifte liegen. Ihr hört nur zu. Das ist die letzte Aufgabe. Schaut noch einmal auf die Karikatur vom Anfang. Denkt zehn Sekunden nach. Ich rufe euch auf.",
  do: [
    "Folie 13 (Abschlussfrage, 3 Minuten): Die Karikatur vom Anfang erscheint wieder, daneben die Aussage des Adligen: „Wenn der Dritte Stand unzufrieden ist, soll er einfach gegen die Politik des Königs stimmen.“ Auftrag: „Was antwortest du dem Adligen?“ Die Antwort kommt mündlich, ohne Schnipsel: „Ich antworte dir, …, weil …“. Mehrere gute Antworten, geprüft werden Position, ein Sachargument aus der Stunde (Zahl aus der Matrix, 97 % der Menschen, 1 von 3 Stimmen, Nationalversammlung) und ein „weil“. Erst zehn Sekunden denken, dann mündlich. Zum Schluss: Wer hat seine Vermutung vom Anfang geändert? (nur Blick ins Heft, keine Abfrage).",
    "Form nach Verlauf wählen: (a) Blitzlicht (empfohlen): 4 bis 6 Personen unterschiedlichen Niveaus gezielt aufrufen, kein Kommentar. (b) Duell zu zweit, leise („Flüsterstufe. Nur dein Partner hört dich.“): eine Person ist der Adlige, die andere antwortet, danach zwei Personen berichten (nur wenn Zeit ist). Unbenotet.",
    "Folie 14 (30 Sekunden, still): Haben wir die Ziele erreicht? Keine Antworten ins Plenum. Folie 15 (30 Sekunden, nur lesen): Ausblick, Soldaten um Paris. „Wie reagieren die Menschen in Paris? Das sehen wir in der nächsten Stunde.“",
  ],
  sol: "Mögliche Antworten an den Adligen (Auswahl, keine Musterlösung). „Das funktioniert nicht, weil nach Ständen Klerus und Adel 2 von 3 Stimmen haben, obwohl der Dritte Stand (Bauern und Bürger) fast alle Menschen vertritt.“ „Das stimmt nur, wenn nach Köpfen abgestimmt wird: dort hätte der Dritte Stand 18 von 25 Stimmen.“ „Ihr wollt nicht nach Köpfen abstimmen, weil ihr eure Vorrechte behalten wollt.“ „Wenn Abstimmen nichts bringt, wehrt sich der Dritte Stand anders, er erklärt sich zur Nationalversammlung.“ Kriterien: Position, ein Sachargument aus der Stunde (Zahl, 97 %, Nationalversammlung), ein „weil“.",
}));
d.push(sp());
d.push(note("SOLLBRUCHSTELLE", "Ist das Auszählen vor Phase 4 nicht fertig, schreiben Sie nur Antrag A an. Wird das Ausrechnen und Schreiben länger als 4 Minuten, entfällt der Vergleich mit der Vermutung und die Antwort an den Adligen läuft als Blitzlicht (2 Minuten). Der Vergleich kann in Stunde 3 als Wiederholung wieder aufgegriffen werden."));

d.push(h3("So sieht das Heft am Ende der Stunde aus"));
d.push(table([2300, 7906], [
  ["Überschrift", "Datum · Überschrift zur Stundenfrage (z. B. „Einfach abstimmen?“)"],
  ["Text und Zeitstrahl", "Blatt eingeklebt, Daten unterstrichen, Ereignisse eingetragen, Verbesserungen grün"],
  ["Abstimmung", "Vermutung vom Anfang, Ergebnisse nach Ständen und nach Köpfen (gerechnet aus der Matrix), Antwort auf die Stundenfrage mit einer Zahl, Vergleich mit der Vermutung, Verbesserungen grün"],
].map(([a, b]) => row([tc(a, 2300, { bold: true, fill: LIGHT }), tc(b, 7906)], 340))));

d.push(h3("Nach der Stunde festhalten (Reflexionslog)"));
d.push(...bullets(["Passten Ziel, Inhalt und Methode zusammen? Wie lief die Zeit? Wie viel echte Lernzeit hatte die Klasse?", "Hat das Auszählen in Phase 2 und 3 funktioniert (Strichliste)? Hat es die Zeit gestört?", "Wo wurde es laut, und welche Ansage half?", "Wie viele nannten in der Abschlussfrage ein Sachargument aus der Stunde? Welche Form (Blitzlicht, Duell) hat funktioniert?", "Wurde der Zeitstrahl bis Anfang Juli erreicht? Was muss in Stunde 3 wiederholt werden?"]));
d.push(h3("Zum Prüfen"));
d.push(...bullets([
  "Das Spiel vereinfacht: Im Spiel stimmen Klerus und Adel immer gleich. Historisch waren die Stände nicht immer einheitlich (viele einfache Pfarrer unterstützten den Dritten Stand, einzelne Adlige schlossen sich später an), bei der Abstimmungsregel hielten die Privilegierten aber oft zusammen. Die Anträge A und B sind Spielanträge, historisch ging der Streit um die Abstimmungsregel (1789 wurde nicht nach Köpfen abgestimmt).",
  "Die Standeskarten verteilen die Rollen: 3 Klerus (12 %), 4 Adel (16 %), 18 Dritter Stand (72 %). Das ist verzerrt, in Frankreich sind Klerus und Adel zusammen etwa 3 % (Klerus ca. 150 000, Adel ca. 500 000, Dritter Stand ca. 20 000 000 Menschen, Zahlen aus Ihrem Gruppenpuzzle-Material).",
  "Zahlen und Daten stammen aus dem Text der 8C (hier gekürzt) und sind gegen das Schulbuch zu lesen (Lehrbuch S. 112/113): 17. Juni Nationalversammlung, 20. Juni Ballhausschwur, Anfang Juli Soldaten und Entlassung des Finanzministers (im gekürzten Text nicht mehr erwähnt).",
  "Die Karikatur (Klerus und Adel) und sind selbst gezeichnet, keine historischen Quellen oder Porträts.",
]));

(async () => {
  const doc = new Document({ styles: { default: { document: { run: { font: "Calibri", size: 22 } } } }, sections: [{ properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 1000, bottom: 850, left: 850, right: 850 } } }, headers: { default: header }, footers: { default: footer }, children: d }] });
  fs.writeFileSync("../Tafelskript_Stunde2_Nationalversammlung_8D.docx", await Packer.toBuffer(doc));
  console.log("ok");
})();
