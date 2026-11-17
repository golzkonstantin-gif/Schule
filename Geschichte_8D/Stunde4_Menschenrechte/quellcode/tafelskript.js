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
  if (parts.ziel) ch.push(p([t("Dient Ziel: ", { bold: true, size: 19, color: "1B5E3A" }), t(parts.ziel, { size: 19, color: "1B5E3A" })], { after: 60 }));
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
d.push(p(t("Ziel: Die Stunde knüpft an die Wege des Widerstands (Stunde 3) an. In der Karikatur sprechen Klerus und Adel vor zwei Bildern (Eid im Ballhaus, Sturm auf die Bastille): „Ihr habt geschworen. Ihr habt die Bastille gestürmt. Wir kamen euch entgegen. Was wollt ihr noch?“ Die Rückschau erschließt die Ereignisse, die Informationsphase erklärt den 4. und den 26. August, und die Klasse wertet die Erklärung der Menschen- und Bürgerrechte in drei Schritten aus (Lesen in eigenen Worten, Einordnen, Begriff). Der Dritte Stand antwortet am Ende mit Beleg: Gnade oder Recht? Schwerpunkt der Hospitation ist die Aktivierung mit Fokus auf dem Kompetenzerwerb (Quelle auswerten).", { size: 21 }), { after: 100 }));

d.push(h3("Kopf der Planung"));
d.push(table([2600, 7606], [
  ["Ziel (Ich kann …)", "1  … Artikel der Erklärung der Menschenrechte in eigenen Worten wiedergeben und den passenden Begriff nennen.  2  … einen Artikel mit der Ständeordnung vergleichen und einem Ereignis von 1789 zuordnen.  3  … mit Beleg und Maßstab begründen, welcher Artikel für den Dritten Stand am wichtigsten ist.*"],
  ["Prüfaufgabe: Woran erkenne ich es?", "Arbeitsblatt Seite 1 (Artikel 2, 4 und 6 selbst ausgefüllt), Zuordnung Bild und Artikel im Partnercheck, Urteil auf Seite 2 (Beleg, Wirkung, Maßstab, Urteil, Gegenargument), mündliche Antwort als Dritter Stand („Ich antworte euch: Wir wollen …, weil Artikel … sagt …, früher …“). Freiwillige Abgabe, Rückmeldung nach drei Stufen (gleich den Zielen). Punkte oder Noten gibt es heute nicht (später)."],
  ["Mindestziel", "Ziel 1 und 2. Ziel 3 hat einen Stern (Urteil auf zwei Sätze kürzbar). „Auswerten“ bleibt immer."],
  ["Beitrag zur Reihe", "Stunde 4 von 11 (Sequenz, Fassung vom 17.11.). Sie zeigt, wie sich die Rechte nach dem Ende der Vorrechte in einer Quelle niederschlagen, baut die Kompetenz Quellenauswertung auf (wiederkehrend, auch im Test) und führt mit „Für wen gelten die Rechte wirklich?“ zur Verfassung 1791 (Stunde 5). Der 4. August ist Rahmen und Merksatz-Inhalt, nicht eigenes Ziel."],
  ["Vorgehen", "Rückschau indirekt (allein, Partnercheck, Plenum), Quellenauswertung direkt: Vormachen (Lehrkraft denkt laut), angeleitet (Artikel 3 mit Hilfe), selbstständig (Artikel 2, 4, 6). Direktes Vorgehen, weil das Vorwissen zur Quellenarbeit gering ist. Urteil indirekt (offenes Format mit Satzanfängen). Direkt oder indirekt wird je Phase entschieden."],
].map(([a, b]) => row([tc(a, 2600, { fill: LIGHT, bold: true }), tc(b, 7606)], 340))));

d.push(h3("Bezug jedes Schritts zum Ziel (Prüfung)"));
d.push(table([4400, 1500, 4306], [
  row([th("Schritt", 4400), th("Ziel", 1500), th("Begründung", 4306)], 320),
  ...[
    ["Karikatur und Vermutung (Folie 1)", "Rahmen", "Stundenfrage, Vorher-Nachher mit der Abschlussantwort"],
    ["Rückschau an den Bildern (Folie 5)", "2", "Ereignisse 1789 kennen, um sie Artikeln zuzuordnen"],
    ["Kette und 4./26. August (Folie 6)", "2", "Rahmen und Gegensatz Gnade/Vorrecht und Recht; Merksatz-Inhalt"],
    ["Drei Schritte, Modell, Prüffrage (Folie 7 bis 9)", "1, 2", "Methode der Auswertung modellieren, Verständnis prüfen"],
    ["Artikel 3 mit Hilfe, Artikel 2, 4, 6 (Folie 10 bis 14)", "1, 2", "angeleitet, dann selbstständig üben; Begriff und Einordnen"],
    ["Zuordnung Bild und Artikel (Folie 13, 15)", "2", "Ereignis zuordnen, Aha: Was sie taten, wird zum Recht"],
    ["Urteil (Folie 16)", "3", "Beleg und Maßstab, zweite Perspektive im Gegenargument"],
    ["Merksatz (Folie 17)", "1, 2", "Ergebnissicherung, Wissen für den Test"],
    ["Abschlussfrage (Folie 18)", "1, 2, 3", "prüft alle Ziele, Vergleich mit der Vermutung"],
    ["Ziele prüfen (Folie 19)", "1, 2, 3", "Zielprüfung, Abgabe"],
    ["Ausblick (Folie 20)", "Brücke", "kein Ziel der Stunde, bereitet Stunde 5 vor"],
  ].map((r) => row(r.map((v, i) => tc(v, [4400, 1500, 4306][i], { bold: i === 1 })), 320)),
]));

d.push(h3("Was aus den bisherigen Stunden übernommen wird"));
d.push(...bullets([
  "Stunde 3 (18.11.): Ballhausschwur und Bastille mit fünf Kriterien, Reihenfolge am Zeitstrahl, „wirksam heißt nicht gerechtfertigt“. Die Rückschau (Folie 5 und 6) wiederholt das an zwei Bildern; Artikel 3 und 2 knüpfen daran an (Macht geht vom Volk aus, Widerstand). Falls in Stunde 3 etwas nicht geschafft wurde, hier als Wiederholung aufgreifen. Die 8C macht Stunde 3 ebenfalls und danach die Menschenrechte: Dasselbe Material passt, da die Rückschau die Vorgeschichte liefert.",
  "Stunde 2 (11.11.): Ständeordnung und Abstimmung nach Ständen (Vorrechte). Sie ist der Vergleichsmaßstab für Schritt 2 „Einordnen“.",
  "Karikatur als Einstieg, Rückschau und Abschluss (Antwort in der Rolle des Dritten Stands): selbst gezeichnet, keine historische Quelle. An der Wand hängen zwei unbeschriftete Bilder, davor sprechen Klerus und Adel. Die Aussage ist bewusst überspitzt (Entgegenkommen als Gnade, historisch nicht freiwillig).",
  "Parallelklasse (Rückmeldung der Lehrkraft): erster Teil Epochenjahr 1789, zweiter Teil Arbeitsblatt zu Artikel 1, 2, 3, 4, 6, 10, 11 mit Begriffen Freiheit, Religionsfreiheit, Meinungsfreiheit, Volkssouveränität, Grundrechte, Gleichheit. Dieselben Artikel und Begriffe gelten hier (Vergleichbarkeit, Test).",
  "Wunsch der Lehrkraft (17.11.): Schwerpunkt der Hospitation ist die Aktivierung mit Fokus auf Kompetenzerwerb (Quelle auswerten). Punkte folgen später, freiwillige Abgabe mit Rückmeldung. „Das ist mir wichtig“ entfällt, weil es eingeführt ist.",
  "Regel „höchstens drei Methoden“: Quellenauswertung (Vormachen, angeleitet, selbstständig, Partnercheck), Urteil, mündliche Abschlussantwort. Karikatur und Rückschau sind Impuls beziehungsweise Routine.",
]));

d.push(h3("Zeitplan (45 Minuten)"));
d.push(table([800, 2300, 4300, 2806], [
  row([th("Min.", 800), th("AVIVA", 2300), th("Phase", 4300), th("Arbeitsform", 2806)], 320),
  ...[["4", "Ankommen", "0  Karikatur, Vermutung, Stundenfrage, Ziele, Ablauf", "still, allein"], ["5", "Vorwissen", "1  Rückschau an den Bildern (4), Kette (1)", "allein, Partnercheck, Plenum"], ["8", "Informieren", "2  4. und 26. August (2), Drei Schritte (1), Modell Artikel 1 (3), Prüffrage (1), Reserve (1)", "hört zu, Stichprobe"], ["20", "Verarbeiten", "3  Artikel 3 mit Hilfe (3), Artikel 2, 4, 6 allein (6), Partnercheck und Zuordnung (4), Kontrolle und Aha (3), Urteil (4)", "allein, zu zweit"], ["8", "Auswerten", "4  Merksatz (2), Antwort als Dritter Stand (4), Ziele (1), Ausblick (1)", "schreibt, mündlich, still"]].map((r) => row(r.map((v, i) => tc(v, [800, 2300, 4300, 2806][i], { bold: i === 0 })), 340)),
]));
d.push(note("ZEIT", "Die Stunde ist dicht. Sollbruchstellen: (1) Rückschau: Plenum nur 30 Sekunden. (2) Reserve in Phase 2 (1 Minute) für erneutes Erklären, sonst weiter. (3) Übung 2: Wer nach 6 Minuten nur zwei Artikel hat, ist richtig. (4) Partnercheck: nur eine Zuordnung genügt. (5) Urteil: auf zwei Sätze kürzen (Ziel 3 mit Stern). Phase 4 „Auswerten“ wird nicht gestrichen. Die Übungsphase endet 8 Minuten vor Schluss."));

d.push(h3("Worauf die Hospitation achten kann (Aktivierung)"));
d.push(...bullets([
  "Alle in jeder Phase aktiv: Ankommen (Vermutung im Heft), Rückschau (allein schreiben, Partner), Modell (Prüffrage mit Stichproben), Übungen (allein, Partner), Urteil (allein), Merksatz (abschreiben), Abschluss (alle denken, wenige sprechen).",
  "Kompetenzaufbau sichtbar: gleiche drei Schritte in Modell, angeleiteter und selbstständiger Phase; Hilfen werden abgebaut (Artikel 1 gelb ausgefüllt, Artikel 3 mit Hilfe, Artikel 2, 4, 6 ohne Hilfe).",
  "Kognitive Aktivierung: Entschlüsseln der Sprechblase an den Bildern, Zuordnung Bild und Artikel mit Begründung, Maßstab und Gegenargument im Urteil.",
  "Differenzierung: gestufte Hilfen auf dem Blatt (Satzanfänge, Wortspeicher), Artikel 4 als leichter Einstieg, Zusatz* (Artikel 10 und 11) für Schnelle.",
  "Stichproben statt Gespräch (3 bis 6 Personen unterschiedlichen Niveaus, unbenotet), wenig Plenum, wörtliche Ansagen.",
]));

d.push(h3("Lautstärke und Ansagen (wörtlich)"));
d.push(table([4700, 5506], [
  ["„Stifte liegen. Ihr hört nur zu.“", "Sie erklären, niemand schreibt."],
  ["„Jetzt arbeitet ihr allein. Die Aufgabe steht rechts.“", "Einzelarbeit, leise."],
  ["„Flüsterstufe. Nur dein Partner hört dich.“", "Rückschau und Partnercheck."],
  ["„Vergleicht und verbessert mit Grün.“", "Kontrolle im eigenen Blatt."],
  ["„Es ist zu laut. Ab jetzt arbeitet ihr allein und schreibt.“", "Wenn es laut wird, keine Diskussion."],
].map(([a, b]) => row([tc(a, 4700, { fill: LIGHT, bold: true }), tc(b, 5506)], 340))));

d.push(h3("Material und Vorbereitung (Details: Materialliste)"));
d.push(...bullets([
  "Arbeitsblatt_Menschenrechte_Quelle_8D.pdf: eine A4-Seite doppelseitig (Seite 1: Quelle und Auswertungstabelle, Seite 2: Zusatz*, Urteil, Rückmeldung). 26-mal drucken, doppelseitig.",
  "Loesung_Arbeitsblatt_Menschenrechte_8D.pdf (für Sie): Musterantworten für beide Seiten (Stichworte genügen).",
  "Folien 1 bis 20 (Schrift mindestens 27 Punkt; Folie 1, 5, 15 und 18 mit der selbst gezeichneten Karikatur beziehungsweise den Bildern aus ihr).",
  "Seitentafel: Stundenfrage, Ziele, Ablauf (siehe Phase 0). Mitte der Tafel leer.",
  "Kreide, Grün zum Verbessern, Uhr oder Timer.",
]));

// Phase 0
d.push(phase("0", "Ankommen: Karikatur und Vermutung", 4, "AVIVA: Ankommen"));
d.push(warum("Klerus und Adel zählen ihr Entgegenkommen auf („Ihr habt geschworen. Ihr habt die Bastille gestürmt. Wir kamen euch entgegen.“) und fragen: „Was wollt ihr noch?“ Das ist die Stundenfrage. Die Vermutung im Heft wird am Ende mit der Antwort verglichen (Vorher-Nachher). Alle haben beim Ankommen etwas zu tun (direkt: Aufträge auf der Folie).", "„Das prüfen wir an einer Quelle. Zuerst: Was meinen Klerus und Adel?“"));
d.push(sp());
d.push(step("kommt an, schreibt leise", {
  say: "Stifte liegen. Ihr hört nur zu. Schaut auf die Folie und lest, was Klerus und Adel sagen. Schreibt ins Heft: Was will der Dritte Stand noch? Ich vermute …",
  ziel: "Rahmen (Stundenfrage), Vorher-Nachher für Ziel 1 bis 3.",
  do: [
    "Folie 1 steht beim Betreten da: selbst gezeichnete Karikatur (keine historische Quelle) mit Klerus, Adel und dem Dritten Stand (Bauern und Bürger), der zuhört. An der Wand hängen zwei unbeschriftete Bilder: der Eid im Ballhaus und der Sturm auf die Bastille. Auftrag: Heft raus, Datum und Überschrift, Vermutung („Was will der Dritte Stand noch? Ich vermute …“, ein bis zwei Sätze). Keine Besprechung.",
    "Folie 2 (Stundenfrage und Begriff „Quelle“, mit Relevanz: bei jeder Quelle und im Test, 45 Sekunden), Folie 3 (Lernziele, zugleich Raster der Rückmeldung, 45 Sekunden), Folie 4 (Ablauf, 30 Sekunden). Alle Folien: Sie sprechen, die Klasse hört zu.",
  ],
  extra: [boardBox("Seitentafel · Stundenfrage, Ziel und Ablauf", [
    bl("Was will der Dritte Stand noch? Was steht dazu in der Erklärung der Menschenrechte?", { after: 40 }),
    bl("Ich kann Artikel in eigenen Worten wiedergeben und den Begriff nennen.", { after: 20 }),
    bl("Ich kann einen Artikel mit früher vergleichen und einem Ereignis zuordnen.", { after: 20 }),
    bl("Ich kann mit Beleg und Maßstab begründen, welcher Artikel am wichtigsten ist.*", { after: 60 }),
    bl("A Karikatur · V Rückschau · I Quelle auswerten · V Blatt, Urteil · A Merksatz, Antwort", { after: 0 }),
  ]), gapP()],
}));

// Phase 1
d.push(phase("1", "Rückschau: Was meinen Klerus und Adel?", 5, "AVIVA: Vorwissen aktivieren"));
d.push(warum("Der Stoff der Stunde 3 wird echt wiederholt (allein, Partnercheck, Plenum), und zwar als Rückschau an zwei Bildern statt als reine Abfrage. Die Bilder sind unbeschriftet, die Klasse entschlüsselt „geschworen“ und „gestürmt“ selbst und verbindet sie mit dem Entgegenkommen (Ursache und Folge). Wer in Stunde 3 etwas nicht geschafft hat, sieht es hier noch einmal.", "„Die Bilder haben mit dem Entgegenkommen zu tun. Wie genau, seht ihr gleich.“"));
d.push(sp());
d.push(step("schreibt allein, vergleicht zu zweit", {
  say: "Stifte liegen. Ihr hört nur zu. Jetzt arbeitet ihr allein. Die Aufgabe steht rechts. … Flüsterstufe. Nur dein Partner hört dich. Vergleicht.",
  ziel: "Ziel 2 (Ereignisse von 1789 kennen).",
  do: [
    "Folie 5 (4 Minuten): Zu jedem Bild ins Heft: Was geschah? Was bewirkte es? (2 Minuten allein), Partnercheck („Was haben die Bilder mit dem Entgegenkommen zu tun?“, 1 Minute), Plenum: Stichproben bei 3 bis 4 Personen unterschiedlichen Niveaus (30 Sekunden, kurze Antwort, unbenotet).",
    "Folie 6, erster Teil (1 Minute): Kette 20. Juni Ballhausschwur (Versammlung bleibt zusammen) → 14. Juli Bastille (König zieht die Soldaten ab) → Sommer: Bauern greifen Schlösser an. „Vergleicht und verbessert mit Grün.“",
  ],
  sol: "Erwartung: Bild links: Eid im Ballhaus am 20. Juni, die Abgeordneten wollen nicht auseinandergehen, die Versammlung bleibt (gewaltlos). Bild rechts: Sturm auf die Bastille am 14. Juli, Gewalt, der König zieht die Soldaten ab, Aufstände im Land. Zusammenhang: Die Ereignisse machten Klerus und Adel Angst, das Entgegenkommen war nicht freiwillig.",
}));

// Phase 2
d.push(phase("2", "Quelle auswerten: 4./26. August, drei Schritte, Modell", 8, "AVIVA: Informieren"));
d.push(warum("Direktes Vorgehen, weil das Vorwissen zur Quellenarbeit gering ist. Zwei Portionen: erst der historische Rahmen (4. und 26. August, 2 Minuten), dann die Methode mit Modell. Lehrkraft denkt laut an Artikel 1, die Klasse sieht, wie die Schritte zusammengehören. Prüffrage mit Stichprobe, eine Minute Reserve zum erneuten Erklären. Es wird nicht weitergegangen, wenn etwas nicht verstanden ist.", "„Entgegenkommen ist Gnade von oben. Was ist der Gegenpol? Das lest ihr jetzt in der Quelle. Auf dem Blatt steht dasselbe in der Tabelle: gelb ist mein Beispiel.“"));
d.push(sp());
d.push(step("hört zu, antwortet in Stichproben", {
  say: "Stifte liegen. Ihr hört nur zu. … Ich mache es vor, ihr schaut zu.",
  ziel: "Ziel 2 (Rahmen, Einordnen) und Ziel 1 und 2 (Methode).",
  do: [
    "Portion 1 (Folie 6, zweiter Teil, 2 Minuten): In der Nacht zum 4. August 1789 beschließt die Nationalversammlung, die Vorrechte abzuschaffen (zum Beispiel die Steuerfreiheit von Klerus und Adel, Abgaben und Dienste der Bauern an den Adel). Klerus und Adel kommen entgegen, weil sie Angst vor den Aufständen haben: nicht freiwillig. Am 26. August folgt die Erklärung der Menschen- und Bürgerrechte, Vorbild ist die amerikanische Unabhängigkeitserklärung von 1776. Warnung vor dem Fehler: Abgeschafft heißt nicht, dass sofort alle gleich leben (viele Abgaben müssen Bauern noch ablösen); nicht der König, sondern die Nationalversammlung schaffte die Vorrechte ab. Prüffrage (Stichprobe, 2 Personen): „Warum kommen Klerus und Adel am 4. August entgegen?“ Erwartung: aus Angst vor den Aufständen.",
    "Portion 2, Teil 1 (Folie 7, 1 Minute): Drei Schritte, immer in dieser Reihenfolge: Lesen (in eigenen Worten), Einordnen (Was war vorher anders? Passt ein Bild?), Begriff (Wortspeicher). Warnung vor dem Fehler: nicht abschreiben, in eigene Worte fassen.",
    "Portion 2, Teil 2 (Folie 8, 3 Minuten): Sie denken laut an Artikel 1: „Was steht da? Alle Menschen haben von Geburt an dieselben Rechte. Was war vorher anders? Die Geburt entschied über den Stand. Wer gewinnt? Der Dritte Stand. Welcher Begriff? Gleichheit.“ Warnung vor dem Fehler: Gleich an Rechten heißt nicht gleich reich.",
    "Prüffrage (Folie 9, 1 Minute) und Reserve (1 Minute): „Was machst du im Schritt 2?“ und „Warum schreibst du nicht einfach ab?“ Stichprobe bei 4 bis 6 Personen unterschiedlichen Niveaus (unbenotet, getrennt von den mündlichen Noten). Bei Unsicherheit Schritt 2 am Artikel 1 wiederholen (Reserve). Dann „Jetzt macht ihr es selbst.“",
  ],
  sol: "Erwartung zur Prüffrage: Ich frage, was vorher anders war, wer gewinnt und ob ein Bild passt. / Ich muss es in eigene Worte fassen, so zeigt sich, ob ich es verstanden habe.",
}));

// Phase 3
d.push(phase("3", "Arbeitsblatt: angeleitet, selbstständig, Zuordnung, Urteil", 20, "AVIVA: Verarbeiten"));
d.push(warum("Progression der direkten Instruktion: erste Anwendung (Artikel 3 mit Hilfe), selbstständiges Üben (Artikel 2, 4, 6 ohne Hilfe), Partnercheck mit Zuordnung (Bild und Artikel), Kontrolle mit Aha („Was sie taten, wird zum Recht“), dann das Urteil als Sachurteil mit Beleg und Werturteil mit offengelegtem Maßstab. Artikel 2 und 4 verknüpfen mit Stunde 3 (Widerstand, Freiheit endet, wo sie anderen schadet).", "„Welcher Artikel ist für den Dritten Stand am wichtigsten? Das entscheidet ihr jetzt.“"));
d.push(sp());
d.push(step("3a  Artikel 3 mit Hilfe (3 Min.)", {
  say: "Stifte liegen. Ihr hört nur zu. Jetzt nehmt ihr das Arbeitsblatt. Jetzt arbeitet ihr allein. Die Aufgabe steht rechts.",
  ziel: "Ziel 1 und 2 (erste Anwendung).",
  do: [
    "Folie 10: Arbeitsblatt Seite 1, Artikel 3 lesen, „Lesen“ und „Einordnen“ mit den Hilfen ausfüllen (Begriff steht schon). Nach 2 Minuten: „Flüsterstufe. Nur dein Partner hört dich. Vergleicht.“ Folie 11 (30 Sekunden): Lösung, „Vergleicht und verbessert mit Grün.“ Stichprobe bei 2 Personen.",
  ],
  sol: "Lesen: Die Herrschaft geht vom Volk aus, nicht vom König allein. Einordnen: Vorher bestimmte der König allein, passt zum Ballhausschwur. Begriff: Volkssouveränität.",
}));
d.push(sp());
d.push(step("3b  Artikel 2, 4, 6 selbst (6 Min.)", {
  say: "Stifte liegen. Jetzt arbeitet ihr allein. Die Aufgabe steht rechts.",
  ziel: "Ziel 1 und 2 (selbstständiges Üben).",
  do: [
    "Folie 12: Selbstständiges Üben ohne Hilfen (Wortspeicher bleibt). Durch die Reihen gehen (Heft-Stichprobe bei 4 bis 6 Personen unterschiedlichen Niveaus, mitlesen, nicht kommentieren). Differenzierung: Wer unsicher ist, beginnt mit Artikel 4. Wer früh fertig ist, bearbeitet das Zusatz* (Artikel 10 und 11).",
  ],
}));
d.push(sp());
d.push(step("3c  Partnercheck und Zuordnung (4 Min.)", {
  say: "Stifte liegen. Flüsterstufe. Nur dein Partner hört dich. Vergleicht Artikel 2, 4 und 6. Dann ordnet ihr zu: Welches Bild passt zu welchem Artikel? Begründet mit einem Wort aus dem Artikel.",
  ziel: "Ziel 2 (einem Ereignis von 1789 zuordnen).",
  do: [
    "Folie 13: Jede Person liest vor, Unterschiede klären, mit Grün verbessern. Dann ordnen die Paare die zwei Bilder (aus Folie 5) Artikeln zu und begründen mit einem Wort aus dem Artikel. Sie hören bei 4 bis 6 Paaren zu, ohne zu helfen. Sollbruchstelle: Reicht die Zeit nicht, genügt eine Zuordnung.",
  ],
}));
d.push(sp());
d.push(step("3d  Kontrolle und Aha (3 Min.)", {
  say: "Stifte liegen. Vergleicht und verbessert mit Grün. … Ihr hört nur zu. Vergleicht mit eurer Zuordnung.",
  ziel: "Ziel 1 und 2 (Kontrolle), Aha für Ziel 2 und Vorbereitung auf Ziel 3.",
  do: [
    "Folie 14 (2 Minuten): Lösung der Tabelle (Stichworte, andere richtige Formulierungen zählen). Stichprobe bei 2 Personen.",
    "Folie 15 (1 Minute, Aha): Ballhaus passt zu Artikel 3 (Herrschaft geht vom Volk aus), Bastille zu Artikel 2 (Recht auf Widerstand). Artikel 2 nennt aber nicht Gewalt; Artikel 4 begrenzt: Die Freiheit endet, wo sie anderen schadet. Haltung: Wir beschreiben, was dort steht, wir erklären nicht, dass Gewalt gut ist (wirksam heißt nicht gerechtfertigt, Stunde 3). Was sie taten, steht ab dem 26. August als Recht in der Erklärung.",
  ],
  sol: "Art. 2: Der Staat muss die Rechte schützen, auch das Recht auf Widerstand. Vorher: Ungehorsam war Verbrechen. Grundrechte. Art. 4: Jeder darf tun, was anderen nicht schadet. Vorher: Einsperren ohne Urteil. Freiheit. Art. 6: Gleiches Gesetz für alle. Vorher: Steuerfreiheit, Ämter nur für Adel. Gleichheit.",
}));
d.push(sp());
d.push(step("3e  urteilt allein (4 Min.)", {
  say: "Stifte liegen. Jetzt arbeitet ihr allein. Die Aufgabe steht rechts.",
  ziel: "Ziel 3 (Beleg und Maßstab).",
  do: [
    "Folie 16: Wähle einen Artikel, der für den Dritten Stand am wichtigsten ist; auf Seite 2 mit den Satzanfängen: Beleg („Artikel … sagt: …“), Wirkung („Das ändert für den Dritten Stand, dass …“), Maßstab („Mein Maßstab: Wichtig ist für mich, was …“, mögliche Bausteine: die meisten Menschen betrifft, am meisten verändert, vor Willkür und Gewalt schützt), Urteil („Ich halte Artikel … für am wichtigsten, weil …“), Gegenargument („Dagegen spricht, dass …“, zum Beispiel aus Sicht von Klerus und Adel). Kein richtiges Urteil. Stichprobe bei 4 bis 6 Personen (unbenotet). Wer das Blatt abgeben möchte, legt es nach der Stunde vorn ab.",
  ],
  sol: "Mögliche Urteile (Auswahl, keine Musterlösung). Artikel 6 (gleiches Gesetz für alle: keine Steuerfreiheit mehr; Maßstab: betrifft die meisten). Artikel 3 (Volk bestimmt; Maßstab: ändert am meisten). Artikel 4 (kein Einsperren ohne Urteil; Maßstab: schützt vor Willkür). Artikel 2 (Widerstand ist ein Recht, ohne Gewalt zu nennen). Gegenargument: Auf dem Papier alles gleich, aber wer darf wählen? Oder aus Sicht von Klerus und Adel: Verlust der Vorrechte.",
}));

// Phase 4
d.push(phase("4", "Auswerten: Merksatz und Antwort als Dritter Stand", 8, "AVIVA: Auswerten"));
d.push(warum("Der Merksatz sichert das Ergebnis (Datum, Begriffe, Recht und Vorrecht) für den Test. Dann greift die Klasse die Aussage von Klerus und Adel vom Anfang auf und antwortet als Dritter Stand mit Beleg aus der Quelle (mehrere gute Antworten: Position, Beleg, Vergleich mit früher, „weil“). Das prüft alle drei Ziele. Der Vergleich mit der Vermutung zeigt den Lernzuwachs, ohne Floskel.", "„Gilt das für alle? Das sehen wir nächste Woche.“"));
d.push(sp());
d.push(step("schreibt ab, denkt, antwortet als Dritter Stand", {
  say: "Stifte liegen. Ihr hört nur zu. Jetzt schreibt ihr von der Folie ab: Merksatz 1. Rahmt ihn rot ein. … Das ist die letzte Aufgabe. Schaut noch einmal auf die Karikatur vom Anfang. Ihr seid der Dritte Stand. Denkt zehn Sekunden nach. Ich rufe euch auf.",
  ziel: "Ziel 1 und 2 (Merksatz), Ziel 1, 2 und 3 (Abschlussantwort).",
  do: [
    "Folie 17 (Merksatz, 2 Minuten): Abschreiben und rot umrahmen, Beispiel (Artikel 6) und Test („Recht gilt für alle, Vorrecht nur für wenige“) mit auf die Folie, kein „Wofür“. Gelenk: „Jetzt antwortet ihr Klerus und Adel.“",
    "Folie 18 (Abschlussfrage, 4 Minuten): Mündlich, in der Ich-Form, ohne Schnipsel: „Ich antworte euch: Wir wollen …, weil Artikel … sagt …, früher …“. Form nach Verlauf: Blitzlicht (4 bis 6 Personen unterschiedlichen Niveaus gezielt aufrufen, kein Kommentar) oder Duell zu zweit (leise: eine Person Klerus und Adel, eine Dritter Stand). Wer hat umgedacht? Vergleich mit der Vermutung vom Anfang.",
    "Folie 19 (Ziele prüfen, 1 Minute, still): Haben wir die Ziele erreicht? Wer das Blatt abgeben möchte, legt es vorn ab. Folie 20 (Ausblick, nur lesen, Brücke zu Stunde 5): „Für wen gelten diese Rechte wirklich?“",
  ],
  sol: "Mögliche Antworten als Dritter Stand (Auswahl, keine Musterlösung). „Wir wollen keine Gnade, sondern Rechte, weil Artikel 6 sagt, dass das Gesetz für alle gleich ist, früher hatten Klerus und Adel Vorrechte.“ „Wir wollen mitbestimmen, weil Artikel 3 sagt, dass die Herrschaft vom Volk ausgeht, früher bestimmte der König allein.“ „Wir wollen frei sein, weil Artikel 4 sagt, dass man alles tun darf, was anderen nicht schadet, früher konnte der König jemanden ohne Urteil einsperren.“",
}));

d.push(h3("So sieht das Heft und das Blatt am Ende aus"));
d.push(table([2300, 7906], [
  ["Überschrift", "Datum · Überschrift zur Stundenfrage (z. B. „Die Menschenrechte“)"],
  ["Vermutung", "„Was will der Dritte Stand noch? Ich vermute …“ (Beginn der Stunde)"],
  ["Rückschau", "zu jedem Bild: Was geschah? Was bewirkte es? Verbesserungen grün"],
  ["Blatt Seite 1", "Tabelle: Artikel 1 (gelb), Artikel 3 mit Hilfe, Artikel 2, 4, 6 selbst; Verbesserungen grün"],
  ["Blatt Seite 2", "Zusatz* (Artikel 10 und 11), Urteil mit Beleg, Wirkung, Maßstab, Urteil und Gegenargument; bei Abgabe Rückmeldung nach Stufen"],
  ["Heft", "Merksatz 1 (abgeschrieben, rot umrahmt)"],
].map(([a, b]) => row([tc(a, 2300, { bold: true, fill: LIGHT }), tc(b, 7906)], 340))));

d.push(h3("Nach der Stunde festhalten (Reflexionslog)"));
d.push(...bullets(["Passten Ziel, Inhalt und Methode zusammen? Wie lief die Zeit? Wie viel echte Lernzeit hatte die Klasse?", "Haben die Schritte „Lesen, Einordnen, Begriff“ gesessen? Wo hakte es (Prüffrage, Artikel 3, selbstständige Phase, Zuordnung)?", "Wie viele Blätter wurden abgegeben, auf welcher Stufe lagen sie?", "Wie lief die Aktivierung (Hospitationsrückmeldung zum Kompetenzerwerb)? Was nimmst du mit?", "Wie lief Stunde 3 (Rückmeldung nachtragen), und wie lief die Rückschau an den Bildern?", "Kam die Pointe „Gnade oder Recht“ in der Abschlussantwort an? Gab es Antworten mit Beleg und Vergleich mit früher?"]));
d.push(h3("Zum Prüfen"));
d.push(...bullets([
  "Wortlaut der Artikel: gekürzte und vereinfachte Fassung, mit der Kollegin und dem Lehrbuch abgleichen (Artikel 3: im Original „Nation“, hier „Nation, also Volk“). Der Test sollte denselben Wortlaut nutzen.",
  "4. August: Aufzählung der abgeschafften Vorrechte und die Ablösepflicht für Bauern gegen das Lehrbuch lesen. Die Aussage „viele Abgaben müssen abgelöst werden“ ist allgemeines Wissen. Dass die Aufstände der Bauern ein wichtiger Grund für den Verzicht waren, ist die gängige Darstellung.",
  "Artikel 10 (Zusatz): „Vorher hatte der katholische Glaube Vorrang“ ist eine Vereinfachung (Protestanten waren benachteiligt). Artikel 11: Zensur durch den König. Beides steht nur in der Lösung und im Zusatz*.",
  "Die Karikatur (Klerus, Adel und Dritter Stand, zwei Bilder an der Wand) ist selbst gezeichnet, keine historische Quelle. Die Aussage der Figuren („Wir kamen euch entgegen.“) ist bewusst überspitzt, eine erfundene Selbstdarstellung. Historisch war der Verzicht am 4. August nicht freiwillig (Angst vor den Bauernaufständen). Die Klasse widerlegt die Aussage am Ende selbst.",
  "Zuordnung Bild und Artikel (Ballhaus zu Artikel 3, Bastille zu Artikel 2) ist eine Deutung, die Klasse darf andere Zuordnungen begründen. Artikel 2 nennt das Recht auf Widerstand, nicht Gewalt.",
  "Die Notenfrage (Punkte, Schlüssel) ist offen: Die Tabelle liefert pro Artikel Felder, das Urteil gibt eine Stufe. Das können wir später ergänzen.",
]));

(async () => {
  const doc = new Document({ styles: { default: { document: { run: { font: "Calibri", size: 22 } } } }, sections: [{ properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 1000, bottom: 850, left: 850, right: 850 } } }, headers: { default: header }, footers: { default: footer }, children: d }] });
  fs.writeFileSync("../Tafelskript_Stunde4_Menschenrechte_8D.docx", await Packer.toBuffer(doc));
  console.log("ok");
})();
