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
d.push(kicker("Tafelskript für die Lehrkraft · Einzelstunde (45 Min.) · 24 Schülerinnen und Schüler · eine Lehrkraft · Mittwoch 25.11.2026 · Hospitation"), h1("Vorrechte weg, Menschenrechte da: Quelle ins Heute übersetzen"));
d.push(p(t("Ziel: Die Stunde knüpft an die Wege des Widerstands (Stunde 3) an und ist nach dem Partnerpuzzle bewusst lehrergesteuert mit Einzelarbeit. In der Karikatur zählen Klerus und Adel ihr Entgegenkommen auf („Ihr habt geschworen. Ihr habt die Bastille gestürmt. Wir kamen euch entgegen. Was wollt ihr noch?“), zwei Bilder an der Wand (Ballhaus, Bastille) liefern die Rückschau. Die Klasse erfährt, was am 4. August (Verzicht auf die Vorrechte) und am 26. August (Erklärung der Menschenrechte) geschah, übersetzt die Artikel in heutige Begriffe (Gleichheit vor dem Gesetz, Eigentumsrecht, Meinungsfreiheit, Volkssouveränität) und stellt fest, was sich für den Dritten Stand (Bauern und Bürger) verbessert hat. Im Aha-Moment „1789 oder heute?“ zeigt sich, dass vieles fast wörtlich im Grundgesetz steht (Vorbildwirkung). Zum Schluss sichert die Klasse das Ergebnis (Merksatz, Abschlussaufgabe als Dritter Stand).", { size: 21 }), { after: 100 }));

d.push(h3("Kopf der Planung"));
d.push(table([2600, 7606], [
  ["Ziel (Ich kann …)", "1  … erklären, was sich durch den Verzicht auf die Vorrechte und die Menschenrechte für den Dritten Stand verbessert hat.  2  … Artikel von 1789 in heutige Begriffe übersetzen (zum Beispiel Eigentumsrecht).  3  … an einem Grundgesetz-Artikel zeigen, dass die Erklärung bis heute nachwirkt.*"],
  ["Prüfaufgabe: Woran erkenne ich es?", "Arbeitsblatt Seite 1 (Begriff heute zu allen Artikeln, Vorher bei 6, 11 und 17), Seite 2 (Aufgabe A „1789 oder heute?“, Aufgabe B „Historische Bedeutung“ mit Satzanfängen), Abschlussaufgabe: schriftlich und mündlich als Dritter Stand („Wir wollen …, weil Artikel … sagt. Früher …, heute …“). Freiwillige Abgabe, Rückmeldung nach drei Stufen (gleich den Zielen). Punkte oder Noten gibt es heute nicht (später)."],
  ["Mindestziel", "Ziel 1 und 2. Ziel 3 hat einen Stern (Transfer gekürzt). „Auswerten“ bleibt immer."],
  ["Beitrag zur Reihe", "Stunde 4 von 11 (Sequenz, Fassung vom 17.11.). Sie zeigt, wie der Verzicht auf die Vorrechte und die Erklärung der Menschenrechte die Lage des Dritten Standes verbessern, baut die Kompetenz Quellenauswertung auf (wiederkehrend, auch im Test) und führt mit „Für wen galten die Rechte 1789 wirklich?“ zur Verfassung 1791 (Stunde 5)."],
  ["Vorgehen", "Rückschau indirekt (allein, Plenum), Informieren und Quellenauswertung direkt: Portionen, Vormachen (Lehrkraft denkt laut), angeleitet (Artikel 3 mit Hilfe), selbstständig (sechs Artikel). Direktes Vorgehen, weil das Vorwissen zur Quellenarbeit gering ist. Aha und Sicherung geführt. Direkt oder indirekt wird je Phase entschieden."],
].map(([a, b]) => row([tc(a, 2600, { fill: LIGHT, bold: true }), tc(b, 7606)], 340))));

d.push(h3("Bezug jedes Schritts zum Ziel (Prüfung)"));
d.push(table([4400, 1500, 4306], [
  row([th("Schritt", 4400), th("Ziel", 1500), th("Begründung", 4306)], 320),
  ...[
    ["Karikatur und Vermutung (Folie 1)", "Rahmen", "Stundenfrage, Vorher-Nachher mit der Abschlussaufgabe"],
    ["Rückschau an den Bildern (Folie 5, 6)", "1", "Ereignisse 1789 als Voraussetzung für „vorher, nachher“"],
    ["Verzicht auf die Vorrechte (Folie 7)", "1", "Vorher und Nachher, Bedeutung des 4. August"],
    ["Quelle ins Heute übersetzen, Modell, Prüffrage (Folie 8 bis 10)", "2", "Methode modellieren, Verständnis prüfen"],
    ["Artikel 3 mit Hilfe, sechs Artikel selbst (Folie 11 bis 14)", "1, 2", "angeleitet, dann selbstständig üben; Begriff heute und Vorher"],
    ["Was hat sich verbessert? (Folie 15)", "1", "Vorher und Seitdem (Bedeutung)"],
    ["1789 oder heute? Echo (Folie 16 bis 18)", "3", "Transfer, Aha: Vorbildwirkung im Grundgesetz"],
    ["Merksatz (Folie 19)", "1, 2, 3", "Ergebnissicherung, Wissen für den Test"],
    ["Abschlussaufgabe (Folie 20)", "1, 2, 3", "prüft alle Ziele, Vergleich mit der Vermutung"],
    ["Ziele prüfen (Folie 21)", "1, 2, 3", "Zielprüfung, Abgabe"],
    ["Ausblick (Folie 22)", "Brücke", "kein Ziel der Stunde, bereitet Stunde 5 vor"],
  ].map((r) => row(r.map((v, i) => tc(v, [4400, 1500, 4306][i], { bold: i === 1 })), 320)),
]));

d.push(h3("Was aus den bisherigen Stunden und Rückmeldungen übernommen wird"));
d.push(...bullets([
  "Rückmeldung zur Hospitation (Lehrkraft, 21.11.): Nach dem Gruppenpuzzle ist eine eher lehrergesteuerte Stunde mit Einzelarbeit sinnvoll (kein ständiges Sozialformfeuerwerk). Die Sicherung am Ende ist wichtig. Die Artikel werden in heutiges Vokabular übersetzt (Gleichheit vor dem Gesetz, Eigentumsrecht, Meinungsfreiheit), um Kontinuität und Vorbildwirkung seit 1789 zu zeigen; der Transfer zu ausgewählten Grundgesetz-Artikeln soll der Aha-Moment werden; die Bedeutung (was wurde verbessert?) soll deutlich werden; das Blatt der Kollegin dient als Vorbild.",
  "Stunde 3 (18.11.): Ballhausschwur und Bastille als Widerstand, wirksam heißt nicht gerechtfertigt. Die Rückschau (Folie 5 und 6) wiederholt das an zwei Bildern. Falls in Stunde 3 etwas nicht geschafft wurde, hier als Wiederholung aufgreifen.",
  "8C (20.11.): Die Stunde 3 hat in der 8C gut funktioniert (Rückmeldung der Lehrkraft, 21.11.). Für die 8C entfällt diese Stunde (die Menschenrechte hatte sie am 13.11.).",
  "Entscheidungen der Lehrkraft (21.11.): Artikel wie im Original (Artikel 6, nicht 5 wie in einem anderen Buch), Wortlaut in der gekürzten Fassung, Eigentumsrecht übernehmen (Artikel 17), Transfer als Aha-Moment („1789 oder heute?“), Blatt der Kollegin anpassen (eigene Fassung mit dem Wortspeicher der Kollegin plus Eigentumsrecht). Das Buch mit den Fotos ist ein anderes Buch und nur Inspiration. Die Nummerierung des Buches wird nicht übernommen.",
  "Karikatur als Einstieg, Rückschau und Abschluss (Antwort in der Rolle des Dritten Stands): selbst gezeichnet, keine historische Quelle. Die Aussage ist bewusst überspitzt (Entgegenkommen als Gnade, historisch nicht freiwillig).",
  "Regel „höchstens drei Methoden“: Quellenauswertung (Vormachen, angeleitet, selbstständig), Aha „1789 oder heute?“, Abschlussaufgabe. Karikatur und Rückschau sind Impuls beziehungsweise Routine.",
]));

d.push(h3("Zeitplan (45 Minuten)"));
d.push(table([800, 2300, 4300, 2806], [
  row([th("Min.", 800), th("AVIVA", 2300), th("Phase", 4300), th("Arbeitsform", 2806)], 320),
  ...[["4", "Ankommen", "0  Karikatur, Vermutung, Stundenfrage, Ziele, Ablauf", "still, allein"], ["4", "Vorwissen", "1  Rückschau an den Bildern (3), Kette (1)", "allein, Plenum"], ["10", "Informieren", "2  Verzicht auf die Vorrechte (3), Erklärung und drei Schritte (1), Modell Artikel 1 (3), Prüffrage (1), Reserve (2)", "hört zu, Stichprobe"], ["18", "Verarbeiten", "3  Artikel 3 mit Hilfe (3), sechs Artikel allein (6), Kontrolle Begriffe und Bilanz (3), Aha „1789 oder heute?“ (4), Reserve (2)", "allein, Plenum"], ["9", "Auswerten", "4  Merksatz (2), Abschlussaufgabe (5), Ziele (1), Ausblick (1)", "schreiben, mündlich, still"]].map((r) => row(r.map((v, i) => tc(v, [800, 2300, 4300, 2806][i], { bold: i === 0 })), 340)),
]));
d.push(note("ZEIT", "Die Stunde ist dicht. Sollbruchstellen: (1) Rückschau: Plenum nur 30 Sekunden. (2) Reserve in Phase 2 (2 Minuten) für erneutes Erklären. (3) Übung 2: Wer nach 6 Minuten nur den Begriff bei allen Artikeln hat, ist richtig; „Vorher“ nur bei Artikel 6. (4) Aha: Zeigen Sie nur die ersten zwei Paare, wenn die Zeit knapp ist (Ziel 3 mit Stern). (5) Abschlussaufgabe: nur mündlich, ohne 2 Minuten Schreiben. Phase 4 „Auswerten“ wird nicht gestrichen. Die Übungsphase endet 9 Minuten vor Schluss."));

d.push(h3("Worauf die Hospitation achten kann (Aktivierung, Kompetenz, Sicherung)"));
d.push(...bullets([
  "Lehrergesteuert mit Einzelarbeit: Portionen, Modell, angeleitet, selbstständig, Stichproben. Partnerarbeit nur kurz (Vergleichen nach Artikel 3).",
  "Alle in jeder Phase aktiv: Ankommen (Vermutung im Heft), Rückschau (allein schreiben), Prüffrage (Stichprobe), Übungen (allein), Aha (allein ankreuzen), Abschluss (alle schreiben, wenige sprechen).",
  "Kompetenzaufbau sichtbar: gleiche drei Schritte (Lesen, Begriff heute, Vorher) in Modell, angeleiteter und selbstständiger Phase; Hilfen werden abgebaut.",
  "Kognitive Aktivierung: Entschlüsseln der Sprechblase an den Bildern, Übersetzen ins heutige Vokabular, Raten „1789 oder heute?“ und das Aha.",
  "Sicherung: Vorher-Nachher-Bilanz, Merksatz, schriftliche und mündliche Abschlussaufgabe, Zielprüfung, Vergleich mit der Vermutung.",
  "Differenzierung: Wortspeicher, Hilfe bei Artikel 3, kurze Artikel zuerst, Zusatz* (Vorher bei 2, 4, 10), Satzanfänge bei Aufgabe B.",
]));

d.push(h3("Lautstärke und Ansagen (wörtlich)"));
d.push(table([4700, 5506], [
  ["„Stifte liegen. Ihr hört nur zu.“", "Sie erklären, niemand schreibt."],
  ["„Jetzt arbeitet ihr allein. Die Aufgabe steht rechts.“", "Einzelarbeit, leise."],
  ["„Flüsterstufe. Nur dein Partner hört dich.“", "kurzer Vergleich nach Artikel 3."],
  ["„Vergleicht und verbessert mit Grün.“", "Kontrolle im eigenen Blatt."],
  ["„Es ist zu laut. Ab jetzt arbeitet ihr allein und schreibt.“", "Wenn es laut wird, keine Diskussion."],
].map(([a, b]) => row([tc(a, 4700, { fill: LIGHT, bold: true }), tc(b, 5506)], 340))));

d.push(h3("Material und Vorbereitung (Details: Materialliste)"));
d.push(...bullets([
  "Arbeitsblatt_Menschenrechte_Quelle_8D.pdf: eine A4-Seite doppelseitig (Seite 1: Quelle, Begriff heute, Vorher; Seite 2: „1789 oder heute?“, Historische Bedeutung, Rückmeldung). 26-mal drucken, doppelseitig.",
  "Loesung_Arbeitsblatt_Menschenrechte_8D.pdf (für Sie): Musterantworten für beide Seiten (Stichworte genügen).",
  "Folien 1 bis 22 (Schrift mindestens 27 Punkt; Folie 1, 5, 6 und 20 mit der selbst gezeichneten Karikatur beziehungsweise den Bildern aus ihr).",
  "Seitentafel: Stundenfrage, Ziele, Ablauf (siehe Phase 0). Mitte der Tafel leer.",
  "Kreide, Grün zum Verbessern, Uhr oder Timer.",
]));

// Phase 0
d.push(phase("0", "Ankommen: Karikatur und Vermutung", 4, "AVIVA: Ankommen"));
d.push(warum("Klerus und Adel zählen ihr Entgegenkommen auf („Ihr habt geschworen. Ihr habt die Bastille gestürmt. Wir kamen euch entgegen.“) und fragen: „Was wollt ihr noch?“ Das ist die Stundenfrage. Die Vermutung im Heft wird am Ende mit der Abschlussaufgabe verglichen (Vorher-Nachher). Alle haben beim Ankommen etwas zu tun (direkt: Aufträge auf der Folie).", "„Was haben Ballhaus und Bastille mit dem Entgegenkommen zu tun?“"));
d.push(sp());
d.push(step("kommt an, schreibt leise", {
  say: "Stifte liegen. Ihr hört nur zu. Schaut auf die Folie und lest, was Klerus und Adel sagen. Schreibt ins Heft: Was will der Dritte Stand noch? Ich vermute …",
  ziel: "Rahmen (Stundenfrage), Vorher-Nachher für Ziel 1 bis 3.",
  do: [
    "Folie 1 steht beim Betreten da: selbst gezeichnete Karikatur (keine historische Quelle) mit Klerus, Adel und dem Dritten Stand (Bauern und Bürger), der zuhört. An der Wand hängen zwei unbeschriftete Bilder: der Eid im Ballhaus und der Sturm auf die Bastille. Auftrag: Heft raus, Datum und Überschrift, Vermutung („Was will der Dritte Stand noch? Ich vermute …“, ein bis zwei Sätze). Keine Besprechung.",
    "Folie 2 (Stundenfrage und Begriff „Quelle“, 45 Sekunden), Folie 3 (Lernziele, zugleich Raster der Rückmeldung, 45 Sekunden), Folie 4 (Ablauf, 30 Sekunden). Alle Folien: Sie sprechen, die Klasse hört zu.",
  ],
  extra: [boardBox("Seitentafel · Stundenfrage, Ziel und Ablauf", [
    bl("Was will der Dritte Stand noch? Was ändern die Abschaffung der Vorrechte und die Menschenrechte?", { after: 40 }),
    bl("Ich kann erklären, was sich für den Dritten Stand verbessert hat.", { after: 20 }),
    bl("Ich kann Artikel von 1789 in heutige Begriffe übersetzen.", { after: 20 }),
    bl("Ich kann an einem Grundgesetz-Artikel zeigen, dass die Erklärung nachwirkt.*", { after: 60 }),
    bl("A Karikatur · V Rückschau · I Vorrechte, Modell · V Blatt, Aha · A Merksatz, Antwort", { after: 0 }),
  ]), gapP()],
}));

// Phase 1
d.push(phase("1", "Rückschau: Was meinen Klerus und Adel?", 4, "AVIVA: Vorwissen aktivieren"));
d.push(warum("Der Stoff der Stunde 3 wird echt wiederholt (allein, Plenum), und zwar als Rückschau an zwei Bildern statt als reine Abfrage. Die Bilder sind unbeschriftet, die Klasse entschlüsselt „geschworen“ und „gestürmt“ selbst und verbindet sie mit dem Entgegenkommen (Ursache und Folge). Kein Partnercheck, weil die Stunde lehrergesteuert ist. Wer in Stunde 3 etwas nicht geschafft hat, sieht es hier noch einmal.", "„Klerus und Adel kamen entgegen, weil sie Angst hatten. Was genau haben sie aufgegeben? Hört zu.“"));
d.push(sp());
d.push(step("schreibt allein, antwortet in Stichproben", {
  say: "Stifte liegen. Ihr hört nur zu. Jetzt arbeitet ihr allein. Die Aufgabe steht rechts.",
  ziel: "Ziel 1 (Ereignisse 1789 als Voraussetzung).",
  do: [
    "Folie 5 (3 Minuten): Zu jedem Bild ins Heft: Was geschah? Was bewirkte es? (2 Minuten allein), Plenum: Stichproben bei 3 bis 4 Personen unterschiedlichen Niveaus (1 Minute, kurze Antwort, unbenotet).",
    "Folie 6 (1 Minute): Kette 20. Juni Ballhausschwur (Versammlung bleibt zusammen) → 14. Juli Bastille (König zieht die Soldaten ab) → Sommer: Bauern greifen Schlösser an; die Zeilen 4. und 26. August nur nennen. „Vergleicht und verbessert mit Grün.“",
  ],
  sol: "Erwartung: Bild links: Eid im Ballhaus am 20. Juni, die Abgeordneten wollen nicht auseinandergehen (gewaltlos). Bild rechts: Sturm auf die Bastille am 14. Juli, Gewalt, der König zieht die Soldaten ab, Aufstände im Land. Zusammenhang: Die Ereignisse machten Klerus und Adel Angst, das Entgegenkommen war nicht freiwillig.",
}));

// Phase 2
d.push(phase("2", "Vorrechte, Erklärung, Quelle ins Heute übersetzen", 10, "AVIVA: Informieren"));
d.push(warum("Direktes Vorgehen, weil das Vorwissen zur Quellenarbeit gering ist. Portion 1: Verzicht auf die Vorrechte (4. August, vorher und nachher). Portion 2: Erklärung vom 26. August und die Methode (Lesen, Begriff heute, Vorher), Modell an Artikel 1. Prüffrage mit Stichprobe, zwei Minuten Reserve zum erneuten Erklären. Es wird nicht weitergegangen, wenn etwas nicht verstanden ist.", "„Entgegenkommen ist Gnade von oben. Was ist der Gegenpol? Das lest ihr jetzt in der Quelle. Auf dem Blatt steht dasselbe in der Tabelle: gelb ist mein Beispiel.“"));
d.push(sp());
d.push(step("hört zu, antwortet in Stichproben", {
  say: "Stifte liegen. Ihr hört nur zu. … Ich mache es vor, ihr schaut zu.",
  ziel: "Ziel 1 (vorher, nachher) und Ziel 2 (Methode, Begriff heute).",
  do: [
    "Portion 1 (Folie 7, 3 Minuten): In der Nacht zum 4. August 1789 beschließt die Nationalversammlung, die Vorrechte abzuschaffen. Klerus und Adel geben nach, weil sie Angst vor den Aufständen der Bauern haben: nicht freiwillig, Gnade von oben. Vorher: Klerus und Adel zahlen keine Steuern, die Bauern leisten Abgaben und Dienste an den Adel, Ämter und Jagd sind Vorrechte. Nachher: Die Steuerfreiheit fällt, viele Abgaben fallen weg oder werden abgelöst. Warnung vor dem Fehler: Abgeschafft heißt nicht, dass sofort alle gleich leben (viele Abgaben müssen Bauern noch ablösen); nicht der König, sondern die Nationalversammlung beschloss es. Prüffrage (Stichprobe, 2 Personen): „Warum geben Klerus und Adel nach?“ Erwartung: aus Angst vor den Aufständen.",
    "Portion 2, Teil 1 (Folie 8, 1 Minute): Am 26. August beschließt die Nationalversammlung die Erklärung der Menschen- und Bürgerrechte, Vorbild ist die amerikanische Unabhängigkeitserklärung von 1776. Drei Schritte, immer in dieser Reihenfolge: Lesen, Begriff heute (Wortspeicher), Vorher (bei Artikel 6, 11 und 17). Warnung vor dem Fehler: nicht abschreiben, den heutigen Begriff suchen.",
    "Portion 2, Teil 2 (Folie 9, 3 Minuten): Sie denken laut an Artikel 1: „Was steht da? Alle Menschen haben von Geburt an dieselben Rechte. Welches Wort sagen wir heute? Gleichheit. Was war vorher anders? Die Geburt entschied über den Stand. Wer gewinnt? Der Dritte Stand.“ Warnung vor dem Fehler: Gleich an Rechten heißt nicht gleich reich.",
    "Prüffrage (Folie 10, 1 Minute) und Reserve (2 Minuten): „Was machst du im Schritt 2 Begriff heute?“ und „Warum steht Gleichheit nicht wörtlich im Artikel?“ Stichprobe bei 4 bis 6 Personen unterschiedlichen Niveaus (unbenotet, getrennt von den mündlichen Noten). Bei Unsicherheit Schritt 2 am Artikel 1 wiederholen (Reserve). Dann „Jetzt macht ihr es selbst.“",
  ],
  sol: "Erwartung zur Prüffrage: Ich suche das Wort, das wir heute dafür sagen. / Die Erklärung ist von 1789, wir sagen heute ein anderes Wort für dieselbe Sache.",
}));

// Phase 3
d.push(phase("3", "Arbeitsblatt: angeleitet, selbstständig, Bilanz, Aha", 18, "AVIVA: Verarbeiten"));
d.push(warum("Progression der direkten Instruktion in Einzelarbeit: erste Anwendung (Artikel 3 mit Hilfe), selbstständiges Üben (sechs Artikel), Kontrolle mit Bilanz (Was hat sich verbessert?), dann der Aha-Moment „1789 oder heute?“. Die Aufgabe A auf Seite 2 sagt nicht, dass es Paare gibt: Das Echo im Grundgesetz soll überraschen.", "„Das war vor 237 Jahren. Was davon gilt noch heute? Dreht Seite 2 um.“"));
d.push(sp());
d.push(step("3a  Artikel 3 mit Hilfe (3 Min.)", {
  say: "Stifte liegen. Ihr hört nur zu. Jetzt nehmt ihr das Arbeitsblatt. Jetzt arbeitet ihr allein. Die Aufgabe steht rechts.",
  ziel: "Ziel 1 und 2 (erste Anwendung).",
  do: [
    "Folie 11: Arbeitsblatt Seite 1, Artikel 3 lesen, Begriff steht schon, bei „Vorher“ die Hilfe ergänzen. Nach 2 Minuten: „Flüsterstufe. Nur dein Partner hört dich. Vergleicht.“ Folie 12 (30 Sekunden): Lösung, „Vergleicht und verbessert mit Grün.“ Stichprobe bei 2 Personen.",
  ],
  sol: "Begriff: Volkssouveränität. Vorher: Der König bestimmte allein.",
}));
d.push(sp());
d.push(step("3b  Sechs Artikel selbst (6 Min.)", {
  say: "Stifte liegen. Jetzt arbeitet ihr allein. Die Aufgabe steht rechts.",
  ziel: "Ziel 1 und 2 (selbstständiges Üben).",
  do: [
    "Folie 13: Artikel 2, 4, 6, 10, 11 und 17: Begriff heute (Wortspeicher), bei 6, 11 und 17 „Vorher“. Durch die Reihen gehen (Heft-Stichprobe bei 4 bis 6 Personen unterschiedlichen Niveaus, mitlesen, nicht kommentieren). Differenzierung: Wer unsicher ist, beginnt mit Artikel 4 und 17. Wer früh fertig ist, bearbeitet das Zusatz* (Vorher bei 2, 4, 10).",
  ],
}));
d.push(sp());
d.push(step("3c  Kontrolle und Bilanz (3 Min.)", {
  say: "Stifte liegen. Vergleicht und verbessert mit Grün. … Ihr hört nur zu. Vergleicht mit euren Einträgen bei Artikel 6, 11 und 17.",
  ziel: "Ziel 2 (Kontrolle der Übersetzung), Ziel 1 (Bilanz: Was hat sich verbessert?).",
  do: [
    "Folie 14 (1,5 Minuten): Begriffe zu allen Artikeln. Stichprobe bei 2 Personen. Hinweis: Artikel 2 nennt das Recht auf Widerstand (Stunde 3), aber nicht Gewalt; Artikel 4 begrenzt die Freiheit (sie endet, wo sie anderen schadet): wirksam heißt nicht gerechtfertigt.",
    "Folie 15 (1,5 Minuten): Vorher und Seitdem bei Artikel 3, 6, 11 und 17. Bedeutung: gleiche Rechte statt Vorrechte nach Geburt, Schutz statt Willkür des Königs, Macht beim Volk. Haltung: ein großer Schritt auf dem Papier; ob die Rechte für alle galten, ist die Frage der nächsten Stunde.",
  ],
  sol: "Art. 1 Gleichheit, Art. 2 Grundrechte (auch Widerstand), Art. 3 Volkssouveränität, Art. 4 Freiheit, Art. 6 Gleichheit vor dem Gesetz, Art. 10 Religionsfreiheit, Art. 11 Meinungsfreiheit, Art. 17 Eigentumsrecht. Vorher: 3 der König bestimmt allein; 6 Vorrechte für Klerus und Adel, keine Steuern; 11 der König kontrolliert, was gedruckt wird; 17 Bauern geben Ernte und Dienste ab, Eigentum nicht gesichert.",
}));
d.push(sp());
d.push(step("3d  Aha: 1789 oder heute? (4 Min.)", {
  say: "Stifte liegen. Ihr hört nur zu. Jetzt dreht ihr Seite 2 um. Jetzt arbeitet ihr allein. Die Aufgabe steht rechts. … Stifte liegen. Ihr hört nur zu. Schaut auf eure Sätze A bis H.",
  ziel: "Ziel 3 (Transfer, Vorbildwirkung).",
  do: [
    "Folie 16 (2 Minuten allein): Aufgabe A auf Seite 2: acht Sätze ohne Jahr, vier aus der Erklärung von 1789 (Artikel 3, 6, 11, 17), vier aus dem Grundgesetz (Artikel 3, 5, 14, 20), gemischt. Ankreuzen: 1789 oder heute, raten ist erlaubt. Sagen Sie nicht, dass es Paare gibt.",
    "Folie 17 und 18 (je 1 Minute, Sie sprechen): Zu jedem Satz von 1789 das Echo im Grundgesetz, fast wörtlich. Aussage: „Was ihr eben als Recht aus 1789 gelesen habt, steht nach 237 Jahren fast wörtlich in unserem Grundgesetz. Die Erklärung war ein Vorbild.“ Haltung: Heute gelten die Rechte für alle, 1789 galten sie nicht für alle (nächste Stunde). Lassen Sie den Aha-Moment wirken, kein Gespräch.",
  ],
  sol: "Lösung Aufgabe A: A 1789 (Art. 6), B heute (GG Art. 20 Abs. 2), C 1789 (Art. 11), D heute (GG Art. 14 Abs. 1), E heute (GG Art. 3 Abs. 1), F 1789 (Art. 3), G 1789 (Art. 17), H heute (GG Art. 5 Abs. 1). Paare: A und E, C und H, F und B, G und D.",
}));

// Phase 4
d.push(phase("4", "Auswerten: Merksatz und Abschlussaufgabe", 9, "AVIVA: Auswerten"));
d.push(warum("Der Merksatz sichert das Ergebnis (Daten, heutige Begriffe, Recht statt Vorrecht, Echo im Grundgesetz) für den Test. Die Abschlussaufgabe prüft alle drei Ziele: Die Klasse antwortet als Dritter Stand schriftlich (Aufgabe B „Historische Bedeutung“) und mündlich mit Artikel, Vorher, Heute. Der Vergleich mit der Vermutung zeigt den Lernzuwachs, ohne Floskel.", "„Gilt das für alle? Das sehen wir nächste Woche.“"));
d.push(sp());
d.push(step("schreibt ab, schreibt, antwortet als Dritter Stand", {
  say: "Stifte liegen. Ihr hört nur zu. Jetzt schreibt ihr von der Folie ab: Merksatz 1. Rahmt ihn rot ein. … Schaut noch einmal auf die Karikatur vom Anfang. Ihr seid der Dritte Stand. Schreibt eure Antwort auf. Jetzt arbeitet ihr allein.",
  ziel: "Ziel 1, 2 und 3 (Merksatz, Abschlussaufgabe).",
  do: [
    "Folie 19 (Merksatz, 2 Minuten): Abschreiben und rot umrahmen, Beispiel (Artikel 6 und Grundgesetz Artikel 3) und Test („Recht gilt für alle, Vorrecht nur für wenige“) mit auf die Folie, kein „Wofür“.",
    "Folie 20 (Abschlussaufgabe, 5 Minuten): 2 Minuten still schreiben auf Seite 2, Aufgabe B „Historische Bedeutung“ (Vorher, Seit 1789, Heute, mit Satzanfängen). Dann 3 Minuten mündlich: Blitzlicht bei 4 bis 6 Personen unterschiedlichen Niveaus, in der Ich-Form: „Ich antworte euch: Wir wollen …, weil Artikel … sagt. Früher …, heute steht im Grundgesetz …“ Wer hat umgedacht? Vergleich mit der Vermutung vom Anfang. Bei Zeitnot: nur mündlich.",
    "Folie 21 (Ziele prüfen, 1 Minute, still): Haben wir die Ziele erreicht? Wer das Blatt abgeben möchte, legt es vorn ab. Folie 22 (Ausblick, nur lesen, Brücke zu Stunde 5): „Für wen galten diese Rechte 1789 wirklich?“",
  ],
  sol: "Mögliche Antworten als Dritter Stand (Auswahl, keine Musterlösung). „Wir wollen Gleichheit, weil Artikel 6 sagt, dass das Gesetz für alle gleich ist. Früher hatten Klerus und Adel Vorrechte, heute steht im Grundgesetz, dass alle Menschen vor dem Gesetz gleich sind.“ „Wir wollen frei reden, weil Artikel 11 das erlaubt. Früher kontrollierte der König den Druck, heute steht die Meinungsfreiheit im Grundgesetz Artikel 5.“ „Wir wollen mitbestimmen, weil Artikel 3 sagt, dass die Herrschaft vom Volk ausgeht, früher bestimmte der König allein, heute steht im Grundgesetz, dass alle Staatsgewalt vom Volke ausgeht.“",
}));

d.push(h3("So sieht das Heft und das Blatt am Ende aus"));
d.push(table([2300, 7906], [
  ["Überschrift", "Datum · Überschrift zur Stundenfrage (z. B. „Die Menschenrechte“)"],
  ["Vermutung", "„Was will der Dritte Stand noch? Ich vermute …“ (Beginn der Stunde)"],
  ["Rückschau", "zu jedem Bild: Was geschah? Was bewirkte es? Verbesserungen grün"],
  ["Blatt Seite 1", "Tabelle: Artikel 1 (gelb), Artikel 3 mit Hilfe, sechs Artikel selbst (Begriff heute, Vorher bei 6, 11, 17); Verbesserungen grün"],
  ["Blatt Seite 2", "Aufgabe A „1789 oder heute?“ angekreuzt, Aufgabe B „Historische Bedeutung“ (Vorher, Seit 1789, Heute), Rückmeldung nach Stufen bei Abgabe"],
  ["Heft", "Merksatz 1 (abgeschrieben, rot umrahmt)"],
].map(([a, b]) => row([tc(a, 2300, { bold: true, fill: LIGHT }), tc(b, 7906)], 340))));

d.push(h3("Nach der Stunde festhalten (Reflexionslog)"));
d.push(...bullets(["Passten Ziel, Inhalt und Methode zusammen? Wie lief die Zeit? Wie viel echte Lernzeit hatte die Klasse?", "Haben die Schritte „Lesen, Begriff heute, Vorher“ gesessen? Wo hakte es (Prüffrage, Artikel 3, selbstständige Phase)?", "Wie lief der Aha-Moment „1789 oder heute?“ (Raten, Reaktion)? Wie gut war die Sicherung (Merksatz, Aufgabe B, mündliche Antworten)?", "Wie viele Blätter wurden abgegeben, auf welcher Stufe lagen sie?", "Hospitationsrückmeldung zur Aktivierung und zum Kompetenzerwerb: Was nimmst du mit?", "Wie lief Stunde 3 (Rückmeldung nachtragen)?"]));
d.push(h3("Zum Prüfen"));
d.push(...bullets([
  "Wortlaut der Artikel: gekürzte und vereinfachte Fassung (Entscheidung der Lehrkraft); Nummerierung wie im Original (Artikel 6, in einem anderen Buch steht 5). Mit der Kollegin abgleichen. Neu: Artikel 17 (Eigentum) und der Begriff Eigentumsrecht; der Wortspeicher der Kollegin enthält ihn nicht (Test vergleichbar halten).",
  "Grundgesetz: Wortlaut der Artikel 3 Abs. 1, 5 Abs. 1, 14 Abs. 1 und 20 Abs. 2 ist leicht gekürzt, bitte mit dem Original abgleichen.",
  "4. August: Aufzählung der abgeschafften Vorrechte und die Ablösepflicht für Bauern gegen das Lehrbuch lesen. Dass die Aufstände der Bauern ein wichtiger Grund für den Verzicht waren, ist die gängige Darstellung.",
  "Vereinfachung: Auf Folie 7 und 15 stehen Stichworte. Zum Beispiel wurden manche Abgaben sofort aufgehoben, andere mussten abgelöst werden.",
  "Die Karikatur (Klerus, Adel und Dritter Stand, zwei Bilder an der Wand) ist selbst gezeichnet, keine historische Quelle. Die Aussage der Figuren („Wir kamen euch entgegen.“) ist bewusst überspitzt, eine erfundene Selbstdarstellung. Historisch war der Verzicht nicht freiwillig.",
  "Vormerken für Stunde 5 (Verfassung 1791): Der König weigerte sich zunächst, die Erklärung zu unterschreiben, und tat es erst nach dem Marsch der Frauen nach Versailles im Oktober 1789 (als Einstieg möglich).",
  "Die Notenfrage (Punkte, Schlüssel) ist offen: Das Blatt liefert Felder und Stufen. Das können wir später ergänzen.",
]));

(async () => {
  const doc = new Document({ styles: { default: { document: { run: { font: "Calibri", size: 22 } } } }, sections: [{ properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 1000, bottom: 850, left: 850, right: 850 } } }, headers: { default: header }, footers: { default: footer }, children: d }] });
  fs.writeFileSync("../Tafelskript_Stunde4_Menschenrechte_8D.docx", await Packer.toBuffer(doc));
  console.log("ok");
})();
