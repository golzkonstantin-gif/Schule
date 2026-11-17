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
const header = new Header({ children: [new Paragraph({ tabStops: [{ type: "right", position: W }], children: [t("Geschichte · Klasse 8D · Widerstand", { size: 17, color: MUTED }), t("\tfür die Lehrkraft", { size: 17, color: MUTED })] })] });
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
d.push(kicker("Tafelskript für die Lehrkraft · Einzelstunde (45 Min.) · 24 Schülerinnen und Schüler (12 Paare) · eine Lehrkraft · Mittwoch 18.11.2026"), h1("Welche Wege des Widerstands hat der Dritte Stand?"));
d.push(p(t("Ziel: Die Stunde knüpft an die Abstimmung an: Der Dritte Stand (Bauern und Bürger) kann nach Ständen nichts durch Wahlen ändern. Die Stunde beginnt mit einer selbst gezeichneten Karikatur (Klerus und Adel sprechen, der Dritte Stand hört zu: „Jetzt kann er sich nur noch fügen.“). In einem Partnerpuzzle (zwei Texte, fünf Kriterien, 10 Punkte) vergleichen die Paare zwei Wege des Widerstands: den gewaltlosen Ballhausschwur und den gewaltsamen Sturm auf die Bastille. Der Zeitstrahl zeigt die Reihenfolge (Aha-Moment), eine Bilanz trennt „wirksam“ von „gerechtfertigt“, und die Klasse beurteilt mit einem Maßstab. Am Ende antwortet sie in der Rolle des Dritten Stands (Bauern und Bürger) mündlich auf die Aussage von Klerus und Adel.", { size: 21, color: MUTED }), { after: 100 }));

d.push(h3("Kopf der Planung"));
d.push(table([2600, 7606], [
  ["Ziel (Ich kann …)", "… den Ballhausschwur und den Sturm auf die Bastille mit fünf Kriterien vergleichen. … an der Reihenfolge der Ereignisse erklären, wie die Wege zusammenwirkten.* … beurteilen, welcher Weg berechtigt war, und das mit einem Maßstab begründen."],
  ["Prüfaufgabe: Woran erkenne ich es?", "Matrix auf dem Blatt (10 Felder, 10 Punkte, freiwillige Abgabe, Rückmeldung und Bewertung nach Raster), Satz im Heft zur Frage „Wer hat die Gewalt angefangen?“, Urteil im Heft mit Maßstab (Sachurteil mit Beleg, dann Werturteil), mündliche Antwort als Dritter Stand (Ich-Form)."],
  ["Mindestziel", "Ziel 1 und 3. Ziel 2 hat einen Stern. „Auswerten“ bleibt immer."],
  ["Beitrag zur Reihe", "Stunde 3 von 11 (Sequenz, Fassung vom 17.11.). Sie zeigt, wie der Dritte Stand mit gewaltlosem und gewaltsamem Widerstand auf die verlorene Abstimmung reagiert, und bereitet Stunde 4 vor (Ende der Vorrechte, Menschenrechte; Artikel 2 nennt den Widerstand gegen Unterdrückung)."],
  ["Vorgehen", "Rückblick direkt (Partnercheck), Partnerpuzzle indirekt (Selbsterarbeitung, dann gegenseitiges Erklären), Aha, Bilanz und Urteil geführt (Beschreiben statt Gutheißen). Direkt oder indirekt wird je Phase entschieden."],
].map(([a, b]) => row([tc(a, 2600, { fill: LIGHT, bold: true }), tc(b, 7606)], 340))));

d.push(h3("Was aus den bisherigen Stunden übernommen wird"));
d.push(...bullets([
  "Stunde 2 (11.11., 8D): Die Abstimmung hat gezeigt, dass der Dritte Stand (Bauern und Bürger) nach Ständen nichts ändern kann; 1789 gab es keine Abstimmung nach Köpfen. Die Karikatur mit Klerus und Adel kehrt als Einstieg und als Abschluss wieder (bewährte Figuren).",
  "Rückmeldung 8C (06.11.): höchstens drei Methoden. Hier: Partnerpuzzle, Zeitstrahl-Auflösung mit Bilanz (Folie), Urteil. Die Karikatur ist nur Impuls.",
  "Entscheidung der Lehrkraft (17.11.): Der Dritte Stand kann nichts durch Wahlen ändern, die Stunde fragt nach den Wegen des Widerstands. Partnerpuzzle mit Ballhausschwur und Bastille in 2er-Gruppen, zwei Texte, fünf Kriterien (Vorgeschichte, Verlauf, Mittel, Opfer und Folgen, Wirkung), 10 Punkte zur Benotung, freiwillige Abgabe. Aha-Moment: Reihenfolge der Ereignisse am Zeitstrahl. Haltung: Gewalt ist wirksam, aber nicht automatisch gerechtfertigt (Bilanz, Urteil mit Maßstab, Ausblick).",
  "Die Menschenrechte (Kollegin hat sie behandelt) folgen in Stunde 4 als Begründung für Widerstand.",
]));

d.push(h3("Zeitplan (45 Minuten)"));
d.push(table([800, 2300, 4300, 2806], [
  row([th("Min.", 800), th("AVIVA", 2300), th("Phase", 4300), th("Arbeitsform", 2806)], 320),
  ...[["4", "Ankommen", "0  Karikatur, Vermutung, Stundenfrage, Ziele", "still, allein"], ["2", "Vorwissen", "1  Rückblick: Zeitstrahl und Abstimmung", "Partnercheck, Flüsterstufe"], ["10", "Informieren", "2  Partnerpuzzle: eigenen Text lesen, Tabelle für den eigenen Text ausfüllen", "allein, leise"], ["24", "Verarbeiten", "3  Austausch (8), Kontrolle (3), Zeitstrahl-Aha (4), Bilanz (1), Urteil (8)", "zu zweit, allein"], ["5", "Auswerten", "4  Antwort als Dritter Stand, Ziele, Abgabe, Ausblick", "mündlich, still"]].map((r) => row(r.map((v, i) => tc(v, [800, 2300, 4300, 2806][i], { bold: i < 3 })), 320)),
]));
d.push(note("ZEIT", "Die Stunde ist voll. Sollbruchstellen: (1) Austausch nach 6 Minuten kürzen, wenn ein Paar nicht wechselt. (2) Zeitstrahl-Aha: Schritt 1 und 2 zusammen zeigen (2 Minuten statt 4). (3) Urteil: nur ein Satz „Ich halte … für …, weil …“ statt drei Zeilen. Ziel 2 (*) fällt zuerst weg. Phase 4 wird nicht gestrichen."));

d.push(h3("Lautstärke und Ansagen (wörtlich)"));
d.push(table([4700, 5506], [
  ["„Stifte liegen. Ihr hört nur zu.“", "Sie erklären, niemand schreibt."],
  ["„Jetzt arbeitet ihr allein. Die Aufgabe steht rechts.“", "Einzelarbeit, leise."],
  ["„Flüsterstufe. Nur dein Partner hört dich.“", "Rückblick und Austausch zu zweit."],
  ["„Vergleicht und verbessert mit Grün.“", "Kontrolle im eigenen Blatt."],
  ["„Es ist zu laut. Ab jetzt arbeitet ihr allein und schreibt.“", "Wenn es laut wird, keine Diskussion."],
].map(([a, b]) => row([tc(a, 4700, { fill: LIGHT, bold: true }), tc(b, 5506)], 340))));

d.push(h3("Material und Vorbereitung (Details: Materialliste)"));
d.push(...bullets([
  "Partnerpuzzle_Texte_A_und_B_8D.pdf: eine Seite, oben Text A (Ballhausschwur), unten Text B (Bastille), 13-mal drucken und in der Mitte schneiden (13 A + 13 B, halbe A4-Seite, 12 Punkt), Text mit fünf Absätzen und Wörtern, keine Satznummern. Partnerpuzzle_Tabelle_8D.pdf: eine ganze A4-Seite pro Person (26 Blatt: 24 plus 2 Reserve), nicht geteilt, mit großen Feldern zum Eintragen, fünf Kriterien und Leitfragen, Name, „freiwillig abgeben“, Punkte. Auf den Tisch legen: links Text A, rechts Text B, dazu je eine Tabelle, umgedreht.",
  "Loesung_Partnerpuzzle_8D.pdf (für Sie): erwartete Einträge je Feld, Punktevergabe, Hinweise zur Haltung. Notenschlüssel für 10 Punkte nach Ihrer Festlegung (Raster Folie 7: 0 bis 4, 5 bis 8, 9 bis 10 Punkte).",
  "Folien 1 bis 20 (Schrift mindestens 27 Punkt; Folie 1 und 18 mit der selbst gezeichneten Karikatur (Klerus, Adel, Dritter Stand)).",
  "Seitentafel: Stundenfrage, Ziele, Ablauf (siehe Phase 0). Mitte der Tafel leer lassen für die Reihenfolge der Ereignisse (optional).",
  "Kreide, Grün zum Verbessern, Uhr oder Timer.",
]));

// Phase 0
d.push(phase("0", "Ankommen: Karikatur und Vermutung", 4, "AVIVA: Ankommen"));
d.push(warum("Die verlorene Abstimmung wird zur Frage („Jetzt kann er sich nur noch fügen.“). Die Vermutung im Heft wird am Ende der Stunde mit dem Ergebnis verglichen (Vorher-Nachher, ohne Fachbegriff). Alle haben beim Ankommen etwas zu tun (direkt, von der Lehrkraft geführt).", "„Klerus und Adel sagen: Er kann sich nur noch fügen. Ob das stimmt, prüfen wir jetzt.“"));
d.push(sp());
d.push(step("kommt an, schreibt leise", {
  say: "Stifte liegen. Ihr hört nur zu. Schaut auf die Folie und lest, was der Adlige sagt. Schreibt ins Heft: Was kann der Dritte Stand noch tun? Ich vermute …",
  do: [
    "Folie 1 steht beim Betreten da: selbst gezeichnete Karikatur (keine historische Quelle) mit drei Figuren: Klerus, Adel und der Dritte Stand (Bauern und Bürger), der zuhört. Der Adlige sagt: „Der Dritte Stand hat abgestimmt und verloren. Jetzt kann er sich nur noch fügen.“ Auftrag: Heft raus, Datum und Überschrift, „Was kann er noch tun? Ich vermute …“ (etwa 2 Minuten, keine Besprechung). Sie lesen beim Durchgehen bei 3 bis 4 Personen mit (nicht kommentieren).",
    "Folie 2 „Das ist mir wichtig“ (20 Sekunden), Folie 3 (Stundenfrage und Begriff „Widerstand“, 45 Sekunden: „Man wehrt sich gegen die Herrschenden, obwohl man sich fügen soll.“), Folie 4 (Lernziele, 30 Sekunden), Folie 5 (Ablauf, 30 Sekunden). Hinweis zum Stern: „Was wir eventuell nicht schaffen, hat einen Stern.“",
  ],
  extra: [boardBox("Seitentafel · Stundenfrage, Ziel und Ablauf", [
    bl("Wie kann sich der Dritte Stand wehren, wenn Wählen nichts bringt?", { after: 40 }),
    bl("Ich kann den Ballhausschwur und den Sturm auf die Bastille mit fünf Kriterien vergleichen.", { after: 20 }),
    bl("Ich kann an der Reihenfolge erklären, wie die Wege zusammenwirkten.*", { after: 20 }),
    bl("Ich kann beurteilen, welcher Weg berechtigt war (mit Maßstab).", { after: 60 }),
    bl("A Karikatur · V Rückblick · I Text lesen · V Puzzle, Urteil · A Antwort", { after: 0 }),
  ]), gapP()],
}));

// Phase 1
d.push(phase("1", "Rückblick: Zeitstrahl und Abstimmung", 2, "AVIVA: Vorwissen aktivieren"));
d.push(warum("Der Stoff der Stunde 2 wird kurz gesichert, bevor Neues kommt (allein, Partnercheck, kurze Abfrage). Der Zeitstrahl endet bei „Soldaten um Paris“ und führt in den Text B hinein.", "„Wählen bringt nichts. Welche Wege bleiben? Das lest ihr jetzt.“"));
d.push(sp());
d.push(step("hört zu, spricht zu zweit", {
  say: "Stifte liegen. Ihr hört nur zu. … Flüsterstufe. Nur dein Partner hört dich. Erklärt eurem Partner in einem Satz: Warum verlor der Dritte Stand die Abstimmung?",
  do: [
    "Folie 6 (Rückblick, 2 Minuten): Zeitstrahl mit vier Einträgen (5. Mai, 17. Juni, 20. Juni, Anfang Juli) und das Ergebnis der Abstimmung in einem Satz: Nach Ständen verlor der Dritte Stand (Bauern und Bürger) immer, obwohl er fast alle Menschen vertritt; 1789 gab es keine Abstimmung nach Köpfen. Partnercheck 1 Minute, danach gezielte Abfrage bei 1 bis 2 Personen (kurz, kein Gespräch).",
  ],
  sol: "Erwartung: Jeder Stand hat nur eine Stimme, Klerus und Adel haben zusammen zwei von drei Stimmen, obwohl sie nur etwa 3 % der Menschen sind.",
}));

// Phase 2
d.push(phase("2", "Partnerpuzzle: Einzelarbeit", 10, "AVIVA: Informieren"));
d.push(warum("Jede Person erarbeitet allein den eigenen, ausführlichen Text (indirekt: Selbsterarbeitung, Unterstreichen, Tabelle; keine Satznummern, keine Hilfen). So muss später jede Person für ihren Text sprechen und niemand kann sich verstecken. Kein Vorlesen.", "„Zwei Wege, zwei Texte. Jeder von euch hat nur einen. Ihr braucht euren Partner, um beide zu kennen.“"));
d.push(sp());
d.push(step("liest allein", {
  say: "Stifte liegen. Ihr hört nur zu. Vor euch liegen ein Text und eine Tabelle. Wer links sitzt, hat Text A, wer rechts sitzt, Text B. Jetzt arbeitet ihr allein. Die Aufgabe steht rechts.",
  do: [
    "Folie 7 (So wird dein Blatt bewertet, 1 Minute): Raster vorstellen, nicht diskutieren (0 bis 4 Punkte, 5 bis 8 Punkte, 9 bis 10 Punkte; je Feld ein Punkt; Abgabe freiwillig, keine Selbsteinschätzung). Folie 8 (Die fünf Kriterien) zeigen. Folie 9 (Ü1, 10 Minuten): eigenen Text (fünf Absätze) genau lesen, unterstreichen, eigene Spalte der Tabelle ausfüllen (Stichworte). Keine Hilfen: Die Schüler suchen selbst im Text.",
    "Warnung vor dem Fehler: Der Ballhausschwur ist Widerstand, obwohl es keine Gewalt gibt: Die Abgeordneten fügen sich dem Verbot des Königs nicht (Ungehorsam). Und „Mittel“ heißt nicht „Waffe“, sondern: Wird Gewalt gebraucht, ja oder nein, und woran erkennt man das? Beim Ballhausschwur steht „sie benutzen keine Waffen, sondern Worte, ihre Unterschrift und ihren Eid“, bei der Bastille „es wird geschossen“.",
    "Heft-Stichprobe beim Durchgehen (nicht kommentieren, unbenotet). Wer früh fertig ist, markiert im Text, woran man erkennt, ob Gewalt im Spiel war.",
  ],
}));

// Phase 3
d.push(phase("3", "Partnerpuzzle: Austausch, Aha, Bilanz, Urteil", 24, "AVIVA: Verarbeiten"));
d.push(warum("Beim gegenseitigen Erklären wird der Gegensatz gewaltlos gegen gewaltsam sichtbar (angeleitet). Der Zeitstrahl zeigt die Reihenfolge (Aha), die Bilanz trennt „wirksam“ von „gerechtfertigt“, das Urteil verlangt einen offengelegten Maßstab (geführt, dann selbstständig).", "„Beide wollten, dass der König nicht mehr allein entscheidet. Aber was hat bei der Reihenfolge gewirkt?“"));
d.push(sp());
d.push(step("3a  erklärt zu zweit (8 Min.)", {
  say: "Stifte liegen. Flüsterstufe. Nur dein Partner hört dich. Erst erklärt die Person links ihren Text, dann die Person rechts. Wer zuhört, schreibt mit.",
  do: [
    "Folie 10 (Ü2): Nach 4 Minuten „Wechselt.“ Jede Person erklärt in eigenen Worten, niemand liest vor. Sie gehen durch die Reihen, hören bei 4 bis 6 Paaren kurz zu, ohne zu helfen. Ziel: Beide Blätter sind vollständig (10 Felder).",
    "Sollbruchstelle: Ist ein Paar nach 6 Minuten noch nicht beim Wechsel, kürzen Sie die zweite Runde.",
  ],
}));
d.push(sp());
d.push(step("3b  kontrolliert (3 Min.)", {
  say: "Stifte liegen. Vergleicht und verbessert mit Grün.",
  do: [
    "Folie 11 (So sieht es aus): Stichworte zu jedem Kriterium für beide Texte. Mittel: gewaltlos gegen gewaltsam (grün und rot). Die Punkte vergeben Sie nach dem Lösungsblatt (nur bei abgegebenen Blättern).",
  ],
  sol: "Vorgeschichte: Saal gesperrt | Hunger, Soldaten. Verlauf: Eid im Ballhaus | Sturm auf die Festung. Mittel: gewaltlos | gewaltsam. Opfer, Folgen: keine Toten (Risiko: Auflösung) | etwa 100 Tote, Gouverneur getötet. Wirkung: Versammlung bleibt | Soldaten abgezogen, Symbol.",
}));
d.push(sp());
d.push(step("3c  Zeitstrahl-Aha (4 Min.)", {
  say: "Stifte liegen. Ihr hört nur zu. Wir bringen die Ereignisse in die Reihenfolge.",
  do: [
    "Folien 12 bis 15 (je 1 Minute, schrittweise): 20. Juni Ballhausschwur (gewaltlos); Anfang Juli der König zieht Soldaten zusammen (Drohung); 14. Juli Bastille (gewaltsam); Mitte Juli der König zieht die Soldaten ab (Wirkung). Fragen (jeweils still denken): „Gegen wen richtet sich der Eid?“ (gegen den König und sein Verbot) „Wie reagiert der König?“ „Was passiert in Paris?“ „Wer hat die Gewalt angefangen?“ (ein Satz ins Heft) „Was hat gewirkt?“",
    "Aha: Der gewaltlose Schritt schuf die Grundlage, der Druck in Paris zwang den König zum Rückzug. Der Anlass der Gewalt war die Drohung (Soldaten), aber Anlass ist keine Rechtfertigung.",
  ],
}));
d.push(sp());
d.push(step("3d  Bilanz (1 Min.)", {
  say: "Stifte liegen. Ihr hört nur zu. Wirksam heißt nicht gerechtfertigt.",
  do: [
    "Folie 16 (Was bewirkte es? Was kostete es?): Ballhausschwur: Versammlung bleibt, König gibt nach | keine Toten, Risiko Auflösung. Bastille: Truppen abgezogen, Symbol | etwa 100 Tote, Gouverneur getötet, nur sieben Gefangene. Ihr Satz: „Wir erklären, warum es dazu kam. Ob es richtig war, beurteilt ihr gleich mit einem Maßstab.“ Zahlen zur Bastille gegen das Lehrbuch prüfen.",
  ],
}));
d.push(sp());
d.push(step("3e  urteilt allein (8 Min.)", {
  say: "Stifte liegen. Jetzt arbeitet ihr allein. Die Aufgabe steht rechts.",
  do: [
    "Folie 17 (Ü3): Maßstab wählen (wirksam, gerecht oder verantwortbar), ins Heft: „Ich halte … für …, weil …“, mit Beleg und Opfern, dann Vergleich mit der Vermutung vom Anfang („Ich hatte vermutet …, jetzt sehe ich …“). Hilfen als Satzanfänge: „Wirksam, weil …“ „Gerecht, weil …“ „Verantwortbar, weil … (die Opfer: …)“. Sachurteil mit Beleg, dann Werturteil mit offengelegtem Maßstab. Heft-Stichprobe bei 4 bis 6 Personen unterschiedlichen Niveaus, nicht kommentieren (unbenotet; das Urteil zählt nicht zu den 10 Punkten).",
  ],
  sol: "Mögliche Urteile (Auswahl, keine Musterlösung). Ballhausschwur: wirksam (Versammlung blieb), gerecht (keine Gewalt, Ziel Verfassung), verantwortbar (keine Toten). Bastille: wirksam (Truppen abgezogen), aber nicht verantwortbar (etwa 100 Tote, getöteter Gouverneur); als Reaktion auf die Soldaten verständlich, aber das macht sie nicht gerecht. Beide Wege zusammen wirkten stärker als jeder allein.",
}));

// Phase 4
d.push(phase("4", "Auswerten: Antwort als Dritter Stand", 5, "AVIVA: Auswerten"));
d.push(warum("Die Klasse nimmt Bezug auf die Aussage von Klerus und Adel vom Anfang und antwortet als Dritter Stand: Das prüft offen, was sie gelernt hat (mehrere gute Antworten, Vorstufe zu den Urteilsaufgaben der Reihe). Der Vergleich mit der Vermutung zeigt den Lernzuwachs. Diese Phase wird nie gestrichen.", "„Klerus und Adel sagten: Jetzt bleibt ihm nur noch, sich zu fügen. Ihr seid der Dritte Stand. Was antwortet ihr?“"));
d.push(sp());
d.push(step("schaut auf die Karikatur, denkt, antwortet als Dritter Stand", {
  say: "Stifte liegen. Ihr hört nur zu. Das ist die letzte Aufgabe. Schaut noch einmal auf die Karikatur vom Anfang. Jetzt seid ihr der Dritte Stand. Denkt zehn Sekunden nach. Ich rufe euch auf.",
  do: [
    "Folie 18 (Abschlussfrage, 3 Minuten): Die Karikatur erscheint wieder mit der Aussage von Klerus und Adel. Auftrag: „Du bist der Dritte Stand. Was antwortest du Klerus und Adel?“ Die Antwort kommt mündlich in der Ich-Form, ohne Schnipsel: „Ich antworte euch: …, weil …“. Mehrere gute Antworten, geprüft werden Position, ein Sachargument aus der Stunde (Eid, Bastille, Reihenfolge, Opfer) und ein „weil“. Form: (a) Blitzlicht (empfohlen): 4 bis 6 Personen unterschiedlichen Niveaus gezielt aufrufen, kein Kommentar. (b) Duell zu zweit, leise: eine Person ist Klerus und Adel, die andere antwortet als Dritter Stand. Unbenotet. Zum Schluss ein Blick ins Heft: Wer hat seine Vermutung geändert? (keine Abfrage).",
    "Folie 19 (Ziele prüfen, 30 Sekunden, still): Haben wir die Ziele erreicht? Wer das Blatt abgeben möchte, legt es jetzt vorn ab. Folie 20 (Ausblick, 30 Sekunden, nur lesen): „Worauf berufen sich die Menschen, wenn sie sich wehren? Das sehen wir in der nächsten Stunde.“ Ihr Satz am Ende: Was passiert, wenn Gewalt zum normalen Mittel wird, sehen wir bei den Jakobinern.",
  ],
  sol: "Mögliche Antworten als Dritter Stand (Auswahl, keine Musterlösung). „Wir fügen uns nicht, weil wir fast alle Menschen vertreten und uns zur Nationalversammlung erklärt haben.“ „Wir wehren uns, aber ohne Gewalt, weil die Bastille viele Tote gekostet hat.“ „Erst der Druck von beiden Seiten hat den König zum Rückzug gebracht, also fügen wir uns nicht.“ Kriterien: Position, ein Sachargument aus der Stunde, ein „weil“.",
}));

d.push(h3("So sieht das Heft und das Blatt am Ende aus"));
d.push(table([2300, 7906], [
  ["Überschrift", "Datum · Überschrift zur Stundenfrage (z. B. „Wege des Widerstands“)"],
  ["Vermutung", "„Was kann der Dritte Stand noch tun? Ich vermute …“ (Beginn der Stunde)"],
  ["Blatt (Matrix)", "zehn Felder gefüllt, Verbesserungen grün; bei Abgabe Punkte und Rückmeldung"],
  ["Satz zur Gewalt", "„Wer hat die Gewalt angefangen?“ (ein Satz)"],
  ["Urteil", "Ich halte … für …, weil … (mit Maßstab, Beleg und Opfern), Vergleich mit der Vermutung"],
].map(([a, b]) => row([tc(a, 2300, { bold: true, fill: LIGHT }), tc(b, 7906)], 340))));

d.push(h3("Nach der Stunde festhalten (Reflexionslog)"));
d.push(...bullets(["Passten Ziel, Inhalt und Methode zusammen? Wie lief die Zeit? Wie viel echte Lernzeit hatte die Klasse?", "Hat das Partnerpuzzle funktioniert (Austausch, Lautstärke)? Wie viele Blätter wurden abgegeben, wie viele Punkte?", "Kam der Aha-Moment am Zeitstrahl an? Was schrieben die Schüler zur Frage „Wer hat die Gewalt angefangen?“", "Wie urteilten die Schüler: Haben sie Gewalt als Lösung gesehen oder differenziert (wirksam, aber nicht verantwortbar)?", "Was muss in Stunde 4 wiederholt werden (Ziel 2)?"]));
d.push(h3("Zum Prüfen"));
d.push(...bullets([
  "Zahlen zur Bastille (etwa 100 Tote unter den Angreifern, sieben Gefangene, Gouverneur nach der Übergabe getötet) stammen aus allgemeinem Wissen und sind gegen das Lehrbuch zu lesen. Wirkung des Ballhausschwurs: Am 23. Juni versuchte der König noch, die Beschlüsse aufzuheben, am 27. Juni forderte er Klerus und Adel auf, sich anzuschließen.",
  "Die Aussage „Der König zieht die Soldaten ab“ (Mitte Juli) steht im Text B und auf Folie 15. Das genaue Datum im Lehrbuch gegenprüfen.",
  "Ob die Gewalt in Paris den König allein zum Nachgeben brachte, ist offen. Im Tafelskript steht die Deutung „der Druck wirkte erst durch beides“, die Klasse belegt sie mit der Reihenfolge.",
  "Die Karikatur (Klerus, Adel und Dritter Stand) ist selbst gezeichnet, keine historische Quelle.",
]));

(async () => {
  const doc = new Document({ styles: { default: { document: { run: { font: "Calibri", size: 22 } } } }, sections: [{ properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 1000, bottom: 850, left: 850, right: 850 } } }, headers: { default: header }, footers: { default: footer }, children: d }] });
  fs.writeFileSync("../Tafelskript_Stunde3_Widerstand_8D.docx", await Packer.toBuffer(doc));
  console.log("ok");
})();
