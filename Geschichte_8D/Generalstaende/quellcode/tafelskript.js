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
const header = new Header({ children: [new Paragraph({ tabStops: [{ type: "right", position: W }], children: [t("Geschichte · Klasse 8D · Generalstände", { size: 17, color: MUTED }), t("\tfür die Lehrkraft", { size: 17, color: MUTED })] })] });
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
d.push(kicker("Tafelskript für die Lehrkraft · Einzelstunde (45 Min.) · 25 Schülerinnen und Schüler · eine Lehrkraft"), h1("Generalstände: Wie können die Stände den Dritten Stand überstimmen?"));
d.push(p(t("Ziel: Die Klasse erarbeitet den Weg zu den Generalständen am Zeitstrahl (Text aus der 8C, hat gut funktioniert). Danach erlebt sie in einer Ständeversammlung selbst, wie Klerus und Adel sich hinter dem Rücken des Dritten Standes absprechen und ihn trotz Mehrheit überstimmen. Das Spiel ersetzt das Schulbeispiel aus der 8C (zu abstrakt) und den Gesprächsanteil (zu lastig). Aufbau nach AVIVA.", { size: 21 }), { after: 80 }));

d.push(h3("Kopf der Planung"));
d.push(table([2600, 7606], [
  ["Ziel (Ich kann …)", "… den Weg bis zu den Generalständen am Zeitstrahl erklären. … erklären, wie die Stände den Dritten Stand überstimmen konnten. … begründen, warum die Abstimmung nach Ständen unfair war."],
  ["Prüfaufgabe: Woran erkenne ich es?", "Zeitstrahl ausgefüllt (Ziel 1). Ü3 Satz 1 und 2 (Ziel 2). Ü3 Satz 3 mit einer Zahl (Ziel 3). offene Abschlussfrage (Rat an den König mit Position, Sachargument und „weil“); Zeitstrahl (Ziel 1) und Ü3 Satz 1 und 2 (Ziel 2) sichern die Fakten."],
  ["Mindestziel", "Ziel 1 und 2. Alles, was mit * steht, ist Zusatz (Ü3 Satz 3, Sportverein-Beispiel). „Auswerten“ bleibt immer."],
  ["Beitrag zur Reihe", "Stunde 1 von 11 (Sequenz Französische Revolution). Sie legt die Ursachen und den Streit um die Abstimmungsregel. Daran knüpft Stunde 2 an: Der Dritte Stand fordert die Abstimmung nach Köpfen und gründet die Nationalversammlung."],
  ["Vorgehen", "Direktes Vorgehen in Phase 1 bis 2 (Vorwissen gering, Zahlen und Regel werden vorgegeben), Öffnung in Phase 3 (Spiel) erst nach dem Modell. Direkt oder indirekt wird je Phase entschieden."],
].map(([a, b]) => row([tc(a, 2600, { fill: LIGHT, bold: true }), tc(b, 7606)], 340))));
d.push(h3("Was aus der 8C übernommen und geändert wurde"));
d.push(...bullets([
  "Übernommen: Text „Wie kam es zu den Generalständen?“ mit Zeitstrahl auf demselben Blatt (hat gut funktioniert), Daten vorgegeben, Ereignisse eintragen.",
  "Geändert, weil die 8C unterrichtsgesprächslastig war: Die Stunde läuft fast ganz über Aufträge (still, Heft, Partner, Kreide-Kette). Gespräch im Plenum gibt es nur in zwei kurzen Momenten (je unter einer Minute, siehe Phase 3d).",
  "Geändert, weil das Schulbeispiel zu abstrakt war: Die Klasse wird selbst in Stände geteilt (3 Klerus, 4 Adel, 18 Dritter Stand). Klerus und Adel sprechen sich schriftlich ab, der Dritte Stand beobachtet, dann wird gezählt: nach Ständen 2 : 1 abgelehnt, nach Köpfen angenommen.",
  "Neu wegen der Klassengröße und Lautstärke: Zwischen den Phasen gibt es wörtliche Ansagen, feste Zeiten und eine einzige Bewegungsphase.",
]));

d.push(table([800, 2300, 4300, 2806], [
  row([th("Min.", 800), th("AVIVA", 2300), th("Phase", 4300), th("Arbeitsform", 2806)], 320),
  ...[["3", "Ankommen", "0  Ankommen, Lernziele, Test", "still, allein"], ["6", "Vorwissen", "1  Wiederholung: Ständegesellschaft", "allein, Kontrolle per Folie"], ["11", "Informieren", "2  Text, Zeitstrahl, Kreide-Kette", "allein, Tafel mit Mitschreiben"], ["18", "Verarbeiten", "3  Ständeversammlung: Absprache, Abstimmung, Auswertung", "Stände, allein, Partner"], ["6", "Auswerten", "4  Merksatz, Abschlussfrage (Berater), Ausblick", "allein, abschreiben"], ["1", "—", "Puffer", "—"]].map((r) => row(r.map((v, i) => tc(v, [800, 2300, 4300, 2806][i], { bold: i < 3 })), 320)),
]));

d.push(h3("Lautstärke und Gesprächsanteil: so läuft die Stunde"));
d.push(table([3300, 6906], [
  ["Ankommen", "Der Auftrag steht beim Betreten auf Folie 1. Niemand muss etwas sagen, jede und jeder hat sofort etwas zu tun."],
  ["Gespräch ersetzen", "Statt Fragen ins Plenum: aufschreiben (Heft), Folie zeigen, Kreide-Kette. Die Klasse antwortet in der Regel schriftlich."],
  ["Fest vorgegebene Zeit", "Jede Phase hat eine sichtbare Zeit auf der Folie. Ankündigung 1 Minute vor Ende: „Noch eine Minute.“"],
  ["Lautstärke wörtlich", "„Flüsterstufe. Nur dein Partner hört dich.“ bzw. „Leise. Jeder arbeitet allein.“ Immer in diesem Wortlaut, ohne Symbole."],
  ["Einzige Bewegung", "Erster und Zweiter Stand gehen einmal an den Tisch vorn (Ansage 3a). Der Dritte Stand (18 von 25) bleibt sitzen. Rückweg mit Ansage in Phase 3d."],
  ["Wenn es zu laut wird", "Phase wird sofort in Einzelarbeit fortgesetzt: „Es ist zu laut. Ab jetzt arbeitet ihr allein und schreibt.“ Keine Diskussion."],
  ["Melden", "„Wer etwas sagen will, schreibt es zuerst auf.“ Aufrufen nur, wer etwas aufgeschrieben hat."],
].map(([a, b]) => row([tc(a, 3300, { fill: LIGHT, bold: true }), tc(b, 6906)], 340))));

d.push(h3("Eindeutige Ansagen (wörtlich)"));
d.push(table([4700, 5506], [
  ["„Stifte liegen. Ihr hört nur zu.“", "Sie erklären, niemand schreibt."],
  ["„Jetzt arbeitet ihr allein. Die Aufgabe steht rechts.“", "Einzelarbeit, leise."],
  ["„Flüsterstufe. Nur dein Partner hört dich.“", "Partner- und Standesarbeit."],
  ["„Wer drankommt, schreibt an die Tafel. Alle prüfen mit.“", "Kreide-Kette beim Zeitstrahl."],
  ["„Jetzt schreibt ihr ab: …“", "Merksatz ins Heft."],
  ["„Vergleicht und verbessert mit Grün.“", "Kontrolle im eigenen Heft."],
].map(([a, b]) => row([tc(a, 4700, { fill: LIGHT, bold: true }), tc(b, 5506)], 340))));

d.push(h3("Material und Vorbereitung"));
d.push(...bullets([
  "Text „Wie kam es zu den Generalständen?“ (ein Blatt pro Person, wie 8C; Zeitstrahl unten, Daten vorgegeben).",
  "Druckvorlage „Standeskarten und Stimmzettel“: Seite 1: 3 Karten Klerus, 4 Karten Adel, 3 Karten Dritter Stand; Seite 2: 10 Karten Dritter Stand; Seite 3: 5 Karten Dritter Stand und 5 Reserve-Karten; Seite 4: zwei Absprache-Zettel (einen reicht). Ausschneiden. Jede Karte enthält unten den Stimmzettel, nichts muss abgerissen werden.",
  "Vor der Stunde: Karten verdeckt auf die Plätze legen. Klerus und Adel gezielt verteilen: eher ruhige, zuverlässige Schülerinnen und Schüler, damit die Absprache leise abläuft. Der Tisch vorn („Hinterzimmer“) steht für sieben Personen bereit (Stühle dazu).",
  "Wenn Schülerinnen oder Schüler fehlen: Klerus bleibt 3, Adel 4, die Reserve-Karten gleichen den Dritten Stand aus (die beiden privilegierten Stände dürfen zusammen nie mehr als ein Drittel der Klasse sein).",
  "Falls die Tests noch nicht zurückgegeben sind: verdeckt auf die Plätze legen, keine Besprechung, Berichtigung mit Grün bis zur nächsten Stunde. Falls schon zurückgegeben, entfällt der Punkt.",
  "Tafel (Seitentafel): Stundenziel „Ich kann …“ (3 Sätze) und Ablauf in Stichpunkten mit * für Phase 4, Satz 3 (siehe Phase 0). Mitte der Tafel leer lassen (Zeitstrahl, Ergebnis der Abstimmung).",
  "Folien 1 bis 18 (Schrift mindestens 24 Punkt), Kreide, Grün zum Verbessern, ein Stapel Heftpapier für die Absprache-Zettel falls nötig.",
]));

// Phase 0
d.push(phase("0", "Ankommen", 3, "AVIVA: Ankommen"));
d.push(warum("Die Klasse wird ruhig, kennt Ziel und Ablauf und weiß, was am Ende gekonnt werden soll (direkt, von der Lehrkraft geführt).", "„Zuerst schauen wir, was ihr aus dem Absolutismus noch wisst, denn das braucht ihr gleich.“"));
d.push(sp());
d.push(step("kommt an, arbeitet leise", {
  say: "Stifte liegen. Ihr hört nur zu. Auf dem Platz liegt eine Karte. Lasst sie verdeckt liegen. Schreibt Datum und Überschrift „Generalstände 1789“ ins Heft.",
  do: [
    "Folie 1 steht beim Betreten da (Auftrag: Heft raus, Datum und Überschrift, Karte nicht umdrehen). Folie 2 (Lernziele) und Folie 3 (Ablauf nach AVIVA) kurz vorlesen, 1 Minute, die Klasse hört nur zu. An der Seitentafel steht dasselbe in Stichpunkten.",
    "Hinweis zum Stern: „Was wir eventuell nicht schaffen, hat einen Stern. Am Ende prüfen wir, ob wir die Ziele erreicht haben.“",
  ],
  extra: [boardBox("Seitentafel · Ziel und Ablauf", [
    bl("Ich kann den Weg bis zu den Generalständen am Zeitstrahl erklären.", { after: 20 }),
    bl("Ich kann erklären, wie die Stände den Dritten Stand überstimmen konnten.", { after: 20 }),
    bl("Ich kann begründen, warum die Abstimmung nach Ständen unfair war.", { after: 60 }),
    bl("A Ankommen · V Wiederholung · I Text und Zeitstrahl · V Ständeversammlung · A Auswerten (Satz 3 *)", { after: 0 }),
  ]), gapP()],
}));

// Phase 1
d.push(phase("1", "Wiederholung: Ständegesellschaft", 6, "AVIVA: Vorwissen aktivieren"));
d.push(warum("Zahlen und Vorrechte der drei Stände aktiv abrufen, denn ohne sie ist die Abstimmung im Spiel nicht zu verstehen (direkt: erst allein, Kontrolle per Folie).", "„Klerus und Adel zahlen keine Steuern. Wie es dazu kam, dass der König ihre Zustimmung doch wieder brauchte, lest ihr jetzt.“"));
d.push(sp());
d.push(step("arbeitet allein, kontrolliert", {
  say: "Jetzt arbeitet ihr allein. Die Aufgabe steht rechts.",
  do: ["Folie 4 (Ü1). 4 Minuten schriftlich im Heft, danach Folie 5 zeigen (Lösung), 2 Minuten „Vergleicht und verbessert mit Grün.“ Kein Plenumsgespräch. Partnercheck nach Ü1 (1 Minute, flüstern) nur, wenn die Klasse ruhig ist; sonst entfällt er. Der Inhalt knüpft an die letzte Stunde an: Die Zahlen und Vorrechte brauchen die Schülerinnen und Schüler gleich im Spiel."],
  extra: [folieBox("4 · Ü1 Ständegesellschaft", [p(t("Ü1 Die drei Stände", { font: "Cambria", size: 22, bold: true, color: NAVY }), { after: 40 }), bl("1. Schreibe die drei Stände untereinander ins Heft."), bl("2. Schreibe dazu die Zahl der Menschen: 150 000 · 500 000 · 20 000 000."), bl("3. Schreibe hinter die beiden Stände mit Vorrechten: „zahlt keine Steuern“."), bl("4. Beende: „Die Steuern zahlt der …“", { after: 40 }), p([t("Zeit: ", { bold: true, size: 19 }), t("4 Minuten · allein · leise     ", { size: 19 }), t("Fertig? Stift hinlegen.", { size: 19, italics: true })], { after: 0 })]), gapP()],
  sol: "Klerus 150 000 · Adel 500 000 · Dritter Stand (Bauern und Bürger) 20 000 000. Klerus und Adel zahlen keine Steuern. Die Steuern zahlt der Dritte Stand (Bauern und Bürger).",
}));
d.push(sp());
d.push(note("SOLLBRUCHSTELLE", "Läuft die Klasse langsam oder laut an, wird Folie 5 sofort ohne Besprechung gezeigt (2 Minuten gewonnen)."));

// Phase 2
d.push(phase("2", "Text, Zeitstrahl, Kreide-Kette", 11, "AVIVA: Informieren"));
d.push(warum("Der Verlauf bis 1789 wird selbst aus dem Text erarbeitet, damit Ursachen und Reihenfolge klar sind (indirekt: Lesen allein, Kreide-Kette als Kontrolle). Die Kreide-Kette bringt Ruhe und lässt alle mitprüfen; ein Partnervergleich vorher wäre bei dieser Klasse zu laut.", "„Der König braucht Geld und ruft die Generalstände. Aber wie wird dort abgestimmt? Das erlebt ihr jetzt selbst.“"));
d.push(sp());
d.push(step("arbeitet allein", {
  say: "Ich teile den Text aus. Jetzt arbeitet ihr allein. Die Aufgabe steht rechts.",
  do: ["Folie 6 (Ü2). Blatt austeilen: oben der Text, unten der Zeitstrahl mit Daten, aber leeren Ereignisfeldern. 6 Minuten: lesen, Jahreszahlen unterstreichen, zu jedem Datum das Ereignis als Stichwort eintragen. Währenddessen leise durch die Reihen gehen, nichts erklären."],
  extra: [folieBox("6 · Ü2 Text und Zeitstrahl", [p(t("Ü2 Wie kam es zu den Generalständen?", { font: "Cambria", size: 22, bold: true, color: NAVY }), { after: 40 }), bl("1. Lies den Text und unterstreiche die Jahreszahlen."), bl("2. Schreibe zu jedem Datum das Ereignis in den Zeitstrahl."), bl("3. Prüfe: Passt jedes Ereignis zum Datum?", { after: 40 }), p([t("Zeit: ", { bold: true, size: 19 }), t("6 Minuten · allein · leise     ", { size: 19 }), t("Fertig? Stift hinlegen.", { size: 19, italics: true })], { after: 0 })]), gapP()],
}));
d.push(sp());
d.push(step("Kreide-Kette, alle prüfen mit", {
  say: "Wer drankommt, schreibt an die Tafel. Alle prüfen mit. Gesprochen wird nicht.",
  do: [
    "Die Datumsangaben stehen schon an der Tafel (links, untereinander). Sieben Schülerinnen und Schüler kommen nacheinander an die Tafel, schreiben das Ereignis zum nächsten Datum (Stichwort, nicht abschreiben) und geben die Kreide weiter. Die anderen vergleichen mit ihrem Blatt und unterstreichen Abweichungen. 4 Minuten, ohne Kommentar.",
    "Danach Folie 7 (Zeitstrahl als Lösung) zeigen. „Vergleicht und verbessert mit Grün.“ 1 Minute.",
  ],
  extra: [boardBox("Mitte · Zeitstrahl (Lösung)", [
    bl("1614    letzte Generalstände (danach 175 Jahre nicht mehr)"),
    bl("1778–1783    Krieg in Amerika, hohe Schulden"),
    bl("1787    Plan „Steuern auch für Adel und Klerus“ scheitert"),
    bl("1788    Missernte, Brotpreis steigt, Staatsbankrott droht"),
    bl("August 1788    König kündigt Generalstände an"),
    bl("Frühjahr 1789    Wahlen, Beschwerdehefte"),
    bl("5. Mai 1789    Eröffnung in Versailles", { after: 0 }),
  ]), gapP()],
  sol: "Die Frage „Warum ruft der König die Generalstände ein, obwohl er sie 175 Jahre nicht brauchte?“ nicht ins Plenum stellen, sie steht im Text als Aufgabe 3 und kann bei Zeit schriftlich im Heft beantwortet werden (*), nicht im Plenum. Antwort: Er braucht Geld, die Privilegierten blockieren Steuern, dazu kommen Hunger und drohender Bankrott.",
}));

// Phase 3
d.push(phase("3", "Ständeversammlung (Spiel)", 18, "AVIVA: Verarbeiten"));
d.push(warum("Die Regel wird erst vorgesprochen und modelliert (direkt), dann erlebt die Klasse sie selbst im Spiel (indirekt, handelnd) und wertet sie schriftlich aus. So entsteht das Sachurteil „unfair, weil …“ aus eigener Erfahrung.", "„Ihr habt erlebt, dass die Regel den Ausgang bestimmt. Jetzt halten wir fest, was das für Frankreich 1789 hieß.“"));
d.push(sp());
d.push(note("SPIELIDEE", "Die Klasse wird in Stände geteilt. Klerus und Adel sprechen sich ab, der Dritte Stand beobachtet. Beide Regeln werden anschließend mit denselben Stimmzetteln gezählt: nach Ständen (jeder Stand eine Stimme) und nach Köpfen (jede Person eine Stimme). Die Klasse sieht, dass dieselben Stimmen zu zwei Ergebnissen führen. Das Spiel gilt nur für die Stunde.", "EAF0FA", NAVY, "B9C6E8"));
d.push(sp());
d.push(step("dreht Karte um, Stände gehen an ihre Plätze", {
  say: "Stifte liegen. Ihr hört nur zu. Dreht jetzt eure Karte um. Lest sie leise. Zeigt sie niemandem.",
  do: [
    "3a  (3 Min.) Folie 8: Antrag „Auch Adel und Klerus zahlen künftig Steuern.“ und die Regel: Jeder Stand hat eine Stimme (so war es bei den Generalständen, siehe Text). Sprechen Sie die Regel einmal vor.",
    "Modell (30 Sekunden, laut denken): „Es gibt drei Stände, jeder hat eine Stimme. Zwei Stimmen reichen. Wenn Klerus und Adel beide Nein sagen, ist der Antrag durch, egal, was der Dritte Stand sagt.“",
    "Warnung vor dem Fehler: Viele denken, der Dritte Stand habe bei der Abstimmung nach Ständen mehr Stimmen, weil er mehr Menschen hat. Es ist eine Stimme für den ganzen Stand, nicht eine pro Person.",
    "Gezielte Abfrage (Stichprobe bei 3 Schülerinnen und Schülern unterschiedlichen Niveaus, unbenotet, getrennt von den mündlichen Noten, kein Gespräch): „Wie viele Stimmen hat der Dritte Stand?“ – „Wie viele braucht man zum Sieg?“ – „Wer hätte schon gewonnen?“ Neutral reagieren, kurz bestätigen oder korrigieren.",
    "Ansage: „Erster und Zweiter Stand, ihr geht jetzt leise an den Tisch vorn. Dritter Stand, ihr bleibt sitzen.“ Die sieben Personen setzen sich an den Tisch vorn. Absprache-Zettel dort ablegen.",
  ],
  sol: "Nur dieser eine Platzwechsel. Falls es unruhig wird: Karten zuerst lesen lassen, dann erst wechseln.",
}));
d.push(sp());
d.push(step("arbeitet in Ständen, Flüsterstufe", {
  say: "Flüsterstufe. Nur dein Partner hört dich. Ihr habt vier Minuten.",
  do: [
    "3b  (5 Min., Folie 9) Erster und Zweiter Stand füllen gemeinsam den Absprache-Zettel aus (drei Fragen: Wie stimmen wir ab? Was versprechen wir uns? Was sagen wir dem Dritten Stand?). Ein Blick auf den Tisch vorn gehört zum Spiel.",
    "Dritter Stand (18 Personen): Partnerarbeit mit dem Sitznachbarn. Auftrag im Heft: (1) Schreibe zwei Argumente auf, warum auch Adel und Klerus Steuern zahlen sollen. (2) Beobachte den Tisch vorn und schreibe eine Beobachtung auf: Was tun Klerus und Adel? Die Beobachtung ist der eigentliche Zielpunkt der Phase. Nicht darauf hinweisen, dass es eine Absprache ist.",
    "Eine Minute vor Ende: „Noch eine Minute.“ Zeit abbrechen, wenn alle sieben den Zettel fertig haben.",
  ],
  sol: "Beobachtungen z. B.: „Die flüstern“, „Die schreiben etwas auf“, „Die sitzen zusammen“, „Die sprechen sich ab“.",
}));
d.push(sp());
d.push(step("stimmt ab, arbeitet leise", {
  say: "Stifte liegen. Jetzt stimmt ihr ab. Kreuzt auf dem Stimmzettel unten auf eurer Karte Ja oder Nein an. Legt die Karte verdeckt auf den Tisch. Danach arbeitet ihr allein.",
  do: [
    "3c  (3 Min., Folie 10) Klerus und Adel stimmen nach Rolle. Der Dritte Stand stimmt, wie er es für richtig hält. Karten einsammeln, in drei Stapel sortieren (Klerus, Adel, Dritter Stand). Während Sie zählen (1,5 Minuten), arbeitet die Klasse allein: „Schreibe auf: Wer gewinnt nach Ständen? Wer gewinnt nach Köpfen? Begründe in einem Satz.“",
  ],
  extra: [boardBox("Links · Ergebnis (wird während des Zählens an die Tafel geschrieben)", [
    bl("nach Ständen:   Klerus ___  ·  Adel ___  ·  Dritter Stand ___   →   ___ : ___   →   ______"),
    bl("nach Köpfen:    Ja ___  ·  Nein ___   →   ______", { after: 0 }),
  ]), gapP()],
  sol: "Erwartet: Klerus 3 Nein, Adel 4 Nein, Dritter Stand Ja (18 Personen, evtl. einzelne Nein-Stimmen). Nach Ständen: Klerus Nein, Adel Nein, Dritter Stand Ja = 2 : 1 abgelehnt. Nach Köpfen: Ja ca. 18, Nein ca. 7 = angenommen. Tatsächliche Zahlen an die Tafel schreiben, nicht die erwarteten.",
}));
d.push(sp());
d.push(step("schaut zu, Stände gehen zurück", {
  say: "Stifte liegen. Ihr hört nur zu. Erster und Zweiter Stand, ihr geht leise an eure Plätze zurück.",
  do: [
    "3d  (3 Min., Folien 11 und 12) Ergebnis an der Tafel zeigen (nach Ständen, dann nach Köpfen), je ein Satz: „Nach Ständen: abgelehnt, 2 zu 1.“ „Nach Köpfen: angenommen.“",
    "Absprache-Zettel vorlesen (Plenum, unter einer Minute). Dabei Folie 11 zeigen: „Zwei Regeln. Dieselben Stimmen. Zwei Ergebnisse.“ Danach Folie 12 (Frankreich): Im Spiel sind die Verhältnisse gröber als in Frankreich; die echten Zahlen stehen auf Folie 12. Klerus und Adel sind etwa 3 %, aber sie haben 2 von 3 Stimmen.",
    "Gezielte Abfrage (2 bis 3 Personen, kurz): „Warum hat der Antrag nach Ständen verloren, obwohl 18 Personen dafür waren?“ – „Was ist anders nach Köpfen?“ Das ersetzt das Gespräch.",
    "Zweiter Gesprächsmoment, ebenfalls unter einer Minute: „Wer will, schreibt jetzt auf, was ihn am meisten überrascht hat.“ Keine Antworten ins Plenum.",
  ],
}));
d.push(sp());
d.push(step("arbeitet allein", {
  say: "Jetzt arbeitet ihr allein. Die Aufgabe steht rechts.",
  do: ["3e  (4 Min., Folie 13) Ü3 im Heft: drei Sätze. Satz 3 ist mit * markiert und darf entfallen, wenn die Zeit knapp wird. Zeit angeben, nach 3 Minuten „Noch eine Minute.“ Stichprobe im Heft: durch die Reihen gehen und bei 4 bis 6 Personen unterschiedlichen Niveaus mitlesen, nicht kommentieren. Im Wechsel mit der gezielten Abfrage einsetzbar."],
  extra: [folieBox("13 · Ü3 Was ist passiert?", [p(t("Ü3 Die Abstimmung", { font: "Cambria", size: 22, bold: true, color: NAVY }), { after: 40 }), bl("1. Der Dritte Stand wurde überstimmt, weil …"), bl("2. Klerus und Adel haben …"), bl("3.* Die Abstimmung nach Ständen ist unfair, weil … (nenne eine Zahl)", { after: 40 }), p([t("Zeit: ", { bold: true, size: 19 }), t("4 Minuten · allein · leise     ", { size: 19 }), t("Fertig? Stift hinlegen.", { size: 19, italics: true })], { after: 0 })]), gapP()],
  sol: "1: …jeder Stand nur eine Stimme hat. 2: …sich abgesprochen und gemeinsam mit Nein gestimmt. 3.*: …der Dritte Stand ca. 97 % der Menschen stellt, aber nur 1 von 3 Stimmen hat.",
}));
d.push(sp());
d.push(note("SOLLBRUCHSTELLE", "Wenn die Zeit knapp wird, entfällt Satz 3 (*). Phase 3d bleibt, Phase 4 wird dann auf Merksatz und Abschlussfrage gekürzt. Satz 3 kann in der nächsten Stunde als Wiederholung wieder aufgegriffen werden."));

// Phase 4
d.push(phase("4", "Auswerten", 6, "AVIVA: Auswerten"));
d.push(warum("Der Lernzuwachs wird gesichert und überprüft: Merksatz, Lernziele, offene Abschlussfrage, die ein Urteil mit Sachargument verlangt (Vorstufe zu den Urteilsaufgaben der Reihe). Diese Phase wird nie gestrichen.", "„Nächste Stunde: Was macht der Dritte Stand, wenn er überstimmt wird?“"));
d.push(sp());
d.push(step("schreibt ab", {
  say: "Jetzt schreibt ihr ab: Merksatz 1. Rahmt ihn rot ein.",
  do: [
    "Folie 14 (Merksatz, 2 Minuten): Kernsatz im roten Rahmen, Beispiel, Test. Alle schreiben ab. Danach „Vergleicht und verbessert mit Grün.“",
    "Folie 15 (3 Minuten): Abschlussfrage mit Zeichnung (König Ludwig XVI. mit Sprechblase, Berater mit „Ich rate dir, …, weil …“): „Du bist Berater von König Ludwig XVI. Er fragt dich: Soll ich erlauben, dass der Dritte Stand (Bauern und Bürger) nach Köpfen abstimmt?“ Rat im Heft: „Ich rate dir, …, weil …“. Es gibt mehrere gute Antworten, geprüft werden Position, ein Sachargument aus der Stunde und ein „weil“. Erst zehn Sekunden denken, dann schreiben.",
    "Form nach Verlauf wählen: (a) Blitzlicht: je eine Stimme „erlauben“ und eine „nicht erlauben“ gezielt aufrufen, kein Kommentar (wenig Zeit, Klasse hört gut zu). (b) Schriftlich, dann zwei bis drei Sätze vorlesen (Klasse unruhig, Heft-Stichprobe bei 4 bis 6 Personen). (c) Berater-Duell zu zweit: A rät dafür, B dagegen, leise (Ruhe und Zeit). Unbenotet.",
    "Folie 16 (30 Sekunden, still): Haben wir die Ziele erreicht? Keine Antworten ins Plenum. Folie 17 (30 Sekunden, nur lesen): Ausblick und Hausaufgabe. „Was passiert, wenn der Dritte Stand fordert, dass nach Köpfen abgestimmt wird, und der König ablehnt? Das sehen wir in der nächsten Stunde.“ Hausaufgabe: Testaufgaben mit Grün berichtigen (falls der Test zurückgegeben wurde).",
  ],
  sol: "Mögliche Argumente (Auswahl, keine Musterlösung). Erlauben: Der Staat braucht Geld, Steuern gehen nur mit dem Dritten Stand (Bauern und Bürger); er stellt fast alle Menschen, sonst droht Unruhe; Hunger und Staatsbankrott sind nah. Nicht erlauben: Adel und Klerus sind wütend (Steuerplan 1787 gescheitert), der König verliert Rückhalt und Macht; der Dritte Stand wäre zu mächtig; seit 1614 wurde nach Ständen abgestimmt. Zwischenlösungen (erlauben, aber nur für die Steuerfrage) sind gut begründbar. Ein Hinweis, falls gefragt: Historisch lehnte der König ab. Kriterien: Position, ein Sachargument aus der Stunde, ein „weil“.",
}));
d.push(sp());
d.push(note("ZUSATZ (Puffer, nur bei Zeit)", "Folie 18: Sportverein-Beispiel (Übertragung): Ein Verein hat drei Abteilungen: Fußball 200 Mitglieder, Tischtennis 20, Schach 10. Jede Abteilung hat bei der Abstimmung eine Stimme. Was ist unfair? Einzelarbeit, Heft, 2 Minuten.", "F4F7FD", NAVY, "B9C6E8"));

d.push(h3("So sieht das Heft am Ende der Stunde aus"));
d.push(table([2300, 7906], [
  ["Überschrift", "Generalstände 1789 · Datum"],
  ["Ü1", "Drei Stände mit Zahlen, Vorrechte, „Die Steuern zahlt der …“ (Verbesserungen grün)"],
  ["Zeitstrahl", "Blatt eingeklebt, Ereignisse zu den Daten eingetragen (1614 bis 5. Mai 1789), Verbesserungen grün"],
  ["Ständeversammlung", "Zwei Argumente, eine Beobachtung, Vermutung zum Ausgang, Ergebnis (Zahlen von der Tafel)"],
  ["Ü3", "Drei Sätze (Satz 3 optional)"],
  ["Merksatz 1", "Abstimmung nach Ständen – rot umrahmt"],
  ["Abschlussfrage", "Rat an den König: „Ich rate dir, …, weil …“ (Position, Sachargument, weil)"],
].map(([a, b]) => row([tc(a, 2300, { bold: true, fill: LIGHT }), tc(b, 7906)], 340))));

d.push(h3("Nach der Stunde festhalten (Reflexionslog)"));
d.push(...bullets(["Passten Ziel, Inhalt und Methode zusammen? Wie lief die Zeit? Wie viel echte Lernzeit hatte die Klasse?", "Wo wurde es laut, und welche Ansage half? Welche Abfrage funktionierte, welche nicht?", "Wie viele nannten in der Abschlussfrage ein Sachargument aus der Stunde (Stichprobe im Heft)? Welche Form (Blitzlicht, schriftlich, Duell) trug am besten?", "Rückmeldung der Klasse: ein Satz im Heft „Das war klar / Das hat gefehlt“ (nur falls Zeit)."]));
d.push(h3("Zum Prüfen"));
d.push(...bullets([
  "Zahlen und Daten (Abgeordnete 290/270/580, Krieg in Amerika, 1787, August 1788) stammen aus dem Text der 8C und sind gegen das Schulbuch zu lesen. Die Anzahl der Menschen (150 000 / 500 000 / 20 000 000) stammt aus Ihrem Gruppenpuzzle-Material.",
  "Das Spiel vereinfacht: 1789 wurde nicht über genau diesen Antrag abgestimmt. Der Streit ging um die Abstimmungsregel (nach Ständen oder nach Köpfen) und blockierte die Generalstände. Die Karten sagen „Auch Adel und Klerus zahlen künftig Steuern“, weil das der Kern der Steuerfrage war. Für die nächste Stunde (Dritter Stand fordert „nach Köpfen“, Nationalversammlung) lässt sich direkt daran anknüpfen.",
  "Die Standeskarten verteilen die Rollen: 3 Klerus (12 %), 4 Adel (16 %), 18 Dritter Stand (72 %). Das ist verzerrt, 1789 waren Klerus und Adel zusammen etwa 3 %. Das Verhältnis der Menschen steht auf Folie 12. Bei der Auswertung kann man darauf hinweisen: In Frankreich war der Unterschied noch viel größer.",
  "Stimmen im Dritten Stand einzelne mit Nein, bleibt das Ergebnis nach Ständen gleich (die Mehrheit im Stand entscheidet), nach Köpfen wird es knapper. Das ist kein Fehler, sondern zeigt, dass auch ein Stand nicht einheitlich denkt (Hinweis für die nächste Stunde).",
]));

(async () => {
  const doc = new Document({ styles: { default: { document: { run: { font: "Calibri", size: 22 } } } }, sections: [{ properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 1000, bottom: 850, left: 850, right: 850 } } }, headers: { default: header }, footers: { default: footer }, children: d }] });
  fs.writeFileSync("../Tafelskript_Generalstaende_8D.docx", await Packer.toBuffer(doc));
  console.log("ok");
})();
