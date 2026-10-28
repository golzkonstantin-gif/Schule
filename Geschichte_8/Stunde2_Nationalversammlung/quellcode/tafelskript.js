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
const header = new Header({ children: [new Paragraph({ tabStops: [{ type: "right", position: W }], children: [t("Geschichte · Klasse 8C · Stunde 2", { size: 17, color: MUTED }), t("\tfür die Lehrkraft", { size: 17, color: MUTED })] })] });
const footer = new Footer({ children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ font: "Calibri", size: 16, color: MUTED, children: ["Tafelskript 8C · Stunde 2 · Seite ", PageNumber.CURRENT] })] })] });

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
d.push(kicker("Tafelskript für die Lehrkraft · Einzelstunde (45 Min.) · 12 Schülerinnen und Schüler · Lehrkraft 2 hospitiert nur · Freitag 06.11.2026"), h1("Der Dritte Stand handelt: Nationalversammlung, Ballhausschwur, Karikatur"));
d.push(p(t("Ziel: Die Klasse sichert zuerst die Abstimmung nach Ständen und nach Köpfen (noch nicht sicher), erarbeitet dann aus einem bearbeiteten Buchtext, wie sich der Dritte Stand zur Nationalversammlung erklärt und den Ballhausschwur leistet, und lernt am Modell, eine Karikatur in drei Schritten zu analysieren. Aufbau nach AVIVA, Methode direkt, dann angeleitet.", { size: 21 }), { after: 80 }));

d.push(h3("Bezug zur Reflexion (Stunde 1, 16.10.)"));
d.push(...bullets([
  "Stunde 1 lief inhaltlich, der Text mit Zeitstrahl war gut. Rückmeldung: Die Stunde war unterrichtsgesprächslastig, das Schulbeispiel zu abstrakt. Deshalb läuft diese Stunde über Aufträge, schriftliche Antworten und gezielte Abfragen statt Gespräch.",
  "Die Abstimmung nach Ständen und nach Köpfen sitzt noch nicht sicher. Deshalb wiederholen wir genau das: Fehlertext (Folie 4 und 5) und ein Merkbild (Folie 6).",
  "Eine zweite Lehrkraft hospitiert nur. Die Hilfekarten liegen am Pult, die Betreuung der Partnerarbeit liegt allein bei Ihnen.",
]));

d.push(h3("Kopf der Planung"));
d.push(table([2600, 7606], [
  ["Ziel (Ich kann …)", "… erklären, warum sich der Dritte Stand (Bauern und Bürger) 1789 zur Nationalversammlung erklärte. … erklären, was die Abgeordneten im Ballhausschwur versprachen. … eine Karikatur analysieren: beschreiben, deuten, beurteilen.*"],
  ["Prüfaufgabe: Woran erkenne ich es?", "Zeitstrahl ausgefüllt, Merksatz im Heft (Ziel 1 und 2). Kurzanalyse der Karikatur „Für die Mühsal geboren“ mit Dreischritt (Ziel 3). Abschlussfrage im Berater-Format."],
  ["Mindestziel", "Ziel 1 und 2 und die Schritte 1 und 2 der Karikatur-Analyse. Alles mit * (Schritt 3 „Beurteilen“) nur bei Zeit; er kann in Stunde 3 wiederholt werden. „Auswerten“ bleibt immer."],
  ["Beitrag zur Reihe", "Stunde 2 von 11 (Sequenz Französische Revolution). Der Dritte Stand wird handelnde Kraft, der Streit um die Abstimmung führt zur Nationalversammlung. Gelenk zu Stunde 3: Der König zieht Soldaten zusammen, in Paris wächst die Angst (Bastille). Die Karikatur-Methode wird in Stunde 7 an „Robespierre als Henker“ wiederholt."],
  ["Vorgehen", "Direkt bei der neuen Methode (Karikatur, Modell laut denken), indirekt beim Text und beim Zeitstrahl (Selbsterarbeitung), angeleitet in der Partnerarbeit mit Hilfekarten. Direkt oder indirekt wird je Phase entschieden."],
].map(([a, b]) => row([tc(a, 2600, { fill: LIGHT, bold: true }), tc(b, 7606)], 340))));

d.push(sp());
d.push(table([800, 2300, 4300, 2806], [
  row([th("Min.", 800), th("AVIVA", 2300), th("Phase", 4300), th("Arbeitsform", 2806)], 320),
  ...[["3", "Ankommen", "0  Ankommen, Ziele, Eingangsdiagnose", "still, allein"], ["6", "Vorwissen", "1  Fehlertext: Abstimmung nach Ständen und Köpfen", "allein, Partnercheck, Folie"], ["12", "Informieren", "2  Text, Quelle, Zeitstrahl, Bild, Merksatz", "allein, Folie, abschreiben"], ["15", "Verarbeiten", "3  Karikatur: Modell, Partnerarbeit, Sicherung", "Lehrkraft, zu zweit"], ["6", "Auswerten", "4  Berater-Frage, Ziele prüfen, Ausblick", "allein, gezielte Abfrage"], ["3", "—", "Puffer", "—"]].map((r) => row(r.map((v, i) => tc(v, [800, 2300, 4300, 2806][i], { bold: i < 3 })), 320)),
]));

d.push(h3("Redeanteil und Ansagen"));
d.push(table([3300, 6906], [
  ["Gespräch ersetzen", "Antworten schriftlich, per Folie, gezielte Abfrage (2 bis 3 Personen, kurz). Kein offenes Plenumsgespräch."],
  ["Partnerarbeit", "Nur in Phase 3 (Karikatur) und beim Partnercheck in Phase 1. Wörtlich: „Flüsterstufe. Nur dein Partner hört dich.“"],
  ["Stichproben", "Im Heft mitlesen bei 4 bis 6 Personen unterschiedlichen Niveaus, unbenotet, getrennt von den mündlichen Noten. Im Wechsel mit der gezielten Abfrage."],
  ["Kreide-Kette", "Nur, wenn Zeit ist, beim Zeitstrahl (12 Schülerinnen und Schüler). Sonst Partnervergleich und Folie."],
].map(([a, b]) => row([tc(a, 3300, { fill: LIGHT, bold: true }), tc(b, 6906)], 340))));
d.push(sp());
d.push(table([4700, 5506], [
  ["„Stifte liegen. Ihr hört nur zu.“", "Sie erklären oder modellieren, niemand schreibt."],
  ["„Jetzt arbeitet ihr allein. Die Aufgabe steht rechts.“", "Einzelarbeit."],
  ["„Arbeitet zu zweit. Flüsterstufe.“", "Partnerarbeit."],
  ["„Jetzt schreibt ihr ab: …“", "Merksatz ins Heft."],
  ["„Vergleicht und verbessert mit Grün.“", "Kontrolle im eigenen Heft."],
].map(([a, b]) => row([tc(a, 4700, { fill: LIGHT, bold: true }), tc(b, 5506)], 340))));

d.push(h3("Material und Vorbereitung"));
d.push(...bullets([
  "Arbeitsblatt „Der Dritte Stand handelt“ (Text, Quelle von Sieyès, Zeitstrahl): 14 Stück (12 plus 2 Reserve), einseitig.",
  "Schnipsel „Karikatur analysieren: Dreischritt“ (Seite 1): 6 Schnipsel, ausschneiden, je einer pro Paar. Hilfekarten Stufe 1 bis 3 (Seite 2): einmal drucken, ausschneiden, am Pult bereitlegen (Hilfe nur auf Nachfrage).",
  "Folien 1 bis 17 (Schrift mindestens 24 Punkt). Das Bild „Für die Mühsal geboren“ (S. 111), das Bild des Ballhausschwurs (S. 113) und die Karikatur zu den drei Ständen aus dem Absolutismus sind auf den Folien. Lehrbuch S. 111, 113 und 116 liegen zur Hand.",
  "Seitentafel: Ziele und Ablauf mit * bei Schritt 3 der Karikatur-Analyse. Mitte der Tafel leer lassen. Kreide, Grün zum Verbessern, Uhr oder Timer.",
  "Material vorab auf die Tische legen (Arbeitsblatt, Schnipsel), damit keine Austeilaktion Zeit kostet.",
  "Zahlen: wie in Stunde 1 (Gedankenspiel 580 zu 560). Siehe Prüfhinweise am Ende.",
]));

// Phase 0
d.push(phase("0", "Ankommen", 3, "AVIVA: Ankommen"));
d.push(warum("Die Klasse wird ruhig, kennt Ziel und Ablauf und beantwortet zu Beginn die Reihenfrage, die in Stunde 11 wieder aufgegriffen wird (Eingangsdiagnose, direkt, von der Lehrkraft geführt).", "„Zuerst schauen wir, was ihr aus der letzten Stunde noch wisst, denn damit beginnt der Streit, um den es heute geht.“"));
d.push(sp());
d.push(step("kommt an, schreibt leise", {
  say: "Stifte liegen. Ihr hört nur zu. Schaut auf die Folie. Schreibt Datum und Überschrift ins Heft und beantwortet dann die Frage in zwei Sätzen.",
  do: [
    "Folie 1 steht beim Betreten da. Zwei Sätze im Heft: „Was ist eine Revolution? Ist sie gut oder schlecht? Warum?“ Die Seite bleibt im Heft und wird in Stunde 11 mit dem Gesamturteil verglichen (Vorher-Nachher, keine Selbsteinschätzung). Sie lesen die Sätze beim Durchgehen mit, nicht kommentieren.",
    "Folie 2 (Ziele) und Folie 3 (Ablauf nach AVIVA) vorlesen, 1 Minute. Hinweis: „Was wir eventuell nicht schaffen, hat einen Stern.“",
  ],
  extra: [boardBox("Seitentafel · Ziel und Ablauf", [
    bl("Ich kann erklären, warum sich der Dritte Stand 1789 zur Nationalversammlung erklärte.", { after: 20 }),
    bl("Ich kann erklären, was die Abgeordneten im Ballhausschwur versprachen.", { after: 20 }),
    bl("Ich kann eine Karikatur analysieren: beschreiben, deuten, beurteilen.*", { after: 60 }),
    bl("A Ankommen · V Abstimmung wiederholen · I Text und Zeitstrahl · V Karikatur (Schritt 3 *) · A Berater-Frage", { after: 0 }),
  ]), gapP()],
}));

// Phase 1
d.push(phase("1", "Fehlertext: Abstimmung wiederholen", 6, "AVIVA: Vorwissen aktivieren"));
d.push(warum("Nach Ständen und nach Köpfen sicher abrufen: Die Regel erklärt, warum der Dritte Stand handelt (direkt: allein, Partnercheck, Kontrolle per Folie).", "„Der Dritte Stand hat gegen die Regel verloren, aber er gibt nicht auf. Wie reagiert er? Das lest ihr jetzt.“"));
d.push(sp());
d.push(step("arbeitet allein, dann zu zweit", {
  say: "Jetzt arbeitet ihr allein. Die Aufgabe steht rechts.",
  do: ["Folie 4 (Ü1, Fehlertext). 3 Minuten allein im Heft, dann 1 Minute Partnercheck (leise): „Vergleicht und verbessert mit Grün.“ Dann Folie 5 (Lösung, 1 Minute) und Folie 6 (Merkbild „Nach Ständen oder nach Köpfen?“, 1 Minute). Gezielte Abfrage bei 2 bis 3 Personen: „Wie viele Stimmen hat der Dritte Stand nach Ständen? Und nach Köpfen?“ Kein Gespräch."],
  extra: [folieBox("4 · Ü1 Fehlertext", [p(t("Ü1 Abstimmung bei den Generalständen", { font: "Cambria", size: 22, bold: true, color: NAVY }), { after: 40 }), bl("1. Bei den Generalständen hat jeder Abgeordnete eine Stimme."), bl("2. Klerus und Adel haben zusammen zwei von drei Stimmen."), bl("3. Der Dritte Stand (Bauern und Bürger) stellt etwa 3 % der Menschen."), bl("4. Bei der Abstimmung nach Köpfen hätte er weniger Gewicht gehabt."), bl("5. Der König erlaubte die Abstimmung nach Köpfen.", { after: 40 }), bl("Aufgabe: Vier Sätze enthalten einen Fehler. Unterstreiche die Fehler und schreibe diese Sätze richtig ins Heft.", { after: 40 }), p([t("Zeit: ", { bold: true, size: 19 }), t("4 Minuten · allein, dann Partnercheck · leise     ", { size: 19 }), t("Fertig? Stift hinlegen.", { size: 19, italics: true })], { after: 0 })]), gapP()],
  sol: "Fehler in 1 (jeder Stand eine Stimme, nicht jeder Abgeordnete), 3 (etwa 97 %), 4 (nach Köpfen mehr Gewicht), 5 (der König lehnte ab). Satz 2 ist richtig. Merkbild: Nach Ständen Klerus Nein, Adel Nein, Dritter Stand Ja = 2 zu 1, abgelehnt. Nach Köpfen Ja 580, Nein 560 = knapp angenommen (Gedankenspiel; 1789 wurde nicht nach Köpfen abgestimmt).",
}));
d.push(sp());
d.push(note("SOLLBRUCHSTELLE", "Läuft die Klasse langsam an, entfällt der Partnercheck, und Folie 5 wird sofort gezeigt (1 Minute gewonnen)."));

// Phase 2
d.push(phase("2", "Text, Quelle, Zeitstrahl, Bild, Merksatz", 12, "AVIVA: Informieren"));
d.push(warum("Wie der Dritte Stand die Macht des Königs infrage stellt, wird selbst aus einem bearbeiteten Buchtext erarbeitet (indirekt: Lesen allein, Kontrolle per Folie). Das Bild des Ballhausschwurs zeigt, wie die Abgeordneten ihren Schwur darstellten. Der Merksatz sichert das Wichtigste.", "„Der Dritte Stand hat sich selbst zur Nationalversammlung gemacht. Wie ein Zeichner damals die Lage darstellte, schauen wir uns gleich genau an.“"));
d.push(sp());
d.push(step("arbeitet allein, vergleicht", {
  say: "Ich teile das Arbeitsblatt aus. Jetzt arbeitet ihr allein. Die Aufgabe steht rechts.",
  do: [
    "Folie 7 (Ü2). 8 Minuten: Text lesen, Daten unterstreichen, Zeitstrahl ausfüllen, ein Satz zu ALLES mit einer Zahl. Nicht erklären, durch die Reihen gehen.",
    "Warnung vor dem Fehler: Das Lehrbuch nennt den 20. Juni für die Nationalversammlung. Historisch ist es der 17. Juni (Nationalversammlung) und der 20. Juni (Ballhausschwur). Das sagen Sie vorab in einem Satz, damit die Klasse nicht verunsichert wird, wenn sie im Buch nachschaut.",
    "Folie 8 (Zeitstrahl, 1 Minute) nach dem Partnervergleich: „Vergleicht und verbessert mit Grün.“ Folie 9 (Bild Ballhausschwur, 1 Minute), keine Besprechung. Folie 10 (Merksatz, 2 Minuten): „Jetzt schreibt ihr ab: Merksatz 1. Rahmt ihn rot ein.“",
  ],
  extra: [folieBox("7 · Ü2 Text, Quelle und Zeitstrahl", [p(t("Ü2 Der Dritte Stand handelt", { font: "Cambria", size: 22, bold: true, color: NAVY }), { after: 40 }), bl("1. Lies den Text. Unterstreiche alle Daten."), bl("2. Schreibe zu jedem Datum das Ereignis in den Zeitstrahl."), bl("3. Schreibe ins Heft: Sieyès sagt, der Dritte Stand sei ALLES. Erkläre das mit einer Zahl.", { after: 40 }), p([t("Zeit: ", { bold: true, size: 19 }), t("8 Minuten · allein · leise     ", { size: 19 }), t("Fertig? Stift hinlegen.", { size: 19, italics: true })], { after: 0 })]), gapP(),
    boardBox("Mitte · Zeitstrahl (Lösung, Folie 8)", [
      bl("Winter 1788/89    Sieyès: Der Dritte Stand ist ALLES"),
      bl("5. Mai 1789    Eröffnung, Streit um die Abstimmung"),
      bl("17. Juni 1789    Nationalversammlung"),
      bl("20. Juni 1789    Ballhausschwur"),
      bl("Anfang Juli 1789    Soldaten um Paris, Angst in Paris", { after: 0 }),
    ]), gapP(),
    folieBox("10 · Merksatz 1 – abschreiben, rot umrahmen", [
      p(t("Nationalversammlung und Ballhausschwur", { bold: true, color: RED, font: "Cambria", size: 22 }), { after: 40 }),
      p(t("Am 17. Juni 1789 erklärt sich der Dritte Stand (Bauern und Bürger) zur Nationalversammlung. Am 20. Juni schwören die Abgeordneten im Ballhaus, nicht auseinanderzugehen, bis Frankreich eine Verfassung hat. Damit stellt der Dritte Stand die Macht des Königs infrage.", { font: "Cambria", size: 21 }), { after: 50 }),
      p(t("Beispiel: Sieyès: Der Dritte Stand ist ALLES, bisher war er NICHTS, er will ETWAS sein.", { size: 20, color: MUTED }), { after: 50 }),
      p(t("Test: Wer erklärte sich zur Nationalversammlung? Was schwor sie?", { size: 20, color: MUTED }), { after: 0 }),
    ]), gapP()],
  sol: "Zu 3: Der Dritte Stand stellt etwa 97 % der Menschen (rund 20 Mio. von 20,65 Mio.), deshalb sagt Sieyès, er sei ALLES. Zeitstrahl wie auf der Tafel oben. Falls Zeit ist, Kreide-Kette beim Zeitstrahl mit Mitschreiben der anderen.",
}));
d.push(sp());
d.push(note("SOLLBRUCHSTELLE", "Wird es knapp, entfällt das Bild (Folie 9) oder wird in 30 Sekunden gezeigt. Merksatz und Zeitstrahl bleiben."));

// Phase 3
d.push(phase("3", "Karikatur: Modell, Partnerarbeit, Sicherung", 15, "AVIVA: Verarbeiten"));
d.push(warum("Neue Kompetenz (SchiC: Analyse einer Karikatur): erst zeigen und laut denken, dann angeleitet zu zweit üben, danach sichern. Direktes Vorgehen, weil die Methode neu ist; Öffnung erst in der Partnerarbeit mit Hilfen.", "„Ihr wisst jetzt, was der Dritte Stand will. Wie ein Zeichner das Leid der Bauern darstellte, untersuchen wir jetzt Schritt für Schritt.“"));
d.push(sp());
d.push(step("hört zu, schaut", {
  say: "Stifte liegen. Ihr hört nur zu. Ich zeige euch, wie man eine Karikatur untersucht.",
  do: [
    "3a  (1 Minute, Folie 11) Dreischritt nennen: beschreiben, deuten, beurteilen. Eine Karikatur übertreibt, um eine Meinung zu zeigen. Warnung vor dem Fehler: Nicht sofort deuten, erst genau beschreiben. Schnipsel liegen auf den Tischen.",
    "3b  (3 Minuten, Folie 12, Modell, laut denken): „Ich sehe einen alten Bauern mit Sense, gebückt. Auf seinem Rücken sitzen ein Geistlicher mit Kreuz und ein Adliger mit Degen und Federhut. Unten fressen Hasen am Kohl, Hühner picken. Das beschreibe ich zuerst. Jetzt deute ich: Der Bauer steht für den Dritten Stand (Bauern und Bürger). Er trägt Klerus und Adel, das heißt, er zahlt die Steuern, die beiden anderen nicht. Die Hasen am Kohl stehen für das Jagdrecht des Adels: Das Wild frisst die Felder der Bauern. Die Beischrift sagt, dass das Spiel bald enden soll. Und jetzt beurteile ich: Der Zeichner steht auf der Seite des Dritten Standes. Er übertreibt, es gab auch arme Adlige und Geistliche. Aber er trifft etwas Richtiges: Klerus und Adel zahlten kaum Steuern. Er will, dass sich etwas ändert, 1789, vor den Generalständen.“",
    "Prüffrage (gezielte Abfrage, 2 Personen, kurz): „Welchen Schritt habe ich zuerst gemacht? Wofür steht der Bauer?“",
  ],
  sol: "Bild: Karikatur zu den drei Ständen, 1789 (kennt die Klasse aus dem Absolutismus). Beischrift im Dialekt: „Man muss hoffen, dass dieses Spiel bald endet.“",
}));
d.push(sp());
d.push(step("arbeitet zu zweit, Flüsterstufe", {
  say: "Arbeitet zu zweit. Flüsterstufe. Nur dein Partner hört dich. Die Aufgabe steht rechts.",
  do: [
    "3c  (7 Minuten, Folie 13) Partnerarbeit an „Für die Mühsal geboren“. Hilfekarten Stufe 1 bis 3 liegen am Pult und werden nur auf Nachfrage geholt. Sie gehen durch die Reihen und lesen bei 4 bis 6 Paaren mit (unbenotet).",
    "Warnung vor dem Fehler: Schritt 2 nicht überspringen. Nicht alle Bildelemente sind sicher zu deuten (z. B. der Hahn auf dem Stock). Offen lassen ist besser als raten.",
  ],
  extra: [folieBox("13 · Ü3 „Für die Mühsal geboren“", [p(t("Ü3 Ihr analysiert die Karikatur", { font: "Cambria", size: 22, bold: true, color: NAVY }), { after: 40 }), bl("1. Beschreibe die Karikatur in drei Sätzen."), bl("2. Deute: Wofür steht der Bauer? Was sagt der Bildtext?"), bl("3.* Beurteile: Ist die Darstellung berechtigt?", { after: 40 }), p([t("Zeit: ", { bold: true, size: 19 }), t("7 Minuten · zu zweit · Flüsterstufe     ", { size: 19 }), t("Fertig? Stift hinlegen.", { size: 19, italics: true })], { after: 0 })]), gapP()],
  sol: "Bildtext (Buch S. 111, Q3): „Jeden Tag mitten auf dem Feld – in Hitze und Kälte – sieht man den armen Bauern – arbeitend das ganze Jahr über – um mit seiner Arbeit das Geld zu verdienen – und damit den Steuereintreiber zu bezahlen.“ Titel: „Ne pour la peine“ (Für die Mühsal geboren), Unterschrift: „L'Homme de Village“.",
}));
d.push(sp());
d.push(step("schaut zu, vergleicht", {
  say: "Stifte liegen. Wir vergleichen. Danach verbessert ihr mit Grün.",
  do: [
    "3d  (3 Minuten, Folie 14) Sicherung: gezielte Abfrage bei 3 Paaren, je ein Schritt (Beschreiben, Deuten, Beurteilen), kein Gespräch. Danach Folie 14 mit möglichen Ergebnissen: „Vergleicht mit eurer Analyse und verbessert mit Grün.“",
  ],
  sol: "Erwartung (Auswahl): 1 großer Bauer mit Hut, Stock, Sack und Werkzeug, Dorf mit Kirche, Tiere, ein weiterer Arbeiter, Titel und Bildtext. 2 Der Bauer steht für alle Bauern und den Dritten Stand (Bauern und Bürger), er arbeitet das ganze Jahr, um Steuern zu zahlen. 3 Der Zeichner will Mitleid und Veränderung; das passt zu den Beschwerden, übertreibt aber, weil nicht alle Bauern gleich arm waren. Gesamtbild: Dritter Stand zahlt, Klerus und Adel zahlen kaum.",
}));
d.push(sp());
d.push(note("SOLLBRUCHSTELLE", "Wird die Zeit knapp, entfällt Schritt 3 „Beurteilen“ (*) in der Partnerarbeit und wird in 1 Minute gemeinsam gemacht oder in Stunde 3 wiederholt. Schritt 1 und 2 und die Sicherung bleiben."));

// Phase 4
d.push(phase("4", "Auswerten", 6, "AVIVA: Auswerten"));
d.push(warum("Der Lernzuwachs wird überprüft: offene Abschlussfrage mit mehreren guten Antworten, die ein Urteil mit Sachargument verlangt (Vorstufe zu den Urteilsaufgaben der Reihe), dann Ziele prüfen. Diese Phase wird nie gestrichen.", "„Nächste Stunde: Der König zieht Soldaten zusammen, was passiert in Paris?“"));
d.push(sp());
d.push(step("arbeitet allein", {
  say: "Stifte liegen. Ihr hört nur zu. Das ist die letzte Aufgabe. Jetzt arbeitet ihr allein. Die Aufgabe steht rechts.",
  do: [
    "Folie 15 (3 Minuten): Abschlussfrage mit Zeichnung (Abgeordneter mit Sprechblase, Berater „Du“): „Du bist Berater der Abgeordneten des Dritten Standes (Bauern und Bürger). Sie fragen dich: Der König hat den Saal sperren lassen. Sollen wir nach Hause gehen oder schwören, zusammenzubleiben?“ Rat im Heft: „Ich rate euch, …, weil …“. Es gibt mehrere gute Antworten, geprüft werden Position, ein Sachargument aus der Stunde und ein „weil“. Erst zehn Sekunden denken, dann schreiben.",
    "Form nach Verlauf wählen: (a) Blitzlicht: je eine Stimme „bleiben und schwören“ und eine „nach Hause“ gezielt aufrufen, kein Kommentar. (b) Schriftlich, dann zwei bis drei Sätze vorlesen, Heft-Stichprobe bei 4 bis 6 Personen. (c) Berater-Duell zu zweit: A rät dafür, B dagegen, leise. Unbenotet.",
    "Folie 16 (30 Sekunden, still): Haben wir die Ziele erreicht? Keine Antworten ins Plenum. Folie 17 (30 Sekunden): Ausblick auf die Bastille.",
  ],
  sol: "Mögliche Argumente (Auswahl, keine Musterlösung). Bleiben und schwören: Sie vertreten fast alle Menschen (Sieyès: ALLES); nur gemeinsam sind sie stark; sonst bleibt alles beim Alten; ohne Verfassung kann der König weiter allein entscheiden. Nach Hause gehen: Der König hat den Saal gesperrt, Gefahr durch Soldaten, Unruhe vermeiden, später neu verhandeln. Zwischenlösung (bleiben, aber verhandeln) ist gut begründbar. Historisch trafen sich die Abgeordneten in der Ballhalle und schworen. Kriterien: Position, ein Sachargument aus der Stunde, ein „weil“.",
}));

d.push(h3("So sieht das Heft am Ende der Stunde aus"));
d.push(table([2300, 7906], [
  ["Überschrift", "Der Dritte Stand handelt · Datum"],
  ["Eingangsdiagnose", "Zwei Sätze zu „Was ist eine Revolution? Gut oder schlecht? Warum?“ (bleibt bis Stunde 11)"],
  ["Ü1", "Fehlertext: vier verbesserte Sätze (Verbesserungen grün)"],
  ["Arbeitsblatt", "Eingeklebt: Text, Zeitstrahl ausgefüllt, ein Satz zu ALLES mit einer Zahl"],
  ["Merksatz 1", "Nationalversammlung und Ballhausschwur – rot umrahmt"],
  ["Ü3", "Analyse „Für die Mühsal geboren“ (Beschreiben, Deuten, Beurteilen*), mit Grün verbessert"],
  ["Abschlussfrage", "Rat an die Abgeordneten: „Ich rate euch, …, weil …“"],
].map(([a, b]) => row([tc(a, 2300, { bold: true, fill: LIGHT }), tc(b, 7906)], 340))));

d.push(h3("Nach der Stunde festhalten (Reflexionslog)"));
d.push(...bullets(["Passten Ziel, Inhalt und Methode zusammen? Wie lief die Zeit? Wie viel echte Lernzeit hatte die Klasse?", "Sitzt die Abstimmung nach Ständen und nach Köpfen jetzt? (Fehlertext, Abfrage)", "Konnten die Paare die Karikatur beschreiben und deuten? Wer brauchte welche Hilfekarte?", "Wie viele nannten in der Abschlussfrage ein Sachargument aus der Stunde? Welche Form trug am besten?", "Rückmeldung der 8D (Stunde 1) und der 8C abgleichen."]));

d.push(h3("Zum Prüfen"));
d.push(...bullets([
  "Daten und Abläufe (17. und 20. Juni, Sieyès, Soldaten und Entlassung des Finanzministers Anfang Juli) stammen aus dem Lehrbuch (S. 112/113) und aus meinem Wissen. Bitte gegen das Buch lesen. Das Buch nennt für die Nationalversammlung den 20. Juni, die Zeitleiste im Buch den 17. Juni. Wir folgen der Zeitleiste.",
  "Sieyès: Das Buch sagt, die Kampfschrift wurde 1788 verfasst; erschienen ist sie im Winter 1788/89. Der Text und der Zeitstrahl sagen deshalb „Winter 1788/89“. Das Zitat ist gekürzt.",
  "Zahlen: Das Merkbild nutzt das Gedankenspiel aus Stunde 1 (580 zu 560, nach Ständen 2 zu 1). Das Buch nennt etwa 600 zu je etwa 300 Abgeordneten. Die Zahlen bleiben wie in der 8C-Stunde 1 und in der 8D (Vergleichbarkeit).",
  "Bilder: Die Bilder (Buch S. 111 und S. 113) sind aus dem Scan übernommen und nur für den Unterricht gedacht. Die Zeichnung des Abgeordneten und des Beraters ist selbst gezeichnet, kein historisches Porträt. Die Deutung einzelner Bildelemente („Für die Mühsal geboren“, z. B. Hahn, Tiere) ist nicht sicher, die Erwartungen sind eine Auswahl.",
  "Zeit: Die Karikatur-Phase hat 15 Minuten, der Puffer 3. Schritt 3 „Beurteilen“ ist mit * markiert.",
]));

(async () => {
  const doc = new Document({ styles: { default: { document: { run: { font: "Calibri", size: 22 } } } }, sections: [{ properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 1000, bottom: 850, left: 850, right: 850 } } }, headers: { default: header }, footers: { default: footer }, children: d }] });
  fs.writeFileSync("../Tafelskript_Stunde2_Nationalversammlung.docx", await Packer.toBuffer(doc));
  console.log("ok");
})();
