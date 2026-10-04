// Tafelscript: Tafelbild (aus tafelbild_schrittfolge.js), Ablauf mit Impulsen, Erwartungshorizont mit Musterfließtext
const fs = require("fs");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, BorderStyle,
  ShadingType, AlignmentType, VerticalAlign, HeightRule, Footer, PageNumber, PageOrientation,
} = require("docx");
const { board } = require("./tafelbild_schrittfolge");

const NAVY = "1E2761", MUTED = "5B6B8C", LIGHT = "F4F7FD", LIGHT2 = "EAF0FA", RED = "D9534F";
const W = 15338;

const t = (text, o = {}) => new TextRun({ text, font: o.font || "Calibri", size: o.size || 20, bold: o.bold, italics: o.italics, color: o.color || "000000" });
const p = (runs, o = {}) => new Paragraph({ children: Array.isArray(runs) ? runs : [runs], alignment: o.align, pageBreakBefore: o.pb, spacing: { before: o.before ?? 0, after: o.after ?? 80 }, keepNext: o.keepNext });
const none = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
const solid = (c, s = 6) => ({ style: BorderStyle.SINGLE, size: s, color: c });
const all = (b) => ({ top: b, bottom: b, left: b, right: b });
const cell = (children, o = {}) => new TableCell({
  children: Array.isArray(children) ? children : [children], width: { size: o.w, type: WidthType.DXA },
  shading: o.fill ? { fill: o.fill, type: ShadingType.CLEAR, color: "auto" } : undefined,
  borders: o.borders || all(solid("B9C6E8")), verticalAlign: o.valign || VerticalAlign.TOP,
  margins: { top: 50, bottom: 50, left: 100, right: 100 }, columnSpan: o.span,
});
const table = (widths, rows) => new Table({ width: { size: widths.reduce((a, b) => a + b, 0), type: WidthType.DXA }, columnWidths: widths, rows });
const row = (cells, h) => new TableRow({ children: cells, height: h ? { value: h, rule: HeightRule.ATLEAST } : undefined, cantSplit: true });
const hdr = (txt, w) => cell(p(t(txt, { bold: true, color: "FFFFFF", size: 18 }), { align: AlignmentType.CENTER, after: 0 }), { w, fill: NAVY, valign: VerticalAlign.CENTER });
const kicker = (text, pb) => p(t(text.toUpperCase(), { size: 17, bold: true, color: MUTED }), { after: 20, pb });
const h1 = (text) => p(t(text, { font: "Cambria", size: 34, bold: true, color: NAVY }), { after: 60 });
const box = (paras, fill = LIGHT2) => table([W], [row([cell(paras, { w: W, fill, borders: all(none) })])]);

// ================= Seite 1: Tafelbild =================
const page1 = [
  kicker("Tafelscript · Deutsch 11 · Kafka, „Brief an den Vater“ · Klausurvorbereitung"),
  h1("Tafelbild im Endstand"),
  p(t("① Abschnitt A im Unterrichtsgespräch (Matrix wächst an der Tafel und im Hefter) · ② Abschnitt C allein, dann zusammentragen · ③ These prüfen · ④ Fließtext. Zeilenangaben wie auf dem Arbeitsblatt; Zitate in Originalschreibung.", { size: 18, color: MUTED }), { after: 80 }),
  board,
];

// ================= Seite 2: Ablauf mit Impulsen =================
const aW = [1000, 2400, 6500, 3238, 2200];
const ablauf = [
  ["0–5'", "Einstieg", "Briefanfang vorlesen: „Du hast mich letzthin einmal gefragt, warum ich behaupte, ich hätte Furcht vor Dir …“ Impuls: „Wovor hat ein 36-Jähriger Furcht – und warum schreibt er, statt zu reden?“ Zwei, drei Vermutungen sammeln. Dann These anschreiben: „Kafkas ‚Furcht‘ gründet auf Kommunikationsproblemen.“ Überleitung: „Bevor wir das prüfen, brauchen wir den Hintergrund – den schreibt ihr jetzt mit.“", "Titel, These", "–"],
  ["5–20'", "Diktat + Selbstkontrolle", "Diktat nach der Vorlesefassung (Kontext + Fachbegriffe); danach Whiteboard-Folien, Fehler farbig anstreichen und zählen. Keine Besprechung. Zum Schluss den Fahrplan an die Tafel: Verstehen → Untersuchen → Schreiben.", "Fahrplan", "Vorlesefassung, Whiteboard; Arbeitsblatt austeilen"],
  ["20–42'", "① Abschnitt A im Gespräch", "Erst still lesen (3'). Dann Zeile für Zeile der Matrix:\n• „Was tut der Vater in Z. 4–5? Mit welchen Worten?“ → Redeverbot\n• „Welche Seite der Nachricht ist das? Gibt es überhaupt eine Sachinformation?“ → Appell + Beziehung\n• „Und was passiert ohne Worte?“ → erhobene Hand, analog\n• „Was bewirkt das beim Sohn? Zeigt es mir mit Zeile.“ → stottern, schweigen\n• „Wer ist hier oben, wer unten – und woran seht ihr das?“ → komplementär\n• „Wie deutet der Vater das Schweigen – und wie der Sohn?“ (Z. 15–16) → Interpunktion, Missverständnis\nDie Klasse überträgt jede Zeile in den Hefter.", "Matrix A, Zeile für Zeile", "Hefter"],
  ["42–55'", "② Abschnitt C allein", "Auftrag: Aufgabe 1 auf dem Arbeitsblatt. Herumgehen; wer hängt, bekommt den Hinweis auf die Hilfefragen. Zwei gute Zeilen für das Zusammentragen vormerken.", "–", "Arbeitsblatt S. 2"],
  ["55–65'", "② C zusammentragen", "Reihum je eine Zeile nennen lassen, die nächste Person nimmt sich selbst dran. Korrigieren und präzisieren: „Ist das schon das Modell – oder noch die Handlung?“ Fehlendes ergänzen, v. a. die Kritik über die Mutter (Z. 9–10).", "Matrix C", "Hefter ergänzen"],
  ["65–72'", "③ These prüfen", "„Bestätigen A und C die These?“ Dann die Gegenstimme einbringen, falls sie nicht kommt: Kafka selbst schreibt „Du verstärktest nur, was war“ (S. 6) und spricht von „Deiner Stärke und meiner Schwäche“ (A, Z. 16). Ergebnis: stützt – mit Ergänzung.", "Kasten ③", "–"],
  ["72–88'", "④ Fließtext", "Auftrag: Aufgabe 2. Kurz an Aufbau und Satzbausteine erinnern (Tafel ④). Optional den Musteranfang auf dem Whiteboard zeigen. Wer nicht fertig wird, schreibt zu Beginn der nächsten Stunde weiter – keine Hausaufgabe.", "Kasten ④", "Arbeitsblatt S. 3"],
  ["88–90'", "Ausblick", "Ein, zwei Einleitungen vorlesen lassen. „In der Übungsklausur wendet ihr genau diese Schritte an einem neuen Ausschnitt an.“", "–", "–"],
];
const aRows = [row([hdr("Zeit", aW[0]), hdr("Phase", aW[1]), hdr("Ihre Impulse / was passiert", aW[2]), hdr("Tafel", aW[3]), hdr("Material", aW[4])], 360)];
ablauf.forEach(([z, ph, imp, tafel, mat]) => aRows.push(row([
  cell(p(t(z, { bold: true, color: NAVY, size: 18 }), { after: 0 }), { w: aW[0], fill: LIGHT }),
  cell(p(t(ph, { bold: true, size: 18, color: NAVY }), { after: 0 }), { w: aW[1], fill: LIGHT }),
  cell(imp.split("\n").map((l) => p(t(l, { size: 17 }), { after: 10 })), { w: aW[2], fill: LIGHT2 }),
  cell(p(t(tafel, { size: 17 }), { after: 0 }), { w: aW[3] }),
  cell(p(t(mat, { size: 17, color: MUTED }), { after: 0 }), { w: aW[4] }),
], 380)));

const page2 = [
  kicker("Tafelscript · Ablauf der Doppelstunde", true),
  h1("Ablauf und Impulse"),
  table(aW, aRows),
];

// ================= Seite 3: Erwartungshorizont Fließtext =================
const muster = [
  ["Einleitung", "In seinem „Brief an den Vater“ (1919) versucht Franz Kafka, seinem Vater zu erklären, warum er Furcht vor ihm hat. In den beiden Abschnitten beschreibt er, wie der Vater mit ihm gesprochen hat: durch Redeverbote (A) und durch Ironie (C)."],
  ["Hauptteil A", "Indem der Vater jede Widerrede verbietet, sendet er kaum eine Sachinformation, sondern einen Appell – „Schweig!“ – und vor allem eine Beziehungsbotschaft: Er allein bestimmt. Seine Drohung „kein Wort der Widerrede!“ (A, Z. 5) unterstreicht er nonverbal mit der „erhobene[n] Hand“ (Z. 5). Die Folge zeigt sich in einer Steigerung: Der Sohn bekommt „eine stockende, stotternde Art des Sprechens“ (Z. 7 f.), schweigt „schließlich“ (Z. 8) und kann am Ende „weder denken noch reden“ (Z. 9). Die Beziehung ist starr komplementär. Verschärft wird sie durch ein Missverständnis: Der Vater deutet das Schweigen als Trotz („contra“, Z. 15), der Sohn erlebt es als Folge von „Stärke“ und „Schwäche“ (Z. 16)."],
  ["Hauptteil C", "Hinzu kommt die Ironie. Die Fragen des Vaters – „Das ist Dir wohl schon zu viel?“ (C, Z. 3 f.) – sind nur scheinbar Fragen, tatsächlich sind sie Vorwürfe. Weil sie von „bösem Lachen und bösem Gesicht“ (Z. 5) begleitet werden, widersprechen sich digitale und analoge Botschaft. Der Sohn fühlt sich „schon bestraft, ehe man noch wußte“, was er falsch gemacht hat (Z. 6). Besonders kränkend ist, dass der Vater ihn über die Mutter anspricht („vom Herrn Sohn“, Z. 10) und ihn damit „nicht einmal des bösen Ansprechens gewürdigt“ (Z. 8) hat. Der Sohn übernimmt dieses Muster und fragt nur noch die Mutter: „Wie geht es dem Vater?“ (Z. 15)."],
  ["Schluss", "Insgesamt stützen beide Stellen die These, dass Kafkas Furcht auf Kommunikationsproblemen gründet: Wo Widerspruch verboten ist und Worte und Gesicht einander widersprechen, kann der Sohn nichts klären. Kafka ergänzt allerdings selbst, dass das Machtgefälle – „Deiner Stärke und meiner Schwäche“ (A, Z. 16) – dazugehört."],
];
const mW = [1900, W - 1900];
const mRows = muster.map(([a, b]) => row([
  cell(p(t(a, { bold: true, color: NAVY, size: 19 }), { after: 0 }), { w: mW[0], fill: LIGHT }),
  cell(p(t(b, { font: "Cambria", size: 19 }), { after: 0 }), { w: mW[1] }),
], 400));

const page3 = [
  kicker("Tafelscript · Erwartungshorizont", true),
  h1("Musterfließtext zu Aufgabe 2"),
  p(t("Eine mögliche Lösung – Schülertexte sind kürzer. Entscheidend: jede Behauptung mit Zitat und Zeile belegt und mit einem Fachbegriff erklärt.", { size: 18, italics: true, color: MUTED }), { after: 80 }),
  table(mW, mRows),
  p(t(""), { after: 60 }),
  box([
    p([t("Woran man einen guten Fließtext erkennt:  ", { bold: true, color: NAVY, size: 19 }), t("Einleitung mit Autor, Titel, Jahr, Thema · Behauptung → Beleg (Zitat + Z.) → Erklärung in jedem Absatz · Fachbegriffe erklären, nicht nur nennen · Überleitungen statt Aufzählung · Schluss mit Bezug zur These · korrekte Zitierweise ([…], [n]) · Rechtschreibung, Grammatik, Zeichensetzung", { size: 18 })], { after: 0 }),
  ]),
];

const doc = new Document({
  styles: { default: { document: { run: { font: "Calibri", size: 20 } } } },
  sections: [{
    properties: { page: { size: { width: 11906, height: 16838, orientation: PageOrientation.LANDSCAPE }, margin: { top: 500, bottom: 500, left: 750, right: 750, footer: 300 } } },
    footers: { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ font: "Calibri", size: 16, color: MUTED, children: ["Tafelscript · Kafka, Brief an den Vater · Seite ", PageNumber.CURRENT] })] })] }) },
    children: [...page1, ...page2, ...page3],
  }],
});
Packer.toBuffer(doc).then((b) => fs.writeFileSync("Tafelscript_Kafka_Brief_an_den_Vater.docx", b));
