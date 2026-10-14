// Tafelscript „Rechtschreib-Check + Kafkas Appell“: Tafelbild im Endstand, Ablauf mit Impulsen,
// Rechtschreib-Check (Vormachen, Lösung), Erwartungshorizont mit Musterrede
const fs = require("fs");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, BorderStyle,
  ShadingType, AlignmentType, VerticalAlign, HeightRule, Footer, PageNumber, PageOrientation,
} = require("docx");
const { CHAT, vergleich, zaehle, REGELN, LEITFRAGEN, REDE_AUFTRAG, REDE_UMFANG, KRITERIEN } = require("./daten");

const NAVY = "1E2761", MUTED = "5B6B8C", LIGHT = "F4F7FD", LIGHT2 = "EAF0FA", RED = "D9534F";
const BOARD = "2E4A3B", BOARD2 = "3A5A48", CHALK = "F3F1EA", DIM = "C9C3A8", YEL = "F5D76E", APP = "F29A8E";

const t = (text, o = {}) => new TextRun({ text, font: o.font || "Calibri", size: o.size || 20, bold: o.bold, italics: o.italics, color: o.color || "000000" });
const p = (runs, o = {}) => new Paragraph({ children: Array.isArray(runs) ? runs : [runs], alignment: o.align, pageBreakBefore: o.pb, spacing: { before: o.before ?? 0, after: o.after ?? 60, line: o.line } });
const none = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
const solid = (c, s = 6) => ({ style: BorderStyle.SINGLE, size: s, color: c });
const all = (b) => ({ top: b, bottom: b, left: b, right: b });
const cell = (children, o = {}) => new TableCell({
  children: Array.isArray(children) ? children : [children], width: { size: o.w, type: WidthType.DXA },
  shading: o.fill ? { fill: o.fill, type: ShadingType.CLEAR, color: "auto" } : undefined,
  borders: o.borders || all(solid("B9C6E8")), verticalAlign: o.valign || VerticalAlign.TOP,
  margins: { top: o.m ?? 60, bottom: o.m ?? 60, left: 120, right: 120 }, columnSpan: o.span,
});
const table = (widths, rows) => new Table({ width: { size: widths.reduce((a, b) => a + b, 0), type: WidthType.DXA }, columnWidths: widths, rows });
const row = (cells, h) => new TableRow({ children: cells, height: h ? { value: h, rule: HeightRule.ATLEAST } : undefined, cantSplit: true });
const kicker = (text, pb) => p(t(text, { size: 17, bold: true, color: MUTED }), { after: 20, pb });
const h1 = (text) => p(t(text, { font: "Cambria", size: 32, bold: true, color: NAVY }), { after: 80 });

// ================= Seite 1: Tafelbild (Kreide) =================
const W = 15338, H = W / 2;
const c = (text, o = {}) => t(text, Object.assign({ color: CHALK, size: 21 }, o));
const line = solid("6E8C7A");
const bc = (children, o = {}) => cell(children, Object.assign({ fill: BOARD, borders: { top: line, bottom: line, left: none, right: none }, m: 90 }, o));
const quad = (label, frage, inhalt, col, fill = BOARD) => bc([
  p([c(label, { bold: true, color: col, size: 24 }), c("   " + frage, { italics: true, color: DIM, size: 18 })], { after: 40 }),
  ...inhalt.map((x) => p(Array.isArray(x) ? x : c(x), { after: 30 })),
], { w: H, fill, borders: all(line) });

const board = table([H, H], [
  row([bc([
    p(c("Franz Kafka: „Brief an den Vater“ (1919) – Worum bittet Kafka?", { font: "Cambria", size: 32, bold: true }), { after: 30 }),
    p([c("Auszug S. 3 · ", { color: DIM, size: 19 }), c("Fokus: der Appell – die Seite, mit der der Sender etwas erreichen will", { size: 19, color: DIM, italics: true })], { after: 0 }),
  ], { w: W, span: 2, borders: { top: none, bottom: line, left: none, right: none } })]),
  row([
    bc([p(c("① Vermutungen", { bold: true, color: YEL, size: 22 })), p(c("(Antworten der Klasse auf „Worum bittet Kafka?“)", { italics: true, color: DIM, size: 19 }))], { w: H }),
    bc([p(c("② Kernbotschaft", { bold: true, color: YEL, size: 22 })), p(c("Kafka bittet seinen Vater, ihm sein Verhalten nicht als Bosheit vorzuwerfen – als Kind hätte er nur Freundlichkeit gebraucht.", { size: 20 }))], { w: H }),
  ], 700),
  row([bc(p([c("③ Analyse: ", { bold: true, color: YEL, size: 22 }), c("Kafkas Nachricht im Kommunikationsquadrat – Schwerpunkt Appell", { size: 20 })], { after: 0 }), { w: W, span: 2, fill: BOARD2, m: 60 })]),
  row([
    quad("Sachinhalt", "Was sagt er über die Sache?", ["Ich war ein ängstliches Kind, aber nicht schwer zu lenken (Z. 5–8)."], CHALK),
    quad("Selbstkundgabe", "Was zeigt er von sich?", ["Ich fühle mich zu Unrecht für böse gehalten.", "Ich habe Angst vor Deinen Vorwürfen."], "9FD3A8"),
  ], 900),
  row([
    quad("Beziehung", "Wie steht er zum Vater?", ["„niemals … an eine Schuld Deinerseits“ (Z. 2):", "Ich klage Dich nicht an – Du bist der Mächtige."], "8EC3F0"),
    quad("APPELL", "Was soll der Vater tun?", [
      [c("offen: ", { bold: true, color: APP }), c("„bitte, nicht zu vergessen“ (Z. 1) · „sollst Du aufhören“ (Z. 3)", { italics: true }), c("  → Halte mich nicht für böse!", { bold: true })],
      [c("verdeckt: ", { bold: true, color: APP }), c("„ein freundliches Wort, ein stilles Bei-der-Hand-Nehmen, ein guter Blick“ (Z. 8 f.)", { italics: true }), c("  → Sei freundlich zu mir!", { bold: true })],
      [c("abgesichert: ", { bold: true, color: APP }), c("vorsichtig, höflich, ohne Vorwurf – aus Furcht", {})],
    ], APP, BOARD2),
  ], 1300),
  row([bc([
    p([c("④ Transfer – Rede an den Vater: ", { bold: true, color: YEL, size: 22 }), c(KRITERIEN.map(([k], i) => (i + 1) + " " + k).join("  ·  "), { size: 20 })], { after: 0 }),
  ], { w: W, span: 2 })]),
  row([bc([
    p([c("⑤ Merksatz: ", { bold: true, color: YEL, size: 22 }), c("Kafka appelliert vorsichtig und verdeckt – aus Furcht vor dem Vater. Ein offener Appell mit Ich-Botschaft sagt klar, was er braucht.", { size: 21, bold: true })], { after: 0 }),
  ], { w: W, span: 2, fill: BOARD2, borders: { top: line, bottom: none, left: none, right: none } })]),
]);

const page1 = [
  p([t("TAFELBILD IM ENDSTAND · ", { size: 17, bold: true, color: MUTED }), t("① Einstieg · ② nach Auftrag 1 · ③ nach Auftrag 2 im Unterrichtsgespräch · ④ vor dem Schreiben · ⑤ nach der Reflexion", { size: 17, color: MUTED })], { after: 60 }),
  board,
];

// ================= Seite 2: Ablauf =================
const aW = [950, 2350, 6650, 2988, 2400];
const hdr = (x, w) => cell(p(t(x, { bold: true, color: "FFFFFF", size: 18 }), { after: 0 }), { w, fill: NAVY });
const ablauf = [
  ["0–2'", "Rechtschreibung: Instruktion", "Schnipsel austeilen. „In euren Texten passieren immer dieselben zwei Fehler – nicht, weil ihr die Regel nicht kennt, sondern weil ihr nicht darauf achtet.“ Regel 1 und 2 zeigen (Folie 2–3), je ein Beispiel laut prüfen.", "–", "Schnipsel, Folie 1–3"],
  ["2–4'", "Vormachen", "Nachricht 1 laut denkend korrigieren (Folie 4, dann 5): Artikelprobe, Verb am Ende → Komma. Siehe nächste Seite.", "–", "Folie 4–5"],
  ["4–8'", "Korrigieren (allein)", "Nachrichten 2–6 auf dem Schnipsel korrigieren, Timer 4 Min. sichtbar. Nicht helfen – die Regeln stehen auf dem Schnipsel. Wer fertig ist, zählt seine Fehler.", "–", "Folie 6, Schnipsel"],
  ["8–10'", "Lösung im Gespräch", "Klasse nennt die Fehler je Nachricht und begründet (Artikelprobe / Verb am Ende), dann Folie zeigen. Abfrage: „Wer hat mindestens 11 von 14?“ Überleitung: „In Nachricht 2 ging es schon um Kafka.“", "–", "Folie 7–8"],
  ["10–13'", "Kafka: Einstieg", "Kontext vom Arbeitsblatt vorlesen, dann den Auszug (Folie 9, AB mit Zeilennummern). Impuls (Folie 10): „Worum bittet Kafka seinen Vater?“", "① Vermutungen", "AB, Folie 9–10"],
  ["13–15'", "Information", "Folie 11: Appell, offen, verdeckt. Rückfrage zum Chat: „‚Ich wünsche mir, dass mein Vater das Handy weglegt‘ – offen oder verdeckt?“", "–", "Folie 11"],
  ["15–18'", "Vorarbeit (allein)", "Auftrag 1 (Folie 12): Kernbotschaft in einem Satz. Zwei vorlesen, den treffendsten anschreiben.", "② Kernbotschaft", "Folie 12, Hefter"],
  ["18–26'", "Analyse (5' allein, 3' Gespräch)", "Auftrag 2 (Folie 13), Formulierungshilfen (Folie 14) bei Bedarf. Dann Frage für Frage ins Quadrat. Impulse: „Welches Wort zeigt, dass Kafka etwas will?“ · „Was soll der Vater mit Wort, Hand, Blick anfangen?“ · „Warum betont Kafka, dass der Vater keine Schuld hat?“ → Furcht.", "③ Quadrat, Appell zuletzt und groß", "Folie 13–14, Hefter"],
  ["26–38'", "Transfer (allein)", "Überleitung: „Max Brods Brief hat gewirkt – der Vater hört zu. Jetzt spricht Franz selbst, und zwar offen.“ Folie 15 bzw. Whiteboard-PDF. Timer 12 Min. Hinweis: Großschreibung und Kommas wie im Chat!", "④ Kriterien", "Folie 15 / Whiteboard, Hefter"],
  ["38–45'", "Vortrag und Reflexion", "Eine bis zwei Reden ohne Unterbrechung (Folie 16, die anderen haken ab), Auswertung nach Kriterien. Reflexion (Folie 17). Merksatz anschreiben. Rest nächste Stunde – keine Hausaufgabe.", "⑤ Merksatz", "Folie 16–17"],
];
const aRows = [row([hdr("Zeit", aW[0]), hdr("Phase", aW[1]), hdr("Ihre Impulse / was passiert", aW[2]), hdr("Tafel", aW[3]), hdr("Material", aW[4])], 360)];
ablauf.forEach(([z, ph, imp, tafel, mat]) => aRows.push(row([
  cell(p(t(z, { bold: true, color: NAVY, size: 18 }), { after: 0 }), { w: aW[0], fill: LIGHT }),
  cell(p(t(ph, { bold: true, size: 18, color: NAVY }), { after: 0 }), { w: aW[1], fill: LIGHT }),
  cell(p(t(imp, { size: 18 }), { after: 0 }), { w: aW[2] }),
  cell(p(t(tafel, { size: 18 }), { after: 0 }), { w: aW[3] }),
  cell(p(t(mat, { size: 17, color: MUTED }), { after: 0 }), { w: aW[4] }),
])));
const page2 = [
  kicker("TAFELSCRIPT · ABLAUF DER STUNDE (45 MIN.: 10 RECHTSCHREIBUNG + 35 KAFKA)", true),
  h1("Ablauf und Impulse"),
  p(t("Teil 1 höchstens 10 Minuten – Tempo ist Ziel. Teil 2 nach dem Muster der Mentorin: Auszug → Information → Vorarbeit → Analyse mit Visualisierung → Transfer → Reflexion. Alle Folien in 27 pt.", { size: 18, italics: true, color: MUTED }), { after: 80 }),
  table(aW, aRows),
];

// ================= Seite 3: Rechtschreib-Check =================
const fehlerRuns = (m, size) => vergleich(m.f, m.r).map((x) => x.fehler ? t(x.text, { size, bold: true, color: RED }) : t(x.text, { size }));
const rest = zaehle(CHAT.slice(1));
const cW = [700, 7300, W - 8000];
const pageR = [
  kicker("TAFELSCRIPT · TEIL 1: RECHTSCHREIB-CHECK", true),
  h1("Vormachen und Lösung"),
  table([W], [row([cell([
    p(t("So denken Sie bei Nachricht 1 laut (Folie 4):", { bold: true, color: NAVY, size: 19 }), { after: 40 }),
    p(t("„Ich glaube dass … – ‚dass‘: Wo steht das Verb? ‚hatte‘, ganz am Ende. Also Nebensatz, also Komma vor dass. – ‚vater‘: Passt ‚der‘ davor? Der Vater – ja, also groß. – ‚geduld‘: die Geduld – groß.“ Dann Folie 5 mit der Korrektur zeigen.", { size: 19, italics: true }), { after: 0 }),
  ], { w: W, fill: LIGHT2, borders: all(none) })])]),
  p(t(""), { after: 60 }),
  table(cW, [
    row([cell(p(t("Nr.", { bold: true, color: "FFFFFF", size: 18 }), { after: 0 }), { w: cW[0], fill: NAVY }), cell(p(t("Korrigierte Nachricht (Fehler rot)", { bold: true, color: "FFFFFF", size: 18 }), { after: 0 }), { w: cW[1], fill: NAVY }), cell(p(t("Begründung", { bold: true, color: "FFFFFF", size: 18 }), { after: 0 }), { w: cW[2], fill: NAVY })]),
    ...CHAT.map((m, i) => {
      const fe = vergleich(m.f, m.r).filter((x) => x.fehler);
      const gross = fe.filter((x) => x.fehler === "gross").map((x) => x.text);
      const komma = fe.filter((x) => x.fehler === "komma").length;
      const grund = [gross.length ? "groß: " + gross.join(", ") + " (Artikelprobe" + (gross.includes("Zuhören") ? "; beim Zuhören: Signalwort" : "") + ")" : "", komma ? "Komma vor dem Nebensatz (Verb am Ende)" : ""].filter(Boolean).join(" · ");
      return row([
        cell(p(t(String(i + 1) + (i ? "" : " *"), { bold: true, size: 18, color: NAVY }), { after: 0 }), { w: cW[0], fill: LIGHT }),
        cell(p([t(m.von + ": ", { bold: true, size: 18, color: MUTED }), ...fehlerRuns(m, 19)], { after: 0 }), { w: cW[1] }),
        cell(p(t(grund, { size: 17 }), { after: 0 }), { w: cW[2] }),
      ]);
    }),
  ]),
  p(t(`* Nachricht 1 machen Sie vor. Die Klasse korrigiert 2–6: ${rest.gross + rest.komma} Fehler (${rest.gross} × groß, ${rest.komma} × Komma).`, { size: 17, italics: true, color: MUTED }), { before: 60 }),
];

// ================= Seite 4: Erwartungshorizont =================
const eW = [5200, W - 5200];
const erwartung = [
  "Offener Appell: Der Vater soll nicht vergessen, dass Kafka ihm keine Schuld gibt, und aufhören, Kafkas Schwäche für Bosheit zu halten.",
  "Verdeckter Appell: Kafka sagt nur, was ihn als Kind erreicht hätte. Gemeint ist: Sei freundlich, nimm mich an der Hand, sieh mich wohlwollend an.",
  "Kafka entlastet den Vater ausdrücklich, damit dieser die Bitte nicht als Angriff versteht. Der Appell ist vorsichtig abgesichert – ein Zeichen seiner Furcht und der komplementären Beziehung.",
];
const ROTE = [
  ["Vater, danke, dass Du mir heute zuhörst.", 1],
  ["Ich will Dir keine Vorwürfe machen – Du hast mich so erzogen, wie Du es für richtig gehalten hast.", 5],
  ["Aber wenn Du laut wirst oder über das lachst, was mir wichtig ist, fühle ich mich klein und bringe kein Wort mehr heraus.", 2],
  ["Wenn Du mein Schweigen für Trotz hältst, verletzt mich das, denn ich schweige nicht aus Bosheit, sondern aus Angst.", 2],
  ["Ich war ein ängstliches Kind, und manchmal bin ich es heute noch.", 0],
  ["Darum bitte ich Dich: Hör auf, mein Verhalten als Bosheit zu deuten.", 3],
  ["Sag mir ab und zu ein freundliches Wort und sieh mich an, wenn ich Dir etwas erzähle.", 4],
  ["Ich wünsche mir, dass wir in Zukunft miteinander reden können, ohne dass einer von uns Angst haben muss.", 3],
];
const page3 = [
  kicker("TAFELSCRIPT · ERWARTUNGSHORIZONT", true),
  h1("Erwartungshorizont zu Auftrag 2 und zur Rede"),
  table(eW, LEITFRAGEN.map(([spur, frage], i) => row([
    cell([p(t(frage, { bold: true, size: 19, color: NAVY }), { after: 20 }), p(t(spur, { size: 17, italics: true, color: MUTED }), { after: 0 })], { w: eW[0], fill: LIGHT }),
    cell(p(t(erwartung[i], { size: 19 }), { after: 0 }), { w: eW[1] }),
  ]))),
  p(t("Musterrede (zum Vorlesen nach den Schülerreden, falls nötig)", { font: "Cambria", size: 26, bold: true, color: NAVY }), { before: 200, after: 40 }),
  p(t("In Klammern das erfüllte Kriterium: 1 Anrede · 2 Ich-Botschaft · 3 offener Appell · 4 Bezug zum Text · 5 respektvoller Ton", { size: 17, italics: true, color: MUTED }), { after: 80 }),
  p(ROTE.flatMap(([s, k]) => [t(s + " ", { font: "Cambria", size: 21 }), ...(k ? [t("(" + k + ")  ", { size: 16, bold: true, color: RED })] : [])]), { after: 120, line: 320 }),
  table([W], [row([cell([
    p([t("Auftrag (Folie 15): ", { bold: true, color: NAVY, size: 18 }), t(REDE_AUFTRAG + " (" + REDE_UMFANG + ")", { size: 18 })], { after: 30 }),
    p([t("Hinweis: ", { bold: true, color: NAVY, size: 18 }), t("Der erfolgreiche Brief von Max Brod stammt aus dem Transfer der Mentorin (Unterrichtsfiktion). Historisch hat Hermann Kafka den „Brief an den Vater“ nie gelesen.", { size: 18 })], { after: 0 }),
  ], { w: W, fill: LIGHT2, borders: all(none) })])]),
];

const footer = (label) => new Footer({ children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ font: "Calibri", size: 16, color: MUTED, children: [label + " · Seite ", PageNumber.CURRENT] })] })] });
const land = { size: { width: 11906, height: 16838, orientation: PageOrientation.LANDSCAPE }, margin: { top: 600, bottom: 500, left: 750, right: 750, footer: 300 } };
const doc = new Document({
  styles: { default: { document: { run: { font: "Calibri", size: 20 } } } },
  sections: [{ properties: { page: land }, footers: { default: footer("Tafelscript · Kafkas Appell") }, children: [...page1, ...page2, ...pageR, ...page3] }],
});
Packer.toBuffer(doc).then((b) => fs.writeFileSync("Tafelscript_Kafkas_Appell.docx", b));
