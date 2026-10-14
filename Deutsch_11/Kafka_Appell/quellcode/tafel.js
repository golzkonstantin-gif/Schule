// Tafelscript „Auswertung des Briefs + Kafkas Appell (AVIVA)“: Tafelbild im Endstand, Ablauf mit Impulsen,
// Erwartungshorizont mit Musterrede
const fs = require("fs");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, BorderStyle,
  ShadingType, AlignmentType, VerticalAlign, HeightRule, Footer, PageNumber, PageOrientation,
} = require("docx");
const { LEITFRAGEN, REDE_SITUATION, REDE_AUFTRAG, REDE_UMFANG, KRITERIEN } = require("./daten");

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
    bc([p(c("① Wiederholung", { bold: true, color: YEL, size: 22 })),
      p(c("verbal · paraverbal · nonverbal", { size: 20 })),
      p(c("vier Ohren: Sachinhalt · Selbstkundgabe · Beziehung · Appell", { size: 20 })),
      p(c("Inhalts- und Beziehungsebene – Beziehung Kafka/Vater: gestört → Brods Brief", { size: 20 }))], { w: H }),
    bc([p(c("② Kernbotschaft", { bold: true, color: YEL, size: 22 })), p(c("Kafka bittet seinen Vater, ihm sein Verhalten nicht als Bosheit vorzuwerfen – als Kind hätte er nur Freundlichkeit gebraucht.", { size: 20 }))], { w: H }),
  ], 900),
  row([bc(p([c("③ Analyse: ", { bold: true, color: YEL, size: 22 }), c("Kafkas Nachricht im Kommunikationsquadrat – Schwerpunkt Appell", { size: 20 })], { after: 0 }), { w: W, span: 2, fill: BOARD2, m: 60 })]),
  row([
    quad("Sachinhalt", "Worüber spricht er?", ["Du hast so auf mich gewirkt, wie Du musstest (Z. 2 f.).", "Freundlichkeit hätte bei mir alles erreicht (Z. 5–7)."], CHALK),
    quad("Selbstkundgabe", "Was zeigt er von sich?", ["Ich fühle mich zu Unrecht für böse gehalten.", "Ich habe Angst vor Deinen Vorwürfen."], "9FD3A8"),
  ], 900),
  row([
    quad("Beziehung", "Wie steht er zum Vater?", ["„niemals … an eine Schuld Deinerseits“ (Z. 1 f.):", "Ich klage Dich nicht an – Du bist der Mächtige."], "8EC3F0"),
    quad("APPELL", "Was soll der Vater tun?", [
      [c("offen: ", { bold: true, color: APP }), c("„bitte, nicht zu vergessen“ (Z. 1) · „sollst Du aufhören“ (Z. 3)", { italics: true }), c("  → Halte mich nicht für böse!", { bold: true })],
      [c("verdeckt: ", { bold: true, color: APP }), c("„ein freundliches Wort, ein stilles Bei-der-Hand-Nehmen, ein guter Blick“ (Z. 5 f.)", { italics: true }), c("  → Sei freundlich zu mir!", { bold: true })],
      [c("abgesichert: ", { bold: true, color: APP }), c("vorsichtig, höflich, ohne Vorwurf – aus Furcht", {})],
    ], APP, BOARD2),
  ], 1300),
  row([bc([
    p([c("④ Merksatz: ", { bold: true, color: YEL, size: 22 }), c("Kafka appelliert vorsichtig und verdeckt – aus Furcht vor dem Vater. Ein offener Appell mit Ich-Botschaft sagt klar, was er braucht.", { size: 21, bold: true })], { after: 0 }),
  ], { w: W, span: 2, fill: BOARD2 })]),
  row([bc([
    p([c("⑤ Transfer – Rede an den Vater: ", { bold: true, color: YEL, size: 22 }), c(KRITERIEN.map(([k], i) => (i + 1) + " " + k).join("  ·  "), { size: 20 })], { after: 0 }),
  ], { w: W, span: 2, borders: { top: line, bottom: none, left: none, right: none } })]),
]);

const page1 = [
  p([t("TAFELBILD IM ENDSTAND · ", { size: 17, bold: true, color: MUTED }), t("Teil 2 (Kafka): ① Vorwissen · ② nach Aufgabe 1 · ③ Auswertung der Aufgaben 2–3 · ④ nach der Reflexion · ⑤ vor dem Schreiben. Teil 1 (Auswertung) läuft über die Folien.", { size: 17, color: MUTED })], { after: 60 }),
  board,
];

// ================= Seite 2: Ablauf =================
const aW = [950, 2350, 6650, 2988, 2400];
const hdr = (x, w) => cell(p(t(x, { bold: true, color: "FFFFFF", size: 18 }), { after: 0 }), { w, fill: NAVY });
const ablauf = [
  ["0–2'", "Verlauf und Lernziele", "Folie 1–2: Teil 1 Auswertung des Briefs, Teil 2 Kafka (Appell). Lernziele vorlesen.", "–", "Folie 1–2"],
  ["2–6'", "Auswertung des Briefs", "Folie 3: Ø 14,5/20 Inhalt, 4,4/10 Sprache, 18,9/30 gesamt (Auswertung vom 08.10.; vorher mit Ihren endgültigen Punkten abgleichen). Lob zuerst (Beziehung, Metakommunikation), dann: Appell oft nur „Ich hoffe, ihr …“ → Teil 2. Folie 4: Rechtschreibung Ø 0,4/2, Zeichensetzung Ø 0,6/2 – Substantive groß, Komma vor dass (auch „Ich hoffe, ihr …“); Training folgt in den nächsten Stunden.", "–", "Folie 3–4"],
  ["6–8'", "A: Ankommen", "Folie 5 (Verlauf, Teil 2), dann Folie 6: „Ich habe euch eben Rückmeldung gegeben. Was ist mein Appell an euch?“ – Rückfrage: offen oder versteckt? („Ich habe nur gesagt, was schlecht war.“)", "–", "Folie 5–6"],
  ["8–13'", "V: Vorwissen aktivieren", "Folie 7: verbal / paraverbal / nonverbal · vier Ohren · Inhalts- und Beziehungsebene. Folie 8: Beziehung Kafka–Vater gestört, Brods Brief → heute: Appell-Ebene.", "① Wiederholung", "Folie 7–8"],
  ["13–16'", "I: Informieren", "Folie 9: Appell, offen, verdeckt (auch AB, Nr. 2). Kontext-Box vorlesen, dann den Auszug (Folie 10, AB).", "–", "AB S. 1, Folie 9–10"],
  ["16–25'", "V: Verarbeiten (allein)", "Aufgaben 1–3 (Folie 11), Timer 9 Min. Aufgaben 1–3 direkt auf dem AB (Schreiblinien). Hilfe zu Aufgabe 1: „Was soll der Vater nach diesem Satz anders machen?“", "–", "AB S. 1–2"],
  ["25–31'", "A: Auswerten", "Kernbotschaft anschreiben (②). Leitfragen der Reihe nach, Ergebnisse ins Quadrat, Appell zuletzt und groß (③). Formulierungshilfen (Folie 12): ganze Sätze. Reflexion (Folie 13) → Merksatz (④).", "② ③ ④", "Folie 12–13"],
  ["31–41'", "Transfer: Rede schreiben", "Folie 14 bzw. Whiteboard-PDF: Rede als Franz, Kriterien an die Tafel (⑤). Timer 10 Min. Hinweis: Substantive groß, Komma vor dass – auch bei „Ich hoffe, …“.", "⑤ Kriterien", "Folie 14 / Whiteboard"],
  ["41–45'", "Vortrag und Rückblick", "Ein bis zwei Reden ohne Unterbrechung, kurze Auswertung nach den Kriterien. Rückblick auf die Lernziele. Wer nicht fertig ist, schreibt nächste Stunde weiter – keine Hausaufgabe.", "–", "Folie 14"],
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
  kicker("TAFELSCRIPT · ABLAUF DER STUNDE (45 MIN.: 6 AUSWERTUNG + 39 KAFKA)", true),
  h1("Ablauf und Impulse"),
  p(t("Teil 1: Auswertung des Briefs (Rechtschreibtraining entfällt, folgt in den nächsten Stunden). Teil 2 nach AVIVA, Arbeitsblatt nach dem Muster der Mentorin. Alle Folien in 27 pt.", { size: 18, italics: true, color: MUTED }), { after: 80 }),
  table(aW, aRows),
];

// ================= Seite 3: Erwartungshorizont =================
const eW = [5200, W - 5200];
const erwartung = [
  "Offener Appell: Der Vater soll nicht vergessen, dass Kafka ihm keine Schuld gibt, und aufhören, Kafkas Schwäche für Bosheit zu halten.",
  "Verdeckter Appell: Kafka sagt nur, was ihn als Kind erreicht hätte. Gemeint ist: Sei freundlich, nimm mich an der Hand, sieh mich wohlwollend an.",
  "Kafka entlastet den Vater ausdrücklich, damit dieser die Bitte nicht als Angriff versteht. Der Appell ist vorsichtig abgesichert – ein Zeichen seiner Furcht und der komplementären Beziehung.",
];
const REFLEXION = ["Reflexion (Aufgabe 4): Warum so vorsichtig und versteckt?", "Aus Furcht vor Zorn und Vorwürfen wagt Kafka keinen direkten Appell. Der Vater, der laut und direkt spricht, hätte die versteckte Bitte wohl überhört – darum bleibt der Brief ein Versuch, der den Vater nie erreicht."];
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
  h1("Erwartungshorizont zu den Aufgaben 2, 4 und 5"),
  table(eW, [...LEITFRAGEN.map(([frage, tipp], i) => row([
    cell([p(t(frage, { bold: true, size: 19, color: NAVY }), { after: 20 }), p(t(tipp, { size: 17, italics: true, color: MUTED }), { after: 0 })], { w: eW[0], fill: LIGHT }),
    cell(p(t(erwartung[i], { size: 19 }), { after: 0 }), { w: eW[1] }),
  ])), row([
    cell(p(t(REFLEXION[0], { bold: true, size: 19, color: NAVY }), { after: 0 }), { w: eW[0], fill: LIGHT }),
    cell(p(t(REFLEXION[1], { size: 19 }), { after: 0 }), { w: eW[1] }),
  ])]),
  p(t("Musterrede (ca. 120 Wörter, zum Vorlesen nach den Schülerreden, falls nötig)", { font: "Cambria", size: 26, bold: true, color: NAVY }), { before: 200, after: 40 }),
  p(t("In Klammern das erfüllte Kriterium: 1 Anrede · 2 Ich-Botschaft · 3 offener Appell · 4 Bezug zum Text · 5 respektvoller Ton", { size: 17, italics: true, color: MUTED }), { after: 80 }),
  p(ROTE.flatMap(([s, k]) => [t(s + " ", { font: "Cambria", size: 21 }), ...(k ? [t("(" + k + ")  ", { size: 16, bold: true, color: RED })] : [])]), { after: 120, line: 320 }),
  table([W], [row([cell([
    p([t("Auftrag (AB Aufgabe 5, Folie 14): ", { bold: true, color: NAVY, size: 18 }), t(REDE_SITUATION + " " + REDE_AUFTRAG + " (" + REDE_UMFANG + ")", { size: 18 })], { after: 30 }),
    p([t("Hinweis: ", { bold: true, color: NAVY, size: 18 }), t("Der erfolgreiche Brief von Max Brod stammt aus dem Transfer der Mentorin (Unterrichtsfiktion). Historisch hat Hermann Kafka den „Brief an den Vater“ nie gelesen.", { size: 18 })], { after: 0 }),
  ], { w: W, fill: LIGHT2, borders: all(none) })])]),
];

const footer = (label) => new Footer({ children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ font: "Calibri", size: 16, color: MUTED, children: [label + " · Seite ", PageNumber.CURRENT] })] })] });
const land = { size: { width: 11906, height: 16838, orientation: PageOrientation.LANDSCAPE }, margin: { top: 600, bottom: 500, left: 750, right: 750, footer: 300 } };
const doc = new Document({
  styles: { default: { document: { run: { font: "Calibri", size: 20 } } } },
  sections: [{ properties: { page: land }, footers: { default: footer("Tafelscript · Kafkas Appell") }, children: [...page1, ...page2, ...page3] }],
});
Packer.toBuffer(doc).then((b) => fs.writeFileSync("Tafelscript_Kafkas_Appell.docx", b));
