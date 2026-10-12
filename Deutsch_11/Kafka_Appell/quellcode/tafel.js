// Tafelscript „Kafkas Appell“ (Tafelbild im Endstand, Ablauf mit Impulsen, Erwartungshorizont mit Musterrede)
// und Vorlesefassung des Mini-Diktats
const fs = require("fs");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, BorderStyle,
  ShadingType, AlignmentType, VerticalAlign, HeightRule, Footer, PageNumber, PageOrientation,
} = require("docx");
const { TEXT, QUELLE, DIKTAT, clean, STOLPER, LEITFRAGEN, REDE_AUFTRAG, KRITERIEN } = require("./daten");

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
const aW = [950, 1900, 7000, 3088, 2400];
const hdr = (x, w) => cell(p(t(x, { bold: true, color: "FFFFFF", size: 18 }), { after: 0 }), { w, fill: NAVY });
const ablauf = [
  ["0–4'", "Einstieg", "Kontext vom Arbeitsblatt vorlesen, dann den Auszug (Folie 1, Klasse liest mit). Impuls (Folie 2): „Worum bittet Kafka seinen Vater?“ Zwei, drei spontane Antworten – nicht bewerten.", "① Vermutungen", "AB, Folie 1–2"],
  ["4–9'", "Information: Mini-Diktat", "Folie 3 steht. Ganz vorlesen, Satz für Satz diktieren (Pausen an ‖), noch einmal vorlesen. Selbstkontrolle mit Folie 4, Fehler zählen, Stolperstellen kurz zeigen.", "–", "Vorlesefassung, Folie 3–4"],
  ["9–13'", "Vorarbeit (allein)", "Auftrag 1 (Folie 5): Kernbotschaft in einem Satz, „Kafka bittet seinen Vater, …“. Zwei Sätze vorlesen lassen, den treffendsten anschreiben. Impuls: „Stimmt das mit unseren Vermutungen überein?“", "② Kernbotschaft", "Folie 5, Hefter"],
  ["13–23'", "Analyse (6' allein, 4' Gespräch)", "Auftrag 2 (Folie 6): drei Leitfragen, jeweils offen oder verdeckt? Danach Frage für Frage im Unterrichtsgespräch, Ergebnisse ins Quadrat. Impulse: „Welches Wort zeigt, dass Kafka etwas will?“ · „Was soll der Vater mit ‚freundliches Wort, Hand, Blick‘ anfangen?“ · „Warum betont Kafka, dass der Vater keine Schuld hat?“ → Furcht.", "③ Quadrat, Appell zuletzt und groß", "Folie 6, Hefter"],
  ["23–38'", "Transfer (allein)", "Überleitung: „Max Brods Brief hat gewirkt – der Vater hört zu. Jetzt spricht Franz selbst, und zwar offen.“ Auftrag und Kriterien (Folie 7 bzw. Whiteboard-PDF). Kriterien an die Tafel. Timer 15 Min.; herumgehen, bei Bedarf auf die Ich-Botschaft hinweisen.", "④ Kriterien", "Folie 7 / Whiteboard, Hefter"],
  ["38–45'", "Vortrag und Reflexion", "Zwei Reden ohne Unterbrechung (Folie 8: die anderen haken ab). Auswertung Kriterium für Kriterium. Reflexion (Folie 9): „Warum hat Kafka im Brief nicht so gesprochen? Welche Fassung hätte der Vater eher gehört?“ Merksatz anschreiben. Wer nicht fertig ist, schreibt nächste Stunde weiter – keine Hausaufgabe.", "⑤ Merksatz", "Folie 8–9"],
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
  kicker("TAFELSCRIPT · ABLAUF DER STUNDE (45 MIN.)", true),
  h1("Ablauf und Impulse"),
  p(t("Phasenfolge nach dem Muster der Mentorin: Auszug → Information → Vorarbeit → Analyse mit Visualisierung → Transfer → Reflexion. Jede Phase hat ein eigenes Produkt (rechts: was danach an der Tafel steht).", { size: 18, italics: true, color: MUTED }), { after: 80 }),
  table(aW, aRows),
];

// ================= Seite 3: Erwartungshorizont =================
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
  ["Ein guter Blick hätte mir als Kind mehr geholfen als jede Drohung.", 4],
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
    p([t("Auftrag (Folie 7): ", { bold: true, color: NAVY, size: 18 }), t(REDE_AUFTRAG, { size: 18 })], { after: 30 }),
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

// ================= Mini-Diktat: Vorlesefassung (A4 hoch, eine Seite) =================
const DW = 10206;
const dict = DIKTAT.map((s, i) => {
  const runs = [t(String(i + 1) + "  ", { bold: true, color: MUTED, size: 24 })];
  s.split(" / ").forEach((part, j, arr) => {
    runs.push(t(part, { font: "Cambria", size: 30 }));
    if (j < arr.length - 1) runs.push(t("  ‖  ", { bold: true, color: RED, size: 30 }));
  });
  return p(runs, { after: 160, line: 360 });
});
const diktat = new Document({
  styles: { default: { document: { run: { font: "Calibri", size: 22 } } } },
  sections: [{
    properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 900, bottom: 800, left: 850, right: 850, footer: 400 } } },
    footers: { default: footer("Mini-Diktat · Der Appell") },
    children: [
      kicker("MINI-DIKTAT · VORLESEFASSUNG FÜR DIE LEHRKRAFT"),
      p(t("Der Appell – offen und verdeckt", { font: "Cambria", size: 38, bold: true, color: NAVY }), { after: 120 }),
      table([DW], [row([cell([
        p(t("Ablauf (ca. 5 Min.)", { bold: true, color: NAVY, size: 21 }), { after: 40 }),
        p(t("1. Ganz vorlesen – nur zuhören.  2. Satz für Satz diktieren, an ‖ kurz pausieren, Satzzeichen nicht ansagen.  3. Noch einmal vorlesen.  4. Selbstkontrolle mit Folie 4: Fehler farbig anstreichen und zählen.", { size: 21 }), { after: 0 }),
      ], { w: DW, fill: LIGHT2, borders: all(none) })])]),
      p(t(""), { after: 160 }),
      ...dict,
      p(t("Stolperstellen (vorher kurz an die Tafel: Appell · dass · Empfänger)", { font: "Cambria", size: 26, bold: true, color: NAVY }), { before: 240, after: 100 }),
      ...STOLPER.map(([h, e]) => p([t(h + ":  ", { bold: true, color: RED, size: 22 }), t(e, { size: 22 })], { after: 60 })),
      p(t(`Umfang: ${clean(DIKTAT).split(/\s+/).length} Wörter – bewusst kurz.`, { size: 19, italics: true, color: MUTED }), { before: 160 }),
    ],
  }],
});
Packer.toBuffer(diktat).then((b) => fs.writeFileSync("Diktat_Appell_Vorlesefassung.docx", b));
