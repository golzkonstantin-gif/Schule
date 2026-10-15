// Lösung zum Arbeitsblatt „Die Ebene des Appells“ (Aufgaben 1–5), für die Lehrkraft
const fs = require("fs");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, BorderStyle,
  ShadingType, AlignmentType, VerticalAlign, Header, Footer, TabStopType,
} = require("docx");
const { LEITFRAGEN, MUSTERREDE, KRITERIEN, REDE_SITUATION, REDE_AUFTRAG } = require("./daten");

const NAVY = "1E2761", MUTED = "5B6B8C", LIGHT = "F4F7FD", LIGHT2 = "EAF0FA", RED = "D9534F", GREEN = "2E9E6B";
const W = 10206;

const t = (text, o = {}) => new TextRun({ text, font: o.font || "Calibri", size: o.size || 21, bold: o.bold, italics: o.italics, color: o.color || "000000" });
const p = (runs, o = {}) => new Paragraph({ children: Array.isArray(runs) ? runs : [runs], alignment: o.align, pageBreakBefore: o.pb, keepNext: o.keepNext, spacing: { before: o.before ?? 0, after: o.after ?? 80, line: o.line } });
const none = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
const solid = (c, s = 6) => ({ style: BorderStyle.SINGLE, size: s, color: c });
const all = (b) => ({ top: b, bottom: b, left: b, right: b });
const cell = (children, o = {}) => new TableCell({
  children: Array.isArray(children) ? children : [children], width: { size: o.w, type: WidthType.DXA },
  shading: o.fill ? { fill: o.fill, type: ShadingType.CLEAR, color: "auto" } : undefined,
  borders: o.borders || all(solid("B9C6E8")), verticalAlign: VerticalAlign.TOP,
  margins: { top: 70, bottom: 70, left: 130, right: 130 },
});
const table = (widths, rows) => new Table({ width: { size: widths.reduce((a, b) => a + b, 0), type: WidthType.DXA }, columnWidths: widths, rows });
const row = (cells) => new TableRow({ children: cells, cantSplit: true });
const box = (paras, fill = LIGHT2) => table([W], [row([cell(paras, { w: W, fill, borders: all(none) })])]);
const aufgabe = (text, o = {}) => p(t(text, { font: "Cambria", size: 25, bold: true, color: NAVY }), { before: o.before ?? 200, after: 70, keepNext: true, pb: o.pb });
const zitat = (x) => t(x, { italics: true, color: NAVY });

// ===== Aufgabe 2: Lösung je Leitfrage =====
const A2 = [
  { art: "offener Appell", col: RED, inhalt: [
    [zitat("„Wobei ich Dich aber immerfort bitte, nicht zu vergessen, daß ich niemals […] an eine Schuld Deinerseits glaube“"), t(" (Z. 1 f.): Der Vater soll nicht vergessen, dass Kafka ihm keine Schuld gibt.")],
    [zitat("„nur sollst Du aufhören, es für eine besondere Bosheit meinerseits zu halten“"), t(" (Z. 3 f.): Der Vater soll aufhören, Kafkas Verhalten als böse Absicht zu deuten.")],
    [t("→ Beide Appelle sind direkt ausgesprochen: „bitte“ und „sollst“ zeigen die Aufforderung.", { bold: true })],
  ] },
  { art: "verdeckter Appell", col: RED, inhalt: [
    [zitat("„ein freundliches Wort, ein stilles Bei-der-Hand-Nehmen, ein guter Blick“"), t(" (Z. 5 f.): Kafka beschreibt nur, was ihn als Kind erreicht hätte – im Konjunktiv („hätten abfordern können“) und als Vermutung („ich kann nicht glauben“).")],
    [t("→ Gemeint ist: Sei freundlich zu mir, zeig mir Nähe, sieh mich wohlwollend an.", { bold: true })],
  ] },
  { art: "Absicherung / Beziehungsebene", col: NAVY, inhalt: [
    [t("Kafka entlastet den Vater ausdrücklich: "), zitat("„niemals im entferntesten an eine Schuld Deinerseits“"), t(" (Z. 1 f.) und "), zitat("„Du wirktest so auf mich, wie Du wirken mußtest“"), t(" (Z. 2 f.).")],
    [t("So soll die Bitte nicht als Vorwurf ankommen. Er rechnet mit Zorn und Gegenvorwürfen – ein Zeichen seiner Furcht und der komplementären Beziehung: Der Unterlegene bittet vorsichtig.")],
  ] },
];

// ===== Aufgabe 3: Kommunikationsquadrat =====
const qW = W / 2;
const quad = (label, punkte, focus) => cell([
  p(t("[ " + label + " ]", { bold: true, size: 21, color: focus ? RED : NAVY }), { after: 40 }),
  ...punkte.map((x) => p(Array.isArray(x) ? x : t("•  " + x, { size: 20 }), { after: 30 })),
], { w: qW, fill: focus ? "FDECEA" : "FFFFFF", borders: all(solid(focus ? RED : "B9C6E8", focus ? 12 : 6)) });
const quadrat = table([qW, qW], [
  row([
    quad("SACHINHALT", ["Der Vater hat so gewirkt, wie er musste (Z. 2 f.).", "Kafka ist dieser Wirkung erlegen (Z. 4 f.).", "Freundlichkeit hätte bei ihm alles erreicht (Z. 5–7)."]),
    quad("SELBSTKUNDGABE", ["Ich fühle mich missverstanden und zu Unrecht für böse gehalten.", "Ich habe Angst vor Deinen Vorwürfen.", "Ich war schwach und hätte Zuwendung gebraucht."]),
  ]),
  row([
    quad("BEZIEHUNG", ["Ich klage Dich nicht an.", "Du bist der Stärkere, ich bin Dir unterlegen (komplementär).", "Ich wünsche mir Frieden zwischen uns."]),
    quad("APPELL", [
      [t("offen: ", { bold: true, color: RED, size: 20 }), t("Vergiss nicht, dass ich Dir keine Schuld gebe. Hör auf, mich für böse zu halten.", { size: 20 })],
      [t("verdeckt: ", { bold: true, color: RED, size: 20 }), t("Sei freundlich zu mir, nimm mich an der Hand, sieh mich wohlwollend an.", { size: 20 })],
    ], true),
  ]),
]);

const header = new Header({ children: [new Paragraph({ tabStops: [{ type: TabStopType.RIGHT, position: W }], children: [t("Deutsch · Klasse 11 · Kafka, „Brief an den Vater“", { size: 17, color: MUTED }), t("\tLösung für die Lehrkraft", { size: 17, color: RED, bold: true })] })] });
const footer = new Footer({ children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [t("Lösung zum Arbeitsblatt „Die Ebene des Appells“", { size: 16, color: MUTED })] })] });

const inhalt = [
  p(t("Lösung: Die Ebene des Appells in Kafkas „Brief an den Vater“", { font: "Cambria", size: 32, bold: true, color: NAVY }), { after: 60 }),
  p(t("Mögliche Lösungen – Schülerantworten dürfen kürzer und anders formuliert sein. Entscheidend: Zitat mit Zeile, Fachbegriff (offen/verdeckt), Deutung.", { size: 18, italics: true, color: MUTED }), { after: 60 }),

  aufgabe("Aufgabe 1: Vorarbeit (Textverständnis)", { before: 120 }),
  box([p(t("Kafka bittet seinen Vater, ihm sein Verhalten nicht als böse Absicht vorzuwerfen – als ängstliches Kind hätte er nur Freundlichkeit und Nähe gebraucht.", { size: 22, italics: true }), { after: 0 })]),

  aufgabe("Aufgabe 2: Analyse des Appells"),
  ...LEITFRAGEN.flatMap(([frage], i) => [
    p([t((i + 1) + ". " + frage + "  ", { bold: true, size: 22 }), t("(" + A2[i].art + ")", { bold: true, size: 20, color: A2[i].col })], { before: i ? 100 : 0, after: 40, keepNext: true }),
    ...A2[i].inhalt.map((runs) => p(runs, { after: 40, line: 290 })),
  ]),
  box([
    p(t("Beispiel für einen ausformulierten Satz (mit Formulierungshilfen):", { bold: true, color: NAVY, size: 19 }), { after: 40 }),
    p(t("Ein offener Appell zeigt sich in der Aufforderung „nur sollst Du aufhören, es für eine besondere Bosheit meinerseits zu halten“ (Z. 3 f.). Verdeckt appelliert Kafka, wenn er beschreibt, dass ihn „ein freundliches Wort, ein stilles Bei-der-Hand-Nehmen, ein guter Blick“ (Z. 5 f.) erreicht hätten. Damit will er erreichen, dass der Vater freundlicher mit ihm umgeht. Seine Bitte sichert er ab, indem er betont, dass er dem Vater keine Schuld gibt (Z. 1 f.).", { size: 20, italics: true }), { after: 0, line: 290 }),
  ], "FFF8EC"),

  aufgabe("Aufgabe 3: Visualisierung", { pb: true }),
  quadrat,

  aufgabe("Aufgabe 4: Reflexion"),
  p([t("Warum so vorsichtig und versteckt? ", { bold: true }), t("Kafka hat Furcht vor dem Vater: Er kennt das Redeverbot, Zorn und Vorwürfe. Deshalb wagt er keinen direkten Appell, sondern entlastet den Vater und formuliert seine Wünsche als Rückblick im Konjunktiv. Schon dass er schreibt statt spricht, zeigt die gestörte, komplementäre Beziehung.")], { after: 60, line: 290 }),
  p([t("Hätte der Vater die Bitte verstanden? ", { bold: true }), t("Eher nicht: Der Vater spricht laut und direkt und hört nicht zu – eine versteckte Bitte hätte er wohl überhört oder als Vorwurf verstanden. Zudem hat er den Brief nie erhalten; der Appell erreicht ihn nicht. Auch eine „Ja, aber“-Antwort ist gut begründbar: Die offenen Bitten („sollst Du aufhören“) sind deutlich.")], { after: 0, line: 290 }),

  aufgabe("Aufgabe 5: Transfer – Musterrede"),
  p(t(REDE_SITUATION + " " + REDE_AUFTRAG, { size: 18, italics: true, color: MUTED }), { after: 60 }),
  p(MUSTERREDE.flatMap(([s, k]) => [t(s + " ", { font: "Cambria", size: 21 }), ...(k ? [t("(" + k + ")  ", { size: 16, bold: true, color: RED })] : [])]), { after: 80, line: 320 }),
  box([p([t("Kriterien: ", { bold: true, color: NAVY, size: 19 }), ...KRITERIEN.flatMap(([k, e], i) => [t((i ? "  ·  " : "") + (i + 1) + " " + k + ": ", { bold: true, size: 19, color: NAVY }), t(e, { size: 19 })])], { after: 0 })]),
];

const doc = new Document({
  styles: { default: { document: { run: { font: "Calibri", size: 21 } } } },
  sections: [{ properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 850, bottom: 750, left: 850, right: 850, header: 380, footer: 380 } } }, headers: { default: header }, footers: { default: footer }, children: inhalt }],
});
Packer.toBuffer(doc).then((b) => fs.writeFileSync("Loesung_Arbeitsblatt_Kafkas_Appell.docx", b));
