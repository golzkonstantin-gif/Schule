// Präsentation „Vor dem Gesetz“ – kafkaesk, Inhalt und Symbole (16:9, Schrift durchgehend 27 pt), nach AVIVA
const pptxgen = require("pptxgenjs");
const D = require("./daten");

const NAVY = "1E2761", NAVY2 = "24306E", MUTED = "5B6B8C", LIGHT = "F4F7FD", LIGHT2 = "EAF0FA", ICE = "CADCFC", WHITE = "FFFFFF";
const RED = "D9534F", GOLD = "E8A33D", GREEN = "2E9E6B", BLUE = "3F7CC4";
const HEAD = "Cambria", BODY = "Calibri";
const FS = 27;

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.333 x 7.5
pres.title = "Kafka – Vor dem Gesetz";

const txt = (s, text, o) => s.addText(text, Object.assign({ isTextBox: true, fontFace: BODY, fontSize: FS, color: NAVY, margin: 0, valign: "top" }, o));
const rect = (s, x, y, w, h, fill, o = {}) => s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, fill: { color: fill }, line: o.line ? { color: o.line, width: o.lw || 1.5 } : { type: "none" }, rectRadius: o.r ?? 0.08 });
function base(title, opts = {}) {
  const s = pres.addSlide();
  s.background = { color: WHITE };
  rect(s, 0, 0, 13.333, 1.0, NAVY, { r: 0 });
  txt(s, title, { x: 0.5, y: 0, w: 9.6, h: 1.0, fontFace: HEAD, fontSize: 30, bold: true, color: WHITE, valign: "middle" });
  if (opts.time) { rect(s, 10.9, 0.15, 2.1, 0.7, RED, { r: 0.12 }); txt(s, opts.time, { x: 10.9, y: 0.15, w: 2.1, h: 0.7, bold: true, color: WHITE, align: "center", valign: "middle" }); }
  else if (opts.phase) txt(s, opts.phase, { x: 9.3, y: 0.15, w: 3.7, h: 0.7, bold: true, color: ICE, align: "right", valign: "middle" });
  return s;
}
const bullets = (arr, o = {}) => arr.map((x, i) => ({ text: x, options: Object.assign({ bullet: o.num ? { type: "number" } : { indent: 24 }, breakLine: i < arr.length - 1, paraSpaceAfter: o.space ?? 12 }, o.opt || {}) }));
function aufgabe(a, notes, extra) {
  const s = base(a.titel, { time: a.zeit });
  txt(s, a.form, { x: 0.5, y: 1.15, w: 12.3, h: 0.5, italic: true, color: MUTED });
  txt(s, bullets(a.schritte, { num: true, space: 10 }), { x: 0.6, y: 1.75, w: 12.1, h: extra ? 3.4 : 4.6 });
  if (extra) extra(s);
  txt(s, "Fertig? Stift hinlegen.", { x: 0.5, y: 6.8, w: 12.3, h: 0.5, bold: true, color: GREEN, align: "right" });
  if (notes) s.addNotes(notes);
  return s;
}

// ===== A: Ankommen – Zitat als Impuls =====
{
  const s = pres.addSlide();
  s.background = { color: NAVY };
  s.addShape(pres.shapes.OVAL, { x: 10.6, y: -1.6, w: 4.6, h: 4.6, fill: { color: NAVY2 }, line: { type: "none" } });
  txt(s, "„", { x: 0.7, y: 0.3, w: 1.5, h: 1.6, fontFace: HEAD, fontSize: 120, color: GOLD });
  txt(s, [
    { text: "Auf Balzacs Spazierstockgriff:", options: { breakLine: true, color: ICE } },
    { text: "Ich breche alle Hindernisse.", options: { breakLine: true, bold: true } },
    { text: " ", options: { breakLine: true, fontSize: 14 } },
    { text: "Auf meinem:", options: { breakLine: true, color: ICE } },
    { text: "Mich brechen alle Hindernisse.", options: { breakLine: true, bold: true, color: GOLD } },
    { text: " ", options: { breakLine: true, fontSize: 14 } },
    { text: "Gemeinsam ist das ‚alle‘.", options: { italic: true, color: ICE } },
  ], { x: 1.6, y: 1.5, w: 10.5, h: 4.4, fontFace: HEAD, fontSize: 34, color: WHITE });
  txt(s, "— " + D.ZITAT.quelle, { x: 1.6, y: 6.2, w: 10.5, h: 0.6, color: ICE });
  s.addNotes("Ansage: „Stifte liegen. Lest den Satz still.“ 30 Sekunden. Impuls: „" + D.ZITAT.impuls + "“ Nachfragen: „Was ist der Unterschied zwischen ‚Ich breche‘ und ‚Mich brechen‘?“ Erwartung: " + D.ZITAT.erwartung + " Tafel (Seitentafel oder oben): Ohnmacht · Resignation. Gelenk: „Für dieses Gefühl gibt es ein eigenes Wort. Ihr kennt es schon.“ Zitat aus Kafkas Nachlass – Wortlaut vor der Stunde mit einer Ausgabe abgleichen.");
}

// ===== Verlauf und Lernziele =====
{
  const s = base("Heute");
  D.VERLAUF.forEach(([inhalt, phase], i) => {
    const y = 1.25 + i * 1.12;
    rect(s, 0.5, y, 12.3, 0.95, i % 2 ? WHITE : LIGHT, { line: "D5DCEC", lw: 0.75 });
    txt(s, phase, { x: 0.75, y, w: 3.2, h: 0.95, bold: true, color: RED, valign: "middle" });
    txt(s, inhalt, { x: 4.0, y, w: 8.6, h: 0.95, valign: "middle" });
  });
  s.addNotes("Ziel und Ablauf stehen auch an der Seitentafel. Was eventuell nicht mehr geschafft wird, hat einen Stern (*): Transfer und die Symbole mit Stern.");
}
{
  const s = base("Lernziele");
  txt(s, D.LERNZIELE.map(([z], i) => ({ text: z, options: { bullet: { indent: 24 }, breakLine: i < D.LERNZIELE.length - 1, paraSpaceAfter: 18 } })), { x: 0.8, y: 1.4, w: 11.8, h: 5.4 });
  s.addNotes("Lernziele vorlesen. Am Ende (Zielcheck) darauf zurückkommen.");
}

// ===== V: Vorwissen – kafkaesk =====
{
  const s = base("Was wisst ihr noch?", { phase: "VORWISSEN" });
  D.VORWISSEN.forEach(([q, h], i) => {
    const y = 1.35 + i * 2.6;
    rect(s, 0.5, y, 12.3, 2.3, i % 2 ? WHITE : LIGHT, { line: "D5DCEC", lw: 0.75 });
    txt(s, [{ text: q, options: { bold: true, breakLine: true } }, { text: h, options: { color: MUTED, italic: true } }], { x: 0.8, y, w: 11.8, h: 2.3, valign: "middle" });
  });
  s.addNotes("Ansage: „Ihr habt eine Minute. Schreibt Stichworte ins Heft.“ Dann Stichproben (4–6 Schüler, unterschiedliche Niveaus). Antworten an die Tafel (①). Erwartung: kafkaesk = bedrohlich, ausweglos, undurchschaubar, man ist einer Macht ausgeliefert. Kafka ohnmächtig gegenüber: dem Vater (Brief an den Vater), dem Schutzmann (Gibs auf!) – und den Behörden.");
}
{
  const s = base("„kafkaesk“", { phase: "VORWISSEN" });
  rect(s, 0.5, 1.2, 12.3, 1.2, LIGHT2);
  txt(s, D.KAFKAESK.duden, { x: 0.8, y: 1.2, w: 11.8, h: 1.2, valign: "middle", bold: true });
  txt(s, D.KAFKAESK.merkmale.map((m, i) => ({ text: m, options: { bullet: { indent: 24 }, breakLine: i < 3, paraSpaceAfter: 10 } })), { x: 0.7, y: 2.65, w: 12.0, h: 4.6 });
  s.addNotes("Erst nach den Schülerantworten aufdecken und mit der Tafel abgleichen: Was hatten wir schon, was fehlt? Steht auch auf dem Arbeitsblatt (Nr. 2).");
}
{
  const s = base("Kafka und die Macht", { phase: "VORWISSEN" });
  D.KAFKA_MACHT.forEach(([k, v], i) => {
    const y = 1.15 + i * 1.32;
    rect(s, 0.5, y, 12.3, 1.22, i % 2 ? WHITE : LIGHT, { line: "D5DCEC", lw: 0.75 });
    txt(s, k, { x: 0.75, y, w: 2.6, h: 1.22, bold: true, color: RED, valign: "middle" });
    txt(s, v, { x: 3.4, y, w: 9.2, h: 1.22, valign: "middle" });
  });
  rect(s, 0.5, 5.2, 12.3, 1.9, WHITE, { line: RED, lw: 2.5, r: 0.04 });
  txt(s, [{ text: "„Der Process“ beginnt so:", options: { bold: true, color: RED, breakLine: true } }, { text: "„" + D.PROCESS_SATZ + "“", options: { italic: true, fontFace: HEAD } }], { x: 0.8, y: 5.25, w: 11.8, h: 1.8, valign: "middle" });
  s.addNotes("Kurz erzählen, nicht abfragen. Vater und Schutzmann kennt die Klasse; neu ist der Bezug zum Staat: Kafka war Jurist in einer Behörde. Impuls zum Process-Satz: „Was ist daran kafkaesk?“ (Er hat nichts getan, erfährt keinen Grund – die Staatsmacht ist übermächtig und undurchschaubar.) Gelenk: „In diesem Roman steht die Parabel, die wir heute lesen.“");
}

// ===== I: Informieren – Kontext, Lesen, Kernbotschaft =====
{
  const s = base("„Vor dem Gesetz“ – Kontext", { phase: "INFORMIEREN" });
  txt(s, bullets(D.KONTEXT, { space: 14 }), { x: 0.7, y: 1.3, w: 12.0, h: 3.6 });
  rect(s, 0.5, 5.2, 12.3, 1.6, LIGHT2);
  txt(s, [{ text: "Leseauftrag: ", options: { bold: true, color: RED } }, { text: D.LESEAUFTRAG }], { x: 0.8, y: 5.2, w: 11.8, h: 1.6, valign: "middle" });
  s.addNotes("Arbeitsblatt austeilen. Ansage: „Stifte liegen. Ihr lest auf dem Arbeitsblatt mit.“ Die Lehrkraft liest den ganzen Text vor (Modell-Lesen, ca. 4 Min.: ruhig, Pausen an den Absätzen, die Rede des Türhüters deutlich). " + D.TEXT_HINWEIS);
}
{
  const s = base("Worum geht es – in einem Satz?", { phase: "INFORMIEREN" });
  txt(s, "Was will der Mann – und was hindert ihn?", { x: 0.5, y: 1.3, w: 12.3, h: 0.8, italic: true, color: MUTED });
  rect(s, 0.5, 3.0, 12.3, 2.6, LIGHT);
  txt(s, D.KERNBOTSCHAFT, { x: 0.9, y: 3.0, w: 11.5, h: 2.6, valign: "middle", fontFace: HEAD });
  s.addNotes("Erst die Frage stellen, 2–3 Schüler formulieren lassen (Kernbotschaft vor der Analyse, Muster der Mentorin). Dann den Satz aufdecken (Animation oder erst später einblenden).");
}

// ===== V: Verarbeiten – Aufgabe 1 =====
aufgabe(D.AUFGABE1, "Ansage: „Jetzt arbeitet ihr allein. Die Aufgabe steht auf dem Arbeitsblatt, Nr. 3.“ 3 Minuten. Herumgehen. Danach Lösung (nächste Folie), alle verbessern mit Grün.");
[0, 3].forEach((k) => {
  const s = base("Lösung: Handlungsschritte (" + (k / 3 + 1) + ")", { phase: "VERARBEITEN" });
  D.SCHRITTE.slice(k, k + 3).forEach(([text, anker], j) => {
    const i = k + j, y = 1.2 + j * 1.95;
    const buchstabe = String.fromCharCode(65 + D.REIHE_GEMISCHT.indexOf(i));
    rect(s, 0.5, y, 12.3, 1.8, j % 2 ? WHITE : LIGHT, { line: "D5DCEC", lw: 0.75 });
    txt(s, (i + 1) + " = " + buchstabe, { x: 0.75, y, w: 1.7, h: 1.8, bold: true, color: GREEN, valign: "middle" });
    txt(s, [{ text: text }, { text: "  (" + D.zeile(anker) + ")", options: { color: MUTED } }], { x: 2.5, y, w: 10.1, h: 1.8, valign: "middle" });
  });
  s.addNotes("Reihenfolge: " + D.SCHRITTE.map((_, i) => String.fromCharCode(65 + D.REIHE_GEMISCHT.indexOf(i))).join(" – ") + ". Zeilenangaben gelten als richtig, wenn sie ± 1 Zeile stimmen. Alle verbessern mit Grün.");
});

// ===== Ich mache vor: das offene Tor =====
{
  const s = base("Ich mache vor: das offene Tor", { phase: "VERARBEITEN" });
  D.VORMACHEN.forEach(([k, v], i) => {
    const y = 1.2 + i * 1.95;
    rect(s, 0.5, y, 12.3, 1.8, i === 2 ? LIGHT2 : LIGHT);
    txt(s, [{ text: k + "  ", options: { bold: true, color: i === 2 ? RED : NAVY } }, { text: v, options: i === 0 ? { italic: true } : {} }], { x: 0.8, y, w: 11.8, h: 1.8, valign: "middle" });
  });
  s.addNotes("Laut denken, Schritt für Schritt (Ich-Phase): „Was steht da? → Was fällt auf? → Wofür steht es?“ Die Klasse trägt die erste Zeile der Tabelle mit ein (" + D.zeile(D.SYMBOLE[0][2]) + "). Genau diese drei Schritte machen alle in Aufgabe 2.");
}
aufgabe(D.AUFGABE2, "Ansage: „Jetzt arbeitet ihr allein. Die Tabelle steht auf dem Arbeitsblatt, Aufgabe 2.“ Timer 9 Minuten. Herumgehen, bei Bedarf die Frage stellen: „Was fällt an dem Wort auf?“ Wer schnell fertig ist: die Zeilen mit Stern.");

// Lösungen in drei Portionen
for (let k = 0; k < D.SYMBOLE.length; k += 3) {
  const stern = D.SYMBOLE[k][4];
  const s = base("Lösung: Symbole" + (stern ? " *" : " (" + (k / 3 + 1) + ")"), { phase: "VERARBEITEN" });
  D.SYMBOLE.slice(k, k + 3).forEach(([b, , anker, sach], i) => {
    const y = 1.15 + i * 2.05;
    rect(s, 0.5, y, 12.3, 1.9, i % 2 ? WHITE : LIGHT, { line: "D5DCEC", lw: 0.75 });
    txt(s, [{ text: b, options: { bold: true, color: RED, breakLine: true } }, { text: D.zeile(anker), options: { color: MUTED } }], { x: 0.7, y, w: 3.4, h: 1.9, valign: "middle" });
    txt(s, sach, { x: 4.1, y, w: 8.6, h: 1.9, valign: "middle" });
  });
  s.addNotes(stern ? "Nur, wenn Zeit ist; sonst kurz mündlich das Schließen des Eingangs – es führt zur Deutung." : (k ? "Erst eine Antwort vom Platz, dann aufdecken. „Vergleicht und verbessert mit Grün.“" : "Erst eine Antwort vom Platz, dann aufdecken. Beim Gesetz ausdrücklich offen lassen: Recht, Wahrheit, Sinn, Gott – oder der Vater. Kafka legt es nicht fest. Wichtig ist die Begründung am Text."));
}

// ===== A: Auswerten – Deutung, Transfer*, Zielcheck =====
{
  const s = base(D.DEUTUNG.frage, { phase: "AUSWERTEN" });
  rect(s, 0.5, 1.2, 6.05, 3.3, LIGHT);
  txt(s, D.DEUTUNG.aussen, { x: 0.75, y: 1.2, w: 5.6, h: 3.3, valign: "middle" });
  rect(s, 6.75, 1.2, 6.05, 3.3, LIGHT2);
  txt(s, D.DEUTUNG.innen, { x: 7.0, y: 1.2, w: 5.6, h: 3.3, valign: "middle" });
  rect(s, 0.5, 4.8, 12.3, 2.2, WHITE, { line: GOLD, lw: 2.5, r: 0.04 });
  txt(s, D.DEUTUNG.zitat, { x: 0.8, y: 4.8, w: 11.8, h: 2.2, valign: "middle", fontFace: HEAD });
  s.addNotes("Erst die Frage stellen, Antworten sammeln, dann die beiden Felder aufdecken. Kern: Der Türhüter sagt „jetzt aber nicht“ und lädt sogar ein, es „trotz meines Verbotes“ zu versuchen. Der Mann fügt sich – die Ohnmacht wird zur Resignation. Rückgriff auf das Zitat vom Anfang. Tafel ③.");
}
{
  const s = base(D.TRANSFER.titel, { phase: "AUSWERTEN" });
  txt(s, bullets(D.TRANSFER.fragen, { space: 24 }), { x: 0.7, y: 1.5, w: 12.0, h: 4.5 });
  s.addNotes("Nur, wenn mehr als 6 Minuten übrig sind; mündlich, 2–3 Beiträge. Sonst entfällt der Transfer – der Zielcheck bleibt. Sportbezug nur anbieten, nicht erzwingen.");
}
{
  const s = base("Zielcheck", { time: "4 Min." });
  txt(s, "Allein, im Heft.", { x: 0.5, y: 1.15, w: 12.3, h: 0.6, italic: true, color: MUTED });
  rect(s, 0.5, 1.9, 12.3, 1.9, LIGHT);
  txt(s, D.ZIELCHECK.auftrag, { x: 0.8, y: 1.9, w: 11.8, h: 1.9, valign: "middle", bold: true });
  txt(s, [{ text: "Satzgerüst: ", options: { bold: true, color: BLUE } }, { text: D.ZIELCHECK.geruest, options: { italic: true } }], { x: 0.8, y: 4.1, w: 11.8, h: 1.3 });
  txt(s, "Fertig? Stift hinlegen.", { x: 0.5, y: 6.8, w: 12.3, h: 0.5, bold: true, color: GREEN, align: "right" });
  s.addNotes("Auswerten wird nie gestrichen. Herumgehen. Achtung Komma vor „weil“. Danach Musterlösung und Kriterien, mit Grün kontrollieren. Wer noch kein Merkmal von „kafkaesk“ nennt, wird in der Wiederholung der nächsten Stunde gezielt aufgegriffen.");
}
{
  const s = base("Zielcheck: Musterlösung");
  rect(s, 0.5, 1.2, 12.3, 3.3, LIGHT);
  txt(s, D.ZIELCHECK.muster, { x: 0.8, y: 1.2, w: 11.8, h: 3.3, valign: "middle" });
  txt(s, [{ text: "Prüft mit Grün: ", options: { bold: true, color: GREEN, breakLine: true } }, ...D.ZIELCHECK.kriterien.map((k, i) => ({ text: "✓ " + k, options: { breakLine: i < 2 } }))], { x: 0.8, y: 4.75, w: 11.8, h: 2.5 });
  s.addNotes("Jedes Symbol ist richtig, wenn Zeile, Deutung und ein Merkmal von „kafkaesk“ stimmen. Rückblick auf die Lernziele: Was davon kann ich jetzt?");
}

pres.writeFile({ fileName: "Praesentation_Vor_dem_Gesetz.pptx" });
