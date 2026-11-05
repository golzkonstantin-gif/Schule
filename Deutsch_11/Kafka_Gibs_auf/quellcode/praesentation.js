// Präsentation „Gibs auf!“ – Stilmittel und eigene Gesprächssituationen (16:9, Schrift durchgehend 27 pt), nach AVIVA
const pptxgen = require("pptxgenjs");
const D = require("./daten");

const NAVY = "1E2761", NAVY2 = "24306E", MUTED = "5B6B8C", LIGHT = "F4F7FD", LIGHT2 = "EAF0FA", ICE = "CADCFC", WHITE = "FFFFFF";
const RED = "D9534F", GOLD = "E8A33D", GREEN = "2E9E6B", BLUE = "3F7CC4", WOOD = "8B5E3C";
const HEAD = "Cambria", BODY = "Calibri";
const FS = 27;

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.333 x 7.5
pres.title = "Kafka – Gibs auf!";

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
function dark(fn) {
  const s = pres.addSlide();
  s.background = { color: NAVY };
  s.addShape(pres.shapes.OVAL, { x: 10.6, y: -1.6, w: 4.6, h: 4.6, fill: { color: NAVY2 }, line: { type: "none" } });
  fn(s);
  return s;
}
const bullets = (arr, o = {}) => arr.map((x, i) => ({ text: x, options: Object.assign({ bullet: o.num ? { type: "number" } : { indent: 24 }, breakLine: i < arr.length - 1, paraSpaceAfter: o.space ?? 12 }, o.opt || {}) }));

// Aufgabenfolie im festen Format: Nummer + Titel, Schritte mit Verb am Anfang, Zeit, Arbeitsform, „Fertig? Stift hinlegen.“
function aufgabe(a, phase, notes, extra) {
  const s = base(a.titel, { time: a.zeit });
  txt(s, a.form, { x: 0.5, y: 1.15, w: 12.3, h: 0.5, italic: true, color: MUTED });
  txt(s, bullets(a.schritte, { num: true, space: 10 }), { x: 0.6, y: 1.75, w: 12.1, h: extra ? 3.4 : 4.6 });
  if (extra) extra(s);
  txt(s, "Fertig? Stift hinlegen.", { x: 0.5, y: 6.8, w: 12.3, h: 0.5, bold: true, color: GREEN, align: "right" });
  if (notes) s.addNotes(notes);
  return s;
}

// ===== Einstieg: Verlauf und Lernziele =====
{
  const s = base("Heute");
  D.VERLAUF.forEach(([inhalt, phase], i) => {
    const y = 1.25 + i * 1.12;
    rect(s, 0.5, y, 12.3, 0.95, i % 2 ? WHITE : LIGHT, { line: "D5DCEC", lw: 0.75 });
    txt(s, phase, { x: 0.75, y, w: 3.2, h: 0.95, bold: true, color: RED, valign: "middle" });
    txt(s, inhalt, { x: 4.0, y, w: 8.6, h: 0.95, valign: "middle" });
  });
  s.addNotes("Ziel und Ablauf stehen auch an der Seitentafel. Was eventuell nicht mehr geschafft wird, hat einen Stern (*). Ansage: „Heute bleiben wir bei ‚Gibs auf!‘ – und dann schreibt ihr euer eigenes ‚Gibs auf!‘.“");
}
{
  const s = base("Lernziele");
  txt(s, D.LERNZIELE.map(([z, stern], i) => ({ text: z + (stern ? " *" : ""), options: { bullet: { indent: 24 }, breakLine: i < D.LERNZIELE.length - 1, paraSpaceAfter: 18 } })), { x: 0.8, y: 1.4, w: 11.8, h: 5.4 });
  s.addNotes("Lernziele vorlesen. Am Ende (Zielcheck) darauf zurückkommen.");
}

// ===== A: Ankommen – Bildimpuls: die Parabel als Bild (Variante A, eine Szene) =====
{
  const s = pres.addSlide();
  s.background = { color: WHITE };
  require("./bildimpuls")(pres).szene(s);
  s.addNotes("Ansage: „Stifte liegen. Schaut euch das Bild an.“ 30 Sekunden wirken lassen. Impuls: „Erzählt die Geschichte anhand des Bildes nach. Was sagt der Schutzmann?“ Zwei, drei Schüler erzählen nacheinander weiter; die Klasse ergänzt den Dialog („Von mir willst du den Weg erfahren?“ – „Ja, da ich ihn selbst nicht finden kann.“ – „Gibs auf, gibs auf.“). Elemente im Bild: früher Morgen, leere Straßen, Turmuhr (7:50) und eigene Uhr, Erzähler mit Koffer in Eile, Bahnhof ohne erkennbaren Weg, Schutzmann wendet sich lachend ab. Gelenk: „Das ist die Bildebene. Was ist die Sachebene?“");
}

// ===== V: Vorwissen aktivieren =====
{
  const s = base("Was wisst ihr noch?", { phase: "VORWISSEN" });
  D.VORWISSEN.forEach(([q, h], i) => {
    const y = 1.25 + i * 1.85;
    rect(s, 0.5, y, 12.3, 1.65, i % 2 ? WHITE : LIGHT, { line: "D5DCEC", lw: 0.75 });
    txt(s, [{ text: q, options: { bold: true, breakLine: true } }, { text: h, options: { color: MUTED, italic: true } }], { x: 0.8, y, w: 11.8, h: 1.65, valign: "middle" });
  });
  s.addNotes("Erst 1 Minute allein im Heft Stichworte, dann Stichproben (4–6 Schüler unterschiedlich). Antworten an die Tafel (① Wiederholung). Erwartung: Parabel = Bildebene (was erzählt wird) wird auf eine Sachebene (allgemeine menschliche Erfahrung) übertragen. Schutzmann = Instanz, die Orientierung geben sollte (Staat, Autorität, Gott – oder der Vater aus dem ‚Brief an den Vater‘). Stilmittel: Parataxe, Hypotaxe, Metapher, Symbol (Uhr = Zeitdruck, Weg = Lebensweg).");
}
{
  const s = base("„Gibs auf!“ – Franz Kafka", { phase: "VORWISSEN" });
  txt(s, D.TEXT.join(" "), { x: 0.5, y: 1.15, w: 12.3, h: 6.2, fontFace: HEAD, lineSpacingMultiple: 1.0 });
  s.addNotes("Arbeitsblatt austeilen (Text mit Zeilennummern, Nr. 1). Lehrkraft liest vor, die Klasse liest still mit. " + D.ORIGINAL_HINWEIS);
}
[false, true].forEach((loes) => {
  const s = base(loes ? "Satzbau: Lösung" : "Satzbau: " + D.SATZBAU.frage, { phase: "VORWISSEN" });
  [["Satz 1", D.SATZBAU.s1], ["Satz 2", D.SATZBAU.s2]].forEach(([k, v], i) => {
    const y = 1.2 + i * 2.95;
    rect(s, 0.5, y, 12.3, loes ? 2.75 : 2.75, i ? LIGHT2 : LIGHT);
    txt(s, [{ text: k + ":  ", options: { bold: true, color: RED } }, { text: loes ? v.loesung : (i ? "„Als ich eine Turmuhr mit meiner Uhr verglich, sah ich, dass es schon viel später war, als ich geglaubt hatte, ich musste mich sehr beeilen, … ich kannte mich in dieser Stadt noch nicht sehr gut aus.“" : "„" + v.text + "“"), options: loes ? {} : { italic: true } }], { x: 0.8, y, w: 11.8, h: 2.75, valign: "middle", fontFace: loes ? BODY : HEAD });
  });
  s.addNotes(loes ? "Klärung der Frage vom 2.11.: Satz 2 ist nicht eindeutig – er beginnt hypotaktisch und wird zur Reihung. Gerade das passt: Der Satz verliert die Ordnung wie der Erzähler." : "Erst allein überlegen (30 Sek.), dann Stichproben. Achtung: Satz 2 ist beides – genau darüber nachdenken lassen. Dann Lösung aufdecken.");
});

// ===== I: Informieren – neue Stilmittel =====
for (let k = 0; k < D.STILMITTEL.length; k += 2) {
  const s = base("Neue Stilmittel", { phase: "INFORMIEREN" });
  D.STILMITTEL.slice(k, k + 2).forEach(([b, kern, , test, , bsp], i) => {
    const y = 1.2 + i * 3.0;
    s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y, w: 12.3, h: 2.8, fill: { color: WHITE }, line: { color: RED, width: 2.5 } });
    txt(s, [{ text: b + ": ", options: { bold: true, color: RED } }, { text: kern }], { x: 0.8, y: y + 0.12, w: 11.8, h: 1.1 });
    txt(s, [{ text: "Beispiel: ", options: { bold: true, color: BLUE } }, { text: bsp, options: { italic: true, color: BLUE, fontFace: HEAD } }], { x: 0.8, y: y + 1.2, w: 11.8, h: 0.75 });
    txt(s, [{ text: "Test: ", options: { bold: true, color: GREEN } }, { text: test }], { x: 0.8, y: y + 1.95, w: 11.8, h: 0.75 });
  });
  s.addNotes(k ? "Alle Stilmittel stehen auf dem Arbeitsblatt (Nr. 2), mit den bekannten zur Wiederholung. Wofür? Stilmittel sind Bedeutungsträger: Sie zeigen, wie der Erzähler die Situation erlebt – das braucht ihr für jede Textanalyse." : "Kurz erklären, je ein Beispiel – bewusst nicht aus „Gibs auf!“, die Stellen im Text sucht die Klasse in Aufgabe 1 selbst. Typischer Fehler vorab: Nicht jede Frage im Text ist rhetorisch – entscheidend ist, ob eine Antwort erwartet wird.");
}
{
  const s = base("Ich mache vor: Wiederholung", { phase: "INFORMIEREN" });
  const [b, , bsp, , wirk] = D.STILMITTEL[0];
  rect(s, 0.5, 1.25, 12.3, 1.2, LIGHT);
  txt(s, bsp, { x: 0.8, y: 1.25, w: 11.8, h: 1.2, fontFace: HEAD, italic: true, valign: "middle", fontSize: 32 });
  txt(s, [
    { text: "1. Was steht da?  ", options: { bold: true } }, { text: "Zweimal direkt hintereinander „gibs auf“.", options: { breakLine: true } },
    { text: "2. Stilmittel:  ", options: { bold: true } }, { text: b + ".", options: { breakLine: true } },
    { text: "3. Wirkung:  ", options: { bold: true } }, { text: wirk },
  ], { x: 0.8, y: 2.75, w: 11.8, h: 4.0, paraSpaceAfter: 12 });
  s.addNotes("Laut denken, Schritt für Schritt: „Was steht da? → Welches Stilmittel? → Was bewirkt es?“ Genau diese drei Schritte machen die Schüler in Aufgabe 1.");
}

// ===== V: Verarbeiten =====
aufgabe(D.AUFGABE1, "VERARBEITEN", "Ansage: „Jetzt arbeitet ihr allein. Die Aufgabe steht auf der Folie.“ Herumgehen. Danach die Lösungen nacheinander aufdecken, alle verbessern mit Grün.");
D.STILMITTEL.slice(1).forEach(([b, , bsp, , wirk]) => {
  const s = base("Lösung: " + b, { phase: "VERARBEITEN" });
  rect(s, 0.5, 1.3, 12.3, 1.5, LIGHT);
  txt(s, [{ text: bsp, options: { italic: true } }, { text: "  (" + D.zeile(bsp) + ")", options: { color: MUTED } }], { x: 0.8, y: 1.3, w: 11.8, h: 1.5, fontFace: HEAD, valign: "middle" });
  txt(s, [{ text: "Wirkung: ", options: { bold: true, color: GREEN } }, { text: wirk }], { x: 0.8, y: 3.1, w: 11.8, h: 3.5 });
  s.addNotes("Erst eine Schülerantwort vom Platz hören, dann aufdecken. Mit Grün verbessern.");
});
aufgabe(D.AUFGABE2, "VERARBEITEN", "Ansage: „Jetzt schreibt ihr allein. Ihr müsst nichts vorlesen, was ihr nicht vorlesen wollt.“ Timer 10 Minuten. Die Schreibhilfe steht auf dem Arbeitsblatt bei Aufgabe 2 (und auf den nächsten Folien). Wer schnell fertig ist: ein zweites Stilmittel einbauen.");
// Schreibhilfe auf zwei Folien (27 pt): Wörtliche Rede + Komma, dann Stilmittel
[[0, 1], [2]].forEach((gruppen, k) => {
  const s = base("Schreibhilfe" + (k ? " (2)" : " (1)"), { phase: "VERARBEITEN" });
  gruppen.forEach((g, i) => {
    const [titel, arr] = D.SCHREIBHILFE[g];
    const w = gruppen.length === 2 ? 6.05 : 12.3, x = 0.5 + i * 6.25;
    rect(s, x, 1.25, w, 5.6, g === 2 ? "FFF8EC" : LIGHT);
    txt(s, titel, { x: x + 0.25, y: 1.35, w: w - 0.5, h: 0.6, bold: true, color: g === 2 ? GOLD : NAVY, fontFace: HEAD });
    txt(s, arr.join("\n"), { x: x + 0.25, y: 2.05, w: w - 0.5, h: 4.7, paraSpaceAfter: 12 });
  });
  s.addNotes(k ? "Stilmittel als Muster – den Inhalt füllt jede und jeder selbst." : "Satzmuster ohne Inhalt – sie zeigen nur, wie Zeichensetzung geht. Achtung Komma: Nach der wörtlichen Rede steht ein Komma vor „sagte er“.");
});

// ===== A: Auswerten =====
aufgabe(D.AUFGABE3, "AUSWERTEN", "* Wenn die Zeit knapp ist, entfällt diese Phase – der Zielcheck bleibt. Ein freiwilliger Text wird vorgelesen. Alle notieren zuerst den abweisenden Satz und seinen Appell auf dem Arbeitsblatt (so sind alle aktiv), dann gemeinsam das Quadrat an der Tafel (②). Abschlussfrage: Ist das auch eine Parabel – was wäre die Sachebene?", (s) => {
  const q = [["Sachinhalt", LIGHT, NAVY], ["Selbstkundgabe", LIGHT, NAVY], ["Beziehung", LIGHT, NAVY], ["Appell", "FDECEA", RED]];
  q.forEach(([k, fill, col], i) => {
    const x = 3.2 + (i % 2) * 3.5, y = 5.0 + Math.floor(i / 2) * 0.85;
    rect(s, x, y, 3.35, 0.75, fill, { line: col === RED ? RED : "D5DCEC", lw: col === RED ? 2.5 : 0.75, r: 0.04 });
    txt(s, k, { x, y, w: 3.35, h: 0.75, bold: true, color: col, align: "center", valign: "middle" });
  });
});
{
  const s = base("Zielcheck", { time: "3 Min." });
  txt(s, "Bestimme das Stilmittel und erkläre die Wirkung in einem Satz. Allein, im Heft.", { x: 0.5, y: 1.15, w: 12.3, h: 0.9, italic: true, color: MUTED });
  D.ZIELCHECK.forEach(([satz], i) => {
    const y = 2.15 + i * 1.5;
    rect(s, 0.5, y, 12.3, 1.3, i % 2 ? WHITE : LIGHT, { line: "D5DCEC", lw: 0.75 });
    txt(s, (i + 1) + "  " + satz, { x: 0.8, y, w: 11.8, h: 1.3, fontFace: HEAD, valign: "middle" });
  });
  s.addNotes("Auswerten wird nie gestrichen. Danach Lösung aufdecken, mit Grün kontrollieren. Herumgehen: Wer noch Fehler hat, wird in der Wiederholung der nächsten Stunde gezielt aufgegriffen.");
}
{
  const s = base("Zielcheck: Lösung");
  D.ZIELCHECK.forEach(([satz, sm, wirk], i) => {
    const y = 1.2 + i * 1.95;
    rect(s, 0.5, y, 12.3, 1.8, i % 2 ? WHITE : LIGHT, { line: "D5DCEC", lw: 0.75 });
    txt(s, [{ text: (i + 1) + "  " + sm + ": ", options: { bold: true, color: GREEN } }, { text: wirk }], { x: 0.8, y, w: 11.8, h: 1.8, valign: "middle" });
  });
  s.addNotes("Mit Grün verbessern. Rückblick auf die Lernziele: Was davon kann ich jetzt?");
}

pres.writeFile({ fileName: "Praesentation_Gibs_auf.pptx" });
