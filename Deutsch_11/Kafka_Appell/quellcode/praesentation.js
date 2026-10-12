// Präsentation „Kafkas Appell“ (16:9) + Whiteboard-PDF mit Redeauftrag und Kriterien
const pptxgen = require("pptxgenjs");
const { TEXT, QUELLE, DIKTAT, clean, STOLPER, LEITFRAGEN, REDE_AUFTRAG, REDE_UMFANG, KRITERIEN } = require("./daten");

const NAVY = "1E2761", NAVY2 = "24306E", MUTED = "5B6B8C", LIGHT = "F4F7FD", LIGHT2 = "EAF0FA", ICE = "CADCFC", WHITE = "FFFFFF";
const RED = "D9534F", GOLD = "E8A33D", GREEN = "2E9E6B", BLUE = "3F7CC4";
const HEAD = "Cambria", BODY = "Calibri";

function deck(title) {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_WIDE"; // 13.333 x 7.5
  pres.title = title;
  const txt = (s, text, o) => s.addText(text, Object.assign({ isTextBox: true, fontFace: BODY, fontSize: 20, color: NAVY, margin: 0, valign: "top" }, o));
  const rect = (s, x, y, w, h, fill, o = {}) => s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, fill: { color: fill }, line: o.line ? { color: o.line, width: o.lw || 1.5 } : { type: "none" }, rectRadius: o.r ?? 0.08 });
  const pill = (s, label, x, y, w, col, size = 14) => { rect(s, x, y, w, 0.44, col, { r: 0.22 }); txt(s, label, { x, y, w, h: 0.44, fontSize: size, bold: true, color: WHITE, align: "center", valign: "middle" }); };
  function base(kicker, title, opts = {}) {
    const s = pres.addSlide();
    s.background = { color: WHITE };
    rect(s, 0, 0, 13.333, 1.05, NAVY, { r: 0 });
    txt(s, kicker.toUpperCase(), { x: 0.5, y: 0.12, w: 9.5, h: 0.3, fontSize: 13, bold: true, color: ICE, charSpacing: 2 });
    txt(s, title, { x: 0.5, y: 0.42, w: 10.3, h: 0.58, fontFace: HEAD, fontSize: 27, bold: true, color: WHITE });
    if (opts.time) { rect(s, 11.0, 0.18, 1.9, 0.7, RED, { r: 0.12 }); txt(s, opts.time, { x: 11.0, y: 0.18, w: 1.9, h: 0.7, fontSize: 20, bold: true, color: WHITE, align: "center", valign: "middle" }); }
    return s;
  }
  function dark(fn) {
    const s = pres.addSlide();
    s.background = { color: NAVY };
    s.addShape(pres.shapes.OVAL, { x: 10.6, y: -1.6, w: 4.6, h: 4.6, fill: { color: NAVY2 }, line: { type: "none" } });
    fn(s);
    return s;
  }
  return { pres, txt, rect, pill, base, dark };
}

// Folie mit Redeauftrag und Kriterien (auch als Whiteboard-PDF)
function redeFolie(d, notes) {
  const { txt, rect, base } = d;
  const s = base("Transfer · allein", "Kafkas Rede an den Vater", { time: "15 Min." });
  rect(s, 0.6, 1.3, 12.1, 1.25, LIGHT);
  txt(s, REDE_AUFTRAG, { x: 0.85, y: 1.38, w: 11.6, h: 0.85, fontFace: HEAD, fontSize: 21, bold: true, valign: "middle" });
  txt(s, REDE_UMFANG, { x: 0.85, y: 2.15, w: 11.6, h: 0.35, fontSize: 16, italic: true, color: MUTED });
  txt(s, "Deine Rede muss enthalten:", { x: 0.6, y: 2.75, w: 8, h: 0.45, fontFace: HEAD, fontSize: 20, bold: true, color: RED });
  KRITERIEN.forEach(([k, e], i) => {
    const y = 3.3 + i * 0.75;
    rect(s, 0.6, y, 0.6, 0.6, NAVY, { r: 0.3 });
    txt(s, String(i + 1), { x: 0.6, y, w: 0.6, h: 0.6, fontSize: 20, bold: true, color: WHITE, align: "center", valign: "middle" });
    txt(s, [{ text: k + ":  ", options: { bold: true } }, { text: e }], { x: 1.4, y, w: 11.3, h: 0.6, fontSize: 19, valign: "middle" });
  });
  if (notes) s.addNotes(notes);
  return s;
}

if (require.main === module) {
  const d = deck("Kafkas Appell");
  const { pres, txt, rect, pill, base, dark } = d;

  // 1 Auszug zum Mitlesen
  {
    const s = base(QUELLE, "Worum bittet Kafka?");
    const size = 21, lh = size / 72 * 1.55, x = 0.6, y = 1.45, w = 12.1;
    TEXT.forEach((l, i) => {
      txt(s, String(i + 1), { x, y: y + i * lh, w: 0.45, h: lh, fontSize: size - 7, color: MUTED, align: "right", valign: "middle" });
      txt(s, l, { x: x + 0.65, y: y + i * lh, w: w - 0.65, h: lh, fontFace: HEAD, fontSize: size, valign: "middle" });
    });
    s.addShape(pres.shapes.LINE, { x: x + 0.55, y, w: 0, h: TEXT.length * lh, line: { color: BLUE, width: 2 } });
    s.addNotes("Erst den Kontext vom Arbeitsblatt vorlesen (Kafka, 36 Jahre, Brief nie angekommen), dann den Auszug. Die Klasse liest mit.");
  }
  // 2 Impuls
  dark((s) => {
    txt(s, "„Wobei ich Dich aber immerfort bitte …“", { x: 0.9, y: 1.4, w: 11.4, h: 1.0, fontFace: HEAD, fontSize: 32, italic: true, color: ICE });
    rect(s, 0.9, 2.9, 11.5, 1.6, RED);
    txt(s, "Worum bittet Kafka seinen Vater?", { x: 1.2, y: 2.9, w: 11, h: 1.6, fontFace: HEAD, fontSize: 34, bold: true, color: WHITE, valign: "middle" });
  }).addNotes("Zwei, drei spontane Antworten an die Tafel (oben links, „Vermutungen“). Nicht bewerten – wir prüfen sie gleich am Text. Überleitung: „Dafür brauchen wir den Fachbegriff – den schreibt ihr jetzt mit.“");
  // 3 Diktat-Platzhalter
  dark((s) => {
    pill(s, "MINI-DIKTAT", 0.9, 2.0, 2.6, RED, 14);
    txt(s, "Der Appell – offen und verdeckt", { x: 0.9, y: 2.7, w: 11.8, h: 0.9, fontFace: HEAD, fontSize: 32, bold: true, color: WHITE });
    txt(s, "Schreibe mit. Satzzeichen werden nicht angesagt – achte selbst auf die Kommas.", { x: 0.9, y: 3.9, w: 11.5, h: 0.6, fontSize: 22, color: ICE });
  }).addNotes("Diese Folie steht während des Diktats (Vorlesefassung). Erst einmal ganz vorlesen, dann diktieren, dann noch einmal vorlesen. Danach weiter zur Lösung.");
  // 4 Diktat-Lösung + Stolperstellen
  {
    const s = base("Mini-Diktat · Selbstkontrolle", "Vergleiche Wort für Wort");
    txt(s, clean(DIKTAT), { x: 0.6, y: 1.3, w: 12.1, h: 2.6, fontFace: HEAD, fontSize: 22, lineSpacingMultiple: 1.25 });
    const cols = [RED, GOLD, GREEN, BLUE];
    STOLPER.forEach(([h, e], i) => {
      const x = 0.6 + (i % 2) * 6.15, y = 4.2 + Math.floor(i / 2) * 1.35;
      rect(s, x, y, 5.95, 1.2, LIGHT);
      rect(s, x, y, 0.12, 1.2, cols[i], { r: 0.02 });
      txt(s, h, { x: x + 0.3, y: y + 0.1, w: 5.5, h: 0.4, fontSize: 17, bold: true, color: cols[i] });
      txt(s, e, { x: x + 0.3, y: y + 0.5, w: 5.5, h: 0.65, fontSize: 15 });
    });
    s.addNotes("Fehler farbig anstreichen, Anzahl notieren. Keine Besprechung, nur Selbstkontrolle – dann die Stolperstellen kurz zeigen.");
  }
  // 5 Vorarbeit
  {
    const s = base("Auftrag 1 · allein", "Kafkas Kernbotschaft", { time: "4 Min." });
    txt(s, "Fasse in einem eigenen Satz zusammen, worum Kafka seinen Vater bittet.", { x: 0.6, y: 1.5, w: 12.1, h: 1.0, fontFace: HEAD, fontSize: 26, bold: true });
    rect(s, 0.6, 2.9, 12.1, 1.3, LIGHT);
    txt(s, "Satzanfang:  „Kafka bittet seinen Vater, …“", { x: 0.9, y: 2.9, w: 11.6, h: 1.3, fontSize: 24, italic: true, valign: "middle" });
    txt(s, "In den Hefter. Danach lesen zwei von euch ihren Satz vor.", { x: 0.6, y: 4.6, w: 12.1, h: 0.6, fontSize: 20, color: MUTED });
    s.addNotes("Zwei Sätze vorlesen lassen, den treffendsten an die Tafel (Kernbotschaft). Mit den Vermutungen vom Anfang vergleichen.");
  }
  // 6 Analyse
  {
    const s = base("Auftrag 2 · allein", "Kafkas Appelle untersuchen", { time: "6 Min." });
    txt(s, "Beantworte die drei Fragen im Hefter. Notiere jeweils: offener oder verdeckter Appell?", { x: 0.6, y: 1.3, w: 12.1, h: 0.55, fontSize: 20, bold: true });
    LEITFRAGEN.forEach(([spur, frage], i) => {
      const y = 2.0 + i * 1.0;
      rect(s, 0.6, y, 12.1, 0.85, i % 2 ? WHITE : LIGHT, { line: "D5DCEC", lw: 0.75 });
      txt(s, String(i + 1), { x: 0.75, y, w: 0.4, h: 0.85, fontSize: 22, bold: true, color: RED, valign: "middle" });
      txt(s, spur, { x: 1.25, y, w: 6.4, h: 0.85, fontFace: HEAD, fontSize: 16, italic: true, valign: "middle" });
      txt(s, frage, { x: 7.85, y, w: 4.7, h: 0.85, fontSize: 18, bold: true, valign: "middle" });
    });
    rect(s, 0.6, 5.2, 12.1, 1.95, "FFF8EC");
    txt(s, "Formulierungshilfen", { x: 0.9, y: 5.3, w: 6, h: 0.4, fontFace: HEAD, fontSize: 18, bold: true, color: GOLD });
    txt(s, "Ein offener Appell zeigt sich in … (Z. …).  ·  Verdeckt appelliert Kafka, wenn er …  ·  Damit will er erreichen, dass der Vater …  ·  Seine Bitte sichert er ab, indem er …", { x: 0.9, y: 5.75, w: 11.5, h: 1.3, fontSize: 17, italic: true });
    s.addNotes("Danach im Unterrichtsgespräch: Frage für Frage, Ergebnisse ins Kommunikationsquadrat an der Tafel (Appell-Feld groß). Kernfrage 3 führt zur Furcht: Kafka bittet vorsichtig, weil er Vorwürfe erwartet.");
  }
  // 7 Transfer
  redeFolie(d, "Überleitung: „Max Brods Brief hat gewirkt – der Vater hört zu. Jetzt spricht Franz selbst – und zwar offen.“ Kriterien stehen auch an der Tafel bzw. als Whiteboard-PDF. Timer 15 Minuten.");
  // 8 Vortrag
  {
    const s = base("Vortrag · zuhören", "Hört Franz zu – und hakt ab");
    txt(s, "Zwei Reden werden ohne Unterbrechung gehalten. Notiere für jede Rede im Hefter, welche Kriterien erfüllt sind.", { x: 0.6, y: 1.35, w: 12.1, h: 0.95, fontSize: 21 });
    KRITERIEN.forEach(([k], i) => {
      const y = 2.5 + i * 0.78;
      rect(s, 0.6, y, 7.6, 0.66, i % 2 ? WHITE : LIGHT, { line: "D5DCEC", lw: 0.75 });
      txt(s, (i + 1) + "  " + k, { x: 0.85, y, w: 5.2, h: 0.66, fontSize: 20, bold: true, valign: "middle" });
      ["Rede 1", "Rede 2"].forEach((r, j) => {
        rect(s, 8.5 + j * 2.15, y, 1.95, 0.66, i % 2 ? WHITE : LIGHT, { line: "D5DCEC", lw: 0.75 });
        if (i === 0) txt(s, r, { x: 8.5 + j * 2.15, y: 2.05, w: 1.95, h: 0.4, fontSize: 15, bold: true, color: MUTED, align: "center" });
        txt(s, "☐", { x: 8.5 + j * 2.15, y, w: 1.95, h: 0.66, fontSize: 22, color: MUTED, align: "center", valign: "middle" });
      });
    });
    s.addNotes("Erst beide Reden ohne Unterbrechung. Dann Auswertung Kriterium für Kriterium: Wo genau war der offene Appell? Welche Ich-Botschaft war am stärksten?");
  }
  // 9 Reflexion
  dark((s) => {
    pill(s, "REFLEXION", 0.9, 1.5, 2.3, RED, 13);
    txt(s, "Im Brief bittet Kafka vorsichtig und versteckt –\nin der Rede offen.", { x: 0.9, y: 2.2, w: 11.4, h: 1.5, fontFace: HEAD, fontSize: 30, bold: true, color: WHITE });
    txt(s, "Warum hat Kafka im Brief nicht so gesprochen?\nWelche Fassung hätte der Vater eher gehört?", { x: 0.9, y: 4.1, w: 11.4, h: 1.4, fontSize: 25, color: ICE, paraSpaceAfter: 6 });
  }).addNotes("Ziel: Rückbezug auf die Furcht – wer Angst hat, appelliert verdeckt und sichert sich ab. Merksatz an die Tafel.");

  pres.writeFile({ fileName: "Praesentation_Kafkas_Appell.pptx" });

  // Whiteboard-PDF: nur Auftrag und Kriterien
  const wb = deck("Rede an den Vater – Auftrag und Kriterien");
  redeFolie(wb);
  wb.pres.writeFile({ fileName: "Whiteboard_Rede_Auftrag_Kriterien.pptx" });
}
