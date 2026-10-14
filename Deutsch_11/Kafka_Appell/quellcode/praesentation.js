// Präsentation „Rechtschreib-Check + Kafkas Appell“ (16:9, Schrift durchgehend 27 pt) + Whiteboard-PDF zur Rede
const pptxgen = require("pptxgenjs");
const { TEXT, QUELLE, APPELL, CHAT, vergleich, zaehle, REGELN, LEITFRAGEN, REDE_AUFTRAG, REDE_UMFANG, KRITERIEN } = require("./daten");

const NAVY = "1E2761", NAVY2 = "24306E", MUTED = "5B6B8C", LIGHT = "F4F7FD", ICE = "CADCFC", WHITE = "FFFFFF";
const RED = "D9534F", GOLD = "E8A33D", GREEN = "2E9E6B", BLUE = "3F7CC4";
const WA_BG = "ECE5DD", WA_ME = "DCF8C6", WA_HEAD = "075E54";
const HEAD = "Cambria", BODY = "Calibri";
const FS = 27; // Schriftgröße auf allen Folien

function deck(title) {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_WIDE"; // 13.333 x 7.5
  pres.title = title;
  const txt = (s, text, o) => s.addText(text, Object.assign({ isTextBox: true, fontFace: BODY, fontSize: FS, color: NAVY, margin: 0, valign: "top" }, o));
  const rect = (s, x, y, w, h, fill, o = {}) => s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, fill: { color: fill }, line: o.line ? { color: o.line, width: o.lw || 1.5 } : { type: "none" }, rectRadius: o.r ?? 0.08 });
  function base(title, opts = {}) {
    const s = pres.addSlide();
    s.background = { color: opts.bg || WHITE };
    rect(s, 0, 0, 13.333, 1.0, opts.head || NAVY, { r: 0 });
    txt(s, title, { x: 0.5, y: 0, w: 10.2, h: 1.0, fontFace: HEAD, fontSize: 30, bold: true, color: WHITE, valign: "middle" });
    if (opts.time) { rect(s, 10.9, 0.15, 2.1, 0.7, RED, { r: 0.12 }); txt(s, opts.time, { x: 10.9, y: 0.15, w: 2.1, h: 0.7, bold: true, color: WHITE, align: "center", valign: "middle" }); }
    return s;
  }
  function dark(fn) {
    const s = pres.addSlide();
    s.background = { color: NAVY };
    s.addShape(pres.shapes.OVAL, { x: 10.6, y: -1.6, w: 4.6, h: 4.6, fill: { color: NAVY2 }, line: { type: "none" } });
    fn(s);
    return s;
  }
  // Chatblase im WhatsApp-Stil; korrigiert = Korrekturen rot hervorheben
  function bubble(s, m, y, h, korrigiert) {
    const w = 9.6, x = m.seite === "r" ? 13.333 - 0.5 - w : 0.5;
    rect(s, x, y, w, h, m.seite === "r" ? WA_ME : WHITE, { r: 0.12 });
    const runs = korrigiert
      ? vergleich(m.f, m.r).map((t) => ({ text: t.text, options: t.fehler ? { color: RED, bold: true, underline: { style: "sng" } } : {} }))
      : [{ text: m.f }];
    const name = m.seite === "l" ? [{ text: m.von, options: { bold: true, color: GREEN, breakLine: true } }] : [];
    txt(s, [...name, ...runs], { x: x + 0.25, y: y + 0.08, w: w - 1.3, h: h - 0.16, color: "111111", valign: "middle" });
    txt(s, m.zeit, { x: x + w - 1.05, y: y + h - 0.5, w: 0.9, h: 0.42, fontSize: 18, color: "8A8A8A", align: "right" });
  }
  return { pres, txt, rect, base, dark, bubble };
}

// Folie mit Redeauftrag und Kriterien (auch als Whiteboard-PDF)
function redeFolie(d, notes) {
  const { txt, rect, base } = d;
  const s = base("Transfer: Kafkas Rede an den Vater", { time: "12 Min." });
  rect(s, 0.5, 1.2, 12.3, 1.75, LIGHT);
  txt(s, REDE_AUFTRAG, { x: 0.75, y: 1.25, w: 11.8, h: 1.1, fontFace: HEAD, bold: true, valign: "middle" });
  txt(s, REDE_UMFANG, { x: 0.75, y: 2.35, w: 11.8, h: 0.5, italic: true, color: MUTED });
  KRITERIEN.forEach(([k, e], i) => {
    const y = 3.15 + i * 0.84;
    rect(s, 0.5, y, 0.66, 0.66, NAVY, { r: 0.33 });
    txt(s, String(i + 1), { x: 0.5, y, w: 0.66, h: 0.66, bold: true, color: WHITE, align: "center", valign: "middle" });
    txt(s, [{ text: k + ":  ", options: { bold: true } }, { text: e }], { x: 1.4, y, w: 11.4, h: 0.66, valign: "middle" });
  });
  if (notes) s.addNotes(notes);
  return s;
}

if (require.main === module) {
  const d = deck("Rechtschreib-Check und Kafkas Appell");
  const { pres, txt, rect, base, dark, bubble } = d;
  const alle = zaehle(CHAT), rest = zaehle(CHAT.slice(1));

  // ===== Teil 1: Rechtschreib-Check (10 Min.) =====
  dark((s) => {
    rect(s, 0.9, 1.6, 4.2, 0.6, RED, { r: 0.3 });
    txt(s, "TEIL 1 · 10 MIN.", { x: 0.9, y: 1.6, w: 4.2, h: 0.6, bold: true, color: WHITE, align: "center", valign: "middle" });
    txt(s, "Rechtschreib-Check", { x: 0.9, y: 2.5, w: 11.5, h: 1.0, fontFace: HEAD, fontSize: 44, bold: true, color: WHITE });
    txt(s, "Zwei Regeln · ein Chat · schnell und genau", { x: 0.9, y: 3.7, w: 11.5, h: 0.7, color: ICE });
  }).addNotes("Ansage: „In euren Texten passieren immer dieselben zwei Fehler – nicht, weil ihr die Regel nicht kennt, sondern weil ihr beim Schreiben nicht darauf achtet. Heute trainieren wir genau das.“ Schnipsel jetzt austeilen (Text nach unten).");

  REGELN.forEach((r, i) => {
    const s = base("Regel " + (i + 1) + ": " + r.titel);
    rect(s, 0.5, 1.25, 12.3, 1.0, i ? "FDECEA" : "E3F1E8");
    txt(s, r.kern, { x: 0.8, y: 1.25, w: 11.8, h: 1.0, fontFace: HEAD, bold: true, valign: "middle", color: i ? RED : GREEN });
    txt(s, r.punkte.map((pt, j) => ({ text: pt, options: { bullet: { indent: 24 }, breakLine: j < r.punkte.length - 1, paraSpaceAfter: 14 } })), { x: 0.8, y: 2.55, w: 11.8, h: 4.5 });
    s.addNotes(i ? "Am Beispiel zeigen: Wo steht das Verb? „…, dass mein Vater das Handy weglegt.“ – Verb am Ende, also Nebensatz, also Komma." : "Artikelprobe vormachen: „Vater – der Vater? Passt. Also groß.“ „Zuhören – beim Zuhören: Signalwort beim, also groß.“");
  });

  // Vormachen an Nachricht 1
  [false, true].forEach((korr) => {
    const s = base(korr ? "Ich mache vor: korrigiert" : "Ich mache vor: Wo sind die Fehler?", { bg: WA_BG, head: WA_HEAD });
    bubble(s, CHAT[0], 1.5, 1.7, korr);
    if (korr) txt(s, [
      { text: "vater, geduld", options: { bold: true, color: RED } }, { text: " → der Vater, die Geduld: Artikelprobe → groß", options: { breakLine: true } },
      { text: "glaube, dass", options: { bold: true, color: RED } }, { text: " → Verb am Ende („hatte“) → Nebensatz → Komma" },
    ], { x: 0.5, y: 3.7, w: 12.3, h: 1.6, paraSpaceAfter: 10 });
    else txt(s, "Ich denke laut: Artikelprobe? · Verb am Ende?", { x: 0.5, y: 3.7, w: 12.3, h: 0.7, italic: true, color: MUTED });
    s.addNotes(korr ? "Die drei Fehler zeigen und laut begründen. Dann sofort weiter zum Auftrag." : "Laut denken, Wort für Wort: „Ich glaube dass … – dass, Verb am Ende? ‚hatte‘ – ja: Komma. vater – der Vater: groß. geduld – die Geduld: groß.“");
  });

  // Auftrag
  {
    const s = base("Jetzt ihr: Korrigiert den Chat", { time: "4 Min." });
    txt(s, [
      { text: "Korrigiere auf deinem Schnipsel die Nachrichten 2 bis 6.", options: { bold: true, breakLine: true } },
      { text: " ", options: { breakLine: true, fontSize: 12 } },
      { text: `Es sind ${rest.gross + rest.komma} Fehler:  ${rest.gross} × Großschreibung,  ${rest.komma} × Komma.`, options: { breakLine: true } },
      { text: " ", options: { breakLine: true, fontSize: 12 } },
      { text: "Fertig? Zähle, wie viele du gefunden hast – und prüfe noch einmal mit den Regeln.", options: { color: MUTED, italic: true } },
    ], { x: 0.6, y: 1.5, w: 12.1, h: 4.5 });
    s.addNotes("Timer 4 Minuten, sichtbar. Tempo betonen: Ziel ist, schneller zu werden. Nicht helfen – die Regeln stehen auf dem Schnipsel.");
  }
  // Lösung
  [[1, 2], [3, 5]].forEach(([a, b], k) => {
    const s = base(`Lösung: Nachrichten ${a + 1} bis ${b + 1}`, { bg: WA_BG, head: WA_HEAD });
    const msgs = CHAT.slice(a, b + 1), h = k ? 1.45 : 2.25, gap = 0.25;
    msgs.forEach((m, i) => bubble(s, m, 1.25 + i * (h + gap), h, true));
    s.addNotes(k ? `Gesamt: ${rest.gross + rest.komma} Fehler. Kurze Abfrage: Wer hat mindestens ${rest.gross + rest.komma - 3} gefunden? Dann Überleitung zu Kafka – „In Nachricht 2 ging es schon um ihn.“` : "Erst die Klasse nennen lassen: „Welche Fehler hast du in Nachricht 2?“ – dann die Folie zeigen. Kurz begründen lassen (Artikelprobe / Verb am Ende).");
  });

  // ===== Teil 2: Kafka (35 Min.) =====
  {
    const s = base("Worum bittet Kafka?");
    txt(s, TEXT.join(" "), { x: 0.5, y: 1.15, w: 12.3, h: 5.75, fontFace: HEAD, lineSpacingMultiple: 1.0 });
    txt(s, QUELLE, { x: 0.5, y: 6.95, w: 12.3, h: 0.45, color: MUTED, align: "right" });
    s.addNotes("Erst den Kontext vom Arbeitsblatt vorlesen (Kafka, 36 Jahre, Brief nie angekommen), dann den Auszug. Die Klasse liest auf dem Arbeitsblatt mit (dort mit Zeilennummern).");
  }
  dark((s) => {
    txt(s, "„Wobei ich Dich aber immerfort bitte …“", { x: 0.9, y: 1.4, w: 11.4, h: 1.0, fontFace: HEAD, fontSize: 32, italic: true, color: ICE });
    rect(s, 0.9, 2.9, 11.5, 1.6, RED);
    txt(s, "Worum bittet Kafka seinen Vater?", { x: 1.2, y: 2.9, w: 11, h: 1.6, fontFace: HEAD, fontSize: 36, bold: true, color: WHITE, valign: "middle" });
  }).addNotes("Zwei, drei spontane Antworten an die Tafel (① Vermutungen). Nicht bewerten.");
  {
    const s = base("Information: Der Appell");
    APPELL.forEach(([k, e], i) => {
      const y = 1.3 + i * 1.85;
      rect(s, 0.5, y, 12.3, 1.65, i ? LIGHT : "FDECEA");
      txt(s, [{ text: k + ": ", options: { bold: true, color: i ? NAVY : RED } }, { text: e }], { x: 0.8, y, w: 11.8, h: 1.65, valign: "middle" });
    });
    s.addNotes("Kurz an die Vier-Seiten-Modell-Stunde erinnern. Ein Beispiel aus dem Chat: „Ich wünsche mir, dass mein Vater das Handy weglegt“ – offen oder verdeckt?");
  }
  {
    const s = base("Auftrag 1: Kafkas Kernbotschaft", { time: "3 Min." });
    txt(s, "Fasse in einem eigenen Satz zusammen, worum Kafka seinen Vater bittet.", { x: 0.6, y: 1.4, w: 12.1, h: 1.3, fontFace: HEAD, bold: true });
    rect(s, 0.6, 3.0, 12.1, 1.2, LIGHT);
    txt(s, "„Kafka bittet seinen Vater, …“", { x: 0.9, y: 3.0, w: 11.6, h: 1.2, italic: true, valign: "middle" });
    txt(s, "In den Hefter. Zwei lesen danach vor.", { x: 0.6, y: 4.6, w: 12.1, h: 0.7, color: MUTED });
    s.addNotes("Zwei Sätze vorlesen lassen, den treffendsten an die Tafel (② Kernbotschaft).");
  }
  {
    const s = base("Auftrag 2: Kafkas Appelle untersuchen", { time: "5 Min." });
    txt(s, "Im Hefter – jeweils: offen oder verdeckt?", { x: 0.6, y: 1.15, w: 12.1, h: 0.6, bold: true });
    LEITFRAGEN.forEach(([spur, frage], i) => {
      const y = 1.9 + i * 1.75;
      rect(s, 0.5, y, 12.3, 1.6, i % 2 ? WHITE : LIGHT, { line: "D5DCEC", lw: 0.75 });
      txt(s, (i + 1) + "  " + frage, { x: 0.75, y: y + 0.08, w: 11.8, h: 0.62, bold: true });
      txt(s, spur, { x: 1.15, y: y + 0.72, w: 11.5, h: 0.8, fontFace: HEAD, italic: true, color: MUTED });
    });
    s.addNotes("Danach im Unterrichtsgespräch Frage für Frage, Ergebnisse ins Kommunikationsquadrat (③, Appell groß). Frage 3 führt zur Furcht.");
  }
  {
    const s = base("Formulierungshilfen");
    const fh = ["Ein offener Appell zeigt sich in … (Z. …).", "Verdeckt appelliert Kafka, wenn er …", "Damit will er erreichen, dass der Vater …", "Seine Bitte sichert er ab, indem er …"];
    txt(s, fh.map((x, i) => ({ text: x, options: { bullet: { indent: 24 }, breakLine: i < fh.length - 1, paraSpaceAfter: 16 } })), { x: 0.8, y: 1.4, w: 11.8, h: 5.0, italic: true });
    s.addNotes("Während Auftrag 2 stehen lassen oder im Gespräch zeigen, wenn Antworten nur in Stichworten kommen. Achtung Komma: „…, dass der Vater …“ – Regel 2!");
  }
  redeFolie(d, "Überleitung: „Max Brods Brief hat gewirkt – der Vater hört zu. Jetzt spricht Franz selbst, und zwar offen.“ Kriterien an die Tafel (④). Timer 12 Minuten. Hinweis: Auf Großschreibung und Kommas achten – wie im Chat!");
  {
    const s = base("Hört Franz zu – und hakt ab");
    txt(s, "Ohne Unterbrechung zuhören. Notiere im Hefter, welche Kriterien erfüllt sind.", { x: 0.5, y: 1.15, w: 12.3, h: 1.0 });
    KRITERIEN.forEach(([k], i) => {
      const y = 2.3 + i * 0.95;
      rect(s, 0.5, y, 8.0, 0.8, i % 2 ? WHITE : LIGHT, { line: "D5DCEC", lw: 0.75 });
      txt(s, (i + 1) + "  " + k, { x: 0.75, y, w: 7.6, h: 0.8, bold: true, valign: "middle" });
      ["Rede 1", "Rede 2"].forEach((r, j) => {
        rect(s, 8.75 + j * 2.1, y, 1.95, 0.8, i % 2 ? WHITE : LIGHT, { line: "D5DCEC", lw: 0.75 });
        txt(s, "☐  " + (j + 1), { x: 8.75 + j * 2.1, y, w: 1.95, h: 0.8, color: MUTED, align: "center", valign: "middle" });
      });
    });
    s.addNotes("Zwei Reden (bei Zeitnot eine). Danach Auswertung Kriterium für Kriterium: Wo war der offene Appell? Welche Ich-Botschaft war am stärksten?");
  }
  dark((s) => {
    txt(s, "Im Brief bittet Kafka vorsichtig und versteckt – in der Rede offen.", { x: 0.9, y: 1.5, w: 11.4, h: 1.7, fontFace: HEAD, fontSize: 32, bold: true, color: WHITE });
    txt(s, [{ text: "Warum hat Kafka im Brief nicht so gesprochen?", options: { breakLine: true } }, { text: "Welche Fassung hätte der Vater eher gehört?" }], { x: 0.9, y: 3.6, w: 11.4, h: 1.6, color: ICE, paraSpaceAfter: 10 });
  }).addNotes("Ziel: Rückbezug auf die Furcht. Merksatz an die Tafel (⑤).");

  pres.writeFile({ fileName: "Praesentation_Kafkas_Appell.pptx" });

  // Whiteboard-PDF: nur Auftrag und Kriterien
  const wb = deck("Rede an den Vater – Auftrag und Kriterien");
  redeFolie(wb);
  wb.pres.writeFile({ fileName: "Whiteboard_Rede_Auftrag_Kriterien.pptx" });
}

module.exports = { deck };
