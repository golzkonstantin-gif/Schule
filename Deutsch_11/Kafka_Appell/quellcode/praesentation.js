// Präsentation (16:9, Schrift durchgehend 27 pt): Testauswertung + Rechtschreibtraining, dann Kafkas Appell nach AVIVA
// + Whiteboard-PDF mit Redeauftrag und Kriterien
const pptxgen = require("pptxgenjs");
const {
  TEXT, QUELLE, APPELL, DEMO, SAETZE, ZUSAMMEN, vergleich, zaehle, REGELN,
  FORMULIERUNGSHILFEN, REDE_SITUATION, REDE_AUFTRAG, REDE_UMFANG, KRITERIEN,
} = require("./daten");

const NAVY = "1E2761", NAVY2 = "24306E", MUTED = "5B6B8C", LIGHT = "F4F7FD", ICE = "CADCFC", WHITE = "FFFFFF";
const RED = "D9534F", GOLD = "E8A33D", GREEN = "2E9E6B", BLUE = "3F7CC4";
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
    s.background = { color: WHITE };
    rect(s, 0, 0, 13.333, 1.0, opts.head || NAVY, { r: 0 });
    txt(s, title, { x: 0.5, y: 0, w: 10.2, h: 1.0, fontFace: HEAD, fontSize: 30, bold: true, color: WHITE, valign: "middle" });
    if (opts.time) { rect(s, 10.9, 0.15, 2.1, 0.7, RED, { r: 0.12 }); txt(s, opts.time, { x: 10.9, y: 0.15, w: 2.1, h: 0.7, bold: true, color: WHITE, align: "center", valign: "middle" }); }
    if (opts.phase) txt(s, opts.phase, { x: 9.3, y: 0.15, w: 3.7, h: 0.7, bold: true, color: ICE, align: "right", valign: "middle" });
    return s;
  }
  function dark(fn) {
    const s = pres.addSlide();
    s.background = { color: NAVY };
    s.addShape(pres.shapes.OVAL, { x: 10.6, y: -1.6, w: 4.6, h: 4.6, fill: { color: NAVY2 }, line: { type: "none" } });
    fn(s);
    return s;
  }
  // Übungssatz: Fehlerfassung oder korrigiert (Korrekturen rot, unterstrichen)
  const satz = (m, korrigiert, nr) => [
    ...(nr ? [{ text: nr + "  ", options: { bold: true, color: MUTED } }] : []),
    ...(korrigiert
      ? vergleich(m.f, m.r).map((x) => ({ text: x.text, options: x.fehler ? { color: RED, bold: true, underline: { style: "sng" } } : {} }))
      : [{ text: m.f }]),
  ];
  return { pres, txt, rect, base, dark, satz };
}

// Folie mit Redeauftrag und Kriterien (auch als Whiteboard-PDF)
function redeFolie(d, notes) {
  const { txt, rect, base } = d;
  const s = base("Transfer: Kafkas Rede an den Vater");
  rect(s, 0.5, 1.2, 12.3, 1.95, LIGHT);
  txt(s, [{ text: REDE_SITUATION, options: { breakLine: true } }, { text: REDE_AUFTRAG, options: { bold: true } }], { x: 0.75, y: 1.25, w: 11.8, h: 1.85, fontFace: HEAD, valign: "middle" });
  txt(s, REDE_UMFANG, { x: 0.5, y: 3.25, w: 12.3, h: 0.5, italic: true, color: MUTED });
  KRITERIEN.forEach(([k, e], i) => {
    const y = 3.85 + i * 0.7;
    rect(s, 0.5, y + 0.03, 0.58, 0.58, NAVY, { r: 0.29 });
    txt(s, String(i + 1), { x: 0.5, y: y + 0.03, w: 0.58, h: 0.58, bold: true, color: WHITE, align: "center", valign: "middle" });
    txt(s, [{ text: k + ":  ", options: { bold: true } }, { text: e }], { x: 1.3, y, w: 11.5, h: 0.64, valign: "middle" });
  });
  if (notes) s.addNotes(notes);
  return s;
}

if (require.main === module) {
  const d = deck("Rechtschreibtraining und Kafkas Appell");
  const { pres, txt, rect, base, dark, satz } = d;
  const allein = zaehle(SAETZE.slice(ZUSAMMEN));

  // ===== Teil 1: Testauswertung und Rechtschreibtraining (ca. 15 Min.) =====
  {
    const s = base("Auswertung des Tests");
    rect(s, 0.5, 1.3, 12.3, 1.3, LIGHT);
    txt(s, [{ text: "Durchschnitt:  ", options: { bold: true } }, { text: "____ von ____ Punkten" }], { x: 0.8, y: 1.3, w: 11.8, h: 1.3, fontSize: 32, valign: "middle" });
    txt(s, [{ text: "Größtes Problem: ", options: { bold: true, color: RED } }, { text: "Rechtschreibung", options: { bold: true } }], { x: 0.5, y: 2.95, w: 12.3, h: 0.6 });
    txt(s, [
      { text: "Substantive werden kleingeschrieben.", options: { bullet: { indent: 24 }, breakLine: true } },
      { text: "Kommas vor Nebensätzen fehlen.", options: { bullet: { indent: 24 } } },
    ], { x: 0.8, y: 3.65, w: 11.8, h: 1.3, paraSpaceAfter: 8 });
    rect(s, 0.5, 5.4, 12.3, 1.2, "E3F1E8");
    txt(s, "Ab jetzt: kleine Trainingseinheiten in jeder Stunde.", { x: 0.8, y: 5.4, w: 11.8, h: 1.2, bold: true, color: GREEN, valign: "middle" });
    s.addNotes("Durchschnittliche Punktzahl vor der Stunde in der pptx eintragen. Ansage: „Inhaltlich wisst ihr viel – Punkte verliert ihr bei der Sprache. Die meisten Fehler sind immer dieselben zwei. Das trainieren wir jetzt.“ Schnipsel austeilen.");
  }
  REGELN.forEach((r, i) => {
    const s = base("Regel " + (i + 1) + ": " + r.titel);
    rect(s, 0.5, 1.25, 12.3, 1.0, i ? "FDECEA" : "E3F1E8");
    txt(s, r.kern, { x: 0.8, y: 1.25, w: 11.8, h: 1.0, fontFace: HEAD, bold: true, valign: "middle", color: i ? RED : GREEN });
    txt(s, r.punkte.map((pt, j) => ({ text: pt, options: { bullet: { indent: 24 }, breakLine: j < r.punkte.length - 1, paraSpaceAfter: 14 } })), { x: 0.8, y: 2.55, w: 11.8, h: 4.5 });
    s.addNotes(i ? "Am Beispiel zeigen: „…, dass sein Sohn Angst hatte.“ – Wo steht das Verb? Am Ende, also Nebensatz, also Komma." : "Artikelprobe vormachen: „Vater – der Vater? Passt, also groß.“ „Zuhören – beim Zuhören: Signalwort, also groß.“");
  });
  // Zeigen, wie es geht
  [false, true].forEach((korr) => {
    const s = base(korr ? "So geht es: korrigiert" : "So geht es: Ich mache vor");
    DEMO.forEach((m, i) => {
      rect(s, 0.5, 1.35 + i * 1.65, 12.3, 1.4, LIGHT);
      txt(s, satz(m, korr), { x: 0.8, y: 1.35 + i * 1.65, w: 11.8, h: 1.4, fontFace: HEAD, valign: "middle" });
    });
    txt(s, korr
      ? [{ text: "Essen, Vater, Sohn, Angst", options: { bold: true, color: RED } }, { text: " → Artikelprobe / Signalwort beim → groß", options: { breakLine: true } }, { text: "nicht, dass", options: { bold: true, color: RED } }, { text: " → Verb am Ende („hatte“) → Komma" }]
      : "Ich denke laut: Artikelprobe? · Signalwort? · Verb am Ende?", { x: 0.5, y: 4.85, w: 12.3, h: 1.6, italic: !korr, color: korr ? NAVY : MUTED, paraSpaceAfter: 8 });
    s.addNotes(korr ? "Die Korrekturen zeigen und noch einmal begründen." : "Wort für Wort laut prüfen: „Beim essen – beim: Signalwort, also groß: beim Essen.“ „Der vater – der Vater: groß. … nicht dass – Verb am Ende? ‚hatte‘ – ja, also Komma vor dass.“");
  });
  // Übung zusammen (Sätze 1–2)
  [false, true].forEach((korr) => {
    const s = base(korr ? `Zusammen: Lösung Satz 1–${ZUSAMMEN}` : `Zusammen: Satz 1–${ZUSAMMEN} auf dem Schnipsel`);
    SAETZE.slice(0, ZUSAMMEN).forEach((m, i) => {
      rect(s, 0.5, 1.35 + i * 1.65, 12.3, 1.4, LIGHT);
      txt(s, satz(m, korr, i + 1), { x: 0.8, y: 1.35 + i * 1.65, w: 11.8, h: 1.4, fontFace: HEAD, valign: "middle" });
    });
    txt(s, korr ? "Satz 2: Der Nebensatz steht vorne – das Komma kommt danach." : "Ihr sagt an, ich schreibe mit – und ihr begründet mit der Regel.", { x: 0.5, y: 4.85, w: 12.3, h: 1.2, color: MUTED, italic: true });
    s.addNotes(korr ? "Kurz bestätigen, dann direkt zur Einzelarbeit." : "Schülerinnen und Schüler sagen die Korrekturen an und begründen (Artikelprobe / Verb am Ende). Erst danach weiter.");
  });
  {
    const s = base(`Jetzt allein: Satz ${ZUSAMMEN + 1}–${SAETZE.length}`, { time: "4 Min." });
    txt(s, [
      { text: "Korrigiere die Sätze auf deinem Schnipsel.", options: { bold: true, breakLine: true } },
      { text: " ", options: { breakLine: true, fontSize: 12 } },
      { text: `Es sind ${allein.gross + allein.komma} Fehler:  ${allein.gross} × Großschreibung,  ${allein.komma} × Komma.`, options: { breakLine: true } },
      { text: " ", options: { breakLine: true, fontSize: 12 } },
      { text: "Achtung: Nicht jeder Satz hat ein Komma.", options: { color: RED, bold: true, breakLine: true } },
      { text: " ", options: { breakLine: true, fontSize: 12 } },
      { text: "Fertig? Zähle, wie viele du gefunden hast.", options: { color: MUTED, italic: true } },
    ], { x: 0.6, y: 1.4, w: 12.1, h: 5.2 });
    s.addNotes("Timer 4 Minuten sichtbar. Tempo betonen. Nicht helfen – die Regeln stehen auf dem Schnipsel.");
  }
  // Lösung allein (je 3 Sätze)
  const rest = SAETZE.slice(ZUSAMMEN);
  for (let k = 0; k < rest.length; k += 3) {
    const teil = rest.slice(k, k + 3);
    const von = ZUSAMMEN + k + 1, bis = von + teil.length - 1;
    const s = base(`Lösung: Satz ${von}–${bis}`);
    teil.forEach((m, i) => {
      rect(s, 0.5, 1.25 + i * 1.95, 12.3, 1.75, LIGHT);
      txt(s, satz(m, true, von + i), { x: 0.8, y: 1.25 + i * 1.95, w: 11.8, h: 1.75, fontFace: HEAD, valign: "middle" });
    });
    s.addNotes(teil.some((m) => m.r.includes(" als ")) ? "Erst die Klasse nennen lassen, dann zeigen. Satz 7: „als jede Drohung“ ist kein Nebensatz (kein Verb am Ende) – also kein Komma." : "Erst die Klasse nennen lassen: „Was hast du in Satz … korrigiert – und warum?“ Dann die Folie zeigen.");
  }

  // ===== Teil 2: Kafka nach AVIVA (ca. 30 Min.) =====
  dark((s) => {
    rect(s, 0.9, 1.4, 3.2, 0.6, RED, { r: 0.3 });
    txt(s, "ANKOMMEN", { x: 0.9, y: 1.4, w: 3.2, h: 0.6, bold: true, color: WHITE, align: "center", valign: "middle" });
    txt(s, "Ich habe euch eben kritisiert.", { x: 0.9, y: 2.4, w: 11.4, h: 0.8, fontFace: HEAD, fontSize: 32, color: ICE });
    txt(s, "Was ist mein Appell an euch?", { x: 0.9, y: 3.3, w: 11.4, h: 1.2, fontFace: HEAD, fontSize: 40, bold: true, color: WHITE });
  }).addNotes("Antworten sammeln (z. B. „Achtet beim Schreiben auf Großschreibung und Kommas!“). Rückfrage: Offen oder versteckt? – „Ich habe nur gesagt, was schlecht war. Den Appell musstet ihr selbst heraushören.“ Überleitung zur Kommunikation.");
  {
    const s = base("Was wisst ihr noch?", { phase: "VORWISSEN" });
    const boxes = [["Wie kommunizieren wir?", "verbal · paraverbal · nonverbal"], ["Wie viele Ohren hat eine Nachricht?", "Schulz von Thun"], ["Welche zwei Ebenen hat jede Nachricht?", "Watzlawick"]];
    boxes.forEach(([q, h], i) => {
      const y = 1.3 + i * 1.85;
      rect(s, 0.5, y, 12.3, 1.65, i % 2 ? WHITE : LIGHT, { line: "D5DCEC", lw: 0.75 });
      txt(s, [{ text: q, options: { bold: true, breakLine: true } }, { text: h, options: { color: MUTED, italic: true } }], { x: 0.8, y, w: 11.8, h: 1.65, valign: "middle" });
    });
    s.addNotes("Schnelle Abfrage, Antworten an die Tafel (Wiederholung links): verbal / paraverbal / nonverbal · vier Ohren: Sachinhalt, Selbstkundgabe, Beziehung, Appell · Inhalts- und Beziehungsebene (Beziehung bestimmt, wie der Inhalt ankommt).");
  }
  {
    const s = base("Kafka und sein Vater", { phase: "VORWISSEN" });
    txt(s, [
      { text: "Die Beziehung ist gestört – der Sohn hat Furcht, der Vater versteht ihn nicht.", options: { bullet: { indent: 24 }, breakLine: true, paraSpaceAfter: 16 } },
      { text: "Max Brod hat deshalb an den Vater geschrieben.", options: { bullet: { indent: 24 }, breakLine: true, paraSpaceAfter: 16 } },
      { text: "Heute: Was will Kafka selbst von seinem Vater?", options: { bullet: { indent: 24 }, bold: true } },
    ], { x: 0.8, y: 1.4, w: 11.8, h: 3.6 });
    rect(s, 0.5, 5.3, 12.3, 1.2, "FDECEA");
    txt(s, "Wir untersuchen die Appell-Ebene.", { x: 0.8, y: 5.3, w: 11.8, h: 1.2, fontFace: HEAD, bold: true, color: RED, valign: "middle" });
    s.addNotes("Kurz an die letzte Stunde (Mentorin) anknüpfen: Selbstkundgabe, Brods Brief. Ziel der Stunde an die Tafel.");
  }
  {
    const s = base("Information: Der Appell", { phase: "INFORMIEREN" });
    APPELL.forEach(([k, e], i) => {
      const y = 1.3 + i * 1.85;
      rect(s, 0.5, y, 12.3, 1.65, i ? LIGHT : "FDECEA");
      txt(s, [{ text: k + ": ", options: { bold: true, color: i ? NAVY : RED } }, { text: e }], { x: 0.8, y, w: 11.8, h: 1.65, valign: "middle" });
    });
    s.addNotes("Steht auch auf dem Arbeitsblatt (2. Information). Rückbezug: „Mein Appell vorhin war verdeckt – ich habe nur beschrieben, was schlecht war.“");
  }
  {
    const s = base("Worum bittet Kafka?", { phase: "INFORMIEREN" });
    txt(s, TEXT.join(" "), { x: 0.5, y: 1.25, w: 12.3, h: 5.4, fontFace: HEAD, lineSpacingMultiple: 1.1 });
    txt(s, QUELLE, { x: 0.5, y: 6.85, w: 12.3, h: 0.5, color: MUTED, align: "right" });
    s.addNotes("Kontext-Box vom Arbeitsblatt vorlesen, dann den Auszug. Die Klasse liest auf dem Arbeitsblatt mit (Zeilennummern, Fußnoten).");
  }
  {
    const s = base("Arbeitsblatt: Aufgaben 1–3", { time: "8 Min." });
    const a = [["1", "Vorarbeit", "Worum bittet Kafka? (ein Satz)"], ["2", "Analyse", "drei Leitfragen, mit Zitat und Zeile"], ["3", "Visualisierung", "Kommunikationsquadrat, Fokus Appell"]];
    a.forEach(([n, k, e], i) => {
      const y = 1.35 + i * 1.6;
      rect(s, 0.5, y, 12.3, 1.4, i % 2 ? WHITE : LIGHT, { line: "D5DCEC", lw: 0.75 });
      txt(s, [{ text: "Aufgabe " + n + ": " + k, options: { bold: true, breakLine: true } }, { text: e, options: { color: MUTED } }], { x: 0.8, y, w: 11.8, h: 1.4, valign: "middle" });
    });
    txt(s, "Allein, im Hefter. Formulierungshilfen stehen auf dem Arbeitsblatt.", { x: 0.5, y: 6.3, w: 12.3, h: 0.6, italic: true, color: MUTED });
    s.addNotes("VERARBEITEN. Herumgehen. Wer bei Aufgabe 1 hängt: „Was soll der Vater nach diesem Satz anders machen?“");
  }
  {
    const s = base("Formulierungshilfen", { phase: "AUSWERTEN" });
    txt(s, FORMULIERUNGSHILFEN.map((x, i) => ({ text: x, options: { bullet: { indent: 24 }, breakLine: i < FORMULIERUNGSHILFEN.length - 1, paraSpaceAfter: 16 } })), { x: 0.8, y: 1.4, w: 11.8, h: 5.0, italic: true });
    s.addNotes("Während der Besprechung stehen lassen: Antworten in ganzen Sätzen einfordern. Ergebnisse ins Quadrat an der Tafel (Appell groß). Achtung Komma: „…, dass der Vater …“ – Regel 2!");
  }
  dark((s) => {
    rect(s, 0.9, 1.4, 3.2, 0.6, RED, { r: 0.3 });
    txt(s, "REFLEXION", { x: 0.9, y: 1.4, w: 3.2, h: 0.6, bold: true, color: WHITE, align: "center", valign: "middle" });
    txt(s, "Kafka hat den Brief nie übergeben.", { x: 0.9, y: 2.4, w: 11.4, h: 0.9, fontFace: HEAD, fontSize: 32, bold: true, color: WHITE });
    txt(s, [{ text: "Warum appelliert er so vorsichtig und versteckt?", options: { breakLine: true } }, { text: "Hätte der Vater die Bitte verstanden?" }], { x: 0.9, y: 3.6, w: 11.4, h: 1.6, color: ICE, paraSpaceAfter: 10 });
  }).addNotes("Aufgabe 4 im Gespräch. Ziel: Furcht → verdeckter, abgesicherter Appell. Merksatz an die Tafel.");
  redeFolie(d, "Aufgabe 5. Überleitung: „Max Brods Brief war erfolgreich – der Vater hört zu. Jetzt spricht Franz selbst, und zwar offen.“ In der Stunde beginnen; vorgetragen wird in der nächsten Stunde (keine Hausaufgabe). Hinweis: Großschreibung und Kommas wie im Training!");

  pres.writeFile({ fileName: "Praesentation_Kafkas_Appell.pptx" });

  const wb = deck("Rede an den Vater – Auftrag und Kriterien");
  redeFolie(wb);
  wb.pres.writeFile({ fileName: "Whiteboard_Rede_Auftrag_Kriterien.pptx" });
}

module.exports = { deck };
