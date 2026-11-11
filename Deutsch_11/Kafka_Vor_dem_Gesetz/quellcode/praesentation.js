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
  s.addNotes("Ziel und Ablauf stehen auch an der Seitentafel. Diesmal ohne Stern: Die Stunde ist bewusst schlank geplant.");
}
{
  const s = base("Lernziele");
  txt(s, D.LERNZIELE.map(([z], i) => ({ text: z, options: { bullet: { indent: 24 }, breakLine: i < D.LERNZIELE.length - 1, paraSpaceAfter: 18 } })), { x: 0.8, y: 1.4, w: 11.8, h: 5.4 });
  s.addNotes("Lernziele vorlesen. Am Ende (Zielcheck) darauf zurückkommen.");
}

// ===== V: Vorwissen – kafkaesk wiederholen, Begriffs-Sprint, Lösung =====
{
  const s = base("Was wisst ihr noch?", { phase: "VORWISSEN" });
  D.VORWISSEN.forEach(([q, h], i) => {
    const y = 1.35 + i * 2.6;
    rect(s, 0.5, y, 12.3, 2.3, i % 2 ? WHITE : LIGHT, { line: "D5DCEC", lw: 0.75 });
    txt(s, [{ text: q, options: { bold: true, breakLine: true } }, { text: h, options: { color: MUTED, italic: true } }], { x: 0.8, y, w: 11.8, h: 2.3, valign: "middle" });
  });
  s.addNotes("Mündlich, 2 Minuten: Stichproben (4–5 Schüler, unterschiedliche Niveaus). Noch nichts anschreiben – die Lösung kommt nach dem Sprint an die Tafel. Grundlage: Arbeitsblatt der Mentorin vom 9.11.");
}
{
  const s = base("Begriffs-Sprint", { time: "1 Min." });
  txt(s, bullets(D.SPRINT.regeln, { num: true, space: 16 }), { x: 0.6, y: 1.5, w: 12.1, h: 5.0 });
  s.addNotes("Regeln wie beim Sprint. Ansage: „Stift in die Hand. Ihr schreibt eine Minute ohne Pause. Der Stift bleibt auf dem Papier, bis ich ‚Stopp‘ sage.“ Erst wenn alle bereit sind: weiterklicken – dann erscheint der Begriff, Timer 1 Minute.");
}
{
  const s = pres.addSlide();
  s.background = { color: NAVY };
  txt(s, D.SPRINT.begriff, { x: 0, y: 2.2, w: 13.333, h: 2.2, fontFace: HEAD, fontSize: 110, bold: true, color: WHITE, align: "center", valign: "middle" });
  txt(s, "1 Minute – Stift bleibt auf dem Papier", { x: 0, y: 5.2, w: 13.333, h: 0.8, color: GOLD, align: "center" });
  s.addNotes("Timer 1 Minute. „Stopp.“ Dann lesen 2–3 Schüler vor (freiwillig oder aufgerufen). Danach die Lösung aufdecken (nächste Folie) und an die Tafel (①).");
}
{
  const s = base("„kafkaesk“ – Lösung", { phase: "VORWISSEN" });
  D.KAFKAESK.bausteine.forEach(([b, e], i) => {
    const x = 0.5 + (i % 2) * 6.25, y = 1.2 + Math.floor(i / 2) * 2.15;
    rect(s, x, y, 6.05, 1.95, i % 3 ? LIGHT : LIGHT2);
    txt(s, [{ text: b, options: { bold: true, color: RED, breakLine: true } }, { text: e }], { x: x + 0.25, y, w: 5.6, h: 1.95, valign: "middle" });
  });
  txt(s, [{ text: "Gefühle: ", options: { bold: true } }, { text: D.KAFKAESK.gefuehle }], { x: 0.5, y: 5.65, w: 12.3, h: 1.4, valign: "middle" });
  s.addNotes("Nach dem Vorlesen aufdecken: Was davon kam im Sprint vor? Die vier Bausteine und die Gefühle kommen an die Tafel (①). Definition (Arbeitsblatt der Mentorin): " + D.KAFKAESK.definition + " Steht auch auf dem Arbeitsblatt (Nr. 2).");
}

// ===== Aufgabe 1: Nervig oder kafkaesk? =====
aufgabe(D.NK_AUFGABE, "Szenarien-Blatt austeilen (halbes A4). Ansage: „Jetzt arbeitet ihr allein. Ihr schreibt nur Nummer, N oder K und den Baustein ins Heft.“ 3 Minuten.", (s) => {
  D.NK_KRITERIUM.forEach(([k, v], i) => {
    const x = 0.5 + i * 6.25;
    rect(s, x, 4.55, 6.05, 2.0, i ? "FDECEA" : LIGHT, { line: i ? RED : "D5DCEC", lw: i ? 2 : 0.75 });
    txt(s, [{ text: k, options: { bold: true, color: i ? RED : NAVY, breakLine: true } }, { text: v }], { x: x + 0.25, y: 4.55, w: 5.6, h: 2.0, valign: "middle" });
  });
});
[0, 3].forEach((k) => {
  const s = base("Nervig oder kafkaesk? – Lösung" + (k ? " (2)" : " (1)"), { phase: "VORWISSEN" });
  D.SZENARIEN.slice(k, k + 3).forEach(([stich, , nk, baustein, grund], j) => {
    const y = 1.2 + j * 1.95;
    rect(s, 0.5, y, 12.3, 1.8, nk === "K" ? "FDECEA" : LIGHT, { line: "D5DCEC", lw: 0.75 });
    txt(s, [{ text: (k + j + 1) + "  " + stich, options: { bold: true, breakLine: true } }, { text: nk === "K" ? "kafkaesk" : "nervig", options: { bold: true, color: nk === "K" ? RED : GREEN } }], { x: 0.75, y, w: 3.4, h: 1.8, valign: "middle" });
    txt(s, [...(baustein ? [{ text: baustein + ": ", options: { bold: true, color: RED } }] : []), { text: grund }], { x: 4.2, y, w: 8.4, h: 1.8, valign: "middle" });
  });
  s.addNotes("Erst eine Begründung vom Platz (mündlich, mit Baustein: „… ist kafkaesk, weil …“), dann aufdecken. Verbessern mit Grün.");
});

// ===== I: Informieren – Kontext mit Process-Satz, Lesen =====
{
  const s = base("„Vor dem Gesetz“ – Kontext", { phase: "INFORMIEREN" });
  txt(s, bullets(D.KONTEXT, { space: 12 }), { x: 0.7, y: 1.2, w: 12.0, h: 2.0 });
  rect(s, 0.5, 3.3, 12.3, 1.9, WHITE, { line: RED, lw: 2.5, r: 0.04 });
  txt(s, "„" + D.PROCESS_SATZ + "“", { x: 0.8, y: 3.3, w: 11.8, h: 1.9, valign: "middle", italic: true, fontFace: HEAD });
  rect(s, 0.5, 5.45, 12.3, 1.5, LIGHT2);
  txt(s, [{ text: "Leseauftrag: ", options: { bold: true, color: RED } }, { text: D.LESEAUFTRAG }], { x: 0.8, y: 5.45, w: 11.8, h: 1.5, valign: "middle" });
  s.addNotes("Impuls zum Process-Satz: „Welcher Baustein ist das?“ (grundlose Schuld). Arbeitsblatt austeilen. Ansage: „Stifte liegen. Ihr lest auf dem Arbeitsblatt mit.“ Die Lehrkraft liest den ganzen Text vor (Modell-Lesen, ca. 4 Min.: ruhig, Pausen an den Absätzen, die Rede des Türhüters deutlich). " + D.TEXT_HINWEIS);
}

// ===== Aufgabe 2: Einleitungssatz mit Deutungshypothese =====
{
  const s = base(D.TATT.titel.replace("Aufgabe 2: ", "Aufgabe 2: "), { time: D.TATT.zeit });
  txt(s, D.TATT.form, { x: 0.5, y: 1.15, w: 12.3, h: 0.5, italic: true, color: MUTED });
  rect(s, 0.5, 1.8, 12.3, 1.1, LIGHT);
  txt(s, [{ text: "Bausteine: ", options: { bold: true } }, { text: D.TATT.bausteine }], { x: 0.8, y: 1.8, w: 11.8, h: 1.1, valign: "middle" });
  txt(s, [{ text: "Gerüst: ", options: { bold: true, color: BLUE } }, { text: D.TATT.geruest, options: { italic: true } }], { x: 0.8, y: 3.2, w: 11.8, h: 1.6 });
  txt(s, "Tipp: Wie verläuft das Gespräch zwischen Türhüter und Mann?", { x: 0.8, y: 5.2, w: 11.8, h: 1.0, color: MUTED });
  s.addNotes("Gemeinsam, Ich – Wir: Die Lehrkraft beginnt („In der Parabel ‚Vor dem Gesetz‘ von Franz Kafka (1915) geht es um …“), die Klasse ergänzt das Thema. Dann Impuls zur Hypothese: „Wer hat im Gespräch die Macht? Wie nennen wir so eine Beziehung?“ → asymmetrisch / komplementär (Watzlawick). Satz an die Tafel (②), alle schreiben auf dem Arbeitsblatt mit. Hinweis Klausur: Eine Hypothese ist vorläufig – „Es lässt sich vermuten, dass …“.");
}
{
  const s = base("Einleitungssatz: Vorschlag", { phase: "INFORMIEREN" });
  rect(s, 0.5, 1.2, 12.3, 2.4, LIGHT);
  txt(s, D.TATT.satz, { x: 0.8, y: 1.2, w: 11.8, h: 2.4, valign: "middle", fontFace: HEAD });
  rect(s, 0.5, 3.85, 12.3, 3.1, WHITE, { line: RED, lw: 2.5, r: 0.04 });
  txt(s, D.TATT.hypothese, { x: 0.8, y: 3.85, w: 11.8, h: 3.1, valign: "middle", fontFace: HEAD });
  s.addNotes("Nur aufdecken, wenn der gemeinsame Satz deutlich abweicht – sonst gilt der Satz der Klasse. „Komplementär“ im Sinne Watzlawicks: Die Rollen ergänzen sich als oben und unten – das heißt nicht, dass das Gespräch gelingt.");
}

// ===== V: Verarbeiten – Ich mache vor, Aufgabe 3 =====
{
  const s = base("Ich mache vor: das offene Tor", { phase: "VERARBEITEN" });
  D.VORMACHEN.forEach(([k, v], i) => {
    const y = 1.2 + i * 1.95;
    rect(s, 0.5, y, 12.3, 1.8, i === 2 ? LIGHT2 : LIGHT);
    txt(s, [{ text: k + "  ", options: { bold: true, color: i === 2 ? RED : NAVY } }, { text: v, options: i === 0 ? { italic: true } : {} }], { x: 0.8, y, w: 11.8, h: 1.8, valign: "middle" });
  });
  s.addNotes("Laut denken, Schritt für Schritt: „Was steht da? → Was fällt auf? → Wofür steht es?“ Alle tragen die erste Zeile der Tabelle mit ein (vgl. " + D.zeile(D.SYMBOLE[0][1]) + "). Tafel ③, erste Zeile.");
}
aufgabe(D.AUFGABE3, "Ansage: „Jetzt arbeitet ihr allein. Die Tabelle steht auf dem Arbeitsblatt, Aufgabe 3.“ Timer 6 Minuten. Herumgehen, Hilfe nur als Frage: „Was fällt an dem Wort auf?“");
{
  const s = base("Lösung: Symbole", { phase: "VERARBEITEN" });
  D.SYMBOLE.slice(1).forEach(([b, anker, sach], i) => {
    const y = 1.15 + i * 2.05;
    rect(s, 0.5, y, 12.3, 1.9, i % 2 ? WHITE : LIGHT, { line: "D5DCEC", lw: 0.75 });
    txt(s, [{ text: b, options: { bold: true, color: RED, breakLine: true } }, { text: "vgl. " + D.zeile(anker), options: { color: MUTED } }], { x: 0.7, y, w: 3.4, h: 1.9, valign: "middle" });
    txt(s, sach, { x: 4.1, y, w: 8.6, h: 1.9, valign: "middle" });
  });
  s.addNotes("Erst eine Antwort vom Platz, dann aufdecken. „Vergleicht und verbessert mit Grün.“ Beim Gesetz ausdrücklich offen lassen: Recht, Wahrheit, Sinn, Gott – oder der Vater. Wichtig ist die Begründung am Text. Tafel ③.");
}

// ===== A: Auswerten – Hypothese prüfen, Zielcheck =====
aufgabe(D.HYPOTHESE, "Eine Minute allein suchen, dann Stichproben. Belege nacheinander aufdecken (nächste Folie). Tafel ④.");
{
  const s = base("Belege für die Asymmetrie", { phase: "AUSWERTEN" });
  D.HYPOTHESE.belege.forEach(([ebene, beleg, anker, erkl], i) => {
    const y = 1.15 + i * 1.62;
    rect(s, 0.5, y, 12.3, 1.5, i % 2 ? WHITE : LIGHT, { line: "D5DCEC", lw: 0.75 });
    txt(s, [{ text: ebene, options: { bold: true, color: RED, breakLine: true } }, { text: "vgl. " + D.zeile(anker), options: { color: MUTED } }], { x: 0.7, y, w: 3.0, h: 1.5, valign: "middle" });
    txt(s, erkl, { x: 3.8, y, w: 8.9, h: 1.5, valign: "middle" });
  });
  rect(s, 0.5, 6.05, 12.3, 1.25, WHITE, { line: GOLD, lw: 2.5, r: 0.04 });
  txt(s, D.HYPOTHESE.fazit, { x: 0.8, y: 6.05, w: 11.8, h: 1.25, valign: "middle", fontFace: HEAD });
  s.addNotes("Belege im Wortlaut: " + D.HYPOTHESE.belege.map(([e, b]) => e + ": " + b).join(" – ") + " Rückgriff auf das Zitat vom Anfang: Der Mann nimmt die untere Rolle an – Ohnmacht wird zur Resignation.");
}
{
  const s = base("Zielcheck", { time: "4 Min." });
  txt(s, "Allein, im Heft.", { x: 0.5, y: 1.15, w: 12.3, h: 0.6, italic: true, color: MUTED });
  rect(s, 0.5, 1.9, 12.3, 1.9, LIGHT);
  txt(s, D.ZIELCHECK.auftrag, { x: 0.8, y: 1.9, w: 11.8, h: 1.9, valign: "middle", bold: true });
  txt(s, [{ text: "Satzgerüst: ", options: { bold: true, color: BLUE } }, { text: D.ZIELCHECK.geruest, options: { italic: true } }], { x: 0.8, y: 4.1, w: 11.8, h: 1.3 });
  txt(s, "Fertig? Stift hinlegen.", { x: 0.5, y: 6.8, w: 12.3, h: 0.5, bold: true, color: GREEN, align: "right" });
  s.addNotes("Auswerten wird nie gestrichen. Herumgehen. Achtung Komma vor „weil“. Danach Musterlösung und Kriterien, mit Grün kontrollieren. Wer noch keinen Baustein von „kafkaesk“ nennt, wird in der Wiederholung der nächsten Stunde gezielt aufgegriffen.");
}
{
  const s = base("Zielcheck: Musterlösung");
  rect(s, 0.5, 1.2, 12.3, 3.5, LIGHT);
  txt(s, D.ZIELCHECK.muster, { x: 0.8, y: 1.2, w: 11.8, h: 3.5, valign: "middle" });
  txt(s, [{ text: "Prüft mit Grün: ", options: { bold: true, color: GREEN, breakLine: true } }, ...D.ZIELCHECK.kriterien.map((k, i) => ({ text: "✓ " + k, options: { breakLine: i < 2 } }))], { x: 0.8, y: 4.9, w: 11.8, h: 2.3 });
  s.addNotes("Jedes Symbol ist richtig, wenn Zeile, Deutung und ein Baustein oder Gefühl von „kafkaesk“ stimmen. Rückblick auf die Lernziele: Was davon kann ich jetzt?");
}

pres.writeFile({ fileName: "Praesentation_Vor_dem_Gesetz.pptx" });
