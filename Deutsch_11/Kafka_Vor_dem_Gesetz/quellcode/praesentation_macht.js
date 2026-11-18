// Präsentation „Vor dem Gesetz“ – Macht und Ohnmacht (Neustart, 45 Min., 16:9, durchgehend 27 pt)
const pptxgen = require("pptxgenjs");
const D = require("./daten_macht");

const NAVY = "1E2761", MUTED = "5B6B8C", LIGHT = "F4F7FD", LIGHT2 = "EAF0FA", ICE = "CADCFC", WHITE = "FFFFFF";
const RED = "D9534F", GOLD = "E8A33D", GREEN = "2E9E6B", BLUE = "3F7CC4", T_FILL = "FBD9D7", M_FILL = "D6E6F8";
const HEAD = "Cambria", BODY = "Calibri", FS = 27;

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE";
pres.title = "Kafka – Vor dem Gesetz: Macht und Ohnmacht";

const txt = (s, text, o) => s.addText(text, Object.assign({ isTextBox: true, fontFace: BODY, fontSize: FS, color: NAVY, margin: 0, valign: "top" }, o));
const rect = (s, x, y, w, h, fill, o = {}) => s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, fill: { color: fill }, line: o.line ? { color: o.line, width: o.lw || 1.5 } : { type: "none" }, rectRadius: o.r ?? 0.08 });
function base(title, opts = {}) {
  const s = pres.addSlide();
  s.background = { color: WHITE };
  rect(s, 0, 0, 13.333, 1.0, NAVY, { r: 0 });
  txt(s, title, { x: 0.5, y: 0, w: 9.8, h: 1.0, fontFace: HEAD, fontSize: 30, bold: true, color: WHITE, valign: "middle" });
  if (opts.time) { rect(s, 10.9, 0.15, 2.1, 0.7, RED, { r: 0.12 }); txt(s, opts.time, { x: 10.9, y: 0.15, w: 2.1, h: 0.7, bold: true, color: WHITE, align: "center", valign: "middle" }); }
  return s;
}
const bullets = (arr, o = {}) => arr.map((x, i) => ({ text: x, options: Object.assign({ bullet: o.num ? { type: "number" } : { indent: 24 }, breakLine: i < arr.length - 1, paraSpaceAfter: o.space ?? 12 }, o.opt || {}) }));
// Aufgabenfolie: oben die Aufgabe (Nummer + Titel, Schritte, Zeit, Arbeitsform), unten immer ein Muster
function aufgabe(a, zeit, form, notes, muster) {
  const s = base(a.titel, { time: zeit });
  txt(s, form, { x: 0.5, y: 1.1, w: 12.3, h: 0.5, italic: true, color: MUTED });
  const start = a.start || 1;
  txt(s, a.schritte.map((x, i) => ({ text: (start + i) + ". " + x, options: { breakLine: i < a.schritte.length - 1, paraSpaceAfter: 6, indentLevel: 0 } })), { x: 0.6, y: 1.6, w: 12.1, h: muster.top - 1.7 });
  const y = muster.top, h = 6.75 - muster.top;
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y, w: 12.3, h, fill: { color: "FFF8EC" }, line: { color: GOLD, width: 2 } });
  txt(s, "Muster", { x: 0.7, y: y + 0.08, w: 3, h: 0.5, bold: true, color: GOLD });
  txt(s, muster.runs, { x: 0.8, y: y + 0.55, w: 11.8, h: h - 0.65, valign: "top", paraSpaceAfter: 4 });
  txt(s, "Fertig? Stift hinlegen.", { x: 0.5, y: 6.85, w: 12.3, h: 0.5, bold: true, color: GREEN, align: "right" });
  if (notes) s.addNotes(notes);
  return s;
}
const zeilen = (rows, notes, title, o = {}) => {
  const s = base(title);
  const h = o.h || (5.9 / rows.length) - 0.12;
  rows.forEach(([k, v], i) => {
    const y = 1.15 + i * (h + 0.12);
    rect(s, 0.5, y, 12.3, h, i % 2 ? WHITE : LIGHT, { line: "D5DCEC", lw: 0.75 });
    txt(s, k, { x: 0.75, y, w: o.kw || 3.2, h, bold: true, color: RED, valign: "middle" });
    txt(s, v, { x: 0.9 + (o.kw || 3.2), y, w: 11.6 - (o.kw || 3.2), h, valign: "middle" });
  });
  if (notes) s.addNotes(notes);
  return s;
};

// ===== Verlauf und Ziel =====
{
  const s = base("Heute");
  D.VERLAUF.forEach(([inhalt], i) => {
    const y = 1.2 + i * 0.97;
    rect(s, 0.5, y, 12.3, 0.85, i % 2 ? WHITE : LIGHT, { line: "D5DCEC", lw: 0.75 });
    txt(s, String(i + 1), { x: 0.75, y, w: 0.6, h: 0.85, bold: true, color: RED, valign: "middle" });
    txt(s, inhalt, { x: 1.4, y, w: 11.2, h: 0.85, valign: "middle" });
  });
  s.addNotes("Arbeitsblatt austeilen. Ziel und Ablauf auch an der Seitentafel. Ansage: „Ihr kennt den Text schon. Heute geht es direkt hinein.“");
}
{
  const s = base("Ziel");
  rect(s, 0.5, 1.4, 12.3, 1.9, LIGHT2);
  txt(s, D.ZIEL, { x: 0.8, y: 1.4, w: 11.8, h: 1.9, valign: "middle", bold: true });
  txt(s, [{ text: "Leitfrage: ", options: { bold: true, color: RED } }, { text: D.LEITFRAGE }], { x: 0.8, y: 3.7, w: 11.8, h: 1.2 });
  s.addNotes("Ziel vorlesen. Am Ende (Zielcheck) darauf zurückkommen.");
}

// ===== Aufgabe 1: Markieren =====
aufgabe(D.AUFGABE1, "5 Min.", "allein, auf dem Arbeitsblatt", "Ansage: „Jetzt arbeitet ihr allein. Zwei Farben: eine für den Türhüter, eine für den Mann. Einkreisen, wie sie es sagen.“ 5 Minuten, herumgehen. Typische Frage: Zählt indirekte Rede? – Ja („fragt, ob …“). Das Muster ist der erste Satz des Türhüters (Z. 2 f.) – so sieht jede weitere Stelle aus.", {
  top: 3.65,
  runs: [
    { text: "Aber der Türhüter ", options: { fontFace: HEAD } },
    { text: "sagt", options: { fontFace: HEAD, bold: true, color: RED, underline: { style: "heavy", color: RED } } },
    { text: ", dass ", options: { fontFace: HEAD } },
    { text: "er ihm jetzt den Eintritt nicht gewähren könne", options: { fontFace: HEAD, highlight: "FBD9D7", breakLine: true } },
    { text: "markiert = was er sagt (Farbe Türhüter)   ·   ", options: { color: MUTED } },
    { text: "eingekreist", options: { bold: true, color: RED } },
    { text: " = wie er es sagt", options: { color: MUTED } },
  ],
});
{
  const s = base("Abgleich: Wie sprechen sie?");
  [["Türhüter", D.EINGEKREIST.tuerhueter, T_FILL, RED], ["Mann vom Lande", D.EINGEKREIST.mann, M_FILL, BLUE]].forEach(([k, arr, fill, col], i) => {
    const x = 0.5 + i * 6.25;
    rect(s, x, 1.15, 6.05, 6.1, fill);
    txt(s, k, { x: x + 0.25, y: 1.25, w: 5.6, h: 0.6, bold: true, color: col });
    txt(s, arr.join("\n"), { x: x + 0.25, y: 1.9, w: 5.6, h: 5.3, paraSpaceAfter: 4 });
  });
  s.addNotes("„Vergleicht und verbessert mit Grün.“ Das sind die eingekreisten Wörter (wie). Der vollständig markierte Text steht auf dem Lösungsblatt – bei Bedarf unter die Dokumentenkamera legen.");
}
{
  const s = base("Was fällt euch auf?");
  D.AUSWERTUNG.forEach(([q], i) => {
    const y = 1.3 + i * 1.95;
    rect(s, 0.5, y, 12.3, 1.75, i % 2 ? WHITE : LIGHT, { line: "D5DCEC", lw: 0.75 });
    txt(s, (i + 1) + "  " + q, { x: 0.8, y, w: 11.8, h: 1.75, valign: "middle" });
  });
  s.addNotes("Gestuft vom Sehen zum Deuten, je 2–3 Stichproben. Erwartung: " + D.AUSWERTUNG.map(([, a], i) => (i + 1) + ") " + a).join(" ") + " Zusatz bei Zeit: " + D.AUSWERTUNG_ZUSATZ.join(" – ") + " Antworten in zwei Spalten an die Tafel (①), die Überschriften Macht / Ohnmacht kommen zuletzt – aus der Klasse.");
}
{
  const s = base("Ergebnis");
  D.TAFEL_SPALTEN.forEach(([k, wie, form, mo], i) => {
    const x = 0.5 + i * 6.25, fill = i ? M_FILL : T_FILL;
    rect(s, x, 1.15, 6.05, 4.4, fill);
    txt(s, [{ text: k, options: { bold: true, breakLine: true } }, { text: wie, options: { breakLine: true } }, { text: form, options: { italic: true, color: MUTED } }], { x: x + 0.25, y: 1.25, w: 5.6, h: 4.2, paraSpaceAfter: 10 });
    rect(s, x, 5.75, 6.05, 1.35, WHITE, { line: i ? BLUE : RED, lw: 2.5 });
    txt(s, "→ " + mo, { x, y: 5.75, w: 6.05, h: 1.35, bold: true, color: i ? BLUE : RED, align: "center", valign: "middle", fontSize: 32 });
  });
  s.addNotes("Erst aufdecken, wenn die Klasse selbst auf Macht und Ohnmacht gekommen ist. Das Tafelbild bleibt die ganze Stunde stehen – bei jedem weiteren Schritt darauf zeigen.");
}

// ===== Aufgabe 2: Einleitung =====
const A2_NOTES = "Schritt für Schritt: 1. TATTE-Bausteine gemeinsam (2 Min.). 2. TATTE-Satz gemeinsam (2 Min.). 3. Inhalt allein (3 Min.): „Erzählt den Ablauf anhand eurer Farben.“ 1–3 Sätze, ein Satz genügt; dann 2 Stichproben. 4. Deutungshypothese gemeinsam: Die Antwort auf „Was wird thematisiert?“ steht schon an der Tafel (Macht/Ohnmacht). Lehrkraft beginnt den Satz, die Klasse vervollständigt. Die Muster zeigen nur die Satzgerüste.";
aufgabe({ titel: D.AUFGABE2.titel + " (1)", schritte: D.AUFGABE2.schritte.slice(0, 2), start: 1 }, "4 Min.", "gemeinsam", A2_NOTES, {
  top: 3.3,
  runs: [
    { text: "TATTE: ", options: { bold: true } }, { text: "Textsorte · Autor · Titel · Thema · Erscheinungsjahr", options: { breakLine: true } },
    { text: "TATTE-Satz: ", options: { bold: true } }, { text: "In der Parabel „…“ von …, die … veröffentlicht wurde, geht es um …", options: { italic: true } },
  ],
});
aufgabe({ titel: D.AUFGABE2.titel + " (2)", schritte: D.AUFGABE2.schritte.slice(2), start: 3 }, "6 Min.", "Inhalt allein · Hypothese gemeinsam", A2_NOTES, {
  top: 3.6,
  runs: [
    { text: "Inhalt: ", options: { bold: true } }, { text: "Ein … will …, doch … . Am Ende …", options: { italic: true, breakLine: true } },
    { text: "Hypothese: ", options: { bold: true } }, { text: "Es lässt sich vermuten, dass Kafka … thematisiert.", options: { italic: true } },
  ],
});

zeilen(D.TATTE.map(([k, v]) => [k, v]), "TATTE: Textsorte, Autor, Titel, Thema, Erscheinungsjahr. Das Thema bewusst allgemein – Macht und Ohnmacht kommen erst in die Hypothese.", "Einleitung: TATTE", { kw: 3.6 });
{
  const s = base("Einleitung: TATTE-Satz");
  txt(s, [{ text: "Gerüst: ", options: { bold: true, color: BLUE } }, { text: "In der Parabel „…“ von …, die … veröffentlicht wurde, geht es um …", options: { italic: true } }], { x: 0.8, y: 1.3, w: 11.8, h: 1.3 });
  rect(s, 0.5, 2.9, 12.3, 3.2, LIGHT);
  txt(s, D.EINLEITUNG.satz1, { x: 0.8, y: 2.9, w: 11.8, h: 3.2, valign: "middle", fontFace: HEAD });
  s.addNotes("Aus den Bausteinen gemeinsam einen Satz bauen: Lehrkraft beginnt („In der Parabel …“), die Klasse ergänzt Autor, Jahr und Thema. Lösung erst danach aufdecken. Entstanden 1914, veröffentlicht 1915 – „veröffentlicht“ ist eindeutig.");
}
{
  const s = base("Einleitung: Inhalt");
  rect(s, 0.5, 1.4, 12.3, 3.0, LIGHT);
  txt(s, D.EINLEITUNG.inhalt, { x: 0.8, y: 1.4, w: 11.8, h: 3.0, valign: "middle", fontFace: HEAD });
  txt(s, "Ein Satz genügt – im Präsens.", { x: 0.8, y: 4.8, w: 11.8, h: 0.8, italic: true, color: MUTED });
  s.addNotes("Erst zwei Schülersätze hören, dann aufdecken. Andere richtige Fassungen gelten.");
}
{
  const s = base("Einleitung: Deutungshypothese");
  txt(s, [{ text: "Was will der Mann?", options: { bold: true, breakLine: true } }, { text: "Was wird dabei thematisiert?", options: { bold: true } }], { x: 0.8, y: 1.3, w: 11.8, h: 1.4 });
  rect(s, 0.5, 3.0, 12.3, 3.6, WHITE, { line: RED, lw: 2.5, r: 0.04 });
  txt(s, D.EINLEITUNG.hypothese, { x: 0.8, y: 3.0, w: 11.8, h: 3.6, valign: "middle", fontFace: HEAD });
  s.addNotes("Erst die Leitfragen stellen (Folie bis zur Hypothese verdecken oder Hypothese als zweite Animation). Hinweis Klausur: Eine Hypothese ist vorläufig – „Es lässt sich vermuten, dass …“. Wir prüfen sie heute mit Bild- und Sachebene und mit den Stilmitteln.");
}

// ===== Spiel =====
{
  const s = base("Die Szene spielen", { time: "5 Min." });
  txt(s, D.SPIEL_REGIE, { x: 0.8, y: 1.25, w: 11.8, h: 1.2, color: MUTED, italic: true });
  txt(s, "Alle anderen achten auf:", { x: 0.8, y: 2.55, w: 11.8, h: 0.6, bold: true });
  txt(s, bullets(D.SPIEL_FRAGEN.map(([k]) => k)), { x: 0.9, y: 3.2, w: 11.6, h: 3.8 });
  s.addNotes("Zwei Freiwillige mit der Spielvorlage. Stuhl als Schemel, „seitwärts“ vom Pult (Tor). Spieldauer ca. 2 Minuten. Danach mündlich, je Merkmal eine Stichprobe: " + D.SPIEL_FRAGEN.map(([k, v]) => k + ": " + v).join(" – ") + " Danach die Fußball-Szene mündlich erzählen (Tafelskript). Brücke: „Wofür stehen Türhüter und Mann also?“");
}

// ===== Aufgabe 3: Bild- und Sachebene =====
{
  const s = base("Bildebene – Sachebene – Symbol");
  D.MERKSATZ.forEach(([k, v], i) => {
    const y = 1.3 + i * 1.95;
    s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y, w: 12.3, h: 1.75, fill: { color: WHITE }, line: { color: RED, width: 2.5 } });
    txt(s, [{ text: k + ": ", options: { bold: true, color: RED } }, { text: v }], { x: 0.8, y, w: 11.8, h: 1.75, valign: "middle" });
  });
  s.addNotes("Figuren sind keine Symbole: Auf der Bildebene haben sie eine Rolle (Wächter, Bittsteller), auf der Sachebene stehen sie für etwas. Beide sind namenlos – nur über Funktion und Herkunft bestimmt: typisierte Stellvertreter. Symbole sind nur Gegenstände.");
}
aufgabe(D.AUFGABE3, "4 Min.", "allein, auf dem Arbeitsblatt", "Ansage: „Jetzt arbeitet ihr allein. Die Bildebene steht schon da – ihr ergänzt die Sachebene: Wofür steht es?“ 4 Minuten. Hilfe: auf die Tafel zeigen (Macht/Ohnmacht). Das Muster stammt bewusst aus „Gibs auf!“, damit die Lösung nicht vorweggenommen wird.", {
  top: 3.3,
  runs: [
    { text: "Aus „Gibs auf!“: der Schutzmann", options: { bold: true, color: RED, breakLine: true } },
    { text: "Bildebene: ", options: { bold: true } }, { text: "ein Polizist, den der Erzähler nach dem Weg fragt", options: { breakLine: true } },
    { text: "Sachebene: ", options: { bold: true } }, { text: "eine Instanz, die Orientierung geben sollte – und sie verweigert" },
  ],
});
[[0, 2], [2, 4]].forEach(([a, b], k) => {
  const s = base("Lösung: Bild- und Sachebene (" + (k + 1) + ")");
  D.BILD_SACH.slice(a, b).forEach(([el, bild, sach, anker], j) => {
    const y = 1.15 + j * 3.05;
    rect(s, 0.5, y, 12.3, 2.9, j ? WHITE : LIGHT, { line: "D5DCEC", lw: 0.75 });
    txt(s, [
      { text: el + "  ", options: { bold: true, color: RED } }, { text: "(vgl. " + D.zeile(anker) + ")", options: { color: MUTED, breakLine: true } },
      { text: "Bildebene: ", options: { bold: true } }, { text: bild, options: { breakLine: true } },
      { text: "Sachebene: ", options: { bold: true } }, { text: sach },
    ], { x: 0.8, y, w: 11.8, h: 2.9, valign: "middle", paraSpaceAfter: 6 });
  });
  s.addNotes("Erst eine Antwort vom Platz, dann aufdecken. „Vergleicht und verbessert mit Grün.“ Beim Gesetz jede am Text begründete Deutung gelten lassen.");
});
{
  const s = base("Stimmt unsere Hypothese?");
  rect(s, 0.5, 1.3, 12.3, 4.5, WHITE, { line: GOLD, lw: 2.5, r: 0.04 });
  txt(s, D.ERGEBNIS, { x: 0.8, y: 1.3, w: 11.8, h: 4.5, valign: "middle", fontFace: HEAD });
  s.addNotes("Ergebnis auf dem Arbeitsblatt notieren lassen (Aufgabe 3, Ergebnis). Zeigen: Das stand schon an der Tafel – Macht und Ohnmacht.");
}

// ===== Aufgabe 4: Stilmittel =====
{
  const s = base("Wiederholung: Stilmittel");
  D.STIL_WIEDERHOLUNG.forEach(([k, sig, bsp], i) => {
    const y = 1.15 + i * 1.53;
    rect(s, 0.5, y, 12.3, 1.42, i % 2 ? WHITE : LIGHT, { line: "D5DCEC", lw: 0.75 });
    txt(s, k, { x: 0.75, y, w: 2.7, h: 1.42, bold: true, color: RED, valign: "middle" });
    txt(s, [{ text: sig, options: { breakLine: true } }, { text: bsp, options: { italic: true, color: BLUE } }], { x: 3.5, y, w: 9.1, h: 1.42, valign: "middle" });
  });
  s.addNotes("Kurz wiederholen, Beispiele bewusst nicht aus dem Text. Kern: Vergleich = mit „wie“; Metapher = Vergleich ohne „wie“ (Probe: ein „wie“ lässt sich einsetzen).");
}
{
  const [sm, stelle, anker, wirkung, bezug] = D.STILMITTEL[0];
  aufgabe(D.AUFGABE4, "4 Min.", "Zeile 1 gemeinsam, dann allein", "Das Muster ist Zeile 1 (Wiederholung): Die Lehrkraft denkt laut – „Was wird wiederholt? → mächtig. Was bewirkt das? → Die Macht wird immer größer. Was heißt das für unsere Hypothese? → Macht.“ Alle schreiben Zeile 1 mit, dann Zeilen 2–4 allein, 4 Minuten.", {
    top: 3.3,
    runs: [
      { text: "1. " + sm + ": ", options: { bold: true, color: RED } }, { text: stelle, options: { italic: true } }, { text: "  (vgl. " + D.zeile(anker) + ")", options: { color: MUTED, breakLine: true } },
      { text: "Wirkung: ", options: { bold: true } }, { text: wirkung, options: { breakLine: true } },
      { text: "Bezug: ", options: { bold: true, color: RED } }, { text: bezug },
    ],
  });
}
[[1, 3], [3, 4]].forEach(([a, b], k) => {
  const s = base("Lösung: Stilmittel (" + (k + 1) + ")");
  D.STILMITTEL.slice(a, b).forEach(([sm, stelle, anker, wirkung, bezug], j) => {
    const h = b - a === 2 ? 2.95 : 5.9, y = 1.15 + j * (h + 0.1);
    rect(s, 0.5, y, 12.3, h, j ? WHITE : LIGHT, { line: "D5DCEC", lw: 0.75 });
    txt(s, [
      { text: sm + ": ", options: { bold: true, color: RED } }, { text: stelle + " ", options: { italic: true } }, { text: "(vgl. " + D.zeile(anker) + ")", options: { color: MUTED, breakLine: true } },
      { text: "Wirkung: ", options: { bold: true } }, { text: wirkung, options: { breakLine: true } },
      { text: "Bezug: ", options: { bold: true, color: RED } }, { text: bezug },
    ], { x: 0.8, y, w: 11.8, h, valign: "middle", paraSpaceAfter: 6 });
  });
  s.addNotes("Erst eine Antwort vom Platz, dann aufdecken, mit Grün verbessern." + (k ? " Metapher: Probe mit „wie“ – Das Gesetz ist (wie) ein Gebäude, in das man eintritt." : ""));
});
{
  const s = base("Die Sprache stützt die Hypothese");
  rect(s, 0.5, 1.3, 12.3, 4.6, WHITE, { line: GOLD, lw: 2.5, r: 0.04 });
  txt(s, D.STIL_ERGEBNIS, { x: 0.8, y: 1.3, w: 11.8, h: 4.6, valign: "middle", fontFace: HEAD });
  s.addNotes("Ergebnis auf dem Arbeitsblatt notieren lassen (Aufgabe 4, Ergebnis).");
}

// ===== Zielcheck =====
aufgabe({ titel: "Zielcheck", schritte: [D.ZIELCHECK.auftrag] }, "3 Min.", "allein, im Heft", "Neues Element (nicht besprochen) – prüft, ob Symbol und Bezug zur Hypothese verstanden sind. Wird nie gestrichen. Herumgehen; wer noch Fehler hat, wird in der Wiederholung der nächsten Stunde gezielt aufgegriffen.", {
  top: 3.6,
  runs: [
    { text: "Der Pelzmantel ist ein Symbol für …, denn … (vgl. Z. …).", options: { italic: true, breakLine: true } },
    { text: "Das stützt die Hypothese, weil …", options: { italic: true, breakLine: true } },
    { text: "Tipp: Lies Z. " + parseInt(D.zeile(D.ZIELCHECK.anker).slice(3), 10) + "–" + (parseInt(D.zeile("entschließt er sich, doch lieber").slice(3), 10) + 1) + " noch einmal.", options: { color: BLUE } },
  ],
});
{
  const s = base("Zielcheck: Musterlösung");
  rect(s, 0.5, 1.2, 12.3, 3.6, LIGHT);
  txt(s, D.ZIELCHECK.muster, { x: 0.8, y: 1.2, w: 11.8, h: 3.6, valign: "middle" });
  txt(s, [{ text: "Prüft mit Grün:", options: { bold: true, color: GREEN, breakLine: true } }, ...D.ZIELCHECK.kriterien.map((k, i) => ({ text: "✓ " + k, options: { breakLine: i < 2 } }))], { x: 0.8, y: 5.0, w: 11.8, h: 2.3 });
  s.addNotes("Rückblick auf das Ziel: Einleitung geschrieben? Bild- und Sachebene unterschieden?");
}

pres.writeFile({ fileName: "../Macht_Ohnmacht/Praesentation_Macht_Ohnmacht.pptx" });
