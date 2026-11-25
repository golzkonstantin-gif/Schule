// Folien „Sparta – ein Staat von Kriegern?“ (Einzelstunde: Input 5 – Einzelarbeit 25 – Auswertung 15). AVIVA, alle Texte für die Klasse ≥ 27 pt, Foliennummern automatisch.
const pptxgen = require("pptxgenjs");
const D = require("./data.js");
const OUT = process.argv[2];
const C = { dark: "2B2318", terra: "AD4426", sand: "F4ECDE", gold: "D9A441", muted: "5A5044", soft: "998C77", white: "FFFFFF" };
const pres = new pptxgen(); pres.layout = "LAYOUT_16x9"; pres.title = D.titel;
const H = "Cambria", B = "Calibri", MIN = 27, TITLE = 36, PH = ["A", "V", "I", "V", "A"];
const txt = (s, text, o) => s.addText(text, { isTextBox: true, fontFace: B, color: C.dark, margin: 0, valign: "top", fontSize: MIN, ...o });
const num = s => { s.slideNumber = { x: 9.3, y: 5.3, w: 0.4, h: 0.25, fontFace: B, fontSize: 12, color: C.soft, align: "right" }; };
function aviva(s, k) { PH.forEach((l, i) => s.addText(l, { isTextBox: true, shape: pres.shapes.ROUNDED_RECTANGLE, rectRadius: 0.06, x: 7.15 + i * 0.5, y: 0.12, w: 0.44, h: 0.44,
  fill: { color: i === k ? C.terra : C.sand }, color: i === k ? C.white : C.soft, fontFace: B, fontSize: MIN, bold: true, align: "center", valign: "middle", margin: 0 })); }
const tag = (s, t, fill = C.dark, color = C.white) => s.addText(t, { isTextBox: true, shape: pres.shapes.ROUNDED_RECTANGLE, rectRadius: 0.08, x: 7.15, y: 0.66, w: 2.4, h: 0.5,
  fill: { color: fill }, color, fontFace: B, fontSize: MIN, bold: true, align: "center", valign: "middle", margin: 0 });
const loesung = s => tag(s, "Lösung", C.gold, C.dark);
const title = (s, t, k) => { aviva(s, k); txt(s, t, { x: 0.45, y: 0.15, w: 6.6, h: 1.05, fontSize: TITLE, bold: true, fontFace: H, valign: "middle" }); };
function panel(s, t, x, y, w, h, o = {}) {
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, rectRadius: 0.06, fill: { color: o.fill || C.white }, line: { color: o.line || C.terra, width: o.lw || 2 } });
  txt(s, t, { x: x + 0.2, y, w: w - 0.4, h, color: o.color || C.dark, bold: o.bold, valign: "middle", align: o.align, italic: o.italic });
}
function steps(s, items, y, gap) { items.forEach((t, i) => {
  s.addText(String(i + 1), { isTextBox: true, shape: pres.shapes.OVAL, x: 0.45, y: y + i * gap, w: 0.5, h: 0.5, fill: { color: C.terra }, color: C.white, fontFace: H, fontSize: MIN, bold: true, align: "center", valign: "middle", margin: 0 });
  txt(s, t, { x: 1.1, y: y + i * gap - 0.02, w: 8.5, h: 0.55, valign: "middle" }); }); }
const slide = (bg = C.white) => { const s = pres.addSlide(); s.background = { color: bg }; return s; };

// 1 Titel
{ const s = slide(C.dark);
  txt(s, "GESCHICHTE · KLASSE 11", { x: 0.7, y: 1.0, w: 8.6, h: 0.5, bold: true, color: C.gold, charSpacing: 2 });
  txt(s, D.titel, { x: 0.7, y: 1.55, w: 8.6, h: 1.5, fontSize: 40, bold: true, fontFace: H, color: C.white });
  txt(s, "Gesellschaft · Wirtschaft · Verfassung", { x: 0.7, y: 3.3, w: 8.6, h: 0.6, color: "E8DFD0" });
  s.addNotes("Material: Text_Sparta (A4, 1 pro Schüler), Fragen_Sparta (A4 doppelseitig: vorne Fragentabelle, hinten stummes Schaubild quer), Hilfen_Lehrkraft_Sparta (nur für dich), Antworten_Sparta.\nZeit: Input 5 Min. (Folien 2–5) – Einzelarbeit 25 Min. (Folie 6) – Auswertung 15 Min. (Folien 7–12)."); }

// 2 Ziel und Ablauf
{ const s = slide(); title(s, "Ziel und Ablauf", 0); tag(s, "Ankommen");
  txt(s, "Ich kann Gesellschaft, Wirtschaft und Verfassung Spartas beschreiben – und erklären, warum Sparta ein Kriegerstaat wurde.", { x: 0.45, y: 1.3, w: 9.1, h: 1.35, bold: true, color: C.terra, valign: "middle" });
  steps(s, ["Sparta auf einen Blick", "Text lesen, Fragen beantworten", "Lösungen und Schaubild", "Warum ein Kriegerstaat?"], 2.85, 0.6);
  num(s); s.addNotes("0–1 Min. Ziel vorlesen. Mindestziel: Fragen 1–6 beantwortet. Die Abschlussfrage „Warum ein Kriegerstaat?“ wird nie gestrichen."); }

// 3 Vorwissen
{ const s = slide(); title(s, "„Das ist aber spartanisch!“", 1); tag(s, "Stichprobe");
  txt(s, "Was meint man heute damit?", { x: 0.45, y: 1.5, w: 9.1, h: 0.7, bold: true });
  panel(s, "einfach · ohne Komfort · hart", 0.45, 2.6, 9.1, 0.95, { line: C.soft, color: C.muted, align: "center" });
  panel(s, "Heute: Woher kommt dieses Bild von Sparta?", 0.45, 3.9, 9.1, 0.95, { fill: C.dark, line: C.dark, color: C.gold, bold: true });
  num(s); s.addNotes("1–2 Min. Zwei Schüler kurz fragen, dann das mittlere Feld zeigen (beim Vorführen: Animation oder einfach stehen lassen). Gelenk: „Heute klärt ihr, wie Sparta wirklich aufgebaut war – und warum es so hart wurde.“"); }

// 4 Sparta auf einen Blick
{ const s = slide(); title(s, "Sparta auf einen Blick", 2); tag(s, "zuhören");
  [["Wo?", "Süden der Peloponnes, Landschaft Lakonien"], ["Was?", "Gegenstück zu Athen: kein Volk an der Macht, sondern ein Kriegerstaat"], ["Wichtig", "Sparta eroberte das Nachbarland Messenien"]].forEach(([a, b], i) => {
    const y = 1.35 + i * 1.3;
    txt(s, a, { x: 0.45, y, w: 1.9, h: 1.1, bold: true, color: C.terra, fontFace: H, valign: "middle" });
    panel(s, b, 2.4, y, 7.15, 1.1, { line: C.terra }); });
  num(s); s.addNotes("2–4 Min. Direkt erklären: „Athen kennt ihr. Sparta war die andere Großmacht Griechenlands – und fast überall das Gegenteil. Merkt euch: Sparta hat ein ganzes Nachbarland unterworfen. Das ist der Schlüssel für heute.“"); }

// 5 Drei Gruppen (Begriffe)
{ const s = slide(); title(s, "Drei Gruppen", 2); tag(s, "Begriffe");
  [["Spartiaten", "Vollbürger"], ["Periöken", "freie „Umwohnende“"], ["Heloten", "Unfreie, gehören dem Staat"]].forEach(([a, b], i) => {
    const y = 1.35 + i * 1.3;
    txt(s, a, { x: 0.45, y, w: 3.0, h: 1.1, bold: true, color: C.terra, fontFace: H, valign: "middle" });
    panel(s, b, 3.5, y, 6.05, 1.1, { line: C.terra }); });
  num(s); s.addNotes("4–5 Min. Begriffe vorlesen; sie stehen auch unten auf dem Textblatt. Warnung vor dem Fehler: „Spartaner“ meint alle Bewohner – „Spartiaten“ nur die Vollbürger."); }

// 6 Auftrag Einzelarbeit
{ const s = slide(); title(s, "Text und Fragen", 3); tag(s, "allein · 25 Min.");
  steps(s, ["Lies den Text (Textblatt).", "Beantworte Fragen 1–8 in Stichworten.", "Notiere die Zeilen.", "Fertig? Fülle das Schaubild (Rückseite) aus."], 1.4, 0.8);
  panel(s, "Text und Fragen liegen nebeneinander.", 0.45, 4.6, 9.1, 0.7, { line: C.soft });
  num(s); s.addNotes("5–30 Min. Umhergehen mit dem Hilfeblatt (Fundstelle → Impuls → Satzanfang). Zuerst zu den Schwächeren. Stichproben im Heft/Blatt: Stehen Zeilen dabei? Frage 3 ist der Schlüssel (Heloten in der Überzahl) – dort genau nachlesen lassen.\nNach 20 Min. ansagen: „Noch 5 Minuten.“"); }

// 7–8 Lösungen
const F = D.fragen;
[[0, 4], [4, 8]].forEach(([a, b], k) => {
  const s = slide(); title(s, "Lösungen " + (a + 1) + "–" + b, 4); loesung(s);
  F.slice(a, b).forEach(([q, z, , kurz], i) => {
    const y = 1.3 + i * 1.02;
    s.addText(String(a + i + 1), { isTextBox: true, shape: pres.shapes.OVAL, x: 0.45, y: y + 0.18, w: 0.5, h: 0.5, fill: { color: C.terra }, color: C.white, fontFace: H, fontSize: MIN, bold: true, align: "center", valign: "middle", margin: 0 });
    txt(s, kurz, { x: 1.1, y, w: 8.45, h: 0.9, valign: "middle", bold: true, color: C.terra });
  });
  num(s); s.addNotes("30–35 Min. Vor jeder Lösung kurz Stichprobe: ein Schüler liest seine Antwort vor. Dann vergleichen und ergänzen.\n" + F.slice(a, b).map(([q, z, e], i) => (a + i + 1) + ". " + q + " (" + z + ") – " + e).join("\n")); });

// 9–10 Schaubild
{ const s = slide(); title(s, "Das Schaubild", 4); tag(s, "gemeinsam");
  s.addImage({ path: __dirname + "/schaubild_stumm.png", x: 0.45, y: 1.25, w: 8.2, h: 4.3 });
  num(s); s.addNotes("35–40 Min. Schaubild an die Tafel zeichnen (Tafelskizze in der Handreichung) und gemeinsam ausfüllen: Schüler sagen vom Platz, du trägst ein, alle füllen die Rückseite ihres Fragenblatts. Bei Unklarheiten hier klären."); }
{ const s = slide(); title(s, "Das Schaubild", 4); loesung(s);
  s.addImage({ path: __dirname + "/../Schaubild_Sparta_Loesung.png", x: 0.45, y: 1.25, w: 8.2, h: 4.3 });
  num(s); s.addNotes("Vergleichen. Die Zeile „Darum ein Kriegerstaat“ bleibt für die Abschlussaufgabe zunächst offen – erst nach Folie 11 eintragen."); }

// 11 Abschlussaufgabe
{ const s = slide(); title(s, "Warum ein Kriegerstaat?", 4); tag(s, "allein · 2 Min.");
  panel(s, D.abschluss, 0.45, 1.35, 9.1, 1.25, { line: C.terra, italic: true });
  txt(s, "Antworte ihm in 2–3 Sätzen.", { x: 0.45, y: 2.85, w: 9.1, h: 0.6 });
  panel(s, "„Ich erkläre es dir: …, weil …“", 0.45, 3.7, 9.1, 0.95, { fill: C.dark, line: C.dark, color: C.gold, bold: true, align: "center" });
  num(s); s.addNotes("40–43 Min. Nie streichen. Schriftlich 2 Min., dann 2–3 gezielt aufrufen (ein sicherer, ein unsicherer Schüler).\nMögliche Argumente:\n– " + D.argumente.join("\n– ")); }

// 12 Lösung + Ausblick
{ const s = slide(); title(s, "So kann es lauten", 4); loesung(s);
  txt(s, "Sparta beherrschte viel mehr Heloten, als es Bürger hatte. Weil sie sich immer wieder erhoben, mussten alle Spartiaten ständig kampfbereit sein.", { x: 0.45, y: 1.35, w: 9.1, h: 2.2 });
  panel(s, "Aber: Fast alles wissen wir von Athenern. Nächstes Mal: Athen und Sparta im Vergleich.", 0.45, 3.85, 9.1, 1.3, { fill: C.dark, line: C.dark, color: C.gold, bold: true });
  num(s); s.addNotes("43–45 Min. Zurück zum Ziel: Haben wir erklärt, warum Sparta ein Kriegerstaat wurde? Ausblick auf die Doppelstunde: Quellen über Sparta stammen fast nur von Nicht-Spartanern (Text Z. " + D.zeileVon("Peloponnes. Was wir") + "–40) – darum prüfen wir dort, wem wir trauen können.\nVollständiges Beispiel: " + D.beispiel); }

pres.writeFile({ fileName: OUT }).then(() => console.log("done"));
