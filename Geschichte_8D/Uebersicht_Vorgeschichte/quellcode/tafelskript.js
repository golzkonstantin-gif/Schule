const fs = require("fs");
const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, BorderStyle, ShadingType, AlignmentType, VerticalAlign, Footer, PageNumber } = require("docx");
const E = require("./daten.js");
const NAVY = "1E2761", MUTED = "5B6B8C", LIGHT = "F4F7FD", RED = "D9534F", GREEN = "1B5E3A", W = 10206;
const t = (x, o = {}) => new TextRun({ text: x, font: o.font || "Calibri", size: o.size || 21, bold: o.bold, italics: o.italics, color: o.color || "000000" });
const p = (r, o = {}) => new Paragraph({ children: Array.isArray(r) ? r : [r], spacing: { before: o.before ?? 0, after: o.after ?? 80 }, keepNext: o.keepNext, indent: o.indent });
const sd = (c) => ({ style: BorderStyle.SINGLE, size: 6, color: c });
const bx = (c) => ({ top: sd(c), bottom: sd(c), left: sd(c), right: sd(c) });
const cell = (ch, w, fill) => new TableCell({ children: Array.isArray(ch) ? ch : [ch], width: { size: w, type: WidthType.DXA }, shading: fill ? { fill, type: ShadingType.CLEAR, color: "auto" } : undefined, borders: bx("B8C2DB"), margins: { top: 60, bottom: 60, left: 100, right: 100 }, verticalAlign: VerticalAlign.TOP });
const table = (ws, rows) => new Table({ width: { size: ws.reduce((a, b) => a + b, 0), type: WidthType.DXA }, columnWidths: ws, rows: rows.map((r) => new TableRow({ children: r.map((x, i) => { const [c, f] = Array.isArray(x) ? x : [x]; return cell(typeof c === "string" ? p(t(c, { size: 20, bold: f === "h" }), { after: 0 }) : c, ws[i], f === "h" ? "EAF0FA" : f); }), cantSplit: true })) });
const h1 = (s) => p(t(s, { font: "Cambria", size: 36, bold: true, color: NAVY }), { after: 60 });
const h3 = (s) => p(t(s, { font: "Cambria", size: 24, bold: true, color: NAVY }), { before: 160, after: 60, keepNext: true });
const kv = (k, v) => p([t(k + " ", { bold: true, color: NAVY }), t(v)], { after: 50 });
const say = (s) => p([t("Ansage: ", { bold: true, color: RED, size: 20 }), t("„" + s + "“", { italics: true })], { after: 50 });
const dient = (s) => p([t("Dient Ziel: ", { bold: true, color: GREEN, size: 19 }), t(s, { color: GREEN, size: 19 })], { after: 50 });
const li = (s) => p(t("•  " + s, { size: 21 }), { indent: { left: 280 }, after: 40 });
const footer = new Footer({ children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ font: "Calibri", size: 16, color: MUTED, children: ["Tafelskript Übersicht Vorgeschichte · 8D · Seite ", PageNumber.CURRENT] })] })] });
const d = [];
d.push(p(t("GESCHICHTE · KLASSE 8D · FRANZÖSISCHE REVOLUTION", { size: 18, bold: true, color: MUTED }), { after: 20 }), h1("Übersicht: Der Weg zur Menschenrechtserklärung"));
d.push(p(t("Baustein für den Anfang der Stunde 5 (Wiederholung der Stunden 1 bis 4, etwa 8 Minuten). Er ersetzt die übliche Wiederholung und bleibt Lernhilfe für Test 1 (kein Hilfsmittel im Test).", { size: 21 }), { after: 100 }));
d.push(table([2600, 7606], [
  [["Ziel (Ich kann …)", "h"], "… die Ereignisse von 1715 bis zur Menschenrechtserklärung in die richtige Reihenfolge bringen und mit „weil“ oder „deshalb“ erklären, wie sie zusammenhängen."],
  [["Prüfaufgabe: Woran erkenne ich es?", "h"], "Die Reihenfolge stimmt nach der Kontrolle (Blatt, Stichprobe im Blatt). Zwei Sätze im Heft: „Klerus und Adel gaben nach, weil …“ und „Die Bauern griffen Schlösser an, deshalb …“."],
  [["Beitrag zur Reihe", "h"], "Die Übersicht kontextualisiert die Vorgeschichte (Absolutismus, Schulden, Generalstände, Nationalversammlung, Ballhaus, Bastille, 4. und 26. August) und wird zur Grundlage für Test 1 und das Gesamturteil in Stunde 11."],
  [["Mindestziel", "h"], "Reihenfolge richtig (Felder 1 bis 14) und ein Pfeil mit „deshalb“. Die zwei Sätze im Heft laufen im Zweifel mündlich (*)."],
  [["Vorgehen", "h"], "Direktes Vorgehen, weil das Vorwissen geordnet werden muss (Wiederholung als Aktivierung: allein, zu zweit, Plenum). Kontrolle gemeinsam, erst danach kleben."],
]));
d.push(h3("Zeitplan (8 Minuten, AVIVA: Vorwissen aktivieren)"));
d.push(table([900, 6000, 3306], [
  [["Min.", "h"], ["Schritt", "h"], ["Arbeitsform", "h"]],
  ["3", "Ordnen: 14 Schnipsel legen (Folie 1)", "zu zweit, Flüsterstufe"],
  ["2,5", "Kontrolle gemeinsam (Folien 2 bis 4, Verbessern mit Grün)", "Plenum, Stichproben"],
  ["1", "Aufkleben", "allein"],
  ["1,5", "drei „deshalb“-Pfeile (Folie 5), Sätze im Heft (Folie 6)", "Plenum, allein"],
]));
d.push(p(t("Sollbruchstellen: Satz im Heft nur mündlich; Pfeile nur 11 zu 13 (der wichtigste Übergang). Die Phase „Auswerten“ der Stunde 5 wird nie gestrichen; hier wird nur gekürzt.", { size: 20, color: MUTED }), { before: 60 }));
d.push(h3("Ablauf mit Ansagen"));
d.push(p(t("1  Ordnen (Folie 1, 3 Minuten)", { bold: true, color: NAVY, size: 23 }), { keepNext: true }), say("Stifte liegen. Ihr hört nur zu. Jetzt arbeitet ihr zu zweit. Flüsterstufe. Nur dein Partner hört dich. Die Aufgabe steht an der Tafel."), dient("Reihenfolge (Ziel, erster Teil)."), li("Jede Person hat Blatt und Schnipsel (14 Streifen, gemischt). Schnipsel legen, nicht kleben."), li("Durch die Reihen gehen, nichts verraten. Wer fertig ist: „Stift hinlegen.“"), li("Warnung vor dem Fehler: 1787 (Steuerplan) kommt nach dem Krieg in Amerika (1778 bis 1783); 17. Juni vor 20. Juni."));
d.push(p(t("2  Kontrolle (Folien 2 bis 4, 2,5 Minuten)", { bold: true, color: NAVY, size: 23 }), { keepNext: true, before: 100 }), say("Stifte liegen. Ihr hört nur zu. Vergleicht und verbessert mit Grün."), dient("Reihenfolge, Zusammenhang."), li("Felder 1 bis 5, 6 bis 10, 11 bis 14 aufdecken. Stichprobe bei 2 bis 3 Personen pro Folie: „Warum steht es vor dem nächsten?“ (kurze Antwort, unbenotet)."), li("Wer alles richtig hatte, hebt nichts hoch. Verbessert wird nur am Blatt, noch nicht geklebt."));
d.push(p(t("3  Kleben (1 Minute)", { bold: true, color: NAVY, size: 23 }), { keepNext: true, before: 100 }), say("Jetzt klebt ihr auf. Stifte liegen erst danach."), dient("Übersicht entsteht (Lernhilfe für Test 1)."), li("Kleber bereit. Falsch Geklebtes entsteht so nicht."));
d.push(p(t("4  „deshalb“ und Sätze (Folien 5 und 6, 1,5 Minuten)", { bold: true, color: NAVY, size: 23 }), { keepNext: true, before: 100 }), say("Jetzt zeichnen wir drei Pfeile ein. … Jetzt arbeitet ihr allein. Die Aufgabe steht an der Tafel."), dient("Zusammenhang erklären (Ziel, zweiter Teil)."), li("Pfeile gemeinsam: 3 zu 5 (Geld fehlt, deshalb Generalstände), 8 zu 9 (Versammlung nicht anerkannt, deshalb Saal gesperrt, deshalb Ballhausschwur), 11 zu 13 (Bastille und Aufstände, deshalb geben Klerus und Adel nach)."), li("Zwei Sätze im Heft. Erwartung (Auswahl): „Klerus und Adel gaben nach, weil die Bauern Schlösser angriffen und der König nach der Bastille nicht mehr alles kontrollierte.“ „Die Bauern griffen Schlösser an, deshalb mussten Klerus und Adel auf ihre Vorrechte verzichten.“"));
d.push(h3("Lösung (Lösungsblatt: Loesung_Uebersicht_Vorgeschichte_8D)"));
d.push(table([700, 2300, 7206], [[["Nr.", "h"], ["Datum", "h"], ["Ereignis", "h"]], ...E.map((e, i) => [String(i + 1), e[0], e[1]])]));
d.push(h3("Zum Prüfen (Vereinfachungen und offene Punkte)"));
d.push(li("1715 (Tod Ludwigs XIV.) ist ein Vorschlag als Start; ist er im Unterricht behandelt? Sonst als Feld streichen und bei 1778 beginnen (dann 13 Felder)."));
d.push(li("Daten aus den Texten der Stunden 1 bis 4 (8C und 8D gleich): 1778 bis 1783, 1787, 1788, August 1788, 5. Mai, 17. und 20. Juni, Anfang Juli, 14. Juli, 4. und 26. August. „Sommer 1789“ bleibt ohne Tag (Aufstände im ganzen Land)."));
d.push(li("Ballhausschwur: Der König lässt den Saal sperren, die Abgeordneten schwören im Ballhaus, nicht auseinanderzugehen, bis es eine Verfassung gibt (wie im Partnerpuzzle, Text A). Bitte gegenlesen."));
d.push(li("Vereinfachung: Die Ereignisse stehen als einfache Reihe. In Wirklichkeit liefen Aufstände, Versammlung und Hof zugleich ab; die Pfeile zeigen nur die wichtigsten Zusammenhänge."));
d.push(h3("Material und Vorbereitung (Materialliste)"));
d.push(table([4400, 5806], [
  [["Material", "h"], ["Menge und Hinweis", "h"]],
  ["Uebersicht_Blatt_Vorgeschichte_8D", "26-mal, A4 quer, einseitig (24 Schüler plus Reserve)"],
  ["Uebersicht_Schnipsel_Vorgeschichte_8D", "26-mal, A4 quer. Stapel mit der Schneidemaschine schneiden (6 Schnitte quer, 1 Schnitt längs), 14 Streifen pro Person in einen Umschlag oder Klemmbeutel. Vorab erledigen."],
  ["Praesentation_Uebersicht_Vorgeschichte_8D", "6 Folien (27 Punkt), Notizen mit Ansagen und Lösung; die Folien kommen an den Anfang von Stunde 5"],
  ["Kleber, grüne Stifte", "pro Person; Kleber auf dem Pult als Reserve"],
  ["Loesung_Uebersicht_Vorgeschichte_8D", "für die Lehrkraft"],
]));
d.push(h3("So sieht das Blatt am Ende aus"));
d.push(p(t("14 Felder in der richtigen Reihenfolge beklebt, falsche Reihen mit Grün verbessert, drei „deshalb“-Pfeile eingezeichnet, im Heft zwei Sätze zur Prüfaufgabe.", { size: 21 })));
d.push(h3("Nach der Stunde festhalten (Reflexionslog)"));
d.push(li("Wie lange brauchte die Klasse zum Ordnen? Wo gab es Verwechslungen (1787 und 1778)? Wie lief die gemeinsame Kontrolle (Lautstärke, Stichproben)?"));
d.push(li("Gelang der Übergang Bastille zu 4. August? Wie viel Zeit blieb für die Verfassung (Stunde 5)?"));
(async () => {
  const doc = new Document({ styles: { default: { document: { run: { font: "Calibri", size: 21 } } } }, sections: [{ properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 1000, bottom: 850, left: 850, right: 850 } } }, footers: { default: footer }, children: d }] });
  fs.writeFileSync("../Tafelskript_Uebersicht_Vorgeschichte_8D.docx", await Packer.toBuffer(doc)); console.log("ok");
})();
