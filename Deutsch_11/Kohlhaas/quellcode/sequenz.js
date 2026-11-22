// Sequenzplan „Michael Kohlhaas“ (Kleist), Deutsch 11 – Mo Doppelstunde, Do Einzelstunde; vollständige Lektüre (Rest in den Weihnachtsferien)
const fs = require("fs");
const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, BorderStyle, ShadingType, AlignmentType, VerticalAlign, PageOrientation, Footer, PageNumber } = require("docx");
const NAVY = "1E2761", MUTED = "5B6B8C", LIGHT = "F4F7FD", LIGHT2 = "EAF0FA", RED = "D9534F", GOLD = "E8A33D";
const W = 15338;
const t = (text, o = {}) => new TextRun({ text, font: o.font || "Calibri", size: o.size || 19, bold: o.bold, italics: o.italics, color: o.color || "000000" });
const p = (runs, o = {}) => new Paragraph({ children: Array.isArray(runs) ? runs : [runs], alignment: o.align, keepNext: o.keepNext, pageBreakBefore: o.pb, spacing: { before: o.before ?? 0, after: o.after ?? 40, line: o.line } });
const solid = (c) => ({ style: BorderStyle.SINGLE, size: 6, color: c });
const all = (b) => ({ top: b, bottom: b, left: b, right: b });
const cell = (children, o = {}) => new TableCell({ children: Array.isArray(children) ? children : [children], width: { size: o.w, type: WidthType.DXA }, shading: o.fill ? { fill: o.fill, type: ShadingType.CLEAR, color: "auto" } : undefined, borders: all(solid("B9C6E8")), verticalAlign: o.valign || VerticalAlign.TOP, margins: { top: 50, bottom: 50, left: 100, right: 100 }, columnSpan: o.span });
const hdr = (x, w) => cell(p(t(x, { bold: true, color: "FFFFFF", size: 17 }), { after: 0 }), { w, fill: NAVY });

// [Datum, Stunde, Thema / Leitfrage, Schwerpunkt und Methode, Text (Reclam), Baustein/Material]
// Gerüst: Inhaltsverzeichnis des Unterrichtsmodells der Lehrkraft (Bausteine B1–B8, Arbeitsblätter AB, Zusatzmaterial Z)
const PHASEN = [
  ["Vor den Ferien: Einführung, Kleist und Hintergründe, Beginn der Gewaltspirale (gemeinsam lesen)", [
    ["Mo 30.11.", "DS", "Rechtsgefühl – was tun, wenn das Recht versagt?", "Einstieg: Fehlentscheidung im Sport, gegen die man nichts tun kann; Maßstäbe meines Handelns. Brücke zu Kafka (der Mann vom Lande wartet – und du?). Erwartungen an den Klappentext. Handlungsverlauf von Anfang an dokumentieren (Vorlage). Lesebeginn.", "Anfang", "B1.1–1.3, AB 1"],
    ["Do 3.12.", "ES", "Die Einleitung", "Der erste Absatz: „einer der rechtschaffensten zugleich und entsetzlichsten Menschen seiner Zeit“ – Antithese, Erzählerhaltung; Leitfrage der Einheit an die Tafel.", "erster Absatz", "B2.3"],
    ["Mo 7.12.", "DS", "Kleist und seine Zeit", "Teil 1 Kleist (ca. 35 Min.): Zitat aus der Kant-Krise als Einstieg („Wie kann jemand so etwas sagen?“), Kleist in fünf Stationen, das historische Vorbild Hans Kohlhase (Berlin, Kohlhasenbrück). Teil 2 Hintergründe (ca. 45 Min.): Bürgertum im 18. Jh., Rechts- und Sozialverhältnisse im 16. Jh. und zu Kleists Zeit, Reformen in Preußen – kurz und gezielt.", "–", "B2.1–2.2, AB 2–3, Z3, Z5"],
    ["Do 10.12.", "ES", "Am Schlagbaum", "Wer sagt was – und wie? (markieren, einkreisen) Machtgefälle Kohlhaas – Burgvogt – Junker; der Passschein.", "Schlagbaum", "B3.1"],
    ["Mo 14.12.", "DS", "Kohlhaas prüft die Lage – Kohlhaas als Advokat", "Kritische Prüfung am Schlagbaum; Gespräch Kohlhaas – Herse als Kommunikationsanalyse (Verhör, Glaubwürdigkeit).", "Tronkenburg, Herse", "B3.2–3.3, Z4"],
    ["Do 17.12.", "ES", "Suche nach legalen Lösungen", "Klage, Bittschrift, Lisbeth – Stationen des Rechtswegs; der Wendepunkt.", "Klage bis Lisbeths Tod", "B3.4"],
    ["Mo 21.12.", "DS", "Kohlhaas’ Weg zur Rache – Zwischenbilanz", "Gewaltspirale als Tafelbild; Meinungslinie „rechtschaffen oder entsetzlich?“ mit Beleg. Leseplan für die Ferien vorstellen, nächste Etappe anlesen.", "bisher Gelesenes", "AB 4"],
  ]],
  ["Weihnachtsferien: den Rest der Novelle lesen", [
    ["Ferien", "–", "Ferienlektüre in drei Etappen", "Pro Etappe drei Stichworte und ein Zitat für den Handlungsüberblick (kein Lesetagebuch); Leitfragen je Etappe; Hinweis auf ein Hörbuch.", "Rest der Novelle", "B1.1–1.2"],
  ]],
  ["Nach den Ferien: Selbstjustiz, Individuum und Staat (B4–B5)", [
    ["Mo 4.1.", "DS", "Lektürecheck und Überblick", "Aussagen zum Text prüfen (richtig/falsch, mit Beleg); Handlungsüberblick vervollständigen.", "ganze Novelle", "B1.2"],
    ["Do 7.1.", "ES", "Figurenkonstellation", "Macht – Recht – Familie – Kirche: Wer steht wo? Visualisierung.", "ganze Novelle", "–"],
    ["Mo 11.1.", "DS", "Rachefeldzug", "Selbstjustiz zur Durchsetzung des Rechtsstandpunktes; die Mandate.", "Rachefeldzug, Mandate", "B4.1"],
    ["Do 14.1.", "ES", "Selbsterhöhung", "Biblische Anspielungen („Statthalter Michaels“): Wie stilisiert sich Kohlhaas?", "Mandate", "B4.2, AB 5"],
    ["Mo 18.1.", "DS", "Kohlhaas in Wittenberg", "Luthers Plakat; Gespräch Kohlhaas – Luther: Gesprächsanalyse (Vier-Seiten-Modell, Argumentation).", "Plakat, Gespräch", "B5.1"],
    ["Do 21.1.", "ES", "Die Logik Kohlhaas’ und Luthers Antwort", "Argumente prüfen; „verstoßen“ aus der Gemeinschaft.", "Gespräch", "B5.2, Z7"],
    ["Mo 25.1.", "DS", "Der Einzelne und die Gesellschaft", "Gesellschaftsvertrag in kurzen Auszügen (Locke/Rousseau); Kohlhaas’ Logik: Wer nicht geschützt wird, gehört nicht dazu.", "Gespräch", "B5.3, Z8–Z10"],
    ["Do 28.1.", "ES", "Guter Wille?", "Kant (kurz): Ist Kohlhaas’ Rechtsgefühl moralisch? Rückbezug auf Kleists Kant-Krise.", "–", "AB 6"],
    ["Mo 1.2.", "DS", "Schreibwerkstatt", "Stellungnahme: „Darf man sich sein Recht selbst nehmen?“ – mit inhaltsunabhängigen Satzmustern.", "–", "–"],
    ["Do 4.2.", "ES", "Rückmeldung", "Rückmeldung zur Stellungnahme; Checkliste.", "–", "–"],
  ]],
  ["Ereignisse in Sachsen und Brandenburg, Novelle, Rezeption (B6–B8)", [
    ["Mo 8.2.", "DS", "Ereignisse in Sachsen", "Staatsratssitzung: Politik und Kalkül statt Recht.", "Staatsrat", "B6.1"],
    ["Do 11.2.", "ES", "Die Abdeckerszene", "Ehre, Volk, Absurdität – die Rappen auf dem Markt.", "Abdecker", "B6.2"],
    ["Mo 15.2.", "DS", "Nagelschmidt", "Bruch der Amnestie, die Falle: Der Staat bricht selbst das Recht.", "Nagelschmidt", "B6.3–6.5"],
    ["Do 18.2.", "ES", "Sicherung", "Wiederholung; Gewaltspirale und Rechtsbruch des Staates gegenüberstellen.", "–", "–"],
    ["Mo 22.2.", "DS", "Ereignisse in Brandenburg", "Das Amulett der Zigeunerin; das Ende: Hinrichtung und Genugtuung zugleich.", "Schluss", "B7.1"],
    ["Do 25.2.", "ES", "Ambivalenz des Schlusses", "Rückkehr zur Leitfrage: rechtschaffen und entsetzlich?", "Schluss", "–"],
    ["Mo 1.3.", "DS", "Novelle und Erzähltechnik", "Definition der Novelle („unerhörte Begebenheit“), Erzählstrategie, Erzählerkommentar.", "Auszüge", "B7.2–7.3, AB 7–8"],
    ["Do 4.3.", "ES", "Grundbegriffe der Erzähltechnik", "Fachbegriffe sichern (für die Klausur).", "–", "AB 7"],
    ["Mo 8.3.", "DS", "Rezeption", "Kohlhaas heute: das Kohlhaas-Motiv; Vergleich Kohlhaas – Kafkas Mann vom Lande.", "–", "B8, AB 9"],
    ["Do 11.3.", "ES", "Transfer", "Leitfrage abschließend: Was tun, wenn das Recht versagt?", "–", "–"],
    ["Mo 15.3.", "DS", "Klausurvorbereitung", "Übungsanalyse eines Auszugs mit Stellungnahme; Schreibwerkstatt.", "Auszug", "–"],
    ["Do 18.3.", "ES", "Rückmeldung", "Rückmeldung; Checkliste (Präsens, „vgl. Z.“, Fachbegriffe, Wirkung deuten).", "–", "–"],
    ["Mo 22.3.", "DS", "Klausur", "Termin und Format offen (siehe offene Fragen).", "Auszug", "–"],
  ]],
];
const OFFEN = [
  ["Dauer und Klausur", "Soll die Sequenz bis Ostern laufen? Wann ist die Klausur, und welches Format (Analyse eines Auszugs mit Stellungnahme wie bei Kafka)?"],
  ["Ferien und Ausfälle", "Fallen Winterferien (Anfang Februar?) oder andere Ausfälle in die Sequenz? Dann verschiebt sich der Plan ab Februar um eine Woche."],
  ["Kontexttexte", "Wie viele Kontexttexte verträgt die Klasse (Textverständnis war zuletzt ein Problem)? Locke, Rousseau und Kant nur in sehr kurzen Auszügen – oder nur einer davon?"],
  ["Zwischenleistung", "Schreibwerkstatt im Februar als Stellungnahme – oder lieber eine Übungsklausur?"],
  ["Textausgabe", "Welche Reclam-Ausgabe genau (Seitenzählung)? Danach werden die Textstellen ergänzt."],
  ["Ferienlektüre", "Leseplan in drei Etappen mit Stichworten und Zitat – reicht das, oder soll es ein kurzes Lesetagebuch geben?"],
  ["Kleist-Fakten", "Biografische Angaben, das Zitat aus der Kant-Krise und den Bezug Kafka – Kleist vor dem Einsatz an einer Quelle prüfen."],
  ["Literatur", "Gedanken aus T. von Brand, „Ganzschriften im Deutschunterricht“, bei Bedarf in eigenen Worten nachreichen (das Buch selbst ist laut Impressum für KI-Nutzung gesperrt)."],
  ["Mentorin", "Gibt es Absprachen oder Material der Mentorin zur Einheit?"],
];

const aW = [1200, 600, 3000, 6338, 2200, 2000];
const rows = [row([hdr("Datum", aW[0]), hdr("Std.", aW[1]), hdr("Thema / Leitfrage", aW[2]), hdr("Schwerpunkt und Methode", aW[3]), hdr("Text (Reclam, S. ergänzen)", aW[4]), hdr("Baustein / Material", aW[5])])];
function row(cells) { return new TableRow({ children: cells, cantSplit: true }); }
PHASEN.forEach(([titel, zeilen]) => {
  rows.push(row([cell(p(t(titel, { bold: true, color: NAVY, size: 19, font: "Cambria" }), { after: 0 }), { w: W, span: 6, fill: LIGHT2 })]));
  zeilen.forEach(([d, s, th, sm, tx, bz]) => rows.push(row([
    cell(p(t(d, { bold: true, color: NAVY, size: 18 }), { after: 0 }), { w: aW[0], fill: LIGHT }),
    cell(p(t(s, { size: 17, color: s === "DS" ? RED : MUTED, bold: true }), { after: 0, align: AlignmentType.CENTER }), { w: aW[1], fill: LIGHT }),
    cell(p(t(th, { bold: true, size: 18 }), { after: 0 }), { w: aW[2] }),
    cell(p(t(sm, { size: 17 }), { after: 0 }), { w: aW[3] }),
    cell(p(t(tx, { size: 17, color: MUTED }), { after: 0 }), { w: aW[4] }),
    cell(p(t(bz, { size: 17, color: MUTED }), { after: 0 }), { w: aW[5] }),
  ])));
});

const kW = [2800, W - 2800];
const kopf = [
  ["Leitfrage", "Was tun, wenn das Recht versagt? – Rahmen: Kohlhaas war „einer der rechtschaffensten zugleich und entsetzlichsten Menschen seiner Zeit“."],
  ["Gerüst", "Die acht Bausteine des Unterrichtsmodells (Inhaltsverzeichnis der Lehrkraft): B1 Einführung, B2 Hintergründe, B3 Gewaltspirale, B4 Selbstjustiz, B5 Individuum und Staat, B6 Ereignisse in Sachsen, B7 Ereignisse in Brandenburg/Novelle, B8 Rezeption; AB = Arbeitsblatt, Z = Zusatzmaterial des Modells."],
  ["Leitidee", "Brücke von Kafka zu Kleist: Der Mann vom Lande wartet ein Leben lang vor dem Gesetz – Kohlhaas nimmt das Recht selbst in die Hand. Beide stehen einer Macht gegenüber, die ihnen ihr Recht verweigert."],
  ["Lektüre", "Vollständig. Das erste Drittel gemeinsam im Unterricht (bis zu den Mandaten), der Rest in den Weihnachtsferien nach Leseplan; Textausgabe Reclam."],
  ["Rhythmus", "Montag Doppelstunde (Erarbeitung, Textarbeit, Schreiben) · Donnerstag Einzelstunde (Lesen, Sichern, Wiederholen)."],
  ["Grundsätze", "Vorwissen in jeder Stunde abprüfen · Textverständnis vor Deutung · kurze Auszüge, ein Schwerpunkt pro Stunde · Erarbeitung allein · Aufgabe und Muster auf jeder Aufgabenfolie · Zielcheck · Transfer · 27 pt · keine Hausaufgaben (Ausnahme: Ferienlektüre)."],
  ["Leistungsnachweis", "Klausur am Ende der Sequenz (Termin offen) – Vorschlag: Analyse eines Gesprächsauszugs mit Stellungnahme, im Format der Kafka-Klausurvorbereitung; Zwischenübung im Februar."],
];
const footer = new Footer({ children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ font: "Calibri", size: 16, color: MUTED, children: ["Sequenzplan · Kleist, „Michael Kohlhaas“ · Deutsch 11 · Seite ", PageNumber.CURRENT] })] })] });
const doc = new Document({
  styles: { default: { document: { run: { font: "Calibri", size: 19 } } } },
  sections: [{ properties: { page: { size: { width: 11906, height: 16838, orientation: PageOrientation.LANDSCAPE }, margin: { top: 600, bottom: 500, left: 750, right: 750, footer: 300 } } }, footers: { default: footer }, children: [
    p(t("Sequenzplan: Heinrich von Kleist, „Michael Kohlhaas“", { font: "Cambria", size: 32, bold: true, color: NAVY }), { after: 20 }),
    p(t("Deutsch · Klasse 11 · Mo 30.11. bis Ostern · Entwurf (Stand 22.11.)", { size: 18, color: MUTED }), { after: 120 }),
    new Table({ width: { size: W, type: WidthType.DXA }, columnWidths: kW, rows: kopf.map(([k, v]) => row([cell(p(t(k, { bold: true, color: NAVY, size: 18 }), { after: 0 }), { w: kW[0], fill: LIGHT }), cell(p(t(v, { size: 18 }), { after: 0 }), { w: kW[1] })])) }),
    p(t(""), { after: 100 }),
    new Table({ width: { size: W, type: WidthType.DXA }, columnWidths: aW, rows }),
    p(t("DS = Doppelstunde (Montag), ES = Einzelstunde (Donnerstag). Seitenangaben werden nach der Reclam-Ausgabe ergänzt.", { size: 16, italics: true, color: MUTED }), { before: 80 }),
    p(t("Offene Fragen (von der Lehrkraft zu klären)", { font: "Cambria", size: 26, bold: true, color: RED }), { before: 0, after: 60, keepNext: true, pb: true }),
    new Table({ width: { size: W, type: WidthType.DXA }, columnWidths: [3000, W - 3600, 600], rows: [
      row([hdr("Thema", 3000), hdr("Frage", W - 3600), hdr("✓", 600)]),
      ...OFFEN.map(([k, v]) => row([cell(p(t(k, { bold: true, color: NAVY, size: 18 }), { after: 0 }), { w: 3000, fill: LIGHT }), cell(p(t(v, { size: 18 }), { after: 0 }), { w: W - 3600 }), cell(p(t("☐", { size: 22 }), { after: 0, align: AlignmentType.CENTER }), { w: 600 })])),
    ] }),
  ] }],
});
Packer.toBuffer(doc).then((b) => fs.writeFileSync("../Sequenzplan_Michael_Kohlhaas.docx", b));
