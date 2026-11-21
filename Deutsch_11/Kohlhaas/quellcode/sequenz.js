// Sequenzplan „Michael Kohlhaas“ (Kleist), Deutsch 11 – Mo Doppelstunde, Do Einzelstunde; vollständige Lektüre (Rest in den Weihnachtsferien)
const fs = require("fs");
const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, BorderStyle, ShadingType, AlignmentType, VerticalAlign, PageOrientation, Footer, PageNumber } = require("docx");
const NAVY = "1E2761", MUTED = "5B6B8C", LIGHT = "F4F7FD", LIGHT2 = "EAF0FA", RED = "D9534F", GOLD = "E8A33D";
const W = 15338;
const t = (text, o = {}) => new TextRun({ text, font: o.font || "Calibri", size: o.size || 19, bold: o.bold, italics: o.italics, color: o.color || "000000" });
const p = (runs, o = {}) => new Paragraph({ children: Array.isArray(runs) ? runs : [runs], alignment: o.align, keepNext: o.keepNext, spacing: { before: o.before ?? 0, after: o.after ?? 40, line: o.line } });
const solid = (c) => ({ style: BorderStyle.SINGLE, size: 6, color: c });
const all = (b) => ({ top: b, bottom: b, left: b, right: b });
const cell = (children, o = {}) => new TableCell({ children: Array.isArray(children) ? children : [children], width: { size: o.w, type: WidthType.DXA }, shading: o.fill ? { fill: o.fill, type: ShadingType.CLEAR, color: "auto" } : undefined, borders: all(solid("B9C6E8")), verticalAlign: o.valign || VerticalAlign.TOP, margins: { top: 50, bottom: 50, left: 100, right: 100 }, columnSpan: o.span });
const hdr = (x, w) => cell(p(t(x, { bold: true, color: "FFFFFF", size: 17 }), { after: 0 }), { w, fill: NAVY });

// [Datum, Stunde, Thema / Leitfrage, Schwerpunkt und Methode, Text (Reclam), Bezug]
const PHASEN = [
  ["Vor den Ferien: einsteigen und das erste Drittel gemeinsam lesen", [
    ["Mo 30.11.", "DS", "Rechtsgefühl – was tun, wenn das Recht versagt?", "Einstieg: Fehlentscheidung im Sport, gegen die man nichts tun kann; Brücke zu Kafka (der Mann vom Lande wartet – und du?). Erster Satz „einer der rechtschaffensten zugleich und entsetzlichsten Menschen“ als Leitfrage der Einheit. Gemeinsam lesen bis zum Schlagbaum.", "Anfang bis Schlagbaum", "Kafka: Macht/Ohnmacht; Vorwissen"],
    ["Do 3.12.", "ES", "Am Schlagbaum", "Wer sagt was – und wie? (markieren, einkreisen) Machtgefälle Kohlhaas – Burgvogt – Junker; Willkür des Passscheins.", "Schlagbaum-Szene", "Markier-Methode aus „Vor dem Gesetz“"],
    ["Mo 7.12.", "DS", "Der Rechtsweg scheitert", "Herses Bericht (Gespräch Kohlhaas – Herse); Kohlhaas’ Klage und die Vetternwirtschaft der Tronkas. Lesen und Handlungsüberblick anlegen.", "Tronkenburg, Herse, Klage", "Kommunikation, kafkaeske Bürokratie"],
    ["Do 10.12.", "ES", "Wendepunkt Lisbeth", "Bittschrift und Tod – erste Stufe der Gewaltspirale (Tafelbild beginnt).", "Lisbeth", "Ursache – Wirkung"],
    ["Mo 14.12.", "DS", "Selbstjustiz", "Überfall auf die Tronkenburg, erstes Mandat; Gewaltspirale fortsetzen. Sprache der Mandate: Selbsterhöhung.", "Rachefeldzug, Mandat", "Stilmittel, Wirkung"],
    ["Do 17.12.", "ES", "Zwischenbilanz: rechtschaffen oder entsetzlich?", "Meinungslinie: Position beziehen und mit Textbeleg begründen; Zielcheck.", "bisher Gelesenes", "Stellungnahme vorbereiten"],
    ["Mo 21.12.", "DS", "Start in die Ferienlektüre", "Leseplan in drei Etappen mit Leitfragen vorstellen; nächstes Stück gemeinsam anlesen. Leitfragen: Wie weit geht Kohlhaas? Wer kann ihn stoppen?", "Anlesen der nächsten Etappe", "Lesemotivation"],
  ]],
  ["Weihnachtsferien: den Rest der Novelle lesen", [
    ["Ferien", "–", "Ferienlektüre in drei Etappen", "Pro Etappe: drei Stichworte und ein Zitat für den Handlungsüberblick (kein Lesetagebuch). Hinweis auf ein Hörbuch. Leitfragen z. B.: Warum sucht Kohlhaas Luther auf? Warum scheitert die Amnestie? Was hat es mit dem Amulett auf sich?", "Rest der Novelle", "vollständige Lektüre"],
  ]],
  ["Nach den Ferien: Analyse mit Kenntnis des ganzen Textes", [
    ["Mo 4.1.", "DS", "Lektürecheck und Überblick", "Aussagen zum Text prüfen (richtig/falsch, mit Beleg) – Muster der Kafka-Stunden; Handlungsüberblick vervollständigen.", "ganze Novelle", "Textverständnis sichern"],
    ["Do 7.1.", "ES", "Figurenkonstellation", "Macht – Recht – Familie – Kirche: Wer steht wo? Visualisierung.", "ganze Novelle", "Bildebene: Rollen"],
    ["Mo 11.1.", "DS", "Kohlhaas und Luther", "Luthers Plakat und das Gespräch: Gesprächsanalyse (Vier-Seiten-Modell, Argumentation; „verstoßen“ aus der Gemeinschaft).", "Luther-Plakat, Gespräch", "Kommunikationsmodelle, Watzlawick"],
    ["Do 14.1.", "ES", "Sicherung Luther; Transfer", "Zielcheck; Vergleich: Luther und der Türhüter als Instanzen.", "Gespräch", "Kafka-Transfer"],
    ["Mo 18.1.", "DS", "Individuum und Staat", "Amnestie, Dresden, Nagelschmidt: Wie verspielt der Staat sein Recht? Kurzer Kontext: Gesellschaftsvertrag.", "Amnestie, Nagelschmidt", "Kontext kurz und gezielt"],
    ["Do 21.1.", "ES", "Wiederholung", "Vorwissen abprüfen; Rechtsbegriffe und Stationen sichern.", "–", "Wiederholung"],
    ["Mo 25.1.", "DS", "Das Ende", "Hinrichtung und Genugtuung zugleich; Amulett; Ambivalenz – Rückkehr zur Leitfrage.", "Schluss", "Deutung, Stellungnahme"],
    ["Do 28.1.", "ES", "Novelle und Erzähler", "„Unerhörte Begebenheit“, Erzählerkommentare, Bauform; Fachbegriffe für die Klausur.", "Auszüge", "Fachbegriffe"],
    ["Mo 1.2.", "DS", "Klausurvorbereitung", "Übungsanalyse eines Auszugs (z. B. Schlagbaum oder Luther) mit Stellungnahme; Schreibwerkstatt mit inhaltsunabhängigen Satzmustern.", "Auszug", "Klausurformat"],
    ["Do 4.2.", "ES", "Rückmeldung", "Rückmeldung zur Übung; Checkliste (Präsens, „vgl. Z.“, Fachbegriffe, Wirkung deuten).", "–", "Überarbeiten"],
    ["ab Mo 8.2.", "DS", "Klausur (Termin offen)", "Vorschlag: Analyse eines Gesprächsauszugs mit Stellungnahme zur Leitfrage.", "Auszug", "Leistungsnachweis"],
  ]],
];

const aW = [1200, 600, 3000, 6338, 2200, 2000];
const rows = [row([hdr("Datum", aW[0]), hdr("Std.", aW[1]), hdr("Thema / Leitfrage", aW[2]), hdr("Schwerpunkt und Methode", aW[3]), hdr("Text (Reclam, S. ergänzen)", aW[4]), hdr("Bezug", aW[5])])];
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
  ["Leitidee", "Brücke von Kafka zu Kleist: Der Mann vom Lande wartet ein Leben lang vor dem Gesetz – Kohlhaas nimmt das Recht selbst in die Hand. Beide stehen einer Macht gegenüber, die ihnen ihr Recht verweigert."],
  ["Lektüre", "Vollständig. Das erste Drittel gemeinsam im Unterricht (bis zu den Mandaten), der Rest in den Weihnachtsferien nach Leseplan; Textausgabe Reclam."],
  ["Rhythmus", "Montag Doppelstunde (Erarbeitung, Textarbeit, Schreiben) · Donnerstag Einzelstunde (Lesen, Sichern, Wiederholen)."],
  ["Grundsätze", "Vorwissen in jeder Stunde abprüfen · Textverständnis vor Deutung · kurze Auszüge, ein Schwerpunkt pro Stunde · Erarbeitung allein · Aufgabe und Muster auf jeder Aufgabenfolie · Zielcheck · Transfer · 27 pt · keine Hausaufgaben (Ausnahme: Ferienlektüre)."],
  ["Leistungsnachweis", "Klausur nach der Sequenz (Termin offen) – Vorschlag: Analyse eines Gesprächsauszugs mit Stellungnahme, im Format der Kafka-Klausurvorbereitung."],
];
const footer = new Footer({ children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ font: "Calibri", size: 16, color: MUTED, children: ["Sequenzplan · Kleist, „Michael Kohlhaas“ · Deutsch 11 · Seite ", PageNumber.CURRENT] })] })] });
const doc = new Document({
  styles: { default: { document: { run: { font: "Calibri", size: 19 } } } },
  sections: [{ properties: { page: { size: { width: 11906, height: 16838, orientation: PageOrientation.LANDSCAPE }, margin: { top: 600, bottom: 500, left: 750, right: 750, footer: 300 } } }, footers: { default: footer }, children: [
    p(t("Sequenzplan: Heinrich von Kleist, „Michael Kohlhaas“", { font: "Cambria", size: 32, bold: true, color: NAVY }), { after: 20 }),
    p(t("Deutsch · Klasse 11 · ab Mo 30.11. · Entwurf (Stand 21.11.)", { size: 18, color: MUTED }), { after: 120 }),
    new Table({ width: { size: W, type: WidthType.DXA }, columnWidths: kW, rows: kopf.map(([k, v]) => row([cell(p(t(k, { bold: true, color: NAVY, size: 18 }), { after: 0 }), { w: kW[0], fill: LIGHT }), cell(p(t(v, { size: 18 }), { after: 0 }), { w: kW[1] })])) }),
    p(t(""), { after: 100 }),
    new Table({ width: { size: W, type: WidthType.DXA }, columnWidths: aW, rows }),
    p(t("DS = Doppelstunde (Montag), ES = Einzelstunde (Donnerstag). Seitenangaben werden nach der Reclam-Ausgabe ergänzt.", { size: 16, italics: true, color: MUTED }), { before: 80 }),
  ] }],
});
Packer.toBuffer(doc).then((b) => fs.writeFileSync("../Sequenzplan_Michael_Kohlhaas.docx", b));
