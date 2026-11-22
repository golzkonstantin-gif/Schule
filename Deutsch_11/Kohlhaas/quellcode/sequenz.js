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
  ["Vor den Ferien: Einführung, Kleist, Hintergründe, Gewaltspirale (gemeinsam lesen)", [
    ["Mo 30.11.", "DS", "Rechtsgefühl", "Gedankenexperiment: Unrecht erlitten, keiner hilft · hinnehmen – beschweren – selbst handeln? (Meinungslinie) · Klappentext: Erwartungen · Vorlage Handlungsverlauf · Lesebeginn", "Anfang", "B1.1–1.3, AB 1"],
    ["Do 3.12.", "ES", "Die Einleitung", "„rechtschaffensten zugleich und entsetzlichsten“ · Antithese · Erzählerhaltung · Leitfrage an die Tafel", "erster Absatz", "B2.3"],
    ["Mo 7.12.", "DS", "Kleist und seine Zeit", "Kant-Krise (Zitat) · Kleist in fünf Stationen · Hans Kohlhase / Kohlhasenbrück · Bürgertum 18. Jh. · Rechtsverhältnisse 16. Jh. · preußische Reformen", "–", "B2.1–2.2, AB 2–3, Z3, Z5"],
    ["Do 10.12.", "ES", "Am Schlagbaum", "Wer sagt was, wie? (markieren, einkreisen) · Machtgefälle · Passschein", "Schlagbaum", "B3.1"],
    ["Mo 14.12.", "DS", "Kohlhaas prüft – Kohlhaas als Advokat", "Prüfung am Schlagbaum · Gespräch Kohlhaas – Herse · Kommunikationsanalyse", "Tronkenburg, Herse", "B3.2–3.3, Z4"],
    ["Do 17.12.", "ES", "Legale Lösungen", "Klage · Bittschrift · Lisbeths Tod · Wendepunkt", "bis Lisbeths Tod", "B3.4"],
    ["Mo 21.12.", "DS", "Weg zur Rache", "Gewaltspirale (Tafelbild) · Meinungslinie „rechtschaffen oder entsetzlich?“ · Leseplan Ferien", "bisher Gelesenes", "AB 4"],
  ]],
  ["Weihnachtsferien", [
    ["Ferien", "–", "Ferienlektüre", "3 Etappen · je 3 Stichworte + 1 Zitat · Leitfragen · Hörbuch-Hinweis", "Rest der Novelle", "B1.1–1.2"],
  ]],
  ["Nach den Ferien: Selbstjustiz, Individuum und Staat (B4–B5)", [
    ["Mo 4.1.", "DS", "Lektürecheck", "Aussagen prüfen (richtig/falsch + Beleg) · Handlungsüberblick", "ganze Novelle", "B1.2"],
    ["Do 7.1.", "ES", "Figurenkonstellation", "Macht – Recht – Familie – Kirche · Visualisierung", "ganze Novelle", "–"],
    ["Mo 11.1.", "DS", "Rachefeldzug", "Selbstjustiz · Mandate", "Mandate", "B4.1"],
    ["Do 14.1.", "ES", "Selbsterhöhung", "biblische Anspielungen · „Statthalter Michaels“", "Mandate", "B4.2, AB 5"],
    ["Mo 18.1.", "DS", "Wittenberg", "Luthers Plakat · Gespräch Kohlhaas – Luther · Vier-Seiten-Modell", "Plakat, Gespräch", "B5.1"],
    ["Do 21.1.", "ES", "Kohlhaas’ Logik", "Argumente prüfen · „verstoßen“", "Gespräch", "B5.2, Z7"],
    ["Mo 25.1.", "DS", "Einzelner und Gesellschaft", "Gesellschaftsvertrag (Kurzauszüge) · Schutz und Zugehörigkeit", "Gespräch", "B5.3, Z8–Z10"],
    ["Do 28.1.", "ES", "Guter Wille?", "Kant kurz · Rückbezug Kant-Krise", "–", "AB 6"],
    ["Mo 1.2.", "DS", "Schreibwerkstatt", "Stellungnahme „Darf man sich sein Recht selbst nehmen?“ · Satzmuster", "–", "–"],
    ["Do 4.2.", "ES", "Rückmeldung", "Rückmeldung · Checkliste", "–", "–"],
  ]],
  ["Sachsen, Brandenburg, Novelle, Rezeption (B6–B8)", [
    ["Mo 8.2.", "DS", "Sachsen", "Staatsrat · Kalkül statt Recht", "Staatsrat", "B6.1"],
    ["Do 11.2.", "ES", "Abdeckerszene", "Ehre · Volk · Absurdität", "Abdecker", "B6.2"],
    ["Mo 15.2.", "DS", "Nagelschmidt", "Amnestiebruch · Staat bricht Recht", "Nagelschmidt", "B6.3–6.5"],
    ["Do 18.2.", "ES", "Sicherung", "Wiederholung · Gewaltspirale vs. Rechtsbruch des Staates", "–", "–"],
    ["Mo 22.2.", "DS", "Brandenburg", "Amulett · Hinrichtung und Genugtuung", "Schluss", "B7.1"],
    ["Do 25.2.", "ES", "Ambivalenz des Schlusses", "Leitfrage: rechtschaffen und entsetzlich?", "Schluss", "–"],
    ["Mo 1.3.", "DS", "Novelle, Erzähltechnik", "„unerhörte Begebenheit“ · Erzählstrategie · Erzählerkommentar", "Auszüge", "B7.2–7.3, AB 7–8"],
    ["Do 4.3.", "ES", "Fachbegriffe Erzähltechnik", "Sichern für die Klausur", "–", "AB 7"],
    ["Mo 8.3.", "DS", "Rezeption", "Kohlhaas-Motiv heute · Rückblick: Kohlhaas – Kafkas Mann vom Lande", "–", "B8, AB 9"],
    ["Do 11.3.", "ES", "Transfer", "Leitfrage abschließend", "–", "–"],
    ["Mo 15.3.", "DS", "Klausurvorbereitung", "Übungsanalyse Auszug + Stellungnahme", "Auszug", "–"],
    ["Do 18.3.", "ES", "Rückmeldung", "Checkliste: Präsens, „vgl. Z.“, Fachbegriffe, Wirkung", "–", "–"],
    ["Mo 22.3.", "DS", "Klausur", "Termin, Format offen", "Auszug", "–"],
  ]],
];
const OFFEN = [
  ["Dauer und Klausur", "Bis Ostern? Klausurtermin? Format (Auszug + Stellungnahme)?"],
  ["Ferien und Ausfälle", "Winterferien / Ausfälle? → Plan ab Februar verschieben"],
  ["Kontexttexte", "Locke, Rousseau, Kant – alle (kurz) oder nur einer?"],
  ["Zwischenleistung", "Stellungnahme oder Übungsklausur?"],
  ["Textausgabe", "Welche Reclam-Ausgabe (Seitenzahlen)?"],
  ["Ferienlektüre", "Stichworte + Zitat genug oder kurzes Lesetagebuch?"],
  ["Kleist-Fakten", "Biografie und Kant-Zitat an Quelle prüfen"],
  ["Literatur", "Gedanken aus von Brand ggf. in eigenen Worten nachreichen (Buch für KI gesperrt)"],
  ["Mentorin", "Absprachen oder Material?"],
];

const aW = [1200, 600, 3200, 6138, 2200, 2000];
const rows = [row([hdr("Datum", aW[0]), hdr("Std.", aW[1]), hdr("Thema / Leitfrage", aW[2]), hdr("Schwerpunkt und Methode", aW[3]), hdr("Text (Reclam, S. ergänzen)", aW[4]), hdr("Baustein / Material", aW[5])])];
function row(cells) { return new TableRow({ children: cells, cantSplit: true }); }
PHASEN.forEach(([titel, zeilen]) => {
  rows.push(row([cell(p(t(titel, { bold: true, color: NAVY, size: 19, font: "Cambria" }), { after: 0 }), { w: W, span: 6, fill: LIGHT2 })]));
  zeilen.forEach(([d, s, th, sm, tx, bz]) => rows.push(row([
    cell(p(t(d, { bold: true, color: NAVY, size: 18 }), { after: 0 }), { w: aW[0], fill: LIGHT }),
    cell(p(t(s, { size: 17, color: s === "DS" ? RED : MUTED, bold: true }), { after: 0, align: AlignmentType.CENTER }), { w: aW[1], fill: LIGHT }),
    cell(p(t(th, { bold: true, size: 18 }), { after: 0 }), { w: aW[2] }),
    cell(sm.split(" · ").map((x) => p(t("• " + x, { size: 17 }), { after: 0 })), { w: aW[3] }),
    cell(p(t(tx, { size: 17, color: MUTED }), { after: 0 }), { w: aW[4] }),
    cell(p(t(bz, { size: 17, color: MUTED }), { after: 0 }), { w: aW[5] }),
  ])));
});

const kW = [2800, W - 2800];
const kopf = [
  ["Leitfrage", "Was tun, wenn das Recht versagt? · Rahmen: „einer der rechtschaffensten zugleich und entsetzlichsten Menschen seiner Zeit“"],
  ["Leitidee", "Rechtsgefühl kippt in Selbstjustiz – wo genau? · Kafka nur als Rückblick (Rezeption)"],
  ["Gerüst", "Bausteine B1–B8 des Unterrichtsmodells · AB = Arbeitsblatt · Z = Zusatzmaterial"],
  ["Lektüre", "vollständig · erstes Drittel gemeinsam · Rest in den Weihnachtsferien · Reclam"],
  ["Rhythmus", "Mo DS: Erarbeitung, Schreiben · Do ES: Lesen, Sichern"],
  ["Grundsätze", "Vorwissen abprüfen · Textverständnis vor Deutung · ein Schwerpunkt · Erarbeitung allein · Aufgabe + Muster · Zielcheck · 27 pt · keine Hausaufgaben (außer Ferienlektüre)"],
  ["Leistungsnachweis", "Klausur am Ende (Termin offen) · Zwischenübung Februar"],
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
