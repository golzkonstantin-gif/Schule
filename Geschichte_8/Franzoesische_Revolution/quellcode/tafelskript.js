const fs = require("fs");
const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, BorderStyle, ShadingType, AlignmentType, VerticalAlign, Header, Footer, PageNumber } = require("docx");

const NAVY = "1E2761", MUTED = "5B6B8C", LIGHT = "F4F7FD", BOARD = "F3F6F1", RED = "D9534F", DARK = "3D4A3F";
const W = 10206;
const t = (text, o = {}) => new TextRun({ text, font: o.font || "Calibri", size: o.size || 22, bold: o.bold, italics: o.italics, color: o.color || "000000" });
const p = (runs, o = {}) => new Paragraph({ children: Array.isArray(runs) ? runs : [runs], alignment: o.align, spacing: { before: o.before ?? 0, after: o.after ?? 100 }, keepNext: o.keepNext, indent: o.indent });
const solid = (c, s = 8) => ({ style: BorderStyle.SINGLE, size: s, color: c });
const all = (b) => ({ top: b, bottom: b, left: b, right: b });
function cell(children, o = {}) {
  return new TableCell({ children: Array.isArray(children) ? children : [children], width: { size: o.w, type: WidthType.DXA }, shading: o.fill ? { fill: o.fill, type: ShadingType.CLEAR, color: "auto" } : undefined, borders: o.borders || all(solid("B9C6E8", 6)), verticalAlign: o.valign || VerticalAlign.CENTER, margins: { top: o.m ?? 70, bottom: o.m ?? 70, left: 110, right: 110 } });
}
const table = (widths, rows) => new Table({ width: { size: widths.reduce((a, b) => a + b, 0), type: WidthType.DXA }, columnWidths: widths, rows });
const row = (cells, h) => new TableRow({ children: cells, height: h ? { value: h, rule: "atLeast" } : undefined, cantSplit: true });
const th = (txt, w, fill = DARK, size = 19) => cell(p(t(txt, { bold: true, color: "FFFFFF", size }), { after: 0, align: AlignmentType.CENTER }), { w, fill, borders: all(solid(DARK, 8)) });
const tc = (txt, w, o = {}) => cell(p(t(txt, { size: o.size || 20, bold: o.bold }), { after: 0 }), { w, fill: o.fill });
const kicker = (s) => p(t(s.toUpperCase(), { size: 18, color: MUTED, bold: true }), { after: 20 });
const h1 = (s) => p(t(s, { font: "Cambria", size: 38, bold: true, color: NAVY }), { after: 60 });
const h3 = (s) => p(t(s, { font: "Cambria", size: 24, bold: true, color: NAVY }), { before: 160, after: 60, keepNext: true });
const S = (s, o = {}) => t(s, Object.assign({ size: 21 }, o));
const header = new Header({ children: [new Paragraph({ tabStops: [{ type: "right", position: W }], children: [t("Geschichte · Klasse 8 · Französische Revolution", { size: 17, color: MUTED }), t("\tfür die Lehrkraft", { size: 17, color: MUTED })] })] });
const footer = new Footer({ children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ font: "Calibri", size: 16, color: MUTED, children: ["Tafelskript · Seite ", PageNumber.CURRENT] })] })] });

const BW = 1700, CW = W - BW, TW = CW - 340;
function step(klasse, parts) {
  const ch = [];
  if (parts.say) ch.push(p([t("Ansage: ", { bold: true, size: 20, color: RED }), t("„" + parts.say + "“", { italics: true, size: 21 })], { after: 60 }));
  (parts.do || []).forEach((x) => ch.push(p(S(x), { after: 50 })));
  if (parts.extra) ch.push(...parts.extra);
  if (parts.sol) ch.push(p([t("Erwartung: ", { bold: true, size: 19, color: MUTED }), t(parts.sol, { size: 19, color: MUTED })], { after: 0 }));
  return table([BW, CW], [row([cell([p(t("Klasse", { size: 16, color: MUTED }), { after: 10 }), p(t(klasse, { bold: true, color: NAVY, size: 20 }), { after: 0 })], { w: BW, fill: LIGHT, valign: VerticalAlign.TOP }), cell(ch, { w: CW, valign: VerticalAlign.TOP })])]);
}
const gapP = () => p(t(""), { after: 30 });
const phase = (nr, title, min) => p([t(`${nr}  `, { font: "Cambria", size: 28, bold: true, color: RED }), t(title, { font: "Cambria", size: 28, bold: true, color: NAVY }), t(`   ${min} Min.`, { size: 20, bold: true, color: "C77C12" })], { before: 180, after: 90, keepNext: true });
const boardBox = (title, children) => table([TW], [row([cell(p(t("TAFEL · " + title, { bold: true, color: "FFFFFF", size: 16 }), { after: 0 }), { w: TW, fill: DARK, borders: all(solid(DARK, 12)), m: 40 })]), row([cell(children, { w: TW, fill: BOARD, borders: all(solid(DARK, 12)), valign: VerticalAlign.TOP, m: 100 })])]);
const bl = (s, o = {}) => p(t(s, { font: "Cambria", size: 20 }), Object.assign({ after: 30 }, o));

// Raster (Lösung)
const RW = [1700, 3100, 3366];
const raster = table(RW, [
  row([th("Wann?", RW[0]), th("Was geschah?", RW[1]), th("Welches Problem steckt dahinter?", RW[2])], 340),
  ...[
    ["A  Mai – Juni 1789", "Der König ruft die Generalstände zusammen (erstmals seit 1614). Der Dritte Stand wird bei Abstimmungen nach Ständen immer überstimmt. Er erklärt sich zur Nationalversammlung (17. Juni) und schwört im Ballhaus (20. Juni), nicht auseinanderzugehen, bis es eine Verfassung gibt.", "leere Staatskasse · Dritter Stand (Bauern und Bürger) ohne Mitsprache"],
    ["B  14. Juli 1789", "Hunger in Paris, Soldaten des Königs rund um die Stadt. Die Pariser stürmen die Bastille, die Festung und das Gefängnis des Königs.", "Hunger und hoher Brotpreis · Angst vor der Willkür des Königs"],
    ["C  August 1789", "Unruhen auf dem Land. Die Nationalversammlung schafft Vorrechte ab (4./5. August) und verkündet die Menschen- und Bürgerrechte (26. August).", "Vorrechte von Adel und Klerus · ungleiche Lasten"],
  ].map((r) => row(r.map((v, i) => tc(v, RW[i], { size: 19, bold: i === 0, fill: i === 0 ? BOARD : undefined })), 600)),
]);

const d = [];
d.push(kicker("Tafelskript für die Lehrkraft · Einzelstunde (45 Min.) · 12 Schülerinnen und Schüler"), h1("Französische Revolution 1789: Gruppenpuzzle"));
d.push(p(t("Ziel: Die Klasse erarbeitet drei Ereignisse des Jahres 1789 (Generalstände und Nationalversammlung, Sturm auf die Bastille, Ende der Vorrechte und Menschenrechte) und ordnet jedes dem Problem zu, das in der letzten Stunde sichtbar wurde (leere Kasse, ungleiche Lasten, Hunger, keine Mitsprache). Anbindung: Die Stimmung ist in der „Münz“-Geschichte gekippt. Jetzt erfährt die Klasse, wie. Es gibt keinen Hörtext. Material: drei Expertentexte, das Raster entsteht an der Tafel und im Heft.", { size: 21 }), { after: 80 }));

d.push(table([800, 6200, 3206], [
  row([th("Min.", 800), th("Phase", 6200), th("Arbeitsform", 3206)], 320),
  ...[["4", "0  Einstieg: Bildimpuls Sturm auf die Bastille", "gemeinsam, mündlich"], ["10", "1  Expertengruppen A, B, C: Text und Raster", "allein, dann in der Gruppe"], ["12", "2  Stammgruppen: Experten erklären sich ihre Themen", "in der Gruppe, 3 Personen"], ["12", "3  Sicherung an der Tafel: Zeitstrahl und Probleme zuordnen", "gemeinsam, an der Tafel"], ["5", "4  Abschluss: Warum begann die Revolution 1789?", "allein, im Heft"], ["2", "Wechsel und Ansagen (Puffer)", "—"]].map((r) => row(r.map((v, i) => cell(p(t(v, { size: 20, bold: i < 2 }), { after: 0 }), { w: [800, 6200, 3206][i] })), 320)),
]));

d.push(h3("Gruppeneinteilung (12 Personen, ohne Namen)"));
d.push(table([2400, 3600, 4206], [
  row([th("Gruppe", 2400), th("Mitglieder", 3600), th("Thema", 4206)], 320),
  ...[["Expertengruppe A", "A1, A2, A3, A4", "Von der leeren Kasse zur Nationalversammlung"], ["Expertengruppe B", "B1, B2, B3, B4", "Der Sturm auf die Bastille"], ["Expertengruppe C", "C1, C2, C3, C4", "Das Ende der Vorrechte und die Menschenrechte"]].map((r) => row(r.map((v, i) => tc(v, [2400, 3600, 4206][i], { bold: i === 0 })), 320)),
  ...[1, 2, 3, 4].map((n) => row([tc(`Stammgruppe ${n}`, 2400, { bold: true, fill: LIGHT }), tc(`A${n}, B${n}, C${n}`, 3600, { fill: LIGHT }), tc("ein Experte pro Thema", 4206, { fill: LIGHT })], 300)),
]));
d.push(p(S("Vorab: Jede Schülerin und jeder Schüler erhält eine Buchstaben-Zahl-Kombination (zum Beispiel A3) und weiß so Expertengruppe und Stammgruppe. Die Zuteilung bleibt bei Ihnen.", { italics: true, color: MUTED }), { before: 60, after: 60 }));

d.push(h3("Eindeutige Ansagen (wörtlich)"));
d.push(table([4700, 5506], [
  ["„Stifte liegen. Ihr schaut nur.“", "Bildimpuls, niemand schreibt."],
  ["„Meldet euch. Ihr schreibt noch nichts auf.“", "Gemeinsames Gespräch."],
  ["„Jetzt arbeitet ihr allein. Die Aufgabe steht rechts.“", "Einzelarbeit am Text."],
  ["„Jetzt arbeitet ihr in eurer Gruppe.“", "Austausch in Expertengruppe oder Stammgruppe."],
  ["„Wer drankommt, schreibt an die Tafel. Alle prüfen mit.“", "Sicherung an der Tafel."],
  ["„Vergleicht und verbessert mit Grün.“", "Kontrolle im eigenen Heft."],
].map(([a, b]) => row([tc(a, 4700, { fill: LIGHT, bold: true }), tc(b, 5506)], 340))));

// Phase 0
d.push(phase("0", "Einstieg", 4));
d.push(step("schaut und meldet sich", {
  say: "Stifte liegen. Ihr schaut nur. Dann meldet ihr euch.",
  do: [
    "Folie 1: Bild vom Sturm auf die Bastille (Gemälde von 1789, z. B. von Jean-Pierre Houël, gemeinfrei; selbst einfügen).",
    "Fragen: Was ist zu sehen? Was ist wohl vorher passiert? Was kennt ihr aus der letzten Stunde, das dazu passt? (Marie, Hunger, Steuern, „Heute sagt niemand etwas“.)",
    "Leitfrage an die Tafel (bleibt stehen): „Wie kippt die Stimmung 1789?“ Dann kurz den Ablauf nennen: Experten, Stammgruppen, Tafel.",
  ],
}));

// Phase 1
d.push(phase("1", "Expertengruppen", 10));
d.push(step("arbeitet allein, dann in der Gruppe", {
  say: "Setzt euch in eure Expertengruppen. Zeichnet zuerst das Raster von der Tafel ab. Jetzt arbeitet ihr allein. Die Aufgabe steht rechts.",
  do: ["Die drei Texte auf die Gruppen verteilen. Raster (leer, drei Spalten) steht an der Tafel. Ablauf: 2 Minuten Raster abzeichnen, 3 Minuten Einzelarbeit, 5 Minuten Gruppe."],
  extra: [boardBox("rechts · Ü1 Expertengruppe", [
    p(t("Ü1 Expertengruppe", { font: "Cambria", size: 22, bold: true, color: NAVY }), { after: 40 }),
    bl("1. Lies deinen Text und unterstreiche Stichworte."), bl("2. Ergänze im Raster nur deine Zeile."), bl("3. Vergleicht in der Gruppe und verbessert."), bl("4. Überlegt: Wie erklärt ihr es in 3 Minuten?", { after: 40 }),
    p([t("Zeit: ", { bold: true, size: 19 }), t("10 Minuten · erst allein, dann in der Gruppe · leise     ", { size: 19 }), t("Fertig? Stift hinlegen.", { size: 19, italics: true })], { after: 0 }),
  ]), gapP()],
  sol: "Leere Raster-Spalten: Wann? · Was geschah? · Welches Problem steckt dahinter? Jede Expertengruppe füllt nur ihre Zeile. Nach 3 Minuten Einzelarbeit: „Jetzt arbeitet ihr in eurer Gruppe.“",
}));

// Phase 2
d.push(phase("2", "Stammgruppen", 12));
d.push(step("arbeitet in der Gruppe", {
  say: "Jetzt arbeitet ihr in der Stammgruppe. Setzt euch zu A, B und C zusammen. Erst erklärt A, dann B, dann C. Die anderen ergänzen ihr Raster.",
  do: ["Je Experte 3 Minuten erklären, 1 Minute Nachfragen. Zeit ansagen und beim Wechsel melden: „A hat fertig, jetzt erklärt B.“ Am Ende 2 Minuten für den Vergleich: Welches Problem kommt in mehreren Ereignissen vor?"],
  extra: [boardBox("rechts · Ü2 Stammgruppe", [
    p(t("Ü2 Stammgruppe", { font: "Cambria", size: 22, bold: true, color: NAVY }), { after: 40 }),
    bl("1. Erkläre dein Thema in eigenen Worten (3 Minuten)."), bl("2. Ergänze das Raster mit den Themen der anderen."), bl("3. Stellt Nachfragen, bis alles klar ist.", { after: 40 }),
    p([t("Zeit: ", { bold: true, size: 19 }), t("12 Minuten · in der Gruppe · leise     ", { size: 19 }), t("Fertig? Stift hinlegen.", { size: 19, italics: true })], { after: 0 }),
  ]), gapP()],
}));

// Phase 3
d.push(phase("3", "Sicherung an der Tafel", 12));
d.push(step("kommt an die Tafel", {
  say: "Setzt euch wieder auf euren Platz. Wer drankommt, schreibt an die Tafel. Alle prüfen mit. … Vergleicht und verbessert mit Grün.",
  do: [
    "Zeitstrahl gemeinsam entwickeln (Schüler nennen die Ereignisse, Sie schreiben oder lassen anschreiben):",
  ],
  extra: [boardBox("Mitte · Zeitstrahl und Probleme", [
    bl("5. Mai: Generalstände  →  Problem: leere Kasse, Dritter Stand überstimmt"),
    bl("17. Juni: Nationalversammlung  ·  20. Juni: Ballhausschwur  →  Problem: keine Mitsprache"),
    bl("14. Juli: Sturm auf die Bastille  →  Problem: Hunger, Angst vor Willkür"),
    bl("4./5. August: Ende der Vorrechte  →  Problem: Ungleichheit der Stände"),
    bl("26. August: Menschen- und Bürgerrechte  →  Ziel: alle Menschen frei und gleich", { after: 0 }),
  ]), gapP(),
    p(S("Fragen nacheinander:  1  Welches Ereignis löst welches Problem? (Zuordnung begründen)  2  Was ändert sich für den Adel, den Klerus und den Dritten Stand (Bauern und Bürger)?  3  Gibt es ein Problem, das 1789 noch nicht gelöst ist? (Hunger, leere Kasse)"), { after: 50 }),
    raster],
  sol: "Das Raster oben ist die Lösung für die Zeilen A, B und C. Die Klasse vergleicht ihr Heft damit und verbessert grün. Zu Frage 2: Adel und Klerus verlieren Vorrechte und müssen Steuern zahlen, der Dritte Stand (Bauern und Bürger) gewinnt Rechte und Mitsprache. Hinweis: Mit „Verfassung“ ist gemeint, dass auch der König an Gesetze gebunden ist (Ende des Absolutismus).",
}));

// Phase 4
d.push(phase("4", "Abschluss", 5));
d.push(step("arbeitet allein", {
  say: "Jetzt arbeitet ihr allein. Die Aufgabe steht rechts.",
  extra: [boardBox("rechts · Ü3 Abschluss", [
    p(t("Ü3 Warum begann die Revolution 1789?", { font: "Cambria", size: 22, bold: true, color: NAVY }), { after: 40 }),
    bl("1. Schreibe Ü3 ins Heft."), bl("2. Nenne zwei Gründe für die Revolution."), bl("3. Begründe, welcher Grund wichtiger war.", { after: 40 }),
    p([t("Zeit: ", { bold: true, size: 19 }), t("4 Minuten · allein · leise     ", { size: 19 }), t("Fertig? Stift hinlegen.", { size: 19, italics: true })], { after: 0 }),
  ]), gapP()],
  sol: "Mögliche Gründe: leere Staatskasse, ungleiche Steuerlast, keine Mitsprache des Dritten Standes (Bauern und Bürger), Hunger und Brotpreis, Angst vor der Willkür des Königs. Zum Schluss 1 bis 2 Antworten vorlesen lassen.",
}));

d.push(h3("So sieht das Heft am Ende der Stunde aus"));
d.push(table([2300, 7906], [
  ["Überschrift", "Die Französische Revolution 1789 (mit Lineal unterstrichen) · Datum rechts"],
  ["Raster", "Wann? · Was geschah? · Welches Problem steckt dahinter? mit drei Zeilen A, B, C (Verbesserungen grün)"],
  ["Ü3", "Zwei Gründe für die Revolution und eine Begründung, welcher wichtiger war"],
].map(([a, b]) => row([tc(a, 2300, { bold: true, fill: LIGHT }), tc(b, 7906)], 340))));

d.push(h3("Zum Prüfen"));
[
  "Die drei Expertentexte sind eigens für diese Stunde geschrieben und an Ihre bisherigen Texte angelehnt. Faktenstand: Generalstände 5. Mai 1789 (zuletzt 1614), 17. Juni Nationalversammlung, 20. Juni Ballhausschwur, 14. Juli Bastille (sieben Gefangene), 4./5. August Ende der Vorrechte, 26. August Menschen- und Bürgerrechte. Bitte mit Ihrem Buch gegenlesen.",
  "Kürzungen und Vereinfachungen: Die Abschaffung der Vorrechte ist im Text zusammengefasst (in Wirklichkeit wurden manche Abgaben erst später und gegen Entschädigung abgeschafft). Das lässt sich im Gespräch richtigstellen.",
  "Wenn die Zeit knapp wird: Ü3 mündlich statt im Heft. Die Fragen 2 und 3 der Sicherung entfallen dann.",
].forEach((x) => d.push(p([S("• "), S(x)], { after: 40, indent: { left: 240 } })));

(async () => {
  const doc = new Document({ styles: { default: { document: { run: { font: "Calibri", size: 22 } } } }, sections: [{ properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 1000, bottom: 850, left: 850, right: 850 } } }, headers: { default: header }, footers: { default: footer }, children: d }] });
  fs.writeFileSync("../Tafelskript_Franzoesische_Revolution.docx", await Packer.toBuffer(doc));
  console.log("ok");
})();
