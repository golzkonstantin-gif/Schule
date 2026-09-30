const fs = require("fs");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, BorderStyle,
  ShadingType, AlignmentType, VerticalAlign, HeightRule, Footer, PageNumber, PageOrientation,
} = require("docx");

const NAVY = "1E2761", MUTED = "5B6B8C", LIGHT = "F4F7FD", LIGHT2 = "EAF0FA", RED = "D9534F";
// Tafel: dunkelgrün, Kreidefarben nach Schulz von Thun (Sache blau, Selbstoffenbarung grün, Beziehung gelb, Appell rot)
const BOARD = "2E4A3B", BOARD2 = "3A5A48", CHALK = "F3F1EA";
const K = { sach: "8EC3F0", selbst: "9BDB9B", bez: "F5D76E", app: "F29A8E" };
const W = 15138; // Inhaltsbreite A4 quer, 1,5 cm Rand

const t = (text, o = {}) => new TextRun({ text, font: o.font || "Calibri", size: o.size || 22, bold: o.bold, italics: o.italics, color: o.color || "000000" });
const p = (runs, o = {}) => new Paragraph({ children: Array.isArray(runs) ? runs : [runs], alignment: o.align, pageBreakBefore: o.pb, spacing: { before: o.before ?? 0, after: o.after ?? 100 }, keepNext: o.keepNext });
const h1 = (text) => p(t(text, { font: "Cambria", size: 40, bold: true, color: NAVY }), { after: 60 });
const kicker = (text, pb) => p(t(text.toUpperCase(), { size: 18, color: MUTED, bold: true }), { after: 20, pb });
const h2 = (label, text) => p([t(label + "  ", { font: "Cambria", size: 26, bold: true, color: RED }), t(text, { font: "Cambria", size: 26, bold: true, color: NAVY })], { before: 160, after: 80, keepNext: true });
const gap = (after = 100) => p(t(""), { after });

const none = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
const solid = (c = NAVY, s = 8) => ({ style: BorderStyle.SINGLE, size: s, color: c });
const allBorders = (b) => ({ top: b, bottom: b, left: b, right: b });

function cell(children, o = {}) {
  return new TableCell({
    children: Array.isArray(children) ? children : [children],
    width: { size: o.w, type: WidthType.DXA },
    shading: o.fill ? { fill: o.fill, type: ShadingType.CLEAR, color: "auto" } : undefined,
    borders: o.borders || allBorders(solid("B9C6E8", 6)),
    verticalAlign: o.valign || VerticalAlign.TOP,
    margins: { top: o.m ?? 60, bottom: o.m ?? 60, left: o.ml ?? 110, right: o.ml ?? 110 },
    columnSpan: o.span,
  });
}
const table = (widths, rows) => new Table({ width: { size: widths.reduce((a, b) => a + b, 0), type: WidthType.DXA }, columnWidths: widths, rows });
const row = (cells, h) => new TableRow({ children: cells, height: h ? { value: h, rule: HeightRule.ATLEAST } : undefined, cantSplit: true });
const hdr = (txt, w, fill = NAVY) => cell(p(t(txt, { bold: true, color: "FFFFFF", size: 20 }), { align: AlignmentType.CENTER, after: 0 }), { w, fill, valign: VerticalAlign.CENTER });
const infoBox = (paras, fill = LIGHT2, w = W) => table([w], [row([cell(paras.map((r) => p(r, { after: 40 })), { w, fill, borders: allBorders(none), m: 120 })])]);

// Kreide-Helfer
const ck = (text, o = {}) => t(text, { font: o.font || "Calibri", size: o.size || 20, bold: o.bold, italics: o.italics, color: o.color || CHALK });
const phase = (n) => ck(["", "①", "②", "③", "④", "⑤", "⑥"][n] + " ", { size: 18, color: "C9C3A8", bold: true });
const chalkBorder = allBorders(solid("6E8C7A", 6));

// ================= Inhalte =================
const seiten = [
  {
    key: "sach", name: "Sachinhalt", frage: "Worüber informiert er?",
    zitat: ["„Es gibt im Moment in diese Mannschaft, oh, einige Spieler vergessen ihnen Profi, was sie sind.“", "„Diese Spieler beklagen mehr als sie spielen.“"],
    inhalt: "Einige Spieler (Strunz, Basler, Scholl) leisten zu wenig, beklagen sich und verweigern die Defensivarbeit.",
    stoerung: "Der Sachinhalt geht im „Rauschen“ unter: Medien und Publikum hören nur noch „Flasche leer“ und „Was erlauben Strunz?“ – nicht die sachliche Kritik.",
  },
  {
    key: "selbst", name: "Selbstoffenbarung", frage: "Was zeigt er von sich?",
    zitat: ["„Ein Trainer ist nicht ein Idiot!“", "„Ich bin müde jetzt Vater diese Spieler …“"],
    inhalt: "Purer Frust, Ohnmacht, Überforderung mit der Mentalität deutscher Profis.",
    stoerung: "Die Selbstoffenbarung überlagert alles: Er lässt seine Wut ungefiltert heraus und schützt damit vor allem sich selbst.",
  },
  {
    key: "bez", name: "Beziehung", frage: "Was hält er von den Spielern?",
    zitat: ["„… schwach wie eine Flasche leer!“", "„Was erlauben Strunz?“"],
    inhalt: "Er entzieht den Spielern den Respekt, den sie seiner Meinung nach selbst vermissen lassen („müssen respektieren die andere Kollegen“).",
    stoerung: "Statt interner Kritik: maximale öffentliche Demütigung einzelner Spieler mit Namen – über die Medien.",
  },
  {
    key: "app", name: "Appell", frage: "Wozu will er bewegen?",
    zitat: ["„… Samstag, diese Spieler müssen zeigen mich, seine Fans, müssen allein die Spiel gewinnen.“"],
    inhalt: "Die Spieler sollen am Samstag laufen, kämpfen und das Spiel gewinnen.",
    stoerung: "Falscher Empfänger: Der Appell gilt den Spielern, geht aber per Pressekonferenz an Journalisten – die Spieler erfahren ihn aus der Zeitung.",
  },
];

// ================= Seite 1: Tafelbild (Endstand) =================
const TW = W;
const c1 = 3100, c2 = 5900, c3 = TW - c1 - c2;
const earTable = table([c1, c2, c3], [
  row([
    cell(p([phase(2), ck("Seite der Nachricht", { bold: true })], { after: 0 }), { w: c1, fill: BOARD2, borders: chalkBorder }),
    cell(p([phase(3), ck("Was sendet Trapattoni? – Belege (Zitate)", { bold: true })], { after: 0 }), { w: c2, fill: BOARD2, borders: chalkBorder }),
    cell(p([phase(4), ck("Kommunikationsstörung", { bold: true })], { after: 0 }), { w: c3, fill: BOARD2, borders: chalkBorder }),
  ], 420),
  ...seiten.map((s) => row([
    cell([p(ck(s.name, { bold: true, size: 24, color: K[s.key] }), { after: 20 }), p(ck(s.frage, { size: 17, italics: true, color: K[s.key] }), { after: 0 })], { w: c1, fill: BOARD, borders: chalkBorder }),
    cell([
      p(ck(s.inhalt, { size: 18 }), { after: 40 }),
      ...s.zitat.map((z) => p(ck(z, { size: 18, italics: true, color: K[s.key] }), { after: 20 })),
    ], { w: c2, fill: BOARD, borders: chalkBorder }),
    cell(p(ck(s.stoerung, { size: 18 }), { after: 0 }), { w: c3, fill: BOARD, borders: chalkBorder }),
  ], 1150)),
]);

const boardW1 = 7400, boardW2 = TW - boardW1;
const board = table([TW], [
  row([cell([
    p([phase(1), ck("„Was erlauben Strunz?“ – Trapattonis Wutrede (10. März 1998)", { font: "Cambria", size: 32, bold: true })], { after: 60 }),
    p([phase(2), ck("Hypothese: ", { bold: true, size: 22, color: "F5D76E" }), ck("Trapattoni verletzt alle vier Seiten einer Nachricht. → Beweist es mit Zitaten!", { size: 22 })], { after: 0 }),
  ], { w: TW, fill: BOARD, borders: allBorders(none), m: 140, ml: 200 })]),
  row([cell([earTable], { w: TW, fill: BOARD, borders: allBorders(none), m: 60, ml: 200 })]),
  row([cell([
    table([boardW1 - 400, boardW2], [row([
      cell([
        p([phase(5), ck("Spieler öffentlich an den Pranger stellen – wozu führt das?", { bold: true, size: 20 })], { after: 40 }),
        p(ck("Demütigung → Spieler machen dicht → Vertrauen weg → Mannschaft gegen Trainer", { size: 18 }), { after: 20 }),
        p(ck("Medien feiern die Sprache, nicht die Kritik → Sache geht verloren", { size: 18 }), { after: 20 }),
        p(ck("⇒ Der Trainer verliert: Autorität, Rückhalt, am Ende das Amt.", { size: 19, bold: true, color: "F5D76E" }), { after: 0 }),
      ], { w: boardW1 - 400, fill: BOARD2, borders: chalkBorder, m: 100 }),
      cell([
        p([phase(6), ck("Jetzt bringt ihr das in Ordnung!", { bold: true, size: 20, color: "F5D76E" })], { after: 40 }),
        p(ck("Schreibt als Trapattoni einen Brief an die Spieler.", { size: 18 }), { after: 20 }),
        p(ck("• alle vier Seiten „heilen“  • Ich-Botschaft  • Empathie  • konkrete Bitte  • Metakommunikation", { size: 17 }), { after: 20 }),
        p(ck("Der Brief ersetzt den angekündigten Test.", { size: 17, italics: true, color: "C9C3A8" }), { after: 0 }),
      ], { w: boardW2, fill: BOARD2, borders: chalkBorder, m: 100 }),
    ])]),
  ], { w: TW, fill: BOARD, borders: allBorders(none), m: 100, ml: 200 })]),
]);

const page1 = [
  kicker("Tafelscript · Deutsch 11 · Kommunikationsstörungen im Sport"),
  h1("Tafelbild im Endstand"),
  p([t("Die Ziffern ① – ⑥ zeigen, ", { size: 20, color: MUTED }), t("wann", { size: 20, color: MUTED, bold: true }), t(" was an die Tafel kommt (Ablauf auf der nächsten Seite). Farbige Kreide: Sache blau · Selbstoffenbarung grün · Beziehung gelb · Appell rot.", { size: 20, color: MUTED })], { after: 100 }),
  board,
];

// ================= Seite 2: Schritt für Schritt =================
const sW = [700, 2600, 4400, 4538, 2900];
const schritte = [
  ["①", "Einstieg: Präsentation", "Überschrift erst nach dem Video aufdecken – vorher nur: „Ein Trainer – ein Fehltritt“.",
    "„Ich präsentiere euch jemanden aus dem Sport, bei dem es zu Kommunikationsstörungen kam – ein Trainer, der sich einen richtigen Fehltritt geleistet hat.“ Foto zeigen, noch keinen Namen.",
    "stellen Vermutungen an (Wer? Was ist passiert?)"],
  ["②", "Hypothese + Beobachtungsauftrag", "Hypothese (gelb unterstrichen). Leere Tabelle: Spalte 1 mit den vier Seiten in Kreidefarben, Überschriften der Spalten 2 und 3.",
    "„Das ist meine Hypothese: Trapattoni verletzt alle vier Seiten. Ihr sollt sie beweisen – mit Zitaten.“ Auftrag vor dem Video: „Überlegt beim Zuschauen: Welche Seite wird wie verletzt?“",
    "notieren beim Video Stichworte und Zitate zu den vier Seiten"],
  ["③", "Video + Belege sammeln", "Spalte 2: Inhalt in Kreideweiß, Zitate in der Farbe der jeweiligen Seite. Überschrift jetzt vollständig anschreiben.",
    "Video (ca. 3,5 Min.). Danach Transkript austeilen – die Zitate sind im Video schwer mitzuschreiben. Impuls: „Wo verletzt er die Sachebene? Belegt es!“ – Seite für Seite, Schüler nehmen sich gegenseitig dran.",
    "belegen jede Seite mit einem Zitat aus dem Transkript"],
  ["④", "Störung benennen", "Spalte 3 erst, wenn alle vier Zeilen belegt sind.",
    "„Wir haben gesehen, WAS er sendet. Wo genau liegt jeweils die Störung?“ Tipp: bei Appell fragen „An wen richtet er sich – und wer hört ihn?“",
    "formulieren die Störung je Seite"],
  ["⑤", "Fazit: öffentlicher Pranger", "Kasten unten links, als Pfeilkette.",
    "„Spieler öffentlich an den Pranger zu stellen – wozu führt das?“ Ergebnis zuspitzen: Der Trainer verliert.",
    "entwickeln Folgen (Spieler, Mannschaft, Medien, Trainer)"],
  ["⑥", "Transfer: Brief (letzte 20–25 Min.)", "Kasten unten rechts mit Auftrag und Kriterien.",
    "„Jetzt bringt ihr das in Ordnung.“ Hinweis: Der Brief ersetzt den angekündigten Test und wird von allen geschrieben. Kriterien siehe Seite 3.",
    "schreiben einzeln einen Brief als Trapattoni an die Spieler"],
];
const sRows = [row([hdr("", sW[0]), hdr("Phase", sW[1]), hdr("Was an die Tafel kommt", sW[2]), hdr("Ihre Moderation / Impuls", sW[3]), hdr("Klasse", sW[4])], 420)];
schritte.forEach(([n, ph, tafel, imp, sus]) => sRows.push(row([
  cell(p(t(n, { size: 30, bold: true, color: RED }), { align: AlignmentType.CENTER, after: 0 }), { w: sW[0], fill: LIGHT, valign: VerticalAlign.CENTER }),
  cell(p(t(ph, { bold: true, size: 20, color: NAVY }), { after: 0 }), { w: sW[1], fill: LIGHT }),
  cell(p(t(tafel, { size: 19 }), { after: 0 }), { w: sW[2] }),
  cell(p(t(imp, { size: 19 }), { after: 0 }), { w: sW[3], fill: LIGHT2 }),
  cell(p(t(sus, { size: 19, color: MUTED }), { after: 0 }), { w: sW[4] }),
], 700)));

const page2 = [
  kicker("Tafelscript · Ablauf", true),
  h1("Wann kommt was an die Tafel?"),
  table(sW, sRows),
  gap(120),
  infoBox([
    [t("Hintergrund: ", { bold: true, color: NAVY, size: 20 }), t("Pressekonferenz des FC Bayern München am 10. März 1998 nach einer 0:1-Niederlage bei Schalke 04. Trapattoni kritisiert namentlich Thomas Strunz, Mario Basler und Mehmet Scholl. Am Saisonende verlässt er den Verein.", { size: 20 })],
    [t("Achtung beim Sachinhalt: ", { bold: true, color: RED, size: 20 }), t("Die Störung ist nicht Trapattonis Deutsch an sich, sondern dass die Aufmerksamkeit auf die Form springt („Flasche leer“) und der Inhalt verloren geht. Das Lachen der Klasse beim Video kann man genau so aufgreifen: „Worüber lacht ihr – und was hat er eigentlich gesagt?“", { size: 20 })],
    [t("Zitate und Eckdaten: ", { bold: true, color: NAVY, size: 20 }), t("Die Zitate folgen verbreiteten Transkripten der Rede. Bitte vor der Stunde mit dem Video abgleichen, das Sie zeigen – Schreibweisen weichen je nach Quelle leicht ab. Auch Datum und Spielergebnis kurz gegenprüfen.", { size: 20 })],
  ]),
];

// ================= Seite 3: Brief als Ersatz für den Test =================
const kW = [3400, 7738, 2000, 2000];
const kriterien = [
  ["Sachinhalt klären", "Benennt die Kritik sachlich und konkret (z. B. Defensivarbeit, Einsatz im Training) – ohne Beleidigung und ohne „Flasche leer“.", "4"],
  ["Selbstoffenbarung ehrlich", "Zeigt Frust und Enttäuschung als Ich-Botschaft („Ich war enttäuscht, als …“) statt als Wutausbruch.", "4"],
  ["Beziehung reparieren", "Entschuldigt sich für die öffentliche Kritik, zeigt Respekt, geht auf die Sicht der Spieler ein (Empathie).", "4"],
  ["Appell an die Richtigen", "Richtet eine konkrete, erfüllbare Bitte direkt an die Spieler (wer, was, bis wann) – z. B. Gespräch unter vier Augen, Einsatz am Samstag.", "4"],
  ["Metakommunikation", "Spricht über die Pressekonferenz selbst: was schiefgelaufen ist und wie man künftig miteinander redet (intern statt über die Medien).", "2"],
  ["Form", "Briefform (Ort, Datum, Anrede, Schluss), passender Ton, sprachlich korrekt.", "2"],
];
const kRows = [row([hdr("Kriterium", kW[0]), hdr("Erwartung", kW[1]), hdr("Punkte", kW[2]), hdr("erreicht", kW[3])], 420)];
kriterien.forEach(([a, b, c]) => kRows.push(row([
  cell(p(t(a, { bold: true, size: 20, color: NAVY }), { after: 0 }), { w: kW[0], fill: LIGHT }),
  cell(p(t(b, { size: 19 }), { after: 0 }), { w: kW[1] }),
  cell(p(t(c, { bold: true, size: 22 }), { align: AlignmentType.CENTER, after: 0 }), { w: kW[2], valign: VerticalAlign.CENTER }),
  cell(p(t("")), { w: kW[3] }),
], 560)));
kRows.push(row([
  cell(p(t("Gesamt", { bold: true, size: 20, color: NAVY }), { after: 0 }), { w: kW[0], fill: LIGHT2 }),
  cell(p(t("")), { w: kW[1], fill: LIGHT2 }),
  cell(p(t("20", { bold: true, size: 22 }), { align: AlignmentType.CENTER, after: 0 }), { w: kW[2], fill: LIGHT2 }),
  cell(p(t("")), { w: kW[3], fill: LIGHT2 }),
], 460));

const page3 = [
  kicker("Tafelscript · Transfer ⑥", true),
  h1("„Jetzt bringt ihr das in Ordnung“ – Brief an die Spieler"),
  infoBox([
    [t("Auftrag (Tafel / mündlich): ", { bold: true, color: NAVY }), t("Schreibe als Giovanni Trapattoni einen Brief an Thomas Strunz, Mario Basler und Mehmet Scholl. Bringe die Missverständnisse aus der Pressekonferenz in Ordnung – so, dass alle vier Seiten der Nachricht wieder stimmen. Nutze die Techniken aus der letzten Stunde.", { size: 21 })],
    [t("Leistungsnachweis: ", { bold: true, color: RED }), t("Der Brief ersetzt den angekündigten Test über die Kommunikationsmodelle und wird von allen geschrieben und bewertet. Er prüft dasselbe – nur angewendet statt beschrieben.", { size: 21 })],
  ]),
  h2("✓", "Bewertungskriterien"),
  table(kW, kRows),
  h2("✎", "Mögliche Formulierungen (Erwartungshorizont)"),
  p([t("Metakommunikation: ", { bold: true, size: 20, color: NAVY }), t("„Liebe Spieler, auf der Pressekonferenz am Dienstag habe ich Dinge über euch gesagt, die ich euch persönlich hätte sagen müssen.“", { size: 20, italics: true })], { after: 40 }),
  p([t("Ich-Botschaft: ", { bold: true, size: 20, color: NAVY }), t("„Ich war nach dem Spiel in Schalke sehr enttäuscht und habe mich hilflos gefühlt, weil ich das Gefühl hatte, dass wir nicht als Mannschaft kämpfen.“", { size: 20, italics: true })], { after: 40 }),
  p([t("Empathie: ", { bold: true, size: 20, color: NAVY }), t("„Ich kann verstehen, dass ihr euch vor ganz Deutschland bloßgestellt gefühlt habt. Das tut mir leid.“", { size: 20, italics: true })], { after: 40 }),
  p([t("Bitte: ", { bold: true, size: 20, color: NAVY }), t("„Ich möchte mit jedem von euch vor dem Training am Donnerstag sprechen. Und ich bitte euch, am Samstag in der Defensive mitzuarbeiten – für die Mannschaft.“", { size: 20, italics: true })], { after: 0 }),
];

// ================= Dokument =================
const pageProps = { page: { size: { width: 11906, height: 16838, orientation: PageOrientation.LANDSCAPE }, margin: { top: 700, bottom: 700, left: 850, right: 850, footer: 350 } } };
const footer = new Footer({ children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ font: "Calibri", size: 16, color: MUTED, children: ["Tafelscript · Trapattonis Wutrede · Seite ", PageNumber.CURRENT] })] })] });
const doc = new Document({
  styles: { default: { document: { run: { font: "Calibri", size: 22 } } } },
  sections: [{ properties: pageProps, footers: { default: footer }, children: [...page1, ...page2, ...page3] }],
});
Packer.toBuffer(doc).then((b) => fs.writeFileSync("Tafelscript_Trapattoni_Wutrede.docx", b));
