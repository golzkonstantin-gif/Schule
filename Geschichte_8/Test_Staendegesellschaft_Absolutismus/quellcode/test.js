const fs = require("fs");
const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, BorderStyle, ShadingType, AlignmentType, VerticalAlign, Footer, PageNumber } = require("docx");

const NAVY = "1E2761", MUTED = "5B6B8C", LIGHT = "F4F7FD", LIGHT2 = "EAF0FA", RED = "D9534F";
const W = 10206;
const t = (text, o = {}) => new TextRun({ text, font: o.font || "Calibri", size: o.size || 24, bold: o.bold, italics: o.italics, color: o.color || "000000" });
const p = (runs, o = {}) => new Paragraph({ children: Array.isArray(runs) ? runs : [runs], alignment: o.align, spacing: { before: o.before ?? 0, after: o.after ?? 100, line: o.line }, keepNext: o.keepNext, indent: o.indent });
const none = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
const solid = (c, s = 8) => ({ style: BorderStyle.SINGLE, size: s, color: c });
const all = (b) => ({ top: b, bottom: b, left: b, right: b });
const thin = all(solid("8A96B5", 6));
function cell(children, o = {}) {
  return new TableCell({ children: Array.isArray(children) ? children : [children], width: { size: o.w, type: WidthType.DXA }, shading: o.fill ? { fill: o.fill, type: ShadingType.CLEAR, color: "auto" } : undefined, borders: o.borders || thin, verticalAlign: o.valign || VerticalAlign.CENTER, margins: { top: o.m ?? 80, bottom: o.m ?? 80, left: 120, right: 120 } });
}
const table = (widths, rows) => new Table({ width: { size: widths.reduce((a, b) => a + b, 0), type: WidthType.DXA }, columnWidths: widths, rows });
const row = (cells, h) => new TableRow({ children: cells, height: h ? { value: h, rule: "atLeast" } : undefined, cantSplit: true });
const line = () => new Paragraph({ children: [t("\t", { color: "9AA5C0" })], tabStops: [{ type: "right", position: 10200, leader: "underscore" }], spacing: { before: 0, after: 0, line: 480 }, keepNext: true });
const lines = (n) => Array.from({ length: n }, line);
const aufgabe = (nr, title, pts) => p([t(`Aufgabe ${nr}  `, { font: "Cambria", size: 28, bold: true, color: RED }), t(title, { font: "Cambria", size: 28, bold: true, color: NAVY }), t(`   (${pts} Punkte)`, { size: 22, color: MUTED })], { before: 240, after: 100, keepNext: true });
const footer = (label) => new Footer({ children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ font: "Calibri", size: 16, color: MUTED, children: [label + " · Seite ", PageNumber.CURRENT] })] })] });

// ---------- Inhalte ----------
const STAENDE = [
  ["Klerus", "ca. 150 000", "keine Steuern · eigene Gerichte · vom Wehrdienst befreit · Kirchenzehnt (10 %)", "Fürsorge für die Armen"],
  ["Adel", "ca. 500 000", "keine Steuern · alleiniges Jagdrecht · ranghohe Posten in der Armee", "verwaltende Aufgaben im Staat · ehrenhaftes Verhalten"],
  ["Dritter Stand (Bauern und Bürger)", "ca. 20 000 000", "keine Vorrechte (ausgeschlossen von der Mitbestimmung)", "Steuern und hohe Abgaben (u. a. Pachtgebühren) · im Krieg Soldaten stellen"],
];
const SAEULEN = [
  ["Verwaltung und Justiz", "Beamte (z. B. Intendanten in den Provinzen) handeln im Namen des Königs; er ist oberster Richter und kann ohne Gerichtsverfahren verhaften lassen („lettres de cachet“). → Kontrolle über das ganze Land."],
  ["Adel und Hof von Versailles", "Der Adel behält Vorrechte (Steuerfreiheit, Jagdrecht) und wird mit Festen, Theater und Geschenken am Hof gehalten. → Er leistet keinen Widerstand."],
  ["Wirtschaft (Merkantilismus)", "Der Staat lenkt die Wirtschaft: Manufakturen, hohe Zölle, Kolonien als Rohstoffquellen und Absatzmärkte; mehr verkaufen als einführen. → Geld für Hof, Beamte und Heer."],
  ["Religion (Gottesgnadentum)", "Der König ist von Gott ausgewählt; 1685 Ende der religiösen Toleranz: ein König, ein Glaube. → Herrschaft ist von Gott gewollt, Einheit im Land."],
  ["Heer (stehendes Heer)", "Berufssoldaten statt Söldner, einheitlich bewaffnet, in Kasernen, dem König unterstellt (bis 400 000 Mann). → Macht nach innen und außen durchsetzen."],
];

// ---------- Test (Schüler) ----------
function testDoc() {
  const c = [];
  c.push(p(t("ENTWURF · GESCHICHTE · KLASSE 8", { size: 18, bold: true, color: MUTED }), { after: 20 }));
  c.push(p(t("Test: Ständegesellschaft und Absolutismus", { font: "Cambria", size: 38, bold: true, color: NAVY }), { after: 120 }));
  c.push(table([5400, 2400, 2406], [row([
    cell(p(t("Name:", { size: 22, color: MUTED }), { after: 0 }), { w: 5400, m: 160 }),
    cell(p(t("Datum:", { size: 22, color: MUTED }), { after: 0 }), { w: 2400, m: 160 }),
    cell(p(t("Punkte:         / 28", { size: 22, color: MUTED }), { after: 0 }), { w: 2406, m: 160 }),
  ])]));
  c.push(p(t("Arbeitszeit: 30 Minuten. Schreibe in ganzen Sätzen und verwende Fachbegriffe.", { size: 22, italics: true, color: MUTED }), { before: 80, after: 40 }));

  // 1
  c.push(aufgabe(1, "Absolutismus", 4));
  c.push(p([t("a) ", { bold: true }), t("Erkläre den Begriff „Absolutismus“.")], { after: 40 }), ...lines(3));
  c.push(p([t("b) ", { bold: true }), t("Passt die Aussage zum Absolutismus? Kreuze an und begründe.")], { before: 100, after: 60 }));
  c.push(table([5600, 1300, 1300, 2006], [
    row(["Aussage", "passt", "passt nicht", "Begründung"].map((h, i) => cell(p(t(h, { bold: true, size: 22, color: "FFFFFF" }), { after: 0, align: AlignmentType.CENTER }), { w: [5600, 1300, 1300, 2006][i], fill: NAVY }))),
    row([cell(p(t("Der König erlässt und ändert Gesetze, ohne jemanden zu fragen.", { size: 22 }), { after: 0 }), { w: 5600 }), cell(p(t("☐", { size: 28 }), { after: 0, align: AlignmentType.CENTER }), { w: 1300 }), cell(p(t("☐", { size: 28 }), { after: 0, align: AlignmentType.CENTER }), { w: 1300 }), cell(p(t(""), { after: 0 }), { w: 2006 })], 1000),
    row([cell(p(t("Ein Parlament muss jeder Steuererhöhung des Königs zustimmen.", { size: 22 }), { after: 0 }), { w: 5600 }), cell(p(t("☐", { size: 28 }), { after: 0, align: AlignmentType.CENTER }), { w: 1300 }), cell(p(t("☐", { size: 28 }), { after: 0, align: AlignmentType.CENTER }), { w: 1300 }), cell(p(t(""), { after: 0 }), { w: 2006 })], 1000),
  ]));

  // 2
  c.push(aufgabe(2, "Die drei Stände", 9));
  c.push(p(t("Ergänze die Tabelle: ungefähre Anzahl, ein Vorrecht (oder „keine“) und eine Pflicht des Standes."), { after: 80 }));
  const wd = [2600, 1700, 3000, 2906];
  c.push(table(wd, [
    row(["Stand", "Anzahl", "ein Vorrecht", "eine Pflicht"].map((h, i) => cell(p(t(h, { bold: true, size: 22, color: "FFFFFF" }), { after: 0, align: AlignmentType.CENTER }), { w: wd[i], fill: NAVY }))),
    ...["Klerus", "Adel", "Dritter Stand (Bauern und Bürger)"].map((s) => row([cell(p(t(s, { bold: true, size: 22 }), { after: 0 }), { w: wd[0], fill: LIGHT }), cell(p(t(""), { after: 0 }), { w: wd[1] }), cell(p(t(""), { after: 0 }), { w: wd[2] }), cell(p(t(""), { after: 0 }), { w: wd[3] })], 1100)),
  ]));

  // 3
  c.push(aufgabe(3, "Mitbestimmung", 3));
  c.push(p(t("Beschreibe, wer in der Ständegesellschaft politisch mitbestimmen durfte und wer nicht. Nenne außerdem eine Folge für den Dritten Stand (Bauern und Bürger)."), { after: 60, keepNext: true }), ...lines(5));

  // 4
  c.push(aufgabe(4, "Die fünf Säulen", 8));
  c.push(p([t("a) ", { bold: true }), t("Nenne die fünf Säulen, auf die Ludwig XIV. seine Macht stützte.")], { after: 60 }));
  c.push(table([W], [row([cell([1, 2, 3, 4, 5].map((n) => p([t(`${n}  `, { bold: true, color: RED }), t("_______________________________________________________")], { after: 60 })), { w: W, borders: all(none), valign: VerticalAlign.TOP })])]));
  c.push(p([t("b) ", { bold: true }), t("Wähle eine Säule aus und erkläre, wie der König damit seine Macht sichert.")], { before: 100, after: 40 }), ...lines(5));

  // 5
  c.push(aufgabe(5, "Im Absolutismus fühle ich mich …", 4));
  c.push(p(t("Wähle einen Stand (Klerus, Adel oder Dritter Stand (Bauern und Bürger)) und beende den Satz. Gehe in deiner Begründung auf mindestens ein Vorrecht und eine Pflicht deines Standes ein."), { after: 60 }));
  c.push(p([t("Ich wähle den Stand: ", { bold: true }), t("____________________________________")], { after: 80 }));
  c.push(p(t("„Im Absolutismus fühle ich mich …", { italics: true }), { after: 20 }), ...lines(5));

  return new Document({ styles: { default: { document: { run: { font: "Calibri", size: 24 } } } }, sections: [{ properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 900, bottom: 800, left: 850, right: 850 } } }, footers: { default: footer("Geschichte 8 · Test Entwurf") }, children: c }] });
}

// ---------- Erwartungshorizont ----------
function loesungDoc() {
  const c = [];
  const S = (s, o = {}) => t(s, Object.assign({ size: 22 }, o));
  const pt = (n) => t(`  [${n} P]`, { bold: true, color: RED, size: 20 });
  const bullet = (runs) => p([S("• "), ...[].concat(runs)], { after: 50, indent: { left: 240 } });
  c.push(p(t("ENTWURF · FÜR DIE LEHRKRAFT", { size: 18, bold: true, color: MUTED }), { after: 20 }));
  c.push(p(t("Erwartungshorizont zum Test", { font: "Cambria", size: 38, bold: true, color: NAVY }), { after: 120 }));
  c.push(p(S("Gesamtpunktzahl: 28 Punkte. Formulierungen dürfen abweichen, entscheidend ist der Inhalt. Folgefehler werden nicht doppelt bestraft. Die Angaben stammen aus den Gruppenpuzzle-Texten (Stände, Säulen) und den Merksätzen der Wiederholungsstunde."), { after: 80 }));

  c.push(aufgabe(1, "Absolutismus", 4));
  c.push(bullet([S("a) Herrschaftsform in Europa im 17. und 18. Jahrhundert, in der ein einzelner Monarch (König oder Fürst) uneingeschränkt (absolut) herrscht."), pt("2")]));
  c.push(p(S("   Verteilung: 1 P für „Herrschaftsform / Monarch herrscht allein“, 1 P für „uneingeschränkt / niemand kann ihn stoppen“ (oder 17./18. Jahrhundert).", { color: MUTED }), { after: 50 }));
  c.push(bullet([S("b) Aussage 1 passt: Der König entscheidet allein, niemand kann ihn stoppen."), pt("1")]));
  c.push(bullet([S("    Aussage 2 passt nicht: Ein Parlament würde den König stoppen können, die Macht wäre begrenzt."), pt("1")]));

  c.push(aufgabe(2, "Die drei Stände", 9));
  c.push(p(S("Je 1 P pro richtiger Zelle. Anzahl: etwa richtige Größenordnung genügt. Vorrecht und Pflicht: ein zutreffendes Beispiel genügt."), { after: 60 }));
  const wd = [2300, 1500, 3300, 3106];
  c.push(table(wd, [
    row(["Stand", "Anzahl", "mögliche Vorrechte", "mögliche Pflichten"].map((h, i) => cell(p(t(h, { bold: true, size: 20, color: "FFFFFF" }), { after: 0, align: AlignmentType.CENTER }), { w: wd[i], fill: NAVY }))),
    ...STAENDE.map((r) => row(r.map((v, i) => cell(p(t(v, { size: 20, bold: i === 0 }), { after: 0 }), { w: wd[i], fill: i === 0 ? LIGHT : undefined })), 500)),
  ]));

  c.push(aufgabe(3, "Mitbestimmung", 3));
  c.push(bullet([S("Der Dritte Stand (Bauern und Bürger) war von der politischen Mitbestimmung ausgeschlossen."), pt("1")]));
  c.push(bullet([S("Der Adel bestimmte das politische Geschehen mit (der Text nennt den Klerus nicht, deshalb nur der Adel erforderlich)."), pt("1")]));
  c.push(bullet([S("Folge: Der Dritte Stand zahlte Steuern und hohe Abgaben und finanzierte den Lebensstil der beiden ersten Stände, ohne mitreden zu dürfen."), pt("1")]));

  c.push(aufgabe(4, "Die fünf Säulen", 8));
  c.push(bullet([S("a) Verwaltung und Justiz · Adel und Hof (Versailles) · Wirtschaft (Merkantilismus) · Religion (Gottesgnadentum) · Heer (stehendes Heer)."), pt("5 (je 1)")]));
  c.push(bullet([S("b) Je nach gewählter Säule, 3 P: 1 P für ein zutreffendes Merkmal, 2 P für die Erklärung der Machtsicherung (Wirkung begründen). Beispiele:"), pt("3")]));
  SAEULEN.forEach(([h, b]) => c.push(p([S(h + ": ", { bold: true, size: 20 }), S(b, { size: 20, color: "333333" })], { after: 40, indent: { left: 600 } })));

  c.push(aufgabe(5, "Im Absolutismus fühle ich mich …", 4));
  c.push(bullet([S("Stand gewählt und passendes Gefühl formuliert."), pt("1")]));
  c.push(bullet([S("Mindestens ein zutreffendes Vorrecht (oder „keine Vorrechte“ beim Dritten Stand) genannt."), pt("1")]));
  c.push(bullet([S("Mindestens eine zutreffende Pflicht genannt."), pt("1")]));
  c.push(bullet([S("Begründung ist schlüssig (Gefühl passt zu Vorrechten und Pflichten, mit „weil“ verbunden)."), pt("1")]));
  c.push(p(S("Beispiele: „… als Adeliger privilegiert, weil ich keine Steuern zahlen muss und allein jagen darf. Dafür muss ich mich ehrenhaft verhalten und Verwaltungsaufgaben übernehmen.“ · „… als Bauer ausgenutzt, weil ich Steuern, hohe Abgaben und Pachtgebühren zahlen muss und von der Mitbestimmung ausgeschlossen bin.“", { color: MUTED }), { before: 40, after: 100 }));

  c.push(p(t("Notenschlüssel (28 Punkte)", { font: "Cambria", size: 26, bold: true, color: NAVY }), { before: 200, after: 80, keepNext: true }));
  const nw = [1500, 2600, 2600];
  c.push(table(nw, [
    row(["Note", "Punkte", "Prozent"].map((h, i) => cell(p(t(h, { bold: true, size: 22, color: "FFFFFF" }), { after: 0, align: AlignmentType.CENTER }), { w: nw[i], fill: NAVY }))),
    ...[["1", "28 – 26", "ab 92 %"], ["2", "25 – 23", "ab 81 %"], ["3", "22 – 19", "ab 67 %"], ["4", "18 – 14", "ab 50 %"], ["5", "13 – 9", "ab 30 %"], ["6", "8 – 0", "unter 30 %"]].map((r) => row(r.map((v, i) => cell(p(t(v, { size: 22, bold: i === 0 }), { after: 0, align: AlignmentType.CENTER }), { w: nw[i], fill: i === 0 ? LIGHT : undefined })))),
  ]));
  c.push(p(S("Hinweis: Der Schlüssel ist ein Vorschlag (üblicher Prozentschlüssel) und kann an Ihre schulinterne Vorgabe angepasst werden.", { italics: true, color: MUTED }), { before: 80, after: 0 }));

  return new Document({ styles: { default: { document: { run: { font: "Calibri", size: 22 } } } }, sections: [{ properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 900, bottom: 800, left: 850, right: 850 } } }, footers: { default: footer("Geschichte 8 · Erwartungshorizont Test") }, children: c }] });
}

(async () => {
  fs.writeFileSync("../Test_Staendegesellschaft_Absolutismus.docx", await Packer.toBuffer(testDoc()));
  fs.writeFileSync("../Erwartungshorizont_Test.docx", await Packer.toBuffer(loesungDoc()));
  console.log("ok");
})();
