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
// Sätze zum Zuordnen (Aufgabe 2b)
const SAETZE = [
  ["A", "keine Steuern zahlen"],
  ["B", "eigene Gerichte haben"],
  ["C", "das alleinige Jagdrecht haben"],
  ["D", "für die Armen sorgen"],
  ["E", "verwaltende Aufgaben im Staat übernehmen"],
  ["F", "Steuern und hohe Abgaben zahlen"],
  ["G", "im Krieg Soldaten stellen"],
  ["H", "den König wählen"],
];
// Wahr oder falsch (Aufgabe 5)
const WF = [
  ["Ludwig XIV. verzichtete auf Söldner und baute ein stehendes Heer auf.", true, ""],
  ["Der Klerus musste den Kirchenzehnt an den Dritten Stand (Bauern und Bürger) zahlen.", false, "Umgekehrt: Der Dritte Stand (Bauern und Bürger) musste dem Klerus den Kirchenzehnt (10 % vom Einkommen) zahlen."],
  ["Der Adel hatte das alleinige Jagdrecht.", true, ""],
  ["Ludwig XIV. erlaubte den Protestanten die freie Religionsausübung.", false, "1685 beendete er die religiöse Toleranz: Wie nur einen König sollte es auch nur einen Glauben geben (Gottesgnadentum)."],
];
const checkbox = (w) => cell(p(t("☐", { size: 28 }), { after: 0, align: AlignmentType.CENTER }), { w });
const th = (txt, w, fill = NAVY) => cell(p(t(txt, { bold: true, size: 22, color: "FFFFFF" }), { after: 0, align: AlignmentType.CENTER }), { w, fill });

// ---------- Test (Schüler) ----------
function testDoc() {
  const c = [];
  c.push(p(t("ENTWURF · GESCHICHTE · KLASSE 8", { size: 18, bold: true, color: MUTED }), { after: 20 }));
  c.push(p(t("Test: Ständegesellschaft und Absolutismus", { font: "Cambria", size: 38, bold: true, color: NAVY }), { after: 120 }));
  c.push(table([4800, 2400, 3006], [row([
    cell(p(t("Name:", { size: 22, color: MUTED }), { after: 0 }), { w: 4800, m: 160 }),
    cell(p(t("Datum:", { size: 22, color: MUTED }), { after: 0 }), { w: 2400, m: 160 }),
    cell(p(t("Punkte:        / 30 (+1)", { size: 22, color: MUTED }), { after: 0 }), { w: 3006, m: 160 }),
  ])]));
  c.push(p(t("Arbeitszeit: 35 Minuten. Schreibe in ganzen Sätzen und verwende Fachbegriffe.", { size: 22, italics: true, color: MUTED }), { before: 80, after: 40 }));

  // 1
  c.push(aufgabe(1, "Absolutismus", 3));
  c.push(p([t("a) ", { bold: true }), t("Erkläre den Begriff „Absolutismus“.")], { after: 40, keepNext: true }), ...lines(3));
  c.push(p([t("b) ", { bold: true }), t("Passt die Aussage zum Absolutismus? Kreuze an und begründe.")], { before: 100, after: 60, keepNext: true }));
  c.push(table([4400, 1100, 1300, 3406], [
    row([th("Aussage", 4400), th("passt", 1100), th("passt nicht", 1300), th("Begründung", 3406)]),
    row([cell(p(t("Ein Parlament muss jeder Steuererhöhung des Königs zustimmen.", { size: 22 }), { after: 0 }), { w: 4400 }), checkbox(1100), checkbox(1300), cell(p(t(""), { after: 0 }), { w: 3406 })], 1100),
  ]));

  // 2
  c.push(aufgabe(2, "Die drei Stände", 7));
  c.push(p([t("a) ", { bold: true }), t("Trage die ungefähre Anzahl der Menschen in jedem Stand in die Tabelle ein.")], { after: 40, keepNext: true }));
  c.push(p([t("b) ", { bold: true }), t("Ordne die Sätze zu: Trage die Buchstaben in das passende Kästchen ein. Ein Satz kann zu mehr als einem Stand passen, ein Satz passt zu keinem.")], { after: 80, keepNext: true }));
  c.push(table([W], [row([cell(SAETZE.map(([l, s]) => p([t(l + "  ", { bold: true, color: RED, size: 22 }), t(s, { size: 22 })], { after: 30 })), { w: W, fill: LIGHT, valign: VerticalAlign.TOP, m: 100 })])]));
  c.push(p(t(""), { after: 60 }));
  const wd = [2900, 2000, 2650, 2656];
  c.push(table(wd, [
    row([th("Stand", wd[0]), th("Anzahl", wd[1]), th("Vorrechte", wd[2]), th("Pflichten", wd[3])]),
    ...["Klerus", "Adel", "Dritter Stand (Bauern und Bürger)"].map((s, i) => row([cell(p(t(s, { bold: true, size: 22 }), { after: 0 }), { w: wd[0], fill: LIGHT }), cell(p(t(""), { after: 0 }), { w: wd[1] }), cell(p(t(i === 2 ? "keine" : "", { size: 22, italics: true, color: MUTED }), { after: 0, align: AlignmentType.CENTER }), { w: wd[2] }), cell(p(t(""), { after: 0 }), { w: wd[3] })], 900)),
  ]));

  // 3
  c.push(aufgabe(3, "Mitbestimmung", 2));
  c.push(p(t("Nenne, wer in der Ständegesellschaft politisch mitbestimmen durfte und wer davon ausgeschlossen war."), { after: 60, keepNext: true }), ...lines(3));

  // 4
  c.push(aufgabe(4, "Die fünf Säulen", 8));
  c.push(p([t("a) ", { bold: true }), t("Nenne die fünf Säulen, auf die Ludwig XIV. seine Macht stützte.")], { after: 60, keepNext: true }));
  c.push(table([W], [row([cell([1, 2, 3, 4, 5].map((n) => p([t(`${n}  `, { bold: true, color: RED }), t("_______________________________________________________")], { after: 60 })), { w: W, borders: all(none), valign: VerticalAlign.TOP })])]));
  c.push(p([t("b) ", { bold: true }), t("Wähle eine Säule aus und erkläre, wie der König damit seine Macht sichert.")], { before: 100, after: 40, keepNext: true }), ...lines(5));

  // 5 Wahr oder falsch
  c.push(aufgabe(5, "Wahr oder falsch?", 6));
  c.push(p(t("Kreuze an. Bei „falsch“ musst du die Aussage begründen oder verbessern."), { after: 60, keepNext: true }));
  c.push(table([4300, 1000, 1150, 3756], [
    row([th("Aussage", 4300), th("wahr", 1000), th("falsch", 1150), th("Begründung (nur bei falsch)", 3756)]),
    ...WF.map(([a]) => row([cell(p(t(a, { size: 22 }), { after: 0 }), { w: 4300 }), checkbox(1000), checkbox(1150), cell(p(t(""), { after: 0 }), { w: 3756 })], 1050)),
  ]));

  // 6 Stellung nehmen
  c.push(aufgabe(6, "Stellung nehmen", 4));
  c.push(p([t("Ein Mitschüler sagt: "), t("„Im Absolutismus war es egal, in welchen Stand man geboren wurde.“", { italics: true })], { after: 40, keepNext: true }));
  c.push(p(t("Nimm Stellung zu dieser Aussage und begründe mit mindestens zwei Argumenten."), { after: 60, keepNext: true }), ...lines(6));

  // Zusatz
  c.push(p([t("Zusatzaufgabe  ", { font: "Cambria", size: 28, bold: true, color: "C98A1E" }), t("(1 Zusatzpunkt)", { size: 22, color: MUTED })], { before: 240, after: 100, keepNext: true }));
  c.push(p(t("Nenne den Finanzminister, der in Frankreich den Merkantilismus einführte."), { after: 60, keepNext: true }), ...lines(1));

  return new Document({ styles: { default: { document: { run: { font: "Calibri", size: 24 } } } }, sections: [{ properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 900, bottom: 800, left: 850, right: 850 } } }, footers: { default: footer("Geschichte 8 · Test Entwurf") }, children: c }] });
}

// ---------- Erwartungshorizont ----------
function loesungDoc() {
  const c = [];
  const S = (s, o = {}) => t(s, Object.assign({ size: 22 }, o));
  const pt = (n) => t(`  [${n} P]`, { bold: true, color: RED, size: 20 });
  const bullet = (runs) => p([S("• "), ...[].concat(runs)], { after: 50, indent: { left: 240 } });
  const sub = (x) => p(S(x, { size: 20, color: "333333" }), { after: 30, indent: { left: 600 } });
  c.push(p(t("ENTWURF · FÜR DIE LEHRKRAFT", { size: 18, bold: true, color: MUTED }), { after: 20 }));
  c.push(p(t("Erwartungshorizont zum Test", { font: "Cambria", size: 38, bold: true, color: NAVY }), { after: 120 }));
  c.push(p(S("Gesamtpunktzahl: 30 Punkte plus 1 Zusatzpunkt. Formulierungen dürfen abweichen, entscheidend ist der Inhalt. Folgefehler werden nicht doppelt bestraft. Die Angaben stammen aus den Gruppenpuzzle-Texten (Stände, Säulen) und den Merksätzen der Wiederholungsstunde."), { after: 80 }));

  c.push(aufgabe(1, "Absolutismus", 3));
  c.push(bullet([S("a) Herrschaftsform in Europa im 17. und 18. Jahrhundert, in der ein einzelner Monarch (König oder Fürst) uneingeschränkt (absolut) herrscht."), pt("2")]));
  c.push(sub("1 P für „Herrschaftsform / Monarch herrscht allein“, 1 P für „uneingeschränkt / niemand kann ihn stoppen“ (oder 17./18. Jahrhundert)."));
  c.push(bullet([S("b) Passt nicht: Ein Parlament könnte den König stoppen, seine Macht wäre begrenzt (Entscheidung und Begründung)."), pt("1")]));

  c.push(aufgabe(2, "Die drei Stände", 7));
  c.push(bullet([S("a) Klerus ca. 150 000 · Adel ca. 500 000 · Dritter Stand (Bauern und Bürger) ca. 20 000 000 (Größenordnung genügt)."), pt("3 (je 1)")]));
  c.push(bullet([S("b) Je 2 richtig eingetragene Buchstaben ergeben 1 Punkt (insgesamt 8 richtige Eintragungen). Falsch eingetragene Buchstaben heben je eine richtige Eintragung auf."), pt("4")]));
  const wd = [2900, 2300, 2500, 2506];
  c.push(table(wd, [
    row([th("Stand", wd[0], "3D4A3F"), th("Anzahl", wd[1], "3D4A3F"), th("Vorrechte", wd[2], "3D4A3F"), th("Pflichten", wd[3], "3D4A3F")]),
    ...[["Klerus", "ca. 150 000", "A, B", "D"], ["Adel", "ca. 500 000", "A, C", "E"], ["Dritter Stand (Bauern und Bürger)", "ca. 20 000 000", "keine (vorgegeben)", "F, G"]].map((r) => row(r.map((v, i) => cell(p(t(v, { size: 22, bold: i === 0 }), { after: 0 }), { w: wd[i], fill: i === 0 ? LIGHT : undefined })), 500)),
  ]));
  c.push(p(S("Satz H („den König wählen“) passt zu keinem Stand.", { color: MUTED }), { before: 60, after: 60 }));

  c.push(aufgabe(3, "Mitbestimmung", 2));
  c.push(bullet([S("Der Adel bestimmte das politische Geschehen mit (der Text nennt den Klerus nicht, deshalb nur der Adel erforderlich)."), pt("1")]));
  c.push(bullet([S("Der Dritte Stand (Bauern und Bürger) war von der politischen Mitbestimmung ausgeschlossen."), pt("1")]));

  c.push(aufgabe(4, "Die fünf Säulen", 8));
  c.push(bullet([S("a) Verwaltung und Justiz · Adel und Hof (Versailles) · Wirtschaft (Merkantilismus) · Religion (Gottesgnadentum) · Heer (stehendes Heer)."), pt("5 (je 1)")]));
  c.push(bullet([S("b) Je nach gewählter Säule, 3 P: 1 P für ein zutreffendes Merkmal, 2 P für die Erklärung der Machtsicherung (Wirkung begründen). Beispiele:"), pt("3")]));
  SAEULEN.forEach(([h, b]) => c.push(p([S(h + ": ", { bold: true, size: 20 }), S(b, { size: 20, color: "333333" })], { after: 40, indent: { left: 600 } })));

  c.push(aufgabe(5, "Wahr oder falsch?", 6));
  c.push(p(S("Wahre Aussagen: 1 P für das richtige Kreuz. Falsche Aussagen: 1 P für das richtige Kreuz, 1 P für die zutreffende Begründung oder Verbesserung.", { color: MUTED }), { after: 60 }));
  WF.forEach(([a, w, b], i) => c.push(bullet([S(`${i + 1}  ${a}  `, { size: 21 }), t(w ? "→ wahr" : "→ falsch", { bold: true, color: w ? "2E9E6B" : RED, size: 21 }), pt(w ? "1" : "2")]), ...(w ? [] : [sub("Begründung: " + b)])));

  c.push(aufgabe(6, "Stellung nehmen", 4));
  c.push(bullet([S("Klare Position: Die Aussage stimmt nicht, denn der Stand entschied über Rechte, Pflichten und Lebenschancen."), pt("1")]));
  c.push(bullet([S("Zwei zutreffende Argumente (je 1 P), zum Beispiel:"), pt("2")]));
  [
    "Vorrechte: Klerus und Adel zahlten keine Steuern (Adel zusätzlich alleiniges Jagdrecht und ranghohe Posten in der Armee), der Dritte Stand (Bauern und Bürger) zahlte Steuern, hohe Abgaben und Pachtgebühren.",
    "Mitbestimmung: Der Dritte Stand war von der politischen Mitbestimmung ausgeschlossen, der Adel bestimmte das politische Geschehen mit.",
    "Geburt: Man wurde in den Stand hineingeboren, für den Dritten Stand gab es kaum Möglichkeit aufzusteigen. In den Klerus kam man nur durch Berufung der Kirche.",
    "Bildung: Klerus und Adel waren meist sehr gebildet, vom Dritten Stand konnten nur wenige lesen und schreiben.",
    "Anzahl: Wenige Privilegierte (ca. 650 000) standen ca. 20 000 000 Menschen gegenüber, die die Last trugen.",
  ].forEach(sub);
  c.push(bullet([S("Schlüssige Darstellung mit Fachbegriffen (z. B. Vorrechte, Steuern, Mitbestimmung) und „weil“-Verknüpfung."), pt("1")]));

  c.push(p([t("Zusatzaufgabe  ", { font: "Cambria", size: 28, bold: true, color: "C98A1E" }), t("(1 Zusatzpunkt)", { size: 22, color: MUTED })], { before: 240, after: 100, keepNext: true }));
  c.push(bullet([S("Colbert (Finanzminister Ludwigs XIV., führte den Merkantilismus ein)."), pt("1")]));
  c.push(p(S("Der Zusatzpunkt wird zur Gesamtpunktzahl addiert. Maximal erreichbar: 31 Punkte.", { color: MUTED }), { before: 40, after: 100 }));

  c.push(p(t("Notenschlüssel (30 Punkte)", { font: "Cambria", size: 26, bold: true, color: NAVY }), { before: 200, after: 80, keepNext: true }));
  const nw = [1500, 2600, 2600];
  c.push(table(nw, [
    row([th("Note", nw[0]), th("Punkte", nw[1]), th("Prozent", nw[2])]),
    ...[["1", "30 – 28", "ab 92 %"], ["2", "27 – 25", "ab 81 %"], ["3", "24 – 21", "ab 67 %"], ["4", "20 – 15", "ab 50 %"], ["5", "14 – 9", "ab 30 %"], ["6", "8 – 0", "unter 30 %"]].map((r) => row(r.map((v, i) => cell(p(t(v, { size: 22, bold: i === 0 }), { after: 0, align: AlignmentType.CENTER }), { w: nw[i], fill: i === 0 ? LIGHT : undefined })))),
  ]));
  c.push(p(S("Hinweis: Der Schlüssel ist ein Vorschlag (üblicher Prozentschlüssel) und kann an Ihre schulinterne Vorgabe angepasst werden.", { italics: true, color: MUTED }), { before: 80, after: 0 }));

  return new Document({ styles: { default: { document: { run: { font: "Calibri", size: 22 } } } }, sections: [{ properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 900, bottom: 800, left: 850, right: 850 } } }, footers: { default: footer("Geschichte 8 · Erwartungshorizont Test") }, children: c }] });
}

(async () => {
  fs.writeFileSync("../Test_Staendegesellschaft_Absolutismus.docx", await Packer.toBuffer(testDoc()));
  fs.writeFileSync("../Erwartungshorizont_Test.docx", await Packer.toBuffer(loesungDoc()));
  console.log("ok");
})();
