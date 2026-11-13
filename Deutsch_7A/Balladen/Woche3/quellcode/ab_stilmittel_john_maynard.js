// Arbeitsblatt „Stilmittel erkennen – John Maynard“ (Moodle, Einzelstunde, selbstständig) + Lösung + Handreichung Polnisch
// John Maynard: Theodor Fontane (1819–1898), gemeinfrei; Zitate nach LB S. 132–133 (Verszählung wie im LB).
// Druck ohne Farbflächen. Aufruf: node ab_stilmittel_john_maynard.js
const fs = require("fs");
const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, BorderStyle, AlignmentType, VerticalAlign, HeightRule, Footer, PageNumber } = require("docx");
const NAVY = "1E2761", MUTED = "5B6B8C", RED = "D9534F", VIO = "7A5BA6", SOLC = "2E7D4F", PLC = "5B4A8A", GRAY = "9AA3B5", TEAL = "2A8C8C", ORA = "C77C12";
const PW = 11906 - 1500;
const t = (text, o = {}) => new TextRun({ text, font: o.font || "Calibri", size: o.size || 22, bold: o.bold, italics: o.it, color: o.color || "000000", underline: o.u ? {} : undefined });
const p = (ch, o = {}) => new Paragraph({ children: Array.isArray(ch) ? ch : [ch], alignment: o.align, spacing: { before: o.before || 0, after: o.after ?? 60, line: o.line }, indent: o.indent, keepNext: o.keepNext });
const B = (style = BorderStyle.SINGLE, size = 6, color = "B9C6E8") => ({ style, size, color });
const borders = (b) => ({ top: b, bottom: b, left: b, right: b });
const NONE = borders(B(BorderStyle.NONE, 0, "FFFFFF"));
const FB = borders(B(BorderStyle.SINGLE, 6, GRAY));
const cell = (children, w, o = {}) => new TableCell({ children: Array.isArray(children) ? children : [children], width: { size: w, type: WidthType.DXA }, borders: o.borders || borders(B()), verticalAlign: o.valign || VerticalAlign.CENTER, margins: { top: o.m ?? 40, bottom: o.m ?? 40, left: o.ml ?? 90, right: o.ml ?? 90 } });
const table = (cols, rows) => new Table({ columnWidths: cols, width: { size: cols.reduce((a, b) => a + b, 0), type: WidthType.DXA }, rows });
const row = (cells, h, rule = HeightRule.ATLEAST) => new TableRow({ children: cells, height: h ? { value: h, rule } : undefined, cantSplit: true });
const s = (x, o = {}) => t(x, Object.assign({ size: 22 }, o));
const solR = (x) => t(x, { size: 20, it: true, bold: true, color: SOLC });
const spacer = (n = 60) => p(t("", { size: 8 }), { after: n });
const pl = (x) => p([t("PL: ", { size: 19, bold: true, color: PLC }), t(x, { size: 19, it: true, color: PLC })], { after: 40 });

// ---------- Inhalte ----------
const SM = [
  // Name, Erkennen, Beispiel (Training), Wirkung, Farbe, Polnisch
  ["Vergleich", "Zwei Dinge werden mit „wie“ oder „als“ verbunden.", "„Er sprintet wie ein Gepard.“", "macht anschaulich: Man sieht es vor sich.", NAVY, "porównanie – dwie rzeczy łączy słowo „wie“ (jak) albo „als“"],
  ["Metapher", "Ein sprachliches Bild ohne „wie“: Ein Wort wird in übertragener Bedeutung benutzt.", "„Unser Torwart ist eine Wand.“", "drückt eine Eigenschaft kurz und stark aus.", TEAL, "metafora – obraz bez słowa „wie“; słowo ma przenośne znaczenie"],
  ["Alliteration", "Benachbarte Wörter beginnen mit demselben Laut.", "„Kampf und Kraft“", "klingt einprägsam, betont die Wörter.", VIO, "aliteracja – sąsiednie słowa zaczynają się od tej samej głoski"],
  ["Anapher", "Mehrere Verse oder Sätze beginnen mit demselben Wort.", "„Wir laufen. Wir kämpfen. Wir siegen.“", "betont, steigert, schafft Rhythmus.", ORA, "anafora – kilka wersów albo zdań zaczyna się tym samym słowem"],
  ["Ellipse", "Ein unvollständiger Satz: Meist fehlt das Verb.", "„Abpfiff. Sieg. Jubel.“", "wirkt knapp, schnell und dramatisch.", MUTED, "elipsa – zdanie niepełne, najczęściej brakuje czasownika"],
  ["Ausruf", "Ein kurzer Satz mit Ausrufezeichen: ein Befehl oder ein starkes Gefühl.", "„Lauf!“ – „Was für ein Tor!“", "zeigt Aufregung, Gefahr oder Freude.", RED, "wykrzyknienie – krótkie zdanie z wykrzyknikiem: rozkaz albo silne uczucie"],
];
const UE1 = [ // Satz, Lösung
  ["Der Ball fliegt wie ein Pfeil ins Tor.", "Vergleich"],
  ["Weiter, gib alles!", "Ausruf"],
  ["Maxi macht mächtig Muskeln.", "Alliteration"],
  ["Letzte Runde. Volle Kraft. Ziel.", "Ellipse"],
  ["Unser Kapitän ist der Motor der Mannschaft.", "Metapher"],
  ["Wir trainieren hart. Wir trainieren lang. Wir trainieren jeden Tag.", "Anapher"],
];
const UE2 = [ // Stilmittel, Suchhilfe, Vers, Zitat, Wirkung
  ["Metapher", "V. 8", "V. 8", "„Die ‚Schwalbe‘ fliegt über den Eriesee“", "Das Schiff ist schnell und leicht wie ein Vogel; die Fahrt wirkt fröhlich."],
  ["Vergleich", "V. 9–10", "V. 9", "„Gischt schäumt um den Bug wie Flocken von Schnee“", "Man sieht die weiße Gischt vor sich; anschaulich."],
  ["Alliteration", "V. 10–12", "V. 11", "„frei und froh“", "betont die unbeschwerte, gute Stimmung vor dem Unglück."],
  ["Ausruf", "V. 18–20", "V. 20", "„Feuer!“", "plötzliche Gefahr, Schreck und Aufregung."],
  ["Anapher", "V. 24–27", "V. 25–26 (auch 27)", "„Am Bugspriet stehn sie … / Am Bugspriet vorn …“", "betont, dass sich alle vorne zusammendrängen; Enge und Angst."],
  ["Ellipse", "V. 46–47", "V. 46 (auch 47)", "„Das Schiff geborsten. Das Feuer verschwelt.“", "knapp und dramatisch, wie ein Schock nach der Rettung."],
];
const WIRK = ["… macht anschaulich, wie …", "… betont, dass …", "… zeigt die Aufregung / die Angst …", "… steigert die Spannung, weil …", "… wirkt knapp und dramatisch, weil …"];

// ---------- Bausteine ----------
const IW = PW - 280;
function kasten(title, col, children, o = {}) {
  const head = p([t(title, { font: "Cambria", size: 27, bold: true, color: col }), ...(o.meta ? [t("     " + o.meta, { size: 19, it: true, color: MUTED })] : [])], { after: 60 });
  return [table([PW], [row([cell([head, ...children], PW, { borders: borders(B(o.dashed ? BorderStyle.DASHED : BorderStyle.SINGLE, 14, col)), m: 100, ml: 140, valign: VerticalAlign.TOP })])]), spacer(o.after ?? 120)];
}
const steps = (arr, col = NAVY) => arr.map((x, i) => p([t(`${i + 1}  `, { bold: true, size: 22, color: col }), s(x)], { after: 30, indent: { left: 300, hanging: 300 } }));
const zeit = (x) => p([t("Zeit: ", { bold: true, size: 19, color: MUTED }), t(x + "      ", { size: 19, color: MUTED }), t("Fertig? Stift hinlegen und noch einmal durchlesen.", { size: 19, it: true, color: MUTED })], { before: 60, after: 0 });

function blatt(mode) { // "s" Schüler, "l" Lösung
  const sol = mode === "l";
  const d = [];
  // Kopf
  d.push(table([PW - 3200, 3200], [row([
    cell([
      p(t("DEUTSCH · KLASSE 7A · MOODLE", { size: 17, bold: true, color: MUTED }), { after: 10 }),
      p(t("Stilmittel erkennen", { font: "Cambria", size: 38, bold: true, color: NAVY }), { after: 0 }),
      p(t("Wie Fontane „John Maynard“ spannend macht (LB S. 132–133)", { font: "Cambria", size: 22, it: true, color: MUTED }), { after: 0 }),
    ], PW - 3200, { borders: NONE, m: 0, ml: 0, valign: VerticalAlign.TOP }),
    cell([p(s(sol ? "LÖSUNG" : "Name:", { bold: sol, color: sol ? SOLC : "000000" }), { after: 140 }), p(s(sol ? "zur Selbstkontrolle" : "Datum:", { color: sol ? SOLC : "000000", size: sol ? 19 : 22 }), { after: 0 })], 3200, { borders: borders(B(BorderStyle.SINGLE, 10, NAVY)), m: 80, ml: 120, valign: VerticalAlign.TOP }),
  ])]));
  d.push(spacer(60));
  d.push(p([t("Ziel: ", { bold: true, color: RED }), s("Ich kann sechs Stilmittel erkennen und sie in „John Maynard“ mit Vers und Zitat nachweisen.", { it: true })], { after: 30 }));
  d.push(p([t("So gehst du vor: ", { bold: true, color: NAVY }), s("① Lies den Merkkasten. ② Löse Ü1. ③ Löse Ü2 mit dem Lehrbuch. ④ * Ü3. ⑤ Kontrolliere mit der Lösung in Moodle und verbessere mit Grün.", { size: 21 })], { after: 100 }));

  // Merkkasten
  const MW = [1700, 3300, 2700, IW - 7700];
  const th = (x, w) => cell(p(t(x, { bold: true, size: 20, color: NAVY }), { after: 0 }), w, { borders: borders(B(BorderStyle.SINGLE, 8, NAVY)) });
  d.push(...kasten("Merkkasten: Stilmittel", RED, [
    p(s("Stilmittel sind besondere sprachliche Mittel. Dichter setzen sie ein, damit ein Text anschaulich, eindringlich oder spannend wirkt.", { size: 21 }), { after: 80 }),
    table(MW, [
      row([th("Stilmittel", MW[0]), th("So erkennst du es", MW[1]), th("Beispiel", MW[2]), th("Wirkung", MW[3])], 360),
      ...SM.map(([n, e, b, w, col]) => row([
        cell(p(t(n, { bold: true, size: 22, color: col, font: "Cambria" }), { after: 0 }), MW[0]),
        cell(p(s(e, { size: 20 }), { after: 0 }), MW[1]),
        cell(p(s(b, { size: 20, it: true }), { after: 0 }), MW[2]),
        cell(p(s(w, { size: 20 }), { after: 0 }), MW[3]),
      ], 500)),
    ]),
    p([t("Tipp: ", { bold: true, size: 20, color: RED }), s("Ein Stilmittel nachweisen heißt: Name + Vers + Zitat + Wirkung. Beispiel: Vergleich, V. 9: „… wie Flocken von Schnee“ – macht die weiße Gischt anschaulich.", { size: 20 })], { before: 80, after: 0 }),
  ]));

  // Ü1
  const W1 = [IW - 3000, 3000];
  d.push(...kasten("Ü1  Stilmittel in Trainingssätzen", NAVY, [
    ...steps(["Lies jeden Satz.", "Schreibe dahinter, welches Stilmittel er enthält. Jedes Stilmittel kommt genau einmal vor.", "Kontrolliere mit dem Merkkasten."]),
    spacer(40),
    table(W1, UE1.map(([x, l], i) => row([
      cell(p([t(`${String.fromCharCode(97 + i)}  `, { bold: true, color: NAVY }), s(x, { size: 22, it: true })], { after: 0 }), W1[0]),
      cell(p(sol ? solR(l) : t(""), { after: 0 }), W1[1], { borders: FB }),
    ], 470))),
    zeit("8 Minuten · allein"),
  ], { meta: "Erst das Werkzeug üben" }));

  // Ü2
  const W2 = [1500, 1050, 950, 3600, IW - 7100];
  const th2 = (x, w) => cell(p(t(x, { bold: true, size: 20, color: NAVY }), { after: 0 }), w, { borders: borders(B(BorderStyle.SINGLE, 8, NAVY)) });
  d.push(new Paragraph({ pageBreakBefore: true, children: [] }));
  d.push(...kasten("Ü2  Stilmittel in „John Maynard“", TEAL, [
    ...steps(["Schlage LB S. 132–133 auf.", "Suche jedes Stilmittel in den angegebenen Versen („Suche in“). Es gibt dort genau ein passendes Beispiel.", "Trage den Vers und das Zitat in Anführungszeichen ein.", "Schreibe in Stichworten, wie das Stilmittel wirkt. Die Bausteine unten helfen dir."], TEAL),
    spacer(40),
    table(W2, [
      row([th2("Stilmittel", W2[0]), th2("Suche in", W2[1]), th2("Vers", W2[2]), th2("Zitat „…“", W2[3]), th2("Wirkung", W2[4])], 360),
      ...UE2.map(([n, h, v, z, w]) => row([
        cell(p(t(n, { bold: true, size: 21, color: (SM.find((x) => x[0] === n) || [])[4] || NAVY, font: "Cambria" }), { after: 0 }), W2[0]),
        cell(p(s(h, { size: 20 }), { after: 0 }), W2[1]),
        cell(p(sol ? solR(v) : t(""), { after: 0 }), W2[2], { borders: FB }),
        cell(p(sol ? solR(z) : t(""), { after: 0 }), W2[3], { borders: FB }),
        cell(p(sol ? solR(w) : t(""), { after: 0 }), W2[4], { borders: FB }),
      ], 900)),
    ]),
    p([t("Bausteine für die Wirkung: ", { bold: true, size: 20, color: TEAL }), s(WIRK.join("  ·  "), { size: 20, it: true })], { before: 80, after: 0 }),
    zeit("20 Minuten · allein · mit dem Lehrbuch"),
  ], { meta: "Suchhilfe: Versbereich" }));

  // Ü3
  d.push(...kasten("* Ü3  Der Countdown", MUTED, [
    p(s("Lies V. 23, 29 und 39. Welches Stilmittel aus dem Merkkasten erkennst du? Begründe in einem Satz, warum Fontane es hier benutzt.", { size: 22 }), { after: 60 }),
    sol
      ? p(solR("Anapher bzw. Wiederholung: „Und noch … Minuten bis Buffalo“. Die Zeit wird immer knapper, während das Feuer schlimmer wird – so steigt die Spannung."), { after: 0 })
      : table([IW], [row([cell(p(t("")), IW, { borders: FB })], 900)]),
  ], { dashed: true, after: 0 }));
  return d;
}

function handreichung() {
  const d = [];
  const hd = (x, plx, col = NAVY) => p([t(x, { bold: true, size: 23, color: col, font: "Cambria" }), t("   " + plx, { size: 19, it: true, color: PLC })], { before: 80, after: 20 });
  d.push(p([t("Stilmittel erkennen – John Maynard", { font: "Cambria", bold: true, size: 28, color: NAVY }), t("   Środki stylistyczne – pomoc", { size: 19, it: true, color: PLC })], { after: 30 }));
  d.push(p([t("Ziel: ", { bold: true, color: RED }), s("Ich kann sechs Stilmittel erkennen und sie in „John Maynard“ mit Vers und Zitat nachweisen.", { it: true })], { after: 20 }));
  d.push(pl("Cel: Potrafię rozpoznać sześć środków stylistycznych i pokazać je w balladzie „John Maynard“ (numer wersu + cytat). Pracujesz sama w domu (Moodle): najpierw przeczytaj ramkę z wyjaśnieniami, potem Ü1, potem Ü2 z podręcznikiem (s. 132–133), * Ü3. Na końcu sprawdź z rozwiązaniem w Moodle i popraw na zielono."));
  d.push(hd("Stilmittel", "środki stylistyczne (nazwy zostają po niemiecku)"));
  const W = [2200, PW - 2200];
  d.push(table(W, SM.map(([n, , , w, col, plx]) => row([
    cell(p(t(n, { bold: true, size: 22, color: col, font: "Cambria" }), { after: 0 }), W[0]),
    cell(p(t(plx, { size: 20, color: "333333" }), { after: 0 }), W[1]),
  ], 380))));
  d.push(pl("Die Wirkung = efekt, działanie: co środek robi z tekstem (np. tekst jest bardziej obrazowy, napięcie rośnie)."));
  d.push(hd("Ü1", "zdania z treningu", NAVY));
  d.push(pl("Przeczytaj każde zdanie i napisz obok po niemiecku, jaki to środek stylistyczny. Każdy środek występuje dokładnie raz. Słowa: der Pfeil = strzała, der Motor = silnik, mächtig = bardzo, potężnie."));
  d.push(hd("Ü2", "John Maynard", TEAL));
  d.push(pl("Szukaj każdego środka tylko w podanych wersach (kolumna „Suche in“) – tam jest dokładnie jeden przykład. Wpisz numer wersu i dokładny cytat w cudzysłowie „…“. W kolumnie „Wirkung“ napisz po niemiecku w kilku słowach, jaki jest efekt. Pomogą ci gotowe początki zdań („Bausteine“). Trudne słowa: die Gischt = piana fal, der Bug = dziób statku, der Bugspriet = belka z przodu statku, geborsten = rozbity, verschwelen = dogasać."));
  d.push(hd("* Ü3", "odliczanie", MUTED));
  d.push(pl("Przeczytaj wersy 23, 29 i 39. Jaki środek stylistyczny tu widzisz? Napisz jedno zdanie po niemiecku: dlaczego Fontane go używa? (np. „…, weil …“)"));
  return d;
}

const footer = (label) => new Footer({ children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ font: "Calibri", size: 16, color: MUTED, children: [label + " · Seite ", PageNumber.CURRENT] })] })] });
const props = { page: { size: { width: 11906, height: 16838 }, margin: { top: 600, bottom: 600, left: 750, right: 750, footer: 300 } } };
const styles = { default: { document: { run: { font: "Calibri", size: 22 } } } };
const save = (name, children, label) => Packer.toBuffer(new Document({ styles, sections: [{ properties: props, footers: { default: footer(label) }, children }] })).then((buf) => fs.writeFileSync(__dirname + "/../" + name + ".docx", buf));
save("AB_Stilmittel_John_Maynard", blatt("s"), "Stilmittel · John Maynard");
save("AB_Stilmittel_John_Maynard_Loesung", blatt("l"), "Stilmittel · John Maynard · Lösung");
save("Handreichung_Polnisch_Stilmittel_John_Maynard", handreichung(), "Stilmittel · pomoc po polsku");
