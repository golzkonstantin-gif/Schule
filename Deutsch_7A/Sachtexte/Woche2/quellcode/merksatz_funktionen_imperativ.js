// Merksatz-Schnipsel: Funktionen von Sachtexten + Imperativ – 2 pro A4 (zum Halbieren), ohne Farbflächen. PL=1 → zweisprachige Fassung für S-PL
const fs = require("fs");
const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, BorderStyle, AlignmentType, VerticalAlign } = require("docx");
const NAVY = "1E2761", MUTED = "5B6B8C", RED = "D9534F", GRAY = "9AA3B5", PLC = "5B4A8A", TEAL = "2A8C8C", VIO = "7A5BA6", YEL = "C98A1E";
const PLM = process.env.PL === "1";
const PW = 11906 - 1100;
const t = (text, o = {}) => new TextRun({ text, font: "Calibri", size: o.size || 20, bold: o.bold, italics: o.it, color: o.color || "000000", underline: o.u ? {} : undefined });
const p = (ch, o = {}) => new Paragraph({ children: Array.isArray(ch) ? ch : [ch], alignment: o.align, spacing: { before: o.before || 0, after: o.after ?? 40 }, indent: o.indent });
const pl = (x, ind = 0) => PLM ? [p(t(x, { size: 17, it: true, color: PLC }), { after: 30, indent: { left: ind } })] : [];
const b = (style, size, color) => ({ style, size, color });
const bs = (x) => ({ top: x, bottom: x, left: x, right: x });
const NB = bs(b(BorderStyle.NONE, 0, "FFFFFF"));
const cell = (ch, w, o = {}) => new TableCell({ children: ch, width: { size: w, type: WidthType.DXA }, borders: o.borders || NB, verticalAlign: VerticalAlign.TOP, margins: { top: o.m ?? 60, bottom: o.m ?? 60, left: o.ml ?? 120, right: o.ml ?? 120 } });
const table = (cols, rows) => new Table({ columnWidths: cols, width: { size: cols.reduce((a, c) => a + c, 0), type: WidthType.DXA }, rows });
const FK = [
  ["informieren", "über Tatsachen berichten, sachlich", NAVY, "informować – przekazywać fakty, rzeczowo"],
  ["unterrichten / anleiten", "erklären, wie etwas geht", TEAL, "uczyć / instruować – wyjaśniać, jak coś się robi"],
  ["aufklären", "zeigen, was richtig ist und was man vermeiden soll", VIO, "uświadamiać – pokazać, co jest dobre, a czego unikać"],
  ["auffordern", "zu etwas bewegen: du-Ansprache, Imperativ, Ausrufezeichen", RED, "wzywać, zachęcać – zwracanie się na „du“, tryb rozkazujący, wykrzyknik"],
  ["überzeugen", "mit Argumenten eine Meinung begründen", YEL, "przekonywać – uzasadniać opinię argumentami"],
];
const IM = [["du", "Verbstamm (+ e): Achte darauf! Lauf schneller!", "forma „du“: temat czasownika (+ e)"], ["du: e → i/ie", "Nimm die Hantel! Lies genau! Gib ab!", "zmiana e → i/ie jak w formie „du“ w czasie teraźniejszym"], ["ihr", "ihr-Form ohne „ihr“: Achtet darauf! Nehmt die Hantel!", "forma „ihr“ bez zaimka"], ["Sie", "Verb + Sie: Nehmen Sie Platz!", "forma grzecznościowa: czasownik + Sie"], ["sein", "Sei leise! Seid leise! Seien Sie leise!", "czasownik nieregularny"]];
function satz() {
  const inner = PW - 300;
  const frame = (x, y) => table([inner], [new TableRow({ children: [cell([p(t(x, { bold: true, size: 21 }), { after: PLM ? 20 : 0 }), ...pl(y)], inner, { borders: bs(b(BorderStyle.SINGLE, 14, RED)), m: 70 })] })]);
  return table([PW], [new TableRow({ children: [cell([
    p([t("Merksatz · ", { bold: true, color: RED, size: 19 }), t("Funktionen von Sachtexten", { bold: true, size: 25, color: NAVY })], { after: 60 }),
    frame("Sachtexte handeln von wirklichen Ereignissen und Sachverhalten. Sie können verschiedene Absichten (Funktionen) haben – oft wechselt die Absicht von Absatz zu Absatz.", "Teksty informacyjne mówią o prawdziwych wydarzeniach. Mogą mieć różne cele (funkcje) – często zmieniają się od akapitu do akapitu."),
    p(t(""), { after: 30 }),
    ...FK.flatMap(([f, e, col, plx]) => [p([t(f + ": ", { bold: true, color: col }), t(e)], { after: PLM ? 0 : 30, indent: { left: 200 } }), ...pl(plx, 200)]),
    p([t("Test: ", { bold: true, color: RED }), t("Was soll ich nach dem Lesen wissen, können oder tun?", { size: 19 })], { before: 40, after: PLM ? 0 : 100 }),
    ...pl("Test: Co mam wiedzieć, umieć albo zrobić po przeczytaniu?"),
    p([t("Merksatz · ", { bold: true, color: RED, size: 19 }), t("Der Imperativ (Befehlsform)", { bold: true, size: 25, color: NAVY })], { before: PLM ? 80 : 0, after: 60 }),
    frame("Mit dem Imperativ fordert man jemanden direkt auf. Das Verb steht am Satzanfang, bei du und ihr ohne Pronomen.", "Trybem rozkazującym (Imperativ) zwracamy się bezpośrednio do kogoś z poleceniem. Czasownik stoi na początku zdania, przy „du“ i „ihr“ bez zaimka."),
    p(t(""), { after: 30 }),
    ...IM.flatMap(([a, c, plx]) => [p([t(a + ": ", { bold: true, color: NAVY }), t(c)], { after: PLM ? 0 : 30, indent: { left: 200 } }), ...pl(plx, 200)]),
    p([t("Achtung: ", { bold: true, color: RED }), t("nicht „Nehm!“ oder „Lese!“, sondern „Nimm!“, „Lies!“", { size: 19 })], { before: 40, after: 0 }),
    ...pl("Uwaga: nie „Nehm!“, „Lese!“, tylko „Nimm!“, „Lies!“"),
  ], PW, { borders: bs(b(BorderStyle.DASHED, 6, GRAY)), m: 120, ml: 200 })] })]);
}
const cut = p(t("✂ - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -", { color: GRAY, size: 16 }), { align: AlignmentType.CENTER, before: 80, after: 80 });
const props = { page: { size: { width: 11906, height: 16838 }, margin: { top: 450, bottom: 350, left: 550, right: 550 } } };
Packer.toBuffer(new Document({ styles: { default: { document: { run: { font: "Calibri", size: 20 } } } }, sections: [{ properties: props, children: PLM ? [satz()] : [satz(), cut, satz()] }] }))
  .then((buf) => fs.writeFileSync(__dirname + "/../Merksatz_Funktionen_Imperativ_Schnipsel" + (PLM ? "_Polnisch" : "") + ".docx", buf));
