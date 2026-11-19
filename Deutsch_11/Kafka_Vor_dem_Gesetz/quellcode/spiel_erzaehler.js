// Spielvorlage mit Erzähler als Dramentext: Personenverzeichnis, Bühne; vor jeder Rede steht, wer spricht und was er dabei tut
// (Regieanweisungen kursiv in Klammern, Rede ohne Anführungszeichen wie im Drama). 3 Exemplare drucken.
const fs = require("fs");
const { Document, Packer, Paragraph, TextRun, AlignmentType, Table, TableRow, TableCell, WidthType, BorderStyle, ShadingType, VerticalAlign } = require("docx");
const D = require("./daten_macht");
const NAVY = "1E2761", MUTED = "5B6B8C", RED = "D9534F", BLUE = "3F7CC4", GREY = "444444";
const t = (text, o = {}) => new TextRun({ text, font: o.font || "Cambria", size: o.size || 22, bold: o.bold, italics: o.italics, color: o.color || "000000", smallCaps: o.caps });
const p = (runs, o = {}) => new Paragraph({ children: Array.isArray(runs) ? runs : [runs], alignment: o.align, indent: o.indent, spacing: { before: o.before ?? 0, after: o.after ?? 100, line: o.line ?? 275 } });
const ohne = (x) => x.replace(/^„|“$/g, "");
const farbe = { "Erzähler": GREY, "Türhüter": RED, "Mann": BLUE };
const name = { "Erzähler": "Erzähler", "Türhüter": "Türhüter", "Mann": "Der Mann vom Lande" };

const zeilen = []; // [Rolle, Absatz]
D.SPIEL_ERZAEHLER.forEach(([rolle, text, regie]) => {
  if (rolle === "Mann") {
    const stumm = text.sagt === "nichts";
    if (stumm) {
      // nur Handlung: eigene Regieanweisung, eingerückt
      zeilen.push(["Mann", p(t("(Der Mann " + text.macht.replace(/; /g, ", ").replace(/ (dich|du|deinen|dir)\b/g, " sich") + ".)", { italics: true, color: BLUE, size: 21 }), { after: 0 })]);
    } else {
      const leise = text.sagt.startsWith("leise: ");
      const rede = ohne(text.sagt.replace(/^leise: /, ""));
      const murmelt = text.sagt.startsWith("murmelt");
      zeilen.push(["Mann", p([
        t(name[rolle].toUpperCase(), { bold: true, color: BLUE, size: 21 }),
        t(" (" + text.macht + (leise ? "; leise" : "") + ")", { italics: true, color: BLUE, size: 21 }),
        t(murmelt ? "" : ": ", { size: 22 }),
        t(murmelt ? " – murmelt unverständlich." : rede, { size: 22, italics: murmelt, color: murmelt ? BLUE : "000000", bold: !murmelt && !leise }),
      ], { after: 0 })]);
    }
    return;
  }
  const runs = [t(name[rolle].toUpperCase(), { bold: true, color: farbe[rolle], size: 21 })];
  if (regie) runs.push(t(" (" + regie + ")", { italics: true, color: MUTED, size: 21 }));
  runs.push(t(": ", { size: 22 }));
  runs.push(t(ohne(text), { size: rolle === "Erzähler" ? 21 : 22, color: rolle === "Erzähler" ? GREY : "000000" }));
  zeilen.push([rolle, p(runs, { after: 0 })]);
});

const W = 11906 - 1800;
const leiste = { "Erzähler": "BFBFBF", "Türhüter": RED, "Mann": BLUE };
const hell = { "Erzähler": "F5F5F5", "Türhüter": "FDEEED", "Mann": "EDF4FC" };
const linie = { style: BorderStyle.SINGLE, size: 6, color: "9AA6C4" };
const none = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
const tabelle = new Table({
  width: { size: W, type: WidthType.DXA }, columnWidths: [420, 160, W - 580],
  rows: zeilen.map(([rolle, absatz], i) => new TableRow({ cantSplit: true, children: [
    new TableCell({ width: { size: 420, type: WidthType.DXA }, borders: { top: none, left: none, right: none, bottom: linie }, verticalAlign: VerticalAlign.TOP, margins: { top: 40, bottom: 40, left: 0, right: 60 }, children: [p(t(String(i + 1), { size: 17, color: MUTED, font: "Calibri" }), { after: 0, align: AlignmentType.RIGHT })] }),
    new TableCell({ width: { size: 160, type: WidthType.DXA }, shading: { fill: leiste[rolle], type: ShadingType.CLEAR, color: "auto" }, borders: { top: none, left: none, right: none, bottom: linie }, children: [p(t(""), { after: 0 })] }),
    new TableCell({ width: { size: W - 580, type: WidthType.DXA }, shading: { fill: hell[rolle], type: ShadingType.CLEAR, color: "auto" }, borders: { top: none, left: none, right: none, bottom: linie }, margins: { top: 40, bottom: 40, left: 140, right: 80 }, children: [absatz] }),
  ] })),
});
const doc = new Document({
  styles: { default: { document: { run: { font: "Cambria", size: 22 } } } },
  sections: [{ properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 600, bottom: 450, left: 900, right: 900 } } }, children: [
    p(t("Vor dem Gesetz", { size: 34, bold: true, color: NAVY }), { after: 0, align: AlignmentType.CENTER }),
    p(t("Szene nach Franz Kafka (gekürzt)", { size: 20, italics: true, color: MUTED }), { after: 140, align: AlignmentType.CENTER }),
    p(t("Personen", { size: 22, bold: true, color: NAVY }), { after: 30 }),
    p([t("ERZÄHLER", { bold: true, color: GREY, size: 20 }), t(" – steht am Rand und liest", { size: 20 })], { after: 0 }),
    p([t("TÜRHÜTER", { bold: true, color: RED, size: 20 }), t(" – steht vor dem Tor, spricht laut", { size: 20 })], { after: 0 }),
    p([t("DER MANN VOM LANDE", { bold: true, color: BLUE, size: 20 }), t(" – spielt vor allem mit dem Körper und spricht fast nichts", { size: 20 })], { after: 80 }),
    p([t("Bühne: ", { bold: true, size: 20 }), t("Ein offenes Tor (Tür des Klassenraums oder Pult). Seitwärts davon ein Stuhl.", { size: 20, italics: true })], { after: 40 }),
    p(t("Die farbige Leiste zeigt, wer dran ist: grau = Erzähler, rot = Türhüter, blau = Mann. Die Nummern helfen beim Wiederfinden.", { size: 18, italics: true, color: MUTED, font: "Calibri" }), { after: 120 }),
    tabelle,
    p(t("(Der Türhüter geht ab. Das Tor ist geschlossen.)", { italics: true, color: MUTED, size: 21 }), { before: 100, indent: { left: 740 } }),
  ] }],
});
Packer.toBuffer(doc).then((b) => fs.writeFileSync("../Macht_Ohnmacht/Spielvorlage_mit_Erzaehler.docx", b));
