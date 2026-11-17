const fs = require("fs");
const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, BorderStyle, ShadingType, AlignmentType, VerticalAlign, PageBreak, Footer, PageNumber } = require("docx");
const NAVY = "1E2761", MUTED = "5B6B8C", LIGHT = "F4F7FD", LIGHT2 = "EAF0FA", RED = "D9534F", GREEN = "2E9E6B";
const W = 10700;
const t = (text, o = {}) => new TextRun({ text, font: o.font || "Calibri", size: o.size || 21, bold: o.bold, italics: o.italics, color: o.color || "000000" });
const p = (runs, o = {}) => new Paragraph({ children: Array.isArray(runs) ? runs : [runs], alignment: o.align, spacing: { before: o.before ?? 0, after: o.after ?? 60, line: o.line }, keepNext: o.keepNext });
const solid = (c, s = 6) => ({ style: BorderStyle.SINGLE, size: s, color: c });
const all = (b) => ({ top: b, bottom: b, left: b, right: b });
const noB = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
const brk = () => new Paragraph({ children: [new PageBreak()] });

const TEXT = {
  A: { titel: "Text A: Der Ballhausschwur (20. Juni 1789)", farbe: GREEN, absätze: [
    "Seit Wochen streiten die Stände in Versailles darüber, wie abgestimmt wird. Der Dritte Stand (Bauern und Bürger) vertritt fast alle Menschen in Frankreich, hat nach Ständen aber nur eine von drei Stimmen. Am 17. Juni 1789 erklärt er sich deshalb zur Nationalversammlung. Er sagt: Wir sprechen für die Nation. Einige Geistliche schließen sich an.",
    "Der König will das nicht hinnehmen. Am Morgen des 20. Juni ist der Sitzungssaal in Versailles abgesperrt, Wachen stehen vor der Tür. Offiziell heißt es, der Saal werde für eine Sitzung des Königs hergerichtet.",
    "Die Abgeordneten weichen aus. Sie gehen in eine nahe gelegene Ballhalle, in der sonst Ballspiele gespielt werden. Dort heben sie die Hand und schwören, nicht auseinanderzugehen, bis Frankreich eine Verfassung hat. Fast alle unterschreiben. Sie benutzen keine Waffen, sondern Worte, ihre Unterschrift und ihren Eid.",
    "Niemand wird verletzt. Aber die Abgeordneten wissen, dass der König sie auflösen oder verhaften lassen könnte. Am 23. Juni befiehlt der König noch einmal, dass die Stände getrennt tagen. Die Abgeordneten bleiben sitzen.",
    "Der König kann die Versammlung nicht zum Gehen zwingen. Am 27. Juni fordert er Klerus und Adel auf, sich der Nationalversammlung anzuschließen. In der Nationalversammlung zählt nun jeder Abgeordnete."],
    wörter: "Ballhalle: Halle, in der man Ballspiele spielte · Eid: feierliches Versprechen · Verfassung: Gesetze, die festlegen, wer wie regiert · Abgeordnete: gewählte Vertreter" },
  B: { titel: "Text B: Der Sturm auf die Bastille (14. Juli 1789)", farbe: RED, absätze: [
    "Im Sommer 1789 hungern viele Menschen in Paris, das Brot ist so teuer wie nie. Anfang Juli zieht der König Soldaten um Paris und Versailles zusammen. Am 11. Juli entlässt er außerdem den beliebten Finanzminister Necker. In Paris glauben viele, dass die Soldaten gegen die Stadt und gegen die Nationalversammlung eingesetzt werden.",
    "Am 14. Juli holen sich Tausende Pariser Gewehre aus einem Waffenlager. Es fehlt Pulver. Das liegt in der Bastille, einer Festung im Osten der Stadt, in der der König Gegner ohne Urteil einsperren lässt. Die Menge zieht vor die Bastille und verlangt die Herausgabe der Waffen.",
    "Es kommt zum Kampf, es wird geschossen. Nach einigen Stunden ergibt sich der Gouverneur, der Befehlshaber der Festung. Die Menge dringt ein und befreit die Gefangenen. Den Gouverneur tötet sie auf dem Weg zum Rathaus.",
    "Etwa 100 Angreifer sterben, auch einige Verteidiger. In der Bastille sitzen nur sieben Gefangene.",
    "Der König erkennt, dass er sich nicht auf seine Soldaten verlassen kann, und zieht sie wenige Tage später ab. Aufstände breiten sich im ganzen Land aus. Die Bastille wird abgerissen, der 14. Juli ist bis heute Nationalfeiertag in Frankreich."],
    wörter: "Bastille: Festung und Gefängnis des Königs in Paris · Gouverneur: Befehlshaber der Festung · Pulver: Schießpulver für Waffen · Nationalfeiertag: Feiertag für das ganze Land" },
};
const KRIT = [
  ["1  Vorgeschichte", "Warum kam es dazu?"],
  ["2  Verlauf", "Was geschah? Wer tat was?"],
  ["3  Mittel", "Gewalt: ja oder nein? Woran erkennst du das?"],
  ["4  Opfer und Folgen", "Wer litt oder riskierte etwas?"],
  ["5  Wirkung", "Was änderte sich danach?"],
];
const LOES = {
  1: ["Der Dritte Stand erklärt sich zur Nationalversammlung, der König lässt den Saal sperren.", "Hunger, Brot teuer; Soldaten um Paris, Finanzminister entlassen, Angst vor dem Militär."],
  2: ["Die Abgeordneten treffen sich in einer Ballhalle und schwören, nicht auseinanderzugehen, bis Frankreich eine Verfassung hat.", "Menschen holen Gewehre, suchen Pulver, greifen die Bastille an, der Gouverneur ergibt sich."],
  3: ["Gewaltlos: Worte und Eid, keine Waffen.", "Gewaltsam: Waffen, Schüsse, Kampf."],
  4: ["Keine Toten; Risiko: Auflösung durch den König.", "Etwa 100 Tote unter den Angreifern, der Gouverneur wird getötet, nur sieben Gefangene."],
  5: ["Die Versammlung bleibt, der König kann sie nicht zum Gehen zwingen, später soll sich der Adel anschließen.", "Der König zieht die Soldaten ab, die Bastille wird zum Symbol der Revolution."],
};

const cell = (children, o = {}) => new TableCell({ children: Array.isArray(children) ? children : [children], width: { size: o.w, type: WidthType.DXA }, shading: o.fill ? { fill: o.fill, type: ShadingType.CLEAR, color: "auto" } : undefined, borders: o.borders || all(solid("8A96B5", 6)), margins: { top: o.m ?? 70, bottom: o.m ?? 70, left: 110, right: 110 }, verticalAlign: o.valign || VerticalAlign.TOP });
const footer = new Footer({ children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ font: "Calibri", size: 16, color: MUTED, children: ["Geschichte 8D · Stunde 3 Widerstand · Seite ", PageNumber.CURRENT] })] })] });

function textblatt(me) {
  const T = TEXT[me];
  const c = [];
  c.push(p([t("GESCHICHTE 8D · PARTNERPUZZLE · TEXT   ", { size: 16, bold: true, color: MUTED }), t("Name: ______________________", { size: 18, color: MUTED })], { after: 30 }));
  c.push(p(t(T.titel, { font: "Cambria", size: 32, bold: true, color: T.farbe }), { after: 120 }));
  T.absätze.forEach((x) => c.push(p(t(x, { font: "Cambria", size: 25 }), { after: 130, line: 300, align: AlignmentType.JUSTIFIED })));
  c.push(p(t(T.wörter, { size: 19, color: MUTED }), { before: 60, after: 120 }));
  c.push(p([t("Aufgabe  ", { bold: true, color: RED, size: 22 }), t("Lies den Text genau. Unterstreiche, was zu den fünf Kriterien passt. Fülle dann die Tabelle für deinen Text aus.", { size: 22 })], { after: 0 }));
  return c;
}
function tabelle() {
  const c = [];
  c.push(p([t("GESCHICHTE 8D · PARTNERPUZZLE · TABELLE   ", { size: 16, bold: true, color: MUTED }), t("Name: ______________________", { size: 18, color: MUTED })], { after: 30 }));
  c.push(p(t("Ballhausschwur und Bastille im Vergleich", { font: "Cambria", size: 32, bold: true, color: NAVY }), { after: 90 }));
  c.push(p([t("Aufgabe  ", { bold: true, color: RED, size: 20 }), t("1  Fülle die Spalte für deinen Text aus (Stichworte).   2  Höre deinem Partner zu und fülle die andere Spalte aus.", { size: 20 })], { after: 100 }));
  const CW = [2500, 4100, 4100];
  const head = (txt, col, w) => cell(p(t(txt, { bold: true, color: "FFFFFF", size: 21 }), { after: 0, align: AlignmentType.CENTER }), { w, fill: col, valign: VerticalAlign.CENTER, borders: all(solid(col, 8)) });
  const rows = [new TableRow({ children: [head("Kriterium", NAVY, CW[0]), head("Text A: Ballhausschwur", GREEN, CW[1]), head("Text B: Bastille", RED, CW[2])] })];
  KRIT.forEach(([n, q], i) => rows.push(new TableRow({ cantSplit: true, height: { value: 2050, rule: "atLeast" }, children: [
    cell([p(t(n, { font: "Cambria", bold: true, size: 23, color: NAVY }), { after: 20 }), p(t(q, { size: 19 }), { after: 0 })], { w: CW[0], fill: i % 2 ? LIGHT : LIGHT2 }),
    cell(p(t(""), { after: 0 }), { w: CW[1] }),
    cell(p(t(""), { after: 0 }), { w: CW[2] }),
  ] })));
  c.push(new Table({ width: { size: W, type: WidthType.DXA }, columnWidths: CW, rows }));
  c.push(p([t("□ Ich gebe das Blatt freiwillig ab.", { size: 22, bold: true }), t("          Punkte: ____ von 10", { size: 22 })], { before: 120, after: 0 }));
  return c;
}
function loesung() {
  const c = [];
  c.push(p(t("LÖSUNG FÜR DIE LEHRKRAFT · GESCHICHTE 8D · STUNDE 3", { size: 17, bold: true, color: RED }), { after: 20 }));
  c.push(p(t("Partnerpuzzle: Erwartungen und Punkte", { font: "Cambria", size: 34, bold: true, color: NAVY }), { after: 60 }));
  c.push(p(t("Je Feld ein Punkt (zwei Texte mal fünf Kriterien = 10 Punkte), wenn die Kernaussage im Feld steht (Stichworte genügen). Das Sachurteil und das Werturteil (Folie 17) zählen nicht zu den 10 Punkten. Die Zahlen zur Bastille bitte gegen das Lehrbuch prüfen. Notenschlüssel nach Ihrer Festlegung (Raster auf Folie 7: 0 bis 4, 5 bis 8, 9 bis 10 Punkte).", { size: 20, color: MUTED }), { after: 120 }));
  const CW = [2300, 4200, 4200];
  const head = (txt, col, w) => cell(p(t(txt, { bold: true, color: "FFFFFF", size: 20 }), { after: 0, align: AlignmentType.CENTER }), { w, fill: col, valign: VerticalAlign.CENTER, borders: all(solid(col, 8)) });
  const rows = [new TableRow({ children: [head("Kriterium", NAVY, CW[0]), head("Ballhausschwur", GREEN, CW[1]), head("Bastille", RED, CW[2])] })];
  KRIT.forEach(([n, q], i) => rows.push(new TableRow({ cantSplit: true, children: [
    cell([p(t(n, { font: "Cambria", bold: true, size: 21, color: NAVY }), { after: 10 }), p(t(q, { size: 17 }), { after: 0 })], { w: CW[0], fill: i % 2 ? LIGHT : LIGHT2 }),
    cell(p(t(LOES[i + 1][0], { size: 20 }), { after: 0 }), { w: CW[1] }),
    cell(p(t(LOES[i + 1][1], { size: 20 }), { after: 0 }), { w: CW[2] }),
  ] })));
  c.push(new Table({ width: { size: W, type: WidthType.DXA }, columnWidths: CW, rows }));
  c.push(p(t("Hinweise zur Haltung", { font: "Cambria", size: 26, bold: true, color: NAVY }), { before: 200, after: 60 }));
  ["Wirksam heißt nicht gerechtfertigt: Beide Wege bewirkten etwas, aber sie kosteten Unterschiedliches (Bilanz, Folie 16).", "Aha: Die Reihenfolge zeigt, dass der gewaltlose Schritt die Grundlage schuf und der Druck in Paris den König zum Rückzug zwang. Der Streit um die Frage, ob die Gewalt den König wirklich allein zum Nachgeben brachte, ist offen; die Klasse soll die Reihenfolge belegen, nicht Schuld verteilen.", "Gewalt war eine Reaktion auf die drohende Gewalt (Soldaten), das ist ein Anlass, keine Rechtfertigung. Die Aufgabe, ein Urteil mit einem offenen Maßstab (wirksam, gerecht, verantwortbar) zu begründen, lässt beide Antworten zu.", "Wirkung beim Ballhausschwur: Nach dem Schwur versuchte der König noch, die Beschlüsse zu kassieren (23. Juni); die Abgeordneten blieben. Am 27. Juni forderte er Klerus und Adel auf, sich anzuschließen. Zum Prüfen mit dem Lehrbuch."].forEach((x) => c.push(p([t("• ", { size: 20 }), t(x, { size: 20 })], { after: 50 })));
  return c;
}
(async () => {
  const sec = (children) => ({ properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 650, bottom: 600, left: 700, right: 700 } } }, footers: { default: footer }, children });
  const d1 = new Document({ styles: { default: { document: { run: { font: "Calibri", size: 21 } } } }, sections: [sec([...textblatt("A"), brk(), ...textblatt("B")])] });
  fs.writeFileSync("../Partnerpuzzle_Texte_A_und_B_8D.docx", await Packer.toBuffer(d1));
  const d3 = new Document({ styles: { default: { document: { run: { font: "Calibri", size: 21 } } } }, sections: [sec(tabelle())] });
  fs.writeFileSync("../Partnerpuzzle_Tabelle_8D.docx", await Packer.toBuffer(d3));
  const d2 = new Document({ styles: { default: { document: { run: { font: "Calibri", size: 21 } } } }, sections: [sec(loesung())] });
  fs.writeFileSync("../Loesung_Partnerpuzzle_8D.docx", await Packer.toBuffer(d2));
  console.log("ok");
})();
