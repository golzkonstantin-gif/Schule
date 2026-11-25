const pptxgen = require("pptxgenjs");
const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE";
pres.title = "Reportage – Aussagen prüfen";
const NAVY = "1E2761", NAVY2 = "24306E", MUTED = "5B6B8C", LIGHT = "F4F7FD", LIGHT2 = "EAF0FA", ICE = "CADCFC", WHITE = "FFFFFF";
const RED = "D9534F", GREEN = "2E9E6B", YEL = "E8A33D", BLUE = "3F7CC4", VIO = "7A5BA6", TEAL_ = "2A8C8C";
const HEAD = "Cambria", BODY = "Calibri";
let FOOT = "Mittwoch, 4.11. · Deutsch 7A";
let pageNo = 1;
function txt(slide, text, o) {
  slide.addText(text, Object.assign({ isTextBox: true, fontFace: BODY, fontSize: 16, color: NAVY, margin: 0, valign: "top" }, o));
}
function base(kicker, title, notes) {
  const s = pres.addSlide();
  pageNo++;
  s.background = { color: WHITE };
  txt(s, kicker.toUpperCase(), { x: 0.6, y: 0.4, w: 9, h: 0.3, fontSize: 12, color: MUTED, charSpacing: 2 });
  txt(s, title, { x: 0.6, y: 0.72, w: 12.1, h: 0.7, fontFace: HEAD, fontSize: 32, bold: true });
  txt(s, FOOT, { x: 0.6, y: 7.0, w: 7, h: 0.25, fontSize: 10, color: MUTED });
  txt(s, String(pageNo), { x: 12.2, y: 7.0, w: 0.5, h: 0.25, fontSize: 10, color: MUTED, align: "right" });
  if (notes) s.addNotes(notes);
  return s;
}
function box(s, x, y, w, h, fill, o = {}) {
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, fill: { color: fill }, line: o.line ? { color: o.line, width: o.lw || 1.5 } : { type: "none" }, rectRadius: o.r ?? 0.08 });
}
function darkBg(s) {
  s.background = { color: NAVY };
  s.addShape(pres.shapes.OVAL, { x: 10.6, y: -1.6, w: 4.6, h: 4.6, fill: { color: NAVY2 }, line: { type: "none" } });
  s.addShape(pres.shapes.OVAL, { x: -1.5, y: 4.6, w: 4.0, h: 4.0, fill: { color: NAVY2 }, line: { type: "none" } });
}
function num(s, n, x, y, col = NAVY, d = 0.5) {
  s.addShape(pres.shapes.OVAL, { x, y, w: d, h: d, fill: { color: col }, line: { type: "none" } });
  txt(s, String(n), { x, y, w: d, h: d, fontSize: d * 32, bold: true, color: WHITE, align: "center", valign: "middle" });
}
// Übung im festen Format: Nummer + Titel, nummerierte Schritte, Zeit, Arbeitsform, „Fertig? Stift hinlegen“
function uebung(kicker, title, steps, meta, notes, extra) {
  const s = base(kicker, title, notes);
  steps.forEach((st, i) => {
    const y = 1.75 + i * 0.8;
    num(s, i + 1, 0.6, y);
    txt(s, st, { x: 1.3, y: y + 0.04, w: 11.3, h: 0.7, fontSize: 21, bold: true });
  });
  if (extra) extra(s, 1.75 + steps.length * 0.8 + 0.1);
  box(s, 0.6, 6.25, 12.1, 0.55, LIGHT2);
  txt(s, [{ text: "Zeit: ", options: { bold: true } }, { text: meta + "      " }, { text: "Fertig? Stift hinlegen.", options: { italic: true, color: MUTED } }], { x: 0.85, y: 6.33, w: 11.7, h: 0.4, fontSize: 16 });
  return s;
}

function section(nr, title, sub, col) {
  const s = pres.addSlide();
  darkBg(s);
  box(s, 0.9, 2.05, 2.2, 0.42, col, { r: 0.21 });
  txt(s, "TEIL " + nr, { x: 0.9, y: 2.05, w: 2.2, h: 0.42, fontSize: 12, bold: true, color: WHITE, align: "center", valign: "middle" });
  txt(s, title, { x: 0.9, y: 2.65, w: 12, h: 1.0, fontFace: HEAD, fontSize: 40, bold: true, color: WHITE });
  txt(s, sub, { x: 0.9, y: 3.75, w: 11.5, h: 0.5, fontSize: 20, italic: true, color: ICE });
  pageNo++;
  return s;
}

// ================= Aufdeck-Folien: „Sehen mit den Fingerkuppen“ – Aussagen prüfen (LB S. 183–184) =================
const AUS = JSON.parse(require("fs").readFileSync(__dirname + "/aussagen_reportage.json", "utf8"));
FOOT = "Deutsch 7A · Sachtexte · Reportage";
uebung("Verarbeiten · Arbeitsblatt", "Sehen mit den Fingerkuppen – Aussagen prüfen", [
  "Prüfe jede Aussage auf dem Arbeitsblatt allein: richtig oder falsch?",
  "Belege mit Zeile und Zitat in Anführungszeichen.",
  "Ist die Aussage falsch? Schreibe auf, was richtig ist.",
], "12 Minuten · allein · Nr. 10 mit *", "Ansage: „Jetzt arbeitet ihr allein. Lehrbuch S. 183–184 auf. Prüft jede Aussage, belegt mit Zeile und Zitat und verbessert falsche Aussagen.“ Danach gemeinsam Aussage für Aussage prüfen: Wer drankommt, nennt Urteil und Beleg, erst dann aufdecken. Alle verbessern mit Grün. Niemand muss an die Tafel.", (s, y) => {
  box(s, 0.6, y + 0.2, 12.1, 0.95, LIGHT);
  txt(s, [{ text: "Beispiel: ", options: { bold: true, color: GREEN } }, { text: "„Die Ausstellung ist hell beleuchtet.“ – falsch · Z. 1: „Herzlich willkommen in der Finsternis!“ · Richtig ist: Die Räume sind völlig dunkel." }], { x: 0.85, y: y + 0.35, w: 11.6, h: 0.7, fontSize: 16 });
});
AUS.forEach(([a, ok, z, zit, r], i) => {
  const n = i + 1;
  const s = base("Auswerten · gemeinsam prüfen", `Aussage ${i === 9 ? "* " : ""}${n} von ${AUS.length}`, `Erst einen Schüler Urteil und Beleg nennen lassen${r ? ", bei falschen Aussagen auch „Richtig ist …“" : ""}. Dann diese Folie zeigen. Alle verbessern mit Grün.`);
  box(s, 0.6, 1.6, 12.1, 1.3, LIGHT);
  txt(s, "„" + a + "“", { x: 0.9, y: 1.75, w: 9.6, h: 1.0, fontSize: 24, bold: true, fontFace: HEAD, valign: "middle" });
  box(s, 10.7, 2.0, 1.8, 0.55, ok ? GREEN : RED, { r: 0.25 });
  txt(s, ok ? "✓ richtig" : "✗ falsch", { x: 10.7, y: 2.0, w: 1.8, h: 0.55, fontSize: 17, bold: true, color: WHITE, align: "center", valign: "middle" });
  box(s, 0.6, 3.15, 12.1, 1.45, WHITE, { line: "C9D3EA" });
  txt(s, [{ text: "Beleg: ", options: { bold: true, color: NAVY } }, { text: z + "  " }, { text: zit, options: { italic: true } }], { x: 0.9, y: 3.3, w: 11.6, h: 1.2, fontSize: 20, valign: "middle" });
  if (r) {
    box(s, 0.6, 4.85, 12.1, 1.1, WHITE, { line: RED, lw: 2 });
    txt(s, [{ text: "Richtig ist: ", options: { bold: true, color: RED } }, { text: r }], { x: 0.9, y: 5.0, w: 11.6, h: 0.8, fontSize: 20, valign: "middle" });
  }
  txt(s, "Vergleiche und verbessere mit Grün.", { x: 0.6, y: 6.35, w: 12.1, h: 0.35, fontSize: 15, italic: true, color: MUTED });
});
pres.writeFile({ fileName: __dirname + "/../Folien_Reportage_Aussagen_pruefen.pptx" });
