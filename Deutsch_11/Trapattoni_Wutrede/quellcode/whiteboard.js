// Whiteboard-Anzeige während des Briefschreibens: Auftrag + Bewertungskriterien (16:9, große Schrift)
const pptxgen = require("pptxgenjs");
const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.333 x 7.5
pres.title = "Brief an die Spieler – Auftrag und Kriterien";

const NAVY = "1E2761", NAVY2 = "24306E", MUTED = "5B6B8C", LIGHT = "F4F7FD", LIGHT2 = "EAF0FA", ICE = "CADCFC", WHITE = "FFFFFF";
const RED = "D9534F", GOLD = "E8A33D";
// Kreidefarben der vier Seiten (wie im Tafelbild)
const K = { sach: "3F7CC4", selbst: "2E9E6B", bez: "C99A1E", app: "D9534F", meta: "8E5BB5" };
const HEAD = "Cambria", BODY = "Calibri";

function txt(s, text, o) {
  s.addText(text, Object.assign({ isTextBox: true, fontFace: BODY, fontSize: 18, color: NAVY, margin: 0, valign: "top" }, o));
}
function box(s, x, y, w, h, fill, o = {}) {
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, fill: { color: fill }, line: o.line ? { color: o.line, width: o.lw || 1.5 } : { type: "none" }, rectRadius: o.r ?? 0.08 });
}
function pill(s, label, x, y, w, col, size = 14, h = 0.42) {
  box(s, x, y, w, h, col, { r: h / 2 });
  txt(s, label, { x, y, w, h, fontSize: size, bold: true, color: WHITE, align: "center", valign: "middle" });
}

const AUFTRAG = "Schreibe als Giovanni Trapattoni einen Brief an Thomas Strunz, Mario Basler und Mehmet Scholl. Bringe die Missverständnisse aus der Pressekonferenz in Ordnung – so, dass alle vier Seiten der Nachricht wieder stimmen. Nutze die Techniken aus der letzten Stunde.";
const INHALT = [
  ["Sachinhalt", "Kritik sachlich und konkret – ohne Beleidigung", K.sach, "4"],
  ["Selbstoffenbarung", "Gefühle als Ich-Botschaft statt Wutausbruch", K.selbst, "4"],
  ["Beziehung", "Entschuldigung, Respekt, Empathie", K.bez, "4"],
  ["Appell", "konkrete Bitte direkt an die Spieler", K.app, "4"],
  ["Metakommunikation", "über die Pressekonferenz selbst sprechen", K.meta, "4"],
];
const SPRACHE = [
  ["Ausdruck und Ton", "respektvoll, treffende Wörter", "2"],
  ["Briefform", "Ort, Datum, Anrede, Schluss, Gruß", "2"],
  ["Rechtschreibung", "auch Groß- und Kleinschreibung", "2"],
  ["Grammatik", "Satzbau, Fälle, Zeiten", "2"],
  ["Zeichensetzung", "Kommas, v. a. bei Nebensätzen und Anrede", "2"],
];

function critList(s, x, y, w, title, pts, items, rowH, sizes) {
  box(s, x, y, w, 0.55, NAVY, { r: 0.06 });
  txt(s, title, { x: x + 0.2, y, w: w - 1.6, h: 0.55, fontFace: HEAD, fontSize: sizes.head, bold: true, color: WHITE, valign: "middle" });
  txt(s, pts, { x: x + w - 1.4, y, w: 1.2, h: 0.55, fontSize: sizes.head, bold: true, color: GOLD, align: "right", valign: "middle" });
  items.forEach((it, i) => {
    const yy = y + 0.65 + i * rowH;
    const [name, desc, col, p] = it.length === 4 ? it : [it[0], it[1], MUTED, it[2]];
    box(s, x, yy, w, rowH - 0.08, LIGHT, { r: 0.05 });
    box(s, x, yy, 0.12, rowH - 0.08, col, { r: 0.02 });
    txt(s, name, { x: x + 0.3, y: yy, w: w * (sizes.ratio || 0.4), h: rowH - 0.08, fontSize: sizes.name, bold: true, color: it.length === 4 ? col : NAVY, valign: "middle" });
    txt(s, desc, { x: x + 0.3 + w * (sizes.ratio || 0.4), y: yy, w: w * (1 - (sizes.ratio || 0.4)) - 1.1, h: rowH - 0.08, fontSize: sizes.desc, valign: "middle" });
    txt(s, p + " P.", { x: x + w - 0.95, y: yy, w: 0.8, h: rowH - 0.08, fontSize: sizes.name, bold: true, align: "right", valign: "middle" });
  });
}

// ============ Folie 1: alles auf einen Blick ============
{
  const s = pres.addSlide();
  s.background = { color: WHITE };
  box(s, 0, 0, 13.333, 1.05, NAVY, { r: 0 });
  txt(s, "Jetzt bringt ihr das in Ordnung!", { x: 0.45, y: 0, w: 8.5, h: 1.05, fontFace: HEAD, fontSize: 32, bold: true, color: WHITE, valign: "middle" });
  pill(s, "Einzelarbeit · ca. 20–25 Min.", 9.15, 0.31, 3.75, RED, 15);

  // links: Auftrag
  box(s, 0.45, 1.3, 5.6, 4.45, LIGHT2);
  txt(s, "Dein Auftrag", { x: 0.7, y: 1.45, w: 5.1, h: 0.45, fontFace: HEAD, fontSize: 22, bold: true });
  txt(s, AUFTRAG, { x: 0.7, y: 1.95, w: 5.1, h: 3.7, fontSize: 21, paraSpaceAfter: 6 });
  box(s, 0.45, 5.9, 5.6, 1.35, "FBECEB");
  txt(s, "Dieser Brief ist euer Test.", { x: 0.7, y: 6.0, w: 5.1, h: 0.45, fontFace: HEAD, fontSize: 20, bold: true, color: RED });
  txt(s, "Note: Inhalt + Sprache (30 Punkte)", { x: 0.7, y: 6.45, w: 5.1, h: 0.4, fontSize: 17 });
  txt(s, "Tipp: „Lieber Thomas, lieber Mario, lieber Mehmet,“ – danach klein weiter.", { x: 0.7, y: 6.82, w: 5.2, h: 0.4, fontSize: 13, italic: true, color: MUTED });

  // rechts: Kriterien
  const sz = { head: 17, name: 15, desc: 13, ratio: 0.42 };
  critList(s, 6.35, 1.25, 6.55, "I  Inhalt", "20 P.", INHALT, 0.44, sz);
  critList(s, 6.35, 4.2, 6.55, "II  Sprache", "10 P.", SPRACHE, 0.44, sz);
}

// ============ Folie 2: nur der Auftrag (groß) ============
{
  const s = pres.addSlide();
  s.background = { color: NAVY };
  s.addShape(pres.shapes.OVAL, { x: 10.6, y: -1.6, w: 4.6, h: 4.6, fill: { color: NAVY2 }, line: { type: "none" } });
  pill(s, "EINZELARBEIT · CA. 20–25 MIN. · DIESER BRIEF IST EUER TEST", 0.8, 0.7, 7.6, RED, 14);
  txt(s, "Jetzt bringt ihr das in Ordnung!", { x: 0.8, y: 1.4, w: 11.5, h: 0.9, fontFace: HEAD, fontSize: 40, bold: true, color: WHITE });
  txt(s, AUFTRAG, { x: 0.8, y: 2.6, w: 11.7, h: 2.8, fontSize: 28, color: WHITE, paraSpaceAfter: 6 });
  txt(s, "Techniken: Ich-Botschaft · Empathie · Wunsch & Bitte · Metakommunikation", { x: 0.8, y: 5.55, w: 11.9, h: 0.5, fontSize: 18, bold: true, color: ICE });
  txt(s, "Bewertet werden Inhalt und Sprache – auch Rechtschreibung, Grammatik und Zeichensetzung.", { x: 0.8, y: 6.15, w: 11.9, h: 0.5, fontSize: 18, color: ICE });
}

// ============ Folie 3: nur die Kriterien (groß) ============
{
  const s = pres.addSlide();
  s.background = { color: WHITE };
  txt(s, "So wird euer Brief bewertet", { x: 0.45, y: 0.3, w: 9, h: 0.7, fontFace: HEAD, fontSize: 32, bold: true });
  txt(s, "Inhalt + Sprache = Gesamtnote · 30 Punkte", { x: 0.45, y: 0.95, w: 9, h: 0.4, fontSize: 18, italic: true, color: MUTED });
  const sz = { head: 20, name: 16, desc: 15, ratio: 0.43 };
  critList(s, 0.45, 1.55, 6.15, "I  Inhalt", "20 P.", INHALT, 0.9, sz);
  critList(s, 6.75, 1.55, 6.15, "II  Sprache", "10 P.", SPRACHE, 0.9, sz);
}

pres.writeFile({ fileName: process.argv[2] || "Whiteboard_Brief_Auftrag_Kriterien.pptx" }).then((f) => console.log(f));
