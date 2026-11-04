// Bildimpuls „Gibs auf!“ zum Nacherzählen der Parabel – zwei Varianten zur Auswahl (16:9):
// A) eine Szene mit allen Elementen, B) Bildgeschichte in drei Bildern. Gezeichnet mit Formen (in PowerPoint bearbeitbar).
// Einziger Text im Bild: das Schild „Bahnhof“ in der Ferne.
const pptxgen = require("pptxgenjs");

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.333 x 7.5
pres.title = "Gibs auf! – Bildimpuls (Varianten)";

const NAVY = "1E2761", INK = "2B2F3A";
const SKY = "DCE7F2", DAWN = "F6DCC4", HOUSE = "9AA8BD", HOUSE2 = "8293AB", ROOF = "6F7F99", WIN = "4A5670";
const STREET = "C9CED6", STREET2 = "B7BDC7", LAMP = "F5D76E", COAT = "7A5C45", SKIN = "F1C9A5", UNI = "24306E", CASE = "8B5E3C";

const shp = (s, type, o) => s.addShape(pres.shapes[type], Object.assign({ line: { type: "none" } }, o));
const box = (s, x, y, w, h, color, o = {}) => shp(s, o.round ? "ROUNDED_RECTANGLE" : "RECTANGLE", Object.assign({ x, y, w, h, fill: { color }, rectRadius: o.r }, o.extra || {}));
const circ = (s, cx, cy, r, color, o = {}) => shp(s, "OVAL", { x: cx - r, y: cy - r, w: 2 * r, h: 2 * r, fill: o.nofill ? { type: "none" } : { color }, line: o.line ? { color: o.line, width: o.lw || 2 } : { type: "none" } });
// Strecke von (x1,y1) nach (x2,y2)
function seg(s, x1, y1, x2, y2, color = INK, width = 3, o = {}) {
  const x = Math.min(x1, x2), y = Math.min(y1, y2), w = Math.max(Math.abs(x2 - x1), 0.001), h = Math.max(Math.abs(y2 - y1), 0.001);
  const flipH = (x2 < x1) !== (y2 < y1);
  s.addShape(pres.shapes.LINE, { x, y, w, h, flipH, line: { color, width, dashType: o.dash, endArrowType: o.arrow } });
}
// Uhr mit Zeigern (Stunde, Minute in Grad, 0 = 12 Uhr)
function uhr(s, cx, cy, r, stunde, minute, o = {}) {
  circ(s, cx, cy, r, o.fill || "FFFFFF", { line: o.rand || INK, lw: o.lw || 2.5 });
  const zeiger = (deg, len, w, col) => { const a = (deg - 90) * Math.PI / 180; seg(s, cx, cy, cx + Math.cos(a) * len, cy + Math.sin(a) * len, col, w); };
  zeiger(stunde, r * 0.5, o.zw || 3, INK);
  zeiger(minute, r * 0.82, o.zw ? o.zw - 1 : 2, o.minCol || INK);
}
// Häuserzeile
function haeuser(s, x0, y0, n, w, hs, o = {}) {
  for (let i = 0; i < n; i++) {
    const h = hs[i % hs.length], x = x0 + i * w, y = y0 - h;
    box(s, x, y, w - 0.04, h, i % 2 ? HOUSE : HOUSE2);
    shp(s, "ISOSCELES_TRIANGLE", { x: x - 0.02, y: y - 0.35, w: w, h: 0.35, fill: { color: ROOF } });
    for (let r = 0; r < Math.floor(h / 0.55); r++) for (let c = 0; c < 2; c++) box(s, x + 0.12 + c * (w / 2 - 0.02), y + 0.18 + r * 0.55, 0.18, 0.26, o.licht && (r + c + i) % 5 === 0 ? LAMP : WIN);
  }
}
function laterne(s, x, yBoden, h) {
  box(s, x, yBoden - h, 0.06, h, INK);
  circ(s, x + 0.03, yBoden - h - 0.08, 0.22, "FFF3C4");
  circ(s, x + 0.03, yBoden - h - 0.08, 0.1, LAMP);
}
function bahnhof(s, x, y, w) {
  box(s, x, y, w, w * 0.45, "B39C86");
  shp(s, "ISOSCELES_TRIANGLE", { x: x - 0.05, y: y - w * 0.2, w: w + 0.1, h: w * 0.2, fill: { color: "8C7560" } });
  shp(s, "ROUND_2_SAME_RECTANGLE", { x: x + w * 0.38, y: y + w * 0.17, w: w * 0.24, h: w * 0.28, fill: { color: WIN } });
  box(s, x + w * 0.12, y + w * 0.04, w * 0.76, w * 0.11, "FFFFFF");
  s.addText("Bahnhof", { x: x + w * 0.12, y: y + w * 0.04, w: w * 0.76, h: w * 0.11, fontFace: "Calibri", fontSize: Math.max(7, w * 9), bold: true, color: INK, align: "center", valign: "middle", margin: 0 });
}
function turm(s, x, yBoden, w, h, std, min) {
  box(s, x, yBoden - h, w, h, "A3927E");
  shp(s, "ISOSCELES_TRIANGLE", { x: x - 0.08, y: yBoden - h - w * 1.1, w: w + 0.16, h: w * 1.1, fill: { color: "6E5E4D" } });
  uhr(s, x + w / 2, yBoden - h + w * 0.62, w * 0.36, std, min, { lw: 2 });
}
// Erzähler: rennend, schaut auf seine Uhr, Koffer in der Hand. Blickrichtung rechts (dir = 1) oder links (-1)
function erzaehler(s, x, y, k = 1, o = {}) {
  const d = o.dir || 1;
  const P = (dx, dy) => [x + dx * k * d, y + dy * k];
  const L = (a, b, col = COAT, w = 7 * k) => seg(s, ...P(...a), ...P(...b), col, w);
  // Beine (Laufschritt)
  L([0, 0.6], [0.32, 1.15], INK, 6 * k); L([0.32, 1.15], [0.55, 1.15], INK, 6 * k);
  L([0, 0.6], [-0.3, 1.05], INK, 6 * k); L([-0.3, 1.05], [-0.42, 1.3], INK, 6 * k);
  // Körper
  L([0, 0.62], [0.12, 0.02], COAT, 12 * k);
  if (o.fragen) {
    // Arm fragend zum Gegenüber ausgestreckt, offene Hand
    L([0.1, 0.12], [0.42, 0.2]); L([0.42, 0.2], [0.68, 0.08]);
    const [fx, fy] = P(0.72, 0.06); circ(s, fx, fy, 0.05 * k, SKIN);
  } else {
    // Arm mit Uhr vor dem Gesicht
    L([0.1, 0.12], [0.35, 0.3]); L([0.35, 0.3], [0.38, 0.0]);
    if (o.uhr !== false) { const [ux, uy] = P(0.38, -0.02); circ(s, ux, uy, 0.07 * k, "FFFFFF", { line: INK, lw: 1.5 }); }
  }
  // Arm nach hinten mit Koffer
  L([0.08, 0.14], [-0.22, 0.42]);
  if (o.koffer !== false) { const [kx, ky] = P(-0.22, 0.42); box(s, kx - 0.16 * k, ky, 0.32 * k, 0.24 * k, CASE, { round: true, r: 0.03 }); }
  // Kopf mit Hut
  const [hx, hy] = P(0.16, -0.18); circ(s, hx, hy, 0.15 * k, SKIN);
  box(s, hx - 0.17 * k, hy - 0.2 * k, 0.34 * k, 0.06 * k, INK); box(s, hx - 0.1 * k, hy - 0.34 * k, 0.2 * k, 0.15 * k, INK);
  // Bewegungslinien
  if (o.tempo !== false) for (let i = 0; i < 3; i++) { const [mx, my] = P(-0.55, 0.15 + i * 0.22); const [nx] = P(-1.05, 0); seg(s, mx, my, nx, my, "8A93A6", 2); }
}
// Schutzmann: Uniform, Mütze. Wendet sich vom Erzähler (links) ab: Blick und Schritt nach rechts,
// Hand vor dem Mund (lacht), der andere Arm schwingt weit nach hinten; Bogen zeigt die Drehung
function schutzmann(s, x, y, k = 1, o = {}) {
  const P = (dx, dy) => [x + dx * k, y + dy * k];
  const L = (a, b, col = UNI, w = 9 * k) => seg(s, ...P(...a), ...P(...b), col, w);
  // Beine: Schritt nach rechts
  L([0, 0.7], [0.3, 1.25], INK, 7 * k); L([0.3, 1.25], [0.48, 1.25], INK, 7 * k);
  L([0, 0.7], [-0.12, 1.3], INK, 7 * k);
  // Körper leicht nach rechts gedreht
  L([0, 0.7], [0.06, 0.02], UNI, 16 * k);
  // Arm schwingt weit nach hinten-oben (zum Erzähler hin, abwinkend)
  L([0.02, 0.12], [-0.35, 0.05]); L([-0.35, 0.05], [-0.6, -0.25]);
  // Arm zum Mund (Lachen)
  L([0.08, 0.14], [0.3, 0.32]); L([0.3, 0.32], [0.26, -0.14]);
  // Kopf mit Mütze, Blick nach rechts (Schirm rechts)
  const [hx, hy] = P(0.1, -0.22); circ(s, hx, hy, 0.16 * k, SKIN);
  box(s, hx - 0.17 * k, hy - 0.3 * k, 0.34 * k, 0.15 * k, UNI); box(s, hx - 0.02 * k, hy - 0.17 * k, 0.26 * k, 0.05 * k, INK);
  circ(s, hx + 0.02 * k, hy - 0.24 * k, 0.03 * k, "E8A33D");
  // Schwungbogen: Drehung weg vom Erzähler
  if (o.schwung !== false) shp(s, "ARC", { x: x - 0.75 * k, y: y - 0.45 * k, w: 0.9 * k, h: 0.9 * k, line: { color: "8A93A6", width: 2.5, dashType: "dash" }, angleRange: [190, 330] });
}
const rahmen = (s, x, y, w, h) => s.addShape(pres.shapes.RECTANGLE, { x, y, w, h, fill: { type: "none" }, line: { color: NAVY, width: 2 } });
const himmel = (s, x, y, w, h) => { box(s, x, y, w, h, SKY); box(s, x, y + h * 0.55, w, h * 0.45, DAWN); };

// ===== Variante A: eine Szene =====
{
  const s = pres.addSlide();
  s.background = { color: "FFFFFF" };
  himmel(s, 0, 0, 13.333, 4.6);
  // Bahnhof in der Ferne, Turm mit Uhr
  bahnhof(s, 9.6, 3.55, 1.5);
  turm(s, 2.0, 4.6, 0.9, 3.4, 230, 300);
  // Häuserzeilen
  haeuser(s, 0, 4.6, 2, 1.0, [1.8, 2.3]);
  haeuser(s, 3.0, 4.6, 4, 0.95, [2.0, 1.6, 2.4, 1.8]);
  haeuser(s, 11.2, 4.6, 3, 0.75, [2.2, 1.7, 2.0]);
  // Straßen: Hauptstraße und Abzweigungen ins Leere
  box(s, 0, 4.6, 13.333, 2.9, STREET);
  shp(s, "TRAPEZOID", { x: 7.0, y: 4.0, w: 3.2, h: 0.62, fill: { color: STREET2 } });
  shp(s, "TRAPEZOID", { x: 4.2, y: 4.25, w: 1.6, h: 0.37, fill: { color: STREET2 } });
  box(s, 0, 5.15, 13.333, 0.05, "FFFFFF"); box(s, 0, 6.8, 13.333, 0.05, "FFFFFF");
  laterne(s, 6.2, 5.2, 1.7); laterne(s, 12.4, 5.2, 1.7);
  // Figuren: Erzähler rennt mit Uhr und Koffer, Schutzmann wendet sich lachend ab
  erzaehler(s, 4.6, 4.95, 1.5);
  schutzmann(s, 9.3, 4.85, 1.5);
}

// ===== Variante B: Bildgeschichte in drei Bildern =====
{
  const s = pres.addSlide();
  s.background = { color: "FFFFFF" };
  const pw = 4.1, ph = 6.3, py = 0.6;
  [0.35, 4.62, 8.89].forEach((px, i) => {
    himmel(s, px, py, pw, ph * 0.62);
    box(s, px, py + ph * 0.62, pw, ph * 0.38, STREET);
    if (i === 0) { // Uhrenvergleich: Turmuhr spät, eigene Uhr früh
      haeuser(s, px, py + ph * 0.62, 2, 0.9, [1.6, 2.0]);
      turm(s, px + 2.2, py + ph * 0.62, 0.95, 2.4, 230, 300);
      erzaehler(s, px + 1.0, py + ph * 0.62 + 0.15, 1.2, { tempo: false });
      // eigene Uhr vergrößert (Lupe), zeigt eine frühere Zeit als die Turmuhr
      seg(s, px + 1.46, py + ph * 0.62 + 0.13, px + 2.75, py + 4.75, "8A93A6", 2, { dash: "dash" });
      uhr(s, px + 3.2, py + 4.95, 0.45, 210, 0, { rand: "8A93A6", lw: 2 });
    } else if (i === 1) { // Rennen und Verirren: Kreuzung, viele Straßen, Bahnhof fern
      haeuser(s, px, py + ph * 0.62, 4, 1.02, [1.7, 2.2, 1.5, 2.0]);
      bahnhof(s, px + 2.7, py + 2.55, 1.1);
      shp(s, "TRAPEZOID", { x: px + 0.3, y: py + ph * 0.62 - 0.3, w: 1.3, h: 0.32, fill: { color: STREET2 } });
      shp(s, "TRAPEZOID", { x: px + 2.3, y: py + ph * 0.62 - 0.3, w: 1.3, h: 0.32, fill: { color: STREET2 } });
      erzaehler(s, px + 1.6, py + ph * 0.62 + 0.2, 1.3);
    } else { // Schutzmann wendet sich lachend ab
      haeuser(s, px, py + ph * 0.62, 4, 1.02, [1.9, 1.6, 2.1, 1.7]);
      erzaehler(s, px + 0.8, py + ph * 0.62 + 0.2, 1.3, { fragen: true, tempo: false });
      schutzmann(s, px + 2.9, py + ph * 0.62 + 0.1, 1.3);
    }
    rahmen(s, px, py, pw, ph);
    s.addText(String(i + 1), { x: px + 0.1, y: py + 0.08, w: 0.5, h: 0.5, fontFace: "Calibri", fontSize: 20, bold: true, color: "FFFFFF", fill: { color: NAVY }, align: "center", valign: "middle", margin: 0 });
  });
}

pres.writeFile({ fileName: "Bildimpuls_Varianten.pptx" });
