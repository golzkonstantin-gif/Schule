const fs = require("fs");
const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, BorderStyle, AlignmentType, VerticalAlign, HeightRule, ShadingType } = require("docx");

const WORDS = [
  ["ABSOLUTISMUS", "Der König regiert allein."],
  ["SONNENKÖNIG", "Beiname Ludwigs XIV."],
  ["VERSAILLES", "Prächtiges Schloss des Königs."],
  ["GOTTESGNADENTUM", "Die Macht des Königs kommt von Gott."],
  ["KLERUS", "Erster Stand: die Geistlichen."],
  ["ADEL", "Zweiter Stand."],
  ["STEUERN", "Abgaben an den Staat."],
  ["VORRECHTE", "Sonderrechte von Klerus und Adel."],
  ["BAUERN", "Sie bestellen die Felder."],
  ["BÜRGER", "Handwerker und Kaufleute in der Stadt."],
];
const N = 15;
let seed = 4711;
const rnd = () => ((seed = (seed * 1103515245 + 12345) & 0x7fffffff) / 0x7fffffff);
const FILL = "EEEEEEEENNNNNIIIISSSSRRRRAAATTTTDDDHHHUUULLLCCCGGGMMMOOOBBWWFFKKZZVPJ";
const DIRS = [[0, 1, 5], [1, 0, 5]]; // waagerecht, senkrecht, schräg (mit Gewicht)

function count(grid, word) {
  let c = 0;
  const dirs = [[0, 1], [1, 0], [1, 1], [0, -1], [-1, 0], [-1, -1], [1, -1], [-1, 1]];
  for (let r = 0; r < N; r++) for (let q = 0; q < N; q++) for (const [dr, dq] of dirs) {
    let ok = true;
    for (let i = 0; i < word.length; i++) { const rr = r + dr * i, qq = q + dq * i; if (rr < 0 || qq < 0 || rr >= N || qq >= N || grid[rr][qq] !== word[i]) { ok = false; break; } }
    if (ok) c++;
  }
  return c;
}
function build() {
  const grid = Array.from({ length: N }, () => Array(N).fill(null));
  const pos = [];
  const order = [...WORDS].sort((a, b) => b[0].length - a[0].length);
  for (const [w] of order) {
    let done = false;
    for (let tries = 0; tries < 400 && !done; tries++) {
      let x = rnd() * 10, d = DIRS[0];
      for (const dd of DIRS) { x -= dd[2]; if (x < 0) { d = dd; break; } }
      const [dr, dq] = d;
            const r0 = Math.floor(rnd() * (N - (dr ? w.length - 1 : 0))), q0 = Math.floor(rnd() * (N - (dq ? w.length - 1 : 0)));
      let ok = true;
      for (let i = 0; i < w.length; i++) { const c = grid[r0 + dr * i][q0 + dq * i]; if (c && c !== w[i]) { ok = false; break; } }
      if (!ok) continue;
      for (let i = 0; i < w.length; i++) grid[r0 + dr * i][q0 + dq * i] = w[i];
      pos.push({ word: w, r: r0, q: q0, dr, dq });
      done = true;
    }
    if (!done) return null;
  }
  const mark = grid.map((row) => row.map((c) => !!c));
  for (let r = 0; r < N; r++) for (let q = 0; q < N; q++) if (!grid[r][q]) grid[r][q] = FILL[Math.floor(rnd() * FILL.length)];
  for (const [w] of WORDS) if (count(grid, w) !== 1) return null;
  return { grid, mark, pos };
}
let G = null;
for (let i = 0; i < 2000 && !G; i++) G = build();
if (!G) throw new Error("kein Gitter");
console.log(G.grid.map((r) => r.join(" ")).join("\n"));
console.log(G.pos.map((p) => `${p.word} (${p.r},${p.q}) ${p.dr},${p.dq}`).join("\n"));

// ---------- docx ----------
const NAVY = "1E2761", MUTED = "5B6B8C";
const t = (text, o = {}) => new TextRun({ text, font: o.font || "Calibri", size: o.size || 20, bold: o.bold, italics: o.italics, color: o.color || "000000" });
const p = (runs, o = {}) => new Paragraph({ children: Array.isArray(runs) ? runs : [runs], alignment: o.align, spacing: { before: o.before ?? 0, after: o.after ?? 40 } });
const none = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
const thin = { style: BorderStyle.SINGLE, size: 4, color: "B9C0D4" };
const CS = 330;
function gridTable(show) {
  return new Table({ width: { size: CS * N, type: WidthType.DXA }, columnWidths: Array(N).fill(CS), rows: G.grid.map((row, r) => new TableRow({ height: { value: CS, rule: HeightRule.EXACT }, children: row.map((ch, q) => new TableCell({
    width: { size: CS, type: WidthType.DXA }, borders: { top: thin, bottom: thin, left: thin, right: thin }, margins: { top: 0, bottom: 0, left: 0, right: 0 }, verticalAlign: VerticalAlign.CENTER,
    shading: show && G.mark[r][q] ? { fill: "CADCFC", type: ShadingType.CLEAR, color: "auto" } : undefined,
    children: [p(t(ch, { size: 22, bold: show && G.mark[r][q], color: show && G.mark[r][q] ? NAVY : "000000" }), { after: 0, align: AlignmentType.CENTER })],
  })) })) });
}
function clueTable() {
  const half = Math.ceil(WORDS.length / 2);
  const col = (list) => new TableCell({ width: { size: 5350, type: WidthType.DXA }, borders: { top: none, bottom: none, left: none, right: none }, margins: { top: 20, bottom: 20, left: 40, right: 100 }, children: list.map(([w, c]) => p([t("• " + c, { size: 18 }), t("  (" + w.length + ")", { size: 16, color: MUTED }), t("  →  ______________", { size: 18, color: "8A96B5" })], { after: 70 })) });
  return new Table({ width: { size: 10700, type: WidthType.DXA }, columnWidths: [5350, 5350], rows: [new TableRow({ children: [col(WORDS.slice(0, half)), col(WORDS.slice(half))] })] });
}
function blatt() {
  return [
    p(t("BUCHSTABENRÄTSEL · GESCHICHTE 8D", { size: 15, bold: true, color: MUTED }), { after: 10 }),
    p(t("Absolutismus und Stände", { font: "Cambria", size: 30, bold: true, color: NAVY }), { after: 30 }),
    p(t("Im Gitter sind zehn Wörter versteckt: waagerecht (von links nach rechts) oder senkrecht (von oben nach unten). Finde zu jedem Hinweis das Wort, kreise es ein und schreibe es hinter den Hinweis. Die Zahl in Klammern ist die Buchstabenzahl.", { size: 17, italics: true, color: MUTED }), { after: 50 }),
    gridTable(false),
    p(t(" ", { size: 8 }), { after: 40 }),
    clueTable(),
  ];
}
const sep = new Paragraph({ children: [new TextRun({ text: "✂ - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -", size: 14, color: "8A96B5" })], spacing: { before: 40, after: 40 }, alignment: AlignmentType.CENTER });
const sol = [
  p(t("Lösung für die Lehrkraft", { font: "Cambria", size: 30, bold: true, color: NAVY }), { after: 60 }),
  gridTable(true),
  p(t(" ", { size: 8 }), { after: 60 }),
  ...G.pos.map((x) => p([t(x.word, { bold: true, size: 20 }), t(`   (Zeile ${x.r + 1}, Spalte ${x.q + 1}, ${x.dr && x.dq ? "schräg" : x.dr ? "senkrecht" : "waagerecht"})`, { size: 18, color: MUTED })], { after: 20 })),
  p(t("Die zehn Lösungswörter werden als Begriffsspeicher in den nächsten Stunden wiederholt.", { size: 19, italics: true, color: MUTED }), { before: 80 }),
];
(async () => {
  const doc = new Document({ styles: { default: { document: { run: { font: "Calibri", size: 20 } } } }, sections: [
    { properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 450, bottom: 400, left: 600, right: 600 } } }, children: [...blatt(), sep, ...blatt()] },
    { properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 700, bottom: 600, left: 700, right: 700 } } }, children: sol },
  ] });
  fs.writeFileSync("../Buchstabenraetsel_Absolutismus_Staende.docx", await Packer.toBuffer(doc));
  console.log("ok");
})();
