const fs = require("fs");
const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, BorderStyle, AlignmentType, VerticalAlign, PageBreak, ShadingType } = require("docx");

// Wörter und Hinweise (Absolutismus und Ständegesellschaft, Wissen aus dem Test)
const WORDS = [
  ["ABSOLUTISMUS", "Der König regiert allein, ohne die Stände."],
  ["SONNENKÖNIG", "Beiname von Ludwig XIV."],
  ["VERSAILLES", "Prächtiges Schloss des Königs mit seinem Hof."],
  ["GOTTESGNADENTUM", "Der König sagt: Meine Macht kommt von Gott."],
  ["KLERUS", "Erster Stand: die Geistlichen."],
  ["ADEL", "Zweiter Stand."],
  ["STEUERN", "Das zahlt der Dritte Stand (Bauern und Bürger), Klerus und Adel nicht."],
  ["VORRECHTE", "Sonderrechte von Klerus und Adel."],
];

// ---- Gitter erzeugen (deterministische Zufallssuche) ----
let seed = 12345;
const rnd = () => ((seed = (seed * 1103515245 + 12345) & 0x7fffffff) / 0x7fffffff);
function tryLayout(order) {
  const cells = new Map(); // "r,c" -> {ch, h, v}
  const placed = [];
  const key = (r, c) => r + "," + c;
  const get = (r, c) => cells.get(key(r, c));
  function canPlace(word, r, c, dir) {
    const dr = dir === "v" ? 1 : 0, dc = dir === "h" ? 1 : 0;
    if (get(r - dr, c - dc) || get(r + dr * word.length, c + dc * word.length)) return -1;
    let cross = 0;
    for (let i = 0; i < word.length; i++) {
      const rr = r + dr * i, cc = c + dc * i, ex = get(rr, cc);
      if (ex) {
        if (ex.ch !== word[i] || ex[dir]) return -1;
        cross++;
      } else {
        // Nachbarn quer dürfen nicht belegt sein
        if (dir === "h" && (get(rr - 1, cc) || get(rr + 1, cc))) return -1;
        if (dir === "v" && (get(rr, cc - 1) || get(rr, cc + 1))) return -1;
      }
    }
    return cross;
  }
  function place(word, r, c, dir) {
    const dr = dir === "v" ? 1 : 0, dc = dir === "h" ? 1 : 0;
    for (let i = 0; i < word.length; i++) {
      const k = key(r + dr * i, c + dc * i);
      const ex = cells.get(k) || { ch: word[i] };
      ex[dir] = true; cells.set(k, ex);
    }
    placed.push({ word, r, c, dir });
  }
  place(order[0], 0, 0, "h");
  for (let w = 1; w < order.length; w++) {
    const word = order[w];
    let best = null;
    for (const p of placed) {
      for (let i = 0; i < word.length; i++) for (let j = 0; j < p.word.length; j++) {
        if (word[i] !== p.word[j]) continue;
        const dir = p.dir === "h" ? "v" : "h";
        const r = p.dir === "h" ? p.r - i : p.r + j;
        const c = p.dir === "h" ? p.c + j : p.c - i;
        const x = canPlace(word, r, c, dir);
        if (x >= 1) {
          let minR = Infinity, maxR = -Infinity, minC = Infinity, maxC = -Infinity;
          for (const k of cells.keys()) { const [a, b] = k.split(",").map(Number); minR = Math.min(minR, a); maxR = Math.max(maxR, a); minC = Math.min(minC, b); maxC = Math.max(maxC, b); }
          const er = dir === "v" ? r + word.length - 1 : r, ec = dir === "h" ? c + word.length - 1 : c;
          minR = Math.min(minR, r); maxR = Math.max(maxR, er); minC = Math.min(minC, c); maxC = Math.max(maxC, ec);
          const area = (maxR - minR + 1) * (maxC - minC + 1) + rnd() * 3 - x * 4;
          if (!best || area < best.area) best = { r, c, dir, area };
        }
      }
    }
    if (!best) return null;
    place(word, best.r, best.c, best.dir);
  }
  let minR = Infinity, maxR = -Infinity, minC = Infinity, maxC = -Infinity;
  for (const k of cells.keys()) { const [a, b] = k.split(",").map(Number); minR = Math.min(minR, a); maxR = Math.max(maxR, a); minC = Math.min(minC, b); maxC = Math.max(maxC, b); }
  return { cells, placed, minR, maxR, minC, maxC, area: (maxR - minR + 1) * (maxC - minC + 1), rows: maxR - minR + 1, cols: maxC - minC + 1 };
}
let bestL = null;
for (let it = 0; it < 4000; it++) {
  const ws = WORDS.map((x) => x[0]);
  ws.sort((a, b) => b.length - a.length + (rnd() - 0.5) * 3);
  const L = tryLayout(ws);
  if (L && (!bestL || L.area + Math.abs(L.rows - L.cols) * 2 < bestL.area + Math.abs(bestL.rows - bestL.cols) * 2)) bestL = L;
}
if (!bestL) throw new Error("kein Gitter");
const L = bestL;
// Nummerierung
const starts = L.placed.map((p) => ({ ...p, r: p.r - L.minR, c: p.c - L.minC }));
const sorted = [...starts].sort((a, b) => a.r - b.r || a.c - b.c);
const numAt = new Map(); let n = 0;
for (const p of sorted) { const k = p.r + "," + p.c; if (!numAt.has(k)) numAt.set(k, ++n); p.num = numAt.get(k); }
const grid = Array.from({ length: L.rows }, () => Array(L.cols).fill(null));
for (const p of starts) for (let i = 0; i < p.word.length; i++) grid[p.r + (p.dir === "v" ? i : 0)][p.c + (p.dir === "h" ? i : 0)] = p.word[i];
const clueOf = Object.fromEntries(WORDS);
const across = sorted.filter((p) => p.dir === "h"), down = sorted.filter((p) => p.dir === "v");
console.log(L.rows + "x" + L.cols, "Wörter:", starts.length);
console.log(grid.map((r) => r.map((x) => x || ".").join("")).join("\n"));
fs.writeFileSync("kreuzwort_daten.json", JSON.stringify({ rows: L.rows, cols: L.cols, grid, words: sorted.map((p) => ({ num: p.num, dir: p.dir, word: p.word, clue: clueOf[p.word], r: p.r, c: p.c })) }, null, 1));
