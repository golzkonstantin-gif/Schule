// Gemeinsames: Farben und die (gemischte) Nummerierung der Schnipsel
const D = require("./data.js");
const C = { dark: "2B2318", terra: "AD4426", sand: "F4ECDE", gold: "D9A441", muted: "5A5044", soft: "998C77", white: "FFFFFF", pale: "FBF8F2", light: "D9CDB8", green: "3C783C", blue: "3E5C76" };
// feste, gemischte Reihenfolge: Nummer 1..18 -> Karten-ID
const ORDER = ["K", "E", "Q", "H", "A", "N", "I", "C", "P", "F", "L", "O", "B", "R", "J", "G", "M", "D"];
const NR = Object.fromEntries(ORDER.map((id, i) => [id, i + 1]));
const card = id => D.karten.find(k => k.id === id);
const nr = ids => ids.split(",").map(s => s.trim()).filter(Boolean).map(id => NR[id]).sort((a, b) => a - b).join(", ");
module.exports = { D, C, ORDER, NR, card, nr };
