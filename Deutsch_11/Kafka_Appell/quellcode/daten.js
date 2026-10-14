// Gemeinsame Daten der Stunde „Testauswertung, Rechtschreibtraining und Kafkas Appell“
// Auszug: Kafka, „Brief an den Vater“, S. 3 (DigBib-Ausgabe), mit dem Original abgeglichen, gekürzt […]
const TEXT = [
  "Wobei ich Dich aber immerfort bitte, nicht zu vergessen, daß ich",
  "niemals im entferntesten an eine Schuld Deinerseits glaube. Du",
  "wirktest so auf mich, wie Du wirken mußtest, nur sollst Du aufhören,",
  "es für eine besondere Bosheit meinerseits zu halten, daß ich dieser",
  "Wirkung erlegen bin. […] ich kann nicht glauben, daß ein freundliches",
  "Wort, ein stilles Bei-der-Hand-Nehmen, ein guter Blick mir nicht alles",
  "hätten abfordern können, was man wollte.",
];
const QUELLE = "Franz Kafka: „Brief an den Vater“ (1919), S. 3";

// Fußnoten: [Wort im Text, Erklärung]
const GLOSSAR = [
  ["immerfort", "immer wieder, ständig"],
  ["Deinerseits", "von Deiner Seite"],
  ["Bosheit", "böse Absicht"],
  ["erlegen", "nicht widerstehen können"],
  ["abfordern", "abverlangen"],
];

const KONTEXT = "Franz Kafka (1883–1924) schrieb 1919 mit 36 Jahren einen über 100 Seiten langen Brief an seinen Vater Hermann, einen erfolgreichen, lauten und bestimmenden Kaufmann. Vor ihm fühlte sich der Sohn klein und verstummte oft. Kafka wollte mit dem Brief erreichen, dass sich die Beziehung bessert. Den Brief gab er seiner Mutter – sie reichte ihn nie weiter.";
const KONTEXT_ZITAT = ["… daß es uns beide ein wenig beruhigen und Leben und Sterben leichter machen kann.", "Franz Kafka, letzter Satz des Briefs (S. 20)"];

// Information zum Appell
const APPELL = [
  ["Appell", "Mit dieser Seite einer Nachricht will der Sender erreichen, dass der Empfänger etwas tut, denkt oder fühlt."],
  ["offen", "direkt ausgesprochen, z. B. als Bitte: „Hör mir bitte zu!“"],
  ["verdeckt", "versteckt in einem Wunsch, einer Frage oder einer Beschreibung: „Du hörst mir nie zu.“"],
];

// ===== Rechtschreibtraining: Großschreibung von Substantiven und Komma vor Nebensätzen =====
// f = Fehlerfassung (Substantive klein, Kommas fehlen), r = richtig; gleiche Wortzahl, Unterschiede werden automatisch markiert
const DEMO = [
  { f: "Beim essen durfte niemand reden.", r: "Beim Essen durfte niemand reden." },
  { f: "Der vater merkte nicht dass sein sohn angst hatte.", r: "Der Vater merkte nicht, dass sein Sohn Angst hatte." },
];
const SAETZE = [
  { f: "Ich glaube dass der vater keine geduld hatte.", r: "Ich glaube, dass der Vater keine Geduld hatte." },
  { f: "Ich hoffe der vater hört mir endlich zu.", r: "Ich hoffe, der Vater hört mir endlich zu." },
  { f: "Wenn Franz etwas erzählte hörte der vater nicht zu.", r: "Wenn Franz etwas erzählte, hörte der Vater nicht zu." },
  { f: "Beim schreiben achte ich auf die kommas.", r: "Beim Schreiben achte ich auf die Kommas." },
  { f: "Kafka hatte angst weil der vater oft laut wurde.", r: "Kafka hatte Angst, weil der Vater oft laut wurde." },
  { f: "Ich wünsche mir dass meine eltern beim zuhören das handy weglegen.", r: "Ich wünsche mir, dass meine Eltern beim Zuhören das Handy weglegen." },
  { f: "Der sohn wusste nicht ob er seinem vater die wahrheit sagen durfte.", r: "Der Sohn wusste nicht, ob er seinem Vater die Wahrheit sagen durfte." },
  { f: "Ich denke Kafka hätte den brief abschicken sollen.", r: "Ich denke, Kafka hätte den Brief abschicken sollen." },
  { f: "Ein freundliches wort hätte dem kind mehr geholfen als jede drohung.", r: "Ein freundliches Wort hätte dem Kind mehr geholfen als jede Drohung." },
  { f: "Weil er sich fürchtete schrieb er seinem vater einen brief.", r: "Weil er sich fürchtete, schrieb er seinem Vater einen Brief." },
];
const ZUSAMMEN = 2; // Sätze 1–2 gemeinsam, ab Satz 3 allein

// Wortweiser Vergleich: [{ text, fehler: "gross" | "komma" | null }]
function vergleich(f, r) {
  const a = f.split(" "), b = r.split(" ");
  if (a.length !== b.length) throw new Error("Wortzahl verschieden: " + r);
  const out = [];
  b.forEach((w, i) => {
    const v = a[i];
    const kommaNeu = w.endsWith(",") && !v.endsWith(",");
    const wort = kommaNeu ? w.slice(0, -1) : w;
    const satz = (wort.match(/[.?!]$/) || [""])[0];
    out.push({ text: satz ? wort.slice(0, -1) : wort, fehler: wort !== v ? "gross" : null });
    if (satz) out.push({ text: satz, fehler: null });
    if (kommaNeu) out.push({ text: ",", fehler: "komma" });
    out.push({ text: i < b.length - 1 ? " " : "", fehler: null });
  });
  return out;
}
const zaehle = (list) => list.reduce((z, m) => { vergleich(m.f, m.r).forEach((x) => { if (x.fehler) z[x.fehler]++; }); return z; }, { gross: 0, komma: 0 });

const REGELN = [
  { titel: "Großschreibung", kern: "Substantive schreibt man groß.", punkte: ["Probe: Passt der, die, das oder ein davor? → groß: der Vater, die Angst", "Auch Verben werden groß, wenn beim, zum oder das davorsteht: beim Zuhören, das Schreiben"] },
  { titel: "Komma zwischen Sätzen", kern: "Vor dass, weil, wenn und ob steht ein Komma.", punkte: ["Ich glaube, dass … · Ich wünsche mir, dass …", "Auch ohne dass: Ich hoffe, er hört mir zu. Probe: Passt „dass“ hinein? → Komma", "Steht der Nebensatz vorne, kommt das Komma danach: Wenn ich erzähle, hört er nicht zu.", "Tipp: Verb am Ende → Nebensatz → Komma"] },
];

// ===== Testauswertung: Brief „Pressekonferenz“ (Auswertung vom 08.10., Deutsch_11/Auswertung_Brief) =====
const TEST = {
  gesamt: "18,9 von 30", inhalt: "14,5 von 20", sprache: "4,4 von 10",
  stark: "Beziehung reparieren und Metakommunikation",
  schwach: "Appell: oft nur „Ich hoffe, ihr …“ statt einer klaren Bitte",
  rs: "Rechtschreibung Ø 0,4 von 2 · Zeichensetzung Ø 0,6 von 2",
};
const LERNZIELE = [
  "Ich erkenne Substantive mit der Artikelprobe und schreibe sie groß.",
  "Ich setze Kommas zwischen Haupt- und Nebensatz – auch ohne „dass“.",
  "Ich erkenne offene und verdeckte Appelle in Kafkas Brief und belege sie.",
  "Ich formuliere selbst einen offenen Appell in einer Rede.",
];
const VERLAUF = [
  ["Teil 1", "Auswertung des Briefs und Rechtschreibtraining", "15 Min."],
  ["Teil 2", "Kafka: Was will Kafka von seinem Vater? – der Appell", "30 Min."],
];

// ===== Kafka: Leitfragen (zeigen auf einzelne Wörter) =====
const LEITFRAGEN = [
  ["Was soll der Vater tun?", "Achte auf „bitte“ (Z. 1) und „sollst Du aufhören“ (Z. 3)."],
  ["Was wünscht sich Kafka, ohne es direkt zu fordern?", "Achte auf „Wort“, „Hand“ und „Blick“ (Z. 5 f.)."],
  ["Warum betont Kafka, dass der Vater keine Schuld hat?", "Achte auf „niemals … an eine Schuld Deinerseits“ (Z. 1 f.)."],
];
const FORMULIERUNGSHILFEN = [
  "Ein offener Appell zeigt sich in … (Z. …).",
  "Verdeckt appelliert Kafka, wenn er …",
  "Damit will er erreichen, dass der Vater …",
  "Seine Bitte sichert er ab, indem er …",
];

// ===== Transfer: Rede an den Vater =====
const REDE_SITUATION = "Max Brods Brief an Hermann Kafka war erfolgreich: Der Vater ist bereit, seinem Sohn zuzuhören.";
const REDE_AUFTRAG = "Schreibe als Franz eine Rede an den Vater, in der du an ihn appellierst.";
const REDE_UMFANG = "ca. 100–150 Wörter · allein · in den Hefter";
const KRITERIEN = [
  ["Anrede", "„Vater, …“"],
  ["Ich-Botschaften", "mindestens zwei: Gefühl + Wirkung"],
  ["Offener Appell", "mindestens eine klare Bitte"],
  ["Bezug zum Text", "Kafkas Wünsche: Wort, Hand, Blick"],
  ["Ton", "ehrlich, mutig, respektvoll – keine Vorwürfe"],
];

module.exports = {
  TEXT, QUELLE, GLOSSAR, KONTEXT, KONTEXT_ZITAT, APPELL,
  DEMO, SAETZE, ZUSAMMEN, vergleich, zaehle, REGELN, TEST, LERNZIELE, VERLAUF,
  LEITFRAGEN, FORMULIERUNGSHILFEN, REDE_SITUATION, REDE_AUFTRAG, REDE_UMFANG, KRITERIEN,
};
