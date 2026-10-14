// Gemeinsame Daten der Stunde „Kafkas Appell“: Rechtschreib-Check (Chat, Regeln), Auszug, Fußnoten, Leitfragen, Redekriterien
// Auszug: Kafka, „Brief an den Vater“, S. 3 (DigBib-Ausgabe), mit dem Original abgeglichen
const TEXT = [
  "Wobei ich Dich aber immerfort bitte, nicht zu vergessen, daß ich",
  "niemals im entferntesten an eine Schuld Deinerseits glaube. Du",
  "wirktest so auf mich, wie Du wirken mußtest, nur sollst Du aufhören,",
  "es für eine besondere Bosheit meinerseits zu halten, daß ich dieser",
  "Wirkung erlegen bin. Ich war ein ängstliches Kind; trotzdem war ich",
  "gewiß auch störrisch, wie Kinder sind; gewiß verwöhnte mich die Mutter",
  "auch, aber ich kann nicht glauben, daß ich besonders schwer lenkbar",
  "war, ich kann nicht glauben, daß ein freundliches Wort, ein stilles",
  "Bei-der-Hand-Nehmen, ein guter Blick mir nicht alles hätten abfordern",
  "können, was man wollte.",
];
const QUELLE = "Franz Kafka: „Brief an den Vater“ (1919), S. 3";

// Fußnoten: [Wort im Text, Erklärung]
const GLOSSAR = [
  ["immerfort", "immer wieder, ständig"],
  ["Deinerseits", "von Deiner Seite"],
  ["Bosheit", "böse Absicht"],
  ["erlegen", "nicht widerstehen können"],
  ["störrisch", "trotzig, eigensinnig"],
  ["lenkbar", "zu lenken, zu erziehen"],
  ["abfordern", "abverlangen"],
];

const KONTEXT = "Franz Kafka (1883–1924) schrieb 1919 mit 36 Jahren einen über 100 Seiten langen Brief an seinen Vater Hermann, einen erfolgreichen, lauten und bestimmenden Kaufmann. Vor ihm fühlte sich der Sohn klein und verstummte oft. Den Brief gab Kafka seiner Mutter – sie reichte ihn nie weiter.";

// Information zum Appell (ersetzt das Diktat; Folie in 27 pt)
const APPELL = [
  ["Appell", "Mit dieser Seite einer Nachricht will der Sender erreichen, dass der Empfänger etwas tut, denkt oder fühlt."],
  ["offen", "direkt ausgesprochen, z. B. als Bitte: „Hör mir bitte zu!“"],
  ["verdeckt", "versteckt in einem Wunsch, einer Frage oder einer Beschreibung: „Du hörst mir nie zu.“"],
];

// Rechtschreib-Check: WhatsApp-Chat mit Fehlern (nur Großschreibung von Substantiven und Kommas vor Nebensätzen)
// Fehlerfassung und richtige Fassung haben gleich viele Wörter – die Unterschiede werden automatisch markiert.
const CHAT = [
  { von: "Mia", seite: "l", zeit: "18:41", f: "Ich glaube dass der vater von Kafka einfach keine geduld hatte.", r: "Ich glaube, dass der Vater von Kafka einfach keine Geduld hatte." },
  { von: "Jonas", seite: "r", zeit: "18:43", f: "Ja, aber Kafka hatte auch angst. Er schreibt dass er als kind nur ein freundliches wort gebraucht hätte.", r: "Ja, aber Kafka hatte auch Angst. Er schreibt, dass er als Kind nur ein freundliches Wort gebraucht hätte." },
  { von: "Mia", seite: "l", zeit: "18:44", f: "Das finde ich traurig. Ich wünsche mir dass mein vater beim zuhören auch mal das handy weglegt.", r: "Das finde ich traurig. Ich wünsche mir, dass mein Vater beim Zuhören auch mal das Handy weglegt." },
  { von: "Jonas", seite: "r", zeit: "18:46", f: "Haha, same. Wenn ich was erzähle schaut meiner nur auf den fernseher.", r: "Haha, same. Wenn ich was erzähle, schaut meiner nur auf den Fernseher." },
  { von: "Mia", seite: "l", zeit: "18:47", f: "Weißt du ob wir morgen die rede halten müssen?", r: "Weißt du, ob wir morgen die Rede halten müssen?" },
  { von: "Jonas", seite: "r", zeit: "18:47", f: "Keine ahnung. Ich hoffe nur dass ich nicht zuerst dran bin.", r: "Keine Ahnung. Ich hoffe nur, dass ich nicht zuerst dran bin." },
];
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
const zaehle = (msgs) => msgs.reduce((z, m) => { vergleich(m.f, m.r).forEach((x) => { if (x.fehler) z[x.fehler]++; }); return z; }, { gross: 0, komma: 0 });

const REGELN = [
  { titel: "Großschreibung", kern: "Substantive schreibt man groß.", punkte: ["Probe: Passt der, die, das oder ein davor? → groß: der Vater, die Angst", "Auch Verben werden groß, wenn beim, zum oder das davorsteht: beim Zuhören, das Schreiben"] },
  { titel: "Komma vor Nebensätzen", kern: "Vor dass, weil, wenn und ob steht ein Komma.", punkte: ["Ich glaube, dass … · Ich wünsche mir, dass …", "Steht der Nebensatz vorne, kommt das Komma danach: Wenn ich erzähle, schaut er weg.", "Tipp: Verb am Ende → Nebensatz → Komma"] },
];

// Leitfragen zur Analyse (zeigen auf einzelne Wörter)
const LEITFRAGEN = [
  ["„bitte“ (Z. 1) · „sollst Du aufhören“ (Z. 3)", "Was genau soll der Vater tun?"],
  ["„ein freundliches Wort, ein stilles Bei-der-Hand-Nehmen, ein guter Blick“ (Z. 8 f.)", "Was wünscht sich Kafka – ohne es direkt zu fordern?"],
  ["„niemals … an eine Schuld Deinerseits“ (Z. 2)", "Warum sichert Kafka seine Bitte so ab?"],
];

// Transfer: Rede an den Vater
const REDE_AUFTRAG = "Max Brods Brief hat gewirkt: Der Vater hört zu. Schreibe die Rede, die Franz jetzt an ihn hält.";
const REDE_UMFANG = "6–8 Sätze · allein · in den Hefter";
const KRITERIEN = [
  ["Anrede", "„Vater, …“"],
  ["Ich-Botschaften", "zwei: Gefühl + Wirkung"],
  ["Offener Appell", "eine klare Bitte"],
  ["Bezug zum Text", "Wort, Hand, Blick"],
  ["Respektvoller Ton", "keine Vorwürfe"],
];

module.exports = { TEXT, QUELLE, GLOSSAR, KONTEXT, APPELL, CHAT, vergleich, zaehle, REGELN, LEITFRAGEN, REDE_AUFTRAG, REDE_UMFANG, KRITERIEN };
