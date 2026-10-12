// Gemeinsame Daten der Stunde „Kafkas Appell“: Auszug, Fußnoten, Mini-Diktat, Leitfragen, Redekriterien
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

// Mini-Diktat: kurz, Pausen an „ / “
const DIKTAT = [
  "Der Appell ist eine der vier Seiten einer Nachricht.",
  "Mit ihm will der Sender erreichen, / dass der Empfänger etwas tut, denkt oder fühlt.",
  "Ein offener Appell wird direkt ausgesprochen, / zum Beispiel als Bitte.",
  "Ein verdeckter Appell versteckt sich / in einem Wunsch, einer Frage oder einer Beschreibung.",
];
const clean = (arr) => arr.map((s) => s.replace(/ \/ /g, " ")).join(" ");
const STOLPER = [
  ["Appell", "zwei p, zwei l – und groß, weil Nomen"],
  ["dass", "leitet den Nebensatz ein: „…, dass der Empfänger …“ – davor ein Komma"],
  ["Empfänger · ausgesprochen · versteckt", "ä in Empfänger (von empfangen); ck in versteckt"],
  ["Aufzählung", "Komma zwischen „tut, denkt oder fühlt“ – aber keins vor „oder“"],
];

// Leitfragen zur Analyse (zeigen auf einzelne Wörter)
const LEITFRAGEN = [
  ["„bitte“ (Z. 1) · „sollst Du aufhören“ (Z. 3)", "Was genau soll der Vater tun?"],
  ["„ein freundliches Wort, ein stilles Bei-der-Hand-Nehmen, ein guter Blick“ (Z. 8 f.)", "Was wünscht sich Kafka – ohne es direkt zu fordern?"],
  ["„niemals … an eine Schuld Deinerseits“ (Z. 2)", "Warum sichert Kafka seine Bitte so ab?"],
];

// Transfer: Rede an den Vater
const REDE_AUFTRAG = "Max Brods Brief hat gewirkt: Hermann Kafka ist bereit, seinem Sohn zuzuhören. Schreibe die Rede, die Franz jetzt an seinen Vater hält.";
const REDE_UMFANG = "8–10 Sätze · allein · in den Hefter";
const KRITERIEN = [
  ["Anrede", "Franz spricht den Vater direkt an („Vater, …“)."],
  ["Ich-Botschaften", "mindestens zwei: Gefühl + Wirkung („Ich fühle mich …, wenn Du …“)."],
  ["Offener Appell", "eine klare Bitte: Was soll der Vater künftig tun?"],
  ["Bezug zum Text", "Kafkas Wünsche aus dem Brief kommen vor (Wort, Hand, Blick)."],
  ["Respektvoller Ton", "keine Vorwürfe – Kafka sucht keine Schuld."],
];

module.exports = { TEXT, QUELLE, GLOSSAR, KONTEXT, DIKTAT, clean, STOLPER, LEITFRAGEN, REDE_AUFTRAG, REDE_UMFANG, KRITERIEN };
