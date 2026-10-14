// Gemeinsame Daten der Stunde „Auswertung des Briefs und Kafkas Appell“
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

const KONTEXT = "Franz Kafka (1883–1924) schrieb 1919 mit 36 Jahren einen über 100 Seiten langen Brief an seinen Vater Hermann, einen erfolgreichen, lauten und bestimmenden Kaufmann. Vor ihm fühlte sich der Sohn klein und verstummte oft. Kafka wollte mit dem Brief erreichen, dass sich die Beziehung bessert.";
const KONTEXT_ZITAT = ["… daß es uns beide ein wenig beruhigen und Leben und Sterben leichter machen kann.", "Franz Kafka, letzter Satz des Briefs (S. 20)"];

// Information zum Appell
const APPELL = [
  ["Appell", "Mit dieser Seite einer Nachricht will der Sender erreichen, dass der Empfänger etwas tut, denkt oder fühlt."],
  ["offen", "direkt ausgesprochen, z. B. als Bitte: „Hör mir bitte zu!“"],
  ["verdeckt", "versteckt in einem Wunsch, einer Frage oder einer Beschreibung: „Du hörst mir nie zu.“"],
];

// ===== Testauswertung: Brief „Pressekonferenz“ (Auswertung vom 08.10., Deutsch_11/Auswertung_Brief) =====
const TEST = {
  gesamt: "18,9 von 30", inhalt: "14,5 von 20", sprache: "4,4 von 10",
  stark: "Beziehung reparieren und Metakommunikation",
  schwach: "Appell: oft nur „Ich hoffe, ihr …“ statt einer klaren Bitte",
  rs: "Rechtschreibung Ø 0,4 von 2 · Zeichensetzung Ø 0,6 von 2",
};
const LERNZIELE = [
  "Ich weiß, wo ich im Brief Punkte gewonnen und verloren habe.",
  "Ich erkenne offene und verdeckte Appelle in Kafkas Brief und belege sie.",
  "Ich formuliere selbst einen offenen Appell in einer Rede.",
];
const VERLAUF = [
  ["Teil 1", "Auswertung des Briefs „Pressekonferenz“", "5 Min."],
  ["Teil 2", "Kafka: Was will Kafka von seinem Vater? – der Appell", "40 Min."],
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
  TEST, LERNZIELE, VERLAUF,
  LEITFRAGEN, FORMULIERUNGSHILFEN, REDE_SITUATION, REDE_AUFTRAG, REDE_UMFANG, KRITERIEN,
};
