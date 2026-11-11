// Gemeinsame Daten der Stunde „Vor dem Gesetz“ – kafkaesk, Inhalt und Symbole der Parabel (Deutsch 11, Do 12.11., 45 Min.)
// Text aus dem Gedächtnis wiedergegeben (Kafka, 1915; gemeinfrei), Schreibung an die heutige Rechtschreibung angepasst.
// Die Textquellen im Netz waren bei der Erstellung gesperrt – bitte vor dem Druck mit einer Ausgabe abgleichen.
const TITEL = "Vor dem Gesetz";
const AUTOR = "Franz Kafka";
const TEXT = [
  "Vor dem Gesetz steht ein Türhüter. Zu diesem Türhüter kommt ein Mann vom Lande und bittet um Eintritt in das Gesetz. Aber der Türhüter sagt, dass er ihm jetzt den Eintritt nicht gewähren könne. Der Mann überlegt und fragt dann, ob er also später werde eintreten dürfen. „Es ist möglich“, sagt der Türhüter, „jetzt aber nicht.“",
  "Da das Tor zum Gesetz offensteht wie immer und der Türhüter beiseite tritt, bückt sich der Mann, um durch das Tor in das Innere zu sehn. Als der Türhüter das merkt, lacht er und sagt: „Wenn es dich so lockt, versuche es doch, trotz meines Verbotes hineinzugehn. Merke aber: Ich bin mächtig. Und ich bin nur der unterste Türhüter. Von Saal zu Saal stehn aber Türhüter, einer mächtiger als der andere. Schon den Anblick des dritten kann nicht einmal ich mehr ertragen.“",
  "Solche Schwierigkeiten hat der Mann vom Lande nicht erwartet; das Gesetz soll doch jedem und immer zugänglich sein, denkt er, aber als er jetzt den Türhüter in seinem Pelzmantel genauer ansieht, seine große Spitznase, den langen, dünnen, schwarzen tatarischen Bart, entschließt er sich, doch lieber zu warten, bis er die Erlaubnis zum Eintritt bekommt.",
  "Der Türhüter gibt ihm einen Schemel und lässt ihn seitwärts von der Tür sich niedersetzen. Dort sitzt er Tage und Jahre. Er macht viele Versuche, eingelassen zu werden, und ermüdet den Türhüter durch seine Bitten. Der Türhüter stellt öfters kleine Verhöre mit ihm an, fragt ihn über seine Heimat aus und nach vielem andern, es sind aber teilnahmslose Fragen, wie sie große Herren stellen, und zum Schlusse sagt er ihm immer wieder, dass er ihn noch nicht einlassen könne.",
  "Der Mann, der sich für seine Reise mit vielem ausgerüstet hat, verwendet alles, und sei es noch so wertvoll, um den Türhüter zu bestechen. Dieser nimmt zwar alles an, aber sagt dabei: „Ich nehme es nur an, damit du nicht glaubst, etwas versäumt zu haben.“",
  "Während der vielen Jahre beobachtet der Mann den Türhüter fast ununterbrochen. Er vergisst die andern Türhüter, und dieser erste scheint ihm das einzige Hindernis für den Eintritt in das Gesetz. Er verflucht den unglücklichen Zufall, in den ersten Jahren rücksichtslos und laut, später, als er alt wird, brummt er nur noch vor sich hin. Er wird kindisch, und da er in dem jahrelangen Studium des Türhüters auch die Flöhe in seinem Pelzkragen erkannt hat, bittet er auch die Flöhe, ihm zu helfen und den Türhüter umzustimmen.",
  "Schließlich wird sein Augenlicht schwach, und er weiß nicht, ob es um ihn wirklich dunkler wird, oder ob ihn nur seine Augen täuschen. Wohl aber erkennt er jetzt im Dunkel einen Glanz, der unverlöschlich aus der Türe des Gesetzes bricht. Nun lebt er nicht mehr lange.",
  "Vor seinem Tode sammeln sich in seinem Kopfe alle Erfahrungen der ganzen Zeit zu einer Frage, die er bisher an den Türhüter noch nicht gestellt hat. Er winkt ihm zu, da er seinen erstarrenden Körper nicht mehr aufrichten kann. Der Türhüter muss sich tief zu ihm hinunterneigen, denn der Größenunterschied hat sich sehr zuungunsten des Mannes verändert. „Was willst du denn jetzt noch wissen?“, fragt der Türhüter, „du bist unersättlich.“ „Alle streben doch nach dem Gesetz“, sagt der Mann, „wieso kommt es, dass in den vielen Jahren niemand außer mir Einlass verlangt hat?“ Der Türhüter erkennt, dass der Mann schon an seinem Ende ist, und, um sein vergehendes Gehör noch zu erreichen, brüllt er ihn an: „Hier konnte niemand sonst Einlass erhalten, denn dieser Eingang war nur für dich bestimmt. Ich gehe jetzt und schließe ihn.“",
];
const TEXT_HINWEIS = "Text aus dem Gedächtnis wiedergegeben und an die heutige Rechtschreibung angepasst (dass, lässt, muss; „sehn“, „hineinzugehn“, „stehn“ wie im Original). Die Textquellen im Netz waren bei der Erstellung gesperrt – bitte vor dem Druck mit einer Ausgabe abgleichen (z. B. Reclam, „Ein Landarzt“ bzw. „Der Process“, Kapitel „Im Dom“). Bei Abweichungen den Text in daten.js korrigieren und neu erzeugen; die Zeilenangaben passen sich an.";

// Text mit festen Zeilen (für das Arbeitsblatt mit Zeilennummern); Absätze beginnen in einer neuen Zeile
function umbrechen(absaetze, breite = 93) {
  const zeilen = [];
  absaetze.forEach((a) => {
    let akt = "";
    a.split(" ").forEach((w) => {
      if ((akt + " " + w).trim().length > breite) { zeilen.push(akt); akt = w; } else akt = (akt + " " + w).trim();
    });
    if (akt) zeilen.push(akt);
  });
  return zeilen;
}
const ZEILEN = umbrechen(TEXT);
// Zeilenangabe zu einem Zitat (Zeile, in der der Anfang steht; ggf. „f.“)
function zeile(zitat) {
  const rest = zitat.replace(/[„“…]/g, "").trim();
  const kern = rest.split(" ").slice(0, 3).join(" ");
  const i = ZEILEN.findIndex((z, j) => { const k = (z + " " + (ZEILEN[j + 1] || "")).indexOf(kern); return k >= 0 && k < z.length; });
  if (i < 0) throw new Error("Zitat nicht im Text: " + zitat);
  return "Z. " + (i + 1) + (ZEILEN[i].includes(rest) ? "" : " f.");
}

// Einstiegszitat (Ankommen) – aus Kafkas Nachlass (Notizen), Wortlaut bitte prüfen
const ZITAT = {
  text: "Auf Balzacs Spazierstockgriff: Ich breche alle Hindernisse. Auf meinem: Mich brechen alle Hindernisse. Gemeinsam ist das ‚alle‘.",
  quelle: "Franz Kafka, Notiz aus dem Nachlass",
  impuls: "Was sagt dieser Satz über Kafka?",
  erwartung: "Balzac (franz. Schriftsteller, 19. Jh.) sieht sich als Sieger über alle Hindernisse; Kafka dreht den Satz um: Er ist das Opfer, die Hindernisse sind stärker. Subjekt und Objekt tauschen die Rolle – Kafka erlebt sich als ohnmächtig, er hat resigniert. Bitterer Witz: „Gemeinsam ist das ‚alle‘.“",
};
// Kontext: erster Satz des „Process“ (Ohnmacht vor staatlichen Autoritäten)
const PROCESS_SATZ = "Jemand musste Josef K. verleumdet haben, denn ohne dass er etwas Böses getan hätte, wurde er eines Morgens verhaftet.";

// Lernziele (Ich kann …)
const LERNZIELE = [
  ["Ich kann den Inhalt von „Vor dem Gesetz“ in Handlungsschritten wiedergeben.", false],
  ["Ich kann die Symbole der Parabel deuten (Bildebene → Sachebene) und am Text belegen.", false],
  ["Ich kann erklären, warum die Parabel „kafkaesk“ ist.", false],
];
const VERLAUF = [
  ["Ein Zitat von Kafka", "Ankommen"],
  ["Wiederholung: „kafkaesk“ – Kafka und die Macht", "Vorwissen"],
  ["„Vor dem Gesetz“ lesen: Worum geht es?", "Informieren"],
  ["Handlungsschritte ordnen · Symbole deuten", "Verarbeiten"],
  ["Deutung · Transfer* · Zielcheck", "Auswerten"],
];

// Vorwissen: „kafkaesk“
const VORWISSEN = [
  ["Was bedeutet „kafkaesk“?", "Erklärt das Wort mit einem Beispiel."],
  ["Wem stand Kafka ohnmächtig gegenüber?", "Denkt an den „Brief an den Vater“ und an „Gibs auf!“."],
];
const KAFKAESK = {
  duden: "kafkaesk: in der Art der Darstellungen Kafkas; auf unergründliche Weise bedrohlich (nach Duden)",
  merkmale: [
    "Ein Einzelner steht einer übermächtigen Instanz gegenüber (Behörde, Gericht, Vater).",
    "Ihre Regeln bleiben undurchschaubar – niemand erklärt sich.",
    "Der Einzelne ist ohnmächtig und fügt sich am Ende (Resignation).",
    "Alles wird nüchtern erzählt – und wirkt doch wie ein Albtraum.",
  ],
};
const KAFKA_MACHT = [
  ["Vater", "„Brief an den Vater“: übermächtig – Kafka fühlt sich klein und schuldig"],
  ["Schutzmann", "„Gibs auf!“: Wer helfen sollte, lacht und wendet sich ab"],
  ["Behörde", "Jurist in einer Versicherungsbehörde in Prag (1908–1922): Bürokratie von innen"],
];
const KONTEXT = [
  "Die Parabel ist ein Teil des Romans „Der Process“ (1914/15 geschrieben).",
  "Josef K. wird verhaftet, ohne zu erfahren, warum. Ein Gericht, das er nie zu sehen bekommt, verhandelt gegen ihn.",
  "Im Dom erzählt ihm ein Geistlicher die Geschichte vom Mann vor dem Gesetz. 1915 hat Kafka sie auch einzeln veröffentlicht.",
];
const LESEAUFTRAG = "Lest mit. Achtet darauf: Was will der Mann – und was hindert ihn?";
const KERNBOTSCHAFT = "Ein Mann vom Lande wartet sein Leben lang vergeblich auf Einlass in das Gesetz – und erfährt erst im Sterben, dass der Eingang nur für ihn bestimmt war.";

// Aufgabe 1: Handlungsschritte ordnen (richtige Reihenfolge; angezeigt wird gemischt nach REIHE_GEMISCHT)
const SCHRITTE = [
  ["Ein Mann vom Lande bittet den Türhüter um Eintritt in das Gesetz.", "Zu diesem Türhüter kommt ein Mann vom Lande und bittet"],
  ["Der Türhüter verweigert den Eintritt: „jetzt aber nicht“.", "jetzt aber nicht"],
  ["Der Türhüter warnt vor seiner Macht und vor noch mächtigeren Türhütern.", "Ich bin mächtig"],
  ["Der Mann wartet jahrelang auf einem Schemel und versucht, den Türhüter zu bestechen.", "Der Türhüter gibt ihm einen Schemel und lässt ihn seitwärts von der Tür sich niedersetzen"],
  ["Kurz vor seinem Tod fragt er, warum niemand sonst Einlass verlangt hat.", "wieso kommt es"],
  ["Der Türhüter antwortet, der Eingang sei nur für ihn bestimmt gewesen, und schließt ihn.", "Hier konnte niemand sonst Einlass erhalten, denn"],
];
const REIHE_GEMISCHT = [3, 5, 0, 2, 4, 1]; // Buchstaben A–F in dieser Reihenfolge
const AUFGABE1 = {
  titel: "Aufgabe 1: Handlungsschritte ordnen",
  zeit: "3 Min.", form: "allein, auf dem Arbeitsblatt",
  schritte: [
    "Lies die sechs Handlungsschritte.",
    "Nummeriere sie in der richtigen Reihenfolge (1–6).",
    "Notiere zu jedem Schritt die Zeile im Text.",
  ],
};

// Aufgabe 2: Symbole deuten – [Bildelement, Achte auf (Wort im Text), Zitat für die Zeile, Sachebene (Lösung), optional]
const SYMBOLE = [
  ["das offene Tor", "„offensteht wie immer“", "Da das Tor zum Gesetz offensteht", "Der Zugang scheint möglich, das Ziel ist sichtbar – und bleibt doch unerreichbar: Hoffnung und Verbot zugleich.", false],
  ["der Türhüter", "„nur der unterste Türhüter“", "Merke aber: Ich bin mächtig", "Eine Autorität (Beamter, Behörde, Staat), die den Zugang verwaltet. Ihre Macht kommt aus einer unsichtbaren Hierarchie – Gründe nennt sie nie.", false],
  ["das Gesetz", "„jedem und immer zugänglich“", "das Gesetz soll doch jedem", "Recht und Gerechtigkeit – offen gedeutet auch Wahrheit, Sinn des Lebens oder Gott: das, wonach jeder Mensch strebt.", false],
  ["der Mann vom Lande", "„vom Lande“", "kommt ein Mann vom Lande", "Ein einfacher Mensch, der die Regeln der Institution nicht kennt – der Einzelne, jeder von uns.", false],
  ["das Warten auf dem Schemel", "„Tage und Jahre“", "Dort sitzt er Tage und Jahre", "Das Leben vergeht mit Warten. Er gehorcht, statt zu handeln – Ohnmacht wird zur Resignation.", false],
  ["die Bestechung", "„nur …, damit du nicht glaubst, etwas versäumt zu haben“", "Ich nehme es nur an", "Alle Mittel sind wirkungslos; die Macht lässt sich nicht beeinflussen. Sie nimmt alles – und gibt nur den Schein von Hoffnung.", false],
  ["die Flöhe im Pelzkragen", "„bittet auch die Flöhe“", "bittet er auch die Flöhe", "Erniedrigung und Verlust der Würde: Er sucht Hilfe beim Kleinsten und verliert das eigentliche Ziel aus dem Blick.", true],
  ["der Glanz aus der Tür", "„unverlöschlich“, „jetzt im Dunkel“", "Wohl aber erkennt er jetzt", "Wahrheit, Erkenntnis, Hoffnung – erst im Sterben sichtbar, zu spät.", true],
  ["das Schließen des Eingangs", "„nur für dich bestimmt“", "Hier konnte niemand sonst", "Der Weg war ganz persönlich; mit seinem Leben ist die Chance vorbei. Paradox: Er selbst hat sie durch sein Warten verspielt.", true],
];
const AUFGABE2 = {
  titel: "Aufgabe 2: Symbole deuten",
  zeit: "9 Min.", form: "allein, auf dem Arbeitsblatt",
  schritte: [
    "Suche die Stelle im Text und notiere die Zeile.",
    "Achte auf die angegebenen Wörter.",
    "Deute in Stichworten: Wofür steht das Bildelement?",
    "Fertig? Bearbeite die Zeilen mit Stern (*).",
  ],
};
// Ich mache vor: das offene Tor (lautes Denken in drei Schritten)
const VORMACHEN = [
  ["1. Was steht da?", "„Da das Tor zum Gesetz offensteht wie immer …“ – das Tor ist offen, immer schon."],
  ["2. Was fällt auf?", "Offen – und trotzdem darf er nicht hinein. Das passt nicht zusammen."],
  ["3. Wofür steht es?", "Der Zugang scheint möglich, das Ziel ist sichtbar – und bleibt doch unerreichbar: Hoffnung und Verbot zugleich."],
];

// Deutung: Ohnmacht von außen und von innen; zurück zum Zitat
const DEUTUNG = {
  frage: "Warum kommt der Mann nicht hinein?",
  aussen: "Von außen: Der Türhüter verbietet – hinter ihm eine unsichtbare Hierarchie von Mächtigeren.",
  innen: "Von innen: Verboten war es nur „jetzt“. Er hätte es versuchen können – er wartet lieber und fügt sich.",
  zitat: "„Mich brechen alle Hindernisse.“ – Kafkas Satz passt zum Mann vom Lande: Ohnmacht wird zur Resignation.",
};
// Transfer* (mündlich)
const TRANSFER = {
  titel: "Transfer *",
  fragen: [
    "Wo steht ihr vor einem „Tor“ mit einem „Türhüter“? (Kader, Nominierung, Behörde, Schule …)",
    "Was würdet ihr anders machen als der Mann vom Lande?",
  ],
};
// Zielcheck (im Heft)
const ZIELCHECK = {
  auftrag: "Erkläre in 2–3 Sätzen, warum „Vor dem Gesetz“ kafkaesk ist. Beziehe dich auf ein Symbol mit Zeilenangabe.",
  geruest: "Der / Das … (Z. …) steht für … . Das ist kafkaesk, weil …",
  kriterien: ["Symbol mit Zeile", "Deutung auf der Sachebene", "ein Merkmal von „kafkaesk“ (übermächtig, undurchschaubar, Ohnmacht, Resignation)"],
};
ZIELCHECK.muster = "Der Türhüter (" + "Z. ?" + ") steht für eine staatliche Autorität, die den Zugang zum Recht verwaltet. Das ist kafkaesk, weil er keinen Grund nennt („jetzt aber nicht“) und hinter ihm eine unsichtbare Hierarchie steht: Der Mann ist dieser Macht ohnmächtig ausgeliefert und wartet sein Leben lang.";
ZIELCHECK.muster = ZIELCHECK.muster.replace("Z. ?", zeile("Vor dem Gesetz steht ein Türhüter"));

module.exports = {
  TITEL, AUTOR, TEXT, TEXT_HINWEIS, ZEILEN, zeile, ZITAT, PROCESS_SATZ, LERNZIELE, VERLAUF, VORWISSEN, KAFKAESK, KAFKA_MACHT,
  KONTEXT, LESEAUFTRAG, KERNBOTSCHAFT, SCHRITTE, REIHE_GEMISCHT, AUFGABE1, SYMBOLE, AUFGABE2, VORMACHEN, DEUTUNG, TRANSFER, ZIELCHECK,
};
