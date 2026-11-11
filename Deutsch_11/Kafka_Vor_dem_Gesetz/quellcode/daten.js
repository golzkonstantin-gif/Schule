// Gemeinsame Daten der Stunde „Vor dem Gesetz“ – kafkaesk, Inhalt und Symbole der Parabel (Deutsch 11, Do 12.11., 45 Min.; schlanke Fassung nach Rückmeldung 11.11.)
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

// Einstiegszitat (Ankommen) – Reaktivierung des Wissens über Kafka; „kafkaesk“ kommt erst danach
// Tagebuch, 21.7.1913 (Liste für und gegen die Heirat mit Felice Bauer) – Datum und Wortlaut bitte mit einer Ausgabe abgleichen
const ZITAT = {
  text: ["Ich werde mich bis zur Besinnungslosigkeit von allen absperren.", "Mit allen mich verfeinden, mit niemandem reden."],
  quelle: "Franz Kafka, Tagebuch, 21. Juli 1913",
  impulse: [
    "„Stifte liegen. Lest den Satz still. Wie wirkt er auf euch?“",
    "„Der Satz ist von Kafka. Wie kann jemand so etwas über sich sagen? Was wisst ihr über Kafka, das es erklärt?“",
  ],
  kontext: "Tagebuch, 21.7.1913: Kafka (30) überlegt, ob er Felice Bauer heiraten soll, und fürchtet, dass er dann nicht mehr schreiben kann.",
  wissen: [
    ["Vater", "übermächtig – Kafka fühlt sich klein und schuldig („Brief an den Vater“)"],
    ["Schreiben", "braucht Einsamkeit – sein eigentliches Leben"],
    ["Beruf", "Jurist in einer Versicherungsbehörde – das Büro als Last"],
    ["Beziehungen", "Verlobungen, Angst vor Bindung"],
  ],
  erwartung: "Eindruck: radikal, verzweifelt, abweisend. Erklärung aus Kafkas Leben: Er will schreiben und braucht dafür völlige Einsamkeit; er fühlt sich beherrscht (Vater, Büro, Erwartung zu heiraten) – der Rückzug ist sein Selbstschutz. Wenn es aus der Klasse kommt: Prag, deutschsprachige Minderheit, Krankheit.",
  gelenk: "„Kafka hat eine ganz eigene Weltsicht. Für Situationen, wie er sie beschreibt, gibt es ein eigenes Wort. Was bedeutet ‚kafkaesk‘?“",
};
const PROCESS_SATZ = "Jemand musste Josef K. verleumdet haben, denn ohne dass er etwas Böses getan hätte, wurde er eines Morgens verhaftet.";

// Lernziele (Ich kann …)
const LERNZIELE = [
  ["Ich kann die Handlung über das Gespräch erschließen und erklären, warum die Kommunikation scheitert.", false],
  ["Ich kann die zentralen Symbole deuten und erklären, warum die Parabel kafkaesk ist.", false],
];
const VERLAUF = [
  ["Ein Zitat: Was wissen wir über Kafka?", "Ankommen"],
  ["„kafkaesk“: Wiederholung · Begriffs-Sprint · nervig oder kafkaesk?", "Vorwissen"],
  ["„Vor dem Gesetz“ lesen · Einleitungssatz mit Hypothese", "Informieren"],
  ["Das Gespräch Schritt für Schritt · Symbole deuten", "Verarbeiten"],
  ["Zielcheck", "Auswerten"],
];

// Vorwissen: „kafkaesk“ mündlich wiederholen (Begriffe der Mentorin, AB vom 9.11.)
const VORWISSEN = [
  ["Was bedeutet „kafkaesk“?", "Erklärt das Wort mit eigenen Worten – denkt an das Arbeitsblatt vom Montag."],
];
const KAFKAESK = {
  definition: "kafkaesk: Situationen, die von einer rätselhaften, absurden und bürokratischen Übermacht geprägt sind. Der Einzelne ist einer anonymen, undurchschaubaren Gewalt ausgeliefert.",
  bausteine: [
    ["labyrinthische Bürokratie", "endlose Wege, unverständliche Regeln, niemand ist zuständig"],
    ["grundlose Schuld", "beschuldigt oder bestraft, ohne zu erfahren, warum"],
    ["vergebliche Logik", "je vernünftiger man handelt, desto tiefer verstrickt man sich"],
    ["unheimliche Normalität", "alle finden das Absurde völlig normal"],
  ],
  gefuehle: "Ohnmacht und Hilflosigkeit · Angst und Isolation · Verwirrung und Orientierungslosigkeit",
};
// Begriffs-Sprint (Regeln wie beim Sprint): 1 Minute ununterbrochen schreiben, danach lesen 2–3 vor, dann Lösung an die Tafel
const SPRINT = {
  regeln: [
    "Stift in die Hand, Heft auf.",
    "Auf mein Zeichen erscheint ein Begriff.",
    "Schreibt eine Minute ohne Pause alles, was euch dazu einfällt.",
    "Der Stift bleibt auf dem Papier – bis „Stopp“.",
  ],
  begriff: "kafkaesk",
};

// Aufgabe 1: Nervig oder kafkaesk? – [Stichwort, Situation, N/K, Baustein bzw. Begründung]
const NK_KRITERIUM = [
  ["nervig", "ärgerlich – aber der Grund ist klar, und du kannst etwas tun"],
  ["kafkaesk", "eine übermächtige Instanz, undurchschaubar, kein Ausweg – du bist ohnmächtig"],
];
const SZENARIEN = [
  ["Kursfahrt", "Deine Zulassung zur Kursfahrt fehlt. Das Sekretariat schickt dich zur Oberstufenkoordination, die schickt dich zur Schulleitung, die Schulleitung zurück ins Sekretariat. Am nächsten Morgen hängt eine Liste am Brett: „Nicht zugelassen“. Hinter deinem Namen steht als Grund: „siehe Antrag“.", "K", "labyrinthische Bürokratie", "Du wirst im Kreis geschickt, niemand ist zuständig, der Grund verweist auf sich selbst."],
  ["Paketbox", "Die Paketbox ist voll. Dein Paket liegt beim Nachbarn, du kannst es heute Abend abholen.", "N", "", "ärgerlich, aber der Grund ist klar, und du kannst etwas tun."],
  ["Klausur", "Laut Liste hast du eine Klausur verpasst. Du zeigst die benotete Arbeit. Die Antwort: „Wer eine Note hat, muss mitgeschrieben haben. Also hast du zweimal geschrieben – Täuschungsversuch.“", "K", "vergebliche Logik", "Gerade der Beweis macht dich schuldig – je logischer, desto tiefer die Verstrickung."],
  ["Grauer Mantel", "Vor dem Training wartet ein Mann in einem grauen Mantel: „Ihr Verfahren läuft. Was Ihnen vorgeworfen wird, darf ich nicht sagen. Trainieren Sie ruhig weiter.“ Seitdem verstummen alle, wenn du die Kabine betrittst.", "K", "grundlose Schuld", "Ein Verfahren ohne Vorwurf – du erfährst nie, warum."],
  ["Extrarunden", "Nach einem Fehlstart lässt dich der Trainer zehn Extrarunden laufen. So steht es in den Teamregeln.", "N", "", "Regel und Grund sind bekannt."],
  ["Letzte Reihe", "Seit Wochen sitzt ein Fremder in der letzten Reihe deines Kurses und schreibt mit, was du sagst. Als du fragst, wer das ist, wundern sich alle: „Der sitzt da doch immer.“", "K", "unheimliche Normalität", "Niemand außer dir findet das seltsam."],
];
const NK_AUFGABE = {
  titel: "Aufgabe 1: Nervig oder kafkaesk?",
  zeit: "4 Min.", form: "allein, auf dem Blatt",
  schritte: [
    "Lies die sechs Situationen.",
    "Kreuze an: nervig oder kafkaesk?",
    "Begründe in Stichworten mit einem Merkmal von „kafkaesk“.",
  ],
};

// Informieren: Kontext und Lesen
const KONTEXT = [
  "Die Parabel ist ein Teil des Romans „Der Process“ (1914/15 geschrieben, 1915 einzeln veröffentlicht).",
  "Im Dom erzählt ein Geistlicher sie Josef K. Der Roman beginnt so:",
];
const LESEAUFTRAG = "Lest mit. Achtet darauf: Wer spricht wie – und wer hat die Macht?";

// Aufgabe 2: Einleitungssatz (TATT) mit Deutungshypothese – gemeinsam
const TATT = {
  titel: "Aufgabe 2: Einleitungssatz mit Deutungshypothese",
  zeit: "3 Min.", form: "gemeinsam, mitschreiben auf dem Arbeitsblatt",
  bausteine: "Textart · Autor · Titel · Entstehungszeit · Thema + Deutungshypothese",
  geruest: "In der Parabel „…“ von … (…) geht es um … . Es lässt sich vermuten, dass …",
  satz: "In der Parabel „Vor dem Gesetz“ von Franz Kafka (1915) geht es um einen Mann vom Lande, der sein Leben lang vergeblich um Einlass in das Gesetz bittet.",
  hypothese: "Es lässt sich vermuten, dass die Kommunikation scheitert, weil ein asymmetrisches (komplementäres) Gespräch stattfindet: Der Türhüter verfügt über Macht und Wissen, der Mann bleibt ohnmächtig.",
};

// Aufgabe 3: Die Handlung über das Gespräch erschließen (Muster der Mentorin) – [Der Mann … (vorgegeben), Der Türhüter … (Lösung), Anker für die Zeile]
const GESPRAECH = [
  ["bittet um Eintritt in das Gesetz.", "sagt, dass er ihm „jetzt den Eintritt nicht gewähren könne“.", "Aber der Türhüter sagt, dass"],
  ["fragt, ob er später eintreten darf.", "„Es ist möglich, jetzt aber nicht.“", "Es ist möglich"],
  ["bückt sich, um durch das Tor zu sehen.", "lacht, lädt ihn ein und droht: „Merke aber: Ich bin mächtig.“", "Als der Türhüter das merkt"],
  ["wartet jahrelang, bittet und versucht zu bestechen.", "stellt „teilnahmslose Fragen, wie sie große Herren stellen“; „Ich nehme es nur an, damit du nicht glaubst, etwas versäumt zu haben.“", ["teilnahmslose Fragen", "Ich nehme es nur an"]],
  ["stellt vor seinem Tod seine einzige Frage.", "brüllt: „… dieser Eingang war nur für dich bestimmt. Ich gehe jetzt und schließe ihn.“", "brüllt er ihn an"],
];
// Zeilenangabe für einen oder mehrere Anker
const zeilen = (anker) => (Array.isArray(anker) ? anker : [anker]).map(zeile).join(", ").replace(/, Z\. /g, ", ");
// Zeilen zu dem, was der Mann sagt oder tut (vorgegeben, damit die Klasse von dort aus weitersucht)
const nr = (a) => parseInt(zeile(a).replace("Z. ", ""), 10);
const bereich = (von, bis) => (bis ? "Z. " + nr(von) + "–" + nr(bis) : zeile(von));
const MANN_Z = [
  bereich("Zu diesem Türhüter kommt ein Mann"),
  bereich("Der Mann überlegt und fragt"),
  bereich("bückt sich der Mann"),
  bereich("Der Türhüter gibt ihm einen Schemel", "um den Türhüter zu bestechen"),
  bereich("Vor seinem Tode sammeln", "Einlass verlangt hat"),
];
const AUFGABE3 = {
  titel: "Aufgabe 3: Das Gespräch Schritt für Schritt",
  zeit: "6 Min.", form: "allein, auf dem Arbeitsblatt",
  schritte: [
    "Lies, was der Mann sagt oder tut.",
    "Notiere, wie der Türhüter reagiert – mit Zitat und Zeile.",
    "Beantworte die beiden Fragen unter der Tabelle.",
  ],
  fragen: ["Wer bestimmt das Gespräch? Woran erkennst du das?", "Vergleiche: Wie wird die Rede der beiden wiedergegeben?"],
};
const GESPRAECH_FAZIT = [
  ["Wer bestimmt?", "Der Türhüter entscheidet, droht, verhört und schließt. Der Mann bittet, fragt, wartet – er nimmt die untere Rolle an."],
  ["Wie sprechen sie?", "Der Türhüter spricht in wörtlicher Rede. Der Mann wird bis kurz vor seinem Tod nur indirekt wiedergegeben – er hat keine eigene Stimme."],
  ["Bild dafür", "„… der Größenunterschied hat sich sehr zuungunsten des Mannes verändert“ (vgl. " + zeile("Größenunterschied hat sich sehr") + ")."],
];
const GESPRAECH_ERGEBNIS = "Unsere Hypothese stimmt: Das Gespräch ist komplementär (asymmetrisch) – deshalb scheitert die Kommunikation. Der Mann redet mit niemandem außer dem Türhüter und ist am Ende ausgesperrt. Kafka: „von allen absperren“.";

// Aufgabe 4: Symbole deuten – [Bildelement, Zitat für die Zeile, Sachebene (Lösung)]; das Tor macht die Lehrkraft vor
const SYMBOLE = [
  ["das offene Tor", "Da das Tor zum Gesetz offensteht", "Der Zugang scheint möglich, das Ziel ist sichtbar – und bleibt doch unerreichbar: Hoffnung und Verbot zugleich."],
  ["der Türhüter", "Merke aber: Ich bin mächtig", "Eine Autorität (Beamter, Behörde, Staat), die den Zugang verwaltet. Ihre Macht kommt aus einer unsichtbaren Hierarchie – Gründe nennt sie nie."],
  ["das Gesetz", "das Gesetz soll doch jedem", "Recht und Gerechtigkeit – offen gedeutet auch Wahrheit, Sinn des Lebens oder Gott: das, wonach jeder Mensch strebt."],
];
const AUFGABE4 = {
  titel: "Aufgabe 4: Symbole deuten",
  zeit: "4 Min.", form: "allein, auf dem Arbeitsblatt",
  schritte: [
    "Suche die Stelle im Text und notiere die Zeile.",
    "Deute: Wofür steht das Bildelement?",
    "Begründe mit einem Wort aus dem Text.",
  ],
};
// Ich mache vor: das offene Tor (lautes Denken in drei Schritten)
const VORMACHEN = [
  ["1. Was steht da?", "„Da das Tor zum Gesetz offensteht wie immer …“ – das Tor ist offen, immer schon."],
  ["2. Was fällt auf?", "Offen – und trotzdem darf er nicht hinein. Das passt nicht zusammen."],
  ["3. Wofür steht es?", "Der Zugang scheint möglich, das Ziel ist sichtbar – und bleibt doch unerreichbar: Hoffnung und Verbot zugleich."],
];

// Zielcheck (im Heft)
const ZIELCHECK = {
  auftrag: "Erkläre in 2–3 Sätzen, warum „Vor dem Gesetz“ kafkaesk ist. Beziehe dich auf ein Symbol mit Zeilenangabe.",
  geruest: "Der / Das … (vgl. Z. …) steht für … . Das ist kafkaesk, weil …",
  kriterien: ["Symbol mit Zeile", "Deutung auf der Sachebene", "ein Baustein oder Gefühl von „kafkaesk“"],
};
ZIELCHECK.muster = "Der Türhüter (vgl. " + zeile("Vor dem Gesetz steht ein Türhüter") + ") steht für eine staatliche Autorität, die den Zugang zum Recht verwaltet. Das ist kafkaesk, weil er keinen Grund nennt („jetzt aber nicht“, vgl. " + zeile("jetzt aber nicht") + ") und hinter ihm immer mächtigere Türhüter stehen (labyrinthische Bürokratie). Der Mann ist dieser Macht ohnmächtig ausgeliefert und wartet sein Leben lang.";

module.exports = {
  TITEL, AUTOR, TEXT, TEXT_HINWEIS, ZEILEN, zeile, ZITAT, PROCESS_SATZ, LERNZIELE, VERLAUF, VORWISSEN, KAFKAESK, SPRINT,
  NK_KRITERIUM, SZENARIEN, NK_AUFGABE, KONTEXT, LESEAUFTRAG, TATT, GESPRAECH, zeilen, MANN_Z, AUFGABE3, GESPRAECH_FAZIT, GESPRAECH_ERGEBNIS, SYMBOLE, AUFGABE4, VORMACHEN, ZIELCHECK,
};
