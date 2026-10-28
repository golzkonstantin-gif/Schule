// Mystery „Lysias“ – Kleisthenes. Lysias ist eine erfundene Figur, alle anderen Angaben folgen Herodot (Buch 5)
// und Aristoteles, Staat der Athener 20–22, sowie dem Lehrbuch („Die antike Welt“, S. 16 f.).
// typ: kontext | problem | wende | reform | folge | ablenker
module.exports = {
  frage: "Lysias, ein Bauer aus Paiania im Binnenland, darf seit Solon in der Volksversammlung abstimmen. Trotzdem hebt er jahrzehntelang immer dann die Hand, wenn es sein adliger Nachbar will. 507 v. Chr. stimmt er zum ersten Mal anders. Warum?",
  karten: [
    { id: "A", typ: "kontext", t: "Lysias ist Bauer in Paiania, einem Dorf im Binnenland Attikas." },
    { id: "B", typ: "kontext", t: "Seit Solon (594 v. Chr.) darf jeder Bürger über 20 in der Volksversammlung abstimmen – auch ein armer Bauer wie Lysias." },
    { id: "C", typ: "problem", t: "Lysias' Nachbar ist ein reicher Adliger. In schlechten Jahren hat er Lysias Saatgut geliehen. Seine Familie ist die mächtigste im ganzen Binnenland." },
    { id: "D", typ: "problem", t: "Bisher waren alle Athener in vier alte Phylen eingeteilt. Dort gaben die großen Adelsfamilien den Ton an: Wer zu einer Familie hielt, stimmte mit ihr." },
    { id: "E", typ: "problem", t: "In der Volksversammlung wird offen mit Handzeichen abgestimmt. Jeder sieht, wie sein Nachbar stimmt." },
    { id: "F", typ: "wende", t: "510 v. Chr. wird der Tyrann Hippias aus Athen vertrieben. Danach kämpfen zwei Adlige um die Führung: Isagoras und Kleisthenes." },
    { id: "G", typ: "wende", t: "Kleisthenes stammt aus der Adelsfamilie der Alkmeoniden. Um gegen Isagoras zu gewinnen, sucht er die Unterstützung des Volkes." },
    { id: "H", typ: "wende", t: "Isagoras holt Truppen aus Sparta zu Hilfe. Doch das Volk erhebt sich und vertreibt die Spartaner von der Akropolis. 508/07 v. Chr. setzt Kleisthenes seine Reform durch." },
    { id: "I", typ: "reform", t: "Kleisthenes teilt Attika in drei Gebiete: Stadt, Küste und Binnenland. Jedes Gebiet wird in zehn Bezirke geteilt, die Trittyen." },
    { id: "J", typ: "reform", t: "Kleisthenes bildet zehn neue Phylen. Jede Phyle bekommt – durch das Los – je eine Trittys aus der Stadt, eine von der Küste und eine aus dem Binnenland." },
    { id: "K", typ: "folge", t: "Die Anhänger von Lysias' Nachbarn wohnen in vielen Dörfern des Binnenlandes. Nach der Reform gehören diese Dörfer zu verschiedenen Phylen." },
    { id: "L", typ: "folge", t: "Lysias gehört jetzt zur Phyle Pandionis. Mit ihm sind darin Fischer von der Küste und Handwerker aus der Stadt, die seinen Nachbarn gar nicht kennen." },
    { id: "M", typ: "folge", t: "Jede Phyle schickt 50 Männer in den neuen Rat der 500. Er bereitet die Volksversammlung vor. Im Rat sitzen Bürger aus allen drei Gebieten zusammen." },
    { id: "N", typ: "folge", t: "Auch das Heer wird nach den zehn Phylen aufgestellt: Lysias kämpft jetzt Seite an Seite mit Männern aus der Stadt und von der Küste." },
    { id: "O", typ: "folge", t: "Bürger ist, wer in der Liste seines Demos (seiner Dorfgemeinde) steht. Man nennt sich nach dem Demos – „Lysias aus Paiania“ –, nicht mehr nach Vater und Familie." },
    { id: "P", typ: "folge", t: "Mit dem Ostrakismos (Scherbengericht) kann das Volk einen zu mächtigen Mann für zehn Jahre aus Athen verbannen." },
    { id: "Q", typ: "ablenker", t: "Lysias baut auf seinen Feldern Oliven und Gerste an." },
    { id: "R", typ: "ablenker", t: "Der Großvater des Kleisthenes war Tyrann der Stadt Sikyon." },
  ],
  // Erwartete Erklärungskette (Lösung)
  kette: [
    ["Ausgangslage", "A, B, C, D, E", "Lysias hat seit Solon das Stimmrecht, ist aber von seinem adligen Nachbarn abhängig und an dessen Familie gebunden – und jeder sieht, wie er stimmt."],
    ["Wende", "F, G, H", "Nach dem Sturz der Tyrannis braucht Kleisthenes das Volk; das Volk setzt sich gegen Isagoras und Sparta durch."],
    ["Reform", "I, J", "Attika wird neu gegliedert: 3 Gebiete × 10 Trittyen, daraus 10 Phylen, die Stadt, Küste und Binnenland mischen."],
    ["Folge", "K, L, M, N, O", "Die Gefolgschaft des Adligen wird auf mehrere Phylen verteilt. Lysias sitzt, berät und kämpft mit Fremden zusammen; er gehört zu seinem Demos statt zu einer Familie."],
    ["Ergebnis", "", "Lysias ist nicht mehr an seinen Adelsherrn gebunden – er kann sein Stimmrecht frei nutzen."],
    ["Zusatz", "P", "Gegen einen neuen Tyrannen hat das Volk jetzt den Ostrakismos."],
  ],
};
