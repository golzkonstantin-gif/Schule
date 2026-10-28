// Mystery „Lysias“ – Kleisthenes. Lysias ist eine erfundene Figur, alle anderen Angaben folgen Herodot (Buch 5)
// und Aristoteles, Staat der Athener 20–22, sowie dem Lehrbuch („Die antike Welt“, S. 16 f.).
// 8 Textabschnitte mit Überschrift; Buchstaben = logische Reihenfolge, gedruckt wird gemischt nummeriert.
module.exports = {
  frage: "Lysias, ein Bauer aus Paiania im Binnenland, darf seit Solon in der Volksversammlung abstimmen. Trotzdem hebt er jahrzehntelang immer dann die Hand, wenn es sein adliger Nachbar will. 507 v. Chr. stimmt er zum ersten Mal anders. Warum?",
  karten: [
    { id: "A", h: "Lysias und sein Nachbar", t: "Lysias ist Bauer in Paiania, einem Dorf im Binnenland Attikas. Seit Solon (594 v. Chr.) darf jeder Bürger über 20 in der Volksversammlung abstimmen – auch ein armer Bauer wie er. Sein Nachbar ist ein reicher Adliger, dessen Familie die mächtigste im ganzen Binnenland ist. In schlechten Jahren hat er Lysias Saatgut geliehen." },
    { id: "B", h: "Die alte Ordnung", t: "Bisher waren alle Athener in vier alte Phylen eingeteilt, die nach Familien und Abstammung gebildet waren. Dort gaben die großen Adelsfamilien den Ton an: Wer zu einer Familie hielt, stimmte mit ihr. In der Volksversammlung wird offen mit Handzeichen abgestimmt – jeder sieht, wie sein Nachbar stimmt." },
    { id: "C", h: "Das Ende der Tyrannis", t: "510 v. Chr. wird der Tyrann Hippias aus Athen vertrieben. Danach kämpfen zwei Adlige um die Führung: Isagoras und Kleisthenes aus der Familie der Alkmeoniden. Um gegen Isagoras zu gewinnen, sucht Kleisthenes die Unterstützung des Volkes." },
    { id: "D", h: "Das Volk greift ein", t: "Isagoras holt Truppen aus Sparta zu Hilfe, Kleisthenes muss Athen verlassen. Doch das Volk erhebt sich, belagert die Spartaner auf der Akropolis und zwingt sie zum Abzug. Kleisthenes kehrt zurück und setzt 508/07 v. Chr. seine Reform durch." },
    { id: "E", h: "Attika wird neu eingeteilt", t: "Kleisthenes teilt Attika in drei Gebiete: Stadt, Küste und Binnenland. Jedes Gebiet besteht aus zehn Bezirken, den Trittyen; jede Trittys umfasst mehrere Demen (Dorfgemeinden). Bürger ist, wer in der Liste seines Demos steht. Man nennt sich nach dem Demos – „Lysias aus Paiania“ –, nicht mehr nach Vater und Familie." },
    { id: "F", h: "Zehn neue Phylen", t: "Aus den Trittyen bildet Kleisthenes zehn neue Phylen. Jede Phyle bekommt durch das Los je eine Trittys aus der Stadt, eine von der Küste und eine aus dem Binnenland. Die Anhänger von Lysias' Nachbarn wohnen in vielen Dörfern des Binnenlandes – nach der Reform gehören diese Dörfer zu verschiedenen Phylen." },
    { id: "G", h: "Lysias' neue Phyle", t: "Lysias gehört jetzt zur Phyle Pandionis, zusammen mit Fischern von der Küste und Handwerkern aus der Stadt, die seinen Nachbarn gar nicht kennen. Jede Phyle schickt 50 Männer in den neuen Rat der 500, der die Volksversammlung vorbereitet. Auch das Heer wird nach Phylen aufgestellt: Lysias kämpft Seite an Seite mit Männern aus allen Gebieten." },
    { id: "H", h: "Ein Mittel gegen Tyrannen", t: "Mit dem Ostrakismos (Scherbengericht) kann das Volk einen Mann, der zu mächtig wird, für zehn Jahre aus Athen verbannen. Jeder Bürger ritzt dazu einen Namen auf eine Tonscherbe. Vermögen und Bürgerrecht behält der Verbannte." },
  ],
  // Erwartete Erklärungskette (Lösung)
  kette: [
    ["Ausgangslage", "A, B", "Lysias hat seit Solon das Stimmrecht, ist aber von seinem adligen Nachbarn abhängig und an dessen Familie gebunden – und jeder sieht, wie er stimmt."],
    ["Wende", "C, D", "Nach dem Sturz der Tyrannis braucht Kleisthenes das Volk; das Volk setzt sich gegen Isagoras und Sparta durch."],
    ["Reform", "E, F", "Attika wird neu gegliedert: 3 Gebiete × 10 Trittyen, daraus 10 Phylen, die Stadt, Küste und Binnenland mischen. Bürger ist man über seinen Demos, nicht über seine Familie."],
    ["Folge", "F, G", "Die Gefolgschaft des Adligen wird auf mehrere Phylen verteilt. Lysias sitzt, berät und kämpft mit Fremden zusammen."],
    ["Ergebnis", "", "Lysias ist nicht mehr an seinen Adelsherrn gebunden – er kann sein Stimmrecht frei nutzen."],
    ["Zusatz", "H", "Gegen einen neuen Tyrannen hat das Volk jetzt den Ostrakismos."],
  ],
};
