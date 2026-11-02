// Mystery „Lysias“ – Kleisthenes. Lysias ist eine erfundene Figur, alle anderen Angaben folgen Herodot (Buch 5)
// und Aristoteles, Staat der Athener 20–22, sowie dem Lehrbuch („Die antike Welt“, S. 16 f.).
// 8 Textabschnitte mit Überschrift; Buchstaben = logische Reihenfolge, gedruckt wird gemischt nummeriert.
module.exports = {
  frage: "Lysias, ein Bauer aus Paiania im Binnenland, darf seit Solon in der Volksversammlung abstimmen. Trotzdem hebt er jahrzehntelang immer dann die Hand, wenn es sein adliger Nachbar will. 507 v. Chr. stimmt er zum ersten Mal anders. Warum?",
  karten: [
    { id: "A", h: "Lysias und sein Nachbar", t: "Lysias ist Bauer in Paiania, einem Dorf im Binnenland Attikas. Seit Solon (594 v. Chr.) darf jeder Bürger über 20 in der Volksversammlung abstimmen – auch ein armer Bauer wie er. Sein Nachbar ist ein reicher Adliger, dessen Familie die mächtigste im ganzen Binnenland ist. In schlechten Jahren hat er Lysias Saatgut geliehen." },
    { id: "B", h: "Die alte Ordnung", t: "___ die großen Adelsfamilien bestimmen in den vier alten Phylen. Diese Phylen sind nach Abstammung gebildet: Die Gefolgsleute einer Adelsfamilie gehören alle zu derselben Phyle wie ihr Herr. Aus jeder Phyle kommen 100 Männer in den Rat der 400 – dort sitzen die Leute einer Familie als geschlossener Block. In der Volksversammlung wird offen mit Handzeichen abgestimmt." },
    { id: "C", h: "Das Ende der Tyrannis", t: "___ wird 510 v. Chr. der Tyrann Hippias aus Athen vertrieben. Zwei Adlige kämpfen nun um die Führung: Isagoras und Kleisthenes aus der Familie der Alkmeoniden. Um gegen Isagoras zu gewinnen, sucht Kleisthenes die Unterstützung des Volkes." },
    { id: "D", h: "Das Volk greift ein", t: "___ holt Isagoras Truppen aus Sparta zu Hilfe, und Kleisthenes muss Athen verlassen. Doch das Volk erhebt sich, belagert die Spartaner auf der Akropolis und zwingt sie zum Abzug. Kleisthenes kehrt zurück und setzt 508/07 v. Chr. seine Reform durch." },
    { id: "E", h: "Attika wird neu eingeteilt", t: "___ teilt Kleisthenes Attika in drei Gebiete: Stadt, Küste und Binnenland. Jedes Gebiet besteht aus zehn Bezirken (Trittyen), jeder Bezirk aus mehreren Dörfern (Demen). Niemand muss umziehen – jeder bleibt in seinem Dorf. Bürger ist jetzt, wer in der Liste seines Dorfes steht, nicht wer zu einer Familie gehört. Man nennt sich nach dem Dorf: „Lysias aus Paiania“ statt „Lysias, Sohn des …“." },
    { id: "F", h: "Zehn neue Phylen", t: "___ bildet Kleisthenes zehn neue Phylen. Jede besteht aus drei Bezirken: einem aus der Stadt, einem von der Küste und einem aus dem Binnenland – zugeteilt durch das Los. Die Gefolgsleute von Lysias' Nachbarn wohnen in sechs Dörfern des Binnenlandes. Diese Dörfer liegen in drei verschiedenen Bezirken. Deshalb gehören die Gefolgsleute jetzt zu drei verschiedenen Phylen." },
    { id: "G", h: "Lysias' neue Phyle", t: "___ gehört Lysias – genau wie sein Nachbar – zur Phyle Pandionis. Doch nur ein Drittel dieser Phyle kommt aus dem Binnenland. Zwei Drittel sind Fischer von der Küste und Handwerker aus der Stadt, die von seinem Nachbarn nicht abhängig sind. Mit ihnen bestimmt Lysias die 50 Männer seiner Phyle für den neuen Rat der 500, und mit ihnen steht er im Heer." },
    { id: "H", h: "Ein Mittel gegen Tyrannen", t: "___ kann das Volk mit dem Ostrakismos (Scherbengericht) einen Mann, der zu mächtig wird, für zehn Jahre aus Athen verbannen. Jeder Bürger ritzt dazu einen Namen auf eine Tonscherbe. Vermögen und Bürgerrecht behält der Verbannte." },
  ],
  // Verbindungswörter für die Wortleiste (gemischt gedruckt) und Lösung als Fließtext
  verbinder: ["dadurch", "aber", "außerdem", "daraufhin", "doch dann", "daraus", "danach"],
  loesung: [["", "A"], ["aber", "B"], ["Doch dann", "C"], ["Daraufhin", "D"], ["Danach", "E"], ["Daraus", "F"], ["Dadurch", "G"], ["", "Ergebnis"], ["Außerdem", "H"]],
  // Vorher-Nachher-Vergleich: [Merkmal, vorher, nachher, Folie: Merkmal kurz, vorher kurz, nachher kurz]
  vergleich: [
    ["Wonach werden die Phylen gebildet?", "nach Familie und Abstammung (4 Phylen)", "nach dem Wohnort: je ein Bezirk aus Stadt, Küste und Binnenland (10 Phylen)", "Phylen", "nach Familie", "nach Wohnort"],
    ["Wer ist Bürger?", "wer zu einer Familie gehört", "wer in der Liste seines Dorfes (Demos) steht", "Bürger", "über die Familie", "über die Dorfliste"],
    ["Wie heißt Lysias?", "„Lysias, Sohn des …“", "„Lysias aus Paiania“", "Name", "„Sohn des …“", "„aus Paiania“"],
    ["Wo sind die Gefolgsleute des Adligen?", "alle zusammen in seiner Phyle", "verteilt auf drei verschiedene Phylen", "Anhänger", "in einer Phyle", "in drei Phylen"],
    ["Mit wem ist Lysias in seiner Phyle?", "mit den Leuten seines Adelsherrn", "zu zwei Dritteln mit Fischern und Handwerkern, die nicht vom Adligen abhängen", "Seine Phyle", "Leute des Adligen", "⅔ aus Stadt, Küste"],
    ["Wer sitzt im Rat?", "Rat der 400: die Leute einer Familie als Block", "Rat der 500: 50 je Phyle, aus allen Gebieten gemischt", "Rat", "400, in Blöcken", "500, gemischt"],
  ],
  gleich: "Wohnort, Besitz, Vermögensklassen, Stimmrecht und offene Abstimmung bleiben gleich.",
  // Erwartete Erklärungskette (Lösung)
  kette: [
    ["Ausgangslage", "A, B", "Lysias hat seit Solon das Stimmrecht, ist aber von seinem adligen Nachbarn abhängig. In der alten Phyle und im Rat der 400 sind die Gefolgsleute der Familie ein geschlossener Block – und jeder sieht, wie er stimmt."],
    ["Wende", "C, D", "Nach dem Sturz der Tyrannis braucht Kleisthenes das Volk; das Volk setzt sich gegen Isagoras und Sparta durch."],
    ["Reform", "E, F", "Attika wird neu gegliedert: 3 Gebiete × 10 Bezirke (Trittyen); jede der 10 neuen Phylen bekommt einen Bezirk aus Stadt, Küste und Binnenland. Bürger ist man über sein Dorf (Demos), nicht über seine Familie."],
    ["Folge", "F, G", "Niemand zieht um, aber die Dörfer der Gefolgsleute gehören jetzt zu drei verschiedenen Phylen. In Lysias' eigener Phyle kommen zwei Drittel aus Stadt und Küste und hängen nicht vom Adligen ab."],
    ["Ergebnis", "", "Lysias ist nicht mehr an seinen Adelsherrn gebunden – er kann sein Stimmrecht frei nutzen."],
    ["Zusatz", "H", "Gegen einen neuen Tyrannen hat das Volk jetzt den Ostrakismos."],
  ],
};
