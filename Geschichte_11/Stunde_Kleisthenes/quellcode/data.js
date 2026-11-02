// Mystery „Thrasymachos“ – Kleisthenes. Thrasymachos und Lysias sind erfundene Figuren, alle anderen Angaben folgen Herodot (Buch 1 und 5)
// und Aristoteles, Staat der Athener 13 und 20–22, sowie dem Lehrbuch („Die antike Welt“, S. 16 f.).
// Perspektive des Adligen (Wunsch der Lehrkraft, 2.11.): Wie brach Kleisthenes die Macht der Adelsfamilien – ohne Land oder Rechte zu nehmen?
// 8 Textabschnitte mit Überschrift; Buchstaben = logische Reihenfolge, gedruckt wird gemischt nummeriert.
module.exports = {
  titel: "Warum verliert Thrasymachos seine Macht?",
  frage: "Thrasymachos aus Paiania ist der mächtigste Adlige im Binnenland Attikas: Hunderte Bauern sind von ihm abhängig. Nach der Reform des Kleisthenes (508/07 v. Chr.) verliert er kein Land, kein Geld und kein Recht – und hat trotzdem kaum noch politische Macht. Warum?",
  karten: [
    { id: "A", h: "Thrasymachos und seine Leute", t: "Thrasymachos stammt aus einer alten Adelsfamilie und besitzt viel Land um Paiania, ein Dorf im Binnenland. Hunderte Bauern der Gegend sind von ihm abhängig: Sie pachten sein Land oder haben bei ihm Schulden – so wie der Bauer Lysias, dem er in schlechten Jahren Saatgut geliehen hat. Wenn Thrasymachos ruft, folgen ihm seine Leute geschlossen." },
    { id: "B", h: "Macht in Athen", t: "___ ist Thrasymachos auch in Athen mächtig. Dort kämpfen Adlige mit den Leuten ihrer Gegend um die Macht – so kam auch Peisistratos mit seinen Anhängern aus dem Bergland an die Macht. Im Rat der 400 geben die Adelsfamilien den Ton an. Und wer Bürger ist, entscheiden die Familienverbände, an deren Spitze Adlige wie Thrasymachos stehen." },
    { id: "C", h: "Das Ende der Tyrannis", t: "___ wird 510 v. Chr. der Tyrann Hippias aus Athen vertrieben. Zwei Adlige kämpfen nun um die Führung: Isagoras und Kleisthenes aus der Familie der Alkmeoniden. Um gegen Isagoras zu gewinnen, sucht Kleisthenes die Unterstützung des Volkes." },
    { id: "D", h: "Das Volk greift ein", t: "___ holt Isagoras Truppen aus Sparta zu Hilfe, und Kleisthenes muss Athen verlassen. Doch das Volk erhebt sich, belagert die Spartaner auf der Akropolis und zwingt sie zum Abzug. Kleisthenes kehrt zurück und setzt 508/07 v. Chr. seine Reform durch." },
    { id: "E", h: "Attika wird neu eingeteilt", t: "___ teilt Kleisthenes Attika in drei Gebiete: Stadt, Küste und Binnenland. Jedes Gebiet besteht aus zehn Bezirken (Trittyen), jeder Bezirk aus mehreren Dörfern (Demen). Bürger ist jetzt, wer in der Liste seines Dorfes steht – darüber entscheidet die Dorfversammlung, nicht mehr ein Familienverband. Man nennt sich nach dem Dorf: „Thrasymachos aus Paiania“." },
    { id: "F", h: "Zehn neue Phylen", t: "___ bildet Kleisthenes zehn neue Phylen. Jede besteht aus drei Bezirken: einem aus der Stadt, einem von der Küste und einem aus dem Binnenland – zugeteilt durch das Los. Die Leute des Thrasymachos wohnen in sechs Dörfern, die in drei verschiedenen Bezirken liegen. Niemand zieht um – aber ihre Dörfer gehören jetzt zu drei verschiedenen Phylen." },
    { id: "G", h: "Plötzlich in der Minderheit", t: "___ in seiner neuen Phyle Pandionis ist Thrasymachos nur noch einer unter vielen: Zwei Drittel kommen aus Stadt und Küste, Fischer und Handwerker, die ihm nichts schulden. Sie bestimmen mit, wer aus der Phyle in den neuen Rat der 500 kommt, der jede Abstimmung vorbereitet. Im Heer kämpfen seine Leute in gemischten Abteilungen. Und Lysias muss nicht mehr fürchten, sein Bürgerrecht zu verlieren." },
    { id: "H", h: "Ein Mittel gegen zu mächtige Männer", t: "___ kann das Volk mit dem Ostrakismos (Scherbengericht) einen Mann, der zu mächtig wird, für zehn Jahre aus Athen verbannen. Jeder Bürger ritzt dazu einen Namen auf eine Tonscherbe. Vermögen und Bürgerrecht behält der Verbannte." },
  ],
  // Verbindungswörter für die Wortleiste (gemischt gedruckt) und Lösung als Fließtext
  verbinder: ["dadurch", "aber", "außerdem", "daraufhin", "doch dann", "daraus", "danach"],
  loesung: [["", "A"], ["Dadurch", "B"], ["Doch dann", "C"], ["Daraufhin", "D"], ["Danach", "E"], ["Daraus", "F"], ["Aber", "G"], ["", "Ergebnis"], ["Außerdem", "H"]],
  ergebnis: "Deshalb behält Thrasymachos zwar Land und Rechte, aber seine Gegend ist kein geschlossener Block mehr: In jeder Phyle, im Rat und im Heer sind seine Leute nur noch ein Drittel unter vielen. Kleisthenes hat dem Adel nichts weggenommen – er hat seine Machtbasis aufgeteilt.",
  // Vorher-Nachher-Vergleich: [Merkmal, vorher, nachher, Folie: Merkmal kurz, vorher kurz, nachher kurz]
  vergleich: [
    ["Worauf beruht seine Macht?", "auf seiner Gegend: die abhängigen Bauern folgen ihm als Block", "seine Gegend ist auf drei Phylen verteilt – in jeder nur ein Drittel", "Machtbasis", "seine Gegend", "auf 3 Phylen verteilt"],
    ["Wie werden die Phylen gebildet?", "nach Familie und Abstammung (4 Phylen)", "nach dem Wohnort: je ein Bezirk aus Stadt, Küste und Binnenland (10 Phylen)", "Phylen", "nach Familie", "gemischt nach Wohnort"],
    ["Wer bestimmt im Rat?", "Rat der 400: die Adelsfamilien geben den Ton an", "Rat der 500: 50 je Phyle, davon zwei Drittel aus Stadt und Küste", "Rat", "Adelsfamilien", "⅔ aus Stadt, Küste"],
    ["Wer entscheidet über das Bürgerrecht?", "die Familienverbände, an der Spitze der Adel", "die Dorfversammlung (Demos)", "Bürgerrecht", "Familienverband", "Dorfversammlung"],
    ["Kann er mit seiner Gegend die Macht ergreifen?", "ja – so kam Peisistratos an die Macht", "kaum: keine Phyle besteht mehr nur aus einer Gegend", "Umsturz", "möglich", "kaum noch"],
    ["Wie heißt er?", "„Thrasymachos, Sohn des …“ – die Familie zählt", "„Thrasymachos aus Paiania“ – das Dorf zählt", "Name", "„Sohn des …“", "„aus Paiania“"],
  ],
  gleich: "Land, Vermögen, Vermögensklasse und Rechte des Thrasymachos bleiben gleich – und niemand zieht um.",
  // Erwartete Erklärungskette (Lösung)
  kette: [
    ["Ausgangslage", "A, B", "Die Macht des Thrasymachos beruht auf seiner Gegend: Abhängige Bauern folgen ihm als Block. Damit zählt er in Athen; Adelsfamilien geben im Rat den Ton an und entscheiden über das Bürgerrecht."],
    ["Wende", "C, D", "Nach dem Sturz der Tyrannis braucht Kleisthenes das Volk; das Volk setzt sich gegen Isagoras und Sparta durch."],
    ["Reform", "E, F", "Bürgerrecht über die Dorfliste. 3 Gebiete × 10 Bezirke (Trittyen); jede der 10 neuen Phylen bekommt einen Bezirk aus Stadt, Küste und Binnenland. Die Dörfer seiner Leute kommen in drei verschiedene Phylen."],
    ["Folge", "G", "In seiner Phyle ist seine Gegend nur ein Drittel: Rat der 500, Heer und Bürgerrecht kann er nicht mehr kontrollieren."],
    ["Ergebnis", "", "Er behält Land und Rechte, aber seine Gegend ist kein Machtblock mehr – die Macht der Adelsfamilien ist gebrochen."],
    ["Zusatz", "H", "Wird ein Adliger trotzdem zu mächtig, kann das Volk ihn mit dem Ostrakismos verbannen."],
  ],
};
