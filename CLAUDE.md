# Unterrichtsplanung – Hinweise für Claude

Dieses Repo enthält Unterrichtsmaterial (Deutsch). Die Lehrkraft plant hier gemeinsam mit Claude Stunden, meldet zurück, wie sie gelaufen sind, und lässt Hospitationsfeedback einfließen. Ziel: Die Planung wird von Stunde zu Stunde besser.

## Vor jeder neuen Planung

1. Das Reflexionslog der Klasse lesen (z. B. `Deutsch_7A/Reflexion.md`): Lernstand, offene Baustellen, Rückmeldungen zu vergangenen Stunden, Hospitationsfeedback.
2. Die Planung ausdrücklich darauf beziehen („Letztes Mal hing die Klasse bei …, deshalb …“).
3. Wenn die Lehrkraft Rückmeldung zu einer Stunde oder Hospitationsfeedback gibt: im Reflexionslog eintragen (Vorlage steht dort) und mitcommitten. Dauerhafte neue Vorlieben zusätzlich hier in `CLAUDE.md` ergänzen.

## Vorlieben der Lehrkraft

- **Erst fragen, was gewünscht ist:** Wenn die Lehrkraft „erst mal Ideen sammeln“ sagt, nur Ideen im Chat – kein Material erstellen.
- **Kleinschrittig und Vorwissen nicht überschätzen.** Lieber einen Schritt zurückgehen und Grundlagen sichern (die Bastelstunde zu Zeitformen ist genau daran gescheitert).
- **Inhalt darf elementar sein, der Ton muss zur Altersstufe passen (7. Klasse, keine Grundschulanmutung).** Also:
  - Fachbegriffe konsequent verwenden (finit/infinit, Hilfsverb/Vollverb) statt kindlicher Metaphern („Verbfamilie“, „Familienname“),
  - Begründungen einfordern („Begründe, warum …“), Fehleranalyse statt Rätsel (z. B. Chatverlauf mit echten Fehlern),
  - Beispielsätze aus der Lebenswelt der Klasse (Sport, Serien, Handy, Klassenfahrt), keine Kinderbeispiele,
  - keine Finger- oder Handzeichenspiele; stattdessen z. B. Speed-Duell zu zweit mit Punkten,
  - Relevanz zeigen (wofür man es braucht, Bezug zum Englischen).
- **Wenige Arbeitsblätter.** Übungen laufen an der Tafel, die Klasse schreibt ins Heft ab und löst dort.
- **Transparenz durch die Lehrkraft, nicht durch Hilfsmittel.** Keine Modus- oder Symbolkarten (ausdrücklich abgelehnt). Stattdessen:
  - feste, wörtliche Ansagen bei jedem Wechsel („Stifte liegen. Ihr hört nur zu.“ / „Jetzt schreibt ihr ab: …“ / „Jetzt arbeitet ihr allein. Die Aufgabe steht rechts.“ / „Wer drankommt, schreibt an die Tafel. Alle prüfen mit.“ / „Vergleicht und verbessert mit Grün.“),
  - Übungen immer im gleichen Format: Nummer + Titel, nummerierte Arbeitsschritte mit Verb am Anfang, Zeit, Arbeitsform, „Fertig? Stift hinlegen“.
- **Bewährt: „Kreide-Kette“.** Die Schüler kommen nacheinander an die Tafel, lösen eine Aufgabe und geben die Kreide an den Nächsten weiter. Das sorgt für eine konzentrierte Atmosphäre und soll bei Tafelübungen standardmäßig eingeplant werden.
- **Zu Beginn jeder Stunde das Vorwissen der letzten Stunde reaktivieren**, mit kurzen Übungen an der Tafel.
- **Gleicher Ablauf je Phase:** gemeinsam entdecken → Merksatz abschreiben → allein üben → an der Tafel vergleichen → mit Grün verbessern.
- **Merksätze** immer mit „Wofür brauche ich das?“, dazu ein Erkennungstrick/Test.
- **Tafelskript** für die Lehrkraft: pro Schritt, was die Klasse tut, die wörtliche Ansage, der Tafelanschrieb, die Lösung. Dazu eine Übersicht, wie das Heft am Ende aussieht.
- **Parallel zur Tafel eine PowerPoint mit Merksätzen und Aufgaben** (gleicher Wortlaut wie im Tafelskript, Lösungen nur in den Notizen). Die **Merksätze kommen auf die Folie**, nicht an die Tafel. Je Merksatz eine Folie: Kernsatz im roten Rahmen, Beispiel in Farbe, „Wofür?“ und „Test“, jeweils direkt vor der passenden Übung. An der Tafel entstehen nur Tafelbilder und Tabellen, die gemeinsam entwickelt werden.

## Gestaltung

- Formate: Word (`.docx`) für Skripte/Arbeitsblätter, PowerPoint (`.pptx`) für Folien, jeweils zusätzlich als PDF.
- Farben: Navy `1E2761` (Grundfarbe), Hellblau `F4F7FD`/`EAF0FA` (Flächen), Grau `5B6B8C` (Nebentext).
- Farbcode Verbformen (gilt für Tafel, Heft, Folien): Hilfsverb haben/sein/werden **rot** `D9534F`, Modalverb **grün** `2E9E6B`, Partizip II **gelb** `E8A33D`, Infinitiv **blau** `3F7CC4`, Verbesserungen im Heft grün.
- Schrift: Überschriften Cambria, Text Calibri. Folien 16:9 im Stil der bisherigen Präsentationen.
- Die Dateien werden per Skript erzeugt; die Skripte liegen jeweils im Unterordner `quellcode/` (Node: `docx`, `pptxgenjs`). Änderungen am Material im Skript machen und neu erzeugen.

## Datenschutz

Keine Klarnamen von Schülerinnen, Schülern oder Kolleginnen und Kollegen ins Repo schreiben. Stattdessen anonymisieren (z. B. „S1“, „Fachleitung“, „Mentor“).

## Struktur

- `Deutsch_7A/` – Material der Klasse 7A, ein Unterordner pro Stunde/Einheit
- `Deutsch_7A/Reflexion.md` – Lernstand, Stundenlog, Hospitationsfeedback, offene Baustellen
