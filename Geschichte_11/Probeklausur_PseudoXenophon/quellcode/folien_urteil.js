// Folien zur Urteilsstunde (Einzelstunde 45 Min.) mit dem Blatt „Argumente und Beispielurteil“. AVIVA, alle Texte für die Klasse ≥ 27 pt, Foliennummern automatisch.
// Anlass (17.11.): Vorwissen zu Kleisthenes fehlte, Tafelbild zu komplex → Reform in einem einfachen Bild, sortierte Argumente, angefangenes Muster.
const pptxgen = require("pptxgenjs");
const OUT = process.argv[2];
const C = { dark: "2B2318", terra: "AD4426", sand: "F4ECDE", gold: "D9A441", muted: "5A5044", soft: "998C77", white: "FFFFFF", green: "3C783C", blue: "3E5C76", grey: "B5AC9E" };
const pres = new pptxgen(); pres.layout = "LAYOUT_16x9"; pres.title = "Urteil: Reformen des Kleisthenes";
const H = "Cambria", B = "Calibri", MIN = 27, TITLE = 36;
const PH = ["A", "V", "I", "V", "A"];
const txt = (s, text, o) => s.addText(text, { isTextBox: true, fontFace: B, color: C.dark, margin: 0, valign: "top", fontSize: MIN, ...o });
const num = s => { s.slideNumber = { x: 9.3, y: 5.3, w: 0.4, h: 0.25, fontFace: B, fontSize: 12, color: C.soft, align: "right" }; };
function aviva(s, k) {
  PH.forEach((l, i) => s.addText(l, { isTextBox: true, shape: pres.shapes.ROUNDED_RECTANGLE, rectRadius: 0.06, x: 7.15 + i * 0.5, y: 0.12, w: 0.44, h: 0.44,
    fill: { color: i === k ? C.terra : C.sand }, color: i === k ? C.white : C.soft, fontFace: B, fontSize: MIN, bold: true, align: "center", valign: "middle", margin: 0 }));
}
const tag = (s, t, fill = C.dark, color = C.white) => s.addText(t, { isTextBox: true, shape: pres.shapes.ROUNDED_RECTANGLE, rectRadius: 0.08, x: 7.15, y: 0.66, w: 2.4, h: 0.5,
  fill: { color: fill }, color, fontFace: B, fontSize: MIN, bold: true, align: "center", valign: "middle", margin: 0 });
const title = (s, t, k) => { aviva(s, k); txt(s, t, { x: 0.45, y: 0.15, w: 6.6, h: 1.05, fontSize: TITLE, bold: true, fontFace: H, valign: "middle" }); };
function panel(s, t, x, y, w, h, o = {}) {
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, rectRadius: 0.06, fill: { color: o.fill || C.white }, line: { color: o.line || C.terra, width: o.lw || 2 } });
  txt(s, t, { x: x + 0.2, y, w: w - 0.4, h, color: o.color || C.dark, bold: o.bold, valign: "middle", align: o.align, fontSize: o.size || MIN });
}
function steps(s, items, y, gap) {
  items.forEach((t, i) => {
    s.addText(String(i + 1), { isTextBox: true, shape: pres.shapes.OVAL, x: 0.45, y: y + i * gap, w: 0.5, h: 0.5, fill: { color: C.terra }, color: C.white, fontFace: H, fontSize: MIN, bold: true, align: "center", valign: "middle", margin: 0 });
    txt(s, t, { x: 1.1, y: y + i * gap - 0.02, w: 8.5, h: 0.55, valign: "middle" });
  });
}
const slide = (bg = C.white) => { const s = pres.addSlide(); s.background = { color: bg }; return s; };
// Satzanfänge fett/terrakotta: Text mit **…** in Runs zerlegen
const rich = (t, o = {}) => String(t).split(/(\*\*[^*]+\*\*)/).filter(Boolean).map(x => x.startsWith("**")
  ? { text: x.slice(2, -2), options: { bold: true, color: C.terra } } : { text: x, options: { color: o.color || C.dark } });

// 1 Titel
{ const s = slide(C.dark);
  txt(s, "GESCHICHTE · KLASSE 11", { x: 0.7, y: 0.9, w: 8.6, h: 0.5, bold: true, color: C.gold, charSpacing: 2 });
  txt(s, "Haben die Reformen des Kleisthenes die Macht des Adels beendet?", { x: 0.7, y: 1.45, w: 8.6, h: 1.9, fontSize: 40, bold: true, fontFace: H, color: C.white });
  txt(s, "Wir schreiben Sach- und Werturteil", { x: 0.7, y: 3.75, w: 8.6, h: 0.6, color: "E8DFD0" });
  s.addNotes("Material: Blatt „Argumente_und_Beispielurteil“ (A4 doppelseitig, 1 pro Schüler). Zum Vergleichen am Ende: Musterurteil in Hilfeblatt_Urteil_Loesung (Rückseite).\n\nWichtig nach der letzten Stunde: Die Reformen werden NICHT tief wiederholt. Folie 5 erklärt nur den Kern in einem Bild – den Rest liefert die Vorderseite des Blatts.");
}

// 2 Ziel und Ablauf (Ankommen)
{ const s = slide(); title(s, "Ziel und Ablauf", 0); tag(s, "Ankommen");
  panel(s, "Ich kann ein Sach- und ein Werturteil zur Frage schreiben.", 0.45, 1.3, 9.1, 0.95, { fill: C.white, line: C.terra, bold: true, color: C.terra });
  steps(s, ["Kurz klären: Worauf beruhte die Macht des Adels?", "Argumente lesen und ankreuzen", "Muster ansehen", "Sachurteil weiterschreiben", "Werturteil schreiben *", "Fazits vergleichen"], 2.4, 0.52);
  num(s); s.addNotes("0–2 Min. Ziel vorlesen, Ablauf zeigen (Seitentafel). * = falls die Zeit nicht reicht, Werturteil als Hausaufgabe. „Fazits vergleichen“ wird nie gestrichen.\n\nMindestziel: Sachurteil mit Fazit, das die Frage beantwortet.");
}

// 3 Vorwissen
{ const s = slide(); title(s, "Worauf beruhte die Macht des Adels?", 1); tag(s, "allein · 1 Min.");
  txt(s, "Vor Kleisthenes hatte der Adel mehr Macht als alle anderen.", { x: 0.45, y: 1.45, w: 9.1, h: 0.6 });
  steps(s, ["Überlege: Woher kam diese Macht?", "Notiere zwei Stichworte im Heft."], 2.35, 0.75);
  panel(s, "Tipp: Denk an Peisistratos – wie kam er an die Macht?", 0.45, 4.2, 9.1, 0.75, { line: C.soft });
  num(s); s.addNotes("2–5 Min. Nur Wissen, das die Klasse hat (Peisistratos-Stunde, Solon). Nach 1 Minute Stichprobe bei 3–4 Schülern unterschiedlichen Niveaus, neutral reagieren. Dann Folie 4.\n\nErwartet: Bauern folgen ihrem Adligen (Peisistratos hatte seine Gegend hinter sich); nur Reiche in die höchsten Ämter (Solon); Reichtum, Land.");
}

// 4 Lösung Vorwissen
{ const s = slide(); title(s, "Die Macht des Adels", 1); tag(s, "Lösung", C.gold, C.dark);
  [["Gefolgschaft", "Die Bauern einer Gegend folgen ihrem Adligen geschlossen."], ["Ämter", "Nur Reiche bekommen die höchsten Ämter."], ["Reichtum", "Der Adel besitzt das meiste Land."]].forEach(([k, t], i) => {
    const y = 1.35 + i * 1.3;
    txt(s, k, { x: 0.45, y, w: 2.6, h: 1.1, bold: true, color: C.terra, fontFace: H, valign: "middle" });
    panel(s, t, 3.1, y, 6.45, 1.1, { line: C.terra });
  });
  num(s); s.addNotes("Gelenk: „Genau an diesen drei Punkten prüfen wir gleich: Was davon hat Kleisthenes geändert – und was nicht?“");
}

// 5 Informieren: die Reform in einem Bild
{ const s = slide(); title(s, "Was machte Kleisthenes?", 2); tag(s, "zuhören");
  const dot = (x, y, c) => s.addShape(pres.shapes.OVAL, { x, y, w: 0.32, h: 0.32, fill: { color: c }, line: { color: c } });
  // Vorher: eine Gruppe, alle Bauern des Adligen
  txt(s, "Vorher", { x: 0.45, y: 1.25, w: 3.4, h: 0.5, bold: true, fontFace: H });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.45, y: 1.8, w: 3.4, h: 2.1, rectRadius: 0.08, fill: { color: C.white }, line: { color: C.dark, width: 2 } });
  s.addText("Adliger", { isTextBox: true, shape: pres.shapes.ROUNDED_RECTANGLE, rectRadius: 0.06, x: 1.05, y: 1.95, w: 2.2, h: 0.55, fill: { color: C.dark }, color: C.white, fontFace: B, fontSize: MIN, bold: true, align: "center", valign: "middle", margin: 0 });
  for (let r = 0; r < 2; r++) for (let c = 0; c < 6; c++) dot(0.75 + c * 0.5, 2.75 + r * 0.5, C.green);
  txt(s, "eine Gruppe = ein Block", { x: 0.45, y: 4.0, w: 3.4, h: 0.5, bold: true, color: C.terra, align: "center" });
  // Pfeil
  s.addShape(pres.shapes.RIGHT_ARROW, { x: 3.95, y: 2.55, w: 0.6, h: 0.5, fill: { color: C.terra }, line: { color: C.terra } });
  // Nachher: drei gemischte Gruppen
  txt(s, "Nachher: neu gemischt", { x: 4.65, y: 1.25, w: 4.9, h: 0.5, bold: true, fontFace: H });
  for (let g = 0; g < 3; g++) {
    const x = 4.65 + g * 1.65;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y: 1.8, w: 1.55, h: 2.1, rectRadius: 0.08, fill: { color: C.white }, line: { color: C.dark, width: 2 } });
    const cols = [C.green, C.green, C.blue, C.blue, C.gold, C.gold, C.grey, C.grey, C.blue, C.gold, C.grey, C.blue];
    for (let i = 0; i < 12; i++) dot(x + 0.12 + (i % 3) * 0.47, 1.95 + Math.floor(i / 3) * 0.47, i < 4 ? C.green : cols[(i + g * 3) % cols.length] === C.green ? C.blue : cols[(i + g * 3) % cols.length]);
  }
  txt(s, "seine Bauern (grün) überall in der Minderheit", { x: 4.65, y: 4.0, w: 4.9, h: 0.9, bold: true, color: C.terra, align: "center" });
  num(s); s.addNotes("5–8 Min. Direkt erklären, nicht erfragen. Laut sprechen:\n„Vorher stimmten alle Bauern aus der Gegend eines Adligen in derselben Gruppe – sie folgten ihm geschlossen, wie bei Peisistratos. Kleisthenes teilte die Bürger in zehn neue Gruppen, die Phylen. Jede wird aus Stadt, Küste und Binnenland zusammengelost. Niemand zieht um. Aber die Bauern eines Adligen landen jetzt in verschiedenen Gruppen – und sind dort überall in der Minderheit.“\n\nWarnung vor dem Fehler: „Der Adlige verliert kein Land und keine Rechte – er verliert seinen Block.“\n\nPrüffragen (Stichprobe 4–6 Schüler, unbenotet):\n• Was ist vorher eine Gruppe? (die Leute aus einer Gegend)\n• Wo landen jetzt die Bauern des Adligen? (in verschiedenen Gruppen)\n• Warum hilft ihm das nicht mehr bei Abstimmungen? (überall in der Minderheit, kein Block)\nNicht weitergehen, bevor die dritte Frage sitzt.");
}

// 6 Was blieb gleich?
{ const s = slide(); title(s, "Was blieb gleich?", 2); tag(s, "zuhören");
  panel(s, "Die höchsten Ämter bekommen weiter nur die Reichen.", 0.45, 1.4, 9.1, 0.95, { line: C.terra });
  panel(s, "Niemand verliert Land – der Adel bleibt reich.", 0.45, 2.55, 9.1, 0.95, { line: C.terra });
  panel(s, "Abwägen heißt: Was wiegt schwerer – das Neue oder das Gebliebene?", 0.45, 3.9, 9.1, 1.0, { fill: C.dark, line: C.dark, color: C.gold, bold: true });
  num(s); s.addNotes("8–10 Min. Rückbezug auf Folie 4: Gefolgschaft wurde gebrochen; Ämter und Reichtum blieben. Prüffrage (2 Schüler): „Nenne eine Sache, die sich geändert hat, und eine, die blieb.“\n\nGelenk: „Diese Argumente stehen sortiert auf eurem Blatt. Heute schreibt ihr daraus euer Urteil.“");
}

// 7 Auftrag 1
{ const s = slide(); title(s, "Argumente lesen", 3); tag(s, "allein · 5 Min.");
  steps(s, ["Lies die Argumente auf der Vorderseite.", "Kreuze in jeder Spalte die zwei wichtigsten an.", "Lies die Zeile „Für den Adel“ besonders genau."], 1.45, 0.85);
  panel(s, "Die Argumente stehen auch an der Wand.", 0.45, 4.3, 9.1, 0.75, { line: C.soft });
  num(s); s.addNotes("10–15 Min. Blatt austeilen. Danach Folien 8 und 9 stehen lassen (Arbeitsgrundlage an der Wand). Umhergehen, mitlesen; bei Unsicherheit: „Was bedeutet das für den Adligen aus Folie 5?“");
}

// 8/9 Argumente an die Wand
const DAFUER = [["Neue Gruppen (Phylen)", "seine Bauern getrennt"], ["Rat der 500", "gemischt, kein Block"], ["Bürgerrecht über das Dorf", "nicht mehr die Familie"], ["Scherbengericht", "zu Mächtige verbannen"]];
const DAGEGEN = [["Ämter nur für Reiche", "Ämter bleiben beim Adel"], ["Areopag bleibt", "Reiche behalten Kontrolle"], ["Kein Land weggenommen", "Adel bleibt reich"], ["Ärmste ohne Ämter", "nur abstimmen, kein Amt"]];
[["Dafür spricht …", DAFUER, C.green], ["Dagegen spricht …", DAGEGEN, C.terra]].forEach(([t, arr, col]) => {
  const s = slide(); title(s, t, 3); tag(s, "Grundlage");
  arr.forEach(([a, b], i) => {
    const y = 1.35 + i * 0.98;
    txt(s, a, { x: 0.45, y, w: 4.4, h: 0.85, bold: true, color: col, valign: "middle" });
    txt(s, "→ " + b, { x: 4.9, y, w: 4.65, h: 0.85, valign: "middle" });
    s.addShape(pres.shapes.LINE, { x: 0.45, y: y + 0.9, w: 9.1, h: 0, line: { color: C.grey, width: 1 } });
  });
  num(s); s.addNotes("Kurzfassung der Vorderseite. Bleibt während der Lese- und Schreibphase sichtbar.");
});

// 10/11 Vormachen: Muster
{ const s = slide(); title(s, "So geht's – ein Muster", 3); tag(s, "a · Einstieg");
  txt(s, rich("**Im Folgenden soll beurteilt werden, ob** die Reformen des Kleisthenes der Vormachtstellung des Adels ein Ende gesetzt haben. **Um diese Frage zu beantworten, muss man klären, worauf die Macht des Adels beruhte:** auf den abhängigen Bauern, den Ämtern und dem Reichtum."),
    { x: 0.45, y: 1.35, w: 9.1, h: 3.9, valign: "top", lineSpacingMultiple: 1.05 });
  num(s); s.addNotes("15–19 Min. Vormachen, laut denken (Rückseite des Blatts):\n„Ich nehme zuerst die Frage auf – fast wörtlich. Dann kläre ich, was ich prüfe: die drei Punkte von vorhin. Daran messe ich gleich jedes Argument.“");
}
{ const s = slide(); title(s, "So geht's – ein Muster", 3); tag(s, "b · dafür");
  txt(s, rich("**Für ein Ende der Vormachtstellung spricht, dass** Kleisthenes die Bürger in zehn neue Gruppen einteilte, die aus Stadt, Küste und Binnenland gemischt wurden. **Dadurch** landeten die Bauern eines Adligen in verschiedenen Gruppen, **sodass** er sie nicht mehr geschlossen für sich abstimmen lassen konnte."),
    { x: 0.45, y: 1.35, w: 9.1, h: 3.9, valign: "top", lineSpacingMultiple: 1.05 });
  num(s); s.addNotes("Laut denken: „Erst sage ich, WAS Kleisthenes tat. Dann – das ist das Wichtigste – was das für den Adel bedeutet: ‚Dadurch …, sodass …‘. Genau diese zweite Hälfte steht auf der Vorderseite unter ‚Für den Adel‘.“\n\nPrüffrage (2–3 Schüler): „Welcher Teil des Satzes erklärt, warum das ein Argument ist?“ (Dadurch …, sodass …)");
}

// 12 Warnung vor dem Fehler
{ const s = slide(); title(s, "Achtung", 3); tag(s, "Warnung", C.terra);
  panel(s, "Nicht nur aufzählen, was Kleisthenes tat!", 0.45, 1.4, 9.1, 0.9, { line: C.terra, bold: true, color: C.terra });
  txt(s, "Erkläre bei jedem Argument, was es für den Adel bedeutet:", { x: 0.45, y: 2.6, w: 9.1, h: 0.6 });
  panel(s, "„Dadurch …, sodass …“", 0.45, 3.4, 9.1, 0.95, { fill: C.dark, line: C.dark, color: C.gold, bold: true, align: "center" });
  num(s); s.addNotes("Häufigster Fehler: reine Liste der Reformen („Er machte zehn Phylen. Er machte den Rat der 500.“) – das ist Beschreibung, kein Urteil.");
}

// 13 Auftrag 2: Sachurteil
{ const s = slide(); title(s, "Sachurteil schreiben", 3); tag(s, "allein · 15 Min.");
  steps(s, ["Schreibe ein zweites Argument dafür (b).", "Schreibe zwei Argumente dagegen (c).", "Gewichte und beantworte die Frage (d).", "Hake die Satzanfänge ab, die du nutzt."], 1.4, 0.8);
  panel(s, "Ins Heft · Rückseite des Blatts hilft", 0.45, 4.6, 9.1, 0.7, { line: C.soft });
  num(s); s.addNotes("19–34 Min. Umhergehen, mitlesen. Zuerst zu den Schwächeren. Impulse gestuft:\n1. „Welches Argument hast du angekreuzt? Lies die Zeile ‚Für den Adel‘.“\n2. „Beginne mit ‚Dagegen spricht jedoch, dass …‘.“\n3. Satz gemeinsam anfangen, Schüler beendet ihn.\n\nBei d: „Welche Seite wiegt für dich schwerer – das Neue (kein Block) oder das Gebliebene (Ämter, Reichtum)? Warum?“\n\nStichproben im Heft: Steht bei jedem Argument ein ‚Dadurch / sodass / weil‘?");
}

// 14 Auftrag 3: Werturteil *
{ const s = slide(); title(s, "Werturteil schreiben *", 3); tag(s, "allein · 5 Min.");
  steps(s, ["Kreuze einen Maßstab an.", "Bewerte die Reformen aus heutiger Sicht.", "Nenne ein Beispiel."], 1.4, 0.8);
  panel(s, "Gleichheit · Leistung · Schutz der Schwachen · gleiche Regeln für alle", 0.45, 4.0, 9.1, 1.0, { line: C.terra, color: C.terra, bold: true, align: "center" });
  num(s); s.addNotes("34–39 Min. * Bei Zeitnot als Hausaufgabe – das Auswerten (Folie 15) hat Vorrang. Satzanfänge: Rückseite, Baustein e.");
}

// 15 Auswerten: Fazits vergleichen
{ const s = slide(); title(s, "Stimmt das Fazit?", 4); tag(s, "alle prüfen");
  txt(s, "Zwei lesen ihr Fazit vor. Alle anderen prüfen:", { x: 0.45, y: 1.35, w: 9.1, h: 0.6 });
  steps(s, ["Beantwortet es die Frage?", "Sagt es, was schwerer wiegt?", "Steht ein „weil“ mit einem Argument?"], 2.15, 0.8);
  panel(s, "Ziel: Ich kann ein Sach- und ein Werturteil schreiben.", 0.45, 4.6, 9.1, 0.7, { line: C.terra, color: C.terra, bold: true });
  num(s); s.addNotes("39–45 Min. Nie streichen. Zwei Schüler gezielt aufrufen (ein sicherer, ein unsicherer), Fazit vorlesen. Danach 2–3 Stichproben zu den Prüfpunkten. Dann Lösungsfolie 16.\n\nMögliche gute Fazits: „teilweise“ (Gefolgschaft gebrochen, Ämter und Reichtum blieben) – „eher ja“ (Grundlage der Adelsmacht zerstört, Weg zur Demokratie) – „eher nein“ (Adel führte weiter, Ende erst 462). Alle richtig, wenn begründet.");
}

// 16–25 Musterlösung komplett: so sieht das fertige Urteil aus (Wunsch der Lehrkraft, 18.11.). Baustein a und erstes Argument von b wortgleich mit dem Blatt.
const Dq = require("./data.js");
const ZQ = (a, b) => "Z. " + Dq.zeileVon(a) + "–" + Dq.zeileVon(b);
{ const s = slide(C.dark);
  txt(s, "MUSTERLÖSUNG", { x: 0.7, y: 1.2, w: 8.6, h: 0.5, bold: true, color: C.gold, charSpacing: 3 });
  txt(s, "So sieht das fertige Urteil aus", { x: 0.7, y: 1.8, w: 8.6, h: 1.2, fontSize: 40, bold: true, fontFace: H, color: C.white });
  txt(s, "Sachurteil a–d · Werturteil e · fett = Satzanfänge", { x: 0.7, y: 3.3, w: 8.6, h: 0.6, color: "E8DFD0" });
  num(s); s.addNotes("Nach dem Auswerten (oder zu Beginn der nächsten Stunde): Die Schüler vergleichen ihr Heft Baustein für Baustein mit dem Muster und ergänzen in einer anderen Farbe. Andere Positionen („eher ja“, „eher nein“) sind ebenfalls richtig, wenn sie abwägen und begründen.");
}
const MUSTER = [
  ["a · Einstieg", "Sachurteil (1/8)", "**Im Folgenden soll beurteilt werden, ob** die Reformen des Kleisthenes der Vormachtstellung des Adels ein Ende gesetzt haben. **Um diese Frage zu beantworten, muss man klären, worauf die Macht des Adels beruhte:** auf den abhängigen Bauern, die ihrem Adligen geschlossen folgten, auf den Ämtern und auf dem Reichtum."],
  ["b · dafür", "Sachurteil (2/8)", "**Für ein Ende der Vormachtstellung spricht, dass** Kleisthenes die Bürger in zehn neue Gruppen (Phylen) einteilte, die aus Stadt, Küste und Binnenland zusammengelost wurden. **Dadurch** landeten die Bauern eines Adligen in verschiedenen Gruppen, **sodass** er sie nicht mehr geschlossen für sich abstimmen lassen konnte."],
  ["b · dafür", "Sachurteil (3/8)", "**Hinzu kommt, dass** im Rat der 500 nun Bürger aus ganz Attika gemischt saßen. **Dadurch konnte ein Adliger** dort **nicht mehr** auf einen eigenen Block zählen. **Auch die Quelle bestätigt dies:** Um 430 v. Chr. klagt ein Adliger, die „gemeinen Leute“ hätten es besser als die „Edlen“ (" + ZQ("Was die Staatsform", "die gemeinen Leute besser") + ")."],
  ["c · dagegen", "Sachurteil (4/8)", "**Dagegen spricht jedoch, dass** Solons Vermögensklassen bestehen blieben: Archont konnte nur werden, wer zu den Reichsten gehörte. **Dadurch** blieben die wichtigsten Ämter beim Adel. **Weiterhin blieb** der Areopag, ein Rat aus ehemaligen Archonten, der die Beamten beaufsichtigte."],
  ["c · dagegen", "Sachurteil (5/8)", "**Zwar** verlor der Adel seinen Block, **aber** er verlor kein Land und blieb reich und angesehen. Adlige führten weiter die Politik – Kleisthenes selbst war Adliger. **Auch die Quelle zeigt,** dass das Volk die militärische Führung anderen überließ (" + ZQ("Jedoch: Alle Ämter", "militärischen Führung") + ")."],
  ["d · Fazit", "Sachurteil (6/8)", "**Schwerer wiegt meines Erachtens** die neue Einteilung der Bürger, **weil** sie die wichtigste Grundlage der Adelsmacht zerstörte: die Bauern, die geschlossen hinter ihrem Adligen standen."],
  ["d · Fazit", "Sachurteil (7/8)", "**Zugegebenermaßen** blieben Ämter und Reichtum beim Adel, **allerdings** entschieden nun Bürger aus ganz Attika gemeinsam. **Daher komme ich zu dem Ergebnis, dass** die Reformen die Vormachtstellung des Adels nur teilweise beendet haben."],
  ["e · Werturteil", "Werturteil (8/8)", "**Gemessen am Maßstab der** Gleichheit **halte ich** die Reformen **für** einen wichtigen, aber unvollständigen Schritt, **weil** viele Bürger nun mitbestimmten, die Ämter aber vom Besitz abhingen. **Ein Beispiel dafür ist,** dass ein Thete zwar abstimmen, aber kein Amt übernehmen durfte."],
];
MUSTER.forEach(([bau, t, text]) => {
  const s = slide(); title(s, t, 4); tag(s, bau, C.gold, C.dark);
  txt(s, rich(text), { x: 0.45, y: 1.35, w: 9.1, h: 3.95, valign: "top", lineSpacingMultiple: 1.05 });
  num(s); s.addNotes("Musterlösung, Baustein " + bau + ". Vergleichen: Steht bei jedem Argument, was es für den Adel bedeutet (Dadurch …, sodass …, weil …)?" + (bau.startsWith("e") ? "\nAndere Maßstäbe ebenso richtig: Leistung (Theten rudern die Flotte, aber kein Amt), Schutz der Schwachen (Scherbengericht schützt vor einem neuen Tyrannen), gleiche Regeln (Abstimmung für alle gleich, Ämter nicht). Kritisch auch: Frauen, Metöken und Sklaven blieben ausgeschlossen." : ""));
});

pres.writeFile({ fileName: OUT }).then(() => console.log("done"));
