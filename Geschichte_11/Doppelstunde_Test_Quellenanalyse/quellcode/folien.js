// Folien Doppelstunde Test + Quellenanalyse – alle Texte für die Klasse mindestens 27 pt (Vorgabe der Lehrkraft).
const pptxgen = require("pptxgenjs");
const D = require("./data.js");
const OUT = process.argv[2];
const C = { dark: "2B2318", terra: "AD4426", sand: "F4ECDE", gold: "D9A441", muted: "5A5044", soft: "998C77", white: "FFFFFF", pale: "FBF8F2", green: "3C783C", blue: "3E5C76" };
const pres = new pptxgen(); pres.layout = "LAYOUT_16x9"; pres.title = "Test und Quelleninterpretation – Aristoteles über Kleisthenes";
const H = "Cambria", B = "Calibri";
const MIN = 27, TITLE = 36;
const PH = ["A", "V", "I", "V", "A"];

function txt(s, text, o) { s.addText(text, { isTextBox: true, fontFace: B, color: C.dark, margin: 0, valign: "top", fontSize: MIN, ...o }); }
function num(s) { s.slideNumber = { x: 9.3, y: 5.3, w: 0.4, h: 0.25, fontFace: B, fontSize: 12, color: C.soft, align: "right" }; }
function aviva(s, k, dark = false) {
  PH.forEach((l, i) => s.addText(l, { isTextBox: true, shape: pres.shapes.ROUNDED_RECTANGLE, rectRadius: 0.06, x: 7.15 + i * 0.5, y: 0.12, w: 0.44, h: 0.44,
    fill: { color: i === k ? C.terra : (dark ? "3A3024" : C.sand) }, color: i === k ? C.white : (dark ? "8C8170" : C.soft), fontFace: B, fontSize: MIN, bold: true, align: "center", valign: "middle", margin: 0 }));
}
function tag(s, t, fill = C.dark, color = C.white) {
  s.addText(t, { isTextBox: true, shape: pres.shapes.ROUNDED_RECTANGLE, rectRadius: 0.08, x: 7.15, y: 0.66, w: 2.4, h: 0.5, fill: { color: fill }, color, fontFace: B, fontSize: MIN, bold: true, align: "center", valign: "middle", margin: 0 });
}
function title(s, t, k, o = {}) { aviva(s, k, o.dark); txt(s, t, { x: 0.45, y: 0.15, w: 6.6, h: 1.05, fontSize: TITLE, bold: true, fontFace: H, valign: "middle", color: o.dark ? C.white : C.dark }); }
function card(s, x, y, w, h, fill = C.sand) { s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, rectRadius: 0.08, fill: { color: fill }, line: { color: fill } }); }
function boxText(s, t, x, y, w, h, fill, color, o = {}) { card(s, x, y, w, h, fill); txt(s, t, { x: x + 0.2, y, w: w - 0.4, h, color, valign: "middle", ...o }); }
function steps(s, items, y, gap, o = {}) {
  items.forEach((t, i) => {
    s.addText(String(i + 1), { isTextBox: true, shape: pres.shapes.OVAL, x: 0.45, y: y + i * gap, w: 0.5, h: 0.5, fill: { color: C.terra }, color: C.white, fontFace: H, fontSize: MIN, bold: true, align: "center", valign: "middle", margin: 0 });
    txt(s, t, { x: 1.1, y: y + i * gap - 0.02, w: 8.5, h: o.h || 0.55, valign: "middle" });
  });
}
const warn = (s, t, y = 4.15, h = 1.0) => boxText(s, "⚠ " + t, 0.45, y, 9.1, h, C.dark, C.gold, { bold: true });
const slide = (bg = C.white) => { const s = pres.addSlide(); s.background = { color: bg }; return s; };
const ZIEL = "Ich kann eine Textquelle nach den Arbeitsschritten analysieren: formale und inhaltliche Aspekte.";

// 1 Test
{ const s = slide(C.dark);
  txt(s, "TEIL 1", { x: 0.7, y: 0.7, w: 8.6, h: 0.5, bold: true, color: C.gold, charSpacing: 3 });
  txt(s, "Test · 45 Minuten", { x: 0.7, y: 1.2, w: 8.6, h: 0.9, fontSize: 40, bold: true, fontFace: H, color: C.white });
  ["Tische auseinander, nur Stift auf dem Tisch.", "Name auf das Blatt.", "Zusatzaufgabe am Ende: +1 Punkt.", "Fertig? Blatt umdrehen, still sitzen bleiben."].forEach((t, i) =>
    txt(s, "•  " + t, { x: 0.7, y: 2.35 + i * 0.6, w: 8.8, h: 0.55, color: "E8DFD0", valign: "middle" }));
  s.addNotes("Teil 1 (0–45 Min.). Tests verdeckt austeilen, gemeinsam starten. Uhrzeit für das Ende an die Tafel. 5 Minuten vor Schluss ansagen. Erwartungshorizont und Notenschlüssel: Test_Erwartungshorizont_Notenschluessel.pdf."); }

// 2 Titel Teil 2
{ const s = slide(C.dark);
  txt(s, "TEIL 2 · GESCHICHTE · KLASSE 11", { x: 0.7, y: 1.0, w: 8.6, h: 0.5, bold: true, color: C.gold, charSpacing: 3 });
  txt(s, "Wie interpretiert man eine Quelle?", { x: 0.7, y: 1.55, w: 8.6, h: 1.4, fontSize: 40, bold: true, fontFace: H, color: C.white });
  txt(s, "Aristoteles über die Reformen des Kleisthenes", { x: 0.7, y: 3.15, w: 8.6, h: 0.6, color: "E8DFD0" });
  s.addNotes("Tests einsammeln. Austeilen: Übersicht „So interpretierst du eine Textquelle“ (für den Hefter) und Arbeitsblatt „Quelle und Schreibplan“ (A4, doppelseitig)."); }

// 3 Ziel und Ablauf
{ const s = slide(); title(s, "Ziel und Ablauf", 0); tag(s, "2 Min.");
  boxText(s, ZIEL, 0.45, 1.3, 9.1, 1.6, C.dark, C.white, { bold: true });
  ["Richtig oder falsch? Kleisthenes", "Die Arbeitsschritte im Überblick", "Analyse: formale Aspekte", "Analyse: Inhalt · Kontext und Urteil*", "Abschlussaufgabe"].forEach((t, i) =>
    txt(s, (i + 1) + "  " + t, { x: 0.6, y: 3.0 + i * 0.47, w: 9, h: 0.47, valign: "middle" }));
  num(s); s.addNotes("0–2 Min. (nach dem Test). Ziel vorlesen, es steht auch an der Seitentafel. „Heute lernt ihr, wie man in der Klausur mit einer Quelle umgeht – Schritt für Schritt.“ * = nur bei Zeit."); }

// 4–5 Richtig oder falsch
const RF = [
  ["Das Verfassungsschaubild im Lehrbuch ist eine Quelle.", false, "Es ist eine Darstellung: heute gezeichnet."],
  ["Kleisthenes teilte die Bürger in zehn neue Phylen ein.", true],
  ["Kleisthenes nahm den Adligen ihr Land weg.", false, "Er mischte die Bürger – das Land blieb."],
  ["Jede Phyle hatte Gebiete in Stadt, Küste und Binnenland.", true],
];
[false, true].forEach(solved => { const s = slide(); title(s, solved ? "Lösung" : "Richtig oder falsch?", 1);
  tag(s, solved ? "Lösung" : "4 Min.", solved ? C.gold : C.dark, solved ? C.dark : C.white);
  RF.forEach(([t, ok, fix], i) => {
    const y = 1.3 + i * 1.0;
    card(s, 0.45, y, 9.1, 0.9, i % 2 ? C.white : C.pale);
    if (!solved) txt(s, (i + 1) + "  " + t, { x: 0.65, y, w: 8.8, h: 0.9, valign: "middle" });
    else if (ok) txt(s, "✓ " + t, { x: 0.65, y, w: 8.8, h: 0.9, valign: "middle", color: C.green });
    else txt(s, "✗ " + fix, { x: 0.65, y, w: 8.8, h: 0.9, valign: "middle", color: C.terra, bold: true });
  });
  num(s); s.addNotes(solved ? "6–7 Min. Vergleichen. Gelenkstelle: „Das Schaubild ist eine Darstellung. Heute arbeiten wir mit einer Quelle – mit dem, was ein Grieche selbst über Kleisthenes geschrieben hat.“"
    : "2–6 Min. Allein ins Heft: R oder F, falsche Aussagen berichtigen (2). Partnercheck (1). Stichproben bei 4–6 Schülern unterschiedlichen Niveaus (1). Quelle = Material aus der Vergangenheit; Darstellung = später darüber geschrieben oder gezeichnet."); });

// 6 Klausur
{ const s = slide(); title(s, "Warum das wichtig ist", 2); tag(s, "Klausur 25.11.");
  boxText(s, "In der Klausur bekommst du eine Quelle. Die Aufgabe heißt oft: „Interpretieren Sie die Quelle.“", 0.45, 1.3, 9.1, 1.6, C.sand, C.dark);
  boxText(s, "Mit den Arbeitsschritten weißt du immer, was als Nächstes kommt.", 0.45, 3.15, 9.1, 1.3, C.dark, C.gold, { bold: true });
  num(s); s.addNotes("7–8 Min. Bedeutung erklären (direkt). Die Arbeitsschritte sind die der Fachschaft (Übersicht im Hefter). Operatoren: analysieren (Schritt 2), einordnen/erläutern (3), beurteilen = Sachurteil, bewerten/Stellung nehmen = Werturteil (4)"); }

// 7 Arbeitsschritte
{ const s = slide(); title(s, "Die Arbeitsschritte", 2); tag(s, "Portion 1");
  steps(s, ["Themenfrage – Was will ich herausfinden?", "Analyse – Wer schreibt? Was steht drin?", "Kontext – Was geschah damals?", "Urteilen – erst Sach-, dann Werturteil"], 1.35, 0.72);
  warn(s, "Erst untersuchen, dann urteilen!", 4.3, 0.85);
  num(s); s.addNotes("8–10 Min. Erklärportion 1 (ca. 2 Min.) mit der Übersicht im Hefter: Die Schritte sind wie ein Trichter – erst die Frage, dann die Quelle selbst, dann der Zusammenhang, zuletzt das Urteil. Die Glaubwürdigkeit gehört ins Sachurteil, die eigene Meinung mit Maßstab ins Werturteil. Prüffragen an 4–6 Schüler: „Was kommt nach der Themenfrage?“ – „In welchem Schritt darf ich sagen, was ich davon halte?“ (4b) – „Wohin gehört: Ist die Quelle glaubwürdig?“ (Sachurteil)."); }

// 8 Themenfrage
{ const s = slide(); title(s, "1 Unsere Themenfrage", 2); tag(s, "vorgegeben", C.sand, C.dark);
  boxText(s, D.themenfrage, 0.45, 1.35, 9.1, 1.7, C.dark, C.white, { bold: true });
  txt(s, "Sie steht über allem – am Ende beantworten wir sie.", { x: 0.45, y: 3.35, w: 9.1, h: 0.9, valign: "middle" });
  num(s); s.addNotes("10 Min. Heute geben wir die Themenfrage vor. In der Klausur steht sie oft in der Aufgabe; sonst stellen die Schüler sie nach dem ersten Lesen selbst auf (Beispiel M11: „Ein Leben in katastrophalen Verhältnissen?“)."); }

// 8 Einleitung
{ const s = slide(); title(s, "2a TATTEA", 2); tag(s, "Portion 2");
  [["T  Titel", "Wie heißt der Text?"], ["A  Autor", "Wer? Amt, Stellung?"], ["T  Textsorte", "Brief, Rede, Bericht …"], ["T  Thema", "Worum geht es?"], ["E  Entstehung", "Wann und wo?"], ["A  Adressat", "Für wen gedacht?"]].forEach(([a, b], i) => {
    const y = 1.28 + i * 0.66;
    s.addText(a, { isTextBox: true, shape: pres.shapes.ROUNDED_RECTANGLE, rectRadius: 0.08, x: 0.45, y, w: 3.3, h: 0.58, fill: { color: C.terra }, color: C.white, fontFace: B, fontSize: MIN, bold: true, align: "left", valign: "middle", margin: [0, 0, 0, 8] });
    txt(s, b, { x: 3.95, y, w: 5.6, h: 0.58, valign: "middle" });
  });
  num(s); s.addNotes("10–11 Min. Erklärportion 2: fünf Fragen, alle Antworten stehen im Kasten zum Autor oder unter der Quelle (Nachweis)."); }

// 9–10 Modell Solon
{ const s = slide(); title(s, "So geht's – ein Beispiel", 2); tag(s, "Lehrkraft");
  D.modell.felder.forEach(([a, b], i) => {
    const y = 1.28 + i * 0.66;
    txt(s, a, { x: 0.45, y, w: 3.3, h: 0.58, valign: "middle", bold: true, color: C.terra });
    txt(s, b, { x: 3.95, y, w: 5.6, h: 0.58, valign: "middle" });
  });
  num(s); s.addNotes("11–12 Min. Vormachen an einer anderen Quelle (Solon kennen sie): laut denken. „Wer hat das geschrieben? Solon selbst – also war er dabei. Was für ein Text? Ein Gedicht. Wann? Um 590, kurz nach seinen Reformen. Für wen? Für die Athener. Worüber? Er verteidigt seine Reformen.“"); }
{ const s = slide(); title(s, "Daraus wird ein Satz", 2); tag(s, "Lehrkraft");
  boxText(s, "„" + D.modell.satz + "“", 0.45, 1.3, 9.1, 2.6, C.sand, C.dark, { italic: true });
  warn(s, "Lebensdaten prüfen: War der Autor dabei?", 4.15, 0.9);
  num(s); s.addNotes("12–13 Min. Satz vorlesen, Bausteine zeigen. Warnung vor dem Fehler: Viele schreiben „Augenzeuge“, obwohl der Autor viel später lebte. Prüffragen an 3–4 Schüler: „Woran erkennst du, ob der Autor dabei war?“ (Lebensdaten/Entstehungszeit mit dem Ereignis vergleichen)."); }

// 11 Auftrag Einleitung
{ const s = slide(); title(s, "Jetzt du: formale Aspekte", 3); tag(s, "4 Min.");
  steps(s, ["Lies den Kasten „Zum Autor“ auf dem Blatt.", "Fülle die sechs TATTEA-Felder in 2a aus.", "Schreibe einen Einleitungssatz.", "Fertig? Lies die Quelle schon einmal."], 1.35, 0.85);
  num(s); s.addNotes("13–17 Min. Erste Anwendung, allein (leicht, alle schaffen sie). Durch die Reihen gehen und mitlesen: Steht „Augenzeuge“ da? Steht die Zeit drin? Folie 12 (Kasten) groß stehen lassen."); }

// 12 Infokasten
{ const s = slide(); title(s, "Zum Autor und zur Schrift", 3);
  ["Aristoteles (384–322 v. Chr.): Philosoph, lebte in Athen, kein Bürger.", "„Der Staat der Athener“, um 330 v. Chr., für seine Schüler und gebildete Leser.", "Kleisthenes’ Reform war damals rund 180 Jahre her."].forEach((t, i) =>
    boxText(s, t, 0.45, 1.3 + i * 1.3, 9.1, 1.15, i % 2 ? C.white : C.sand, C.dark));
  num(s); s.addNotes("Arbeitsgrundlage an die Wand (gekürzt; vollständig auf dem Blatt)."); }

// 13 Lösung Einleitung
{ const s = slide(); title(s, "Lösung: formale Aspekte", 3); tag(s, "Lösung", C.gold, C.dark);
  txt(s, "„" + D.einleitung.satz + "“", { x: 0.45, y: 1.3, w: 9.1, h: 3.9, italic: true });
  num(s); s.addNotes("17–19 Min. Zwei Schüler lesen ihren Satz vor (unterschiedliches Niveau), dann aufdecken. Vergleichen, im Blatt ergänzen. Gelenkstelle: „Jetzt wissen wir, wer spricht. Als Nächstes: Was sagt er eigentlich?“ Ist etwas unklar, hier nachsteuern, nicht weitergehen."); }

// 14 Inhalt – Regeln
{ const s = slide(); title(s, "2b Inhaltliche Aspekte", 2); tag(s, "Portion 3");
  steps(s, ["Wesentliche Aussage je Abschnitt.", "Eigene Worte, indirekte Rede: „er habe …“", "Zeilen angeben: (Z. 6)", "Schlüsselbegriffe und Textsprache"], 1.35, 0.72);
  warn(s, "Noch keine Meinung – die kommt erst in Schritt 4.", 4.3, 0.85);
  num(s); s.addNotes("19–21 Min. Erklärportion 3: Indirekte Rede zeigt, dass ich wiedergebe, was der Autor sagt – nicht, was stimmt. Prüffragen an 4–6 Schüler: „Mach daraus indirekte Rede: Er teilte die Bürger ein.“ (Er habe die Bürger eingeteilt.) – „Darf ich schreiben: Das war klug von ihm?“ (Nein, erst in Schritt 4)."); }

// 15 Modell Abschnitt 1
{ const s = slide(); title(s, "So geht's: Z. 1–4", 2); tag(s, "Lehrkraft");
  txt(s, "„" + D.abschnitte[0][3] + "“", { x: 0.45, y: 1.3, w: 9.1, h: 3.9, italic: true });
  num(s); s.addNotes("21–23 Min. Vormachen, laut denkend: „Zeile 1 bis 4 gehören zusammen: Es geht um den Machtkampf. Wo endet der Abschnitt? Ab Zeile 5 kommt etwas Neues – die Phylen. Was ist der Kern? Kleisthenes gewinnt das Volk. Ich schreibe es in indirekter Rede und setze die Zeilen dahinter.“ Das Beispiel steht schon im Schreibplan (✓)."); }

// 16 Auftrag Inhalt
{ const s = slide(); title(s, "Jetzt ihr: Inhalt", 3); tag(s, "10 Min.");
  steps(s, ["Lest Z. 5–14 zu zweit.", "Schreibt Z. 5–8 mit dem Satzanfang.", "Schreibt Z. 9–12 mit dem Satzanfang.", "Schreibt Z. 12–14 ganz allein.", "Notiert Begriffe, kreuzt die Sprache an."], 1.35, 0.75);
  num(s); s.addNotes("23–33 Min. Angeleitet → selbstständig (Teillösungen im Schreibplan werden weniger). Partnerarbeit, beide schreiben. Umhergehen, mitlesen; Stolpersteine und Hilfsimpulse in der Handreichung. Die Quelle steht auf den nächsten Folien groß an der Wand."); }

// 17–19 Quelle
[[0, 4, "Z. 1–4"], [4, 8, "Z. 5–8"], [8, 14, "Z. 9–14"]].forEach(([a, b, z]) => { const s = slide(); title(s, "Aristoteles", 3); tag(s, z, C.sand, C.dark);
  txt(s, D.zeilen.slice(a, b).join(" "), { x: 0.45, y: 1.3, w: 9.1, h: 4.0, fontSize: MIN });
  num(s); s.addNotes("Arbeitsgrundlage an die Wand. Gleicher Wortlaut wie auf dem Blatt."); });

// 20–22 Lösung Inhalt
D.abschnitte.slice(1).forEach(([z, , , l]) => { const s = slide(); title(s, "Lösung: " + z, 3); tag(s, "Lösung", C.gold, C.dark);
  txt(s, "„" + l + "“", { x: 0.45, y: 1.3, w: 9.1, h: 3.9, italic: true });
  num(s); s.addNotes(z === "Z. 5–8" ? "33–36 Min. Je Abschnitt ein Paar vom Platz vorlesen lassen, dann aufdecken. Andere Formulierungen sind richtig, wenn Kern, indirekte Rede und Zeilen stimmen." : "Vergleichen, im Blatt ergänzen."); });

// Lösung Schlüsselbegriffe und Textsprache
{ const s = slide(); title(s, "Lösung: Begriffe, Sprache", 3); tag(s, "Lösung", C.gold, C.dark);
  boxText(s, D.begriffe, 0.45, 1.3, 9.1, 1.2, C.sand, C.dark, { bold: true });
  txt(s, "Sprache: sachlich und erklärend – er nennt mit „damit …“ den Zweck (Z. 7, 11, 13).", { x: 0.45, y: 2.75, w: 9.1, h: 1.6 });
  num(s); s.addNotes("Kurz vergleichen. Gelenkstelle: „Aristoteles erklärt, was Kleisthenes wollte – aber konnte er das wissen? Das prüfen wir im Urteil.“"); }

// 23 Einordnen und Urteilen*
{ const s = slide(); title(s, "Kontext und Urteil*", 3); tag(s, "nur bei Zeit", C.sand, C.dark);
  boxText(s, "Kontext: Nach dem Sturz der Tyrannis (510) kämpfte der Adel wieder um die Macht.", 0.45, 1.3, 9.1, 1.65, C.sand, C.dark);
  boxText(s, "Sachurteil: Aristoteles schrieb 180 Jahre später – er war nicht dabei.", 0.45, 3.15, 9.1, 1.4, C.dark, C.gold, { bold: true });
  num(s); s.addNotes("*36–38 Min. Nur bei Zeit, sonst direkt zur Abschlussaufgabe. Kurz direkt erklären, zwei Prüffragen: „Woher wusste Aristoteles das?“ (aus älteren Schriften) – „Ist die Quelle deshalb wertlos?“ (nein, aber kein Augenzeugenbericht; er deutet die Reform aus seiner Zeit). Kontext, Sach- und Werturteil werden in der nächsten Doppelstunde vertieft."); }

// 24 Abschlussaufgabe
{ const s = slide(); title(s, "Abschlussaufgabe", 4); tag(s, "4 Min.");
  txt(s, "Ein Mitschüler zeigt dir seinen Klausur-Anfang:", { x: 0.45, y: 1.25, w: 9.1, h: 0.5, valign: "middle" });
  boxText(s, "„Aristoteles war dabei, als Kleisthenes die Demokratie einführte. Kleisthenes teilte die Bürger in zehn Phylen ein. Das war eine sehr gute Idee.“", 0.45, 1.8, 9.1, 2.1, C.sand, C.dark, { italic: true });
  boxText(s, "Gib ihm drei Ratschläge: „Ich rate dir …, weil …“", 0.45, 4.1, 9.1, 0.9, C.dark, C.white, { bold: true });
  num(s); s.addNotes("38/40–43 Min. Prüft genau das Ziel (formale und inhaltliche Aspekte). Allein ins Heft, dann gezielt 3 Schüler unterschiedlichen Niveaus je einen Ratschlag vorlesen lassen. Diese Phase nie kürzen. Mögliche Ratschläge in der Handreichung."); }

// 25 Lösung Abschluss
{ const s = slide(); title(s, "So könnte es lauten", 4); tag(s, "Lösung", C.gold, C.dark);
  ["Ich rate dir, die Lebensdaten zu prüfen, weil Aristoteles 180 Jahre später schrieb.", "Ich rate dir, indirekte Rede und Zeilen zu nutzen, weil du wiedergibst, was er sagt.", "Ich rate dir, die Meinung wegzulassen, weil sie erst in Schritt 4 gehört."].forEach((t, i) =>
    boxText(s, t, 0.45, 1.3 + i * 1.3, 9.1, 1.15, i % 2 ? C.white : C.sand, C.dark));
  num(s); s.addNotes("43–44 Min. Weitere gute Antworten: TATTEA-Angaben fehlen (Titel, Textsorte, Entstehung, Adressat); „Demokratie eingeführt“ steht so nicht in der Quelle."); }

// 26 Ziel erreicht?
{ const s = slide(C.dark); aviva(s, 4, true);
  txt(s, "Haben wir das Ziel erreicht?", { x: 0.45, y: 0.15, w: 6.6, h: 1.05, fontSize: TITLE, bold: true, fontFace: H, color: C.white, valign: "middle" });
  txt(s, ZIEL, { x: 0.45, y: 1.5, w: 9.1, h: 1.6, color: C.gold, bold: true });
  txt(s, "Nächstes Mal: Kontext, Sach- und Werturteil.", { x: 0.45, y: 3.6, w: 9.1, h: 0.9, color: "E8DFD0" });
  num(s); s.addNotes("44–45 Min. Rückbezug aufs Ziel. Übersicht und Schreibplan im Hefter aufbewahren – sie sind die Vorlage für die Klausur."); }

pres.writeFile({ fileName: OUT }).then(() => console.log("done"));
