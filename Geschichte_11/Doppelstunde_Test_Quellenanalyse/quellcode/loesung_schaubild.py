# Lösung Testaufgabe 4: trägt die Antworten in genau das stumme Schaubild des Testblatts (stumm.png) ein,
# damit die Lehrkraft Feld für Feld vergleichen kann. Erzeugt einen ImageMagick-Befehl (PIL ist hier nicht nutzbar).
# Aufruf: python3 loesung_schaubild.py <stumm.png> <ausgabe.png>
import math, subprocess, sys

SRC, OUT = sys.argv[1], sys.argv[2]
W, H = 3200, 2160
ROT = "#AD4426"          # Lösungsfarbe (Terrakotta)
GRUEN, KREUZ = "#2E7D32", "#AD4426"
FONT_B = "/usr/share/fonts/truetype/crosextra/Carlito-Bold.ttf"
COLS = [430, 1210, 1990, 2770]          # Spaltenmitten (Ämter und Klassen)
cmd = ["convert", SRC]

def text(t, cx, cy, size, color=ROT, font=FONT_B):
    # gravity center: Versatz von der Bildmitte
    cmd.extend(["-font", font, "-pointsize", str(size), "-fill", color, "-stroke", "none", "-gravity", "center",
                "-annotate", f"{cx - W / 2:+.0f}{cy - H / 2:+.0f}", t])

def pfeil(x1, y1, x2, y2, color=ROT, breite=9):
    a = math.atan2(y2 - y1, x2 - x1); L, B = 46, 22
    bx, by = x2 - L * math.cos(a), y2 - L * math.sin(a)
    p1 = (bx + B * math.sin(a), by - B * math.cos(a)); p2 = (bx - B * math.sin(a), by + B * math.cos(a))
    cmd.extend(["-gravity", "NorthWest", "-stroke", color, "-strokewidth", str(breite), "-fill", "none",
                "-draw", f"line {x1},{y1} {bx:.0f},{by:.0f}",
                "-stroke", "none", "-fill", color, "-draw", f"polygon {x2},{y2} {p1[0]:.0f},{p1[1]:.0f} {p2[0]:.0f},{p2[1]:.0f}"])

def haken(cx, cy):
    cmd.extend(["-gravity", "NorthWest", "-stroke", GRUEN, "-strokewidth", "11", "-fill", "none",
                "-draw", f"polyline {cx - 26},{cy + 2} {cx - 6},{cy + 24} {cx + 30},{cy - 26}"])

def kreuz(cx, cy):
    d = 24
    cmd.extend(["-gravity", "NorthWest", "-stroke", KREUZ, "-strokewidth", "10", "-fill", "none",
                "-draw", f"line {cx - d},{cy - d} {cx + d},{cy + d}", "-draw", f"line {cx - d},{cy + d} {cx + d},{cy - d}"])

# a) Ämter (oben) und Volksversammlung (unten)
for cx, t, s in [(COLS[0], "Areopag", 72), (COLS[1], "Archonten", 72), (COLS[2], "Rat der 400", 72)]:
    text(t, cx, 250, s)
text("Geschworenen-", COLS[3], 215, 64); text("gericht", COLS[3], 290, 64)
text("Volksversammlung", 1229, 1990, 70)

# Klassennamen auf der Linie und Schefferzahlen in der Lücke
for cx, n in zip(COLS, ["Fünfhundertscheffler", "Ritter", "Hopliten", "Theten"]):
    text(n, cx, 1110, 50 if len(n) > 12 else 60)
for x, z in zip([301, 1082, 1861, 2706], ["500", "300", "200", "200"]):
    text(z, x, 1212, 54)

# b) Rechte ankreuzen (Zeilen: abstimmen, klagen, in den Rat, Archont werden)
BOX_X = [157, 936, 1715, 2496]; ROWS = [1371, 1507, 1643, 1779]
RECHTE = [[1, 1, 1, 1], [1, 1, 1, 0], [1, 1, 1, 0], [1, 1, 0, 0]]
for bx, r in zip(BOX_X, RECHTE):
    for y, ok in zip(ROWS, r):
        (haken if ok else kreuz)(bx, y)

# c) Pfeile: Wer kann in welches Amt gewählt werden? (Arbeitsanweisung in der Mitte abdecken)
cmd.extend(["-gravity", "NorthWest", "-stroke", "none", "-fill", "white", "-draw", "rectangle 580,610 2620,715"])
pfeil(653, 995, 984, 395)      # 1. Klasse → Archonten
pfeil(728, 995, 1744, 395)     # 1. Klasse → Rat der 400
pfeil(1208, 995, 1984, 395)    # 2. Klasse → Rat der 400
pfeil(2029, 995, 2077, 395)    # 3. Klasse → Rat der 400
text("kein Amt", COLS[3], 900, 60)

cmd.append(OUT)
subprocess.run(cmd, check=True)
print("done")
