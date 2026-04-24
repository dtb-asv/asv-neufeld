from PIL import Image, ImageDraw, ImageFont
import pandas as pd
from datetime import datetime, timedelta
import os
import zipfile

# -----------------------------
# EINSTELLUNGEN
# -----------------------------
excel = "Terminefussball_2026_Frühjahr.xlsx"
sheet = "ICS2"
background_path = "background.jpg"



modus = "single"  # "single" oder "alle"
ziel_datum_input = "2026-04-20"
if ziel_datum_input:
    ziel_datum = datetime.strptime(ziel_datum_input, "%Y-%m-%d")
else:
    ziel_datum = datetime.today()

# -----------------------------
# Excel laden
# -----------------------------
df = pd.read_excel(excel, sheet_name=sheet)
df["DATUM"] = pd.to_datetime(df["DATUM"])

# Nachwuchs (ohne KM & U23)
df = df[~df["LIGA"].str.contains("KM|U23", case=False, na=False)]

# Jahr bestimmen
jahr = df["DATUM"].dt.year.mode()[0]

# -----------------------------
# Wochen bestimmen
# -----------------------------
if modus == "single":
    ziel_datum = datetime.strptime(ziel_datum_input, "%Y-%m-%d")
    df["KW"] = df["DATUM"].dt.isocalendar().week
    wochen = [ziel_datum.isocalendar()[1]]

elif modus == "alle":
    df["KW"] = df["DATUM"].dt.isocalendar().week
    wochen = sorted(df["KW"].dropna().unique())

# -----------------------------
# Fonts
# -----------------------------
font_title = ImageFont.truetype("arial.ttf", 60)
font_big = ImageFont.truetype("arial.ttf", 45)
font_small = ImageFont.truetype("arial.ttf", 28)
font_kw = ImageFont.truetype("arial.ttf", 38)

# -----------------------------
# Layout
# -----------------------------
y_start = 220
block_height = 130
max_spiele = 7

# -----------------------------
# Bild Funktion
# -----------------------------
def neues_bild():
    img = Image.open(background_path).convert("RGBA")
    img = img.resize((1080, 1350))
    overlay = Image.new("RGBA", img.size, (0, 0, 0, 120))
    return Image.alpha_composite(img, overlay)

# -----------------------------
# HAUPTSCHLEIFE (alle Wochen)
# -----------------------------
for kw in wochen:

    df_woche = df[df["KW"] == kw].sort_values(by="DATUM")

    if df_woche.empty:
        continue

    # Ordner erstellen
    folder = f"output/{jahr}/KW{kw}"
    os.makedirs(folder, exist_ok=True)

    count = 0
    seite = 1

    img = neues_bild()
    draw = ImageDraw.Draw(img)

    # Header
    draw.text((50, 50), "SPIELPLAN NACHWUCHS", font=font_title, fill="white")
    draw.text((50, 120), f"KW {kw}", font=font_kw, fill=(200,200,200))

    y = y_start

    # Spiele
    for _, row in df_woche.iterrows():

        if count == max_spiele:
            img.convert("RGB").save(f"{folder}/Spielplan_KW{kw}_Seite_{seite}.png")

            seite += 1
            count = 0
            img = neues_bild()
            draw = ImageDraw.Draw(img)

            # Header neu
            draw.text((50, 50), "SPIELPLAN NACHWUCHS", font=font_title, fill="white")
            draw.text((50, 120), f"KW {kw}", font=font_kw, fill=(200,200,200))

            y = y_start

        liga = str(row.get("LIGA", ""))
        gegner = str(row.get("GEGNER", ""))
        datum = row["DATUM"].strftime("%d.%m.%Y")
        zeit = str(row.get("STARTZEIT"))
        ort = str(row.get("ORT", ""))
        typ = str(row.get("TYP", "")).strip().lower()

        if typ == "heim":
            label = "Heim"
            spiel = f"ASV Neufeld vs {gegner}"
        else:
            label = "Auswärts"
            spiel = f"{gegner} vs ASV Neufeld"

        # Zeichnen
        draw.text((50, y), liga, font=font_big, fill="white")

        text = f"{datum} | {zeit} Uhr | {label} {ort}"
        draw.text((220, y+10), text, font=font_small, fill=(220,220,220))

        draw.text((220, y+55), spiel.upper(), font=font_small, fill="white")

        draw.line((50, y+110, 1000, y+110), fill=(255,255,255,60), width=1)

        y += block_height
        count += 1

    # letzte Seite speichern
    img.convert("RGB").save(f"{folder}/Spielplan_KW{kw}_Seite_{seite}.png")

print("🚀 Alle Spielpläne erstellt!")
# -----------------------------
# ZIP EXPORT
# -----------------------------
zip_name = f"output_{jahr}.zip"

with zipfile.ZipFile(zip_name, 'w', zipfile.ZIP_DEFLATED) as zipf:
    for root, dirs, files in os.walk(f"output/{jahr}"):
        for file in files:
            filepath = os.path.join(root, file)
            arcname = os.path.relpath(filepath, f"output/{jahr}")
            zipf.write(filepath, arcname)

print(f"📦 ZIP erstellt: {zip_name}")