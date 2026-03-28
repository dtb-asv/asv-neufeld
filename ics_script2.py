import pandas as pd
from ics import Calendar, Event
import hashlib
import subprocess
import pytz

# -----------------------------
# Einstellungen
# -----------------------------
datei = "Terminefussball_2026_Frühjahr.xlsx"
sheet = "ICS2"
tz = pytz.timezone("Europe/Vienna")

# -----------------------------
# Gegner automatisch erkennen
# -----------------------------
def get_gegner(row):
    if pd.notna(row.get("GEGNER")):
        return str(row["GEGNER"])

    titel = str(row["TITEL"])

    if "vs" in titel.lower():
        return titel.split("vs")[-1].strip()
    if "@" in titel:
        return titel.split("@")[-1].strip()

    return "Unbekannt"

# -----------------------------
# Event erstellen
# -----------------------------
def create_event(row):
    typ = str(row.get("TYP", "")).strip().lower()
    liga = str(row.get("LIGA", ""))
    gegner = get_gegner(row)

    # Emoji
    if typ == "heim":
        prefix = "🏠 Heim"
    elif typ == "auswärts":
        prefix = "🚗 Auswärts"
    else:
        prefix = "📅 Spiel"
    if typ == "heim":
       titel = f"{prefix}: {liga} ASV Neufeld vs {gegner}"
    elif typ == "auswärts":
       titel = f"{prefix}: {liga} {gegner} vs ASV Neufeld"

    # Zeiten
    start_naiv = pd.to_datetime(f"{row['DATUM']} {row['STARTZEIT']}")
    end_naiv = pd.to_datetime(f"{row['DATUM']} {row['ENDZEIT']}")

    start = tz.localize(start_naiv, is_dst=None)
    end = tz.localize(end_naiv, is_dst=None)

    ort = str(row.get("ORT", ""))
    
    beschreibung = str(row.get("Beschreibung", "") or "")
    if beschreibung.lower() == "nan":
       beschreibung = ""
    beschreibung = str(row.get("BESCHREIBUNG", ""))
    
    beschreibung_full = f"""
{beschreibung}

⚽ Gegner: {gegner}
🏆 Liga: {liga}
"""

    uid = hashlib.md5(f"{titel}{start}{end}_{ort}".encode()).hexdigest()

    e = Event()
    e.name = titel
    e.begin = start
    e.end = end
    e.location = ort
    e.description = beschreibung_full
    e.uid = uid

    return e

# -----------------------------
# Excel einlesen
# -----------------------------
df = pd.read_excel(datei, sheet_name=sheet)

# -----------------------------
# 1️⃣ Gesamt-Kalender
# -----------------------------
cal_all = Calendar()

for _, row in df.iterrows():
    try:
        cal_all.events.add(create_event(row))
    except Exception as ex:
        print("Fehler:", ex)

with open("alle_spiele.ics", "w", encoding="utf-8") as f:
    f.writelines(cal_all)

print("✅ Gesamt-Kalender erstellt")

# -----------------------------
# 2️⃣ Kalender pro Liga
# -----------------------------
ligen = df["LIGA"].dropna().unique()

for liga in ligen:
    cal = Calendar()
    liga_df = df[df["LIGA"] == liga]

    for _, row in liga_df.iterrows():
        try:
            cal.events.add(create_event(row))
        except Exception as ex:
            print("Fehler:", ex)

    filename = f"{liga}.ics".replace(" ", "_")

    with open(filename, "w", encoding="utf-8") as f:
        f.writelines(cal)

    print(f"✅ {filename} erstellt")

# -----------------------------
# Git Upload (smart)
# -----------------------------
try:
    result = subprocess.run(["git", "status", "--porcelain"], capture_output=True, text=True)

    if result.stdout.strip() == "":
        print("ℹ️ Keine Änderungen")
    else:
        subprocess.run(["git", "add", "."], check=True)
        subprocess.run(["git", "commit", "-m", "Update Kalender"], check=True)
        subprocess.run(["git", "push"], check=True)
        print("🚀 GitHub aktualisiert!")

except subprocess.CalledProcessError as e:
    print("❌ Git Fehler:", e)