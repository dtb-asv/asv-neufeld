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
# Gegner erkennen
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

    # Emoji + Titel
    if typ == "heim":
        prefix = "🏠 Heim"
    elif typ == "auswärts":
        prefix = "🚗 Auswärts"
    else:
        prefix = "📅 Spiel"

    titel = f"{prefix}: ASV Neufeld vs {gegner}"
    # Status mitnehmen
    status = str(row.get("STATUS", "Aktiv")).lower()

    if status == "abgesagt":
       titel = f"❌ ABGESAGT: {titel}"

    # Zeiten (DST sicher)
    start_naiv = pd.to_datetime(f"{row['DATUM']} {row['STARTZEIT']}")
    end_naiv = pd.to_datetime(f"{row['DATUM']} {row['ENDZEIT']}")

    start = tz.localize(start_naiv, is_dst=None)
    end = tz.localize(end_naiv, is_dst=None)

    ort = str(row.get("Ort", ""))
    beschreibung = row.get("BESCHREIBUNG", "")
    if pd.isna(beschreibung):
      beschreibung = ""
    else:
      beschreibung = str(beschreibung)

    beschreibung_full = f"""
{beschreibung}

⚽ Gegner: {gegner}
🏆 Liga: {liga}

"""

    uid = hashlib.md5(f"{titel}_{start}_{end}_{ort}_{status}".encode()).hexdigest()

    e = Event()
    e.name = titel
    e.begin = start
    e.end = end
    e.location = ort
    e.description = beschreibung_full
    e.uid = uid

    # ⏰ Erinnerung 2 Stunden vorher
    e.alarms.append({"trigger": -7200})

    return e

# -----------------------------
# Excel einlesen
# -----------------------------
df = pd.read_excel(datei, sheet_name=sheet)

# -----------------------------
# Kalender erstellen
# -----------------------------
cal_all = Calendar()
cal_home = Calendar()
cal_away = Calendar()

ligen_cals = {}

for liga in df["LIGA"].dropna().unique():
    ligen_cals[liga] = Calendar()

# -----------------------------
# Events verteilen
# -----------------------------
for _, row in df.iterrows():
    try:
        event = create_event(row)
        cal_all.events.add(event)

        typ = str(row.get("TYP", "")).lower()
        liga = row.get("LIGA")

        if typ == "heim":
            cal_home.events.add(event)
        elif typ == "auswärts":
            cal_away.events.add(event)

        if pd.notna(liga):
            ligen_cals[liga].events.add(event)

    except Exception as ex:
        print("❌ Fehler:", ex)

# -----------------------------
# Dateien speichern
# -----------------------------
def save_cal(name, cal):
    filename = f"{name}.ics".replace(" ", "_")
    with open(filename, "w", encoding="utf-8") as f:
        f.writelines(cal)
    print(f"✅ {filename} erstellt")

save_cal("alle_spiele", cal_all)
save_cal("heimspiele", cal_home)
save_cal("auswaertsspiele", cal_away)

for liga, cal in ligen_cals.items():
    save_cal(liga, cal)

# -----------------------------
# Git Upload
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