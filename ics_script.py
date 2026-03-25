import pandas as pd
from ics import Calendar, Event
import hashlib
import subprocess
import pytz

# -----------------------------
# Einstellungen
# -----------------------------
datei = "Heimspiele_ASVNEUFELD_ICS.xlsx"
sheet = "ICS"
ics_file = "ASV_Neufeld_Heimspiele.ics"

# Zeitzone Österreich
tz = pytz.timezone("Europe/Vienna")

# -----------------------------
# Excel einlesen
# -----------------------------
df = pd.read_excel(datei, sheet_name=sheet)
cal = Calendar()

for index, row in df.iterrows():
    try:
        titel = str(row["TITEL"])
        ort = str(row["Ort"])
        beschreibung = str(row["Beschreibung"])

        # Zeiten korrekt mit Zeitzone (MEZ/CEST automatisch)
        start_naiv = pd.to_datetime(f"{row['DATUM']} {row['Startzeit']}")
        end_naiv = pd.to_datetime(f"{row['DATUM']} {row['Endzeit']}")

        start = tz.localize(start_naiv, is_dst=None)
        end = tz.localize(end_naiv, is_dst=None)

        # UID neu berechnen bei Zeitänderung
        uid = hashlib.md5(f"{titel}_{start}_{end}_{ort}_{beschreibung}".encode()).hexdigest()

        e = Event()
        e.name = titel
        e.begin = start
        e.end = end
        e.location = ort
        e.description = beschreibung
        e.uid = uid

        cal.events.add(e)

    except Exception as ex:
        print(f"Fehler in Zeile {index}: {ex}")

# ICS speichern
with open(ics_file, "w", encoding="utf-8") as f:
    f.writelines(cal)

print("✅ ICS Datei erstellt mit korrekten Zeiten für MEZ/CEST!")

# -----------------------------
# Git Auto Upload
# -----------------------------
try:
    subprocess.run(["git", "add", "."], check=True)
    subprocess.run(["git", "commit", "-m", "Update Kalender"], check=True)
    subprocess.run(["git", "push"], check=True)
    print("🚀 GitHub wurde automatisch aktualisiert!")
except subprocess.CalledProcessError as e:
    print("❌ Git Fehler:", e)