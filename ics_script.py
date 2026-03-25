import pandas as pd
from ics import Calendar, Event
import hashlib
import subprocess
import pytz

# Datei + Reiter
datei = "Heimspiele_ASVNEUFELD_ICS.xlsx"
sheet = "ICS"

# Output direkt im Repo-Ordner
ics_file = "ASV_Neufeld_Heimspiele.ics"

# Excel laden
df = pd.read_excel(datei, sheet_name=sheet)
cal = Calendar()

# Zeitzone Österreich
tz = pytz.timezone("Europe/Vienna")

for index, row in df.iterrows():
    try:
        titel = str(row["TITEL"])
        ort = str(row["Ort"])
        beschreibung = str(row["Beschreibung"])

        start = tz.localize(pd.to_datetime(f"{row['DATUM']} {row['Startzeit']}"))
        end = tz.localize(pd.to_datetime(f"{row['DATUM']} {row['Endzeit']}"))

        uid = hashlib.md5(f"{titel}_{start}_{ort}".encode()).hexdigest()

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

print("✅ ICS erstellt!")

# -----------------------------
# 🔄 Git Auto Upload
# -----------------------------
try:
    subprocess.run(["git", "add", "."], check=True)
    subprocess.run(["git", "commit", "-m", "Update Kalender"], check=True)
    subprocess.run(["git", "push"], check=True)

    print("🚀 GitHub wurde automatisch aktualisiert!")

except Exception as e:
    print("❌ Git Fehler:", e)