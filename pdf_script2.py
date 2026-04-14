from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.units import cm
from reportlab.platypus import Image
from reportlab.lib.utils import ImageReader
import pandas as pd

# Excel laden
df = pd.read_excel("Terminefussball_2026_Frühjahr.xlsx", sheet_name="ICS2")

styles = getSampleStyleSheet()

doc = SimpleDocTemplate("Spielplan_ASV_Neufeld_Design.pdf", pagesize=landscape(A4))

# doc = SimpleDocTemplate("Spielplan_ASV_Neufeld_Design.pdf", pagesize=A4)

elements = []

from datetime import datetime, timedelta

# Datum festlegen
ziel_datum = datetime.today()

# Montag berechnen
montag = ziel_datum - timedelta(days=ziel_datum.weekday())

# Sonntag berechnen
sonntag = montag + timedelta(days=6)

kw = montag.isocalendar().week

print("Zeitraum:", montag, "bis", sonntag)

df["DATUM"] = pd.to_datetime(df["DATUM"])

df = df[(df["DATUM"] >= montag) & (df["DATUM"] <= sonntag)]

# Titel
titel = f"⚽ Spielplan KW {kw} ({montag.strftime('%d.%m')} - {sonntag.strftime('%d.%m.%Y')})"
elements.append(Paragraph(titel, styles["Title"]))
elements.append(Spacer(1, 10))
def draw_background(canvas, doc):
    canvas.saveState()

    logo = ImageReader("asv_logo.png")

    # Größe (groß für Hintergrund)
    width, height = A4

    canvas.drawImage(
        logo,
        x=width/2 - 6*cm,
        y=height/2 - 6*cm,
        width=12*cm,
        height=12*cm,
        mask='auto',
        preserveAspectRatio=True,
        anchor='c'
    )

    canvas.restoreState()
#logo = Image("asv_logo.png")

# Größe anpassen (wichtig!)
#logo.drawHeight = 4*cm
#logo.drawWidth = 5*cm
#logo.hAlign = "CENTER"

#elements.append(logo)
#elements.append(Spacer(1, 10))  

# Sortieren
df = df.sort_values(by="DATUM")

# Spiele durchgehen
for _, row in df.iterrows():
    typ = str(row.get("TYP", "")).strip().lower()
    liga = str(row.get("LIGA", ""))
    gegner = str(row.get("GEGNER", ""))
    datum = row["DATUM"].strftime("%d.%m.%Y")
    zeit = str(row.get("STARTZEIT"))
    ort = str(row.get("ORT", ""))
    status = str(row.get("STATUS", "")).strip().lower()

    # Heim/Auswärts
    if typ == "heim":
        spiel = f"🏠 {liga}: ASV Neufeld vs {gegner}"
    else:
        spiel = f"🚗 {liga}: {gegner} vs ASV Neufeld"

    # Freundschaft erkennen (optional)
    beschreibung = str(row.get("BESCHREIBUNG", "")).lower()
    if "freundschaft" in beschreibung:
        spiel = f"🤝 Freundschaftsspiel: ASV Neufeld vs {gegner}"

    # Status
    if status == "abgesagt":
        spiel = f"❌ ABGESAGT: {spiel}"

    # Block bauen
    
    elements.append(Paragraph(spiel, styles["Heading3"]))
    elements.append(Paragraph(f"📅 {datum} | ⏰ {zeit}", styles["Normal"]))
    elements.append(Paragraph(f"📍 {ort}", styles["Normal"]))
    elements.append(Spacer(1, 12))

# PDF erzeugen
#kw = montag.isocalendar().week
from datetime import datetime
timestamp = datetime.now().strftime("%Y%m%d_%H%M")
filename = f"Spielplan_KW{kw}_{timestamp}.pdf"
#doc = SimpleDocTemplate(filename, pagesize=landscape(A4))
doc = SimpleDocTemplate(
    filename,
    pagesize=landscape(A4),
    leftMargin=1*cm,
    rightMargin=1*cm,
    topMargin=1*cm,
    bottomMargin=1*cm
)
doc.build(elements, onFirstPage=draw_background, onLaterPages=draw_background)
#doc.build(elements)


print("✅ Design PDF erstellt")