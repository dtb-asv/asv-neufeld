from PIL import Image, ImageDraw, ImageFont
import pandas as pd
from datetime import datetime
from pathlib import Path
import os


# =========================================================
# GRUNDEINSTELLUNGEN
# =========================================================

# Ordner, in dem dieses Python-Programm liegt
BASE_DIR = Path(__file__).resolve().parent

# Excel-Datei
EXCEL_DATEI = BASE_DIR / "Excel2" / "Terminefussball_2026_HerbstV2.0.xlsx"

# Tabellenblatt
SHEET = "ICS2"

# Hintergrundbild
BACKGROUND_PATH = BASE_DIR / "background.jpg"

# Ausgabeordner
OUTPUT_DIR = BASE_DIR / "output"


# =========================================================
# AUSWAHL
# =========================================================

# "single" = nur eine bestimmte Woche
# "alle"   = alle vorhandenen Wochen erstellen

MODUS = "single"


# Bei MODUS = "single":
# Einfach ein Datum aus der gewünschten Woche eingeben

ZIEL_DATUM_INPUT = "2026-09-07"


# =========================================================
# HILFSFUNKTIONEN
# =========================================================

def zeit_formatieren(zeit):

    if pd.isna(zeit):
        return ""

    # Falls Zeit bereits datetime.time ist
    if hasattr(zeit, "strftime"):
        try:
            return zeit.strftime("%H:%M")
        except:
            pass

    zeit_string = str(zeit)

    # Entfernt Sekunden, falls vorhanden
    if len(zeit_string) >= 5:
        return zeit_string[:5]

    return zeit_string


def font_laden(groesse):

    # Windows Arial
    try:
        return ImageFont.truetype("arial.ttf", groesse)

    except:
        # Fallback
        return ImageFont.load_default()


# =========================================================
# EXCEL PRÜFEN
# =========================================================

print("")
print("==============================================")
print("        ASV NEUFELD SPIELPLAN GENERATOR")
print("==============================================")
print("")


if not EXCEL_DATEI.exists():

    print("❌ FEHLER!")
    print("")
    print("Excel-Datei wurde nicht gefunden:")
    print(EXCEL_DATEI)
    print("")

    input("Enter drücken zum Beenden...")

    raise SystemExit


if not BACKGROUND_PATH.exists():

    print("❌ FEHLER!")
    print("")
    print("background.jpg wurde nicht gefunden:")
    print(BACKGROUND_PATH)
    print("")

    input("Enter drücken zum Beenden...")

    raise SystemExit


print("✅ Excel-Datei gefunden")
print("📁", EXCEL_DATEI)
print("")


# =========================================================
# EXCEL LADEN
# =========================================================

try:

    df = pd.read_excel(
        EXCEL_DATEI,
        sheet_name=SHEET
    )

except Exception as e:

    print("❌ Fehler beim Laden der Excel-Datei!")
    print(e)

    input("Enter drücken zum Beenden...")

    raise SystemExit


print("✅ Excel erfolgreich geladen")


# =========================================================
# DATUM KONVERTIEREN
# =========================================================

df["DATUM"] = pd.to_datetime(
    df["DATUM"],
    errors="coerce"
)


# Ungültige Datumswerte entfernen
df = df.dropna(
    subset=["DATUM"]
)


# =========================================================
# NUR SPIELE
# =========================================================

if "ART" in df.columns:

    df = df[
        df["ART"]
        .astype(str)
        .str.lower()
        .eq("spiel")
    ]


# =========================================================
# KM UND U23 ENTFERNEN
# =========================================================

df = df[
    ~df["LIGA"]
    .astype(str)
    .str.contains(
        "KM|U23",
        case=False,
        na=False
    )
]


# =========================================================
# ABGESAGTE SPIELE ENTFERNEN
# =========================================================

if "STATUS" in df.columns:

    df = df[
        ~df["STATUS"]
        .astype(str)
        .str.contains(
            "abgesagt",
            case=False,
            na=False
        )
    ]


# =========================================================
# KALENDERWOCHE BERECHNEN
# =========================================================

df["KW"] = (
    df["DATUM"]
    .dt
    .isocalendar()
    .week
    .astype(int)
)


df["JAHR"] = (
    df["DATUM"]
    .dt
    .year
)


# =========================================================
# WOCHEN AUSWÄHLEN
# =========================================================

if MODUS.lower() == "single":

    try:

        ziel_datum = datetime.strptime(
            ZIEL_DATUM_INPUT,
            "%Y-%m-%d"
        )

    except:

        print("❌ Datum falsch eingegeben!")
        print("")
        print("Richtig wäre zum Beispiel:")
        print("2026-09-07")

        input("Enter drücken zum Beenden...")

        raise SystemExit


    ziel_kw = (
        ziel_datum
        .isocalendar()
        .week
    )


    ziel_jahr = (
        ziel_datum
        .year
    )


    wochen = [
        (ziel_jahr, ziel_kw)
    ]


elif MODUS.lower() == "alle":

    wochen = (
        df[
            ["JAHR", "KW"]
        ]
        .drop_duplicates()
        .sort_values(
            ["JAHR", "KW"]
        )
        .values
        .tolist()
    )


else:

    print("❌ MODUS falsch!")
    print("")
    print("Erlaubt:")
    print("single")
    print("alle")

    input("Enter drücken zum Beenden...")

    raise SystemExit


# =========================================================
# FONTS
# =========================================================

font_title = font_laden(60)

font_big = font_laden(45)

font_small = font_laden(28)

font_kw = font_laden(38)


# =========================================================
# LAYOUT
# =========================================================

BILD_BREITE = 1080

BILD_HOEHE = 1350


Y_START = 220

BLOCK_HOEHE = 130

MAX_SPIELE = 7


# =========================================================
# NEUES BILD
# =========================================================

def neues_bild():

    img = Image.open(
        BACKGROUND_PATH
    ).convert("RGBA")


    img = img.resize(
        (
            BILD_BREITE,
            BILD_HOEHE
        )
    )


    overlay = Image.new(

        "RGBA",

        img.size,

        (
            0,
            0,
            0,
            120
        )
    )


    return Image.alpha_composite(
        img,
        overlay
    )


# =========================================================
# HAUPTPROGRAMM
# =========================================================

print("")
print("==============================================")
print("        SPIELPLÄNE WERDEN ERSTELLT")
print("==============================================")
print("")


gesamt_anzahl = 0


for jahr, kw in wochen:


    print("")
    print("----------------------------------------------")
    print(
        f"📅 Bearbeite KW {kw} / {jahr}"
    )
    print("----------------------------------------------")


    # Spiele der Woche auswählen

    df_woche = df[

        (df["KW"] == int(kw))

        &

        (df["JAHR"] == int(jahr))

    ].sort_values(

        by=["DATUM", "STARTZEIT"]

    )


    if df_woche.empty:

        print(
            f"⚠️ Keine Spiele in KW {kw}"
        )

        continue


    # Ausgabeordner

    folder = (

        OUTPUT_DIR

        / str(jahr)

        / f"KW{kw}"

    )


    folder.mkdir(

        parents=True,

        exist_ok=True

    )


    print(
        f"⚽ {len(df_woche)} Spiele gefunden"
    )


    # =====================================================
    # SEITEN ERSTELLEN
    # =====================================================

    seiten = []

    count = 0

    seite = 1


    img = neues_bild()

    draw = ImageDraw.Draw(img)


    # HEADER

    draw.text(

        (50, 50),

        "SPIELPLAN NACHWUCHS",

        font=font_title,

        fill="white"

    )


    draw.text(

        (50, 120),

        f"KW {kw}",

        font=font_kw,

        fill=(200, 200, 200)

    )


    y = Y_START


    # =====================================================
    # SPIELE DURCHGEHEN
    # =====================================================

    for _, row in df_woche.iterrows():


        # Neue Seite nach 7 Spielen

        if count == MAX_SPIELE:


            seiten.append(

                img.convert("RGB")

            )


            seite += 1

            count = 0


            img = neues_bild()

            draw = ImageDraw.Draw(img)


            # HEADER

            draw.text(

                (50, 50),

                "SPIELPLAN NACHWUCHS",

                font=font_title,

                fill="white"

            )


            draw.text(

                (50, 120),

                f"KW {kw}",

                font=font_kw,

                fill=(200, 200, 200)

            )


            y = Y_START


        # =================================================
        # DATEN
        # =================================================

        liga = str(

            row.get(
                "LIGA",
                ""
            )

        )


        gegner = str(

            row.get(
                "GEGNER",
                ""
            )

        )


        datum = (

            row["DATUM"]

            .strftime(

                "%d.%m.%Y"

            )

        )


        zeit = zeit_formatieren(

            row.get(
                "STARTZEIT",
                ""
            )

        )


        ort = str(

            row.get(
                "ORT",
                ""
            )

        )


        typ = (

            str(

                row.get(
                    "TYP",
                    ""
                )

            )

            .strip()

            .lower()

        )


        # =================================================
        # HEIM / AUSWÄRTS
        # =================================================

        if typ == "heim":


            label = "Heim"


            spiel = (

                f"ASV Neufeld vs {gegner}"

            )


        else:


            label = "Auswärts"


            spiel = (

                f"{gegner} vs ASV Neufeld"

            )


        # =================================================
        # ZEICHNEN
        # =================================================

        draw.text(

            (50, y),

            liga,

            font=font_big,

            fill="white"

        )


        text = (

            f"{datum} | "

            f"{zeit} Uhr | "

            f"{label} "

            f"{ort}"

        )


        draw.text(

            (220, y + 10),

            text,

            font=font_small,

            fill=(220, 220, 220)

        )


        draw.text(

            (220, y + 55),

            spiel.upper(),

            font=font_small,

            fill="white"

        )


        draw.line(

            (

                50,

                y + 110,

                1000,

                y + 110

            ),

            fill=(

                255,

                255,

                255,

                60

            ),

            width=1

        )


        y += BLOCK_HOEHE

        count += 1


    # =====================================================
    # LETZTE SEITE HINZUFÜGEN
    # =====================================================

    seiten.append(

        img.convert("RGB")

    )


    # =====================================================
    # PDF SPEICHERN
    # =====================================================

    pdf_datei = (

        folder

        / f"Spielplan_KW{kw}_{jahr}.pdf"

    )


    erste_seite = seiten[0]


    weitere_seiten = seiten[1:]


    erste_seite.save(

        pdf_datei,

        "PDF",

        resolution=100.0,

        save_all=True,

        append_images=weitere_seiten

    )


    print("✅ PDF erstellt:")

    print(pdf_datei)


    gesamt_anzahl += 1


# =========================================================
# FERTIG
# =========================================================

print("")
print("==============================================")
print("🚀 FERTIG!")
print("==============================================")

print("")
print(
    f"📄 {gesamt_anzahl} Spielplan-PDF(s) erstellt"
)

print("")
print(
    "Ausgabeordner:"
)

print(
    OUTPUT_DIR
)

print("")


input(
    "Enter drücken zum Beenden..."
)