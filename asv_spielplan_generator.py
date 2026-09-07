
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from pathlib import Path
from datetime import datetime
import threading
import pandas as pd
from PIL import Image, ImageDraw, ImageFont

BASE_DIR = Path(__file__).resolve().parent
DEFAULT_EXCEL = BASE_DIR / "Excel2" / "Terminefussball_2026_HerbstV2.0.xlsx"
BACKGROUND = BASE_DIR / "background.png"
OUTPUT_DIR = BASE_DIR / "output"
SHEET = "ICS2"

def font(size):
    for name in ["arial.ttf", "C:/Windows/Fonts/arial.ttf"]:
        try:
            return ImageFont.truetype(name, size)
        except:
            pass
    return ImageFont.load_default()

def fmt_time(value):
    if pd.isna(value):
        return ""
    if hasattr(value, "strftime"):
        try:
            return value.strftime("%H:%M")
        except:
            pass
    s = str(value)
    return s[:5] if len(s) >= 5 else s

def new_image():
    if BACKGROUND.exists():
        img = Image.open(BACKGROUND).convert("RGBA").resize((1080, 1350))
        overlay = Image.new("RGBA", img.size, (0, 0, 0, 120))
        return Image.alpha_composite(img, overlay)
    return Image.new("RGBA", (1080, 1350), (25, 25, 25, 255))

def load_games(excel_file):
    df = pd.read_excel(excel_file, sheet_name=SHEET)
    required = ["DATUM", "STARTZEIT", "TYP", "GEGNER", "LIGA"]
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise ValueError("Folgende Spalten fehlen: " + ", ".join(missing))

    df["DATUM"] = pd.to_datetime(df["DATUM"], errors="coerce")
    df = df.dropna(subset=["DATUM"])

    if "ART" in df.columns:
        df = df[df["ART"].astype(str).str.strip().str.lower().eq("spiel")]

    df = df[~df["LIGA"].astype(str).str.contains("KM|U23", case=False, na=False)]

    # SPIELFREI nicht anzeigen
    # Wenn in der Spalte GEGNER "spielfrei" steht, wird die gesamte Zeile entfernt.
    df = df[
        ~df["GEGNER"].astype(str).str.contains(
            "spielfrei",
            case=False,
            na=False
        )
    ]

    if "STATUS" in df.columns:
        df = df[~df["STATUS"].astype(str).str.contains("abgesagt", case=False, na=False)]

    df["KW"] = df["DATUM"].dt.isocalendar().week.astype(int)
    df["JAHR"] = df["DATUM"].dt.year.astype(int)
    return df

def create_week_pdf(df, year, week):
    games = df[(df["JAHR"] == int(year)) & (df["KW"] == int(week))].sort_values(
        by=["DATUM", "STARTZEIT"], kind="stable"
    )
    if games.empty:
        return None

    folder = OUTPUT_DIR / str(year) / f"KW{week:02d}"
    folder.mkdir(parents=True, exist_ok=True)

    title_font = font(60)
    league_font = font(42)
    text_font = font(26)
    kw_font = font(36)

    pages = []
    img = new_image()
    draw = ImageDraw.Draw(img)

    def header():
        draw.text((50, 45), "SPIELPLAN NACHWUCHS", font=title_font, fill="white")
        draw.text((50, 120), f"KW {int(week):02d}", font=kw_font, fill=(220,220,220))

    header()
    y = 220
    count = 0

    for _, row in games.iterrows():
        if count == 7:
            pages.append(img.convert("RGB"))
            img = new_image()
            draw = ImageDraw.Draw(img)
            header()
            y = 220
            count = 0

        liga = str(row.get("LIGA", ""))
        gegner = str(row.get("GEGNER", ""))
        datum = row["DATUM"].strftime("%a, %d.%m.%Y")
        zeit = fmt_time(row.get("STARTZEIT", ""))
        ort = str(row.get("ORT", "")).strip()
        typ = str(row.get("TYP", "")).strip().lower()

        if typ == "heim":
            label = "HEIM"
            spiel = f"ASV NEUFELD vs {gegner.upper()}"
        else:
            label = "AUSWÄRTS"
            spiel = f"{gegner.upper()} vs ASV NEUFELD"

        draw.text((50, y), liga, font=league_font, fill="white")
        info = f"{datum} | {zeit} Uhr | {label}"
        if ort and ort.lower() != "nan":
            info += f" | {ort}"
        draw.text((220, y + 8), info, font=text_font, fill=(220,220,220))
        draw.text((220, y + 58), spiel, font=text_font, fill="white")
        draw.line((50, y + 112, 1030, y + 112), fill=(255,255,255,80), width=2)

        y += 130
        count += 1

    pages.append(img.convert("RGB"))
    pdf = folder / f"Spielplan_KW{int(week):02d}_{year}.pdf"
    pages[0].save(pdf, "PDF", resolution=100.0, save_all=True, append_images=pages[1:])
    return pdf

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("ASV Neufeld – Spielplan Generator")
        self.geometry("720x430")
        self.resizable(False, False)

        self.excel = tk.StringVar(value=str(DEFAULT_EXCEL))
        self.date = tk.StringVar(value=datetime.today().strftime("%Y-%m-%d"))
        self.mode = tk.StringVar(value="single")
        self.status = tk.StringVar(value="Bereit.")

        self.build()

    def build(self):
        pad = {"padx": 20, "pady": 8}

        ttk.Label(self, text="ASV NEUFELD – SPIELPLAN GENERATOR",
                  font=("Arial", 18, "bold")).pack(pady=(22, 18))

        frame = ttk.Frame(self)
        frame.pack(fill="x", **pad)
        ttk.Label(frame, text="Excel-Datei:", width=15).grid(row=0, column=0, sticky="w")
        ttk.Entry(frame, textvariable=self.excel, width=60).grid(row=0, column=1, padx=5)
        ttk.Button(frame, text="Durchsuchen", command=self.choose_excel).grid(row=0, column=2)

        ttk.Separator(self).pack(fill="x", padx=20, pady=10)

        mode_frame = ttk.Frame(self)
        mode_frame.pack(fill="x", **pad)
        ttk.Radiobutton(mode_frame, text="Eine Kalenderwoche erstellen",
                        variable=self.mode, value="single",
                        command=self.toggle_date).grid(row=0, column=0, sticky="w")
        ttk.Radiobutton(mode_frame, text="Alle Kalenderwochen erstellen",
                        variable=self.mode, value="all",
                        command=self.toggle_date).grid(row=1, column=0, sticky="w")

        date_frame = ttk.Frame(self)
        date_frame.pack(fill="x", **pad)
        ttk.Label(date_frame, text="Datum (TT.MM.JJJJ oder JJJJ-MM-TT):").pack(side="left")
        self.date_entry = ttk.Entry(date_frame, textvariable=self.date, width=18)
        self.date_entry.pack(side="left", padx=10)

        self.progress = ttk.Progressbar(self, mode="indeterminate")
        self.progress.pack(fill="x", padx=20, pady=(15, 5))

        ttk.Label(self, textvariable=self.status, wraplength=660).pack(padx=20, pady=8)

        ttk.Button(self, text="PDF(s) ERSTELLEN", command=self.start,
                   width=28).pack(pady=12)

        ttk.Label(self, text="Ausgabe: C:\\Users\\DTB\\asv-neufeld\\output",
                  foreground="gray").pack()

    def choose_excel(self):
        p = filedialog.askopenfilename(
            title="Excel-Datei auswählen",
            filetypes=[("Excel-Dateien", "*.xlsx *.xls")]
        )
        if p:
            self.excel.set(p)

    def toggle_date(self):
        state = "normal" if self.mode.get() == "single" else "disabled"
        self.date_entry.configure(state=state)

    def parse_date(self, value):
        value = value.strip()
        for f in ("%Y-%m-%d", "%d.%m.%Y"):
            try:
                return datetime.strptime(value, f)
            except ValueError:
                pass
        raise ValueError("Datum bitte als 07.09.2026 oder 2026-09-07 eingeben.")

    def start(self):
        if not Path(self.excel.get()).exists():
            messagebox.showerror("Fehler", "Die Excel-Datei wurde nicht gefunden.")
            return

        if self.mode.get() == "single":
            try:
                self.parse_date(self.date.get())
            except ValueError as e:
                messagebox.showerror("Fehler", str(e))
                return

        threading.Thread(target=self.run_generation, daemon=True).start()

    def run_generation(self):
        try:
            self.progress.start(10)
            self.status.set("Excel wird geladen ...")
            df = load_games(Path(self.excel.get()))

            if self.mode.get() == "single":
                d = self.parse_date(self.date.get())
                targets = [(d.isocalendar().year, d.isocalendar().week)]
            else:
                targets = (
                    df[["JAHR", "KW"]]
                    .drop_duplicates()
                    .sort_values(["JAHR", "KW"])
                    .itertuples(index=False, name=None)
                )

            targets = list(targets)
            created = []
            for year, week in targets:
                self.status.set(f"Erstelle KW {int(week):02d}/{year} ...")
                result = create_week_pdf(df, year, week)
                if result:
                    created.append(result)

            self.progress.stop()
            if created:
                self.status.set(f"Fertig! {len(created)} PDF-Datei(en) erstellt.")
                messagebox.showinfo(
                    "Fertig",
                    f"{len(created)} PDF-Datei(en) wurden erstellt.\n\n"
                    f"Ordner:\n{OUTPUT_DIR}"
                )
            else:
                self.status.set("Keine passenden Nachwuchsspiele gefunden.")
                messagebox.showwarning("Keine Spiele", "Es wurden keine passenden Spiele gefunden.")

        except Exception as e:
            self.progress.stop()
            self.status.set("Fehler: " + str(e))
            messagebox.showerror("Fehler", str(e))

if __name__ == "__main__":
    App().mainloop()
