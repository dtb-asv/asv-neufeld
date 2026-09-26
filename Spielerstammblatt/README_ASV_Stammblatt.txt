ASV NEUFELD – SPIELERSTAMMBLATT IMPORTER (Version 1)
=====================================================

FUNKTION
- Liest PDF, JPG, JPEG und PNG.
- Digitale PDFs werden direkt ausgelesen.
- Eingescannte PDFs/Bilder werden mit OCR verarbeitet.
- Angehaengte DSGVO-Seiten werden ignoriert, wenn sie kein Spielerstammblatt sind.
- Mehrere Stammblatt-Seiten in einer PDF werden unterstuetzt.
- Mannschaft (z.B. U10) wird beim Import mitgegeben.
- Alle erkannten Werte koennen vor dem Excel-Export kontrolliert/korrigiert werden.
- Checkboxen werden absichtlich NICHT uebernommen.

BENÖTIGT (Windows)
1. Python 3 muss installiert sein.
2. In Eingabeaufforderung/PowerShell ausfuehren:
   pip install pymupdf pytesseract pillow openpyxl
3. Fuer Scans/JPG wird Tesseract OCR benoetigt.
   Installationsordner normalerweise:
   C:\Program Files\Tesseract-OCR\
   Das Programm erkennt diesen Pfad automatisch.
   Fuer bessere deutsche Erkennung sollte das deutsche Sprachpaket (deu) installiert sein.

START
- Doppelklick auf START_ASV_Stammblatt.bat
  oder
- python asv_stammblatt_importer.py

ABLAUF
1. Mannschaft waehlen, z.B. U10.
2. "Stammblätter auswählen" oder "Ordner auswählen" anklicken.
3. PDF/JPG/PNG-Dateien auswaehlen.
4. Original links mit den erkannten Daten rechts vergleichen.
5. Fehler direkt in den Feldern korrigieren.
6. Mit "Naechster" alle Spieler kontrollieren.
7. "Excel exportieren" anklicken.

WICHTIG
Handschriftliche Texterkennung ist nie 100 % sicher. Deshalb ist die Kontrollansicht fester Bestandteil des Programms.
