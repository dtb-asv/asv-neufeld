ASV NEUFELD – SPIELERSTAMMBLATT IMPORTER V3
=============================================

V3 arbeitet komplett lokal und benötigt KEINEN OpenAI/API-Key.

NEU IN V3
- PyMuPDF wird über "import pymupdf" verwendet (keine fitz-Warnung mehr).
- Bei Scan-PDF/JPG/PNG wird NICHT mehr die ganze Seite als Datenblock ausgewertet.
- Stattdessen wird jedes bekannte Feld des ASV-Spielerstammblatts separat ausgeschnitten
  und separat per Tesseract OCR gelesen.
- Dadurch können Überschriften und Nachbarzeilen nicht mehr in falsche Excel-Felder rutschen.
- Digitale PDFs werden weiterhin direkt ausgelesen.
- DSGVO-Folgeseiten werden ignoriert.
- Alle Ergebnisse bleiben vor Excel-Export editierbar.

INSTALLATION
1. INSTALL_PYTHON_PAKETE.bat ausführen.
2. Tesseract OCR für Windows muss installiert sein.
3. START_ASV_Stammblatt.bat starten.

WICHTIG
Handschrift-OCR mit Tesseract ist nicht perfekt. V3 verbessert vor allem die korrekte
Feldzuordnung. Handschriftliche Namen/E-Mails/Telefonnummern bitte in der Kontrollansicht
mit dem Original vergleichen und bei Bedarf korrigieren.
