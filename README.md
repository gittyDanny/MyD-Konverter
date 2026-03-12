[README.md](https://github.com/user-attachments/files/25952360/README.md)
# 🔑 MyD Konverter

**XMLs und Migrationsprotokolle hochladen und das Leben genießen.**

Ein Desktop-Tool zur Analyse und Konvertierung von SAP MyDesign (MyD) Migrationsdaten – von XML- und TXT-Dateien zu übersichtlichen Excel-Reports.

---

## ✨ Features

- 📂 **XML-Parsing** – Liest MyD-Migrationsdateien (Produktstamm, Preise, Serien, Langtexte etc.)
- 📄 **TXT-Parsing** – Verarbeitet SAP-Migrationsprotokolle und extrahiert Fehler/Warnungen
- 📊 **Excel-Export** – Generiert formatierte Excel-Dateien mit mehreren Sheets, Filtern und Auto-Spaltenbreite
- 🖥️ **GUI** – Einfache Drag & Drop Oberfläche (oder Ordner-Auswahl per Dialog)
- 🎨 **PwC-Design** – Professionelles Erscheinungsbild mit Corporate Theme

---

## 🚀 Installation

```bash
# Repository klonen
git clone https://github.com/gittyDanny/MyD-Konverter.git
cd MyD-Konverter

# Abhängigkeiten installieren
pip install -r requirements.txt

# Starten
python main.py
```

---

## 📁 Projektstruktur

```
MyD_Konverter/
├── main.py              # Einstiegspunkt
├── config/
│   └── theme.py         # Farben & Styling
├── core/
│   ├── xml_parser.py    # XML-Dateien parsen
│   ├── txt_parser.py    # Migrationsprotokolle parsen
│   └── excel_writer.py  # Excel-Export
├── gui/
│   └── app.py           # GUI (CustomTkinter)
├── key_icon.ico         # App-Icon
└── requirements.txt     # Python-Abhängigkeiten
```

---

## 🛠️ Technologien

- **Python 3.11+**
- **CustomTkinter** – Moderne GUI
- **OpenPyXL** – Excel-Erzeugung
- **lxml** – XML-Parsing

---

## 📸 Screenshot

<!-- Screenshot hier einfügen -->
<!-- ![MyD Konverter](screenshot.png) -->

---

## 👤 Autor

**Daniil Ioffe**

---

## 📝 Lizenz

Dieses Projekt ist frei verfügbar – keine Einschränkungen.
