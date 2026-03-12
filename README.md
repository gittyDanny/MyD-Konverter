# 🔑 MyD Konverter

**XMLs und Migrationsprotokolle hochladen und das Leben genießen.**

Ein Desktop-Tool zur Analyse und Konvertierung von SAP Migrate your Data (MyD) Migrationsdaten – von XML- und TXT-Dateien zu übersichtlichen Excel-Reports.

---

## ✨ Features

- 📂 **XML-Parsing** – Liest MyD-Migrationsdateien (Produktstamm, Preise, Serien, Langtexte etc.)
- 📄 **TXT-Parsing** – Verarbeitet SAP-Migrationsprotokolle und extrahiert Fehler/Warnungen
- 📊 **Excel-Export** – Generiert formatierte Excel-Dateien mit mehreren Sheets, Filtern und Auto-Spaltenbreite
- 🖥️ **GUI** – Einfache Benutzeroberfläche (oder Ordner-Auswahl per Dialog)

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

<img width="1920" height="1200" alt="image" src="https://github.com/user-attachments/assets/0b649af8-02cd-47a8-8bdd-7400d8a8f65d" />


---

## 👤 Autor

**Daniil Ioffe**

---

## 📝 Lizenz

Dieses Projekt ist frei verfügbar – keine Einschränkungen.
