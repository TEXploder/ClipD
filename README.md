# ClipD — Smart Clipboard Guardian

ClipD ist ein eleganter Zwischenablagen-Helfer für Windows.  
Er speichert automatisch deine Clipboard-Historie, zeigt Vorschauen direkt am Mauszeiger an und lässt dich mit globalen Hotkeys blitzschnell durch kopierte Inhalte navigieren.

## Features
- Hintergrunddienst mit Tray-Icon (schließt ins Tray, kein störendes Fenster).
- Verzögerungsfreie Vorschau der Zwischenablage (Hotkeys: `Strg + Alt + Pfeil Hoch/Runter`, `Strg + Alt + V`).
- Verschlüsselte Persistenz (Fernet AES-128) im `%APPDATA%`\ClipD-Ordner.
- Optionaler Bildschirm-Aufnahmeschutz via `SetWindowDisplayAffinity`.
- On-the-fly Skalen-/Dauer-/Farbanpassung für die Overlay-Vorschau.
- Verlauf mit Suche, Datum/Uhrzeit, „Verlauf leeren“-Button und Drag&Drop-fähigen Karten.
- Erster Start installiert sich automatisch in `%ProgramFiles%\TEX-Programme\ClipD` und trägt Autostart & Startmenü-Verknüpfung ein.

## Steuerung

| Aktion | Hotkey/Trigger |
| --- | --- |
| Verlauf öffnen | Tray-Doppelklick oder `Strg + Alt + V` |
| Zwischen Einträgen navigieren | `Strg + Alt + Pfeil hoch/runter` |
| Aktuelle Vorschau einfügen | `Strg + V` |
| Verlaufsfenster | Tray → „Verlauf anzeigen“ |
| ClipD beenden | Tray → „Beenden“ |

## Installation / Build

### 1. Lokale Entwicklung
```bash
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
python clipboard_guardian.py
```

### 2. Onefile Build (Nuitka)
```bash
build_clipd.bat
```

### 3. Onefile Build (PyInstaller + Cython)
```bash
build_clipd_cython.bat
```

**Hinweis:** Beide Batch-Dateien erwarten `logo.ico` und `logo.png` im Projektverzeichnis; Nuitka benötigt wahlweise die Visual Studio Clang-Komponente oder wechselt automatisch auf MinGW64. Der Cython-Build kompiliert `clipboard_guardian.py` zu einem nativen Module (`clipd_core`) und nutzt `run_clipd.py` als Stub.

## Konfiguration & Daten

- `settings.json` liegt unter `%APPDATA%\ClipD`.  
  * `show_preview_overlay`: Vorschau ein-/ausblenden  
  * `capture_protection_enabled`: Bildschirmaufnahme-Schutz  
  * `first_run`: steuert, ob das Welcome-Fenster erscheint  
  * Hotkeys & Farbschema werden ebenfalls dort verwaltet
- Verlaufsspeicher (`history.bin`) liegt verschlüsselt daneben; Schlüssel (`key.bin`) ebenfalls im AppData-Ordner.

## Lizenz / Rechtliches

ClipD verwendet PySide6 (Qt for Python) unter LGPL 3.0.  
Diese README deckt nicht die gesamte Lizenzpflicht ab. Du musst:
1. Den LGPL-Text (und Qt-Lizenzhinweise) beilegen.  
2. Dokumentieren, welche Versionen du mitlieferst.  
3. Ermöglichen, dass Nutzer die Qt-/PySide-Komponenten austauschen (die Onefile-Pakete entpacken zur Laufzeit – Nuter können die entpackten DLLs ersetzen).  

Andere Abhängigkeiten (`cryptography`, `Cython`, Nuitka/PyInstaller etc.) haben meist MIT/BSD/Apache-Lizenzen. Belege ebenfalls deren Lizenztexte.

Für eine echte statische Link-Variante (ohne Ersatzmöglichkeit) benötigst du eine kommerzielle Qt-Lizenz.  

## Credits

- Built with ❤️ by TEXploder.
- PySide6, Qt, Cython, Nuitka, PyInstaller.
- Grafiken: `logo.ico`, `logo.png` (eigenes Branding einfügen).

Viel Spaß mit ClipD! Pull Requests und Issues sind willkommen. 🚀
