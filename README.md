# smgw-cms-data-extractor

**DE: Python-Skript zur Umwandlung von CMS-Dateien eines Smart Meter Gateways (SMGW/iMSys) in strukturierte Excel- und CSV-Dateien mit Tageswerten**

**EN: Python script to convert German Smart Meter Gateway (aka SMGW/iMSys) cms-files into structured Excel and CSV files with daily values.**

---

### Was das Skript macht

Dieses Skript liest eine CMS-verpackte SMGW-Exportdatei wie zum Beispiel:

`zaehler_id.sm_data.xml.cms`

Es extrahiert die darin eingebetteten XML-Messdaten, verarbeitet die kumulativen Bezugszählerstände und erzeugt daraus:

- eine **Excel-Datei** (`.xlsx`)
- eine **CSV-Datei** (`.csv`)

Die Excel-Ausgabe enthält:

- **Tagesendwerte**
- **Tarifzonen-Berechnungen** für **Intelligent Octopus Go**
  - **Go-Zeit:** 00:00 bis 05:00 Uhr
  - **Standardzeit:** 05:00 bis 24:00 Uhr

Das Skript ist für deutsche Smart-Meter-HAN-Exporte gedacht, bei denen die Zeitstempel in **UTC** vorliegen und die Zählerstände als kumulative Registerwerte geliefert werden.

### Definition des Tagesendwerts

Das Skript definiert den **Tagesendwert** als:

> den **ersten verfügbaren kumulativen Zählerstand um 00:00 Uhr lokaler Zeit des Folgetags**, der dem Vortag zugeordnet wird.

Beispiel:

- Messung um `2026-03-21 00:00:03` lokaler Zeit
- Dieser Wert gilt als **Tagesendwert für den 20.03.2026**

Warum diese Methode sinnvoll ist:

Ein Messwert von `23:45` enthält die letzten 15 Minuten vor Mitternacht noch **nicht**.  
Der erste Wert um `00:00` des Folgetags ist daher der fachlich sauberste Tagesendwert für ein kumulatives Zählerregister.

### Logik der Tarifzonen

Für das Tarifblatt geht das Skript von einem **Kalendertag** von `00:00` bis `24:00` aus.

Pro Tag werden diese drei kumulativen Bezugszählerstände verwendet:

- `00:00` lokale Zeit
- `05:00` lokale Zeit
- `00:00` lokale Zeit des Folgetags

Daraus werden berechnet:

- **Octopus Go-Verbrauch (00:00–05:00)**  
  `Zählerstand 05:00 - Zählerstand 00:00`

- **Octopus Standard-Verbrauch (05:00–24:00)**  
  `Zählerstand Folgetag 00:00 - Zählerstand 05:00`

- **Gesamtverbrauch des Tages (00:00–24:00)**  
  `Zählerstand Folgetag 00:00 - Zählerstand 00:00`

Der Gesamtverbrauch ist damit rechnerisch identisch zu:

`Go-Verbrauch + Standard-Verbrauch`

wird aber direkt aus den beiden Tagesrandwerten berechnet.

## Voraussetzungen

### Betriebssystem

Empfohlen:

- **Windows 10 oder Windows 11**

Das Skript sollte auch unter Linux oder macOS laufen, aber die nachfolgenden Hinweise beziehen sich auf Windows.

### Python

Erforderlich:

- **Python 3.10 oder neuer**

Außerdem wird folgendes Python-Paket benötigt:

- `openpyxl`

Installation:

```bash
pip install openpyxl
```

Das Skript enthält eine eingebaute Fallback-Logik für die Zeitzone **Europe/Berlin**, falls `tzdata` nicht installiert ist.  
Die Installation von `tzdata` ist optional, kann aber auf manchen Systemen trotzdem sinnvoll sein:

```bash
pip install tzdata
```

---

## Eingabedatei

Das Skript erwartet eine CMS-Exportdatei des Smart Meters, zum Beispiel:

```text
Zählerstände-2026-01-01 bis_2026-03-21_1lgz0072999211.sm_data.xml.cms
```

---

## Verwendung

### Grundlegender Aufruf

```bash
python smgw_cms_data_extractor.py "Zählerstände-2026-01-01 bis_2026-03-21_1lgz0072999211.sm_data.xml.cms"
```

### Mit explizitem Ausgabedateinamen

```bash
python smgw_cms_data_extractor.py "Zählerstände-2026-01-01 bis_2026-03-21_1lgz0072999211.sm_data.xml.cms" -o "output.xlsx"
```

---

## Ausgabedateien

Das Skript erzeugt:

- eine Excel-Datei (`.xlsx`)
- eine CSV-Datei (`.csv`)

### Aufbau der Excel-Datei

Die Excel-Arbeitsmappe enthält mindestens diese Tabellenblätter:

#### `Tagesendwerte`

Typische Spalten:

- Datum
- Verwendeter lokaler Zeitstempel
- Bezugszählerstand (1.8.0) in kWh
- Einspeisezählerstand (2.8.0) in kWh

#### `Tarifzonen`

Typische Spalten:

- Datum
- Lokaler Zeitstempel 00:00
- Lokaler Zeitstempel 05:00
- Lokaler Zeitstempel Folgetag 00:00
- Bezugszählerstand um 00:00
- Bezugszählerstand um 05:00
- Bezugszählerstand um Folgetag 00:00
- Octopus Go-Verbrauch 00:00–05:00
- Octopus Standard-Verbrauch 05:00–24:00
- Gesamtverbrauch 00:00–24:00

---

## Beispielhafter Ablauf unter Windows

1. Python installieren
2. `openpyxl` installieren
3. Das Skript als `smgw_cms_data_extractor.py` speichern
4. Die `.sm_data.xml.cms`-Datei in denselben Ordner kopieren
5. In diesem Ordner die Eingabeaufforderung öffnen
6. Folgenden Befehl ausführen:

```bash
python smgw_cms_data_extractor.py "your_file.sm_data.xml.cms"
```

7. Die erzeugte Excel-Datei in Excel öffnen

---

## Hinweise

- Das Skript ist für **deutsche SMGW-CMS-Exportdateien** gedacht
- Der Schwerpunkt liegt derzeit auf **kumulativen Bezugszählerständen** für Tages- und Tarifberechnungen
- Die Tariflogik ist auf **Intelligent Octopus Go (00:00–05:00)** zugeschnitten

---

## Lizenz

MIT
