# smgw-cms-data-extractor

**EN: Python script to convert German Smart Meter Gateway (aka SMGW/iMSys) cms-files into structured Excel and CSV files with daily values.**

---

### What this script does

This script reads a CMS-wrapped SMGW export file such as:

`meter_id.sm_data.xml.cms`

It extracts the embedded XML meter data, parses the cumulative import meter readings, and generates:

- an **Excel workbook** (`.xlsx`)
- a **CSV file** (`.csv`)

The Excel output contains:

- **Daily end values**
- **Tariff zone calculations** for **Intelligent Octopus Go**
  - **Go time:** 00:00 to 05:00
  - **Standard time:** 05:00 to 24:00

The script is designed for German smart meter HAN exports where timestamps are stored in **UTC** and meter values are provided as cumulative counters.

### Definition of the daily end value

The script defines the **daily end value** as:

> the **first available cumulative meter reading at 00:00 local time of the following day**, assigned to the previous day.

Example:

- Reading at `2026-03-21 00:00:03` local time
- This is treated as the **daily end value for 2026-03-20**

Why this method is used:

A reading at `23:45` does **not** yet include the last 15 minutes before midnight.  
The first reading at `00:00` of the next day is the cleanest end-of-day value for a cumulative meter register.

### Tariff zone logic

For the tariff sheet, the script assumes a **calendar day** from `00:00` to `24:00`.

For each day it uses these three cumulative import meter readings:

- `00:00` local time
- `05:00` local time
- `00:00` local time of the following day

From these values it calculates:

- **Octopus Go consumption (00:00–05:00)**  
  `reading at 05:00 - reading at 00:00`

- ****Octopus Standard consumption (05:00–24:00)**  
  `reading at next day's 00:00 - reading at 05:00`

- **Total daily consumption (00:00–24:00)**  
  `reading at next day's 00:00 - reading at 00:00`

So the total daily consumption is mathematically identical to:

`Go consumption + Standard consumption`

but it is calculated directly from the two daily boundary values.

---

## Requirements / Voraussetzungen

### Operating system

Recommended:

- **Windows 10 or Windows 11**

The script should also run on Linux or macOS, but the instructions below are written for Windows.

### Python

Required:

- **Python 3.10 or newer**

You also need the Python package:

- `openpyxl`

Install it with:

```bash
pip install openpyxl
```

The script contains a built-in fallback for **Europe/Berlin** timezone handling if `tzdata` is not installed.  
Installing `tzdata` is optional, but can still be useful on some systems:

```bash
pip install tzdata
```

---

## Input file

The script expects a CMS export file from the smart meter, for example:

```text
Zählerstände-2026-01-01 bis_2026-03-21_1lgz0072999211.sm_data.xml.cms
```

---

## Usage / Verwendung

### Basic command

```bash
python smgw_cms_data_extractor.py "Zählerstände-2026-01-01 bis_2026-03-21_1lgz0072999211.sm_data.xml.cms"
```

### With explicit output file name

```bash
python smgw_cms_data_extractor.py "Zählerstände-2026-01-01 bis_2026-03-21_1lgz0072999211.sm_data.xml.cms" -o "output.xlsx"
```

---

## Output files

The script creates:

- an Excel file (`.xlsx`)
- a CSV file (`.csv`)

### Excel structure

The Excel workbook contains at least these sheets:

#### 1. `Tagesendwerte`

Contains the daily end values.

Typical columns:

- Date / Datum
- Used local timestamp
- Import reading (1.8.0) in kWh
- Export reading (2.8.0) in kWh

#### 2. `Tarifzonen`

Contains the tariff-relevant import readings and calculated daily consumption split.

Typical columns:

- Date / Datum
- Local timestamp 00:00
- Local timestamp 05:00
- Local timestamp next day 00:00
- Import reading at 00:00
- Import reading at 05:00
- Import reading at next day 00:00
- Go consumption 00:00–05:00
- Standard consumption 05:00–24:00
- Total daily consumption 00:00–24:00

---

## Example workflow on Windows

1. Install Python
2. Install `openpyxl`
3. Save the script as `smgw_cms_data_extractor.py`
4. Copy your `.sm_data.xml.cms` file into the same folder
5. Open Command Prompt in that folder
6. Run:

```bash
python smgw_cms_data_extractor.py "your_file.sm_data.xml.cms"
```

7. Open the generated Excel file in Excel

---

## Notes

- The script is intended for **German SMGW CMS export files**
- It currently focuses on **cumulative import readings** for daily and tariff calculations
- The tariff logic is tailored to **Intelligent Octopus Go (00:00–05:00)**

---

## License

MIT
