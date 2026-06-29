# SAP AgriTrade Automation

> ⚠️ **Private Repository** — Contains production-grade SAP scripting tied to live business infrastructure. All company-specific data (T-codes, user IDs, file paths, partner names) has been anonymized in the public-facing scripts in [`Automation-Portfolio`](https://github.com/NandesHungarian/Automation-Portfolio).

---

## Overview

A Python-based end-to-end automation system for **daily and weekly sales reporting in agricultural commodity trading**, built around SAP ERP integration. The system replaces a manual 45–90 minute reporting workflow with a single script execution.

**Domain:** Agricultural commodity trading (grains, oilseeds) — sales and logistics operations
**SAP modules involved:** SD (Sales & Distribution), custom reporting transactions
**Environment:** Windows · SAP GUI with Scripting API · Microsoft Excel · Outlook

---

## System Architecture

```
┌─────────────────────────────────────────────────────────┐
│                    sap_sales_automation_.py              │
│                                                         │
│  1. SAP Login ──► 2. Run Transaction ──► 3. Export XLS  │
│                                                         │
│  4. Excel Processing:                                   │
│     ├─ Delete internal/purchase rows                    │
│     ├─ Load daily FX rate files (EUR/HUF, USD/HUF)      │
│     ├─ Detect missing freight → Tkinter popup           │
│     ├─ Calculate: Flat EUR, Freight EUR, Net Flat EUR    │
│     ├─ Reorder columns to fixed management layout       │
│     ├─ Apply conditional formatting (anomaly detection) │
│     └─ Draw pivot summary table (by location/commodity) │
│                                                         │
│  5. Save .xlsx ──► 6. Optional: Generate HTML map       │
└─────────────────────────────────────────────────────────┘
         │
         ▼
┌─────────────────────────────────────────────────────────┐
│                    map_generator_.py                     │
│                                                         │
│  Reads processed .xlsx ──► Extracts shipment locations  │
│  ──► Folium/Leaflet map ──► Self-contained .html output │
└─────────────────────────────────────────────────────────┘
```

---

## Scripts

### `sap_sales_automation_.py`

The core automation script. Handles the full pipeline from SAP login to finished management report.

**Key features:**

| Feature | Detail |
|---------|--------|
| SAP GUI Scripting | Automates login, transaction navigation, and ALV report export via `win32com` |
| Smart date logic | Runs daily report on weekdays, weekly report on Mondays automatically |
| FX rate lookup | Loads daily EUR/HUF and USD/HUF rate files, matches by closest prior date |
| Freight detection | Flags non-FCA contracts with missing freight costs, prompts user via Tkinter GUI, caches per contract base number |
| Column normalization | Maps columns by header name (not fixed index) — robust against SAP layout changes |
| Anomaly detection | Marks rows red: HUF prices below threshold, internal partners, missing freight on DDP/CPT contracts |
| Summary table | Draws a pivot-style table at the bottom of the sheet: quantity and weighted avg net price by location, commodity, and crop year |
| Map integration | After saving the report, optionally launches `map_generator_.py` |

**SAP interaction flow:**
1. Checks if SAP GUI is already running; launches `saplogon.exe` if not
2. Connects to the configured ERP system via `win32com.client.GetObject("SAPGUI")`
3. Enters credentials from `~/sap_config.txt`
4. Navigates to the custom sales transaction via T-code
5. Sets plant entity, date range, and executes
6. Exports via ALV toolbar → Excel → catches the new workbook

**Currency conversion logic:**
- `Flat Price` in HUF → divide by EUR/HUF rate for that doc date
- `Flat Price` in USD → multiply by USD/HUF, then divide by EUR/HUF
- Freight costs follow the same logic per their own currency column
- `Net Flat EUR = Flat EUR − Freight EUR`
- FX rates are loaded from daily price indication Excel files stored in a shared folder

---

### `map_generator_.py`

Standalone interactive HTML logistics map generator.

**Key features:**
- Reads the finished management report `.xlsx`
- Geocodes base locations using a predefined coordinate dictionary (no API key required)
- Plots shipment markers on a Folium/Leaflet.js map
- Color-codes by commodity type (e.g. rapeseed vs. sunflower)
- Adds popup info per marker: partner, quantity, net price, IncoTerm
- Exports a fully self-contained `.html` file — no server, no dependencies

---

## Tech Stack

| Tool | Purpose |
|------|---------|
| `win32com.client` | SAP GUI Scripting API, Excel COM automation |
| `tkinter` + `ttk` | Freight input popup GUI |
| `openpyxl` | Excel cell-level formatting, formula writing |
| `pandas` | Data loading and analysis in `map_generator_.py` |
| `folium` | Interactive Leaflet.js map generation |
| `datetime`, `glob`, `shutil` | File handling, date logic, temp file cleanup |

---

## Setup

```bash
pip install -r requirements.txt
```

On first run, `sap_sales_automation_.py` creates `~/sap_config.txt`:

```
USERNAME=Your_SAP_Username
PASSWORD=Your_SAP_Password
```

Fill in your credentials and re-run. SAP GUI must be installed with the Scripting API enabled (SAP Logon → Options → Scripting → Enable scripting).

---

## Security Notes

- Credentials are stored in a local plaintext file — **never committed to git** (`.gitignore` covers this)
- All T-codes, plant codes, partner names, and file paths in the public version are anonymized
- The private version of this repo contains the full production configuration

---

## Output

**Daily report (`daily_report_YYYY-MM-DD.xlsx`):**
- Filtered to active sales team members
- Columns reordered to fixed management layout
- EUR net prices calculated
- Anomaly rows highlighted in red
- Daily summary table appended

**Weekly report (`weekly_report_YYYY-MM-DD.xlsx`):**
- Same as daily, plus a separate Friday-only sales breakdown table

**Optional: `logistics_map_YYYY-MM-DD.html`**
- Interactive map with all active shipment locations
- Opens directly in any browser
