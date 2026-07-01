# SOC Audit Architect

An automation tool designed to streamline Microsoft Sentinel audits. It handles file management, rule decomposition, and report generation so analysts can focus on the actual security analysis.

---

## What It Does

When launched, the tool walks you through an audit session and automatically produces:

| Output | Description |
|---|---|
| **Timestamped session folder** | Isolates every audit engagement (e.g. `2026-02-17_Audit`) |
| **Historical JSON archive** | Moves the raw Sentinel export to a `Historical Master JSON's` folder — preserving an untouched source of truth |
| **Unpacked rules** | Decomposes the master JSON into individual rule files inside `OLD` for granular review and KQL diffing between audits |
| **Interactive Excel report** | Formatted spreadsheet with KQL extraction, colour-coded rows, and `Pass / Fail / In-Progress` dropdown menus per rule |
| **Audit signature** | `Audit_Signature.txt` logging the engineer name, timestamp, client, and rule count for compliance and non-repudiation |

---

## Quick Start

1. Ensure your Sentinel Audits SharePoint folder is synced to OneDrive
2. Run `SOC Audit Architect.exe` (or `python converter_1.1.py` from source)
3. Follow the prompts — select your Sentinel JSON export(s), confirm the client name and date
4. The tool builds the full directory structure and opens the completed audit folder

---

## Installation (from source)

**Requirements:** Python 3.8+, Windows 10/11

```bash
pip install -r requirements.txt
python converter_1.1.py
```

**Dependencies:**
- `pandas` — DataFrame management and data mapping
- `xlsxwriter` — Excel report generation and conditional formatting
- `openpyxl` — Supporting Excel engine
- `tkinter` — Built-in Python GUI (no install needed)

---

## How It Works

### 1. Multi-File Merge
Supports merging multiple Sentinel JSON exports into a single deduplicated ruleset. Duplicate rule names are detected and removed automatically, with a count reported at completion.

### 2. Versioned Session Management
If an audit folder for the same date already exists, the tool automatically increments the version (`_v2`, `_v3`) rather than overwriting existing work.

### 3. Rule Decomposition
Parses the `resources` array in the master JSON and saves each security rule as a standalone `.json` file. Filenames are sanitised to remove characters illegal on Windows file systems.

### 4. Excel Report
Built with `pandas` + `xlsxwriter`:
- 16-column layout covering rule name, version, KQL, severity, enabled/disabled status, and audit fields
- KQL columns rendered in `Consolas` monospace font
- Dropdown menus in the Audit Judgment column with five status options
- Colour-coded rows triggered by conditional formatting based on the selected judgment

### 5. Audit Signature
Writes a plain-text compliance artifact capturing the engineer name, timestamp, client, session version, total unique rules, and duplicates removed.

---

## Directory Structure Created

```
Client Folder/
├── Historical Master JSON's/
│   └── ClientName_YYYY-MM-DD_v1_MASTERJSON.json
└── YYYY-MM-DD_Audit/
    ├── ClientName_Sentinel_Audit_YYYY-MM-DD_v1.xlsx
    ├── Audit_Signature.txt
    ├── OLD/          ← Individual rule JSON files
    ├── NEW/
    ├── hold/
    └── Updated JSON/
```

---

## Demonstration

See **[SOC Audit Architect Demonstration.pdf](./SOC%20Audit%20Architect%20Demonstration.pdf)** for a full visual walkthrough of the generated directory structure and final Excel report.
