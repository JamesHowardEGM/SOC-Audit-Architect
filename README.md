#TL;DR
# ⚙️ SOC Audit Architect

An automation tool designed to streamline Microsoft Sentinel audits. It handles the "heavy lifting" of file management, rule decomposition, and reporting so you can focus on the actual security analysis.

---

## 📂 What the Tool Does
When you run the tool, it automatically organizes your workspace and generates the following files:

* **Automated Folders**: Creates a timestamped session folder (e.g., `2026-02-17_Audit`) to keep every audit engagement isolated and organized.
* **Historical Archive**: Relocates your raw Sentinel export to a `Historical Master JSON's` folder to ensure you always have an untouched "Source of Truth".
* **Rule Unpacker**: Decomposes the master file into individual JSON rules inside `OLD` and `NEW` folders for granular review.
* **Interactive Excel Report**: Generates a formatted spreadsheet with light blue headers, KQL extractions, and built-in "Pass/Fail" dropdown menus.
* **Compliance Signature**: Produces an `Audit_Signature.txt` file that logs the timestamp and user profile to verify when the audit was completed.

---

## 🚀 Quick Start
1. **Prepare**: Ensure you have a parent `Audits` folder with a sub-folder for your client (e.g., `Audits/Client Name`).
2. **Run**: Launch the `SOC Audit Architect.exe`.
3. **Select**: Use the file explorer to pick your Sentinel JSON export from your `Downloads`.
4. **Result**: The tool will build the directory structure and pop out your finished Excel report and Audit Signature.

---

## 🛠️ Requirements
* **OS**: Windows 10 or 11.
* **Access**: Read/Write permissions for your local `Audits` and `Downloads` folders.
* **Dependencies**: (If running from code) `pandas`, `xlsxwriter`, `openpyxl`.

---

## 📖 Walkthrough
For a full visual guide with screenshots of the generated directory structure and the final report, please refer to the **[SOC Audit Architect Demonstration.pdf](./SOC%20Audit%20Architect%20Demonstration.pdf)** included in this repository.

---
---
---

# SOC-Audit-Architect
An automation tool for Microsoft Sentinel audits that archives master JSON files, unpacks individual rules, and generates formatted Excel reports with Audit Signatures.

# ⚙️ SOC Audit Architect: Operational Workflow & Logic

This document details the computational steps and automated file management logic utilized by the **SOC Audit Architect**.

---

## 1. Automated Workspace Initialization (Directory Orchestration)
The system's primary directive is to establish a standardized, multi-tier environment to ensure audit consistency.

* **Parent Pathing**: The architect operates within a root directory structure (e.g., `Audits` > `New Client`).
* **The "Directory Engine"**: Upon execution, the software utilizes `os.makedirs` to check for and instantiate a specific folder hierarchy:
    * **`Historical Master JSON's`**: A top-level archive directory for raw "Source of Truth" exports.
    * **`[YYYY-MM-DD]_Audit`**: A unique, timestamped session folder containing all outputs for the current audit engagement.
* **Session Sub-folders**: Inside the dated audit folder, the script bootstraps the following environment:
    * **`NEW`**: Storage for freshly processed data artifacts.
    * **`OLD`**: Contains decomposed, granular snapshots of individual security rules for historical comparison.
    * **`hold`** & **`Updated JSON`**: Internal directories used for staging and data versioning.

## 2. File Ingestion & Data Integrity (The "Librarian")
The script acts as a gatekeeper to ensure that raw telemetry from Microsoft Sentinel is preserved without corruption.

* **Atomic Move Operation**: Utilizing the `shutil` library, the script locates the raw Sentinel export (typically in the `Downloads` directory) and relocates it to the `Historical Master JSON's` archive.
* **Standardized Serialization**: During relocation, the file is renamed to a strict convention: `FPC_[YYYY-MM-DD]_MASTERJSON.json`.
* **Version Preservation**: By archiving the original file before any processing occurs, the script ensures that the "Source of Truth" remains immutable for long-term audit history.

## 3. Rule Decomposition (The "Unpacker")
The software decomposes monolithic master files into granular components to facilitate precise peer review.

* **Resource Iteration**: The script parses the JSON structure and enters a loop to iterate through every security rule object found within the `resources` array.
* **Regex Sanitization**: Rule names often contain characters illegal in the Windows File System (e.g., `:`, `\`, `*`). The script uses Regular Expressions (`re`) to strip these characters, ensuring a successful `write` operation.
* **Granular Persistence**: Each security rule is saved as a standalone `.json` file within the **`OLD`** folder. This allows for line-by-line KQL code "diffing" between different audit periods.

## 4. Interactive Report Generation (The "Architect")
The software utilizes the `pandas` and `xlsxwriter` engines to transform raw data into a high-fidelity, interactive auditor interface.

* **DataFrame Orchestration**: Rule metadata (Name, KQL, Severity, etc.) is mapped into a structured 16-column `pandas` DataFrame.
* **UI Formatting**: The `xlsxwriter` engine applies professional styling, including a light blue header fill (`#ADD8E6`) and text-wrapping logic for long KQL queries.
* **Logic Integration**:
    * **Dropdown Menus**: The script injects "Data Validation" objects into the "Audit Result" column, allowing for selectable `Pass`, `Fail`, or `In-Progress` statuses.
    * **Conditional Formatting**: Absolute reference formulas (e.g., `=$P2="Pass"`) are applied to trigger row-wide background color changes based on the auditor's selection.

## 5. Final Result & Non-Repudiation (The "Signature")
Upon completion, the software delivers a finalized audit package inside the dated session folder:

* **The Audit Report**: A formatted, interactive Excel workbook (e.g., `new_client_Sentinel_Audit_2026-02-17.xlsx`).
* **The Compliance Artifact**: An `Audit_Signature.txt` file identifying the local Windows user profile and the precise timestamp of completion to ensure accountability.


---

* # 📋 System Requirements & Dependencies

To ensure the **SOC Audit Architect** functions correctly, the following environment and libraries must be present on the host system.

---

## 1. Environment Specifications
* **Operating System**: Windows 10/11 (Required for directory pathing and `.exe` compatibility).
* **Python Version**: Python 3.8 or higher.
* **Permissions**: The user must have **Read/Write** permissions for the `Downloads` folder and the local `Audits` directory.

---

## 2. Core Python Dependencies
The script utilizes several high-performance libraries to handle data processing and Excel orchestration:

* **Pandas**: Used for data manipulation and structured DataFrame management.
* **XlsxWriter**: The primary engine used to create the formatted Excel reports and inject conditional formatting.
* **Openpyxl**: Required as a supporting engine for advanced Excel file handling.
* **Tkinter**: The standard Python interface for the graphical file-selection window.

---

## 3. Directory Prerequisites
Before execution, the software expects the following structure to be established by the user:
* A **Parent Folder** (e.g., `C:\Users\James-OneStep\OneDrive\Documents\Audits`).
* A **Client Sub-folder** (e.g., `New Client`).
* **Note**: All other folders (`Historical Master JSON's`, `OLD`, `NEW`) will be automatically generated by the "Directory Engine" upon the first run.

---

## 🛠️ Installation of Dependencies
If running from the source code rather than the `.exe`, you can install all necessary libraries using the provided `requirements.txt` file:

```bash
pip install -r requirements.txt
