⚙️ SOC Audit Architect

An automation tool designed to streamline Microsoft Sentinel audits. It handles the "heavy lifting" of file management, rule decomposition, and reporting so you can focus on the actual security analysis.

📂 What the Tool Does

When you run the tool, it automatically organizes your workspace and generates the following files:

    Intelligent Merging: Overcomes Azure Portal limitations by allowing you to merge multiple JSON exports (e.g., from different pages) into a single master audit.

    Duplicate Filtering: Automatically detects and removes overlapping rules during the merge process to ensure data accuracy.

    Evidence Protection: Automatically detects if an audit already exists for that date and applies versioning (e.g., _v2, _v3) to prevent overwriting historical data.

    Historical Archive: Relocates your merged master JSON to a Historical Master JSON's folder to ensure an untouched "Source of Truth."

    Rule Unpacker: Decomposes the master file into individual JSON rules inside OLD and NEW folders for granular review.

    Interactive Excel Report: Generates a formatted spreadsheet with light blue headers, KQL extractions, and built-in "Pass/Fail" dropdown menus.

    Compliance Signature: Produces an Audit_Signature.txt file logging the engineer's name, timestamp, session version, and the count of duplicates removed.

🚀 Quick Start

    Prepare: Ensure you have a parent Audits folder with a sub-folder for your client.

    Run: Launch the SOC Audit Architect.exe.

    Merge: Enter the number of JSON files you need to combine.

    Select: Pick your files and choose your client's root directory.

    Result: The tool builds the structure and opens the folder containing your Excel report and Audit Signature.

🛠️ Requirements & Setup

    OS: Windows 10 or 11.

    Access: Read/Write permissions for Audits and Downloads folders.

💻 Dependencies (If running from code)

You can install all necessary external libraries with this one-line command in your VS Code terminal (Ctrl + Shift + `):
PowerShell

python -m pip install pandas xlsxwriter openpyxl

Library	Type			Purpose
Pandas		External	Processes Sentinel JSON data into structured tables.
XlsxWriter	External	Builds Excel reports with branding and dropdowns.
Openpyxl	External	Supports modern .xlsx file handling.
Tkinter		Built-in	Handles the file-selection and pop-up windows.
OS / Shutil	Built-in	Manages directories and archives files.



📖 In-depth Operational Logic
1. Automated Workspace Initialization

The system utilizes os.makedirs to instantiate a multi-tier hierarchy. If the folder [YYYY-MM-DD]_Audit already exists, the Evidence Protection logic increments the name to _v2 to preserve the previous session's work.
2. The "Librarian" (Merge & Duplicate Logic)

The tool acts as a gatekeeper for data integrity.

    De-duplication: It utilizes a hashing set to track rule names. If a rule appears in multiple JSON parts, it is recorded only once.

    Standardized Serialization: The final merged file is saved to the Historical Master JSON's folder with a version-controlled naming convention: Client_Date_v#_MASTERJSON.json.

3. Rule Decomposition (The "Unpacker")

The script parses the master file and iterates through the resources array. It uses Regular Expressions (re) to sanitize rule names, ensuring that characters like : or / do not crash the file-writing process as rules are saved individually in the OLD folder.
4. Interactive Report Generation (The "Architect")

The xlsxwriter engine transforms raw telemetry into a 16-column interface:

    UI Formatting: Applies light blue fills (#ADD8E6) and Consolas font styles for KQL.

    Logic Integration: Injects Data Validation for "Audit Judgments" and Row-wide Conditional Formatting based on status (e.g., Red for "Urgent Attention").

5. Final Result & Non-Repudiation

The Audit_Signature.txt provides an immutable record of the work, capturing:

    Engineer Identity: Captured at the start of the session.

    Audit Versioning: Distinguishes between multiple attempts on the same day.

    Integrity Metrics: Logs the total unique rules versus duplicates skipped.



📂 Directory & File Breakdown

    .vs: A hidden folder created by Visual Studio. It stores your local display settings, window layouts, and IntelliSense cache so your editor remembers exactly where you left off in converter_1.1.py.

    Dev Folder - Updated Versions: This is your active "Staging" area. It is specifically labeled as your work folder, where the code for the future v2.0 (Noise Reduction features) is being written.

    Old versions (do.not.delete.this.): Your historical archive. This contains Version 1.0 of the tool, preserved in case you ever need to reference the original single-file logic.

    venv: The Virtual Environment. This folder contains a local copy of Python and the specific libraries (like pandas) needed to run the script without affecting the rest of your computer.

    .gitignore: A configuration file that tells version control systems (like Git) to ignore the .vs and venv folders so you don't accidentally upload massive, unnecessary system files to a repository.

    converter_1.1.py: The Current Stable Production script. This is the v1.1 version we just finalized, featuring the multi-file merge and auto-versioning for template audits.

    Change Log.txt: Your project's history book. It documents every fix and feature added during the transition from 1.0 to 1.1.

    README.txt & Demonstration Files: These provide documentation and presentation materials for the team to understand how to use the "SOC Audit Architect" correctly.

🛡️ Client Data Structure

While the above are your tool files, the script saves its output into your synced SharePoint client folders. Each client listed (e.g., AQW - Aqwest, BG - Bellevue Gold) will now receive the standardized _Audit folders generated by your converter_1.1.py.