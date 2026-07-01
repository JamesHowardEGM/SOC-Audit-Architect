import pandas as pd
import json
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog
import os
import re
import shutil
from datetime import datetime

# --- HELPER FUNCTIONS ---
def sanitize_filename(filename):
    return re.sub(r'[\\/*?:"<>|]', "", filename)

def load_json_safe(file_path):
    try:
        with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
            return json.load(f).get('resources', [])
    except:
        with open(file_path, 'r', encoding='cp1252', errors='ignore') as f:
            return json.load(f).get('resources', [])

def create_excel_logic(rules, save_path):
    rows = []
    for r in rules:
        p = r.get('properties', {})
        status_val = 'Enabled' if p.get('enabled') == True else 'Disabled'
        rows.append({
            'Rule Name': p.get('displayName', 'N/A'),
            'Current Version': p.get('templateVersion', 'N/A'),
            'Newest Version': '', 
            'Last Modified': 'N/A', 
            'Enabled/Disabled': status_val,
            'Changes/Modifications': '', 'Notes': '', 'Comments': '',
            'Old KQL': p.get('query', ''), 'Updated KQL': '', 
            'Changes': '', 'Test Results': '', 'Notes ': '', 
            'Before': '', 'After': '', 'Audit Judgment': 'Default'
        })
    
    df = pd.DataFrame(rows)
    writer = pd.ExcelWriter(save_path, engine='xlsxwriter')
    df.to_excel(writer, index=False, sheet_name='Rules')
    workbook, worksheet = writer.book, writer.sheets['Rules']
    
    h_fmt = workbook.add_format({'bold': True, 'bg_color': '#ADD8E6', 'border': 1, 'align': 'center', 'valign': 'vcenter', 'text_wrap': True})
    b_fmt = workbook.add_format({'text_wrap': True, 'align': 'center', 'valign': 'vcenter', 'border': 1})
    k_fmt = workbook.add_format({'text_wrap': True, 'align': 'left', 'valign': 'vcenter', 'font_name': 'Consolas', 'border': 1})
    
    white, yellow, orange, blue, red = [workbook.add_format({'bg_color': c, 'border': 1, 'align': 'center', 'valign': 'vcenter'}) for c in ['#FFFFFF', '#FFFF00', '#FFC000', '#00B0F0', '#FF0000']]
    red.set_font_color('white')

    num_rows = len(df)
    worksheet.set_column(0, 0, 45, b_fmt)
    worksheet.set_column(1, 7, 20, b_fmt)
    worksheet.set_column(8, 9, 85, k_fmt)
    worksheet.set_column(10, 14, 20, b_fmt)
    worksheet.set_column(15, 15, 30, b_fmt)

    judgments = ['Default', 'Update on Hold', 'No. of alerts same before/after', 'OSG Modified Rules', 'Urgent Attention Needed']
    worksheet.data_validation(1, 15, num_rows, 15, {'validate': 'list', 'source': judgments})

    for opt, fmt in zip(judgments, [white, yellow, orange, blue, red]):
        worksheet.conditional_format(1, 0, num_rows, 15, {'type': 'formula', 'criteria': f'=$P2="{opt}"', 'format': fmt})

    for col, val in enumerate(df.columns):
        worksheet.write(0, col, val, h_fmt)
    
    worksheet.set_row(0, 45)
    writer.close()

# --- MAIN APP ---
def start_automated_audit():
    # --- FIXED: Syntax and Double Quotes Corrected ---
    sync_confirm = messagebox.askyesno("SharePoint Sync Check",
        "Before starting, please confirm:\n\n"
        "1. Have you navigated to your 'Sentinel Audits' SharePoint folder in your browser?\n"
        "2. Have you clicked 'Add shortcut to OneDrive'?\n\n"
        "This is required to save the audit files to the correct OneDrive sync location.\n\n"
        "If this message does not make sense, please read the README documentation before proceeding.")
    
    if not sync_confirm:
        messagebox.showwarning("Incomplete Setup", "Please set up the OneDrive shortcut before running the architect.")
        return

    engineer_name = simpledialog.askstring("Engineer ID", "Enter your full name:")
    if not engineer_name: return

    # 1. Multi-File Merge Logic
    num_files = simpledialog.askinteger("Merge", "How many JSON files to merge?", initialvalue=1)
    if not num_files or num_files < 1: return

    combined_rules = []
    seen_rule_names = set()
    duplicates_found = 0
    files_to_delete = []
    
    initial_downloads = os.path.join(os.path.expanduser("~"), "Downloads")

    for i in range(num_files):
        master_file = filedialog.askopenfilename(
            initialdir=initial_downloads,
            title=f"Select JSON Part {i+1} of {num_files}",
            filetypes=[("JSON files", "*.json")]
        )
        if not master_file: return
        
        files_to_delete.append(master_file)
        part_rules = load_json_safe(master_file)
        for r in part_rules:
            rule_name = r.get('properties', {}).get('displayName', 'Unnamed Rule')
            if rule_name not in seen_rule_names:
                combined_rules.append(r)
                seen_rule_names.add(rule_name)
            else:
                duplicates_found += 1

    # 2. Project Context
    client_name = simpledialog.askstring("Naming", "Which client is this for?")
    if not client_name: return
    
    today = datetime.now().strftime("%Y-%m-%d")
    date_input = simpledialog.askstring("Naming", "Confirm Audit Date:", initialvalue=today)
    if not date_input: return

    # Default to OneDrive root if present, otherwise fall back to home directory
    onedrive_shortcut_path = os.path.join(
        os.path.expanduser("~"),
        "OneDrive"
    )

    client_dir = filedialog.askdirectory(
        initialdir=onedrive_shortcut_path if os.path.exists(onedrive_shortcut_path) else None,
        title=f"Select Root Folder for {client_name}"
    )
    if not client_dir: return

    # 3. Automatic Versioning
    project_root = os.path.join(client_dir, f"{date_input}_Audit")
    version = 1
    original_root = project_root
    while os.path.exists(project_root):
        version += 1
        project_root = f"{original_root}_v{version}"
    
    history_dir = os.path.join(client_dir, "Historical Master JSON's")
    new_master_name = f"{client_name}_{date_input}_v{version}_MASTERJSON.json"
    archive_path = os.path.join(history_dir, new_master_name)

    try:
        # 4. Save Master Record
        os.makedirs(history_dir, exist_ok=True)
        with open(archive_path, 'w', encoding='utf-8') as f:
            json.dump({"resources": combined_rules}, f, indent=4)

        # 5. Build Standard Structure
        for folder in ['hold', 'NEW', 'OLD', 'Updated JSON']:
            os.makedirs(os.path.join(project_root, folder), exist_ok=True)

        # 6. Unpack Rules
        old_path = os.path.join(project_root, 'OLD')
        for r in combined_rules:
            name = sanitize_filename(r.get('properties', {}).get('displayName', 'Rule'))
            with open(os.path.join(old_path, f"{name}.json"), 'w', encoding='utf-8') as out:
                json.dump(r, out, indent=4)

        # 7. Excel Architect Report
        excel_filename = f"{client_name}_Sentinel_Audit_{date_input}_v{version}.xlsx"
        create_excel_logic(combined_rules, os.path.join(project_root, excel_filename))

        # 8. Signature File
        with open(os.path.join(project_root, "Audit_Signature.txt"), 'w') as sig:
            sig.write(f"SOC ENGINEER: {engineer_name}\n")
            sig.write(f"TIMESTAMP: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
            sig.write(f"CLIENT: {client_name}\n")
            sig.write(f"SESSION VERSION: {version}\n")
            sig.write(f"TOTAL UNIQUE RULES: {len(combined_rules)}\n")
            sig.write(f"DUPLICATES REMOVED: {duplicates_found}")

        # 9. Cleanup original Downloads
        for file_path in files_to_delete:
            try: os.remove(file_path)
            except: pass

        messagebox.showinfo("Success", f"Audit v{version} Built!\nSaved to OneDrive Sync Folder.")
        os.startfile(project_root)
        
    except Exception as e:
        messagebox.showerror("Error", f"Automation failed: {e}")

def run_app():
    root = tk.Tk()
    root.title("SOC Audit Architect")
    root.geometry("600x450")
    root.configure(bg="#f0f2f5")
    tk.Label(root, text="SOC Audit Architect", font=("Segoe UI", 22, "bold"), pady=40, bg="#f0f2f5").pack()
    tk.Button(root, text="START NEW PROJECT", command=start_automated_audit, 
              width=40, height=3, bg="#28a745", fg="white", font=("Segoe UI", 12, "bold")).pack(pady=30)
    tk.Label(root, text="This software was created by EGM.", font=("Segoe UI", 10, "italic"), bg="#f0f2f5").pack(side="bottom", pady=20)
    root.mainloop()

if __name__ == "__main__":
    run_app()