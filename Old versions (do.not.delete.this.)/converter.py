import pandas as pd
import json
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog
import os
import re
import shutil
from datetime import datetime

def sanitize_filename(filename):
    """Ensures rule names don't crash the file system."""
    return re.sub(r'[\\/*?:"<>|]', "", filename)

def load_json_safe(file_path):
    """Handles various encodings to ensure the master file opens every time."""
    try:
        with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
            return json.load(f).get('resources', [])
    except:
        with open(file_path, 'r', encoding='cp1252', errors='ignore') as f:
            return json.load(f).get('resources', [])

def create_excel_logic(rules, save_path):
    """Builds the 16-column audit sheet with precise formatting."""
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
    
    # Header format: Light Blue
    h_fmt = workbook.add_format({'bold': True, 'bg_color': '#ADD8E6', 'border': 1, 'align': 'center', 'valign': 'vcenter', 'text_wrap': True})
    # Body format: Center aligned
    b_fmt = workbook.add_format({'text_wrap': True, 'align': 'center', 'valign': 'vcenter', 'border': 1})
    # KQL format: Consolas font
    k_fmt = workbook.add_format({'text_wrap': True, 'align': 'left', 'valign': 'vcenter', 'font_name': 'Consolas', 'border': 1})
    
    # Conditional Formatting Colors
    white, yellow, orange, blue, red = [workbook.add_format({'bg_color': c, 'border': 1, 'align': 'center', 'valign': 'vcenter'}) for c in ['#FFFFFF', '#FFFF00', '#FFC000', '#00B0F0', '#FF0000']]
    red.set_font_color('white')

    num_rows = len(df)
    worksheet.set_column(0, 0, 45, b_fmt)  # Rule Name Match
    worksheet.set_column(1, 7, 20, b_fmt)
    worksheet.set_column(8, 9, 85, k_fmt)
    worksheet.set_column(10, 14, 20, b_fmt)
    worksheet.set_column(15, 15, 30, b_fmt)

    # Dropdowns for Audit Judgment
    judgments = ['Default', 'Update on Hold', 'No. of alerts same before/after', 'OSG Modified Rules', 'Urgent Attention Needed']
    worksheet.data_validation(1, 15, num_rows, 15, {'validate': 'list', 'source': judgments})

    # Apply Conditional Formats
    for opt, fmt in zip(judgments, formats := [white, yellow, orange, blue, red]):
        worksheet.conditional_format(1, 0, num_rows, 15, {'type': 'formula', 'criteria': f'=$P2="{opt}"', 'format': fmt})

    for col, val in enumerate(df.columns):
        worksheet.write(0, col, val, h_fmt)
    
    worksheet.set_row(0, 45)
    writer.close()

def run_app():
    root = tk.Tk()
    root.title("SOC Audit Architect")
    root.geometry("600x450")
    root.configure(bg="#f0f2f5")

    def start_automated_audit():
        # 1. Capture SOC Engineer Identity
        engineer_name = simpledialog.askstring("Engineer ID", "Enter your full name:")
        if not engineer_name: return

        # 2. Librarian: Locate Source
        initial_dir = os.path.join(os.path.expanduser("~"), "Downloads")
        master_file = filedialog.askopenfilename(
            initialdir=initial_dir,
            title="Select Azure_Sentinel_analytics_rules JSON",
            filetypes=[("JSON files", "*.json")]
        )
        if not master_file: return
        
        # 3. Define Metadata
        client_input = simpledialog.askstring("Naming", "Which client is this for?")
        if not client_input: return
        
        today = datetime.now().strftime("%Y-%m-%d")
        date_input = simpledialog.askstring("Naming", "What is the date today?", initialvalue=today)
        if not date_input: return

        # 4. Target Pathing
        client_dir = filedialog.askdirectory(title=f"Select Root Folder for {client_input}")
        if not client_dir: return

        # Structure Paths
        project_root = os.path.join(client_dir, f"{date_input}_Audit")
        history_dir = os.path.join(client_dir, "Historical Master JSON's")
        new_master_name = f"{client_input}_{date_input}_MASTERJSON.json"
        archive_path = os.path.join(history_dir, new_master_name)

        # Overwrite Protection
        if os.path.exists(project_root):
            if not messagebox.askyesno("Warning", f"{date_input}_Audit already exists. Overwrite?"): return

        try:
            # 5. Librarian: Archive & Rename
            if not os.path.exists(history_dir): os.makedirs(history_dir)
            shutil.move(master_file, archive_path)

            # 6. Build Project Workspace
            for folder in ['hold', 'NEW', 'OLD', 'Updated JSON']:
                os.makedirs(os.path.join(project_root, folder), exist_ok=True)

            # 7. Unpack Baseline Rules
            rules = load_json_safe(archive_path)
            old_path = os.path.join(project_root, 'OLD')
            for r in rules:
                name = sanitize_filename(r.get('properties', {}).get('displayName', 'Rule'))
                with open(os.path.join(old_path, f"{name}.json"), 'w', encoding='utf-8') as out:
                    json.dump(r, out, indent=4)

            # 8. Generate Excel
            excel_filename = f"{client_input}_Sentinel_Audit_{date_input}.xlsx"
            create_excel_logic(rules, os.path.join(project_root, excel_filename))

            # 9. Audit Signature (Removed Readme/Workings per request)
            with open(os.path.join(project_root, "Audit_Signature.txt"), 'w') as sig:
                sig.write(f"SOC ENGINEER: {engineer_name}\n")
                sig.write(f"CREATION DATE: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
                sig.write(f"CLIENT: {client_input}\n")
                sig.write(f"TOTAL RULES: {len(rules)}")

            messagebox.showinfo("Success", f"Audit built for {client_input}.\nMaster archived and Signature stamped.")
            os.startfile(project_root)
            
        except Exception as e:
            messagebox.showerror("Error", f"Automation failed: {e}")

    # UI Design
    tk.Label(root, text="SOC Audit Architect", font=("Segoe UI", 22, "bold"), pady=40, bg="#f0f2f5").pack()
    tk.Button(root, text="🚀 START NEW PROJECT (Full Automation)", command=start_automated_audit, 
              width=40, height=3, bg="#28a745", fg="white", font=("Segoe UI", 12, "bold")).pack(pady=30)
    tk.Label(root, text="This software was created by EGM.", font=("Segoe UI", 10, "italic"), bg="#f0f2f5").pack(side="bottom", pady=20)
    
    root.mainloop()

if __name__ == "__main__":
    run_app()