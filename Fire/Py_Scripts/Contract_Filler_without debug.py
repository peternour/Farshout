# -*- coding: utf-8 -*-
import os
import time
import pandas as pd
import xlwings as xw
from tkinter import Tk, filedialog, messagebox

# ==============================
# USER CONFIGURATION AREA
# ==============================
HEADER_MAPPING = {
    "اسم العميل عربى": "اسم العميل",  # "Customer Name Arabic" : "Customer Name"
    "كود العميل": "كود العميل",        # "Customer Code" : "Customer Code"
    "الرقم القومى": "رقم قومي",        # "National ID" : "National ID"
    "النشاط الرئيسى": "نوع النشاط الرئيسي",  # "Main Activity" : "Main Activity Type"
    "النشاط الثانى": "فرعي 1",          # "Secondary Activity" : "Sub 1"
    "النشاط الثالث": "فرعي 2",         # "Tertiary Activity" : "Sub 2"
    "برنامج الاقراض": "مصدر التمويل / اسم المشورع الممول",  # "Lending Program" : "Funding Source/Project Name"
    "تليفون العميل": "موبايل",         # "Customer Phone" : "Mobile"
    "عنوان النشاط": "عنـــوان المشــروع",  # "Activity Address" : "Project Address"
    "مبلغ القرض": "القرض",             # "Loan Amount" : "Loan"
    "عدد الاقساط": "الاقساط"           # "Number of Installments" : "Installments"
}

TEMPLATE_SHEET_NAME = "تفصيلي"  # "Detailed"
START_ROW = 6

# ==============================
# MAIN SCRIPT
# ==============================
def main():
    root = Tk()
    root.withdraw()
    
    try:
        # Select source file
        messagebox.showinfo("Source Data", "Please select the source data file")
        source_path = filedialog.askopenfilename(
            title="SOURCE FILE SELECTION",
            filetypes=[("Excel files", "*.xlsx;*.xls;*.xltm")]
        )
        if not source_path:
            return

        # Select template file
        messagebox.showinfo("Template Selection", "Please select the contract file to update")
        template_path = filedialog.askopenfilename(
            title="CONTRACT FILE SELECTION",
            filetypes=[("Excel files", "*.xlsx;*.xls;*.xltm")]
        )
        if not template_path:
            return

        # Confirm before modifying template
        confirm = messagebox.askyesno("Confirmation Required", 
            f"ARE YOU SURE? THIS OPERATION CANNOT BE UNDONE!\n"
            f"DATA WILL BE ADDED TO:\n{template_path}\n"
            f"This action cannot be reversed!")
        
        if not confirm:
            messagebox.showinfo("Operation Cancelled", "The operation has been cancelled.")
            return

        # Process files
        copy_data_xlwings(source_path, template_path)
        messagebox.showinfo("Success", f"Template updated successfully!\nData has been added to:\n{template_path}")
        
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred:\n{str(e)}")
    finally:
        root.destroy()

def copy_data_xlwings(source_path, template_path):
    # Load source data
    df = pd.read_excel(source_path, engine='openpyxl')
    if df.empty:
        messagebox.showerror("Error", "Source file is empty!")
        return
            
    # Verify source headers
    source_headers = df.columns.tolist()
    missing_headers = [h for h in HEADER_MAPPING.keys() if h not in source_headers]
    if missing_headers:
        raise ValueError(f"Missing headers in source file: {', '.join(missing_headers)}")
    
    # Process template
    app = xw.App(visible=False)
    try:
        wb = app.books.open(template_path)
        try:
            sheet = wb.sheets[TEMPLATE_SHEET_NAME]
        except:
            raise ValueError(f"Sheet not found: {TEMPLATE_SHEET_NAME}")
        
        # Get template headers
        header_row = sheet.range((START_ROW-1, 1), (START_ROW-1, 50)).value
        template_headers = {}
        for idx, header in enumerate(header_row):
            if header:
                template_headers[str(header).strip()] = idx + 1
        
        # Verify template headers
        missing_template_headers = [
            h for h in HEADER_MAPPING.values() 
            if str(h).strip() not in template_headers
        ]
        if missing_template_headers:
            raise ValueError(f"Missing headers in template: {', '.join(missing_template_headers)}")
        
        # Write data to template
        current_row = START_ROW
        for _, row in df.iterrows():
            for src_header, tpl_header in HEADER_MAPPING.items():
                col_index = template_headers[str(tpl_header).strip()]
                sheet.range((current_row, col_index)).value = row[src_header]
            current_row += 1
        
        # Save with backup
        backup_path = template_path.replace(".xlsx", "_backup.xlsx")
        wb.save(backup_path)
        wb.save()
        
    finally:
        app.quit()

if __name__ == "__main__":
    main()