# -*- coding: utf-8 -*-
import os
import time
import logging
import pandas as pd
import xlwings as xw
from tkinter import Tk, filedialog, messagebox
import sys

# ==============================
# USER CONFIGURATION AREA
# ==============================

# 1. Define your header mapping here (source header → template header)
HEADER_MAPPING = {
    "اسم العميل عربى": "اسم العميل",
    "كود العميل": "كود العميل",
    "الرقم القومى": "رقم قومي",
    "النشاط الرئيسى": "نوع النشاط الرئيسي",
    "النشاط الثانى": "فرعي 1",
    "النشاط الثالث": "فرعي 2", 
    "برنامج الاقراض": "مصدر التمويل / اسم المشورع الممول", 
    "تليفون العميل": "موبايل", 
    "عنوان النشاط": "عنـــوان المشــروع",
    "مبلغ القرض": "القرض", 
    "عدد الاقساط": "الاقساط"
}

# 2. Configure template settings
TEMPLATE_SHEET_NAME = "تفصيلي"
START_ROW = 6  # Always start writing from this row

# 3. Configure logging
LOG_FILE = "excel_processor_debug.log"

# ==============================
# MAIN SCRIPT
# ==============================

def setup_logging():
    """Configure logging to file and console with detailed debugging"""
    logging.basicConfig(
        level=logging.DEBUG,  # Set to DEBUG for maximum information
        format='%(asctime)s - %(levelname)s - %(message)s',
        handlers=[
            logging.FileHandler(LOG_FILE, encoding='utf-8'),
            logging.StreamHandler()
        ]
    )
    logging.info("Logging system initialized")

def main():
    setup_logging()
    logging.info("Starting Excel Processor")
    
    # Set up Tkinter
    root = Tk()
    root.withdraw()
    
    try:
        # File selection dialog for source data
        messagebox.showinfo("source", "source Data")
        source_path = filedialog.askopenfilename(
            title="SOURCE",
            filetypes=[("Excel files", "*.xlsx;*.xls;*.xltm")]
        )
        if not source_path:
            logging.info("Operation cancelled: No source file selected")
            return
        logging.info(f"Selected source file: {source_path}")

        # File selection dialog for template
        messagebox.showinfo("TEMPLATE", "SELECT CONTRACT FILE FOR UPDATE")
        template_path = filedialog.askopenfilename(
            title="CONTRACT",
            filetypes=[("Excel files", "*.xlsx;*.xls;*.xltm")]
        )
        if not template_path:
            logging.info("Operation cancelled: No template file selected")
            return
        logging.info(f"Selected template file: {template_path}")

        # Confirm before modifying template
        confirm = messagebox.askyesno("UPDATE INFO", 
            f"ARE YOU SURE THIS OPERATION CANT BE UNDONE\n"
            f"THE DATA WILL ADD TO:\n{template_path}\n"
            f"لا يمكن التراجع عن هذا الإجراء!")
        
        if not confirm:
            logging.info("Operation cancelled by user")
            messagebox.showinfo("تم الإلغاء", "تم إلغاء العملية.")
            return

        # Process files
        copy_data_xlwings(source_path, template_path)
        
        messagebox.showinfo("تم بنجاح", f"تم تحديث القالب بنجاح!\nتمت إضافة البيانات إلى:\n{template_path}")
        logging.info("Operation completed successfully")
        
    except Exception as e:
        error_msg = f"حدث خطأ:\n{str(e)}\n\nتفاصيل الخطأ مسجلة في: {LOG_FILE}"
        logging.exception("Critical error occurred")
        messagebox.showerror("خطأ", error_msg)
    finally:
        # Ensure Tkinter root window is destroyed
        try:
            root.destroy()
        except:
            pass

def verify_environment():
    """Verify that required packages are installed and accessible"""
    logging.info("Verifying environment...")
    try:
        import pandas as pd
        import xlwings as xw
        logging.info("All required packages are available")
    except ImportError as e:
        logging.error(f"Missing package: {str(e)}")
        raise

def copy_data_xlwings(source_path, template_path):
    """
    Process Excel files using xlwings with extensive debugging
    """
    start_time = time.time()
    logging.info("Starting data copy process")
    
    # Verify environment first
    verify_environment()
    
    # STEP 1: Load source data with detailed debugging
    logging.info("Loading source data...")
    try:
        # Read with explicit encoding and handling
        df = pd.read_excel(source_path, engine='openpyxl')
        logging.info(f"Source file loaded successfully. Shape: {df.shape}")
        
        if df.empty:
            logging.error("Source file is empty!")
            messagebox.showerror("خطأ", "SRC FILE EMPTY!")
            return
            
        # Log detailed information about source data
        logging.debug(f"Source columns: {list(df.columns)}")
        logging.debug(f"First row sample: {df.iloc[0].to_dict()}")
        logging.debug(f"Data types:\n{df.dtypes}")
    except Exception as e:
        logging.error(f"Error loading source file: {str(e)}")
        raise
    
    # Check source headers
    source_headers = df.columns.tolist()
    logging.info("Source headers: " + ", ".join(source_headers))
    
    missing_headers = [h for h in HEADER_MAPPING.keys() if h not in source_headers]
    if missing_headers:
        err_msg = f"Missing headers in source: {', '.join(missing_headers)}"
        logging.error(err_msg)
        messagebox.showerror("hEADER ERR", err_msg)
        raise ValueError(err_msg)
    
    # STEP 2: Open template with detailed debugging
    logging.info("Opening template file...")
    app = None
    try:
        app = xw.App(visible=True)  # Make visible to see what's happening
        logging.info(f"Excel application started: {app}")
        
        wb = app.books.open(template_path)
        logging.info(f"Opened template: {template_path}")
        
        # Log all sheets in workbook
        sheet_names = [s.name for s in wb.sheets]
        logging.info(f"Available sheets: {', '.join(sheet_names)}")
        
        try:
            sheet = wb.sheets[TEMPLATE_SHEET_NAME]
            logging.info(f"Accessed sheet: {TEMPLATE_SHEET_NAME}")
        except Exception: 
            err_msg = f"Sheet not found: {TEMPLATE_SHEET_NAME}. Available sheets: {', '.join(sheet_names)}"
            logging.error(err_msg)
            messagebox.showerror("ورقة غير موجودة", err_msg)
            raise ValueError(err_msg)
        
        # STEP 3: Get template headers with detailed verification
        logging.info("Reading template headers...")
        header_row_num = START_ROW - 1  # Headers are expected one row above data
        
        # Read headers with expanded range
        header_range = sheet.range((header_row_num, 1), (header_row_num, 50))
        template_headers = {}
        
        for cell in header_range: 
            if cell.value is not None:
                try:
                    # Handle different data types
                    header_value = str(cell.value).strip()
                    template_headers[header_value] = cell.column
                    logging.debug(f"Found header: '{header_value}' at column {cell.column}")
                except Exception as e:
                    logging.error(f"Error processing header at column {cell.column}: {str(e)}")
        
        logging.info(f"Template headers found: {', '.join(template_headers.keys())}")
        
        # Verify all target headers exist in the template
        missing_template_headers = []
        for h in HEADER_MAPPING.values():
            normalized_h = h.strip()
            if normalized_h not in template_headers:
                missing_template_headers.append(h)
                logging.warning(f"Target header not found: '{h}' (normalized: '{normalized_h}')")

        if missing_template_headers:
            err_msg = f"Missing target headers in template: {', '.join(missing_template_headers)}"
            logging.error(err_msg)
            messagebox.showerror("رؤوس مفقودة", err_msg)
            raise ValueError(err_msg)
        
        # STEP 4: Copy data with detailed verification
        current_row = START_ROW
        logging.info(f"Starting data filling at row: {current_row}")

        total_rows = len(df)
        logging.info(f"Preparing to write {total_rows} rows")
        
        # Log column positions for debugging
        col_positions = {}
        for h in HEADER_MAPPING.values():
            normalized_h = h.strip()
            col_positions[h] = template_headers.get(normalized_h, "NOT FOUND")
        logging.info(f"Column positions: {col_positions}")
        
        # Test write to verify Excel is writable
        test_cell = sheet.range((current_row, 1))
        original_value = test_cell.value
        test_cell.value = "XLWINGS_TEST"
        logging.info(f"Wrote test value to cell ({current_row}, 1): 'XLWINGS_TEST'")
        
        # Write actual data
        for i, row in df.iterrows():
            logging.debug(f"Processing row {i+1}/{total_rows}")
            
            for src_header, tpl_header in HEADER_MAPPING.items():
                normalized_tpl = tpl_header.strip()
                
                if normalized_tpl in template_headers:
                    col_index = template_headers[normalized_tpl]
                    cell_value = row[src_header]
                    
                    # Log first row writes in detail
                    if i == 0:
                        logging.info(f"First row: Writing to row {current_row}, col {col_index} ({tpl_header}): {cell_value}")
                    
                    # Write value to the cell
                    sheet.range((current_row, col_index)).value = cell_value
                else:
                    logging.warning(f"Skipping column: '{tpl_header}' not found in template")
            
            current_row += 1
        
        # Verify test write
        test_value = sheet.range((START_ROW, 1)).value
        logging.info(f"Test cell value after writes: '{test_value}' (should be 'XLWINGS_TEST')")
        
        # Verify first data point
        first_data_col = template_headers[list(HEADER_MAPPING.values())[0].strip()]
        first_data_value = sheet.range((START_ROW, first_data_col)).value
        expected_value = df.iloc[0][list(HEADER_MAPPING.keys())[0]]
        logging.info(f"First data point: {first_data_value} (expected: {expected_value})")
        
        # STEP 5: Save changes with verification
        logging.info("Saving template...")
        
        # Create backup before saving
        backup_path = template_path.replace(".xlsx", "_backup.xlsx")
        wb.save(backup_path)
        logging.info(f"Created backup at: {backup_path}")
        
        wb.save()
        logging.info(f"Template saved successfully to: {template_path}")
        
        # Verify file size changed
        original_size = os.path.getsize(template_path)
        logging.info(f"Final file size: {original_size} bytes")
        
        # STEP 6: Close resources
        wb.close()
        app.quit()
        
        elapsed = time.time() - start_time
        logging.info(f"Operation completed in {elapsed:.2f} seconds")
        
        # Final verification
        if test_value != "XLWINGS_TEST":
            logging.error("Test write verification failed! Changes may not have been saved.")
            messagebox.showerror("خطأ في الحفظ", "فشل التحقق من الحفظ. يرجى التحقق من ملف النسخة الاحتياطية.")
        else:
            logging.info("Test write verification successful")
        
        return True
        
    except Exception as e:
        logging.error(f"Error during processing: {str(e)}")
        # Try to save for debugging if possible
        try:
            crash_path = template_path.replace(".xlsx", "_CRASH.xlsx")
            wb.save(crash_path)
            logging.info(f"Saved crash recovery file to: {crash_path}")
        except:
            pass
            
        if app:
            try:
                app.quit()
            except:
                pass
        raise

if __name__ == "__main__":
    main()