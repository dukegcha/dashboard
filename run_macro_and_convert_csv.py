import os
import win32com.client
import time

RAW_DIR = r"C:\Users\duke.cha\Desktop\all_gi\raw"
CLEANED_DIR = r"C:\Users\duke.cha\Desktop\all_gi\cleaned"
CSV_DIR = r"C:\Users\duke.cha\Desktop\all_gi\csv"
MACRO_NAME = "PERSONAL.XLSB!savelograw"

# Ensure output directories exist
os.makedirs(CLEANED_DIR, exist_ok=True)
os.makedirs(CSV_DIR, exist_ok=True)

PERSONAL_XLSB_PATH = r"C:\\Users\\duke.cha\\AppData\\Roaming\\Microsoft\\Excel\\XLSTART\\PERSONAL.XLSB"

def run_macro_on_files():
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False  # Suppress Excel pop-ups and alerts
    try:
        # Open PERSONAL.XLSB so macros are available
        personal_wb = excel.Workbooks.Open(PERSONAL_XLSB_PATH)
        for filename in os.listdir(RAW_DIR):
            if filename.lower().endswith((".xls", ".xlsx", ".xlsm")):
                raw_path = os.path.join(RAW_DIR, filename)
                wb = excel.Workbooks.Open(raw_path)
                # Run the macro
                excel.Application.Run(MACRO_NAME)
                wb.Close(SaveChanges=True)
                # Give Excel a moment to finish writing files
                time.sleep(1)
        personal_wb.Close(SaveChanges=False)
    finally:
        excel.Quit()

def convert_cleaned_to_csv():
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    try:
        for filename in os.listdir(CLEANED_DIR):
            if filename.lower().endswith((".xls", ".xlsx", ".xlsm")):
                cleaned_path = os.path.join(CLEANED_DIR, filename)
                wb = excel.Workbooks.Open(cleaned_path)
                # Save as CSV
                base = os.path.splitext(filename)[0]
                csv_path = os.path.join(CSV_DIR, base + ".csv")
                wb.SaveAs(csv_path, FileFormat=6)  # 6 = xlCSV
                wb.Close(False)
    finally:
        excel.Quit()

if __name__ == "__main__":
    print("Running macro on raw files...")
    run_macro_on_files()
    print("Converting cleaned files to CSV...")
    convert_cleaned_to_csv()
    print("Done.")
