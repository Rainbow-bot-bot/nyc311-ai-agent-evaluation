import openpyxl
import os

from build_sheet1 import build_sheet_1
from build_sheet2 import build_sheet_2
from build_sheet3 import build_sheet_3
from build_sheet4 import build_sheet_4
from build_sheet5 import build_sheet_5
from build_sheet6 import build_sheet_6

print("==================================================")
print("STARTING MASTER WORKBOOK GENERATION")
print("==================================================")

wb = openpyxl.Workbook()
wb.remove(wb.active) # Remove default sheet

build_sheet_1(wb)
build_sheet_2(wb)
build_sheet_3(wb)
build_sheet_4(wb)
build_sheet_5(wb)
build_sheet_6(wb)

FINAL_PATH = "D:/项目2/Gemini/最终成果.xlsx"
wb.save(FINAL_PATH)

print("==================================================")
print(f"SUCCESS! Master workbook saved to {FINAL_PATH}")
file_size = os.path.getsize(FINAL_PATH)
print(f"File Size: {file_size:,} bytes ({file_size / 1024:.1f} KB)")
print(f"Sheet Names: {wb.sheetnames}")
print("==================================================")
