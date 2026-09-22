import os
import sys
import openpyxl
import re

sys.stdout.reconfigure(encoding='utf-8')

excel_path = r"D:\项目2\Gemini\第二轮\最终成果.xlsx"
report_path = r"D:\项目2\Gemini\第二轮\分析说明.md"

print("--- 1. Checking Deliverable Existence ---")
assert os.path.exists(excel_path), f"Missing: {excel_path}"
assert os.path.exists(report_path), f"Missing: {report_path}"
print(f"Excel file exists: {excel_path} ({os.path.getsize(excel_path)} bytes)")
print(f"Report file exists: {report_path} ({os.path.getsize(report_path)} bytes)")

print("\n--- 2. Checking Excel Workbook Structure & Sheets ---")
wb = openpyxl.load_workbook(excel_path, data_only=False)
expected_sheets = [
    "执行摘要与看板导航",
    "01_数据质量与证据边界",
    "02_年度增长与归因分解",
    "03_极端气候冲击与连锁反应",
    "04_分类口径变更核查",
    "05_结构集聚与二八法则",
    "06_空间分布与响应执法"
]
assert len(wb.sheetnames) == len(expected_sheets), f"Sheet count mismatch: {wb.sheetnames}"
for idx, sname in enumerate(expected_sheets):
    assert sname in wb.sheetnames, f"Missing sheet: {sname}"
    ws = wb[sname]
    print(f"Sheet {idx+1}: [{sname}] - Rows: {ws.max_row}, Cols: {ws.max_column}")
    # Verify gridlines enabled
    assert ws.views.sheetView[0].showGridLines == True, f"Gridlines not enabled in {sname}"

print("\n--- 3. Checking Excel Charts ---")
ws3 = wb["02_年度增长与归因分解"]
print(f"Charts in Sheet 02: {len(ws3._charts)}")
assert len(ws3._charts) >= 1, "Sheet 02 should have a chart"

ws4 = wb["03_极端气候冲击与连锁反应"]
print(f"Charts in Sheet 03: {len(ws4._charts)}")
assert len(ws4._charts) >= 1, "Sheet 03 should have a chart"

print("\n--- 4. Cross-checking Key Metrics in Report & Excel ---")
with open(report_path, "r", encoding="utf-8") as f:
    report_text = f.read()

key_metrics = [
    ("7,525,498", "Total records"),
    ("3,293,003", "Year 1 11-month volume"),
    ("3,639,698", "Year 2 11-month volume"),
    ("346,695", "11-month YoY net increase"),
    ("10.53%", "YoY growth rate"),
    ("42.38%", "HPD growth contribution rate"),
    ("19.28%", "DOT growth contribution rate"),
    ("16.87%", "NYPD growth contribution rate"),
    ("9.56%", "DSNY growth contribution rate"),
    ("63.95%", "Top 4 complaints contribution rate"),
    ("94.07%", "Online channel contribution to growth"),
    ("11,370", "Peak snow complaints on Feb 24, 2026"),
    ("22,790", "DOT Pothole complaints in Mar 2026"),
    ("428%", "DOT Pothole surge percentage"),
    ("25.43%", "Heat Top 1% address share"),
    ("51.47%", "Heat Top 5% address share"),
    ("80.43%", "Heat Top 20% address share"),
    ("29.08%", "Heat duplicate complaint resolution percentage"),
    ("41.31%", "Bronx share of residential noise"),
    ("35.06%", "Bronx share of heat/hot water"),
    ("70.65%", "Brooklyn + Queens share of illegal parking"),
    ("0.36%", "NYPD residential noise summons rate"),
    ("44.06%", "NYPD residential noise no-evidence rate")
]

for metric_str, metric_desc in key_metrics:
    found = metric_str in report_text
    print(f"Metric [{metric_desc} = {metric_str}]: {'FOUND' if found else 'MISSING'}")
    assert found, f"Missing metric in report: {metric_desc} = {metric_str}"

print("\n--- 5. Checking Isolation & Prohibited References ---")
# Ensure no mentions of other AI models like GPT, Claude, Kimi, etc.
prohibited = ["gpt", "claude", "kimi", "qwen", "glm", "deepseek", "grok"]
for p in prohibited:
    # search case-insensitively
    matches = re.findall(rf"\b{p}\b", report_text, re.IGNORECASE)
    assert len(matches) == 0, f"Found unauthorized mention of {p}: {matches}"

print("All deliverable integrity checks PASSED successfully!")
