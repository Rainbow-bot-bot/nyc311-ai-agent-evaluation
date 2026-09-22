"""步骤8（续）：用 Excel 打开最终文件并导出 PDF，用于人工目视复核排版是否可读。
仅为自检用，产出 _预览.pdf 可在复核后删除。"""
import win32com.client as win32
import pathlib, os

BASE = pathlib.Path(r"D:\项目2\Claude")
src = BASE / "最终成果.xlsx"
pdf = BASE / "_预览.pdf"
if pdf.exists():
    pdf.unlink()

app = win32.Dispatch("Excel.Application")
app.Visible = False
app.DisplayAlerts = False
try:
    wb = app.Workbooks.Open(str(src), ReadOnly=True)
    print("Excel 成功打开，工作表:", wb.Sheets.Count)
    for i in range(1, wb.Sheets.Count + 1):
        sh = wb.Sheets(i)
        sh.PageSetup.Orientation = 2          # 横向
        sh.PageSetup.Zoom = False
        sh.PageSetup.FitToPagesWide = 1
        sh.PageSetup.FitToPagesTall = False
        print(f"  {i}. {sh.Name}  已用区域={sh.UsedRange.Address}  图表={sh.ChartObjects().Count}")
    wb.ExportAsFixedFormat(0, str(pdf))
    wb.Close(False)
    print("PDF 预览已导出:", pdf, os.path.getsize(pdf), "bytes")
finally:
    app.Quit()
