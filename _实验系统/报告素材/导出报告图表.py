"""只读导出当前Excel中的报告图表，不重写已审定布局。"""

# 历史源代码保留供回查；执行和导入均在写入之前停止。
if __name__ == "__main__":
    raise SystemExit('历史Q/R生成或旧版式脚本已停用，禁止改写现行成果。当前维护见资料/复现说明.md；评分同步用评测数据处理.py --sync-scores，验收用报告素材/终审.py。')
raise RuntimeError('历史Q/R生成或旧版式脚本已停用，禁止改写现行成果。当前维护见资料/复现说明.md；评分同步用评测数据处理.py --sync-scores，验收用报告素材/终审.py。')

from pathlib import Path
import json,re,time,sys
import win32com.client as w
A=Path(__file__).parent
def export_report_assets(destination=A):
 destination=Path(destination);destination.mkdir(exist_ok=True)
 items=json.loads((A/'图表来源.json').read_text(encoding='utf8'))
 app=w.DispatchEx('Excel.Application');app.Visible=True;app.DisplayAlerts=False;wb=None
 old_errors=app.ErrorCheckingOptions.BackgroundChecking
 try:
  wb=app.Workbooks.Open(str(A.parents[1]/'交付成果/项目2_AI评测分析.xlsx'),0,True)
  app.ErrorCheckingOptions.BackgroundChecking=False
  # Only this unsaved export view loses interface controls; formal Excel retains them.
  for ws in wb.Worksheets:
   for index in range(ws.ListObjects.Count,0,-1):ws.ListObjects(index).Unlist()
  dim=wb.Worksheets('分维度比较');dim.Range('A1:H46').Font.Size=15
  for start in (1,13,25,37):
   dim.Rows(start).RowHeight=30;dim.Rows(start+1).RowHeight=42
   dim.Range(f'A{start+2}:H{start+9}').RowHeight=28
   if start in (25,37):dim.Range(f'A{start+1}:H{start+1}').Font.Size=12
  detail=wb.Worksheets('评分细项')
  for start in range(1,170,28):
   detail.Rows(start+1).Hidden=True;detail.Rows(start+2).Hidden=True
   detail.Range(f'A{start+3}:M{start+22}').Font.Size=18
   detail.Range(f'A{start+3}:M{start+22}').RowHeight=30
   detail.Range(f'A{start+3}:M{start+3}').Font.Size=14
   detail.Range(f'A{start+23}:M{start+23}').Font.Size=14
   detail.Rows(start+23).RowHeight=44
  wb.Worksheets('评测总览').Range('C3').Value2='第二轮';app.CalculateFullRebuild()
  for item in items:
   if len(sys.argv)>1 and item['file'] not in sys.argv[1:]:continue
   print(item['file'],flush=True)
   source=item['source'];sheet=wb.Worksheets(source.split()[0]);sheet.Activate()
   chart=re.search(r'图表(\d+)',source)
   if chart:assert sheet.ChartObjects(int(chart[1])).Chart.Export(str(destination/item['file']),'PNG');continue
   address=re.search(r'[A-Z]+\d+:[A-Z]+\d+',source)[0];rg=sheet.Range(address)
   sheet.Range('AZ1').Select()
   for attempt in range(2):
    try:
     rg.CopyPicture(1,2);break
    except Exception:
     if attempt:raise
     time.sleep(.5)
   time.sleep(.3);co=sheet.ChartObjects().Add(0,0,rg.Width,rg.Height)
   try:
    co.Activate();co.Chart.Paste();time.sleep(.2);assert co.Chart.Export(str(destination/item['file']),'PNG')
   finally:co.Delete()
 finally:
  if wb is not None:wb.Close(False)
  app.ErrorCheckingOptions.BackgroundChecking=old_errors
  app.Quit()
if __name__=='__main__':export_report_assets()
