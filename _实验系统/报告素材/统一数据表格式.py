
# 历史源代码保留供回查；执行和导入均在写入之前停止。
if __name__ == "__main__":
    raise SystemExit('历史Q/R生成或旧版式脚本已停用，禁止改写现行成果。当前维护见资料/复现说明.md；评分同步用评测数据处理.py --sync-scores，验收用报告素材/终审.py。')
raise RuntimeError('历史Q/R生成或旧版式脚本已停用，禁止改写现行成果。当前维护见资料/复现说明.md；评分同步用评测数据处理.py --sync-scores，验收用报告素材/终审.py。')

from pathlib import Path
import json,time
A=Path(__file__).parent
SHEETS=['运行记录','评分与用量','API计价','行为记录','核验记录','字段说明']
def apply_data_tables(wb):
 app=wb.Application; oldcalc=app.Calculation;app.Calculation=-4135;app.ScreenUpdating=False
 active=wb.ActiveSheet.Name;results=[]
 try:
  for name in SHEETS:
   ws=wb.Worksheets(name);table=ws.ListObjects(1);rg=table.Range
   before=rg.Formula; formats=[ws.Cells(2,c).NumberFormat for c in range(1,rg.Columns.Count+1)]
   oldheights=[ws.Rows(r).RowHeight for r in range(1,14)] if name=='字段说明' else []
   rg.ClearFormats();rg.Font.Name='微软雅黑';rg.Font.Size=11;rg.Font.Bold=False
   rg.WrapText=True;rg.VerticalAlignment=-4108;rg.HorizontalAlignment=-4131
   table.TableStyle='TableStyleMedium1';table.ShowTableStyleRowStripes=True;table.ShowTableStyleColumnStripes=False;table.ShowTableStyleFirstColumn=False;table.ShowTableStyleLastColumn=False;table.ShowAutoFilter=True
   for c,header in enumerate(before[0],1):
    header=str(header)
    width=18
    if header=='AI':width=13
    elif header in ('轮次','单位','计价档'):width=12
    elif '编号' in header:width=43 if header in ('参与编号','证据编号','行为编号') else 36
    elif header in ('动作目的','核验项目','原始文字值','对照文字值','含义','允许值','空值规则'):width=48
    elif header in ('模型','API计价说明','指标','数值口径','判断来源','运行工具','字段'):width=30
    elif header in ('工具名称','数据版本','执行阶段','表名'):width=24
    elif len(header)>7:width=22
    ws.Columns(c).ColumnWidth=width
    ws.Range(ws.Cells(2,c),ws.Cells(rg.Rows.Count,c)).NumberFormat=formats[c-1]
   table.HeaderRowRange.Font.Bold=True;table.HeaderRowRange.RowHeight=36
   table.DataBodyRange.Rows.AutoFit()
   for r in range(2,rg.Rows.Count+1):
    if ws.Rows(r).RowHeight<26:ws.Rows(r).RowHeight=26
   if name=='字段说明':
    model=ws.Range('H2:L10');model.ClearFormats();model.Font.Name='微软雅黑';model.Font.Size=11;model.WrapText=True;model.VerticalAlignment=-4108
    try:mt=ws.ListObjects('tblModelConditions')
    except Exception:mt=ws.ListObjects.Add(1,model,None,1);mt.Name='tblModelConditions'
    mt.TableStyle='TableStyleMedium1';mt.ShowTableStyleRowStripes=True
    ws.Range('H1:L1').Font.Name='微软雅黑';ws.Range('H1:L1').Font.Size=11
    ws.Range('H11:L13').Font.Name='微软雅黑';ws.Range('H11:L13').Font.Size=11
    for r in range(11,14):ws.Rows(r).RowHeight=max(ws.Rows(r).RowHeight,40)
   ws.Activate();app.ActiveWindow.DisplayGridlines=True;app.ActiveWindow.Zoom=90
   app.ActiveWindow.FreezePanes=False;app.ActiveWindow.SplitColumn=2;app.ActiveWindow.SplitRow=1;app.ActiveWindow.FreezePanes=True
   app.ActiveWindow.ScrollRow=1;app.ActiveWindow.ScrollColumn=1;ws.Range('A1').Select()
   assert rg.Formula==before,name
   results.append({'sheet':name,'rows':table.ListRows.Count,'native_table':table.Name,'style':str(table.TableStyle),'content_unchanged':True})
   print(name,'完成',flush=True)
  apply_comparison_tables(wb)
  wb.Worksheets(active).Activate()
 finally:app.Calculation=oldcalc;app.ScreenUpdating=True
 return results
def apply_comparison_tables(wb):
 ws=wb.Worksheets('分维度比较');before=ws.UsedRange.Formula
 specs=[(2,10,8),(14,22,8),(26,34,8),(38,46,8)]+[(52+14*i,62+14*i,6) for i in range(8)]
 for i,(head,end,cols) in enumerate(specs,1):
  rg=ws.Range(ws.Cells(head,1),ws.Cells(end,cols));rg.UnMerge();rg.ClearFormats()
  name=f'tblDimension{i}'
  try:table=ws.ListObjects(name)
  except Exception:table=ws.ListObjects.Add(1,rg,None,1);table.Name=name
  table.TableStyle='TableStyleMedium1';table.ShowTableStyleRowStripes=True
  rg.Font.Name='微软雅黑';rg.Font.Size=11;rg.Font.Bold=False;rg.WrapText=True;rg.VerticalAlignment=-4108
  table.HeaderRowRange.Font.Bold=True;table.HeaderRowRange.RowHeight=36
  table.DataBodyRange.Rows.AutoFit()
  for r in range(head+1,end+1):
   if ws.Rows(r).RowHeight<26:ws.Rows(r).RowHeight=26
  oldtitle=ws.Range(ws.Cells(head-1,1),ws.Cells(head-1,8));label=ws.Cells(head-1,1).Value2;oldtitle.UnMerge();oldtitle.ClearFormats()
  title=ws.Range(ws.Cells(head-1,1),ws.Cells(head-1,cols));title.Merge();title.Value2=label;title.Interior.Color=4210752;title.Font.Color=16777215;title.Font.Name='微软雅黑';title.Font.Size=12;title.Font.Bold=True;title.RowHeight=28
  if cols==6:ws.Range(ws.Cells(head,7),ws.Cells(end,8)).ClearFormats()
 # Native headers normalize line breaks; underlying data and formulas must remain intact.
 for head,end,cols in specs:
  for r in range(head+1,end+1):
   for c in range(1,cols+1):assert ws.Cells(r,c).Formula==before[r-1][c-1],(r,c)
 ws.Activate();wb.Application.ActiveWindow.Zoom=90;wb.Application.ActiveWindow.DisplayGridlines=True
 ws.Range('A1').Select();wb.Application.ActiveWindow.ScrollRow=1;wb.Application.ActiveWindow.ScrollColumn=1
 apply_table_edges(wb)
 print('分维度比较12张原生表格，得分和公式保持一致',flush=True)

def apply_table_edges(wb):
 count=0
 for name in ['分维度比较']+SHEETS:
  ws=wb.Worksheets(name)
  for table in ws.ListObjects:
   for edge in (7,8,9,10):
    border=table.Range.Borders(edge);border.LineStyle=1;border.Weight=2;border.Color=0
   count+=1
 return count

if __name__=='__main__':
 import win32com.client as w
 app=w.DispatchEx('Excel.Application');app.Visible=True;app.DisplayAlerts=False
 wb=app.Workbooks.Open(str(A.parents[1]/'交付成果/项目2_AI评测分析.xlsx'))
 if wb.ReadOnly:raise RuntimeError('目标工作簿只读')
 result=apply_data_tables(wb);wb.Save()
 (A/'数据表格式验收.json').write_text(json.dumps(result,ensure_ascii=False,indent=2),encoding='utf-8')
 # The model reference shares row heights with the field dictionary.
 ws=wb.Worksheets('字段说明');ws.Activate();rg=ws.Range('H1:L14');rg.CopyPicture(1,2);time.sleep(.3)
 co=ws.ChartObjects().Add(0,0,rg.Width,rg.Height)
 try:co.Activate();co.Chart.Paste();time.sleep(.3);co.Chart.Export(str(A/'运行条件.png'),'PNG')
 finally:co.Delete()
 wb.Worksheets('运行记录').Activate();wb.Save()
