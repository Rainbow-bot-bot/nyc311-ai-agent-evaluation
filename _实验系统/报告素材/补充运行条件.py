from pathlib import Path
import time,json
A=Path(__file__).parent
ROWS=[
 ('GPT','gpt-6-astra','gpt-6-astra','中（medium）','Codex'),
 ('Grok','grok-4.6','grok-4.6','中（medium）','Cursor'),
 ('DeepSeek','deepseek-v4.1-flash','deepseek-v4.1-flash','high','Claude Code'),
 ('GLM','glm-5.3','glm-5.3','high','Claude Code / Brain Router'),
 ('Qwen','qwen3.8-max-0902','qwen3.8-max-0902','high','Claude Code'),
 ('Gemini','gemini-3.8-flash','Gemini（子版本未核实）','中（medium）','Antigravity'),
 ('Claude','Claude Opus 5 → Grok','无独立运行','中（medium）','Cursor'),
 ('Kimi','Kimi（子版本未核实）','kimi-k3','high','Claude Code'),
]
def apply_conditions(wb,export=False):
 s=wb.Worksheets('字段说明')
 try:s.ListObjects('tblModelConditions').Unlist()
 except Exception:pass
 rg=s.Range('H1:L14');rg.UnMerge();rg.Clear()
 s.Columns('H').ColumnWidth=13;s.Columns('I:J').ColumnWidth=30;s.Columns('K').ColumnWidth=19;s.Columns('L').ColumnWidth=30
 rg.Font.Name='微软雅黑';rg.Font.Size=12;rg.Font.Bold=False;rg.Font.Color=0;rg.Interior.Color=16777215
 rg.WrapText=True;rg.VerticalAlignment=-4108;rg.RowHeight=30;rg.Borders.LineStyle=1;rg.Borders.Color=14277081
 s.Range('H1:L1').Merge();s.Range('H1').Value2='参评模型、思考强度与 Harness';s.Range('H1:L2').Interior.Color=4210752;s.Range('H1:L2').Font.Color=16777215;s.Range('H1:L2').Font.Bold=True
 s.Range('H2:L2').Value2=(('AI','第一轮模型','第二轮模型','思考强度','Harness（执行环境）'),)
 s.Range('H3:L10').Value2=tuple(ROWS)
 for r in (4,6,8,10):s.Range(f'H{r}:L{r}').Interior.Color=15921906
 model=s.ListObjects.Add(1,s.Range('H2:L10'),None,1);model.Name='tblModelConditions'
 s.Range('H2:L10').ClearFormats();model.TableStyle='TableStyleMedium1';model.ShowTableStyleRowStripes=True
 s.Range('H2:L10').Font.Name='微软雅黑';s.Range('H2:L10').Font.Size=11;s.Range('H2:L10').WrapText=True;s.Range('H2:L10').VerticalAlignment=-4108
 s.Range('H11:L11').Merge();s.Range('H11').Value2='运行说明';s.Range('H11:L11').Interior.Color=4210752;s.Range('H11:L11').Font.Color=16777215;s.Range('H11:L11').Font.Bold=True
 notes=[('思考强度','海外模型设为中，国内模型设为Claude Code中的high。'),('执行环境','Harness指模型使用的执行环境；GLM经Brain Router路由。同名强度设置不代表跨平台相同思考预算。'),('版本与缺失','模型版本取自运行记录。Claude第二轮未运行；Gemini第二轮、Kimi第一轮的具体子版本未核实。')]
 for r,(label,note) in enumerate(notes,12):
  s.Cells(r,8).Value2=label;s.Range(f'I{r}:L{r}').Merge();s.Cells(r,9).Value2=note
  s.Range(f'H{r}:L{r}').Interior.Color=15921906 if r%2==0 else 16777215
  s.Cells(r,8).Font.Bold=True;s.Rows(r).RowHeight=42
 rg.Font.Size=11
 if export:
  s.Activate();rg.CopyPicture(1,2);time.sleep(.3);co=s.ChartObjects().Add(0,0,rg.Width,rg.Height)
  try:
   co.Activate();co.Chart.Paste();time.sleep(.3);assert co.Chart.Export(str(A/'运行条件.png'),'PNG')
  finally:co.Delete()
 return ROWS
if __name__=='__main__':
 import win32com.client as w
 app=w.DispatchEx('Excel.Application');app.Visible=False;app.DisplayAlerts=False
 try:
  wb=app.Workbooks.Open(str(A.parents[1]/'交付成果/项目2_AI评测分析.xlsx'))
  assert not wb.ReadOnly
  apply_conditions(wb,True);wb.Save();wb.Close(False)
 finally:app.Quit()
