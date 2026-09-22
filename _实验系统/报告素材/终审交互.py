from pathlib import Path
import json,hashlib,datetime,sys
import pythoncom
import win32com.client as w
A=Path(__file__).parent
sys.path.insert(0,str(A.parents[1]))
import 评测数据处理 as pipeline
tables=pipeline.load_core().extract(A.parents[1]/'_实验系统/telemetry/benchmark.sqlite')[0]
metrics=dict(tables)['评分与用量']
app=w.dynamic.Dispatch(pythoncom.CoCreateInstance('Excel.Application',None,pythoncom.CLSCTX_LOCAL_SERVER,pythoncom.IID_IDispatch))
book=None;out=[]
try:
 app.Visible=False;app.DisplayAlerts=False
 path=Path(sys.argv[1]) if len(sys.argv)>1 else A.parents[1]/'交付成果/项目2_AI评测分析.xlsx'
 book=app.Workbooks.Open(str(path),0,True)
 runs=book.Worksheets('运行记录').Range('A2:R16').Value2
 s=book.Worksheets('单AI分析')
 for ai in ['GPT','Grok','DeepSeek','GLM','Qwen','Gemini','Claude','Kimi']:
  for rnd in ['第一轮','第二轮']:
   s.Range('C3').Value2=ai;s.Range('G3').Value2=rnd;app.CalculateFullRebuild()
   r=next((r for r in runs if r[0]==ai and r[1]==rnd),None)
   selected=[m for m in metrics if m['AI']==ai and m['轮次']==rnd]
   expected=[]
   for kind in ['质量分项','可靠性分项']:
    items=[m for m in selected if m['指标类别']==kind]
    expected.append(sum(m['数值'] for m in items)+sum(m['数值'] for m in selected if kind=='质量分项' and m['指标类别']=='质量封顶') if items else '—')
   actual=[s.Range(c).Value2 for c in ['B5','E5']]
   assert actual==expected,(ai,rnd,expected,actual)
   assert ('仅作观察' in str(s.Range('J3').Value2))==(ai=='Gemini' and rnd=='第二轮')
   out.append({'AI':ai,'轮次':rnd,'Q/R':actual,'状态':s.Range('J3').Value2})
 for rnd in ['第一轮','第二轮']:
  ov=book.Worksheets('评测总览');ov.Range('C3').Value2=rnd;app.CalculateFullRebuild()
  for row in range(9,15):
   ai=ov.Cells(row,2).Value2;r=next(r for r in runs if r[0]==ai and r[1]==rnd)
   assert ov.Cells(row,3).Value2==r[13]
  ov.ExportAsFixedFormat(0,str(A/f'终审总览-{rnd}.pdf'))
 for ai,rnd in [('GPT','第二轮'),('Kimi','第一轮'),('Claude','第一轮')]:
  s.Range('C3').Value2=ai;s.Range('G3').Value2=rnd;app.CalculateFullRebuild();s.ExportAsFixedFormat(0,str(A/f'终审单AI-{ai}.pdf'))
 (path.parent/'待替换交互.json' if len(sys.argv)>1 else A/'终审交互.json').write_text(json.dumps({'时间':datetime.datetime.now().astimezone().isoformat(),'文件':str(path),'工作簿SHA256':hashlib.sha256(path.read_bytes()).hexdigest(),'案例':out,'总览轮次':['第一轮','第二轮']},ensure_ascii=False,indent=2),encoding='utf8');print('16种AI/轮次切换与2轮总览核对通过',flush=True)
 if len(sys.argv)>1:
  import time
  app.Visible=True;app.WindowState=-4137;s.Activate();s.Range('C3').Value2='Gemini';s.Range('G3').Value2='第二轮';app.CalculateFullRebuild();app.ActiveWindow.ScrollRow=1;app.ActiveWindow.ScrollColumn=1
  print(s.Range('J3').Value2,flush=True)
  time.sleep(50)
finally:
 if book is not None:book.Close(False)
 app.Quit()
