"""只读检查现行Excel的全部有限选择器；期望分数来自S分项。"""
from pathlib import Path
import sys,json,hashlib,gc,time
from collections import defaultdict
from decimal import Decimal,ROUND_HALF_UP
ROOT=Path(__file__).resolve().parents[2];sys.path.insert(0,str(ROOT))
import 评测数据处理 as pipeline

def main():
    import win32com.client,win32process,psutil
    path=Path(sys.argv[1]).resolve() if len(sys.argv)>1 else ROOT/'交付成果/项目2_AI评测分析.xlsx'
    entries=pipeline.read_current_scores(path)
    totals=defaultdict(float);dims=defaultdict(float)
    for entry in entries.values():
        s=entry['评分原文'];totals[s['AI'],s['轮次']]+=s['得分'];dims[s['AI'],s['轮次'],s['条款'][0]]+=s['得分']
    def want(ai,rd):return totals.get((ai,rd))
    def same(actual,expected):
        if expected is None:return actual in (None,'','—','–')
        return isinstance(actual,(int,float)) and abs(actual-expected)<1e-8
    systems=['GPT','Grok','DeepSeek','GLM','Qwen','Gemini','Claude','Kimi']
    before=hashlib.sha256(path.read_bytes()).hexdigest()
    app=win32com.client.DispatchEx('Excel.Application');pid=win32process.GetWindowThreadProcessId(app.Hwnd)[1]
    book=score=single=overview=dimension=None
    try:
        app.Visible=False;app.DisplayAlerts=False;book=app.Workbooks.Open(str(path),0,True)
        score=book.Worksheets('评分细项');single=book.Worksheets('单AI分析');overview=book.Worksheets('评测总览');dimension=book.Worksheets('分维度比较')
        for ai in systems:
            score.Range('B2').Value2=ai;app.CalculateFull()
            assert same(score.Range('D22').Value2,want(ai,1)),ai
            assert same(score.Range('E22').Value2,want(ai,2)),ai
            single.Range('C3').Value2=ai
            for rd,label in [(1,'第一轮'),(2,'第二轮')]:
                single.Range('G3').Value2=label;app.CalculateFull()
                assert same(single.Range('B5').Value2,want(ai,rd)),(ai,rd,'本轮')
                assert same(single.Range('E5').Value2,want(ai,3-rd)),(ai,rd,'另一轮')
        runs=book.Worksheets('运行记录').Range('A2:R16').Value2
        for rd,label in [(1,'第一轮'),(2,'第二轮')]:
            overview.Range('C3').Value2=label;app.CalculateFull()
            for row in range(9,15):
                ai=overview.Cells(row,2).Value2
                assert same(overview.Cells(row,3).Value2,want(ai,1)),('总览首轮',ai)
                assert same(overview.Cells(row,4).Value2,want(ai,2)),('总览次轮',ai)
                record=next(r for r in runs if r[0]==ai and r[1]==label)
                assert same(overview.Cells(row,5).Value2,record[7]/60 if record[7] is not None else None),('总览用时',ai,rd)
        maxima=dict(zip('ABCDEF',[15,20,20,25,10,10]))
        for label in ['第一轮','第二轮','两轮变化']:
            dimension.Range('B3').Value2=label
            for metric in ['得分','得分率']:
                dimension.Range('F3').Value2=metric;app.CalculateFull()
                for row,ai in enumerate(systems,6):
                    for col,dim in enumerate('ABCDEF',2):
                        x=dims.get((ai,1,dim));y=dims.get((ai,2,dim))
                        expected=x if label=='第一轮' else y if label=='第二轮' else (y-x if x is not None and y is not None else None)
                        actual=dimension.Cells(row,col).Value2
                        if metric=='得分' or expected is None:assert same(actual,expected),(ai,label,metric,dim,actual,expected)
                        else:
                            percent=float((Decimal(str(expected))/Decimal(maxima[dim])*100).quantize(Decimal('0.1'),rounding=ROUND_HALF_UP))
                            assert abs(float(str(actual).replace('%',''))-percent)<1e-8,(ai,label,metric,dim,actual,percent)
        print(json.dumps({'评分细项':8,'单AI分析':16,'总览':2,'分维度比较':6,'分维度核对':'8个系统×6维','评分来源':'正式Excel S分项','工作簿SHA256':before},ensure_ascii=False))
    finally:
        if book is not None:book.Close(False)
        score=single=overview=dimension=book=None
        app.Quit();app=None;gc.collect()
        for _ in range(20):
            if not psutil.pid_exists(pid):break
            time.sleep(.2)
        if psutil.pid_exists(pid):
            # Only this DispatchEx instance, after its read-only workbook was closed.
            process=psutil.Process(pid);process.terminate();process.wait(timeout=5)
        if psutil.pid_exists(pid):raise RuntimeError(f'本次Excel实例尚未退出：{pid}')
        print(f'Excel实例已退出：{pid}')
    assert hashlib.sha256(path.read_bytes()).hexdigest()==before

if __name__=='__main__':main()
