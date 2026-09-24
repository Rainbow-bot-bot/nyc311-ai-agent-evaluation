
# 历史源代码保留供回查；执行和导入均在写入之前停止。
if __name__ == "__main__":
    raise SystemExit('历史Q/R生成或旧版式脚本已停用，禁止改写现行成果。当前维护见资料/复现说明.md；评分同步用评测数据处理.py --sync-scores，验收用报告素材/终审.py。')
raise RuntimeError('历史Q/R生成或旧版式脚本已停用，禁止改写现行成果。当前维护见资料/复现说明.md；评分同步用评测数据处理.py --sync-scores，验收用报告素材/终审.py。')

from pathlib import Path
import json,re,sys
import pythoncom
from win32com.client import dynamic
A=Path(__file__).parent;R=A.parents[1]

def apply(wb):
    previous_autofill=wb.Application.AutoCorrect.AutoFillFormulasInLists
    try:
        wb.Application.AutoCorrect.AutoFillFormulasInLists=False
        sys.path.insert(0,str(R))
        import 评测数据处理 as pipeline
        data=dict(pipeline.load_core().extract(R/'_实验系统/telemetry/benchmark.sqlite')[0])
        for name in ('运行记录','评分与用量'):
            ws=wb.Worksheets(name);rows=data[name];keys=list(rows[0]);n=len(rows)
            ws.Range(ws.Cells(2,1),ws.Cells(n+1,len(keys))).Value2=tuple(tuple(r.get(k) for k in keys) for r in rows)
            table=ws.ListObjects(1);table.Resize(ws.Range(ws.Cells(1,1),ws.Cells(n+1,table.Range.Columns.Count)))
        ws=wb.Worksheets('运行记录')
        for row in range(2,len(data['运行记录'])+2):
            ws.Cells(row,14).Formula=f'=IF(COUNTIFS(评分与用量!$H:$H,J{row},评分与用量!$C:$C,"质量分项")=21,SUMIFS(评分与用量!$E:$E,评分与用量!$H:$H,J{row},评分与用量!$C:$C,"质量分项")+SUMIFS(评分与用量!$E:$E,评分与用量!$H:$H,J{row},评分与用量!$C:$C,"质量封顶"),"")'
        ov=wb.Worksheets('评测总览');ov.Range('C3').Value2='第二轮'
        for addr,formula in {'B5':'=COUNT(C9:C13)','E5':'=AVERAGE(C9:C13)','H5':'=AVERAGE(D9:D13)','K5':'=MIN(E9:E13)'}.items():ov.Range(addr).Formula=formula
        ov.Range('F3').Value2='五个配对系统；Gemini第二轮仅作观察'
        ov.Range('N14').Formula='=IF(C3="第二轮","观察","已完成")'
        ov.Range('B19').Value2='两轮分数变化 · Gemini不计配对汇总'
        ov.Range('B29').Value2='五个配对系统 · 分项变化与封顶调整'
        ov.Range('M30').Value2='分项增量占比'
        for row in range(31,36):
            for col in ('G','I'):
                f=ov.Range(f'{col}{row}').Formula
                f=re.sub(r'\+SUMIFS\([^)]*"Gemini"[^)]*\)','',f).replace('/6','/5')
                ov.Range(f'{col}{row}').Formula=f
        ov.Range('B36').Value2='分项合计'
        for row,label in [(37,'封顶调整'),(38,'最终Q')]:
            rg=ov.Range(f'B{row}:N{row}');rg.Clear();rg.Font.Name='等线';rg.Font.Size=11;rg.Font.Bold=row==38;rg.Borders.LineStyle=1;rg.Borders.Color=13882323;rg.RowHeight=19
            for span in ('B:D','E:F','G:H','I:J','K:L','M:N'):
                left,right=span.split(':');ov.Range(f'{left}{row}:{right}{row}').Merge()
            ov.Range(f'B{row}').Value2=label
        ov.Range('G37').Value2=0;ov.Range('I37').Formula='=SUMIFS(评分与用量!E:E,评分与用量!C:C,"质量封顶",评分与用量!B:B,"第二轮")/5'
        ov.Range('K37').Formula='=I37-G37'
        for col in ('G','I','K'):ov.Range(f'{col}38').Formula=f'=SUM({col}36:{col}37)'
        ov.Range('G31:N38').HorizontalAlignment=-4108;ov.Range('G31:N38').NumberFormat='0.0';ov.Range('M31:M36').NumberFormat='0.0%'
        ov.Range('S29').Value2='GPT、Grok、DeepSeek、GLM、Qwen五个系统等权。Gemini第二轮Gate不通过，保留观察记录。'
        ov.Range('S31').Value2='分项增量占比以封顶前增分为分母；Qwen封顶调整另列，不拆摊至质量维度。'
        ov.Range('S33').Value2='最终Q平均增加0.6分，中位增加6分。Qwen分项合计90，主线时间窗错误，封顶为59。'
        ov.Range('P35:AD38').UnMerge()
        ov.Range('P35:R38').Merge();ov.Range('S35:AD38').Merge()
        ov.Range('P35').Value2='本次裁定'
        ov.Range('P35:R38').Interior.Color=4210752;ov.Range('P35:R38').Font.Color=16777215
        ov.Range('S35:AD38').Interior.Color=16777215;ov.Range('P35:AD38').Borders.LineStyle=1;ov.Range('P35:AD38').Borders.Color=13882323
        ov.Range('P35:AD38').VerticalAlignment=-4108;ov.Range('P35:AD38').WrapText=True
        ov.Range('S35').Value2='GPT的C2、C4重新评定为满分，Q为100；原评分及本次裁定均可查。'
        ov.Range('E18').Value2='Gemini第二轮为观察评分。Qwen第二轮分项90、封顶59；费用为API目录价估算。'
        ov.Range('B30:N38').RowHeight=19
        ov.PageSetup.PrintArea='$A$1:$AD$38'
        dim=wb.Worksheets('分维度比较')
        for row in range(15,23):
            dim.Cells(row,7).Formula=f'=IF(COUNT(B{row}:F{row})=0,"",IF(A{row}="Qwen",MIN(SUM(B{row}:F{row}),59),SUM(B{row}:F{row})))'
        dim.Range('H19').Value2='分项90；封顶59'
        for addr in ('H20','H44'):dim.Range(addr).Value2='观察；Gate不通过'
        dim.Range('F109').Value2='Q分项90，封顶59'
        dim.Range('F123').Value2='第二轮仅作观察'
        details=wb.Worksheets('评分细项')
        details.Range('E134').Formula='=MIN(SUM(E118:E121,E123:E126,E128:E132,L118:L121,L123:L126),59)'
        details.Range('A136').Value2='Qwen第二轮分项合计90；主要发现使用11个月与9个月比较，按冻结规则封顶59。原分项保留，调整为−31。'
        details.Range('A24').Value2='GPT第二轮C2改为8、C4改为4，Q由98改为100。峰日明细、描述排除与固定构成检查支持这两项满分，原扣分缺少对应依据。'
        single=wb.Worksheets('单AI分析')
        status=single.Range('J3').Formula
        suffix='&IF(AND($C$3="Gemini",$G$3="第二轮"),"（仅作观察）","")'
        if suffix not in status:single.Range('J3').Formula=status+suffix
        single.Range('E26').Value2='Qwen第二轮分项90，封顶59；Gemini第二轮仅作观察。分项与总分的差额见评分与用量“质量封顶”。'
        checks=wb.Worksheets('核验记录');headers=list(checks.Range('A1:W1').Value2[0])
        for field in ('原始文字值','对照文字值'):
            col=headers.index(field)+1
            for row,record in enumerate(data['核验记录'],2):
                value=record.get(field)
                if value is not None:
                    cell=checks.Cells(row,col);cell.NumberFormat='@';cell.Value2=str(value)
        fields=wb.Worksheets('字段说明')
        for row in range(2,fields.UsedRange.Rows.Count+1):
            if fields.Cells(row,2).Value2=='成果质量分':fields.Cells(row,3).Value2='A-E质量分项合计＋质量封顶调整'
        raw=wb.Worksheets('评分与用量');end=len(data['评分与用量'])+1
        raw.Rows(end-1).Copy();raw.Rows(end).PasteSpecial(-4122);wb.Application.CutCopyMode=False
        for sheet in (raw,ws):
            for table in sheet.ListObjects:
                for edge in (7,8,9,10):table.Range.Borders(edge).LineStyle=1;table.Range.Borders(edge).Weight=2
        from 语言格式修正 import apply as apply_layout
        apply_layout(wb,data,pipeline.load_core())
        wb.Application.CalculateFullRebuild()
        assert ov.Range('E5').Value2==89.2
        assert [dim.Cells(row,7).Value2 for row in range(15,21)]==[100,98,95,94,59,78]
        assert abs(ov.Range('K38').Value2-.6)<1e-8
    finally:
        wb.Application.AutoCorrect.AutoFillFormulasInLists=previous_autofill

if __name__=='__main__':
    app=dynamic.Dispatch(pythoncom.CoCreateInstance('Excel.Application',None,pythoncom.CLSCTX_LOCAL_SERVER,pythoncom.IID_IDispatch));wb=None
    try:
        app.Visible=False;app.DisplayAlerts=False
        wb=app.Workbooks.Open(str(R/'交付成果/项目2_AI评测分析.xlsx'))
        if wb.ReadOnly:raise RuntimeError('正式工作簿被占用')
        app.Calculation=-4135
        apply(wb);app.Calculation=-4105
        wb.Save();print('裁定已更新；五系统Q均值89.2，增量0.6')
    finally:
        if wb is not None:wb.Close(False)
        app.Quit()
