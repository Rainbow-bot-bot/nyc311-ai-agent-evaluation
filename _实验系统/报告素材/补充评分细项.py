"""在既有Excel中增加报告附表；只读引用评分与用量，不改评分。"""
from pathlib import Path
import json,time

A=Path(__file__).parent
CAPS={'A1':5,'A2':5,'A3':5,'A4':5,'B1':10,'B2':5,'B3':5,'B4':5,'C1':8,'C2':8,'C3':7,'C4':4,'C5':3,'D1':4,'D2':4,'D3':3,'D4':4,'E1':3,'E2':2,'E3':2,'E4':3,'R1':25,'R2':25,'R3':20,'R4':20,'R5':10}
AIS=['GPT','Grok','DeepSeek','GLM','Qwen','Gemini','Claude']

def apply_score_details(wb,export=False):
    try:ws=wb.Worksheets('评分细项')
    except Exception:ws=wb.Worksheets.Add(After=wb.Worksheets('分维度比较'));ws.Name='评分细项'
    print('开始排版',flush=True)
    wb.Application.Calculation=-4135;wb.Application.ScreenUpdating=False
    ws.UsedRange.UnMerge();ws.UsedRange.Clear();print('旧区域清理完成',flush=True)
    source=wb.Worksheets('评分与用量').UsedRange.Value2
    index={};names={}
    for n,r in enumerate(source[1:],2):
        code=str(r[3]).split(' ')[0]
        if code in CAPS:
            key=(r[0],r[1],code)
            if key in index:raise ValueError(f'重复评分 {key}')
            index[key]=(n,r[4]);names[code]=str(r[3]).split(' ',1)[1]
    # existing merged cells cleared above
    checks=[]
    for col,width in [('A',6),('B',24),('C',7),('D',9),('E',9),('F',8),('G',3),('H',6),('I',24),('J',7),('K',9),('L',9),('M',8)]:ws.Columns(col).ColumnWidth=width
    def color(s):return int(s[:2],16)+256*int(s[2:4],16)+65536*int(s[4:],16)
    blocks=[]
    for i,ai in enumerate(AIS):
        print(ai,flush=True)
        start=1+i*28;end=start+24
        rg=ws.Range(ws.Cells(start,1),ws.Cells(end,13))
        rg.Font.Name='微软雅黑';rg.Font.Size=12;rg.Font.Bold=False;rg.Font.Color=color('26323B')
        rg.Interior.Color=16777215;rg.RowHeight=20;rg.VerticalAlignment=-4108
        title=ws.Range(ws.Cells(start,1),ws.Cells(start,13));title.Merge();title.Value2=f'{ai}  |  评分细项'
        title.Font.Size=18;title.Font.Bold=True;title.RowHeight=32;title.Interior.Color=color('26323B');title.Font.Color=16777215
        for j,name in enumerate(AIS):
            cell=ws.Cells(start+1,1+j*2)
            ws.Hyperlinks.Add(Anchor=cell,Address='',SubAddress=f"'评分细项'!A{1+j*28}",TextToDisplay=name)
            cell.Value2=name
            cell.Font.Size=10;cell.Font.Color=color('16726A');cell.Font.Bold=name==ai
        for base in (1,8):
            head=ws.Range(ws.Cells(start+3,base),ws.Cells(start+3,base+5))
            head.Value2=(('编号','评分项','满分','第一轮','第二轮','分差'),)
            head.Interior.Color=color('E8EDF0');head.Font.Bold=True;head.RowHeight=24
            ws.Range(ws.Cells(start+3,base+2),ws.Cells(start+21,base+5)).HorizontalAlignment=-4108
        sections=[(1,4,'A  数据判断 · 20分','A'),(1,9,'B  证据计算 · 25分','B'),(1,14,'C  调查推进 · 30分','C'),(8,4,'D  最终交付 · 15分','D'),(8,9,'E  复杂度与范围 · 10分','E'),(8,14,'R  过程可靠性 · 100分','R')]
        locations={}
        for base,off,label,prefix in sections:
            band=ws.Range(ws.Cells(start+off,base),ws.Cells(start+off,base+5));band.Merge();band.Value2=label
            band.Interior.Color=color('E8F2F0');band.Font.Color=color('225E58');band.Font.Bold=True;band.RowHeight=24
            for j,code in enumerate(c for c in CAPS if c.startswith(prefix)):
                row=start+off+1+j;locations[code]=(row,base)
                line=ws.Range(ws.Cells(row,base),ws.Cells(row,base+5));line.Interior.Color=color('F5F7F8' if j%2 else 'FFFFFF')
                line.Borders(9).LineStyle=1;line.Borders(9).Color=color('E1E6E9')
                ws.Cells(row,base).Value2=code;ws.Cells(row,base).Font.Color=color('687780')
                ws.Cells(row,base+1).Value2=names[code];ws.Cells(row,base+2).Value2=CAPS[code]
                ws.Cells(row,base+2).Font.Color=color('687780')
                for col,rnd in ((base+3,'第一轮'),(base+4,'第二轮')):
                    hit=index.get((ai,rnd,code))
                    if hit:ws.Cells(row,col).Formula=f"='评分与用量'!E{hit[0]}"
                    checks.append((row,col,hit[1] if hit else None,ai,code,rnd))
                first=ws.Cells(row,base+3).Address;second=ws.Cells(row,base+4).Address
                delta=ws.Cells(row,base+5);delta.Formula=f'=IF(COUNT({first},{second})=2,{second}-{first},"")'
                delta.NumberFormat='+0;-0;"—"';delta.Font.Color=color('16726A')
        for base,prefix,label in ((1,'Q','成果质量 Q'),(8,'R','过程可靠性 R')):
            row=start+21;line=ws.Range(ws.Cells(row,base),ws.Cells(row,base+5))
            line.Interior.Color=color('E8EDF0');line.Font.Bold=True;line.RowHeight=28
            ws.Range(ws.Cells(row,base),ws.Cells(row,base+1)).Merge();ws.Cells(row,base).Value2=label;ws.Cells(row,base+2).Value2=100
            for shift,rnd in ((3,'第一轮'),(4,'第二轮')):
                codes=[c for c in CAPS if c.startswith('R')==(prefix=='R')]
                if (ai,rnd,codes[0]) in index:
                    refs=','.join(ws.Cells(locations[c][0],locations[c][1]+shift).Address for c in codes)
                    ws.Cells(row,base+shift).Formula=f'=SUM({refs})'
            first=ws.Cells(row,base+3).Address;second=ws.Cells(row,base+4).Address
            ws.Cells(row,base+5).Formula=f'=IF(COUNT({first},{second})=2,{second}-{first},"")'
            ws.Cells(row,base+5).NumberFormat='+0;-0;"—"'
        note=ws.Range(ws.Cells(start+23,1),ws.Cells(end,13));note.Merge();note.WrapText=True;note.Font.Size=11;note.Font.Color=color('52616B')
        note.Value2=('Claude第一轮由Claude主体完成、Grok收尾，81分评价混合成品；R不发布，第二轮未形成独立评分。' if ai=='Claude' else 'Q为成果质量，R为过程可靠性，两项各100分、分别评价。分差＝第二轮－第一轮；得分理由见报告对应AI章节。')+' 空白表示未评分。'
        blocks.append({'ai':ai,'range':f'A{start}:M{end}','file':f'{ai}细项.png'})
    ws.Activate();wb.Application.ActiveWindow.DisplayGridlines=False;wb.Application.ActiveWindow.Zoom=95
    ws.Range('A1').Select();wb.Application.ActiveWindow.ScrollRow=1;wb.Application.ActiveWindow.ScrollColumn=1
    ws.PageSetup.PrintArea=f'$A$1:$M${end}';ws.PageSetup.Zoom=False;ws.PageSetup.FitToPagesWide=1;ws.PageSetup.FitToPagesTall=False
    ws.ResetAllPageBreaks()
    for i in range(1,len(AIS)):ws.HPageBreaks.Add(Before=ws.Cells(1+i*28,1))
    wb.Application.Calculation=-4105;wb.Application.ScreenUpdating=True
    wb.Application.CalculateFullRebuild()
    for row,col,expected,ai,code,rnd in checks:assert ws.Cells(row,col).Value2==expected,(ai,code,rnd)
    if export:
        for block in blocks:
            rg=ws.Range(block['range']);rg.Select();rg.CopyPicture(1,2);time.sleep(.25)
            co=ws.ChartObjects().Add(0,0,rg.Width,rg.Height)
            try:
                co.Activate();co.Chart.Paste();time.sleep(.25)
                assert co.Chart.Export(str(A/block['file']),'PNG')
            finally:co.Delete()
        (A/'评分细项来源.json').write_text(json.dumps(blocks,ensure_ascii=False,indent=2),encoding='utf-8')
    ws.Range('A1').Select();wb.Application.ActiveWindow.ScrollRow=1
    return blocks

if __name__=='__main__':
    import win32com.client
    app=win32com.client.DispatchEx('Excel.Application');app.Visible=False;app.DisplayAlerts=False
    wb=None
    try:
        wb=app.Workbooks.Open(str(A.parents[1]/'交付成果/项目2_AI评测分析.xlsx'))
        if wb.ReadOnly:raise RuntimeError('正式Excel为只读')
        print(apply_score_details(wb,True));wb.Save();wb.Close(False);wb=None
    finally:
        if wb is not None:wb.Close(False)
        app.Quit()
