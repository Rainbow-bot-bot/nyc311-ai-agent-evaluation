"""正式工作簿的局部排版修正；由最终裁定生成步骤调用。"""
def apply(wb,data,core):
    fields=wb.Worksheets('字段说明')
    dictionary={(r['表名'],r['字段']):r['允许值'] for r in core.build_field_dictionary_rows(list(data.items()))}
    for row in range(2,fields.UsedRange.Rows.Count+1):
        key=(fields.Cells(row,1).Value2,fields.Cells(row,2).Value2)
        if key in dictionary:fields.Cells(row,4).Value2=dictionary[key]
    dim=wb.Worksheets('分维度比较')
    for start in range(51,150,14):
        for table in dim.ListObjects:
            if table.Range.Row==start+1:table.Resize(dim.Range(f'A{start+1}:E{start+11}'))
        dim.Range(f'F{start+1}:F{start+11}').Clear()
        dim.Range(f'A{start}:H{start}').UnMerge();dim.Range(f'A{start}:E{start}').Merge()
        dim.Range(f'F{start}:H{start}').Clear()
    notes={107:'第二轮Q分项合计90，主线时间窗错误，封顶59；R为95。',121:'第二轮Q 78、R 81仅作观察，不进入正式排名与两轮配对汇总。',135:'第一轮Q 81评价Claude与Grok的混合成品，R不评分；第二轮未独立运行。',149:'两轮均未形成完整交付，Q/R及两轮变化留空。'}
    for start,note in notes.items():
        dim.Range(f'F{start+2}:F{start+11}').ClearContents()
        rg=dim.Range(f'A{start+12}:F{start+12}');rg.UnMerge()
        rg=dim.Range(f'A{start+12}:E{start+12}');rg.Merge();rg.Value2=note
        rg.Font.Size=11;rg.Font.Bold=False;rg.WrapText=True;rg.VerticalAlignment=-4108;rg.RowHeight=32
        rg.Borders.LineStyle=1;rg.Borders.Color=13882323
    detail=wb.Worksheets('评分细项')
    detail.Range('A164').Value2='Gemini第二轮Q 78、R 81为观察评分；不进入正式排名与两轮配对汇总。第一轮Q/R保留。'
    detail.Rows(164).RowHeight=32;detail.Range('A164:M164').WrapText=True
    s=wb.Worksheets('单AI分析')
    s.Range('B1').Value2='单个 AI 评测明细'
    for cell in ['B5','E5','H5','K5','N5','Q5','T5']:
        s.Range(cell).Font.Size=20;s.Range(cell).Font.Name='等线';s.Range(cell).Font.Bold=True
    if s.Range('R8').Value2=='Token 构成':
        s.Range('R8:V18').Cut(s.Range('AH8:AL18'))
        s.Range('AH:AL').EntireColumn.Hidden=True
        s.Range('R20:V24').Cut(s.Range('R8:V12'))
    rg=s.Range('R14:V24');rg.UnMerge();rg.Clear()
    s.Range('R14:V14').Merge();s.Range('R14').Value2='本轮资格'
    s.Range('R14:V14').Interior.Color=1973790;s.Range('R14:V14').Font.Color=16777215;s.Range('R14:V14').Font.Bold=True
    s.Range('R14:V14').Font.Name='等线';s.Range('R14:V14').Font.Size=11;s.Rows(14).RowHeight=18
    s.Range('R15:V19').Merge();s.Range('R15').Formula='=IF(AND(C3="Gemini",G3="第二轮"),"观察评分：不进入正式排名或两轮配对汇总。",IF(C3="Claude",IF(G3="第一轮","混合交付：Q评价Claude与Grok的成品；R不评分。","未开展独立运行，评分留空。"),IF(C3="Kimi","未形成完整交付，评分留空。",IF(AND(C3="Qwen",G3="第二轮"),"分项合计90，主线时间窗错误，最终Q封顶59。","正式比较样本。"))))'
    s.Range('R20:V24').Merge();s.Range('R20').Value2='下方保留Token与动作的数量、占比及构成图。逐项评分与核验依据见相应明细表。'
    for address in ['R15:V19','R20:V24']:
        box=s.Range(address);box.Font.Size=12;box.Font.Bold=False;box.WrapText=True;box.VerticalAlignment=-4108;box.Borders.LineStyle=1;box.Borders.Color=13882323
