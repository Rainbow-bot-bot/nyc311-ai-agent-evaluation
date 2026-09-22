from pathlib import Path
import json, zipfile, xml.etree.ElementTree as ET
import pyarrow.parquet as pq
import pyarrow.compute as pc
R=Path(__file__).resolve().parent;P=R/'汇总';S=Path(r'D:\项目1\原始数据');report={}
totals={};cats={};boroughs={}
for year in [2025,2026]:
    totals[year]=0;cats[year]={};boroughs[year]={}
    for month in range(1,9):
        f=next(S.glob(f'created_date={year}-{month:02d}-01*.parquet'))
        t=pq.read_table(f,columns=['unique_key','complaint_type','borough']);totals[year]+=t.num_rows
        for field,dest in [('complaint_type',cats),('borough',boroughs)]:
            z=t.group_by(field).aggregate([('unique_key','count')]).to_pylist()
            for x in z:dest[year][x[field]]=dest[year].get(x[field],0)+x['unique_key_count']
d=json.loads((P/'workbook_data.json').read_text(encoding='utf-8'))
for row in d['categories']:
    for year in [2025,2026]:assert cats[year].get(row['complaint_type'],0)==row[str(year)]
for row in d['boroughs']:
    for year in [2025,2026]:assert boroughs[year].get(row['borough'],0)==row[str(year)]
report['independent_pyarrow_totals']=totals;report['all_category_and_borough_counts_match']=True
metadata_total=sum(pq.ParquetFile(f).metadata.num_rows for f in S.glob('*.parquet'))
assert metadata_total==d['quality']['total'];report['metadata_total']=metadata_total
ns={'s':'http://schemas.openxmlformats.org/spreadsheetml/2006/main','c':'http://schemas.openxmlformats.org/drawingml/2006/chart','a':'http://schemas.openxmlformats.org/drawingml/2006/main'}
with zipfile.ZipFile(R/'最终成果.xlsx') as z:
    assert z.testzip() is None
    wk=ET.fromstring(z.read('xl/workbook.xml'));report['sheet_names']=[x.attrib['name'] for x in wk.findall('s:sheets/s:sheet',ns)]
    assert report['sheet_names']==['看板','月度趋势','类型与区域','数据质量']
    sheetdata={}
    for i in range(1,5):
        root=ET.fromstring(z.read(f'xl/worksheets/sheet{i}.xml'));cells={c.attrib['r']:c for c in root.findall('.//s:c',ns)};sheetdata[i]=cells
        assert not [c.attrib['r'] for c in cells.values() if c.attrib.get('t')=='e']
    def value(i,cell):return float(sheetdata[i][cell].find('s:v',ns).text)
    assert value(1,'A6')==totals[2026] and value(1,'G6')==metadata_total
    assert abs(value(1,'D6')-(totals[2026]/totals[2025]-1))<1e-12
    for i,row in enumerate(d['categories'],19):
        assert value(3,f'B{i}')==row['2025'] and value(3,f'C{i}')==row['2026']
        assert value(3,f'D{i}')==row['change']
    for r in [19,19+len(d['categories'])//2,18+len(d['categories'])]:
        assert sheetdata[3][f'D{r}'].find('s:f',ns).text==f'C{r}-B{r}'
    report['saved_formula_cells']=sum(c.find('s:f',ns) is not None for cells in sheetdata.values() for c in cells.values())
    report['chart_bindings']=[]
    for name in z.namelist():
        if '/charts/chart' in name and name.endswith('.xml'):
            root=ET.fromstring(z.read(name));series=root.findall('.//c:ser',ns)
            report['chart_bindings'].append({'file':name,'series':[{'ranges':[f.text for f in s.findall('.//c:f',ns)],'line_colors':[x.attrib for x in s.findall('.//a:ln/a:solidFill/a:srgbClr',ns)]} for s in series]})
    assert len(report['chart_bindings'])==2
    assert all(s['ranges'] and s['line_colors'] for c in report['chart_bindings'] for s in c['series'])
    report['zip_xml_and_cached_values_verified']=True
(R/'核验'/'independent_checks.json').write_text(json.dumps(report,ensure_ascii=False,indent=2),encoding='utf-8')
print(json.dumps(report,ensure_ascii=False,indent=2))
