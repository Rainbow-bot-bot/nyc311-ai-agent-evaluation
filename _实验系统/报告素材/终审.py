from pathlib import Path
import json,re,hashlib,zipfile,sqlite3,sys,io,contextlib,datetime
from collections import defaultdict
import openpyxl,pypdfium2 as pdf
from docx import Document
R=Path(r'D:\项目2');D=R/'交付成果';A=Path(__file__).parent
sys.path.insert(0,str(R))
import 评测数据处理 as pipeline
result={};fail=[]
db=R/'_实验系统/telemetry/benchmark.sqlite';before=hashlib.sha256(db.read_bytes()).hexdigest()
con=sqlite3.connect(db.resolve().as_uri()+'?mode=ro',uri=True)
result['SQLite完整性']=con.execute('pragma integrity_check').fetchone()[0]
tables,evidence,_,_=pipeline.load_core().extract(db)
w=openpyxl.load_workbook(D/'项目2_AI评测分析.xlsx',data_only=True)
result['工作表']=w.sheetnames
for name,rows in tables:
 vals=list(w[name].values); heads=list(vals[0]); expected=[tuple(row.get(k) for k in heads[:len(rows[0])]) for row in rows]
 # Match by source columns; appended report formulas are checked separately.
 keys=list(rows[0]); idx=[heads.index(k) for k in keys]; actual=[tuple(r[i] for i in idx) for r in vals[1:]]
 exp=[tuple(r.get(k) for k in keys) for r in rows]
 norm=lambda row:tuple(v.strftime('%Y-%m-%d %H:%M:%S') if isinstance(v,datetime.datetime) else v for v in row)
 if [norm(r) for r in actual]!=[norm(r) for r in exp]:fail.append('SQLite/Excel不一致:'+name)
 result[name+'记录数']=len(actual)
wf=openpyxl.load_workbook(D/'项目2_AI评测分析.xlsx',data_only=False)
result['缺失样本图表占位']=[];result['Excel错误单元格']=[]
for s in w:
 for row in s:
  for c in row:
   if c.data_type=='e':
    intentional=s.title=='单AI分析' and c.column in (28,31) and 7<=c.row<=11 and c.value=='#N/A' and 'NA()' in str(wf[s.title][c.coordinate].value)
    result['缺失样本图表占位' if intentional else 'Excel错误单元格'].append(f'{s.title}!{c.coordinate}:{c.value}')
scores=defaultdict(float);counts=defaultdict(int)
for r in list(w['评分与用量'].values)[1:]:
 if r[2] in ['质量分项','可靠性分项','质量封顶']:
  kind='R' if str(r[3]).startswith('R') else 'Q';scores[r[0],r[1],kind]+=r[4];counts[r[0],r[1],kind]+=1
for r in list(w['运行记录'].values)[1:]:
 for kind,col in [('Q',13),('R',14)]:
  k=r[0],r[1],kind
  if (scores[k] if counts[k] else None)!=r[col]:fail.append('评分合计:'+str(k))
result['评分合计核对']='通过' if not any('评分合计' in x for x in fail) else '失败'
result['评分项数']={str(k):v for k,v in counts.items()}
result['断链']=[]
for f in [D/'开始阅读.html',D/'查证.html',D/'项目2_自主交付研究报告.md',R/'README.md']:
 for link in re.findall(r'(?:href|src)="([^"]+)"',f.read_text(encoding='utf8'))+re.findall(r'\]\(([^)]+)\)',f.read_text(encoding='utf8')):
  if not link.startswith(('#','http','javascript','data:')) and not (f.parent/link.split('#')[0]).exists():result['断链'].append([f.name,link])
result['PDF']={}
from PIL import Image,ImageDraw
for f in D.glob('*.pdf'):
 p=pdf.PdfDocument(str(f));texts=[page.get_textpage().get_text_range() for page in p]
 result['PDF'][f.name]={'页数':len(p),'每页字数':[len(t) for t in texts],'乱码替代符':sum(t.count('\ufffd') for t in texts)}
 for start in range(0,len(p),8):
  canvas=Image.new('RGB',(1200,1800),'#ddd');draw=ImageDraw.Draw(canvas)
  for n in range(start,min(start+8,len(p))):
   im=p[n].render(scale=.7).to_pil().convert('RGB');im.thumbnail((580,425));x=((n-start)%2)*600;y=((n-start)//2)*450;canvas.paste(im,(x,y+20));draw.text((x+5,y+3),str(n+1),fill='black')
  canvas.save(A/f'终审-{f.stem}-{start+1}.png')
for src,manifest in [(D/'项目2_AI数据分析Agent自主交付评测报告.docx','图表来源.json'),(A/'项目2_简洁版报告.docx','简洁版图表来源.json')]:
 m=json.loads((A/manifest).read_text(encoding='utf8')); doc=Document(src)
 with zipfile.ZipFile(src) as z: hashes={hashlib.sha256(z.read(n)).hexdigest() for n in z.namelist() if n.startswith('word/media/')}
 result[manifest]={'图片数':len(doc.inline_shapes),'来源数':len(m),'图片一致':hashes=={r['sha256'] for r in m},'素材未变':all(hashlib.sha256((A/r['file']).read_bytes()).hexdigest()==r['sha256'] for r in m)}
with zipfile.ZipFile(D/'项目2_AI评测分析.xlsx') as z:
 result['原生图表数']=len([n for n in z.namelist() if re.match(r'xl/charts/chart\d+.xml$',n)])
result['Notebook执行']=[];env={};n=json.loads((D/'01_评测数据整理与验收.ipynb').read_text(encoding='utf8'))
for i,c in enumerate(n['cells']):
 if c['cell_type']=='code':
  try:
   with contextlib.redirect_stdout(io.StringIO()):exec(''.join(c['source']),env)
   result['Notebook执行'].append([i,'通过'])
  except Exception as e:result['Notebook执行'].append([i,str(e)]);fail.append('Notebook单元'+str(i));break
result['数据库未改']=before==hashlib.sha256(db.read_bytes()).hexdigest()
if result['断链']:fail.append('交付入口断链')
if result['Excel错误单元格']:fail.append('Excel非预期错误')
from 终审判定 import failures
fail.extend(failures(result))
result['失败']=list(dict.fromkeys(fail))
result['交付文件哈希']={f.name:hashlib.sha256(f.read_bytes()).hexdigest() for f in D.iterdir() if f.is_file()}
result['裁定SHA256']=hashlib.sha256((A/'最终评分裁定.json').read_bytes()).hexdigest()
result['验收时间']=datetime.datetime.now().astimezone().isoformat()
result['检查范围']='数据与文件自动核对；评分裁定、原生交互和视觉审查另行核验。'
(A/'项目终审.json').write_text(json.dumps(result,ensure_ascii=False,indent=2),encoding='utf8')
print(json.dumps({k:v for k,v in result.items() if k!='评分项数'},ensure_ascii=False,indent=2))
sys.exit(1 if result['失败'] else 0)
