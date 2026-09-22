from pathlib import Path
import win32com.client as w
A=Path(__file__).parent;D=A.parents[1]/'交付成果'
# Refuse an open target; never close a user's document.
try:
 old=w.GetActiveObject('Word.Application')
except Exception:
 old=None
if old is not None:
 for opened in old.Documents:
  if str(opened.FullName) in [str(D/'项目2_AI数据分析Agent自主交付评测报告.docx'),str(A/'项目2_简洁版报告.docx')]:
   raise RuntimeError('目标文档已打开，请先保存并关闭：'+str(opened.FullName))
for target in [D/'项目2_AI数据分析Agent自主交付评测报告.docx',A/'项目2_简洁版报告.docx']:
 if target.with_name('~$'+target.name[2:]).exists():
  raise RuntimeError('文档存在编辑锁，停止导出：'+str(target))
app=w.DispatchEx('Word.Application');doc=None
try:
 app.Visible=False
 for src,dst in [(D/'项目2_AI数据分析Agent自主交付评测报告.docx',D/'项目2_研究报告.pdf'),(A/'项目2_简洁版报告.docx',D/'项目2_研究报告_概览.pdf')]:
  temp=A/('待替换_'+dst.name)
  doc=app.Documents.Open(str(src),False,True);doc.Repaginate();print(dst.name,doc.ComputeStatistics(2));doc.ExportAsFixedFormat(str(temp),17);doc.Close(False);doc=None
  temp.replace(dst)
finally:
 if doc is not None:doc.Close(False)
 app.Quit()
