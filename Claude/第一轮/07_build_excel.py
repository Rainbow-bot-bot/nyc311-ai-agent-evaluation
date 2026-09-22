"""步骤7：把 中间汇总/*.csv 组装为最终 Excel 看板 最终成果.xlsx。
只做排版与图表，不重新计算指标（所有数字来自步骤6的汇总表）。
"""
import pandas as pd, pathlib, xlsxwriter

BASE = pathlib.Path(r"D:\项目2\Claude")
AGG = BASE / "中间汇总"
XLSX = BASE / "最终成果.xlsx"

def load(name):
    return pd.read_csv(AGG / f"{name}.csv", encoding="utf-8-sig")

T = {p.stem: load(p.stem) for p in sorted(AGG.glob("t*.csv"))}

wb = xlsxwriter.Workbook(str(XLSX), {"nan_inf_to_errors": True})

F = {
    "title":   wb.add_format({"bold": True, "font_size": 18, "font_color": "#1F3864"}),
    "h1":      wb.add_format({"bold": True, "font_size": 14, "font_color": "#FFFFFF",
                              "bg_color": "#1F3864", "align": "left", "valign": "vcenter"}),
    "h2":      wb.add_format({"bold": True, "font_size": 11, "font_color": "#1F3864",
                              "bottom": 2, "border_color": "#1F3864"}),
    "hdr":     wb.add_format({"bold": True, "bg_color": "#D9E2F3", "border": 1,
                              "border_color": "#8EA9DB", "align": "center", "valign": "vcenter",
                              "text_wrap": True}),
    "cell":    wb.add_format({"border": 1, "border_color": "#BFBFBF"}),
    "cellc":   wb.add_format({"border": 1, "border_color": "#BFBFBF", "align": "center"}),
    "num":     wb.add_format({"border": 1, "border_color": "#BFBFBF", "num_format": "#,##0"}),
    "num1":    wb.add_format({"border": 1, "border_color": "#BFBFBF", "num_format": "#,##0.0"}),
    "num2":    wb.add_format({"border": 1, "border_color": "#BFBFBF", "num_format": "#,##0.00"}),
    "pct1":    wb.add_format({"border": 1, "border_color": "#BFBFBF", "num_format": '#,##0.0"%"'}),
    "pct2":    wb.add_format({"border": 1, "border_color": "#BFBFBF", "num_format": '#,##0.00"%"'}),
    "date":    wb.add_format({"border": 1, "border_color": "#BFBFBF", "num_format": "yyyy-mm",
                              "align": "center"}),
    "date2":   wb.add_format({"border": 1, "border_color": "#BFBFBF", "num_format": "yyyy-mm-dd",
                              "align": "center"}),
    "note":    wb.add_format({"font_size": 9, "font_color": "#595959", "text_wrap": True,
                              "valign": "top"}),
    "kpi_lab": wb.add_format({"font_size": 10, "font_color": "#404040", "align": "center",
                              "bg_color": "#F2F2F2", "border": 1, "border_color": "#BFBFBF"}),
    "kpi_val": wb.add_format({"font_size": 14, "bold": True, "font_color": "#1F3864",
                              "align": "center", "bg_color": "#F2F2F2", "border": 1,
                              "border_color": "#BFBFBF"}),
    "body":    wb.add_format({"text_wrap": True, "valign": "top", "font_size": 11}),
    "bodyb":   wb.add_format({"text_wrap": True, "valign": "top", "font_size": 11, "bold": True}),
    "warn":    wb.add_format({"text_wrap": True, "valign": "top", "font_size": 11,
                              "bg_color": "#FFF2CC", "border": 1, "border_color": "#BF9000"}),
}

# 数值列 -> 格式名 的推断规则
def fmt_for(col, series):
    c = str(col)
    if c.endswith("%") or "占比" in c or "同比" in c:
        return F["pct2"] if series.abs().max() < 1000 else F["pct1"]
    if "指数" in c or "倍数" in c:
        return F["num2"]
    if "小时" in c or "(天)" in c:
        return F["num2"]
    if pd.api.types.is_integer_dtype(series) or (
            pd.api.types.is_float_dtype(series) and series.dropna().mod(1).eq(0).all()):
        return F["num"]
    return F["num1"]

_widths = {}   # 同一工作表内多个表格共用列宽，取各自需求的最大值，避免后写的表覆盖前面的

def write_table(ws, df, row, col=0, widths=None, first_col_width=None, date_cols=()):
    """写一个带表头的表格，返回下一空行行号。"""
    for j, c in enumerate(df.columns):
        ws.write(row, col + j, str(c), F["hdr"])
    ws.set_row(row, 30)
    for j, c in enumerate(df.columns):
        s = df[c]
        if c in date_cols:
            f = F["date"] if str(c) == "月份" else F["date2"]
        elif pd.api.types.is_numeric_dtype(s):
            f = fmt_for(c, s)
        else:
            f = F["cell"] if j == 0 else F["cellc"]
        for i, v in enumerate(s):
            r = row + 1 + i
            if pd.isna(v):
                ws.write_blank(r, col + j, None, f)
            elif c in date_cols:
                ws.write_datetime(r, col + j, pd.Timestamp(v).to_pydatetime(), f)
            else:
                ws.write(r, col + j, v, f)
    if widths:
        store = _widths.setdefault(ws.get_name(), {})
        for j, w in enumerate(widths):
            k = col + j
            store[k] = max(store.get(k, 0), w)
            ws.set_column(k, k, store[k])
    return row + 1 + len(df) + 2

def section(ws, row, text, span=10):
    ws.merge_range(row, 0, row, span, text, F["h1"])
    ws.set_row(row, 24)
    return row + 2

def _height(text, span, per_line_chars):
    """按文本长度估算合并单元格所需行高，宁可多留也不裁切。"""
    lines = -(-len(text) // per_line_chars)
    return max(30, 16 * lines + 6)

def note(ws, row, text, span=10, per_line=44):
    ws.merge_range(row, 0, row, span, text, F["note"])
    ws.set_row(row, _height(text, span, per_line))
    return row + 2

def warn(ws, row, text, span=10, per_line=40):
    """页首黄色提示框。"""
    ws.merge_range(row, 0, row, span, text, F["warn"])
    ws.set_row(row, _height(text, span, per_line))
    return row + 2


for part in ["08_excel_sheets.py", "09_excel_sheets2.py",
             "10_excel_sheets3.py", "11_excel_sheets4.py"]:
    exec(compile((BASE / part).read_text(encoding="utf-8"), part, "exec"))

wb.close()
print("已生成:", XLSX)
