from pathlib import Path
import gc
import hashlib
import importlib.util
import json
import time

import psutil
import win32com.client as win32

ROOT = Path(r"D:\项目2")
DB = ROOT / "_实验系统" / "telemetry" / "benchmark.sqlite"
OUT = ROOT / "交付成果" / "项目2_AI评测分析.xlsx"
QA = ROOT / "_实验系统" / "最终工作簿验收.json"

REPORT_SHEETS = ["评测总览", "单AI分析"]
DATA_SHEETS = ["运行记录", "评分与用量", "API计价", "行为记录", "核验记录", "字段说明"]

def load_core():
    path = ROOT / "_实验系统" / "评测数据处理核心.py"
    spec = importlib.util.spec_from_file_location("project2_cleaning_core", path)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module

def excel_pids():
    return {p.info["pid"] for p in psutil.process_iter(["pid", "name"])
            if (p.info["name"] or "").lower() == "excel.exe"}

def has_visible_window(pid):
    found = []
    import win32gui, win32process
    def check(hwnd, _):
        if not win32gui.IsWindowVisible(hwnd):
            return
        _, window_pid = win32process.GetWindowThreadProcessId(hwnd)
        if window_pid == pid and win32gui.GetWindowText(hwnd):
            found.append(hwnd)
    win32gui.EnumWindows(check, None)
    return bool(found)

def cleanup_excel(before):
    time.sleep(0.2)
    for pid in excel_pids() - before:
        if has_visible_window(pid):
            continue
        try:
            proc = psutil.Process(pid)
            proc.terminate()
            proc.wait(3)
        except (psutil.NoSuchProcess, psutil.TimeoutExpired):
            try:
                psutil.Process(pid).kill()
            except psutil.NoSuchProcess:
                pass

def rgb(r, g, b):
    return r + (g << 8) + (b << 16)

DARK = rgb(28, 28, 28)
MID = rgb(210, 210, 210)
LIGHT = rgb(247, 247, 247)
RED = rgb(220, 220, 220)
GRID = rgb(210, 210, 210)
TEXT = rgb(28, 28, 28)
MUTED = rgb(92, 92, 92)
NAVY = rgb(28, 28, 28)
BLUE = rgb(0, 163, 255)
TEAL = rgb(0, 204, 140)
AMBER = rgb(255, 204, 0)
PURPLE = rgb(168, 92, 255)
PANEL = rgb(245, 245, 245)
PANEL_EDGE = rgb(196, 196, 196)
PALE_BLUE = rgb(236, 236, 236)
PALE_TEAL = rgb(240, 240, 240)
TABLE_GRID = rgb(186, 186, 186)
ACCENT_DARK = rgb(28, 28, 28)
ACCENT_BLUE = BLUE
ACCENT_TEAL = TEAL
OUTLIER_BLACK = rgb(28, 28, 28)
CHART_LIGHT_GRAY = rgb(218, 218, 218)
CHART_PALE_BLUE = rgb(210, 229, 239)
CHART_BG = rgb(255, 255, 255)
KPI_BG_BLUE = rgb(255, 255, 255)
KPI_BG_GREEN = rgb(247, 247, 247)
CHART_BLUE = rgb(0, 168, 255)
CHART_MINT = rgb(0, 214, 150)
CHART_LIME = rgb(190, 230, 20)
CHART_AMBER = rgb(255, 196, 0)
CHART_CORAL = rgb(255, 92, 140)
CHART_LILAC = rgb(168, 92, 255)
MODEL_COLORS = (CHART_BLUE, CHART_MINT, CHART_LIME,
                CHART_LILAC, CHART_AMBER, CHART_CORAL)
CATEGORY_COLORS = (CHART_BLUE, CHART_MINT, CHART_AMBER,
                   CHART_LILAC, CHART_CORAL)
KPI_ACCENTS = (BLUE, TEAL, AMBER, PURPLE)
FONT = "等线"
BLACK = rgb(0, 0, 0)
HEAT_LO = (210, 232, 255)
HEAT_HI = (21, 75, 175)
HEADER_FILL = rgb(236, 236, 236)
HEADER_LINE = rgb(70, 70, 70)

def chart_by_title(ws, title):
    for i in range(1, ws.ChartObjects().Count + 1):
        co = ws.ChartObjects(i)
        try:
            if co.Chart.HasTitle and co.Chart.ChartTitle.Text == title:
                return co
        except Exception:
            pass
    return None

def delete_chart(ws, title, object_name=None):
    if object_name:
        try:
            ws.ChartObjects(object_name).Delete()
            return
        except Exception:
            pass
    co = chart_by_title(ws, title)
    if co:
        co.Delete()

def style_bar(chart, pct=False):
    chart.PlotVisibleOnly = False
    chart.ChartArea.Format.Fill.Visible = -1
    chart.ChartArea.Format.Fill.Solid()
    chart.ChartArea.Format.Fill.ForeColor.RGB = rgb(255, 255, 255)
    chart.ChartArea.Format.Line.Visible = 0
    chart.ChartArea.Format.Shadow.Visible = 0
    chart.PlotArea.Format.Fill.Visible = -1
    chart.PlotArea.Format.Fill.Solid()
    chart.PlotArea.Format.Fill.ForeColor.RGB = rgb(255, 255, 255)
    chart.PlotArea.Format.Line.Visible = 0
    chart.HasLegend = True
    chart.Legend.Position = -4107
    chart.Legend.Font.Name = "微软雅黑"
    chart.Legend.Font.Size = 8
    chart.Legend.Font.Color = TEXT

    if chart.HasTitle:
        title_font = chart.ChartTitle.Format.TextFrame2.TextRange.Font
        title_font.Name = "微软雅黑"
        title_font.Size = 11
        title_font.Bold = -1
        title_font.Fill.ForeColor.RGB = TEXT

    try:
        category = chart.Axes(1)
        category.ReversePlotOrder = True
        category.HasTitle = True
        category.AxisTitle.Text = "AI"
        category.AxisTitle.Font.Name = "微软雅黑"
        category.AxisTitle.Font.Size = 8
        category.TickLabelPosition = 4
        category.TickLabels.Font.Name = "微软雅黑"
        category.TickLabels.Font.Size = 9
        category.TickLabels.Font.Color = TEXT
        category.Format.Line.Visible = 0
    except Exception:
        pass

    try:
        value = chart.Axes(2)
        value.ReversePlotOrder = False
        value.HasTitle = True
        value.AxisTitle.Text = "占比" if pct else "元"
        value.AxisTitle.Font.Name = "微软雅黑"
        value.AxisTitle.Font.Size = 8
        value.HasMajorGridlines = True
        value.MajorGridlines.Format.Line.ForeColor.RGB = GRID
        value.Format.Line.Visible = 0
        value.TickLabelPosition = 4
        value.TickLabels.Font.Name = "微软雅黑"
        value.TickLabels.Font.Size = 9
        value.TickLabels.Font.Color = TEXT
        if pct:
            value.MinimumScale = 0
            value.MaximumScaleIsAuto = True
            value.MajorUnitIsAuto = True
            value.TickLabels.NumberFormat = "0%"
    except Exception:
        pass

    ser = chart.SeriesCollection(1)
    ser.Format.Fill.ForeColor.RGB = CHART_BLUE
    ser.Format.Line.Visible = 0
    ser.HasDataLabels = True
    labels = ser.DataLabels()
    labels.ShowValue = True
    labels.ShowCategoryName = False
    labels.ShowSeriesName = False
    labels.Font.Name = "微软雅黑"
    labels.Font.Size = 8
    labels.Font.Color = TEXT
    labels.NumberFormat = "0.0%" if pct else "0.00"

def add_single_bar(ws, title, categories, values, start, end, pct=False):
    delete_chart(ws, title)
    a = ws.Range(start)
    b = ws.Range(end)
    co = ws.ChartObjects().Add(a.Left, a.Top,
                               b.Left + b.Width - a.Left,
                               b.Top + b.Height - a.Top)
    ch = co.Chart
    ch.ChartType = 57
    ch.HasTitle = True
    ch.ChartTitle.Text = title
    ch.SeriesCollection().NewSeries()
    ch.SeriesCollection(1).Formula = (
        '=SERIES("' + title + '",' + categories + ',' + values + ',1)'
    )
    style_bar(ch, pct=pct)
    return co

def add_stack(ws, title, labels, value_cells, start, end, colors,
              object_name=None):
    # Existing charts already point at the selector-driven helper cells. Reuse
    # them: rebuilding Series formulas in a visible Excel session can leave an
    # incomplete chart even though the formulas themselves look valid.
    co = None
    if object_name:
        try:
            co = ws.ChartObjects(object_name)
        except Exception:
            pass
    if co is None:
        co = chart_by_title(ws, title)
    if co is not None:
        if object_name:
            co.Name = object_name
        return co
    a = ws.Range(start)
    b = ws.Range(end)
    co = ws.ChartObjects().Add(a.Left, a.Top,
                               b.Left + b.Width - a.Left,
                               b.Top + b.Height - a.Top)
    if object_name:
        co.Name = object_name
    ch = co.Chart
    ch.ChartType = 59
    ch.HasTitle = True
    ch.ChartTitle.Text = title
    ch.PlotVisibleOnly = False
    for idx, (label, value_cell) in enumerate(zip(labels, value_cells), 1):
        ref = "'单AI分析'!$" + value_cell[0] + "$" + value_cell[1:]
        ch.SeriesCollection().NewSeries()
        ch.SeriesCollection(idx).Formula = (
            '=SERIES("' + label + '",\'单AI分析\'!$C$3,' + ref + ',' + str(idx) + ')'
        )
    ch.ChartArea.Format.Fill.Visible = -1
    ch.ChartArea.Format.Fill.Solid()
    ch.ChartArea.Format.Fill.ForeColor.RGB = rgb(255, 255, 255)
    ch.ChartArea.Format.Line.Visible = 0
    ch.ChartArea.Format.Shadow.Visible = 0
    ch.PlotArea.Format.Fill.Visible = -1
    ch.PlotArea.Format.Fill.Solid()
    ch.PlotArea.Format.Fill.ForeColor.RGB = rgb(255, 255, 255)
    ch.PlotArea.Format.Line.Visible = 0
    ch.HasLegend = True
    ch.Legend.Position = -4107
    ch.Legend.Font.Name = "微软雅黑"
    ch.Legend.Font.Size = 8
    try:
        category = ch.Axes(1)
        category.TickLabelPosition = -4134
        category.Format.Line.Visible = 0
        category.TickLabels.Font.Name = "微软雅黑"
        category.TickLabels.Font.Size = 9
    except Exception:
        pass

    try:
        value = ch.Axes(2)
        value.MinimumScale = 0
        value.MaximumScale = 1
        value.MajorUnit = 0.25
        value.TickLabels.NumberFormat = "0%"
        value.HasMajorGridlines = False
        value.Format.Line.Visible = 0
        value.TickLabelPosition = -4134
        value.TickLabels.Font.Name = "微软雅黑"
        value.TickLabels.Font.Size = 9
    except Exception:
        pass
    for idx in range(1, ch.SeriesCollection().Count + 1):
        ser = ch.SeriesCollection(idx)
        ser.Format.Fill.ForeColor.RGB = colors[idx - 1]
        ser.Format.Line.Visible = 0
        ser.HasDataLabels = True
        labels = ser.DataLabels()
        labels.ShowPercentage = True
        try:
            labels.ShowSeriesName = True
        except Exception:
            pass
        try:
            labels.ShowValue = False
        except Exception:
            pass
        labels.Font.Name = "微软雅黑"
        labels.Font.Size = 8
        labels.Font.Color = TEXT
    return co

def _section_like(ws, source_cell, target_range, text):
    try:
        ws.Range(target_range).UnMerge()
    except Exception:
        pass
    ws.Range(target_range).MergeCells = True
    source = ws.Range(source_cell)
    target = ws.Range(target_range)
    target.Value2 = text
    target.Interior.Color = source.Interior.Color
    target.Font.Name = source.Font.Name
    target.Font.Size = source.Font.Size
    target.Font.Bold = source.Font.Bold
    target.Font.Color = source.Font.Color
    target.HorizontalAlignment = source.HorizontalAlignment
    target.VerticalAlignment = source.VerticalAlignment


def _clear_zero_labels(ws, chart, cells):
    series = chart.SeriesCollection(1)
    for index, coord in enumerate(cells, 1):
        value = ws.Range(coord).Value2
        if not isinstance(value, (int, float)) or abs(value) < 1e-12:
            try:
                series.Points(index).DataLabel.Delete()
            except Exception:
                pass


def _row_has_text(ws, row, first=2, last=22):
    try:
        vals = ws.Range(ws.Cells(row, first), ws.Cells(row, last)).Value2
    except Exception:
        return True
    if vals is None:
        return False
    if not isinstance(vals, (tuple, list)):
        return vals not in (None, "")
    for item in vals:
        if isinstance(item, (tuple, list)):
            if any(value not in (None, "") for value in item):
                return True
        elif item not in (None, ""):
            return True
    return False


CELL_WIDTH = 8.11
CELL_HEIGHT = 16
KEEP_MERGE_AREAS = {
    "评测总览": (
        "C3:D3", "B4:N6", "B8:N16", "B20:N26",
    ),
    "单AI分析": (
        "C3:D3", "G3:H3", "B4:V6",
        "R8:V18", "B20:P24", "R20:V24",
        "B29:H33", "J29:P33", "R29:V33",
    ),
}


def _box(rng):
    return (rng.Row, rng.Column,
            rng.Row + rng.Rows.Count - 1,
            rng.Column + rng.Columns.Count - 1)


def _inside(inner, outer):
    ir1, ic1, ir2, ic2 = inner
    or1, oc1, or2, oc2 = outer
    return ir1 >= or1 and ic1 >= oc1 and ir2 <= or2 and ic2 <= oc2


def _text_columns(text):
    units = 0
    for char in str(text or ""):
        units += 2 if ord(char) > 127 else 1
    return max(1, min(21, (units + 6) // 7))


def _uniform_report_grid(wb):
    for name in REPORT_SHEETS:
        ws = wb.Worksheets(name)
        for col in range(1, 23):
            ws.Columns(col).ColumnWidth = CELL_WIDTH
        for row in range(1, 81):
            ws.Rows(row).RowHeight = CELL_HEIGHT


def _shrink_nontable_merges(wb):
    """Keep table merges. Other merges span only as many equal cells as the text needs."""
    for name in REPORT_SHEETS:
        ws = wb.Worksheets(name)
        keep = [_box(ws.Range(area)) for area in KEEP_MERGE_AREAS[name]]
        seen = set()
        jobs = []
        for row in range(1, 43):
            for col in range(2, 23):
                cell = ws.Cells(row, col)
                if not cell.MergeCells:
                    continue
                area = cell.MergeArea
                addr = area.Address
                if addr in seen:
                    continue
                seen.add(addr)
                box = _box(area)
                if any(_inside(box, kept) for kept in keep):
                    continue
                jobs.append(area)
        for area in jobs:
            first = area.Cells(1, 1)
            formula = first.Formula
            value = first.Value2
            font_name = first.Font.Name
            font_size = first.Font.Size
            font_bold = first.Font.Bold
            font_color = first.Font.Color
            fill = first.Interior.Color
            start_row, start_col = first.Row, first.Column
            display = value
            if isinstance(formula, str) and formula.startswith("="):
                display = value
            need = _text_columns(display)
            area.UnMerge()
            area.ClearFormats()
            last_col = min(22, start_col + need - 1)
            target = ws.Range(ws.Cells(start_row, start_col),
                              ws.Cells(start_row, last_col))
            if last_col > start_col:
                target.MergeCells = True
            if isinstance(formula, str) and formula.startswith("="):
                first.Formula = formula
            else:
                first.Value2 = value
            target.Font.Name = font_name or FONT
            target.Font.Size = font_size or 11
            target.Font.Bold = font_bold
            target.Font.Color = font_color
            target.Interior.Color = fill
            target.VerticalAlignment = -4108
            target.HorizontalAlignment = -4131
            target.WrapText = False


def _snap_charts(wb):
    overview = wb.Worksheets("评测总览")
    single = wb.Worksheets("单AI分析")
    _place_chart(overview, 1, "P4", "V16")
    _place_chart(overview, 2, "P19", "V26")
    _place_chart(overview, 3, "B30", "K36")
    _place_chart(overview, 4, "M30", "V36")
    _place_chart(single, 1, "B8", "H17")
    _place_chart(single, 2, "J8", "P17")
    _place_chart(single, 3, "B35", "H41")
    _place_chart(single, 4, "J35", "P41")


def _remerge(ws, address):
    target = ws.Range(address)
    origin = ws.Range(address.split(":")[0])
    if origin.MergeCells:
        origin.MergeArea.UnMerge()
    try:
        target.UnMerge()
    except Exception:
        pass
    target.MergeCells = True
    return target


def _restore_page_geometry(wb):
    """Put column widths, row heights, and page chrome back. Do not rebuild tables."""
    widths = [2.0, 15.67] + [10.33] * 20
    heights = {
        "评测总览": [
            29.0, 16.1, 24.0, 17.0, 17.0, 17.0, 17.0, 26.0, 23.0, 23.0, 23.0,
            23.0, 23.0, 23.0, 23.0, 23.0, 17.0, 20.0, 22.1, 22.1, 20.0, 20.0,
            20.0, 20.0, 20.0, 20.0, 17.0, 22.1, 20.0, 20.0, 20.0, 20.0, 20.0,
            20.0, 20.0, 15.6, 13.8, 13.8, 13.8, 13.8, 13.8, 13.8,
        ],
        "单AI分析": [
            29.0, 16.1, 24.0, 17.0, 17.0, 17.0, 17.0, 21.0, 18.0, 18.0, 18.0,
            18.0, 17.0, 21.0, 18.0, 18.0, 18.0, 18.0, 17.0, 25.1, 22.1, 20.0,
            20.0, 20.0, 17.0, 17.0, 22.1, 20.0, 20.0, 20.0, 25.1, 27.0, 23.0,
            23.0, 23.0, 20.0, 18.0, 18.0, 18.0, 20.0, 20.0, 13.8,
        ],
    }
    chrome = {
        "评测总览": (
            "B1:P1", "R1:V1", "B2:V2", "F3:N3",
            "B18:N18", "B19:N19", "B28:V28",
        ),
        "单AI分析": (
            "B1:P1", "R1:V1", "B2:V2", "J3:P3",
            "B18:H18", "J18:P18", "B26:V26", "B27:V27",
        ),
    }
    for name in REPORT_SHEETS:
        ws = wb.Worksheets(name)
        for col, width in enumerate(widths, 1):
            ws.Columns(col).ColumnWidth = width
        for row, height in enumerate(heights[name], 1):
            ws.Rows(row).RowHeight = height
        for area in chrome[name]:
            _remerge(ws, area)
    _snap_charts(wb)


def _uniform_empty_cells(wb):
    _uniform_report_grid(wb)


def _paint_table_black(ws, area):
    rng = ws.Range(area)
    for edge in (7, 8, 9, 10, 11, 12):
        line = rng.Borders(edge)
        line.LineStyle = 1
        line.Weight = 2
        line.Color = BLACK


def _write_round_table(ws):
    """两轮变化 uses B:P so it lines up with Token (B:H) + 动作 (J:P)."""
    def round_metric(column, rnd):
        return (
            f'=IF(SUMPRODUCT((运行记录!$A$2:$A$16=$C$3)*(运行记录!$B$2:$B$16="{rnd}")'
            f'*ISNUMBER(运行记录!${column}$2:${column}$16))=0,"",'
            f'SUMIFS(运行记录!${column}$2:${column}$16,运行记录!$A$2:$A$16,$C$3,'
            f'运行记录!$B$2:$B$16,"{rnd}"))'
        )

    q1, q2 = round_metric("N", "第一轮"), round_metric("N", "第二轮")
    r1, r2 = round_metric("O", "第一轮"), round_metric("O", "第二轮")
    t1, t2 = round_metric("M", "第一轮"), round_metric("M", "第二轮")
    try:
        ws.Range("B20:P24").UnMerge()
    except Exception:
        pass
    ws.Range("B20:P24").ClearContents()
    ws.Range("B20:P24").ClearFormats()
    title = ws.Range("B20:P20")
    title.MergeCells = True
    title.Value2 = "两轮变化"
    title.Interior.Color = NAVY
    title.Font.Color = rgb(255, 255, 255)
    title.Font.Bold = True
    title.Font.Name = FONT
    title.Font.Size = 11
    title.HorizontalAlignment = -4131
    title.VerticalAlignment = -4108
    ws.Rows(20).RowHeight = 20
    ws.Rows(21).RowHeight = 20
    for area, text in (("B21:C21", "指标"), ("D21:H21", "第一轮"),
                       ("I21:L21", "第二轮"), ("M21:P21", "变化")):
        cell = ws.Range(area)
        cell.MergeCells = True
        cell.Value2 = text
        cell.Interior.Color = NAVY
        cell.Font.Color = rgb(255, 255, 255)
        cell.Font.Bold = True
        cell.Font.Name = FONT
        cell.Font.Size = 11
        cell.HorizontalAlignment = -4108
        cell.VerticalAlignment = -4108
    for row, label, left, right, delta, fmt in (
        (22, "成果质量 Q", q1, q2,
         '=IF(OR(D22="",I22=""),"—",I22-D22)', "0"),
        (23, "可靠性 R", r1, r2,
         '=IF(OR(D23="",I23=""),"—",I23-D23)', "0"),
        (24, "耗时（分钟）", t1, t2,
         '=IF(OR(D24="",I24=""),"—",I24-D24)', "0.0"),
    ):
        lab = ws.Range(f"B{row}:C{row}")
        lab.MergeCells = True
        lab.Value2 = label
        lab.Font.Name = FONT
        lab.Font.Size = 11
        lab.HorizontalAlignment = -4131
        lab.VerticalAlignment = -4108
        for area, formula in ((f"D{row}:H{row}", left),
                              (f"I{row}:L{row}", right),
                              (f"M{row}:P{row}", delta)):
            cell = ws.Range(area)
            cell.MergeCells = True
            ws.Range(area.split(":")[0]).Formula = formula
            ws.Range(area.split(":")[0]).NumberFormat = fmt
            cell.HorizontalAlignment = -4152
            cell.VerticalAlignment = -4108
            cell.Font.Name = FONT
            cell.Font.Size = 11
        ws.Range(f"B{row}:P{row}").Interior.Color = rgb(255, 255, 255)
    _paint_table_black(ws, "B20:P24")


def _align_range(rng, horizontal, vertical=-4108):
    rng.HorizontalAlignment = horizontal
    rng.VerticalAlignment = vertical


def _align_report_cells(wb):
    """Line tables to the same column grid. Do not move charts."""
    single = wb.Worksheets("单AI分析")
    _write_round_table(single)
    _align_range(single.Range("B18:H18"), -4131)
    _align_range(single.Range("J18:P18"), -4131)
    _align_range(single.Range("B20:P20"), -4131)
    for header in ("B21:P21", "B29:H29", "J29:P29",
                   "R8:V8", "R14:V14", "R20:V20", "R29:V29"):
        _align_range(single.Range(header), -4108)
    for labels in ("B22:C24", "B30:D33", "J30:L33",
                   "R9:T12", "R15:T18", "R30:S32"):
        _align_range(single.Range(labels), -4131)
    for nums in ("D22:P24", "E30:H33", "M30:P33",
                 "U9:V12", "U15:V18", "T30:V32"):
        _align_range(single.Range(nums), -4152)
    note = single.Range("R21:V24")
    note.WrapText = True
    _align_range(note, -4131)
    overview = wb.Worksheets("评测总览")
    _align_range(overview.Range("B8:N8"), -4108)
    _align_range(overview.Range("B20:N20"), -4108)
    _align_range(overview.Range("B9:B16"), -4131)
    _align_range(overview.Range("B21:B26"), -4131)
    _align_range(overview.Range("C9:J16"), -4152)
    _align_range(overview.Range("K9:N16"), -4131)
    _align_range(overview.Range("C21:N26"), -4152)


def _pack_single_page(wb):
    """Close the holes left when charts were moved off the composition tables."""
    ws = wb.Worksheets("单AI分析")
    _write_round_table(ws)

    try:
        ws.Range("B29:P36").UnMerge()
    except Exception:
        pass
    ws.Range("B29:P36").ClearContents()
    ws.Range("B29:P36").ClearFormats()
    for area, formula in (
        ("B29:H29", '=IF(COUNT(U9:U12)=0,"Token 构成  ·  无记录",'
         '"Token 构成  ·  "&TEXT(SUM(U9:U12),"#,##0"))'),
        ("J29:P29", '=IF(COUNT(U15:U18)=0,"动作构成  ·  无记录",'
         '"动作构成  ·  "&TEXT(SUM(U15:U18),"0")&" 次")'),
    ):
        head = ws.Range(area)
        head.MergeCells = True
        head.Formula = formula
        head.Interior.Color = NAVY
        head.Font.Color = rgb(255, 255, 255)
        head.Font.Name = FONT
        head.Font.Size = 10
        head.Font.Bold = True
    for offset, left_label, right_label in (
        (0, "未缓存输入", "直接贡献交付"),
        (1, "缓存读取", "必要探索"),
        (2, "缓存写入", "管理动作"),
        (3, "输出", "确认无收益"),
    ):
        row = 30 + offset
        src_l = f"U{9 + offset}"
        src_r = f"U{15 + offset}"
        left = ws.Range(f"B{row}:D{row}")
        right = ws.Range(f"J{row}:L{row}")
        left.MergeCells = True
        right.MergeCells = True
        left.Value2 = left_label
        right.Value2 = right_label
        ws.Range(f"E{row}").Formula = (
            f'=IF(COUNT($U$9:$U$12)=0,"",IFERROR({src_l}/SUM($U$9:$U$12),0))')
        ws.Range(f"E{row}").NumberFormat = "0.0%"
        ws.Range(f"M{row}").Formula = (
            f'=IF(COUNT($U$15:$U$18)=0,"",IFERROR({src_r}/SUM($U$15:$U$18),0))')
        ws.Range(f"M{row}").NumberFormat = "0.0%"
        for area, source in ((f"F{row}:H{row}", src_l), (f"N{row}:P{row}", src_r)):
            cell = ws.Range(area)
            cell.MergeCells = True
            cell.Formula = f'=IF(ISNUMBER({source}),{source},"")'
            cell.NumberFormat = "#,##0"
            cell.HorizontalAlignment = -4152
        for area in (f"B{row}:H{row}", f"J{row}:P{row}"):
            line = ws.Range(area)
            line.Interior.Color = rgb(255, 255, 255)
            line.Font.Name = FONT
            line.Font.Size = 9
            line.Font.Color = TEXT
    _paint_table_black(ws, "B29:H33")
    _paint_table_black(ws, "J29:P33")

    try:
        ws.Range("R29:V36").UnMerge()
    except Exception:
        pass
    ws.Range("R29:V36").ClearContents()
    ws.Range("R29:V36").ClearFormats()
    head = ws.Range("R29:V29")
    head.MergeCells = True
    head.Value2 = "构成说明"
    head.Interior.Color = NAVY
    head.Font.Color = rgb(255, 255, 255)
    head.Font.Name = FONT
    head.Font.Size = 10
    head.Font.Bold = True
    for row, label, formula in (
        (30, "缓存读取占比",
         '=IF(COUNT(U9:U12)=0,"",IFERROR(U10/SUM(U9:U12),0))'),
        (31, "直接贡献＋必要探索",
         '=IF(COUNT(U15:U18)=0,"",IFERROR(SUM(U15:U16)/SUM(U15:U18),0))'),
        (32, "确认无收益占比",
         '=IF(COUNT(U15:U18)=0,"",IFERROR(U18/SUM(U15:U18),0))'),
    ):
        lab = ws.Range(f"R{row}:S{row}")
        lab.MergeCells = True
        lab.Value2 = label
        lab.Font.Name = FONT
        lab.Font.Size = 9
        lab.Font.Color = MUTED
        val = ws.Range(f"T{row}:V{row}")
        val.MergeCells = True
        val.Formula = formula
        val.NumberFormat = "0.0%"
        val.Font.Name = FONT
        val.Font.Size = 11
        val.Font.Bold = True
        val.HorizontalAlignment = -4152
        ws.Range(f"R{row}:V{row}").Interior.Color = rgb(255, 255, 255)
    note = ws.Range("R33:V33")
    note.MergeCells = True
    note.Value2 = "按当前 AI 与轮次的已记录项计算"
    note.Font.Name = FONT
    note.Font.Size = 8
    note.Font.Color = MUTED
    _paint_table_black(ws, "R29:V33")


def _format_report_pages(wb):
    font_name = FONT

    def format_chart(chart):
        try:
            chart.ChartArea.Format.Fill.Visible = -1
            chart.ChartArea.Format.Fill.Solid()
            chart.ChartArea.Format.Fill.ForeColor.RGB = rgb(255, 255, 255)
            chart.ChartArea.Format.Line.Visible = -1
            chart.ChartArea.Format.Line.ForeColor.RGB = HEADER_LINE
            chart.ChartArea.Format.Line.Weight = 1
            chart.ChartArea.Format.Shadow.Visible = 0
            chart.PlotArea.Format.Fill.Visible = -1
            chart.PlotArea.Format.Fill.Solid()
            chart.PlotArea.Format.Fill.ForeColor.RGB = rgb(255, 255, 255)
            chart.PlotArea.Format.Line.Visible = 0
        except Exception:
            pass

        try:
            if chart.HasTitle:
                title_font = chart.ChartTitle.Format.TextFrame2.TextRange.Font
                title_font.Name = font_name
                title_font.Size = 11
                title_font.Bold = -1
                title_font.Fill.ForeColor.RGB = TEXT
        except Exception:
            pass

        try:
            if chart.HasLegend:
                chart.Legend.Font.Name = font_name
                chart.Legend.Font.Size = 8
                chart.Legend.Font.Color = TEXT
        except Exception:
            pass

        for axis_type in (1, 2):
            try:
                axis = chart.Axes(axis_type)
                axis.TickLabels.Font.Name = font_name
                axis.TickLabels.Font.Size = 9
                axis.TickLabels.Font.Color = TEXT
                axis.HasMajorGridlines = False
                axis.Format.Line.Visible = 0
                if axis.HasTitle:
                    axis.AxisTitle.Font.Name = font_name
                    axis.AxisTitle.Font.Size = 9
                    axis.AxisTitle.Font.Color = TEXT
            except Exception:
                pass

        try:
            for index in range(1, chart.SeriesCollection().Count + 1):
                series = chart.SeriesCollection(index)
                if series.HasDataLabels:
                    series.DataLabels().Font.Name = font_name
                    series.DataLabels().Font.Size = 8
                    series.DataLabels().Font.Color = TEXT
                    try:
                        series.DataLabels().Font.NameFarEast = font_name
                    except Exception:
                        pass
        except Exception:
            pass

    def set_box(chart_object, ws, start_cell, end_cell):
        start = ws.Range(start_cell)
        end = ws.Range(end_cell)
        chart_object.Left = start.Left
        chart_object.Top = start.Top
        chart_object.Width = end.Left + end.Width - start.Left
        chart_object.Height = end.Top + end.Height - start.Top
        # Excel's move-and-size placement drifts a few points after zoom/view
        # changes; keep the audited cell geometry fixed.
        chart_object.Placement = 3

    overview = wb.Worksheets("评测总览")
    overview.UsedRange.Font.Name = font_name
    try:
        overview.UsedRange.Font.NameFarEast = font_name
    except Exception:
        pass
    overview.Range("E3").Font.Name = font_name
    try:
        overview.Range("E3").Font.NameFarEast = font_name
    except Exception:
        pass

    for index in range(1, overview.ChartObjects().Count + 1):
        format_chart(overview.ChartObjects(index).Chart)

    set_box(overview.ChartObjects(1), overview, "P4", "V16")
    set_box(overview.ChartObjects(2), overview, "P19", "V26")
    set_box(overview.ChartObjects(3), overview, "B30", "K36")
    set_box(overview.ChartObjects(4), overview, "M30", "V36")

    single = wb.Worksheets("单AI分析")
    single.UsedRange.Font.Name = font_name
    try:
        single.UsedRange.Font.NameFarEast = font_name
    except Exception:
        pass
    for cell in ("E3", "I3"):
        single.Range(cell).Font.Name = font_name
        try:
            single.Range(cell).Font.NameFarEast = font_name
        except Exception:
            pass

    for index in range(1, single.ChartObjects().Count + 1):
        format_chart(single.ChartObjects(index).Chart)

    set_box(single.ChartObjects(1), single, "B8", "H17")
    set_box(single.ChartObjects(2), single, "J8", "P17")
    set_box(single.ChartObjects(3), single, "B35", "H41")
    set_box(single.ChartObjects(4), single, "J35", "P41")
    for index in (3, 4):
        try:
            single.ChartObjects(index).Chart.HasTitle = False
        except Exception:
            pass

    for cell_range in ("B1:P1", "B2:P2", "B27:V27"):
        single.Range(cell_range).Font.Name = font_name


def _apply_visual_system(wb):
    """Give the two reader-facing pages one restrained, legible visual system."""
    def frame_black(rng, weight=2):
        for edge in (7, 8, 9, 10):
            line = rng.Borders(edge)
            line.LineStyle = 1
            line.Weight = weight
            line.Color = BLACK

    def paint_table(ws, area):
        rng = ws.Range(area)
        for edge in (7, 8, 9, 10, 11, 12):
            line = rng.Borders(edge)
            line.LineStyle = 1
            line.Weight = 2
            line.Color = BLACK

    def section(ws, area):
        rng = ws.Range(area)
        rng.Interior.Color = HEADER_FILL
        rng.Font.Color = TEXT
        rng.Font.Bold = True
        rng.Font.Size = 11
        rng.Font.Name = FONT
        frame_black(rng)

    def table_grid(ws, area, header_row):
        paint_table(ws, area)

    def kpi(ws, label, value, whole, background):
        box = ws.Range(whole)
        box.Interior.Color = background
        frame_black(box)
        lab = ws.Range(label)
        lab.Font.Color = MUTED
        lab.Font.Size = 8
        val = ws.Range(value)
        val.Font.Color = ACCENT_DARK
        val.Font.Bold = True
        val.Font.Size = 13

    def chart_panel(co, series_colors, point_colors=None, background=CHART_BG):
        ch = co.Chart
        for area in (ch.ChartArea, ch.PlotArea):
            area.Format.Fill.Visible = -1
            area.Format.Fill.Solid()
            area.Format.Fill.ForeColor.RGB = background
        ch.ChartArea.Format.Line.Visible = -1
        ch.ChartArea.Format.Line.ForeColor.RGB = HEADER_LINE
        ch.ChartArea.Format.Line.Weight = 1
        ch.PlotArea.Format.Line.Visible = 0
        if ch.HasTitle:
            title_font = ch.ChartTitle.Format.TextFrame2.TextRange.Font
            title_font.Fill.ForeColor.RGB = ACCENT_DARK
            title_font.Size = 11
        for idx, color in enumerate(series_colors, 1):
            if idx > ch.SeriesCollection().Count:
                break
            series = ch.SeriesCollection(idx)
            series.Format.Fill.ForeColor.RGB = color
            series.Format.Line.Visible = 0
        if point_colors and ch.SeriesCollection().Count:
            for point_index, color in point_colors.items():
                try:
                    ch.SeriesCollection(1).Points(point_index).Format.Fill.ForeColor.RGB = color
                except Exception:
                    pass
        ch.HasLegend = True
        ch.Legend.Position = -4107
        ch.Legend.IncludeInLayout = True
        ch.Legend.Font.Name = FONT
        ch.Legend.Font.Size = 8
        ch.Legend.Font.Color = TEXT
        stacked = ch.ChartType in (52, 53, 54, 58, 59, 60)
        _label_chart_series(ch, stacked)
        for axis_type in (1, 2):
            try:
                axis = ch.Axes(axis_type)
                axis.TickLabelPosition = 4
                title_text = ""
                try:
                    title_text = str(axis.AxisTitle.Text or "")
                except Exception:
                    pass
                if title_text in ("", "坐标轴标题", "Axis Title"):
                    axis.HasTitle = False
            except Exception:
                pass

    overview = wb.Worksheets("评测总览")
    overview.Activate()
    single = wb.Worksheets("单AI分析")
    single.Activate()
    for ws in (overview, single):
        ws.Activate()
        wb.Application.ActiveWindow.DisplayGridlines = True
        # Remove the old blue-tinted template colors before applying the
        # monochrome sections below. Empty cells remain native Excel cells.
        ws.Range("B1:V36").Interior.Pattern = -4142
        ws.Range("B1:V36").Font.Color = TEXT
        ws.Range("B1:V1").Font.Name = FONT
        ws.Range("B1:V1").Font.Size = 14
        ws.Range("B1:V1").Font.Bold = True
        ws.Range("B1:V1").Interior.Color = ACCENT_DARK
        ws.Range("B1:V1").Font.Color = rgb(255, 255, 255)
        ws.Rows(1).RowHeight = 22
        frame_black(ws.Range("B1:V1"))
        ws.Range("B2:V2").Interior.Color = PALE_BLUE
        ws.Range("B2:V2").Font.Color = MUTED
        ws.Range("B2:V2").Font.Size = 9
        ws.Range("B3:V3").Font.Size = 10
        for selector in (("C3:D3", "G3:H3") if ws.Name == "单AI分析"
                         else ("C3:D3",)):
            ws.Range(selector).Interior.Color = PANEL
        ws.Range("R1:V1").Font.Size = 9
        ws.Range("R1:V1").Font.Bold = False
        ws.Range("R1:V1").Font.Color = rgb(255, 255, 255)

    for index, (label, value, whole) in enumerate((
        ("B4:D4", "B5:D6", "B4:D6"),
        ("E4:G4", "E5:G6", "E4:G6"),
        ("H4:J4", "H5:J6", "H4:J6"),
        ("K4:N4", "K5:N6", "K4:N6"),
    )):
        kpi(overview, label, value, whole,
            KPI_BG_BLUE if index % 2 == 0 else KPI_BG_GREEN)
        overview.Range(whole).Borders(7).Color = KPI_ACCENTS[index % 4]
        overview.Range(whole).Borders(7).Weight = 4
    overview.Range("B8:N8").Interior.Color = NAVY
    overview.Range("B8:N8").Font.Color = rgb(255, 255, 255)
    overview.Range("B8:N8").Font.Bold = True
    overview.Range("B8:N8").Font.Size = 10
    overview.Rows(8).RowHeight = 20
    for row in range(9, 17):
        overview.Range(f"B{row}:N{row}").Interior.Color = rgb(255, 255, 255)
    overview.Range("C9:D14").Font.Color = BLUE
    overview.Range("C9:D14").Font.Bold = True
    overview.Range("B15:N16").Font.Color = MUTED
    section(overview, "B19:N19")
    section(overview, "B28:V28")
    overview.Range("B20:N20").Interior.Color = NAVY
    overview.Range("B20:N20").Font.Color = rgb(255, 255, 255)
    overview.Range("B20:N20").Font.Bold = True
    for row in range(21, 27):
        overview.Range(f"B{row}:N{row}").Interior.Color = rgb(255, 255, 255)
    for index, colors, point_colors in (
        (1, [CHART_BLUE, CHART_MINT], None),
        (2, [CHART_BLUE], dict(enumerate(MODEL_COLORS, 1))),
        (3, [CHART_BLUE], dict(enumerate(MODEL_COLORS, 1))),
        (4, [CHART_BLUE], dict(enumerate(MODEL_COLORS, 1))),
    ):
        chart_panel(overview.ChartObjects(index), colors, point_colors)
    for index in (3, 4):
        chart = overview.ChartObjects(index).Chart
        chart.ChartArea.Format.Fill.ForeColor.RGB = CHART_BG
        chart.PlotArea.Format.Fill.ForeColor.RGB = CHART_BG
        try:
            axis = chart.Axes(2)
            axis.HasMajorGridlines = True
            axis.MajorGridlines.Format.Line.ForeColor.RGB = GRID
            axis.MajorGridlines.Format.Line.Weight = 0.5
        except Exception:
            pass
    table_grid(overview, "B8:N16", 8)
    table_grid(overview, "B20:N26", 20)
    paint_table(overview, "B8:N16")
    paint_table(overview, "B20:N26")

    for index, (label, value, whole) in enumerate((
        ("B4:D4", "B5:D6", "B4:D6"),
        ("E4:G4", "E5:G6", "E4:G6"),
        ("H4:J4", "H5:J6", "H4:J6"),
        ("K4:M4", "K5:M6", "K4:M6"),
        ("N4:P4", "N5:P6", "N4:P6"),
        ("Q4:S4", "Q5:S6", "Q4:S6"),
        ("T4:V4", "T5:V6", "T4:V6"),
    )):
        kpi(single, label, value, whole,
            KPI_BG_BLUE if index % 2 == 0 else KPI_BG_GREEN)
        single.Range(whole).Borders(7).Color = KPI_ACCENTS[index % 4]
        single.Range(whole).Borders(7).Weight = 4
    section(single, "B20:P20")
    section(single, "B27:V27")
    single.Range("B21:P21").Interior.Color = NAVY
    single.Range("B21:P21").Font.Color = rgb(255, 255, 255)
    single.Range("B21:P21").Font.Bold = True
    for row in range(22, 25):
        single.Range(f"B{row}:P{row}").Interior.Color = rgb(255, 255, 255)
    for index, colors, point_colors in (
        (1, [CHART_BLUE], dict(enumerate(CATEGORY_COLORS, 1))),
        (2, [CHART_MINT], dict(enumerate(CATEGORY_COLORS, 1))),
        (3, [CHART_BLUE, CHART_MINT, CHART_AMBER, CHART_CORAL], None),
        (4, [CHART_BLUE, CHART_MINT, CHART_AMBER, CHART_CORAL], None),
    ):
        chart_panel(single.ChartObjects(index), colors, point_colors)
    table_grid(single, "B21:P24", 21)
    paint_table(single, "B21:P24")

    # The right-hand source summary belongs to the same monochrome report.
    # Its original white merged note block was the largest visual void.
    for start, end, header_row, first_data, last_data in (
        ("R8:V12", "R8:V8", 8, 9, 12),
        ("R14:V18", "R14:V14", 14, 15, 18),
    ):
        area = single.Range(start)
        area.Interior.Color = rgb(255, 255, 255)
        paint_table(single, start)
        head = single.Range(end)
        head.Interior.Color = NAVY
        head.Font.Color = rgb(255, 255, 255)
        head.Font.Bold = True
        for row in range(first_data, last_data + 1):
            line = single.Range(f"R{row}:V{row}")
            line.Interior.Color = rgb(255, 255, 255)
            line.Font.Color = TEXT
    paint_table(single, "R20:V24")
    single.Range("R20:V20").Interior.Color = NAVY
    single.Range("R20:V20").Font.Color = rgb(255, 255, 255)
    single.Range("R20:V20").Font.Bold = True
    single.Range("R21:V24").Interior.Color = rgb(255, 255, 255)
    single.Range("R21:V24").Font.Color = TEXT

    for row in range(33, 37):
        single.Range(f"B{row}:H{row}").Interior.Color = rgb(255, 255, 255)
        single.Range(f"J{row}:P{row}").Interior.Color = rgb(255, 255, 255)
    single.Range("R30:V36").Interior.Color = rgb(255, 255, 255)
    for gap in ("B30:H32", "J30:P32"):
        g = single.Range(gap)
        g.Interior.Pattern = -4142
        g.Borders.LineStyle = 0
    paint_table(single, "B29:H29")
    paint_table(single, "B33:H36")
    paint_table(single, "J29:P29")
    paint_table(single, "J33:P36")
    paint_table(single, "R29:V36")
    for area in ("B29:H29", "J29:P29", "R29:V29"):
        head = single.Range(area)
        head.Interior.Color = NAVY
        head.Font.Color = rgb(255, 255, 255)
        head.Font.Name = FONT
        head.Font.Size = 10
    for value_row in (31, 33, 35):
        single.Range(f"R{value_row}:V{value_row}").Font.Size = 12


def _compose_single_detail(wb):
    """Keep the two breakdowns and their actual values in one compact cell-aligned card."""
    ws = wb.Worksheets("单AI分析")
    whole = ws.Range("B29:V36")
    whole.UnMerge()
    whole.ClearContents()
    whole.ClearFormats()
    whole.Interior.Color = PANEL
    for edge in (7, 8, 9, 10):
        rule = whole.Borders(edge)
        rule.LineStyle = 1
        rule.Weight = 2
        rule.Color = PANEL_EDGE

    for area, formula in (
        ("B29:H29", '=IF(COUNT(U9:U12)=0,"Token 构成  ·  无记录",'
         '"Token 构成  ·  "&TEXT(SUM(U9:U12),"#,##0"))'),
        ("J29:P29", '=IF(COUNT(U15:U18)=0,"动作构成  ·  无记录",'
         '"动作构成  ·  "&TEXT(SUM(U15:U18),"0")&" 次")'),
    ):
        header = ws.Range(area)
        header.MergeCells = True
        header.Formula = formula
        header.Interior.Color = PALE_BLUE
        header.Font.Name = "微软雅黑"
        header.Font.Size = 11
        header.Font.Bold = True
        header.Font.Color = NAVY
        header.VerticalAlignment = -4108
        header.Borders(9).LineStyle = 1
        header.Borders(9).Color = PANEL_EDGE

    ws.Range("I29:I36").Interior.Color = PANEL
    ws.Range("Q29:Q36").Interior.Color = PANEL
    ws.Range("R29:V36").Interior.Color = PANEL
    for row, left_label, right_label in (
        (33, "未缓存输入", "直接贡献交付"),
        (34, "缓存读取", "必要探索"),
        (35, "缓存写入", "管理动作"),
        (36, "输出", "确认无收益"),
    ):
        left = ws.Range(f"B{row}:D{row}")
        right = ws.Range(f"J{row}:L{row}")
        left.MergeCells = True
        right.MergeCells = True
        left.Value2 = left_label
        right.Value2 = right_label
        for cell, source, total in (
            (ws.Range(f"E{row}"), f"U{row-24}", "$U$9:$U$12"),
            (ws.Range(f"M{row}"), f"U{row-18}", "$U$15:$U$18"),
        ):
            cell.Formula = (f'=IF(COUNT({total})=0,"",'
                            f'IFERROR({source}/SUM({total}),0))')
            cell.NumberFormat = "0.0%"
        for area, source in ((f"F{row}:H{row}", f"U{row-24}"),
                             (f"N{row}:P{row}", f"U{row-18}")):
            cell = ws.Range(area)
            cell.MergeCells = True
            cell.Formula = f'=IF(ISNUMBER({source}),{source},"")'
            cell.NumberFormat = "#,##0"
            cell.HorizontalAlignment = -4152
        for area, color in ((f"B{row}:H{row}", BLUE),
                            (f"J{row}:P{row}", TEAL)):
            line = ws.Range(area)
            line.Interior.Color = rgb(255, 255, 255)
            line.Font.Name = "微软雅黑"
            line.Font.Size = 9
            line.Font.Color = TEXT
            line.VerticalAlignment = -4108
            line.Borders(9).LineStyle = 1
            line.Borders(9).Color = PANEL_EDGE
            accent = line.Borders(7)
            accent.LineStyle = 1
            accent.Color = color
            accent.Weight = 3

    for index in (1, 2):
        chart = ws.ChartObjects(index).Chart
        chart.PlotArea.Format.Fill.Visible = -1
        chart.PlotArea.Format.Fill.Solid()
        chart.PlotArea.Format.Fill.ForeColor.RGB = CHART_BG
        try:
            axis = chart.Axes(2)
            axis.HasMajorGridlines = True
            axis.MajorGridlines.Format.Line.ForeColor.RGB = GRID
            axis.MajorGridlines.Format.Line.Weight = 0.5
        except Exception:
            pass

    for area, names, values in (
        ("B18:H18", "$AA$7:$AA$11", "$AB$7:$AB$11"),
        ("J18:P18", "$AD$7:$AD$11", "$AE$7:$AE$11"),
    ):
        insight = ws.Range(area)
        insight.UnMerge()
        insight.MergeCells = True
        insight.Formula = (
            f'=IF(COUNT({values})=0,"暂无分项记录",'
            f'"最低分项得分率  ·  "&INDEX({names},MATCH(MIN({values}),{values},0))'
            f'&"  "&TEXT(MIN({values}),"0%"))'
        )
        insight.Interior.Color = (KPI_BG_BLUE if area.startswith("B")
                                  else KPI_BG_GREEN)
        insight.Font.Name = "微软雅黑"
        insight.Font.Size = 9
        insight.Font.Color = ACCENT_DARK
        insight.Font.Bold = True
        insight.VerticalAlignment = -4108
        for edge in (7, 8, 9, 10):
            insight.Borders(edge).LineStyle = 1
            insight.Borders(edge).Color = PANEL_EDGE

    for index, title in ((3, "Token 构成"), (4, "动作构成")):
        chart = ws.ChartObjects(index).Chart
        chart.HasTitle = True
        chart.ChartTitle.Text = title
        chart.HasLegend = True
        chart.Legend.Position = -4107
        chart.ChartArea.Format.Line.Visible = 0
        for axis_type in (1, 2):
            try:
                axis = chart.Axes(axis_type)
                axis.TickLabelPosition = 4
                axis.HasTitle = False
            except Exception:
                pass
        chart.PlotArea.Format.Fill.Visible = -1
        chart.PlotArea.Format.Fill.Solid()
        chart.ChartArea.Format.Fill.ForeColor.RGB = CHART_BG
        chart.PlotArea.Format.Fill.ForeColor.RGB = CHART_BG
        for idx in range(1, chart.SeriesCollection().Count + 1):
            series = chart.SeriesCollection(idx)
            series.Format.Fill.ForeColor.RGB = CATEGORY_COLORS[(idx - 1) % len(CATEGORY_COLORS)]
        _label_chart_series(chart, True)

    insight_header = ws.Range("R29:V29")
    insight_header.MergeCells = True
    insight_header.Value2 = "构成说明"
    insight_header.Interior.Color = NAVY
    insight_header.Font.Name = "微软雅黑"
    insight_header.Font.Size = 11
    insight_header.Font.Bold = True
    insight_header.Font.Color = rgb(255, 255, 255)
    insight_header.VerticalAlignment = -4108
    for row, label, formula, color in (
        (30, "缓存读取占比", '=IF(COUNT(U9:U12)=0,"",'
         'IFERROR(U10/SUM(U9:U12),0))', ACCENT_DARK),
        (32, "直接贡献＋必要探索", '=IF(COUNT(U15:U18)=0,"",'
         'IFERROR(SUM(U15:U16)/SUM(U15:U18),0))', ACCENT_DARK),
        (34, "确认无收益占比", '=IF(COUNT(U15:U18)=0,"",'
         'IFERROR(U18/SUM(U15:U18),0))', ACCENT_DARK),
    ):
        label_cell = ws.Range(f"R{row}:V{row}")
        label_cell.MergeCells = True
        label_cell.Value2 = label
        label_cell.Font.Name = "微软雅黑"
        label_cell.Font.Size = 9
        label_cell.Font.Color = MUTED
        value_cell = ws.Range(f"R{row+1}:V{row+1}")
        value_cell.MergeCells = True
        value_cell.Formula = formula
        value_cell.NumberFormat = "0.0%"
        value_cell.Font.Name = "微软雅黑"
        value_cell.Font.Size = 16
        value_cell.Font.Bold = True
        value_cell.Font.Color = color
        value_cell.VerticalAlignment = -4108
        value_cell.Borders(9).LineStyle = 1
        value_cell.Borders(9).Color = PANEL_EDGE
    note = ws.Range("R36:V36")
    note.MergeCells = True
    note.Value2 = "按当前 AI 与轮次的已记录项计算"
    note.Font.Name = "微软雅黑"
    note.Font.Size = 9
    note.Font.Color = MUTED
    note.VerticalAlignment = -4108


def _series_values(ser):
    vals = ser.Values
    if vals is None:
        return ()
    if isinstance(vals, (tuple, list)):
        return vals
    return (vals,)


def _flatten_values(vals):
    if vals is None:
        return
    if not isinstance(vals, (tuple, list)):
        yield vals
        return
    for item in vals:
        if isinstance(item, (tuple, list)):
            yield from item
        else:
            yield item


def _range_has_number(rng):
    try:
        for value in _flatten_values(rng.Value2):
            if isinstance(value, (int, float)):
                return True
    except Exception:
        return False
    return False


def _gray_heatmap(rng):
    if rng is None:
        return
    try:
        raw = rng.Value2
    except Exception:
        return
    rows = raw if isinstance(raw, (tuple, list)) else ((raw,),)
    if rows and not isinstance(rows[0], (tuple, list)):
        rows = tuple((row,) for row in rows)
    numeric = []
    for r_idx, row in enumerate(rows, 1):
        for c_idx, value in enumerate(row, 1):
            if isinstance(value, (int, float)):
                numeric.append((r_idx, c_idx, float(value)))
    if len(numeric) < 2:
        return
    low = min(item[2] for item in numeric)
    high = max(item[2] for item in numeric)
    span = (high - low) or 1.0
    for r_idx, c_idx, value in numeric:
        t = (value - low) / span
        shade = (
            int(round(HEAT_LO[i] + (HEAT_HI[i] - HEAT_LO[i]) * t))
            for i in range(3)
        )
        red, green, blue = tuple(shade)
        try:
            cell = rng.Cells(r_idx, c_idx)
            cell.Interior.Color = rgb(red, green, blue)
            cell.Font.Color = rgb(255, 255, 255) if t >= 0.62 else TEXT
        except Exception:
            pass


def _set_font(obj):
    try:
        obj.Name = FONT
    except Exception:
        pass
    try:
        obj.NameFarEast = FONT
    except Exception:
        pass


def _readable_report_type(wb):
    """Enlarge report type. Do not move tables or charts."""
    overview = wb.Worksheets("评测总览")
    overview.Range("B1:V1").Font.Size = 16
    overview.Range("B2:V2").Font.Size = 10
    overview.Range("B3:V3").Font.Size = 11
    overview.Range("B4:N4").Font.Size = 10
    overview.Range("B5:N6").Font.Size = 16
    overview.Range("B8:N8").Font.Size = 11
    overview.Range("B9:N16").Font.Size = 11
    overview.Range("B19:N19").Font.Size = 12
    overview.Range("B20:N20").Font.Size = 11
    overview.Range("B21:N26").Font.Size = 11
    overview.Range("B28:V28").Font.Size = 12

    single = wb.Worksheets("单AI分析")
    single.Range("B1:V1").Font.Size = 16
    single.Range("B2:V2").Font.Size = 10
    single.Range("B3:V3").Font.Size = 11
    single.Range("B4:V4").Font.Size = 10
    single.Range("B5:V6").Font.Size = 16
    single.Range("B18:P18").Font.Size = 11
    single.Range("B20:P20").Font.Size = 12
    single.Range("B21:P24").Font.Size = 11
    single.Range("B26:V26").Font.Size = 10
    single.Range("B27:V27").Font.Size = 12
    single.Range("B29:P33").Font.Size = 11
    single.Range("R8:V18").Font.Size = 11
    single.Range("R20:V20").Font.Size = 12
    note = single.Range("R21:V24")
    note.Font.Name = FONT
    note.Font.Size = 12
    note.Font.Color = TEXT
    note.WrapText = True
    note.VerticalAlignment = -4108
    note.HorizontalAlignment = -4131
    note.Value2 = (
        "Q / R 满分 100。AWR 不含管理动作。"
        "目录价不是实际账单，缺失保留为空或“—”。"
    )
    single.Range("R29:V29").Font.Size = 12
    single.Range("R30:S32").Font.Size = 11
    single.Range("T30:V32").Font.Size = 14
    single.Range("R33:V33").Font.Size = 10


def _paint_table_headers(wb):
    """Black fill + white type on table header rows only. Section banners stay light."""
    for name, areas in (
        ("评测总览", ("B8:N8", "B20:N20")),
        ("单AI分析", (
            "B20:P20", "B21:P21", "B29:H29", "J29:P29",
            "R8:V8", "R14:V14", "R20:V20", "R29:V29",
        )),
    ):
        ws = wb.Worksheets(name)
        ws.Range("B1:V1").Interior.Color = NAVY
        ws.Range("B1:V1").Font.Color = rgb(255, 255, 255)
        ws.Range("B1:V1").Font.Bold = True
        for area in areas:
            rng = ws.Range(area)
            rng.Interior.Color = NAVY
            rng.Font.Color = rgb(255, 255, 255)
            rng.Font.Bold = True
            rng.Font.Name = FONT


def _place_chart(ws, index, start, end):
    co = ws.ChartObjects(index)
    a, b = ws.Range(start), ws.Range(end)
    co.Left = a.Left
    co.Top = a.Top
    co.Width = b.Left + b.Width - a.Left
    co.Height = b.Top + b.Height - a.Top
    co.Placement = 3


def _expand_plot(ch):
    """Fill the frame with data; keep a gutter so end labels are readable."""
    try:
        n_series = ch.SeriesCollection().Count
    except Exception:
        n_series = 1
    stacked = ch.ChartType in (52, 53, 54, 58, 59, 60)
    try:
        ch.ChartGroups(1).GapWidth = 45
    except Exception:
        pass
    try:
        if (not stacked) and n_series == 1:
            ch.HasLegend = False
    except Exception:
        pass
    try:
        if ch.HasLegend:
            ch.Legend.Position = -4107
            ch.Legend.IncludeInLayout = True
            ch.Legend.Font.Size = 8
            ch.Legend.Font.Name = FONT
    except Exception:
        pass
    try:
        title_h = 16 if ch.HasTitle else 6
        legend_h = 18 if ch.HasLegend else 4
        left = 6 if stacked else 58
        right = 10 if stacked else 36
        pa = ch.PlotArea
        pa.Top = title_h
        pa.Left = left
        pa.Width = max(80, ch.ChartArea.Width - left - right)
        pa.Height = max(36, ch.ChartArea.Height - title_h - legend_h)
    except Exception:
        pass


def _tighten_density(wb):
    """Snap charts back to the packed cell grid. Do not retune row heights."""
    overview = wb.Worksheets("评测总览")
    single = wb.Worksheets("单AI分析")
    _place_chart(overview, 1, "P4", "V16")
    _place_chart(overview, 2, "P19", "V26")
    _place_chart(overview, 3, "B30", "K36")
    _place_chart(overview, 4, "M30", "V36")
    _place_chart(single, 1, "B8", "H17")
    _place_chart(single, 2, "J8", "P17")
    _place_chart(single, 3, "B35", "H41")
    _place_chart(single, 4, "J35", "P41")
    for name in REPORT_SHEETS:
        ws = wb.Worksheets(name)
        for i in range(1, ws.ChartObjects().Count + 1):
            _expand_plot(ws.ChartObjects(i).Chart)
    _paint_table_headers(wb)


def _reshape_composition_chart(ws, ch, chart_index):
    """Part-to-whole as one 100% stacked bar, not four sparse bars."""
    if chart_index == 3:
        labels = ("未缓存输入", "缓存读取", "缓存写入", "输出")
        cells = ("U9", "U10", "U11", "U12")
    else:
        labels = ("直接贡献", "必要探索", "管理动作", "确认无收益")
        cells = ("U15", "U16", "U17", "U18")
    ch.ChartType = 59
    while ch.SeriesCollection().Count > 0:
        ch.SeriesCollection(ch.SeriesCollection().Count).Delete()
    raw = []
    for idx, (lab, cell) in enumerate(zip(labels, cells), 1):
        ch.SeriesCollection().NewSeries()
        ser = ch.SeriesCollection(idx)
        ser.Name = lab
        ser.Values = ws.Range(cell)
        ser.XValues = (" ",)
        ser.Format.Line.Visible = 0
        ser.Format.Fill.ForeColor.RGB = CATEGORY_COLORS[(idx - 1) % len(CATEGORY_COLORS)]
        try:
            raw.append(float(ws.Range(cell).Value2 or 0))
        except (TypeError, ValueError):
            raw.append(0.0)
    total = sum(raw)
    try:
        ch.ChartGroups(1).GapWidth = 28
    except Exception:
        pass
    ch.HasTitle = False
    ch.HasLegend = True
    ch.Legend.Position = -4107
    ch.Legend.IncludeInLayout = True
    ch.Legend.Font.Name = FONT
    ch.Legend.Font.Size = 8
    try:
        ch.Axes(1).TickLabelPosition = -4142
        ch.Axes(1).Format.Line.Visible = 0
        ch.Axes(2).TickLabelPosition = -4142
        ch.Axes(2).HasMajorGridlines = False
        ch.Axes(2).Format.Line.Visible = 0
        ch.Axes(2).MinimumScale = 0
        ch.Axes(2).MaximumScale = 1
    except Exception:
        pass
    for idx, number in enumerate(raw, 1):
        ser = ch.SeriesCollection(idx)
        share = (number / total) if total else 0
        if share < 0.08:
            ser.HasDataLabels = False
            continue
        ser.HasDataLabels = True
        ser.Points(1).HasDataLabel = True
        label = ser.Points(1).DataLabel
        label.Text = f"{share:.0%}"
        label.Position = -4108
        label.Font.Name = FONT
        label.Font.Size = 10
        label.Font.Color = rgb(255, 255, 255) if idx in (1, 4) else TEXT


def _apply_font_report(wb):
    for name in REPORT_SHEETS:
        ws = wb.Worksheets(name)
        try:
            ws.Range("B1:V36").Font.Name = FONT
        except Exception:
            pass
        for chart_index in range(1, ws.ChartObjects().Count + 1):
            ch = ws.ChartObjects(chart_index).Chart
            try:
                if ch.HasTitle:
                    ch.ChartTitle.Font.Name = FONT
            except Exception:
                pass
            try:
                if ch.HasLegend:
                    ch.Legend.Font.Name = FONT
            except Exception:
                pass
            for axis_type in (1, 2):
                try:
                    ch.Axes(axis_type).TickLabels.Font.Name = FONT
                except Exception:
                    pass
            try:
                for series_index in range(1, ch.SeriesCollection().Count + 1):
                    series = ch.SeriesCollection(series_index)
                    if series.HasDataLabels:
                        series.DataLabels().Font.Name = FONT
            except Exception:
                pass


def _apply_heatmaps(wb):
    def by_column(ws, area):
        rng = ws.Range(area)
        for col in range(1, rng.Columns.Count + 1):
            _gray_heatmap(rng.Columns(col))

    overview = wb.Worksheets("评测总览")
    by_column(overview, "C9:E14")
    by_column(overview, "G9:I14")
    by_column(overview, "C21:M26")
    single = wb.Worksheets("单AI分析")
    for area in ("U9:U12", "U15:U18", "D22:P24",
                 "E30:E33", "M30:M33", "F30:H33", "N30:P33",
                 "T30:V32"):
        by_column(single, area)


def _force_solid_fill(obj, color):
    try:
        fill = obj.Format.Fill
        fill.Visible = -1
        fill.Solid()
        fill.ForeColor.RGB = color
    except Exception:
        pass
    try:
        obj.Format.Line.Visible = 0
    except Exception:
        pass


def _label_chart_series(ch, stacked):
    try:
        ch.ChartGroups(1).GapWidth = 28 if stacked else 55
    except Exception:
        pass
    n_series = ch.SeriesCollection().Count
    if n_series < 1:
        return
    point_totals = None
    series_max = 1.0
    if stacked:
        first_vals = _series_values(ch.SeriesCollection(1))
        point_totals = [0.0] * max(1, len(first_vals))
        for idx in range(1, n_series + 1):
            for point_index, value in enumerate(_series_values(ch.SeriesCollection(idx))):
                try:
                    point_totals[point_index] += float(value or 0)
                except (TypeError, ValueError, IndexError):
                    pass
    else:
        nums = []
        for idx in range(1, n_series + 1):
            for value in _series_values(ch.SeriesCollection(idx)):
                try:
                    nums.append(abs(float(value or 0)))
                except (TypeError, ValueError):
                    pass
        series_max = max(nums) if nums else 1.0
    dark_fills = {CHART_BLUE, CHART_LILAC, CHART_CORAL}
    for idx in range(1, n_series + 1):
        ser = ch.SeriesCollection(idx)
        try:
            ser.HasDataLabels = True
            labels = ser.DataLabels()
            labels.ShowCategoryName = False
            labels.ShowSeriesName = False
            try:
                labels.ShowValue = not stacked
            except Exception:
                pass
            if stacked:
                position = -4108
            else:
                position = 3
            try:
                labels.Position = position
            except Exception:
                pass
            labels.Font.Name = FONT
            labels.Font.Size = 9
            labels.Font.Bold = False
            fill = CHART_BLUE
            try:
                fill = ser.Format.Fill.ForeColor.RGB
            except Exception:
                pass
            for point_index, value in enumerate(_series_values(ser), 1):
                try:
                    numeric = float(value)
                except (TypeError, ValueError):
                    numeric = None
                hide = numeric is None or abs(numeric) < 1e-12
                share = 0
                if stacked and numeric is not None and point_totals:
                    total = point_totals[point_index - 1]
                    share = (numeric / total) if total else 0
                    hide = hide or share < 0.08
                try:
                    ser.Points(point_index).HasDataLabel = not hide
                except Exception:
                    continue
                if hide:
                    continue
                try:
                    data_label = ser.Points(point_index).DataLabel
                    if stacked:
                        data_label.Text = f"{share:.0%}"
                    data_label.Font.Name = FONT
                    data_label.Font.Size = 9
                    data_label.Font.Bold = False
                    short = (not stacked) and numeric is not None and abs(numeric) < 0.22 * series_max
                    try:
                        pt_fill = ser.Points(point_index).Format.Fill.ForeColor.RGB
                    except Exception:
                        pt_fill = fill
                    if short:
                        try:
                            data_label.Position = 2
                        except Exception:
                            pass
                        data_label.Font.Color = TEXT
                    elif pt_fill in dark_fills:
                        try:
                            data_label.Position = 3
                        except Exception:
                            pass
                        data_label.Font.Color = rgb(255, 255, 255)
                    else:
                        try:
                            data_label.Position = 3
                        except Exception:
                            pass
                        data_label.Font.Color = TEXT
                except Exception:
                    pass
        except Exception:
            pass


def _polish_charts(wb):
    placeholders = {"", "坐标轴标题", "Axis Title", "图表标题", "Chart Title"}
    for sheet_name in REPORT_SHEETS:
        ws = wb.Worksheets(sheet_name)
        for i in range(1, ws.ChartObjects().Count + 1):
            ch = ws.ChartObjects(i).Chart
            try:
                if ch.HasTitle and str(ch.ChartTitle.Text or "") in placeholders:
                    titles = {
                        1: "成果质量与可靠性",
                        2: "完成耗时（分钟）",
                        3: "费用参考值（元）" if sheet_name == "评测总览" else "Token 构成",
                        4: "无收益动作占比（AWR）" if sheet_name == "评测总览" else "动作构成",
                    }
                    ch.ChartTitle.Text = titles.get(i, ch.ChartTitle.Text)
            except Exception:
                pass
            for axis_type in (1, 2):
                try:
                    ch.Axes(axis_type).HasTitle = False
                except Exception:
                    pass
            composition = sheet_name == "单AI分析" and i in (3, 4)
            if composition:
                _reshape_composition_chart(ws, ch, i)
            else:
                try:
                    if sheet_name == "评测总览" and i in (3, 4):
                        ch.ChartType = 57
                except Exception:
                    pass
                stacked = ch.ChartType in (52, 53, 54, 58, 59, 60)
                palette = MODEL_COLORS if sheet_name == "评测总览" else CATEGORY_COLORS
                for idx in range(1, ch.SeriesCollection().Count + 1):
                    try:
                        ser = ch.SeriesCollection(idx)
                        color = palette[(idx - 1) % len(palette)]
                        if (not stacked) and ch.SeriesCollection().Count == 1:
                            color = CHART_BLUE
                        _force_solid_fill(ser, color)
                    except Exception:
                        pass
                if (not stacked) and ch.SeriesCollection().Count == 1:
                    try:
                        point_palette = (
                            MODEL_COLORS if sheet_name == "评测总览" else CATEGORY_COLORS
                        )
                        for point_index, color in enumerate(point_palette, 1):
                            _force_solid_fill(
                                ch.SeriesCollection(1).Points(point_index),
                                color,
                            )
                    except Exception:
                        pass
                _label_chart_series(ch, stacked)
                n_series = ch.SeriesCollection().Count
                need_legend = stacked or n_series > 1
                ch.HasLegend = need_legend
                if need_legend:
                    ch.Legend.Position = -4107
                    ch.Legend.IncludeInLayout = True
                    try:
                        ch.Legend.Font.Name = FONT
                        ch.Legend.Font.Size = 8
                        ch.Legend.Font.Bold = False
                    except Exception:
                        pass
            try:
                if ch.HasTitle:
                    ch.ChartTitle.Font.Name = FONT
                    ch.ChartTitle.Font.Size = 11
            except Exception:
                pass
            if not composition:
                for axis_type in (1, 2):
                    try:
                        axis = ch.Axes(axis_type)
                        axis.TickLabels.Font.Name = FONT
                        axis.TickLabels.Font.Size = 9
                        axis.HasTitle = False
                    except Exception:
                        pass
                try:
                    ch.Axes(1).ReversePlotOrder = True
                except Exception:
                    pass
                try:
                    axis = ch.Axes(2)
                    axis.HasMajorGridlines = True
                    axis.MajorGridlines.Format.Line.ForeColor.RGB = GRID
                    axis.MajorGridlines.Format.Line.Weight = 0.5
                    axis.MinimumScale = 0
                    axis.MaximumScaleIsAuto = True
                    if sheet_name == "单AI分析" and i in (1, 2):
                        axis.MaximumScale = 1
                    elif sheet_name == "评测总览" and i == 1:
                        axis.MaximumScale = 100
                except Exception:
                    pass
            try:
                line = ch.ChartArea.Format.Line
                line.Visible = -1
                line.ForeColor.RGB = GRID
                line.Weight = 0.75
            except Exception:
                pass
            _expand_plot(ch)


def _hide_report_gridlines(wb):
    xl = wb.Application
    current = wb.ActiveSheet.Name
    for name in REPORT_SHEETS:
        wb.Worksheets(name).Activate()
        try:
            xl.ActiveWindow.DisplayGridlines = False
            xl.ActiveWindow.Zoom = 80
        except Exception:
            pass
    try:
        wb.Worksheets(current).Activate()
    except Exception:
        wb.Worksheets("评测总览").Activate()


def apply_report_updates(wb):
    overview = wb.Worksheets("评测总览")
    overview.Range("E3").Value2 = "▼ 下拉切换"
    overview.Range("E3").Font.Size = 9
    overview.Range("E3").Font.Color = MUTED

    _section_like(
        overview,
        "B19",
        "B28:V28",
        "成本与动作效率",
    )
    overview.Rows(28).RowHeight = 22

    add_single_bar(
        overview,
        "费用参考值（元）",
        "'评测总览'!$B$9:$B$14",
        "'评测总览'!$Z$9:$Z$14",
        "B30", "K36",
    )
    add_single_bar(
        overview,
        "无收益动作占比（AWR）",
        "'评测总览'!$B$9:$B$14",
        "'评测总览'!$H$9:$H$14",
        "M30", "V36",
        pct=True,
    )

    single = wb.Worksheets("单AI分析")
    for cell in ("E3", "I3"):
        single.Range(cell).Value2 = "▼ 下拉选择"
        single.Range(cell).Font.Size = 9
        single.Range(cell).Font.Color = MUTED

    single.Range("H5").Formula = (
        '=IF(SUMPRODUCT((运行记录!$A$2:$A$16=$C$3)'
        '*(运行记录!$B$2:$B$16=$G$3)'
        '*ISNUMBER(运行记录!$M$2:$M$16))=0,"—",'
        'SUMIFS(运行记录!$M$2:$M$16,'
        '运行记录!$A$2:$A$16,$C$3,'
        '运行记录!$B$2:$B$16,$G$3))'
    )

    _section_like(
        single,
        "B20",
        "B27:V27",
        "Token 与动作结构",
    )
    single.Rows(27).RowHeight = 22

    add_stack(
        single,
        "Token 构成",
        ["未缓存输入", "缓存读取", "缓存写入", "输出"],
        ["U9", "U10", "U11", "U12"],
        "B29", "H39",
        [CHART_BLUE, CHART_MINT, CHART_AMBER, CHART_LILAC],
        object_name="chart_token_composition",
    )
    add_stack(
        single,
        "动作构成",
        ["直接贡献", "必要探索", "管理动作", "确认无收益"],
        ["U15", "U16", "U17", "U18"],
        "J29", "P39",
        [CHART_BLUE, CHART_LIME, CHART_AMBER, CHART_LILAC],
        object_name="chart_action_composition",
    )

    _format_report_pages(wb)
    _apply_visual_system(wb)
    _pack_single_page(wb)
    _polish_charts(wb)
    _apply_heatmaps(wb)
    _apply_font_report(wb)
    _uniform_empty_cells(wb)
    _format_report_pages(wb)
    _readable_report_type(wb)
    _paint_table_headers(wb)

    # Final minimal repair: keep the overview comparison row labels intact and
    # preserve native worksheet gridlines on both reader-facing report pages.
    overview = wb.Worksheets("评测总览")
    try:
        overview.Range("B21:C26").UnMerge()
    except Exception:
        pass
    for row in range(21, 27):
        label = overview.Range(f"B{row}:C{row}")
        label.MergeCells = True
        overview.Range(f"B{row}").Formula = f"=B{row - 12}"
        label.HorizontalAlignment = -4131
        label.VerticalAlignment = -4108
        label.Font.Name = FONT
        label.Font.Size = 11
    _paint_table_black(overview, "B20:N26")

    single = wb.Worksheets("单AI分析")
    overview.PageSetup.PrintArea = "$A$1:$V$36"
    single.PageSetup.PrintArea = "$A$1:$W$41"

    for ws in (overview, single):
        ws.Activate()
        wb.Application.ActiveWindow.DisplayGridlines = True

def apply_research_table(wb):
    """在现有总览下追加配对分项表；保持原图表、数据表和选择器。"""
    spec = json.loads((ROOT / "_实验系统" / "报告素材" / "研究表定义.json").read_text(encoding="utf-8"))
    ws = wb.Worksheets(spec["sheet"])
    ws.Range("B39:V50").UnMerge()
    ws.Range("B39:V50").Clear()
    ws.Range("B39:V50").Font.Name = FONT
    ws.Range("B39:V50").Font.Size = 11
    ws.Rows("39:50").RowHeight = 25
    ws.Range("B39:N39").Merge()
    ws.Range("B39").Value2 = spec["title"]
    ws.Range("B39:N40").Interior.Color = rgb(235,235,235)
    ws.Range("B39:N40").Font.Bold = True
    for row, values in enumerate([spec["headers"]] + spec["rows"],40):
        for block, value in zip(spec["blocks"],values):
            left,right=block.split(":")
            rg=ws.Range(f"{left}{row}:{right}{row}")
            rg.Merge()
            rg.HorizontalAlignment = -4131 if left == "B" else -4108
            rg.VerticalAlignment = -4108
            rg.Borders.LineStyle = 1
            rg.Borders.Color = rgb(210,210,210)
            cell=ws.Range(f"{left}{row}")
            if isinstance(value,str) and value.startswith("="): cell.Formula=value
            else: cell.Value2=value
            if row>40: rg.NumberFormat="0.0%" if left=="M" else "0.0"
    ws.Range("B46:N46").Font.Bold=True
    for row,note in enumerate(spec["notes"],47):
        ws.Range(f"B{row}:V{row}").Merge()
        ws.Range(f"B{row}").Value2=note
        ws.Range(f"B{row}:V{row}").Font.Size=10
    ws.PageSetup.PrintArea="$A$1:$V$50"
    ws.PageSetup.FitToPagesWide=1
    ws.PageSetup.FitToPagesTall=False


def put_matrix(ws, matrix):
    rows = len(matrix)
    cols = len(matrix[0])
    ws.Range(ws.Cells(1, 1), ws.Cells(rows, cols)).Value = tuple(
        tuple(row) for row in matrix
    )

def resize_table(ws, table_name, ref):
    try:
        table = ws.ListObjects.Item(table_name)
        table.Resize(ws.Range(ref))
    except Exception:
        table = ws.ListObjects.Add(1, ws.Range(ref), None, 1)
        table.Name = table_name
        table.TableStyle = "TableStyleLight1"

def refresh_data_com(wb, tables, field_rows):
    data = dict(tables)

    ws = wb.Worksheets("运行记录")
    headers = list(data["运行记录"][0])
    matrix = [headers] + [[row.get(h) for h in headers] for row in data["运行记录"]]
    ws.Range("A2:R1000").ClearContents()
    put_matrix(ws, matrix)
    last = len(matrix)
    ws.Range("M1:R1").Value = (("原生分钟", "成果质量分", "独立可靠性分",
                                "总Token", "API费用低（元）", "API费用高（元）"),)
    ws.Range("M2").Formula = '=IF(ISNUMBER(H2),H2/60,"")'
    ws.Range(f"M2:M{last}").FillDown()
    ws.Range("N2").Formula = '=IF(COUNTIFS(评分与用量!$H:$H,J2,评分与用量!$C:$C,"质量分项")=21,SUMIFS(评分与用量!$E:$E,评分与用量!$H:$H,J2,评分与用量!$C:$C,"质量分项"),"")'
    ws.Range(f"N2:N{last}").FillDown()
    ws.Range("O2").Formula = '=IF(COUNTIFS(评分与用量!$H:$H,J2,评分与用量!$C:$C,"可靠性分项")=5,SUMIFS(评分与用量!$E:$E,评分与用量!$H:$H,J2,评分与用量!$C:$C,"可靠性分项"),"")'
    ws.Range(f"O2:O{last}").FillDown()
    ws.Range("P2").Formula = '=IF(COUNTIFS(评分与用量!$H:$H,J2,评分与用量!$C:$C,"Token用量",评分与用量!$D:$D,"<>推理Token（输出子集）")=4,SUMIFS(评分与用量!$E:$E,评分与用量!$H:$H,J2,评分与用量!$C:$C,"Token用量",评分与用量!$D:$D,"<>推理Token（输出子集）"),"")'
    ws.Range(f"P2:P{last}").FillDown()
    ws.Range("Q2").Formula = '=IF(COUNTIFS(评分与用量!$H:$H,J2,评分与用量!$C:$C,"原生计价",评分与用量!$D:$D,"费用低")=1,SUMIFS(评分与用量!$E:$E,评分与用量!$H:$H,J2,评分与用量!$C:$C,"原生计价",评分与用量!$D:$D,"费用低"),IF(COUNTIFS(API计价!$L:$L,J2,API计价!$D:$D,"低")=4,SUMIFS(API计价!$J:$J,API计价!$L:$L,J2,API计价!$D:$D,"低"),""))'
    ws.Range(f"Q2:Q{last}").FillDown()
    ws.Range("R2").Formula = '=IF(COUNTIFS(评分与用量!$H:$H,J2,评分与用量!$C:$C,"原生计价",评分与用量!$D:$D,"费用高")=1,SUMIFS(评分与用量!$E:$E,评分与用量!$H:$H,J2,评分与用量!$C:$C,"原生计价",评分与用量!$D:$D,"费用高"),IF(COUNTIFS(API计价!$L:$L,J2,API计价!$D:$D,"高")=4,SUMIFS(API计价!$J:$J,API计价!$L:$L,J2,API计价!$D:$D,"高"),""))'
    ws.Range(f"R2:R{last}").FillDown()
    resize_table(ws, "tblRuns", f"A1:R{last}")

    simple = [
        ("评分与用量", "tblMetrics"),
        ("API计价", "tblPricing"),
        ("行为记录", "tblActions"),
    ]
    for name, table_name in simple:
        ws = wb.Worksheets(name)
        headers = list(data[name][0])
        matrix = [headers] + [[row.get(h) for h in headers] for row in data[name]]
        ws.Range("A2:Z5000").ClearContents()
        put_matrix(ws, matrix)
        last = len(matrix)
        end = ws.Cells(last, len(headers))
        resize_table(ws, table_name, ws.Range(ws.Cells(1, 1), end).Address)

    ws = wb.Worksheets("核验记录")
    headers = list(data["核验记录"][0])
    matrix = [headers] + [[row.get(h) for h in headers] for row in data["核验记录"]]
    ws.Range("A2:W1000").ClearContents()
    put_matrix(ws, matrix)
    last = len(matrix)
    ws.Range("W1").Value2 = "数值差（对照减原值）"
    ws.Range("W2").Formula = '=IF(COUNT(F2:G2)=2,G2-F2,"")'
    ws.Range(f"W2:W{last}").FillDown()
    resize_table(ws, "tblChecks", f"A1:W{last}")

    ws = wb.Worksheets("字段说明")
    headers = ["表名", "字段", "含义", "允许值", "空值规则"]
    matrix = [headers] + [[row.get(h) for h in headers] for row in field_rows]
    ws.Range("A2:E1000").ClearContents()
    put_matrix(ws, matrix)
    last = len(matrix)
    resize_table(ws, "tblFields", f"A1:E{last}")

def apply_comparison_tables(wb):
    """生成报告使用的维度对照表，并统一工作簿表格样式。"""
    spec = json.loads((ROOT / "_实验系统" / "报告素材" / "比较表定义.json").read_text(encoding="utf-8"))
    try:
        ws = wb.Worksheets("分维度比较")
    except Exception:
        ws = wb.Worksheets.Add(After=wb.Worksheets("单AI分析"))
        ws.Name = "分维度比较"
    ws.Cells.Clear()
    ws.Columns("A").ColumnWidth=24
    ws.Columns("B:G").ColumnWidth=13
    ws.Columns("H").ColumnWidth=24
    for block in spec["tables"]:
        start=block["start"]
        ws.Range(f"A{start}:H{start}").Merge()
        ws.Cells(start,1).Value2=block["title"]
        for i,row in enumerate(block["rows"],start+1):
            for j,val in enumerate(row,1):
                if isinstance(val,str) and val.startswith("="): ws.Cells(i,j).Formula=val
                elif val is not None: ws.Cells(i,j).Value2=val
        ws.Range(f"A{start}:H{start+1}").Font.Bold=True
        ws.Range(f"A{start}:H{start+1}").Interior.Color=rgb(64,64,64)
        ws.Range(f"A{start}:H{start+1}").Font.Color=rgb(255,255,255)
    ws.UsedRange.Font.Name=FONT
    ws.UsedRange.Font.Size=11
    ws.UsedRange.RowHeight=25
    ws.UsedRange.VerticalAlignment=-4108
    ws.UsedRange.Borders.LineStyle=1
    ws.UsedRange.Borders.Color=rgb(210,210,210)
    ws.PageSetup.PrintArea="$A$1:$H$161"
    ws.PageSetup.Orientation=2
    ws.PageSetup.Zoom=False
    ws.PageSetup.FitToPagesWide=1
    ws.PageSetup.FitToPagesTall=False
    for block in spec["tables"]:
        row=block["start"]+1
        ws.Rows(row).RowHeight=38
        ws.Range(f"A{row}:H{row}").WrapText=True
    for sheet_name in DATA_SHEETS:
        for table in wb.Worksheets(sheet_name).ListObjects:
            table.TableStyle="TableStyleLight1"
            table.Range.Font.Name=FONT
            table.Range.Font.Size=11
            table.Range.FormatConditions.Delete()
            table.Range.Interior.Color=rgb(255,255,255)
            table.Range.Font.Color=rgb(0,0,0)
            table.Range.Borders.LineStyle=1
            table.Range.Borders.Color=rgb(210,210,210)
            table.HeaderRowRange.Interior.Color=rgb(64,64,64)
            table.HeaderRowRange.Font.Color=rgb(255,255,255)
            table.HeaderRowRange.Font.Bold=True
            table.HeaderRowRange.RowHeight=30
            table.HeaderRowRange.WrapText=True
    areas={"评测总览":[("B8:N16","B8:N8"),("B20:N26","B20:N20"),("B40:N46","B40:N40")],"单AI分析":[("B21:P24","B21:P21")]}
    for name, ranges in areas.items():
        sheet=wb.Worksheets(name)
        for area,head in ranges:
            sheet.Range(area).FormatConditions.Delete()
            sheet.Range(area).Interior.Color=rgb(255,255,255)
            sheet.Range(area).Font.Color=rgb(0,0,0)
            sheet.Range(area).Borders.LineStyle=1
            sheet.Range(area).Borders.Color=rgb(210,210,210)
            sheet.Range(head).Interior.Color=rgb(64,64,64)
            sheet.Range(head).Font.Color=rgb(255,255,255)
            sheet.Range(head).Font.Bold=True
    wb.Worksheets("评测总览").Range("B39:N39").Interior.Color=rgb(64,64,64)
    wb.Worksheets("评测总览").Range("B39:N39").Font.Color=rgb(255,255,255)


def finish_note_tables(wb):
    """把说明纳入表格，保留图表引用并隐藏辅助计算列。"""
    def style(ws, area, header=False):
        r=ws.Range(area)
        r.Font.Name=FONT
        r.Font.Size=11
        r.Font.Color=rgb(255,255,255) if header else rgb(0,0,0)
        r.Font.Bold=header
        r.Interior.Color=rgb(64,64,64) if header else rgb(255,255,255)
        r.Borders.LineStyle=1
        r.Borders.Color=rgb(210,210,210)
        r.WrapText=True
        r.VerticalAlignment=-4108
    def note(ws,row,label,text,last='N'):
        ws.Range(f'B{row}:{last}{row}').UnMerge()
        ws.Range(f'B{row}:{last}{row}').ClearContents()
        ws.Range(f'B{row}:D{row}').Merge()
        ws.Range(f'E{row}:{last}{row}').Merge()
        ws.Range(f'B{row}').Value2=label
        ws.Range(f'E{row}').Value2=text
        style(ws,f'B{row}:D{row}',True)
        style(ws,f'E{row}:{last}{row}')
        ws.Rows(row).RowHeight=44
    ov=wb.Worksheets('评测总览')
    for name in ('评测总览','单AI分析'):
        wb.Worksheets(name).Columns('B:V').ColumnWidth=10
    ov.Range('E3').Value2='▼'
    for row,label in [(18,'比较说明'),(47,'比较范围'),(48,'计算方法'),(49,'补充比较'),(50,'解释范围')]:
        text=ov.Range(f'E{row}').Value2 or ov.Range(f'B{row}').Value2
        note(ov,row,label,text)
    ws=wb.Worksheets('单AI分析')
    ws.Range('E3').Value2='▼'
    ws.Range('I3').Value2='▼'
    text=ws.Range('E26').Value2 or ws.Range('B26').Value2
    note(ws,26,'比较说明',text,'V')
    ws.Range('R21:V24').UnMerge()
    ws.Range('R21:V24').ClearContents()
    for row,label,text in [(21,'Q / R','满分均为 100'),(22,'AWR','不含管理动作'),(23,'目录价','供成本比较，非实际账单'),(24,'缺失值','保留空白或“—”')]:
        ws.Range(f'R{row}:S{row}').Merge()
        ws.Range(f'T{row}:V{row}').Merge()
        ws.Range(f'R{row}').Value2=label
        ws.Range(f'T{row}').Value2=text
        style(ws,f'R{row}:V{row}')
        ws.Rows(row).RowHeight=34
    vals=[ws.Range(f'U{r}').Formula or ws.Range(f'T{r}').Formula for r in (30,31,32)]
    ws.Range('R30:V33').UnMerge()
    ws.Range('R30:V33').ClearContents()
    for row,label,formula in zip((30,31,32),('缓存读取占比','直接贡献＋必要探索','确认无收益占比'),vals):
        ws.Range(f'R{row}:T{row}').Merge()
        ws.Range(f'U{row}:V{row}').Merge()
        ws.Range(f'R{row}').Value2=label
        ws.Range(f'U{row}').Formula=formula
        ws.Range(f'U{row}:V{row}').NumberFormat='0.0%'
        style(ws,f'R{row}:V{row}')
        ws.Rows(row).RowHeight=32
    ws.Range('R33:V33').Merge()
    ws.Range('R33').Value2='计算范围：当前 AI、当前轮次的已记录项'
    style(ws,'R33:V33')
    ws.Rows(33).RowHeight=36
    for area in ['B18:H18','J18:P18']:
        style(ws,area)
    # 两张图独立成块，标题不跨越中间留白。
    ov.Range('B28:V28').UnMerge()
    ov.Range('B28:V28').Clear()
    for area,title in [('B28:K28','API 目录价'),('M28:V28','无收益动作占比')]:
        ov.Range(area).Merge()
        ov.Range(area.split(':')[0]).Value2=title
        style(ov,area,True)
    ov.Rows('1:50').RowHeight=16
    ov.Rows(1).RowHeight=23
    ov.Rows('3:6').RowHeight=18
    ov.Rows('8:16').RowHeight=18
    ov.Rows(18).RowHeight=32
    ov.Rows('19:26').RowHeight=18
    ov.Rows('29:36').RowHeight=18
    ov.Rows('37:38').RowHeight=8
    ov.Rows('39:46').RowHeight=20
    ov.Rows('47:50').RowHeight=30
    ws.Rows('1:41').RowHeight=17
    ws.Rows(1).RowHeight=23
    ws.Rows('3:6').RowHeight=20
    ws.Rows('21:24').RowHeight=30
    ws.Rows(26).RowHeight=32
    ws.Rows('30:33').RowHeight=28
    for index,area in [(1,'B8:H17'),(2,'J8:P17'),(3,'B35:H41'),(4,'J35:P41')]:
        chart=ws.ChartObjects(index)
        target=ws.Range(area)
        chart.Left=target.Left
        chart.Top=target.Top
        chart.Width=target.Width
        chart.Height=target.Height
    for index,area in [(1,'P4:V16'),(2,'P19:V26'),(3,'B29:K36'),(4,'M29:V36')]:
        chart=ov.ChartObjects(index)
        target=ov.Range(area)
        chart.Left=target.Left
        chart.Top=target.Top
        chart.Width=target.Width
        chart.Height=target.Height
    for sheet in (ov,ws):
        for chart in sheet.ChartObjects():
            chart.Chart.PlotVisibleOnly=False
            chart.Placement=2
        sheet.Columns('X:XFD').Hidden=True
        sheet.Activate()
        wb.Application.ActiveWindow.DisplayGridlines=False
        wb.Application.ActiveWindow.ScrollColumn=1
        wb.Application.ActiveWindow.ScrollRow=1
        wb.Application.ActiveWindow.Zoom=80
    ov.Activate()
    ov.Range('A1').Select()
    # 总览改成左侧表格、右侧四图、底部比较表与说明的横向矩形。
    notes=[ov.Range(f'E{r}').Value2 for r in range(47,51)]
    ov.Range('B39:N46').Cut(ov.Range('B29'))
    ov.Range('P28:AD36').UnMerge()
    ov.Range('P28:AD36').Clear()
    ov.Range('P28:AD28').Merge()
    ov.Range('P28').Value2='比较说明'
    style(ov,'P28:AD28',True)
    for row,label,text in zip((29,31,33,35),('比较范围','计算方法','补充比较','解释范围'),notes):
        ov.Range(f'P{row}:R{row+1}').Merge()
        ov.Range(f'S{row}:AD{row+1}').Merge()
        ov.Range(f'P{row}').Value2=label
        ov.Range(f'S{row}').Value2=text
        style(ov,f'P{row}:R{row+1}',True)
        style(ov,f'S{row}:AD{row+1}')
    ov.Range('B37:V50').UnMerge()
    ov.Range('B37:V50').Clear()
    ov.Columns('B:AD').ColumnWidth=8.7
    ov.Columns('X:AD').Hidden=False
    ov.Columns('AE:XFD').Hidden=True
    ov.Rows('1:36').RowHeight=17
    ov.Rows(1).RowHeight=24
    ov.Rows(18).RowHeight=34
    ov.Rows(28).RowHeight=10
    ov.Rows('29:36').RowHeight=23
    for index,area in [(1,'P4:V14'),(2,'X4:AD14'),(3,'P16:V26'),(4,'X16:AD26')]:
        chart=ov.ChartObjects(index);target=ov.Range(area)
        chart.Left=target.Left;chart.Top=target.Top
        chart.Width=target.Width;chart.Height=target.Height
    ov.Range('B28:N28').Clear()
    ov.PageSetup.PrintArea='$A$1:$AD$36'
    ov.Activate()
    wb.Application.ActiveWindow.Zoom=85
    ws.Rows('21:24').RowHeight=24
    ws.Rows('30:33').RowHeight=24
    for index,area in [(3,'B35:H41'),(4,'J35:P41')]:
        chart=ws.ChartObjects(index);target=ws.Range(area)
        chart.Top=target.Top;chart.Height=target.Height
    ws.Activate()
    wb.Application.ActiveWindow.Zoom=75
    ov.Activate()
    ov.Range('A1').Select()


def align_rectangle(wb):
    s=wb.Worksheets('评测总览')
    s.Range('X1:AZ27').Cut(s.Range('BA1'))
    s.Range('X1:AD27').UnMerge()
    s.Range('X1:AD27').Clear()
    s.Range('R1:V1').UnMerge()
    s.Range('R1:V1').Clear()
    s.Range('R1:AD1').Merge()
    s.Range('R1').Value2='查看单个 AI  ›'
    s.Range('R1:AD1').Interior.Color=rgb(32,32,32)
    s.Range('R1:AD1').Font.Color=rgb(255,255,255)
    s.Rows(28).RowHeight=20
    s.Columns('B:AD').Hidden=False
    s.Columns('AE:XFD').Hidden=True
    for i,a in [(1,'P4:V14'),(2,'X4:AD14'),(3,'P16:V26'),(4,'X16:AD26')]:
        c=s.ChartObjects(i);t=s.Range(a)
        c.Width=t.Width;c.Height=t.Height;c.Left=t.Left;c.Top=t.Top


def finish_screen_layout(wb):
    """屏幕布局：标题与刻度分开，保存经过真实窗口检查的比例。"""
    for name,zoom in [('评测总览',70),('单AI分析',65)]:
        sheet=wb.Worksheets(name)
        report=sheet.Range('B1:AD36' if name=='评测总览' else 'B1:V41')
        report.Font.Name='微软雅黑'
        report.Font.Size=14
        report.Font.Bold=True
        sheet.Range('B1:P1').Font.Size=20
        sheet.Range('B2:V2').Font.Size=11
        sheet.Range('B3:V3').Font.Size=12
        if name=='评测总览':
            sheet.Rows('8:16').RowHeight=19
            sheet.Rows('20:26').RowHeight=19
            for area in ['B18:N18','P28:AD36']:
                sheet.Range(area).Font.Size=12
            sheet.Range('B29:N36').Font.Size=12
        else:
            for area in ['B26:V26','R21:V24','R30:V33']:
                sheet.Range(area).Font.Size=12
            sheet.Range('B18:P18').Font.Size=11
            sheet.Range('B5:V6').Font.Size=22
        if name=='评测总览':
            sheet.Columns('B').ColumnWidth=13
            sheet.Columns('C:D').ColumnWidth=7.5
        for obj in sheet.ChartObjects():
            chart=obj.Chart
            chart.ChartArea.Font.Name='微软雅黑'
            chart.ChartArea.Font.Size=12
            chart.ChartArea.Font.Bold=True
            chart.ChartArea.AutoScaleFont=False
            if chart.HasTitle:
                chart.ChartTitle.Font.Size=12
                chart.ChartTitle.Top=5
                chart.ChartTitle.Left=max(0,(obj.Width-chart.ChartTitle.Width)/2)
            try:
                chart.Axes(2).TickLabelPosition=-4134
                chart.Axes(2).TickLabels.Font.Size=10
                chart.Axes(1).TickLabels.Font.Size=11
                chart.PlotArea.Top=34
                chart.PlotArea.Height=max(60,obj.Height-72)
            except Exception:
                pass
        if name=='评测总览':
            chart=sheet.ChartObjects(1).Chart
            chart.ChartGroups(1).GapWidth=40
            chart.ChartGroups(1).Overlap=0
            chart.Axes(2).MaximumScale=110
            chart.PlotArea.InsideTop=36
            chart.PlotArea.InsideHeight=chart.Parent.Height-58
            for i,pos in [(1,2),(2,2)]:
                series=chart.SeriesCollection(i)
                series.DataLabels().Position=pos
                series.DataLabels().AutoScaleFont=False
                series.DataLabels().Font.Size=9
                series.DataLabels().Font.Bold=True
        else:
            for index in (1,2):
                chart=sheet.ChartObjects(index).Chart
                chart.Axes(2).MaximumScale=1.2
                for series in chart.SeriesCollection():
                    series.DataLabels().Position=2
                    series.DataLabels().Font.Size=11
                    series.DataLabels().Font.Bold=True
        for index in range(1,5 if name=='评测总览' else 3):
            obj=sheet.ChartObjects(index)
            chart=obj.Chart
            chart.Axes(2).TickLabelPosition=-4127
            chart.PlotArea.InsideTop=35
            chart.PlotArea.InsideHeight=max(70,obj.Height-65)
            chart.ChartTitle.IncludeInLayout=False
            chart.ChartTitle.Top=3
            for series in chart.SeriesCollection():
                if series.HasDataLabels:
                    series.DataLabels().Position=2
                    series.DataLabels().Font.Bold=True
        if name=='评测总览':
            sheet.Range('B18:N18').Font.Size=11
            chart=sheet.ChartObjects(1).Chart
            chart.Legend.Position=-4160
            chart.Legend.Font.Size=9
            chart.Legend.Width=55
            chart.Legend.Height=15
            chart.Legend.Left=chart.Parent.Width-60
            chart.Legend.Top=5
        sheet.Activate()
        wb.Application.ActiveWindow.Zoom=zoom
        wb.Application.ActiveWindow.ScrollColumn=1
        wb.Application.ActiveWindow.ScrollRow=1
        sheet.Range('A1').Select()
    wb.Worksheets('评测总览').Activate()


def apply_final_visual_system(wb):
    """两页统一字体层级；单AI页收紧纵向布局并保持分块间隔。"""
    headers={
        '评测总览':['B1:AD1','B8:N8','B19:N20','B29:N30','P28:AD28','B18:D18','P29:R36'],
        '单AI分析':['B1:V1','R8:V8','R14:V14','B20:P21','R20:V20','B26:D26','B29:H29','J29:P29','R29:V29']}
    for name in ('评测总览','单AI分析'):
        s=wb.Worksheets(name)
        r=s.Range('B1:AD36' if name=='评测总览' else 'B1:V41')
        r.Font.Name='微软雅黑';r.Font.Size=13;r.Font.Bold=False
        for a in headers[name]:s.Range(a).Font.Bold=True
        s.Range('B1:P1').Font.Size=18
        s.Range('B2:V3').Font.Size=11
        s.Range('B4:V4').Font.Size=12
        s.Range('B5:N6' if name=='评测总览' else 'B5:V6').Font.Size=20
        s.Range('B5:N6' if name=='评测总览' else 'B5:V6').Font.Bold=True
        if name=='评测总览':
            s.Range('C9:D16').Font.Bold=True
            s.Range('B36:N36').Font.Bold=True
            s.Range('B18:N18').Font.Size=11
            s.Range('P29:AD36').Font.Size=12
            s.Range('B29:N36').Font.Size=12
        else:
            s.Columns('B:V').ColumnWidth=11.5
            s.Rows('8:18').RowHeight=16
            s.Rows('21:24').RowHeight=22
            s.Rows(26).RowHeight=30
            s.Range('B27:V27').UnMerge();s.Range('B27:V27').Clear()
            s.Rows('27:28').RowHeight=8
            s.Rows('29:33').RowHeight=22
            s.Range('B18:P18').Font.Size=11
            s.Range('B26:V26').Font.Size=11
            s.Range('R21:V24').Font.Size=12
            s.Range('R30:V33').Font.Size=12
            for i,a in [(1,'B8:H17'),(2,'J8:P17'),(3,'B35:H40'),(4,'J35:P40')]:
                c=s.ChartObjects(i);t=s.Range(a)
                c.Width=t.Width;c.Height=t.Height;c.Left=t.Left;c.Top=t.Top
            s.PageSetup.PrintArea='$A$1:$W$41'
        for i in range(1,5):
            c=s.ChartObjects(i).Chart
            c.ChartArea.Font.Bold=False
            c.ChartArea.Font.Size=11
            if c.HasTitle:
                c.ChartTitle.Font.Bold=True;c.ChartTitle.Font.Size=12
                c.ChartTitle.Top=3
            for ax in (1,2):
                try:c.Axes(ax).TickLabels.Font.Bold=False;c.Axes(ax).TickLabels.Font.Size=10
                except Exception:pass
            for series in c.SeriesCollection():
                if series.HasDataLabels:
                    series.DataLabels().Font.Bold=False
                    series.DataLabels().Font.Size=10 if (name!='评测总览' or i!=1) else 9
        if name=='单AI分析':
            for i in (3,4):
                c=s.ChartObjects(i).Chart
                c.PlotArea.InsideTop=23
                c.PlotArea.InsideHeight=25
                c.Legend.Position=-4107
                c.Legend.Top=c.Parent.Height-18
                c.Legend.Height=15
                c.Legend.Font.Size=10
        s.Activate()
        wb.Application.ActiveWindow.Zoom=70 if name=='评测总览' else 65
        wb.Application.ActiveWindow.ScrollRow=1
        wb.Application.ActiveWindow.ScrollColumn=1
        s.Range('A1').Select()
    wb.Worksheets('评测总览').Activate()


def bind_live_structure_labels(wb):
    s=wb.Worksheets('单AI分析')
    for chart_index,source in [(3,'E'),(4,'M')]:
        chart=s.ChartObjects(chart_index).Chart
        for i in range(1,5):
            row=29+i
            series=chart.SeriesCollection(i)
            series.Values=f"='单AI分析'!${source}${row}"
            series.HasDataLabels=False
            series.ApplyDataLabels()
            labels=series.DataLabels()
            labels.ShowValue=True
            labels.ShowSeriesName=False
            labels.ShowCategoryName=False
            labels.NumberFormatLinked=False
            labels.NumberFormat='[>=0.08]0%;[<=0]"";""'
            labels.Position=-4108
            labels.Font.Name='微软雅黑'
            labels.Font.Size=10
            labels.Font.Bold=False
    s.Range('BA30:BB33').ClearContents()



def publish():
    if not OUT.exists():
        raise FileNotFoundError("主工作簿不存在")

    core = load_core()
    tables, _, _, _ = core.extract(DB)
    field_rows = core.build_field_dictionary_rows(tables)
    expected_rows = {name: len(rows) for name, rows in tables}
    expected_rows["字段说明"] = len(field_rows)

    before = excel_pids()
    xl = win32.DispatchEx("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    wb = xl.Workbooks.Open(str(OUT))
    qa = {}
    success = False

    try:
        if wb.ReadOnly:
            raise RuntimeError("主工作簿已在 Excel 中打开；请先关闭后再运行重建脚本")
        refresh_data_com(wb, tables, field_rows)
        apply_report_updates(wb)
        apply_research_table(wb)
        apply_comparison_tables(wb)
        finish_note_tables(wb)
        xl.CalculateFullRebuild()
        align_rectangle(wb)
        finish_screen_layout(wb)
        apply_final_visual_system(wb)
        bind_live_structure_labels(wb)
        import importlib.util
        detail_path = ROOT / '_实验系统' / '报告素材' / '补充评分细项.py'
        detail_spec = importlib.util.spec_from_file_location('report_score_details', detail_path)
        detail_module = importlib.util.module_from_spec(detail_spec)
        detail_spec.loader.exec_module(detail_module)
        detail_module.apply_score_details(wb)
        condition_spec = importlib.util.spec_from_file_location('report_conditions', Path(__file__).parent / '_实验系统' / '报告素材' / '补充运行条件.py')
        condition_module = importlib.util.module_from_spec(condition_spec)
        condition_spec.loader.exec_module(condition_module)
        condition_module.apply_conditions(wb)
        data_style_spec = importlib.util.spec_from_file_location('native_data_style', ROOT / '_实验系统' / '报告素材' / '统一数据表格式.py')
        data_style_module = importlib.util.module_from_spec(data_style_spec)
        data_style_spec.loader.exec_module(data_style_module)
        data_style_module.apply_data_tables(wb)

        qa["sheet_names"] = [
            wb.Worksheets(i).Name
            for i in range(1, wb.Worksheets.Count + 1)
        ]
        expected_sheets = REPORT_SHEETS + ["分维度比较", "评分细项"] + DATA_SHEETS
        if qa["sheet_names"] != expected_sheets:
            raise AssertionError(qa["sheet_names"])

        table_names = {
            "运行记录": "tblRuns",
            "评分与用量": "tblMetrics",
            "API计价": "tblPricing",
            "行为记录": "tblActions",
            "核验记录": "tblChecks",
            "字段说明": "tblFields",
        }
        qa["table_rows"] = {}
        for sheet, table in table_names.items():
            count = wb.Worksheets(sheet).ListObjects.Item(table).ListRows.Count
            qa["table_rows"][sheet] = count
            if count != expected_rows[sheet]:
                raise AssertionError((sheet, count, expected_rows[sheet]))

        qa["charts"] = {
            "评测总览": wb.Worksheets("评测总览").ChartObjects().Count,
            "单AI分析": wb.Worksheets("单AI分析").ChartObjects().Count,
        }
        if qa["charts"] != {"评测总览": 4, "单AI分析": 4}:
            raise AssertionError(qa["charts"])

        qa["formula_errors"] = []
        for sheet in qa["sheet_names"]:
            ws = wb.Worksheets(sheet)
            try:
                errors = ws.UsedRange.SpecialCells(-4123, 16)
                for area in errors.Areas:
                    qa["formula_errors"].append({
                        "sheet": sheet,
                        "address": area.Address,
                    })
            except Exception:
                pass
        if qa["formula_errors"]:
            raise AssertionError(qa["formula_errors"])

        qa["external_connections"] = wb.Connections.Count
        if qa["external_connections"] != 0:
            raise AssertionError(qa["external_connections"])

        qa["selectors"] = {
            "评测总览": wb.Worksheets("评测总览").Range("C3").Value2,
            "单AI分析": [
                wb.Worksheets("单AI分析").Range("C3").Value2,
                wb.Worksheets("单AI分析").Range("G3").Value2,
            ],
        }
        qa["hints"] = {
            "评测总览": wb.Worksheets("评测总览").Range("E3").Value2,
            "单AI分析_AI": wb.Worksheets("单AI分析").Range("E3").Value2,
            "单AI分析_轮次": wb.Worksheets("单AI分析").Range("I3").Value2,
        }

        qa["report_gridlines"] = {}
        for sheet in REPORT_SHEETS:
            wb.Worksheets(sheet).Activate()
            qa["report_gridlines"][sheet] = bool(xl.ActiveWindow.DisplayGridlines)
        if not all(qa["report_gridlines"].values()):
            raise AssertionError(qa["report_gridlines"])

        overview = wb.Worksheets("评测总览")
        qa["overview_round_labels"] = [
            overview.Range(f"B{row}").Text for row in range(21, 27)
        ]
        expected_labels = ["GPT", "Grok", "DeepSeek", "GLM", "Qwen", "Gemini"]
        if qa["overview_round_labels"] != expected_labels:
            raise AssertionError(qa["overview_round_labels"])

        qa["overview_round_label_merges"] = [
            overview.Range(f"B{row}").MergeArea.Address for row in range(21, 27)
        ]
        expected_merges = [
            "$B$" + str(row) + ":$C$" + str(row) for row in range(21, 27)
        ]
        if qa["overview_round_label_merges"] != expected_merges:
            raise AssertionError(qa["overview_round_label_merges"])

        qa["print_areas"] = {
            "评测总览": overview.PageSetup.PrintArea,
            "单AI分析": wb.Worksheets("单AI分析").PageSetup.PrintArea,
        }
        expected_print_areas = {
            "评测总览": "$A$1:$AD$36",
            "单AI分析": "$A$1:$W$41",
        }
        if qa["print_areas"] != expected_print_areas:
            raise AssertionError(qa["print_areas"])

        import sys
        sys.path.insert(0,str(ROOT / '_实验系统/报告素材'))
        from 应用最终裁定 import apply as apply_final_adjudication
        apply_final_adjudication(wb)
        from 同步原生数据 import apply as apply_native_facts
        apply_native_facts(wb)
        qa['print_areas']['评测总览']=wb.Worksheets('评测总览').PageSetup.PrintArea
        success = True
    finally:
        try:
            if wb is not None:
                wb.Close(SaveChanges=success)
        except Exception:
            pass
        try:
            if xl is not None:
                xl.Quit()
        except Exception:
            pass
        wb = None
        xl = None
        gc.collect()
        cleanup_excel(before)

    qa["workbook"] = str(OUT)
    qa["sha256"] = hashlib.sha256(OUT.read_bytes()).hexdigest()
    qa["bytes"] = OUT.stat().st_size
    QA.write_text(
        json.dumps(qa, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    print(json.dumps(qa, ensure_ascii=False))

if __name__ == "__main__":
    raise SystemExit("此入口属于旧Q/R整表生成流程，已停用以保护现行S评分和手工排版。请按资料/复现说明.md维护交付成果/项目2_AI评测分析.xlsx。")
