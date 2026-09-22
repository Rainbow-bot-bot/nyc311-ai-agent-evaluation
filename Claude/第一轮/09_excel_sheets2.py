"""步骤7（续2）：趋势、同比归因、口径变更工作表。"""

# =========================================================
# 4 月度趋势
# =========================================================
ws = wb.add_worksheet("4 月度趋势")
ws.hide_gridlines(2)
r = warn(ws, 0,
    "本页及以后各页均为分析口径（7,339,707 条），即已剔除「3 数据质量」3.2 的 185,791 条异常聚集记录。")

r = section(ws, r, "  4.1 月度记录量（首月与末月不完整，已在表内标注）")
d = T["t04_monthly"]
r = write_table(ws, d, r, widths=[12, 12, 14, 14, 22], date_cols=("月份",))
full = d[d["月份完整性"] == "完整月"].reset_index(drop=True)
top = r
r = write_table(ws, full[["月份", "记录数"]], r, col=7, widths=[12, 12], date_cols=("月份",))
ws.merge_range(top - 1, 7, top - 1, 8, "（下图数据源：仅 23 个完整月）", F["note"])
nf = len(full)
ch = wb.add_chart({"type": "column"})
ch.add_series({
    "name": "月度记录数",
    "categories": ["4 月度趋势", top + 1, 7, top + nf, 7],
    "values":     ["4 月度趋势", top + 1, 8, top + nf, 8],
    "fill": {"color": "#4472C4"},
})
ch.set_title({"name": "2024-10 ~ 2026-08 完整月记录量（分析口径）"})
ch.set_y_axis({"name": "记录数", "min": 200000})
ch.set_x_axis({"name": "创建月份", "num_format": "yyyy-mm"})
ch.set_legend({"none": True}); ch.set_size({"width": 900, "height": 320})
ws.insert_chart(r, 0, ch)
r += 18
r = note(ws, r,
    "23 个完整月的记录量在 25.1 万 ~ 34.8 万之间波动。2026 年的每一个月都高于 2025 年同月。"
    "月度波动同时受自然月天数、季节性问题（供暖、冰雪、街头噪音）和天气事件影响，"
    "单月比较意义有限，因此下节改用整年窗口。")

r = section(ws, r, "  4.2 同比窗口定义与总量")
r = write_table(ws, T["t05_yoy_window"], r, widths=[30, 14, 10, 14])
r = note(ws, r,
    "为了用满全部数据又避免季节错配，同比窗口按数据首日对齐切分：Y1 = 2024-09-07~2025-09-06（365 天），"
    "Y2 = 2025-09-07~2026-09-05（364 天）。两窗口天数差 1 天，因此增长率以日均记录数计算："
    "9,557.8 → 10,580.0，即 +10.7%；按总量计为 3,488,583 → 3,851,124，净增 362,541 条（+10.4%）。"
    "本看板所有「占净增长%」都以 362,541 为分母。")

# =========================================================
# 5 同比增长归因
# =========================================================
ws = wb.add_worksheet("5 同比增长归因")
ws.hide_gridlines(2)
r = section(ws, 0, "  5.1 净增长 362,541 条按问题类型拆分（按变动绝对值排序，取前 20）")
d = T["t06_yoy_by_type"]
r = write_table(ws, d, r, widths=[30, 14, 14, 12, 12, 14])
nb = len(d)
ch = wb.add_chart({"type": "bar"})
ch.add_series({
    "name": "增减量",
    "categories": ["5 同比增长归因", r - nb - 2, 0, r - 3, 0],
    "values":     ["5 同比增长归因", r - nb - 2, 3, r - 3, 3],
    "fill": {"color": "#4472C4"},
    "invert_if_negative": True,
    "invert_if_negative_color": "#C00000",
})
ch.set_title({"name": "Y2 vs Y1 记录量增减（条，蓝=增 / 红=减）"})
ch.set_x_axis({"name": "增减量（条）"})
ch.set_y_axis({"reverse": True, "label_position": "low"})
ch.set_legend({"none": True}); ch.set_size({"width": 820, "height": 620})
ws.insert_chart(r, 0, ch)
r += 33
r = note(ws, r,
    "增长高度集中：供暖热水（+63,648，+21.7%）、违章停车（+57,569，+10.5%）、冰雪（+55,130，+626%）、"
    "路面状况（+48,928，+70.5%）四类合计占净增长 62.2%，其中三类是天气敏感型（合计 46.3%）。"
    "Water Maintenance 的 +12,034 与「6 口径变更警示」中 Water System 的消失是同一件事，"
    "属于口径迁移而非新增需求，其 Y1 基数只有 9 条，因此不计算同比百分比。"
    "下降项里，Drug Activity（−19,352）与 Lead（−13,118）的降幅主要来自 Y1 内的短期高峰"
    "（Lead 在 2024-10/11 达 5,529/9,103 条，Drug Activity 在 2025-07/08 达 6,038/5,447 条），"
    "之后回落到低位并保持，属于水平位移；本数据不含政策或系统变更记录，不足以判断原因。")

r = section(ws, r, "  5.2 净增长按受理机构拆分")
d = T["t07_yoy_by_agency"]
r = write_table(ws, d, r, widths=[12, 44, 14, 14, 12, 12, 14])
r = note(ws, r,
    "住房局 HPD（+146,949，占净增长 40.5%）、警局 NYPD（+79,174，占 21.8%）与交通局 DOT"
    "（+67,712，占 18.7%）是增长主体，三者合计占净增长 81%。"
    "但增长的性质不同：HPD +19.8%、DOT +32.4% 是明显提速，NYPD 只有 +5.0%，"
    "在其庞大基数（Y1 156.9 万条）上属于温和增长。"
    "EDC 的 −38.8% 全部来自 Noise - Helicopter（该类型月度记录数从约 1,300~2,800 条降到 2026-08 的 229 条）。"
    "表中两个机构 Y1 为 0：治安官办公室（OOS）自 2025-09 起接入，全部记录都是 Cannabis Retailer"
    "（与 6.3 中该类型的首次出现月一致），是新增业务线而非增长；"
    "NYC311-PRD/HIQA 只在 2026-03 与 2026-04 出现 371 条 Sidewalk Condition，量级可忽略。"
    "这两个机构的 Y1 基数为 0，因此不计算同比百分比。")

r = section(ws, r, "  5.3 两个冬季直接对比（12-01 ~ 次年 03-31，各 121 天）")
d = T["t08_winter"]
r = write_table(ws, d, r, widths=[30, 14, 18, 26, 26])
r = note(ws, r,
    "同长度冬季窗口对比：全部记录 1,125,791 → 1,324,555（+17.7%）；冰雪 8,806 → 63,930（7.3 倍）；"
    "供暖热水 197,122 → 232,644（+18.0%）；路面状况 23,357 → 45,574（+95.1%）。"
    "三类天气敏感型问题在同一个冬季同时大幅上升，且冰雪类只在 11 月至次年 3 月出现。"
    "推断层级说明：已验证的事实是这三类记录同时同期大幅上升；"
    "「2025-26 冬季比 2024-25 冬季更寒冷多雪」是与记录形态一致的合理推断，"
    "但本数据不含任何气象观测，无法在此数据内验证，属于未验证解释。")

r = section(ws, r, "  5.4 天气敏感型问题的月度形态")
d = T["t09_weather_monthly"]
r = write_table(ws, d, r, widths=[12, 12, 14, 12, 12], date_cols=("月份",))
nm = len(d)
ch = wb.add_chart({"type": "line"})
for j, (nmz, color) in enumerate([("冰雪", "#2E75B6"), ("供暖热水", "#C00000"),
                                  ("路面状况", "#7F7F7F"), ("树木受损", "#548235")], start=1):
    ch.add_series({
        "name": nmz,
        "categories": ["5 同比增长归因", r - nm - 2, 0, r - 3, 0],
        "values":     ["5 同比增长归因", r - nm - 2, j, r - 3, j],
        "line": {"color": color, "width": 2},
    })
ch.set_title({"name": "四类天气敏感型问题的月度记录量"})
ch.set_y_axis({"name": "记录数"})
ch.set_x_axis({"name": "创建月份", "num_format": "yyyy-mm"})
ch.set_size({"width": 900, "height": 340})
ws.insert_chart(r, 0, ch)
r += 19
r = note(ws, r,
    "季节性特征清晰：冰雪类只在 11 月—次年 3 月出现（其余月份为 0 或个位数）；供暖热水在 1 月见顶、"
    "8-9 月见底；路面状况在 2026-03 出现全期最高的 28,690 条（2025-03 为 7,035 条），"
    "时点紧随冰雪高峰之后，与「冬季结束后路面损坏集中上报」的顺序一致；树木受损在夏季 6-7 月最高。")

# =========================================================
# 6 口径变更警示
# =========================================================
ws = wb.add_worksheet("6 口径变更警示")
ws.hide_gridlines(2)
r = warn(ws, 0,
    "读趋势前必须先看这页：数据中至少有一次问题类型口径切换，若直接按 complaint_type 看趋势会得出错误结论。")

r = section(ws, r, "  6.1 DEP（环保局）用水与污水类型的月度切换")
d = T["t10_taxonomy_switch"]
r = write_table(ws, d, r, widths=[12, 18, 20, 14, 20, 12, 12], date_cols=("月份",))
nm = len(d)
ch = wb.add_chart({"type": "line"})
for j, (nmz, color, dash) in enumerate(
        [("Water System（旧）", "#C00000", "solid"),
         ("Water Maintenance（新）", "#2E75B6", "solid"),
         ("Sewer（旧）", "#ED7D31", "dash"),
         ("Sewer Maintenance（新）", "#548235", "dash"),
         ("四类合计", "#7F7F7F", "solid")], start=1):
    ch.add_series({
        "name": nmz,
        "categories": ["6 口径变更警示", r - nm - 2, 0, r - 3, 0],
        "values":     ["6 口径变更警示", r - nm - 2, j, r - 3, j],
        "line": {"color": color, "width": 2.25, "dash_type": dash},
    })
ch.set_title({"name": "旧类型归零、新类型接棒，四类合计未同步崩塌"})
ch.set_y_axis({"name": "记录数"})
ch.set_x_axis({"name": "创建月份", "num_format": "yyyy-mm"})
ch.set_size({"width": 900, "height": 340})
ws.insert_chart(r, 0, ch)
r += 19

r = section(ws, r, "  6.2 逐日证据：切换发生在 2026-07-29")
d = T["t11_taxonomy_daily"]
r = write_table(ws, d, r, widths=[12, 20, 22, 14, 22], date_cols=("日期",))
r = note(ws, r,
    "2026-07-28 当天 Water System 仍有 139 条、Sewer 仍有 33 条；2026-07-29 起 Water System 降至 1 条、"
    "Sewer 降至 11 条，2026-07-30 之后两者恒为 0，而 Water Maintenance 与 Sewer Maintenance 同步跳升。"
    "这是典型的口径替换而非需求消失：如果按类型看，会得到「用水投诉在 2026 年 8 月归零」的错误结论；"
    "把四类合并后，2026-08 合计 13,985 条，与前几个月同一量级，DEP 机构总量 22,233 条也在正常区间。"
    "本看板的处理方式：涉及趋势的分析都把这四类合并看待，季节性表已剔除这四类以避免误读。")

r = section(ws, r, "  6.3 其他窗口内新增或停用的问题类型（记录数 ≥ 2,000）")
r = write_table(ws, T["t12_type_lifespan"], r,
                widths=[32, 10, 14, 14, 12, 16], date_cols=("首次出现月", "最后出现月"))
r = note(ws, r,
    "Snow or Ice 的「2024-11 首现 / 2026-03 末现」是季节性造成的，不是口径变更。"
    "Cannabis Retailer（2025-09 起）属于窗口内真实新增类型；School Maintenance、Non-Residential Heat、"
    "Bike/Roller/Skate Chronic、Construction Safety Enforcement 在窗口内停止出现，"
    "本数据无法区分是业务下线、改名归并还是自然消失，因此不对这些类型做趋势解读。")
