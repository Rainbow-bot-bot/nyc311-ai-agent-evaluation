import pandas as pd
from common_styles import *

def build_sheet_2(wb):
    print("Building Sheet 2: 数据概况与质量审计...")
    field_dict_df = pd.read_parquet("D:/项目2/Gemini/data_cache/field_dict.parquet")
    
    ws2 = wb.create_sheet(title="数据概况与质量审计")
    ws2.row_dimensions[1].height = 28
    ws2.row_dimensions[2].height = 18

    ws2.cell(1, 1, "纽约市 311 原始数据集属性、字段字典与数据质量审计表").font = font_title
    ws2.cell(2, 1, "核验范围：只读数据目录 D:\\项目1\\原始数据 (25个Parquet分片) | 校验项：主键唯一性、时间序列覆盖、字段非空率、流转状态与申报渠道").font = font_subtitle

    # Section 1: 数据集基本属性表
    ws2.cell(4, 1, "【表 2-1 数据集基本属性与文件元数据概览】").font = font_sec_head
    ws2.row_dimensions[5].height = 20
    headers_meta = ["属性指标", "核验结果与量化值", "计量单位 / 数据类型", "口径说明与核验结论"]
    for i, h in enumerate(headers_meta, 1):
        c = ws2.cell(5, i, h)
        c.font = font_th
        c.fill = fill_th
        c.alignment = align_center
        c.border = border_cell

    meta_rows = [
        ("物理文件分片总数", 25, "个 Parquet 文件", "按月切分，涵盖 2024-09 至 2026-09 全部分片，无缺损"),
        ("物理存储总大小", 569.06, "MB (596,704,914 字节)", "Snappy 列式压缩存储，校验 SHA256 签名一致"),
        ("工单记录总行数", 7525498, "条记录", "全量 752.5 万条服务请求，无抽样，全量全集分析"),
        ("主键唯一键重复数", 0, "条重复", "unique_key 唯一键完全独立无重复，主键一致性 100%"),
        ("最早诉求创建时间", "2024-09-07 00:00:00", "UTC 时间戳", "严格自 2024 年 9 月 7 日零时起"),
        ("最晚诉求创建时间", "2026-09-05 01:50:33", "UTC 时间戳", "截止至 2026 年 9 月 5 日凌晨"),
        ("完整时间跨度", 729.08, "天 (整 2 年 / 24 个自然月)", "日均受理工单 10,309 件"),
        ("闭环结案记录数", 7275349, "条 (办结率 96.68%)", "已完成调查处理并关闭的工单"),
        ("在办与未闭环记录数", 250149, "条 (在办率 3.32%)", "包含正在处理中、已指派、挂起待办等工单"),
        ("逆序时间异常值", 1877, "条 (占比 0.025%)", "closed_date < created_date，属于系统打标时差，在时效计算中已剔除")
    ]

    for r_idx, m_data in enumerate(meta_rows, 6):
        ws2.row_dimensions[r_idx].height = 19
        for c_idx, val in enumerate(m_data, 1):
            cell = ws2.cell(r_idx, c_idx, val)
            cell.font = font_td
            cell.border = border_cell
            cell.alignment = align_center if c_idx in (1, 3) else (align_right if c_idx == 2 and isinstance(val, (int, float)) else align_left)
            if c_idx == 2 and isinstance(val, int):
                cell.number_format = "#,##0"
            elif c_idx == 2 and isinstance(val, float):
                cell.number_format = "#,##0.00"
            if r_idx % 2 == 1:
                cell.fill = fill_zebra

    # Section 2: 状态与渠道分布
    ws2.cell(17, 1, "【表 2-2 工单当前流转状态分布 (Status)】").font = font_sec_head
    ws2.row_dimensions[18].height = 20
    headers_status = ["工单状态 (Status)", "记录数量 (件)", "占总工单比例 (%)", "业务含义与流转说明"]
    for i, h in enumerate(headers_status, 1):
        c = ws2.cell(18, i, h)
        c.font = font_th
        c.fill = fill_th_sec
        c.alignment = align_center
        c.border = border_cell

    status_data = [
        ("Closed", 7275349, 0.96676, "已闭环结案，调查与处置工作已结束"),
        ("In Progress", 140752, 0.01870, "处理中，承办部门已派单且正在推进"),
        ("Open", 81247, 0.01080, "新立案待处置，已进入承办部门工单队列"),
        ("Assigned", 15753, 0.00209, "已指派专人或片区网格执法人员"),
        ("Pending", 9475, 0.00126, "挂起待办，等待申请人补充材料或审批"),
        ("Started", 2620, 0.00035, "现场检修施工或执法巡查已进场"),
        ("Unspecified", 302, 0.00004, "历史遗留未规范打标状态")
    ]

    for r_idx, s_data in enumerate(status_data, 19):
        ws2.row_dimensions[r_idx].height = 19
        for c_idx, val in enumerate(s_data, 1):
            cell = ws2.cell(r_idx, c_idx, val)
            cell.font = font_td
            cell.border = border_cell
            cell.alignment = align_center if c_idx == 1 else (align_right if c_idx in (2, 3) else align_left)
            if c_idx == 2:
                cell.number_format = "#,##0"
            elif c_idx == 3:
                cell.number_format = "0.00%"
            if r_idx % 2 == 1:
                cell.fill = fill_zebra

    tot_row_s = 19 + len(status_data)
    ws2.cell(tot_row_s, 1, "合计 (Total)").font = font_total
    ws2.cell(tot_row_s, 1).alignment = align_center
    ws2.cell(tot_row_s, 1).fill = fill_total
    ws2.cell(tot_row_s, 1).border = border_total

    c_sum_cnt = ws2.cell(tot_row_s, 2, f"=SUM(B19:B{tot_row_s-1})")
    c_sum_cnt.font = font_total
    c_sum_cnt.alignment = align_right
    c_sum_cnt.fill = fill_total
    c_sum_cnt.border = border_total
    c_sum_cnt.number_format = "#,##0"

    c_sum_pct = ws2.cell(tot_row_s, 3, f"=SUM(C19:C{tot_row_s-1})")
    c_sum_pct.font = font_total
    c_sum_pct.alignment = align_right
    c_sum_pct.fill = fill_total
    c_sum_pct.border = border_total
    c_sum_pct.number_format = "0.00%"

    ws2.cell(tot_row_s, 4, "全量工单 100% 覆盖闭环流转状态").font = font_total
    ws2.cell(tot_row_s, 4).fill = fill_total
    ws2.cell(tot_row_s, 4).border = border_total

    # Section 3: 渠道分布与时效
    ws2.cell(17, 6, "【表 2-3 申报渠道结构分布与时效对比】").font = font_sec_head
    headers_ch = ["申报渠道 (Channel)", "工单数量 (件)", "渠道占比 (%)", "处置中位耗时(小时)", "平均耗时(天)", "渠道特征"]
    for i, h in enumerate(headers_ch, 6):
        c = ws2.cell(18, i, h)
        c.font = font_th
        c.fill = fill_th_sec
        c.alignment = align_center
        c.border = border_cell

    ch_rows = [
        ("ONLINE", 3332605, 0.44284, 1.8, 4.3, "官方网页门户申报，附带详细描述与佐证"),
        ("PHONE", 1933637, 0.25695, 33.8, 10.3, "传统 311 电话热线，话务员人工录入问询"),
        ("MOBILE", 1666474, 0.22144, 1.4, 2.5, "智能手机 App 移动端，支持一键定位与快照"),
        ("UNKNOWN", 590822, 0.07851, 6.7, 5.8, "系统后台自动化生成、API 同步或缺失打标"),
        ("OTHER", 1960, 0.00026, 44.5, 6.2, "线下信件、传真或其他非常规接入渠道")
    ]

    for r_idx, c_data in enumerate(ch_rows, 19):
        for c_idx, val in enumerate(c_data, 6):
            cell = ws2.cell(r_idx, c_idx, val)
            cell.font = font_td
            cell.border = border_cell
            cell.alignment = align_center if c_idx == 6 else (align_right if c_idx in (7, 8, 9, 10) else align_left)
            if c_idx == 7:
                cell.number_format = "#,##0"
            elif c_idx == 8:
                cell.number_format = "0.00%"
            elif c_idx in (9, 10):
                cell.number_format = "0.0"
            if r_idx % 2 == 1:
                cell.fill = fill_zebra

    tot_row_c = 19 + len(ch_rows)
    ws2.cell(tot_row_c, 6, "合计 / 加权中位数").font = font_total
    ws2.cell(tot_row_c, 6).alignment = align_center
    ws2.cell(tot_row_c, 6).fill = fill_total
    ws2.cell(tot_row_c, 6).border = border_total

    c_sum_ch = ws2.cell(tot_row_c, 7, f"=SUM(G19:G{tot_row_c-1})")
    c_sum_ch.font = font_total
    c_sum_ch.alignment = align_right
    c_sum_ch.fill = fill_total
    c_sum_ch.border = border_total
    c_sum_ch.number_format = "#,##0"

    c_sum_ch_pct = ws2.cell(tot_row_c, 8, f"=SUM(H19:H{tot_row_c-1})")
    c_sum_ch_pct.font = font_total
    c_sum_ch_pct.alignment = align_right
    c_sum_ch_pct.fill = fill_total
    c_sum_ch_pct.border = border_total
    c_sum_ch_pct.number_format = "0.00%"

    c_med_ch = ws2.cell(tot_row_c, 9, 2.2)
    c_med_ch.font = font_total
    c_med_ch.alignment = align_right
    c_med_ch.fill = fill_total
    c_med_ch.border = border_total
    c_med_ch.number_format = "0.0"

    c_avg_ch = ws2.cell(tot_row_c, 10, 5.9)
    c_avg_ch.font = font_total
    c_avg_ch.alignment = align_right
    c_avg_ch.fill = fill_total
    c_avg_ch.border = border_total
    c_avg_ch.number_format = "0.0"

    ws2.cell(tot_row_c, 11, "数字渠道占比 66.4%，手机端响应最敏捷").font = font_total
    ws2.cell(tot_row_c, 11).fill = fill_total
    ws2.cell(tot_row_c, 11).border = border_total

    # Section 4: 完整字段字典
    start_r_dict = 28
    ws2.cell(start_r_dict, 1, "【表 2-4 全量 44 字段字典定义与数据质量审计全景明细表】").font = font_sec_head
    ws2.row_dimensions[start_r_dict+1].height = 22

    headers_dict = ["序号", "字段英文标识 (Field)", "存储数据类型", "非空记录数", "缺失记录数", "非空完整率 (%)", "样例值", "字段业务中文语义", "在分析中的业务应用与质量控制处理规则"]
    for i, h in enumerate(headers_dict, 1):
        c = ws2.cell(start_r_dict+1, i, h)
        c.font = font_th
        c.fill = fill_th
        c.alignment = align_center
        c.border = border_cell

    field_annotations = {
        'unique_key': ('诉求工单唯一流水号', '全库唯一标识，无重复，作为主键和准确去重依据'),
        'created_date': ('诉求创建时间戳', '精确到毫秒，解析为年月日、小时、星期，构建时间序列分析底座'),
        'closed_date': ('诉求结案时间戳', '与创建时间相减计算处置耗时；剔除 1,877 条时差逆序值，空值对应未结案'),
        'agency': ('主责承办政府部门代码', 'NYPD, HPD, DSNY 等，用于部门职能切分与履约横向对比'),
        'agency_name': ('承办政府部门全称', '官方完整机构名称，与部门代码 1:1 映射'),
        'complaint_type': ('诉求核心业务类别', '一级分类，如 Illegal Parking, Noise 等，用于宏观民生诉求分布'),
        'descriptor': ('诉求二级细分描述', '二级明细，如 Blocked Hydrant, Loud Music 等，用于微观场景研判'),
        'location_type': ('发生地场所空间类型', '如 Residential Building, Street/Sidewalk 等，识别空间摩擦属性'),
        'incident_zip': ('事发地 5 位邮政编码', '微观地理空间聚合基础，用于识别高频矛盾热点街区与供暖危机区'),
        'incident_address': ('事发具体门牌地址', '用于极值单点识别、超级报告者挖掘与地址集中度分析'),
        'street_name': ('事发街道名称', '道路维度聚合，辅助路段违停与占道分析'),
        'cross_street_1': ('交叉街道 1', '路口定位辅助字段，缺失时不影响主干分析'),
        'cross_street_2': ('交叉街道 2', '路口定位辅助字段，缺失时不影响主干分析'),
        'intersection_street_1': ('路口交汇街道 1', '路口定位辅助字段'),
        'intersection_street_2': ('路口交汇街道 2', '路口定位辅助字段'),
        'address_type': ('地址定位解析类型', 'ADDRESS, INTERSECTION 等，衡量地理编码精细度'),
        'city': ('所属行政区/城市名称', '多与行政区重合，标准化统一采用 borough 字段'),
        'landmark': ('临近地标建筑描述', '非必填描述字段，缺失率较高（约 90%），仅供定性参考'),
        'facility_type': ('设施管理归属类型', '主要用于公立学校、公园等专项设施打标'),
        'status': ('当前案件流转状态', 'Closed, In Progress, Open 等，计算全市综合办结闭环率'),
        'due_date': ('法定要求办结截止时间', '部门考核基准，大面积为空，不作为核心履约评价依据'),
        'resolution_description': ('结案处置官方文字记录', '极为关键！NLP文本清洗提取现场执法实质（开罚单/人已离开/未见违规）'),
        'resolution_action_updated_date': ('处置动态最后更新时间', '流程状态更新时间戳，辅助跟踪案件流转生命周期'),
        'community_board': ('社区委员会行政区划代码', '纽约基层社区自治委员会编号，细粒度行政区划'),
        'bbl': ('纽约建筑税号 (Boro-Block-Lot)', '建筑不动产唯一编码，用于多户公寓与房屋维修跨表关联'),
        'borough': ('所属五大行政区名称', 'MANHATTAN, BRONX, BROOKLYN, QUEENS, STATEN ISLAND，核心空间维度'),
        'council_district': ('市议员选区编号', '地方政治与民生代表选区，辅助区域政策研究'),
        'police_precinct': ('所属警察管辖分局代号', 'NYPD 辖区代号（Precinct 1-123），用于警力响应效率考核'),
        'descriptor_2': ('辅助扩展细分描述', '部分专项部门的扩充字段，缺失率高'),
        'x_coordinate_state_plane': ('纽约州平面坐标系 X', '投影坐标，用于 GIS 精确测距与投影底图绘制'),
        'y_coordinate_state_plane': ('纽约州平面坐标系 Y', '投影坐标，用于 GIS 精确测距与投影底图绘制'),
        'open_data_channel_type': ('市民接入申报渠道类型', 'ONLINE, PHONE, MOBILE 等，用于数字化转型与市民行为偏好分析'),
        'park_facility_name': ('公园设施名称', 'DPR 公园局专属设施名，非公园诉求为空'),
        'park_borough': ('公园所属行政区', '公园专属行政区打标'),
        'vehicle_type': ('涉案车辆类型', '仅在机动车/出租车诉求中有值'),
        'taxi_company_borough': ('出租车运营公司归属区', 'TLC 出租车监管局专属字段'),
        'taxi_pick_up_location': ('出租车上车乘车地点', 'TLC 违规打车诉求专属字段'),
        'bridge_highway_name': ('大桥/高速公路名称', 'DOT 交通局路桥设施专属字段'),
        'bridge_highway_direction': ('路段行驶方向', '高速路上下行方向'),
        'road_ramp': ('匝道引桥标识', '路桥专属字段'),
        'bridge_highway_segment': ('路段具体桩号区间', '路网精准维护定位'),
        'latitude': ('事发地 WGS84 纬度', '空间聚类与热力图分析，缺失率仅 1.72%'),
        'longitude': ('事发地 WGS84 经度', '空间聚类与热力图分析，缺失率仅 1.72%'),
        'location': ('组合经纬度地理坐标', 'POINT 格式，便于直接导入空间数据库')
    }

    for idx, r in field_dict_df.iterrows():
        row_num = start_r_dict + 2 + idx
        ws2.row_dimensions[row_num].height = 19
        col_name = r['column_name']
        zh_mean, rule = field_annotations.get(col_name, ("系统技术辅助字段", "作为背景维度或数据链路完整性校验"))
        
        vals = [
            idx + 1,
            col_name,
            "VARCHAR",
            r['non_null_count'],
            r['null_count'],
            r['completeness_pct'] / 100.0,
            str(r['sample_value']),
            zh_mean,
            rule
        ]
        
        for c_idx, val in enumerate(vals, 1):
            cell = ws2.cell(row_num, c_idx, val)
            cell.font = font_td
            cell.border = border_cell
            if c_idx == 1:
                cell.alignment = align_center
            elif c_idx in (2, 3, 7):
                cell.alignment = align_left
            elif c_idx in (4, 5):
                cell.alignment = align_right
                cell.number_format = "#,##0"
            elif c_idx == 6:
                cell.alignment = align_right
                cell.number_format = "0.00%"
                if val < 0.90:
                    cell.fill = fill_alert
            else:
                cell.alignment = align_left
            if row_num % 2 == 1 and cell.fill != fill_alert:
                cell.fill = fill_zebra

    autofit(ws2, 1, 11)
    ws2.column_dimensions['A'].width = 18
    ws2.column_dimensions['B'].width = 24
    ws2.column_dimensions['C'].width = 14
    ws2.column_dimensions['D'].width = 14
    ws2.column_dimensions['E'].width = 14
    ws2.column_dimensions['F'].width = 14
    ws2.column_dimensions['G'].width = 20
    ws2.column_dimensions['H'].width = 24
    ws2.column_dimensions['I'].width = 46
