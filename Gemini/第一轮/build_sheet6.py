import pandas as pd
from openpyxl.chart import BarChart, Reference
from common_styles import *

def build_sheet_6(wb):
    print("Building Sheet 6: 发现4_部门响应与时效画像...")
    agency_df = pd.read_parquet("D:/项目2/Gemini/data_cache/agency.parquet")
    complaint_types_df = pd.read_parquet("D:/项目2/Gemini/data_cache/complaint_types.parquet")
    
    ws6 = wb.create_sheet(title="发现4_部门响应与时效画像")
    ws6.row_dimensions[1].height = 28
    ws6.row_dimensions[2].height = 18

    ws6.cell(1, 1, "核心发现 4：市政部门履约效能画像 —— 警务敏捷响应 vs 住建长期滞后与处置落差").font = font_title
    ws6.cell(2, 1, "核心结论：NYPD 处置中位数仅 1.3 小时（但未见违规率达 45%）；HPD 与 DOB 涉及复杂房屋维修中位数需 3.6-6.6 天，P90 耗时超 37-145 天").font = font_subtitle

    # Table 6-1: Agency Performance Leaderboard
    ws6.cell(4, 1, "【表 6-1 市政主要承办部门履约时效与流转效率总榜】").font = font_sec_head
    ws6.row_dimensions[5].height = 22
    headers_ag = ["承办部门 (Agency)", "机构全称与核心管辖业务", "全周期受理量", "闭环结案量", "综合办结率", "处置中位数 (小时)", "P75 分位数 (小时)", "P90 分位数 (小时)", "平均耗时 (小时)", "折合工作日 (天)", "履约敏捷度画像"]
    for i, h in enumerate(headers_ag, 1):
        c = ws6.cell(5, i, h)
        c.font = font_th
        c.fill = fill_th
        c.alignment = align_center
        c.border = border_cell

    agency_profiles = {
        'NYPD': ("New York City Police Department (纽约市警察局)", "【即时敏捷型】警车街头巡逻就近指派，1.3 小时极速处置噪音违停"),
        'HPD': ("Department of Housing Preservation and Development (房屋维护局)", "【中长周期维修型】供暖优先(41h)，但结构/虫害修缮多需数周"),
        'DSNY': ("Department of Sanitation (纽约市环卫局)", "【定班清扫型】垃圾清运、街头违章抛洒，常规 1.4 天清运完毕"),
        'DOT': ("Department of Transportation (纽约市交通局)", "【两极分化型】红绿灯故障 2.2h 极速修，道路坑洼需数周摊铺"),
        'DEP': ("Department of Environmental Protection (环境保护局)", "【专业水务型】自来水爆管 3.3h 抢修，商业油烟与噪音需多日检测"),
        'DOB': ("Department of Buildings (楼宇管理局)", "【法定长周期型】违法加建、电梯与脚手架安全，涉及法定调查程序"),
        'DPR': ("Department of Parks and Recreation (公园与休闲局)", "【季节滞后型】树木修剪与倒伏清理，非紧急诉求积压严重"),
        'DOHMH': ("Department of Health and Mental Hygiene (卫生与心理健康局)", "【检验专业型】餐厅卫生、鼠患调查，需专业卫生检查员上门"),
        'DHS': ("Department of Homeless Services (无家可归者服务局)", "【社工寻访型】街头流浪救助，外勤社工 7.4 小时上门寻访评估"),
        'TLC': ("Taxi and Limousine Commission (出租车管理委员会)", "【漫长行政审理型】针对网约车/出租车拒载投诉，需举证听证(中位40天)"),
        'DCWP': ("Department of Consumer and Worker Protection (消费者保护局)", "【纠纷调解型】商家欺诈与劳工纠纷，中位 5.2 天调解结案"),
        'EDC': ("New York City Economic Development Corporation (经济开发署)", "【重资产工程型】滨水码头与工业设施维保，耗时极长(中位262天)"),
        'OOS': ("Office of Special Enforcement (特别执法办公室)", "【专项执法型】非法短租与民宿专项整治"),
        'DOE': ("Department of Education (纽约市教育局)", "【校园设施型】公立学校校舍与周边环境诉求")
    }

    for idx, r in agency_df.iterrows():
        row_num = 6 + idx
        ws6.row_dimensions[row_num].height = 20
        ag = r['agency']
        ag_full, ag_eval = agency_profiles.get(ag, ("市政专业部门", "常规履行法定管理职责"))
        med_h = r['median_resolution_hours']
        avg_d = r['avg_resolution_hours'] / 24.0 if pd.notna(r['avg_resolution_hours']) else 0
        
        vals = [
            ag,
            ag_full,
            r['total_requests'],
            r['closed_requests'],
            r['close_rate_pct'] / 100.0,
            med_h,
            r['p75_resolution_hours'],
            r['p90_resolution_hours'],
            r['avg_resolution_hours'],
            avg_d,
            ag_eval
        ]
        for c_idx, val in enumerate(vals, 1):
            cell = ws6.cell(row_num, c_idx, val)
            cell.font = font_td
            cell.border = border_cell
            if c_idx == 1:
                cell.alignment = align_center
                cell.font = font_td_bold
                if ag == 'NYPD':
                    cell.fill = fill_success
                elif ag in ('DOB', 'TLC', 'EDC'):
                    cell.fill = fill_alert
            elif c_idx == 2:
                cell.alignment = align_left
            elif c_idx in (3, 4):
                cell.alignment = align_right
                cell.number_format = "#,##0"
            elif c_idx == 5:
                cell.alignment = align_right
                cell.number_format = "0.0%"
            elif c_idx in (6, 7, 8, 9, 10):
                cell.alignment = align_right
                cell.number_format = "#,##0.0"
                if c_idx == 6 and val <= 2.0:
                    cell.fill = fill_success
                elif c_idx == 6 and val >= 100.0:
                    cell.fill = fill_alert
            else:
                cell.alignment = align_left
            if row_num % 2 == 1 and cell.fill not in (fill_success, fill_alert):
                cell.fill = fill_zebra

    # Total row for Agency
    tot_r_ag = 6 + len(agency_df)
    ws6.cell(tot_r_ag, 1, "主要部门汇总").font = font_total
    ws6.cell(tot_r_ag, 1).alignment = align_center
    ws6.cell(tot_r_ag, 1).fill = fill_total
    ws6.cell(tot_r_ag, 1).border = border_total

    ws6.cell(tot_r_ag, 2, "覆盖全市 99.9% 工单承办部门").fill = fill_total
    ws6.cell(tot_r_ag, 2).border = border_total
    ws6.cell(tot_r_ag, 2).font = font_total

    c_tot_ag_cnt = ws6.cell(tot_r_ag, 3, f"=SUM(C6:C{tot_r_ag-1})")
    c_tot_ag_cnt.font = font_total
    c_tot_ag_cnt.alignment = align_right
    c_tot_ag_cnt.fill = fill_total
    c_tot_ag_cnt.border = border_total
    c_tot_ag_cnt.number_format = "#,##0"

    c_tot_ag_cls = ws6.cell(tot_r_ag, 4, f"=SUM(D6:D{tot_r_ag-1})")
    c_tot_ag_cls.font = font_total
    c_tot_ag_cls.alignment = align_right
    c_tot_ag_cls.fill = fill_total
    c_tot_ag_cls.border = border_total
    c_tot_ag_cls.number_format = "#,##0"

    c_tot_ag_rate = ws6.cell(tot_r_ag, 5, f"=D{tot_r_ag}/C{tot_r_ag}")
    c_tot_ag_rate.font = font_total
    c_tot_ag_rate.alignment = align_right
    c_tot_ag_rate.fill = fill_total
    c_tot_ag_rate.border = border_total
    c_tot_ag_rate.number_format = "0.0%"

    c_tot_ag_med = ws6.cell(tot_r_ag, 6, 2.2)
    c_tot_ag_med.font = font_total
    c_tot_ag_med.alignment = align_right
    c_tot_ag_med.fill = fill_total
    c_tot_ag_med.border = border_total
    c_tot_ag_med.number_format = "0.0"

    for c_idx in range(7, 11):
        cell = ws6.cell(tot_r_ag, c_idx, "-")
        cell.fill = fill_total
        cell.alignment = align_center
        cell.border = border_total
        cell.font = font_total

    ws6.cell(tot_r_ag, 11, "部门间响应速度极差高达百倍以上，呈现鲜明专业分工特征").fill = fill_total
    ws6.cell(tot_r_ag, 11).border = border_total
    ws6.cell(tot_r_ag, 11).font = font_total

    # Table 6-2: Top 20 Complaint Types Resolution Time
    start_r_types = 24
    ws6.cell(start_r_types, 1, "【表 6-2 全市 TOP 20 核心民生诉求处置时效明细表】").font = font_sec_head
    ws6.row_dimensions[start_r_types+1].height = 22
    headers_types = ["排名", "民生诉求类型 (Complaint Type)", "主责承办部门", "结案总量 (件)", "占全市总比重", "中位数耗时 (小时)", "平均耗时 (天)", "业务处置特征与公众感知"]
    for i, h in enumerate(headers_types, 1):
        c = ws6.cell(start_r_types+1, i, h)
        c.font = font_th
        c.fill = fill_th_sec
        c.alignment = align_center
        c.border = border_cell

    type_narratives = {
        'Illegal Parking': "【极速处置】巡警现场开罚单或劝离，1.5 小时结案",
        'Noise - Residential': "【极速处置】民警上门走访告诫，1.3 小时结案（45% 未见违规）",
        'HEAT/HOT WATER': "【优先攻坚】涉及冬季基本生存权，住建局优先于其他报修（中位 1.7 天）",
        'Blocked Driveway': "【快速疏通】车道被堵严重影响居民出行，巡警 1.8 小时内到场",
        'Noise - Street/Sidewalk': "【极速响应】街头流动噪音，巡警 1.0 小时即时巡查",
        'UNSANITARY CONDITION': "【长期拖延】发霉、鼠患、垃圾堆积，需预约入户核查（中位 10.4 天，平均 22.9 天）",
        'Street Condition': "【批量排期】路面破损、坑洼修补，交通局工程队统一排期（中位 1.2 天）",
        'Water System': "【紧急抢修】水管爆裂/水压骤降，DEP 防汛抢修队 3.3 小时火速排查",
        'PLUMBING': "【慢速维修】下水管道堵塞/反水，涉及入户开工与房东交涉（中位 7.9 天）",
        'Abandoned Vehicle': "【核查拖移】废弃无牌车辆，警方现场贴条核实（中位 2.2 小时）",
        'Dirty Condition': "【保洁清运】环卫工人定点清扫清运，中位 24.7 小时（约 1 天）",
        'Noise - Commercial': "【极速处理】酒吧餐厅音乐外溢，巡警到场警告（中位 0.9 小时）",
        'PAINT/PLASTER': "【室内整修】墙皮脱落、铅漆隐患，住建局要求房东整改（中位 7.3 天）",
        'Noise': "【专业降噪】DEP 负责的空调压缩机与建筑工地噪音（中位 2.8 天）",
        'Traffic Signal Condition': "【生命线工程】交通红绿灯故障，DOT 紧急抢修队 2.2 小时修复！",
        'DOOR/WINDOW': "【安防修缮】门窗损坏漏风，房东修缮周期长（中位 7.3 天）",
        'Encampment': "【社会关怀】流浪人员帐篷露营，警社联动救助劝导（中位 2.2 小时）",
        'Noise - Vehicle': "【流动取证难】汽车警报器、改装排气管炸街（中位 1.1 小时）",
        'WATER LEAK': "【漏水排查】楼上漏水渗水纠纷，预约入户中位 6.7 天",
        'Derelict Vehicles': "【报废清理】环卫局收缴报废车辆，中位 41.5 小时"
    }

    for idx, r in complaint_types_df.head(20).iterrows():
        row_num = start_r_types + 2 + idx
        ws6.row_dimensions[row_num].height = 19
        ctype = r['complaint_type']
        narr = type_narratives.get(ctype, "标准工单处置流程")
        vals = [
            idx + 1,
            ctype,
            r['agency'],
            r['total_requests'],
            r['pct_of_total'] / 100.0,
            r['median_resolution_hours'],
            r['avg_resolution_days'],
            narr
        ]
        for c_idx, val in enumerate(vals, 1):
            cell = ws6.cell(row_num, c_idx, val)
            cell.font = font_td
            cell.border = border_cell
            if c_idx in (1, 3):
                cell.alignment = align_center
            elif c_idx == 2:
                cell.alignment = align_left
                cell.font = font_td_bold
            elif c_idx == 4:
                cell.alignment = align_right
                cell.number_format = "#,##0"
            elif c_idx == 5:
                cell.alignment = align_right
                cell.number_format = "0.00%"
            elif c_idx in (6, 7):
                cell.alignment = align_right
                cell.number_format = "#,##0.0"
                if c_idx == 6 and val <= 2.0:
                    cell.fill = fill_success
                elif c_idx == 6 and val >= 100.0:
                    cell.fill = fill_alert
            else:
                cell.alignment = align_left
            if row_num % 2 == 1 and cell.fill not in (fill_success, fill_alert):
                cell.fill = fill_zebra

    # Embedded Bar Chart for Agency Median Hours
    chart_ag = BarChart()
    chart_ag.type = "bar"
    chart_ag.style = 10
    chart_ag.title = "各部门工单处置耗时中位数对比 (小时)"
    chart_ag.x_axis.title = "耗时中位数 (小时)"
    chart_ag.y_axis.title = "承办部门代码"
    chart_ag.height = 12
    chart_ag.width = 18

    data_ref_a = Reference(ws6, min_col=6, min_row=5, max_row=16)
    cats_ref_a = Reference(ws6, min_col=1, min_row=6, max_row=16)
    chart_ag.add_data(data_ref_a, titles_from_data=True)
    chart_ag.set_categories(cats_ref_a)
    chart_ag.legend = None
    ws6.add_chart(chart_ag, "I24")

    autofit(ws6, 1, 8)
    ws6.column_dimensions['A'].width = 10
    ws6.column_dimensions['B'].width = 30
    ws6.column_dimensions['C'].width = 12
    ws6.column_dimensions['D'].width = 15
    ws6.column_dimensions['E'].width = 13
    ws6.column_dimensions['F'].width = 15
    ws6.column_dimensions['G'].width = 14
    ws6.column_dimensions['H'].width = 46
