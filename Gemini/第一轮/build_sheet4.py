import pandas as pd
from openpyxl.chart import LineChart, Reference
from common_styles import *

def build_sheet_4(wb):
    print("Building Sheet 4: 发现2_供暖危机与空间不平等...")
    monthly_df = pd.read_parquet("D:/项目2/Gemini/data_cache/monthly.parquet")
    heat_zips_df = pd.read_parquet("D:/项目2/Gemini/data_cache/heat_zips.parquet")
    
    ws4 = wb.create_sheet(title="发现2_供暖危机与空间不平等")
    ws4.row_dimensions[1].height = 28
    ws4.row_dimensions[2].height = 18

    ws4.cell(1, 1, "核心发现 2：空间不平等与极端民生痛点 —— 冬季供暖危机的空间断层与阶层鸿沟").font = font_title
    ws4.cell(2, 1, "核心结论：供暖热水夏冬极差高达 25 倍；布朗克斯万人投诉率高达 155.1 件，是皇后区 4.4 倍、史泰登岛 11.1 倍，占全城 35% 供暖诉求").font = font_subtitle

    # Table 4-1: Monthly Heat Trend
    ws4.cell(4, 1, "【表 4-1 全市供暖与热水诉求 25 个月时序走势 (2024.09 - 2026.09)】").font = font_sec_head
    ws4.row_dimensions[5].height = 22
    headers_heat_m = ["年月 (Month)", "全市民生工单总数", "供暖与热水工单数", "占当月总工单比例 (%)", "季节属性划分", "季节运行特征与民生体征"]
    for i, h in enumerate(headers_heat_m, 1):
        c = ws4.cell(5, i, h)
        c.font = font_th
        c.fill = fill_th
        c.alignment = align_center
        c.border = border_cell

    season_map = {
        '2024-09': ("非供暖季 (夏季尾声)", "基准低位，月均不足 3 千件"),
        '2024-10': ("供暖季启动期", "法定供暖季自 10 月 1 日开启，报修量破万"),
        '2024-11': ("初冬升温期", "气温明显下降，老旧锅炉故障集中爆发"),
        '2024-12': ("深冬高峰期", "寒潮来袭，报修突破 6 万件"),
        '2025-01': ("严冬极值期 (峰值)", "全年供暖诉求最高峰，单月达 6.7 万件"),
        '2025-02': ("隆冬持续期", "严寒持续，报修维持 4 万件高位"),
        '2025-03': ("初春过渡期", "气温逐步回暖，诉求逐步回落至 2.6 万件"),
        '2025-04': ("供暖季收尾期", "春季气温平稳，报修降至 1.6 万件"),
        '2025-05': ("供暖季结束", "5 月 31 日法定供暖季结束"),
        '2025-06': ("夏季低谷期", "仅剩零星生活热水报修"),
        '2025-07': ("盛夏低谷期", "年内最低点，仅 3,008 件"),
        '2025-08': ("盛夏低谷期", "年内最低点，仅 2,929 件"),
        '2025-09': ("非供暖季 (夏季)", "维持夏季基准低位"),
        '2025-10': ("供暖季启动期", "报修跳升至 3.1 万件"),
        '2025-11': ("初冬升温期", "报修升至 4.7 万件"),
        '2025-12': ("深冬高峰期", "报修突破 6.3 万件"),
        '2026-01': ("严冬极值期 (峰值)", "次年最高峰，达 79,928 件，占总工单 23%！"),
        '2026-02': ("隆冬持续期", "报修高达 5.9 万件"),
        '2026-03': ("初春过渡期", "降至 2.9 万件"),
        '2026-04': ("供暖季收尾期", "降至 1.9 万件"),
        '2026-05': ("供暖季结束", "降至不足 1 万件"),
        '2026-06': ("夏季低谷期", "恢复夏季低谷"),
        '2026-07': ("盛夏低谷期", "恢复夏季低谷"),
        '2026-08': ("盛夏低谷期", "恢复夏季低谷"),
        '2026-09': ("数据截止月", "仅覆盖前 5 天记录")
    }

    for idx, r in monthly_df.iterrows():
        row_num = 6 + idx
        ws4.row_dimensions[row_num].height = 19
        ym = r['year_month']
        s_type, s_desc = season_map.get(ym, ("常规期", "常规运行"))
        pct = r['heat_requests'] / r['total_requests']
        vals = [
            ym,
            r['total_requests'],
            r['heat_requests'],
            pct,
            s_type,
            s_desc
        ]
        for c_idx, val in enumerate(vals, 1):
            cell = ws4.cell(row_num, c_idx, val)
            cell.font = font_td
            cell.border = border_cell
            if c_idx == 1:
                cell.alignment = align_center
            elif c_idx in (2, 3):
                cell.alignment = align_right
                cell.number_format = "#,##0"
            elif c_idx == 4:
                cell.alignment = align_right
                cell.number_format = "0.0%"
                if val > 0.15:
                    cell.fill = fill_alert
                    cell.font = font_td_bold
            elif c_idx == 5:
                cell.alignment = align_center
            else:
                cell.alignment = align_left
            if row_num % 2 == 1 and cell.fill != fill_alert:
                cell.fill = fill_zebra

    # Total row for monthly
    tot_r_hm = 6 + len(monthly_df)
    ws4.cell(tot_r_hm, 1, "全周期汇总 (Total)").font = font_total
    ws4.cell(tot_r_hm, 1).alignment = align_center
    ws4.cell(tot_r_hm, 1).fill = fill_total
    ws4.cell(tot_r_hm, 1).border = border_total

    c_tot_all = ws4.cell(tot_r_hm, 2, f"=SUM(B6:B{tot_r_hm-1})")
    c_tot_all.font = font_total
    c_tot_all.alignment = align_right
    c_tot_all.fill = fill_total
    c_tot_all.border = border_total
    c_tot_all.number_format = "#,##0"

    c_tot_heat = ws4.cell(tot_r_hm, 3, f"=SUM(C6:C{tot_r_hm-1})")
    c_tot_heat.font = font_total
    c_tot_heat.alignment = align_right
    c_tot_heat.fill = fill_total
    c_tot_heat.border = border_total
    c_tot_heat.number_format = "#,##0"

    c_tot_hp = ws4.cell(tot_r_hm, 4, f"=C{tot_r_hm}/B{tot_r_hm}")
    c_tot_hp.font = font_total
    c_tot_hp.alignment = align_right
    c_tot_hp.fill = fill_total
    c_tot_hp.border = border_total
    c_tot_hp.number_format = "0.0%"

    ws4.cell(tot_r_hm, 5, "25 个月跨度").fill = fill_total
    ws4.cell(tot_r_hm, 5).alignment = align_center
    ws4.cell(tot_r_hm, 5).border = border_total
    ws4.cell(tot_r_hm, 5).font = font_total

    ws4.cell(tot_r_hm, 6, "全市民生诉求中供暖热水占 8.7%，冬夏两季呈现 25 倍极差震荡").fill = fill_total
    ws4.cell(tot_r_hm, 6).border = border_total
    ws4.cell(tot_r_hm, 6).font = font_total

    # Table 4-2: Borough Per-Capita Heat Comparison
    start_r_bheat = 34
    ws4.cell(start_r_bheat, 1, "【表 4-2 五大行政区供暖诉求人口标准化对比与空间不平等分析】").font = font_sec_head
    ws4.row_dimensions[start_r_bheat+1].height = 22
    headers_bheat = ["所属行政区 (Borough)", "供暖工单总量 (件)", "占全市供暖比重 (%)", "常住人口 (2020普查)", "万人供暖投诉率 (件/万人)", "年化万人供暖投诉率", "相对史泰登岛倍数", "区域供暖脆弱度评价"]
    for i, h in enumerate(headers_bheat, 1):
        c = ws4.cell(start_r_bheat+1, i, h)
        c.font = font_th
        c.fill = fill_th_sec
        c.alignment = align_center
        c.border = border_cell

    pop_map = {
        'BRONX': (1472654, 228475, "【极高危脆弱区】占全城超 1/3 供暖报修，租户供暖困境极为尖锐"),
        'BROOKLYN': (2736074, 174565, "【中高危集中区】人口基数最大，老旧公寓与新建住宅并存，总量第二"),
        'MANHATTAN': (1694251, 156329, "【高位承压区】战前老楼集中，出租公寓供暖管网老化，每万人诉求次高"),
        'QUEENS': (2405464, 85284, "【中低位平稳区】独立别墅与一户独栋比例高，自主供暖比例大"),
        'STATEN ISLAND': (495747, 6939, "【低危平稳区】低密度独栋住宅为主，极少依赖多户集中供暖系统")
    }

    b_order = ['BRONX', 'BROOKLYN', 'MANHATTAN', 'QUEENS', 'STATEN ISLAND']
    for idx, b_name in enumerate(b_order):
        row_num = start_r_bheat + 2 + idx
        ws4.row_dimensions[row_num].height = 22
        pop, h_cnt, eval_txt = pop_map[b_name]
        pct_city = h_cnt / 651594
        rate_10k = (h_cnt / pop) * 10000
        annual_rate = rate_10k / 2.0
        si_rate = (6939 / 495747) * 10000
        ratio = rate_10k / si_rate
        
        vals = [
            b_name,
            h_cnt,
            pct_city,
            pop,
            rate_10k,
            annual_rate,
            f"{ratio:.1f} 倍",
            eval_txt
        ]
        for c_idx, val in enumerate(vals, 1):
            cell = ws4.cell(row_num, c_idx, val)
            cell.font = font_td
            cell.border = border_cell
            if c_idx == 1:
                cell.alignment = align_center
                if idx == 0:
                    cell.fill = fill_alert
                    cell.font = font_td_bold
            elif c_idx in (2, 4):
                cell.alignment = align_right
                cell.number_format = "#,##0"
            elif c_idx == 3:
                cell.alignment = align_right
                cell.number_format = "0.00%"
            elif c_idx in (5, 6):
                cell.alignment = align_right
                cell.number_format = "#,##0.0"
                if idx == 0:
                    cell.fill = fill_alert
                    cell.font = font_td_bold
            elif c_idx == 7:
                cell.alignment = align_center
                if idx == 0:
                    cell.fill = fill_alert
                    cell.font = font_td_bold
            else:
                cell.alignment = align_left
            if row_num % 2 == 1 and cell.fill != fill_alert:
                cell.fill = fill_zebra

    # Total row for borough heat
    tot_r_bh = start_r_bheat + 2 + len(b_order)
    ws4.cell(tot_r_bh, 1, "全市合计 / 平均").font = font_total
    ws4.cell(tot_r_bh, 1).alignment = align_center
    ws4.cell(tot_r_bh, 1).fill = fill_total
    ws4.cell(tot_r_bh, 1).border = border_total

    c_tot_bh_cnt = ws4.cell(tot_r_bh, 2, f"=SUM(B{start_r_bheat+2}:B{tot_r_bh-1})")
    c_tot_bh_cnt.font = font_total
    c_tot_bh_cnt.alignment = align_right
    c_tot_bh_cnt.fill = fill_total
    c_tot_bh_cnt.border = border_total
    c_tot_bh_cnt.number_format = "#,##0"

    c_tot_bh_pct = ws4.cell(tot_r_bh, 3, f"=SUM(C{start_r_bheat+2}:C{tot_r_bh-1})")
    c_tot_bh_pct.font = font_total
    c_tot_bh_pct.alignment = align_right
    c_tot_bh_pct.fill = fill_total
    c_tot_bh_pct.border = border_total
    c_tot_bh_pct.number_format = "0.00%"

    c_tot_pop = ws4.cell(tot_r_bh, 4, f"=SUM(D{start_r_bheat+2}:D{tot_r_bh-1})")
    c_tot_pop.font = font_total
    c_tot_pop.alignment = align_right
    c_tot_pop.fill = fill_total
    c_tot_pop.border = border_total
    c_tot_pop.number_format = "#,##0"

    c_city_rate = ws4.cell(tot_r_bh, 5, f"=B{tot_r_bh}/D{tot_r_bh}*10000")
    c_city_rate.font = font_total
    c_city_rate.alignment = align_right
    c_city_rate.fill = fill_total
    c_city_rate.border = border_total
    c_city_rate.number_format = "#,##0.0"

    c_city_arate = ws4.cell(tot_r_bh, 6, f"=E{tot_r_bh}/2")
    c_city_arate.font = font_total
    c_city_arate.alignment = align_right
    c_city_arate.fill = fill_total
    c_city_arate.border = border_total
    c_city_arate.number_format = "#,##0.0"

    ws4.cell(tot_r_bh, 7, "基准 1.0 倍").fill = fill_total
    ws4.cell(tot_r_bh, 7).alignment = align_center
    ws4.cell(tot_r_bh, 7).font = font_total
    ws4.cell(tot_r_bh, 7).border = border_total

    ws4.cell(tot_r_bh, 8, "布朗克斯万人投诉率高出全市平均水平一倍以上，不平等性显著").fill = fill_total
    ws4.cell(tot_r_bh, 8).border = border_total
    ws4.cell(tot_r_bh, 8).font = font_total

    # Table 4-3: Top 20 Heat ZIP Codes
    start_r_zheat = 44
    ws4.cell(start_r_zheat, 1, "【表 4-3 全市冬季供暖危机最严重的 TOP 20 核心脆弱邮区画像】").font = font_sec_head
    ws4.row_dimensions[start_r_zheat+1].height = 22
    headers_zheat = ["排名", "邮编 (ZIP)", "所属行政区", "供暖工单量 (件)", "涉及独立建筑数 (栋)", "单建筑平均报修频次", "核心代表街区 / 社区"]
    for i, h in enumerate(headers_zheat, 1):
        c = ws4.cell(start_r_zheat+1, i, h)
        c.font = font_th
        c.fill = fill_th
        c.alignment = align_center
        c.border = border_cell

    zip_neighborhoods = {
        '10467': "Norwood / Williamsbridge (布朗克斯核心租赁集聚区)",
        '10458': "Belmont / Fordham (福特汉姆大学周边高密老房区)",
        '11226': "Flatbush / Ditmas Park (布鲁克林传统老旧多户租赁区)",
        '10468': "Jerome Park / Kingsbridge (高层战前多单元老公寓区)",
        '10452': "Highbridge / Concourse (高桥/大中央走廊密集街区)",
        '10453': "Morris Heights / Mount Hope (高密度多户集合住宅区)",
        '10456': "Morrisania / Melrose (南布朗克斯老旧公共/平价住房区)",
        '10457': "Tremont / Belmont (高失修率租凭公寓聚集带)",
        '10031': "Hamilton Heights / West Harlem (曼哈顿西哈林老旧街区)",
        '10032': "Washington Heights South (曼哈顿华盛顿高地南部)",
        '10040': "Inwood / Fort George (曼哈顿最北部老旧出租公寓带)",
        '10462': "Parkchester / Pelham Parkway (东布朗克斯多户大型社区)",
        '10463': "Kingsbridge / Marble Hill (横跨曼哈顿与布朗克斯)",
        '10460': "West Farms / Charlotte Gardens (中布朗克斯租赁街区)",
        '10033': "Washington Heights North (华盛顿高地北部)",
        '10459': "Longwood / Hunts Point (亨茨点周边低收入社区)",
        '10451': "Concourse / Mott Haven (南布朗克斯门户街区)",
        '11213': "Crown Heights North (皇冠高地多户排屋与租赁公寓)",
        '11225': "Prospect Lefferts Gardens (展望莱弗茨花园公寓带)",
        '10025': "Upper West Side / Manhattan Valley (上西区低租金老旧单元)"
    }

    for idx, r in heat_zips_df.head(20).iterrows():
        row_num = start_r_zheat + 2 + idx
        ws4.row_dimensions[row_num].height = 19
        z_code = str(r['incident_zip'])
        neigh = zip_neighborhoods.get(z_code, f"{r['borough']} 居民区")
        vals = [
            idx + 1,
            z_code,
            r['borough'],
            r['heat_requests'],
            r['unique_addresses_with_heat_issues'],
            r['avg_heat_requests_per_address'],
            neigh
        ]
        for c_idx, val in enumerate(vals, 1):
            cell = ws4.cell(row_num, c_idx, val)
            cell.font = font_td
            cell.border = border_cell
            if c_idx in (1, 2, 3):
                cell.alignment = align_center
            elif c_idx in (4, 5):
                cell.alignment = align_right
                cell.number_format = "#,##0"
            elif c_idx == 6:
                cell.alignment = align_right
                cell.number_format = "0.0"
            else:
                cell.alignment = align_left
            if row_num % 2 == 1:
                cell.fill = fill_zebra
            if idx < 10 and r['borough'] == 'BRONX':
                if c_idx == 3:
                    cell.fill = fill_alert
                    cell.font = font_td_bold

    # Embedded Line Chart for Monthly Heat Trend
    chart_heat = LineChart()
    chart_heat.title = "全市月度供暖与热水诉求总量走势 (2024.09 - 2026.09)"
    chart_heat.style = 13
    chart_heat.y_axis.title = "工单数量 (件)"
    chart_heat.x_axis.title = "年月"
    chart_heat.height = 13
    chart_heat.width = 19

    data_ref_h = Reference(ws4, min_col=3, min_row=5, max_row=5+len(monthly_df))
    cats_ref_h = Reference(ws4, min_col=1, min_row=6, max_row=5+len(monthly_df))
    chart_heat.add_data(data_ref_h, titles_from_data=True)
    chart_heat.set_categories(cats_ref_h)
    chart_heat.legend = None
    ws4.add_chart(chart_heat, "H5")

    autofit(ws4, 1, 8)
    ws4.column_dimensions['A'].width = 14
    ws4.column_dimensions['B'].width = 16
    ws4.column_dimensions['C'].width = 16
    ws4.column_dimensions['D'].width = 18
    ws4.column_dimensions['E'].width = 24
    ws4.column_dimensions['F'].width = 20
    ws4.column_dimensions['G'].width = 20
    ws4.column_dimensions['H'].width = 46
