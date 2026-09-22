import pandas as pd
from openpyxl.chart import BarChart, Reference
from common_styles import *

def build_sheet_3(wb):
    print("Building Sheet 3: 发现1_超级报告者与地址畸变...")
    top_addrs_df = pd.read_parquet("D:/项目2/Gemini/data_cache/top_addrs.parquet")
    addr_buckets_df = pd.read_parquet("D:/项目2/Gemini/data_cache/addr_buckets.parquet")
    
    ws3 = wb.create_sheet(title="发现1_超级报告者与地址畸变")
    ws3.row_dimensions[1].height = 28
    ws3.row_dimensions[2].height = 18

    ws3.cell(1, 1, "核心发现 1：记录数 ≠ 现实事件 —— 极端“超级报告者”与地址高集中度剖析").font = font_title
    ws3.cell(2, 1, "核心结论：全市仅 0.02% 地址贡献了 10.9% 工单；第一大极值点单地址产生 17.2 万工单，导致该区统计指标严重失真").font = font_subtitle

    # Table 3-1: TOP 20 Outlier Addresses
    ws3.cell(4, 1, "【表 3-1 全市工单投诉频次极值地址 TOP 20 深度画像】").font = font_sec_head
    ws3.row_dimensions[5].height = 24
    headers_addr = ["排名", "事发详细门牌地址 (Address)", "所属行政区", "邮编", "工单总量 (件)", "噪音工单", "违停工单", "供暖工单", "APP移动端%", "网页端%", "电话端%", "极端异常特征诊断与研判"]
    for i, h in enumerate(headers_addr, 1):
        c = ws3.cell(5, i, h)
        c.font = font_th
        c.fill = fill_th
        c.alignment = align_center
        c.border = border_cell

    for idx, r in top_addrs_df.head(20).iterrows():
        row_num = 6 + idx
        ws3.row_dimensions[row_num].height = 20
        
        addr_str = str(r['incident_address'])
        if '230 STREET' in addr_str:
            diag = "【超级报告者暴击点】个人App高频刷单，24小时持续报送派对噪音，日均 348 件"
        elif r['parking_requests'] > r['total_requests'] * 0.8:
            diag = "【路权争夺黑洞】商业装卸货区/私家车道被占高频报案，日均多轮重复举报"
        elif r['noise_requests'] > r['total_requests'] * 0.8:
            diag = "【邻里恶性矛盾】针对特定住宅的重复噪音投诉，存在人为主观执念放大"
        elif 'AIRPORT' in addr_str:
            diag = "【交通枢纽聚合】肯尼迪机场大型公共设施天然集中诉求点"
        else:
            diag = "【多户大型综合体】住户基数大与设施老化引发的诉求叠加"
            
        vals = [
            idx + 1,
            r['incident_address'],
            r['borough'],
            r['incident_zip'],
            r['total_requests'],
            r['noise_requests'],
            r['parking_requests'],
            r['heat_requests'],
            r['mobile_pct'] / 100.0,
            r['online_pct'] / 100.0,
            r['phone_pct'] / 100.0,
            diag
        ]
        
        for c_idx, val in enumerate(vals, 1):
            cell = ws3.cell(row_num, c_idx, val)
            cell.font = font_td
            cell.border = border_cell
            if c_idx in (1, 3, 4):
                cell.alignment = align_center
            elif c_idx in (5, 6, 7, 8):
                cell.alignment = align_right
                cell.number_format = "#,##0"
            elif c_idx in (9, 10, 11):
                cell.alignment = align_right
                cell.number_format = "0.0%"
            else:
                cell.alignment = align_left
            if row_num % 2 == 1:
                cell.fill = fill_zebra
            if idx == 0 and c_idx in (2, 5, 12):
                cell.fill = fill_alert
                cell.font = font_td_bold

    # Table 3-2: Address Concentration Buckets
    start_r_bkt = 28
    ws3.cell(start_r_bkt, 1, "【表 3-2 全市 86.5 万个独立地址投诉频次阶梯集中度分布】").font = font_sec_head
    ws3.row_dimensions[start_r_bkt+1].height = 22
    headers_bkt = ["频次分层区间", "独立地址数 (个)", "地址占比 (%)", "贡献工单总量 (件)", "工单占比 (%)", "平均每地址工单数", "最小工单", "最大工单", "业务特征与治理启示"]
    for i, h in enumerate(headers_bkt, 1):
        c = ws3.cell(start_r_bkt+1, i, h)
        c.font = font_th
        c.fill = fill_th_sec
        c.alignment = align_center
        c.border = border_cell

    bkt_notes = [
        "普通市民一次性偶然报修，典型的健康长尾诉求",
        "偶发生活摩擦或家庭设施小修，处置闭环率最高",
        "低频复发性问题，存在轻微邻里摩擦或建筑局部瑕疵",
        "中频矛盾集中点，多为老旧多户租赁公寓或商住混合街区",
        "高频重点治理对象，已具备微观黑点特征，消耗大量基层警力",
        "极端超级报告者与异常刷单点！仅 151 个地址吞噬全市 11% 资源！"
    ]

    for idx, r in addr_buckets_df.iterrows():
        row_num = start_r_bkt + 2 + idx
        ws3.row_dimensions[row_num].height = 20
        vals = [
            r['bucket'],
            r['address_count'],
            r['address_pct'] / 100.0,
            r['total_complaints'],
            r['complaints_pct'] / 100.0,
            r['avg_complaints_per_address'],
            r['min_in_bucket'],
            r['max_in_bucket'],
            bkt_notes[idx]
        ]
        for c_idx, val in enumerate(vals, 1):
            cell = ws3.cell(row_num, c_idx, val)
            cell.font = font_td
            cell.border = border_cell
            if c_idx == 1:
                cell.alignment = align_left
            elif c_idx in (2, 4, 7, 8):
                cell.alignment = align_right
                cell.number_format = "#,##0"
            elif c_idx in (3, 5):
                cell.alignment = align_right
                cell.number_format = "0.00%"
            elif c_idx == 6:
                cell.alignment = align_right
                cell.number_format = "0.0"
            else:
                cell.alignment = align_left
            if row_num % 2 == 1:
                cell.fill = fill_zebra
            if idx == 5:
                cell.fill = fill_alert
                cell.font = font_td_bold

    # Total row for buckets
    tot_r_bkt = start_r_bkt + 2 + len(addr_buckets_df)
    ws3.cell(tot_r_bkt, 1, "全市汇总 (Total)").font = font_total
    ws3.cell(tot_r_bkt, 1).alignment = align_center
    ws3.cell(tot_r_bkt, 1).fill = fill_total
    ws3.cell(tot_r_bkt, 1).border = border_total

    c_bkt_addr = ws3.cell(tot_r_bkt, 2, f"=SUM(B{start_r_bkt+2}:B{tot_r_bkt-1})")
    c_bkt_addr.font = font_total
    c_bkt_addr.alignment = align_right
    c_bkt_addr.fill = fill_total
    c_bkt_addr.border = border_total
    c_bkt_addr.number_format = "#,##0"

    c_bkt_addr_p = ws3.cell(tot_r_bkt, 3, f"=SUM(C{start_r_bkt+2}:C{tot_r_bkt-1})")
    c_bkt_addr_p.font = font_total
    c_bkt_addr_p.alignment = align_right
    c_bkt_addr_p.fill = fill_total
    c_bkt_addr_p.border = border_total
    c_bkt_addr_p.number_format = "0.00%"

    c_bkt_cnt = ws3.cell(tot_r_bkt, 4, f"=SUM(D{start_r_bkt+2}:D{tot_r_bkt-1})")
    c_bkt_cnt.font = font_total
    c_bkt_cnt.alignment = align_right
    c_bkt_cnt.fill = fill_total
    c_bkt_cnt.border = border_total
    c_bkt_cnt.number_format = "#,##0"

    c_bkt_cnt_p = ws3.cell(tot_r_bkt, 5, f"=SUM(E{start_r_bkt+2}:E{tot_r_bkt-1})")
    c_bkt_cnt_p.font = font_total
    c_bkt_cnt_p.alignment = align_right
    c_bkt_cnt_p.fill = fill_total
    c_bkt_cnt_p.border = border_total
    c_bkt_cnt_p.number_format = "0.00%"

    c_bkt_avg = ws3.cell(tot_r_bkt, 6, 8.7)
    c_bkt_avg.font = font_total
    c_bkt_avg.alignment = align_right
    c_bkt_avg.fill = fill_total
    c_bkt_avg.border = border_total
    c_bkt_avg.number_format = "0.0"

    ws3.cell(tot_r_bkt, 7, 1).border = border_total
    ws3.cell(tot_r_bkt, 7).fill = fill_total
    ws3.cell(tot_r_bkt, 7).alignment = align_right
    ws3.cell(tot_r_bkt, 7).font = font_total

    ws3.cell(tot_r_bkt, 8, 270483).border = border_total
    ws3.cell(tot_r_bkt, 8).fill = fill_total
    ws3.cell(tot_r_bkt, 8).alignment = align_right
    ws3.cell(tot_r_bkt, 8).number_format = "#,##0"
    ws3.cell(tot_r_bkt, 8).font = font_total

    ws3.cell(tot_r_bkt, 9, "严重幂律分布：1% 地址占 37% 工单，0.02% 占 11% 工单").font = font_total
    ws3.cell(tot_r_bkt, 9).fill = fill_total
    ws3.cell(tot_r_bkt, 9).border = border_total

    # Table 3-3: Bronx 10466 Comparison
    start_r_bx = 38
    ws3.cell(start_r_bx, 1, "【表 3-3 畸变案例解剖：布朗克斯 10466 邮区在剔除单一刷单地址前后的对照】").font = font_sec_head
    ws3.row_dimensions[start_r_bx+1].height = 22
    headers_bx = ["统计场景口径", "工单总量 (件)", "噪音工单 (件)", "噪音占比 (%)", "供暖工单 (件)", "违停工单 (件)", "全市邮区排名变化", "统计失真诊断与政策启示"]
    for i, h in enumerate(headers_bx, 1):
        c = ws3.cell(start_r_bx+1, i, h)
        c.font = font_th
        c.fill = fill_th
        c.alignment = align_center
        c.border = border_cell

    bx_data = [
        ("1. 原始全量口径 (含 655 E 230 St)", 247699, 194448, 0.7850, 8160, 15128, "全市第 1 名 (名义投诉之王)", "被单一移动端刷单账号严重污染，虚假造成‘全市治安/噪音最差街区’假象"),
        ("2. 剔除极值单点口径 (修正后)", 74042, 22881, 0.3090, 8109, 14637, "跌出全市前 15 名 (回落至真实水平)", "噪音占比骤降 47.6 个百分点，回归普通混合居住街区正常水平，避免了警力错配")
    ]

    for idx, r in enumerate(bx_data):
        row_num = start_r_bx + 2 + idx
        ws3.row_dimensions[row_num].height = 24
        for c_idx, val in enumerate(r, 1):
            cell = ws3.cell(row_num, c_idx, val)
            cell.font = font_td_bold if idx == 1 else font_td
            cell.border = border_cell
            if c_idx == 1:
                cell.alignment = align_left
                cell.fill = fill_highlight if idx == 1 else fill_alert
            elif c_idx in (2, 3, 5, 6):
                cell.alignment = align_right
                cell.number_format = "#,##0"
            elif c_idx == 4:
                cell.alignment = align_right
                cell.number_format = "0.0%"
                cell.fill = fill_highlight if idx == 1 else fill_alert
            elif c_idx == 7:
                cell.alignment = align_center
            else:
                cell.alignment = align_left

    # Add Chart for Top 10 Addresses
    chart1 = BarChart()
    chart1.type = "col"
    chart1.style = 10
    chart1.title = "全市工单投诉量最高的前 10 个地址 (件)"
    chart1.y_axis.title = "工单总量 (件)"
    chart1.x_axis.title = "事发门牌地址"
    chart1.height = 12
    chart1.width = 18

    data_ref = Reference(ws3, min_col=5, min_row=5, max_row=15)
    cats_ref = Reference(ws3, min_col=2, min_row=6, max_row=15)
    chart1.add_data(data_ref, titles_from_data=True)
    chart1.set_categories(cats_ref)
    chart1.legend = None
    ws3.add_chart(chart1, "E44")

    autofit(ws3, 1, 12)
    ws3.column_dimensions['A'].width = 8
    ws3.column_dimensions['B'].width = 28
    ws3.column_dimensions['C'].width = 14
    ws3.column_dimensions['D'].width = 10
    ws3.column_dimensions['E'].width = 14
    ws3.column_dimensions['F'].width = 13
    ws3.column_dimensions['G'].width = 13
    ws3.column_dimensions['H'].width = 13
    ws3.column_dimensions['I'].width = 12
    ws3.column_dimensions['J'].width = 12
    ws3.column_dimensions['K'].width = 12
    ws3.column_dimensions['L'].width = 46
