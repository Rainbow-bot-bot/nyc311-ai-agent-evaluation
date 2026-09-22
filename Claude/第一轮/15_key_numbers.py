"""步骤9：把看板文字里用到的关键数字集中打印出来，用于核对叙述与计算结果一致。"""
import pandas as pd, pathlib
AGG = pathlib.Path(r"D:\项目2\Claude\中间汇总")
def L(n): return pd.read_csv(AGG / f"{n}.csv", encoding="utf-8-sig")

pd.set_option("display.width", 250)
for n in ["t01_overview", "t28_excluded_pairs", "t30_noise_yoy_effect", "t05_yoy_window",
          "t06_yoy_by_type", "t07_yoy_by_agency", "t08_winter", "t19_dup_overall",
          "t20_dup_by_type", "t23_borough", "t16_nypd_outcome", "t21_seasonality",
          "t13_duration_by_agency", "t14_duration_by_type", "t17_nypd_outcome_by_type",
          "t18_hpd_heat_reason", "t25_channel_monthly", "t27_top_types", "t29_excluded_monthly"]:
    print(f"\n########## {n} ##########")
    print(L(n).to_string(index=False))

t5 = L("t05_yoy_window")
n1, n2 = int(t5.loc[0, "记录数"]), int(t5.loc[1, "记录数"])
d1, d2 = float(t5.loc[0, "日均记录数"]), float(t5.loc[1, "日均记录数"])
print(f"\n>>> 净增量 = {n2-n1:,}；总量同比 = {100*(n2-n1)/n1:.1f}%；日均同比 = {100*(d2-d1)/d1:.1f}%")

t6 = L("t06_yoy_by_type")
w3 = t6[t6["问题类型"].isin(["Snow or Ice", "HEAT/HOT WATER", "Street Condition"])]
print(f">>> 天气三类增量合计 = {int(w3['增减量'].sum()):,}；占净增长 = "
      f"{100*w3['增减量'].sum()/(n2-n1):.1f}%")

t7 = L("t07_yoy_by_agency")
for a in ["HPD", "DOT", "NYPD"]:
    row = t7[t7["机构代码"] == a].iloc[0]
    print(f">>> {a}: {int(row['增减量']):+,} ({row['同比%']}%), 占净增长 {row['占净增长%']}%")

t8 = L("t08_winter")
print(">>> 冬季对比:")
for c in t8.columns[1:]:
    a, b = int(t8.loc[0, c]), int(t8.loc[1, c])
    print(f"    {c}: {a:,} -> {b:,}  ({100*(b-a)/a:+.1f}%, {b/a:.2f}x)")

t27 = L("t27_top_types")
print(f">>> 前10类累计 = {t27.loc[9,'累计占比%']}%；前20类 = {t27.loc[19,'累计占比%']}%")

t25 = L("t25_channel_monthly")
f_, l_ = t25.iloc[0], t25.iloc[-2]
print(f">>> 渠道 {f_['月份']} 网页{f_['网页 ONLINE%']}/电话{f_['电话 PHONE%']} → "
      f"{l_['月份']} 网页{l_['网页 ONLINE%']}/电话{l_['电话 PHONE%']}")
