"""
预处理：从只读原始 Parquet 分片构建精简中间表 (slim parquet)。
- 只读 D:\\项目1\\原始数据，不复制/改写原始文件。
- 派生字段：created_dt, closed_dt, created_month, resolution_hours, is_closed, borough_clean。
- 输出：work/slim.parquet （仅用于本项目复现的中间产物）
保留可追溯逻辑：所有派生规则集中在此文件。
"""
import glob, os
import pandas as pd
import numpy as np
import pyarrow as pa
import pyarrow.parquet as pq

SRC = r"D:\项目1\原始数据"
OUT = r"D:\项目2\增强实验\Qwen\work"
os.makedirs(OUT, exist_ok=True)

KEEP = ["unique_key","created_date","closed_date","agency","complaint_type",
        "descriptor","borough","city","incident_zip","status",
        "open_data_channel_type","resolution_description"]

def parse_dt(s):
    # 原始为字符串 'YYYY-MM-DDTHH:MM:SS.sss'；空/NaN -> NaT
    return pd.to_datetime(s, errors="coerce", format="ISO8601")

def norm(s):
    # 去除首尾空白；空串 -> NA
    s = s.astype("string").str.strip()
    return s.replace({"": pd.NA, "nan": pd.NA, "None": pd.NA, "NaN": pd.NA})

files = sorted(glob.glob(os.path.join(SRC, "*.parquet")))
writer = None
total = 0
for f in files:
    df = pd.read_parquet(f, columns=KEEP)
    out = pd.DataFrame()
    out["unique_key"] = pd.to_numeric(df["unique_key"], errors="coerce").astype("int64")
    out["created_dt"] = parse_dt(df["created_date"])
    out["closed_dt"]  = parse_dt(df["closed_date"])
    out["created_month"] = out["created_dt"].dt.strftime("%Y-%m")
    for c in ["agency","complaint_type","descriptor","city","status",
              "open_data_channel_type"]:
        out[c] = norm(df[c])
    out["borough_raw"] = norm(df["borough"])
    # borough 清洗：把 'Unspecified' 归为 NA（真实事件无法定位行政区）
    out["borough"] = out["borough_raw"].replace({"Unspecified": pd.NA})
    out["incident_zip"] = norm(df["incident_zip"])
    out["channel"] = out["open_data_channel_type"].replace({"UNKNOWN": pd.NA})
    # is_closed：以 closed_dt 是否存在为准（status 字段与 closed_dt 高度一致，见 QA）
    out["is_closed"] = out["closed_dt"].notna()
    # resolution_hours：仅对已关闭记录有意义
    delta = (out["closed_dt"] - out["created_dt"]).dt.total_seconds() / 3600.0
    out["resolution_hours"] = delta.where(out["is_closed"])
    # 记录异常：已关闭但 closed < created（负时长）
    out = out[["unique_key","created_dt","closed_dt","created_month","agency",
               "complaint_type","descriptor","borough","borough_raw","city",
               "incident_zip","status","channel","is_closed","resolution_hours"]]
    # 转为 pyarrow table 写出（category 压缩字符串）
    for c in ["created_month","agency","complaint_type","descriptor","borough",
              "borough_raw","city","status","channel"]:
        out[c] = out[c].astype("category")
    tbl = pa.Table.from_pandas(out, preserve_index=False)
    if writer is None:
        writer = pq.ParquetWriter(os.path.join(OUT,"slim.parquet"), tbl.schema,
                                  compression="zstd")
    writer.write_table(tbl)
    total += len(out)
    print(os.path.basename(f), len(out), "cum", total, flush=True)

writer.close()
print("DONE total rows written:", total)
