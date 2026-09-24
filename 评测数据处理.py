"""评测数据整理入口。

输入：正式Excel的S分项及批注；SQLite保留运行、用量和旧评分记录。
输出：查证页、取数 SQL 和当前表结构。
"""

from pathlib import Path
import argparse
import hashlib
import importlib.util
import json

ROOT = Path(__file__).resolve().parent
DEFAULT_DB = ROOT / "_实验系统" / "telemetry" / "benchmark.sqlite"


def load_core():
    """加载内部整理模块。"""
    path = ROOT / "_实验系统" / "评测数据处理核心.py"
    spec = importlib.util.spec_from_file_location("project2_cleaning_core", path)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def read_current_scores(workbook=None):
    """只读正式 Excel 的 S 分项和批注；不保存或重排工作簿。"""
    import re
    from openpyxl import load_workbook
    workbook = workbook or ROOT / "交付成果/项目2_AI评测分析.xlsx"
    standard = (ROOT / "资料/评判标准.md").read_text(encoding="utf-8-sig")
    criteria = {code: float(maximum) for code, maximum in re.findall(
        r"^\| ([A-F]\d+) [^|]+\|\s*([0-9.]+)\s*\|", standard, re.M)}
    if len(criteria) != 25 or sum(criteria.values()) != 100:
        raise ValueError("现行标准应包含25项，满分合计100")
    fields = {"实际工作与缺口": "实际完成与缺口", "原件": "原件目录",
              "核查依据": "依据", "核查方式": "核查方式", "评测方补做": "评测方补做"}
    book = load_workbook(workbook, data_only=True)
    scores = {}
    try:
        sheet = book["评分与用量"]
        for row in sheet:
            if row[2].value != "S分项":
                continue
            ai, round_label, code = row[0].value, row[1].value, row[9].value
            if round_label not in ("第一轮", "第二轮") or code not in criteria:
                raise ValueError(f"未知轮次或条款：{ai}/{round_label}/{code}")
            rd = 1 if round_label == "第一轮" else 2
            key = f"{ai}-{rd}-{code}"
            value, maximum = row[4].value, criteria[code]
            if not isinstance(value, (int, float)) or value / maximum not in (0, .25, .5, .75, 1):
                raise ValueError(f"{key} 得分不符合统一档位：{value}")
            if key in scores or not row[4].comment:
                raise ValueError(f"{key} 重复或缺少得分批注")
            comment = row[4].comment.text
            labels = {m.group(1): m.group(2).strip() for m in re.finditer(
                r"^(评分项|实际工作与缺口|原件|核查依据|核查方式|问题编号|评测方补做)：(.*?)(?=\n(?:评分项|实际工作与缺口|原件|核查依据|核查方式|问题编号|评测方补做)：|\Z)",
                comment, re.M | re.S)}
            if any(not labels.get(k) for k in fields):
                raise ValueError(f"{key} 批注缺少必要字段")
            issue = row[10].value
            if not issue or labels.get("问题编号") != issue:
                raise ValueError(f"{key} K列与批注的问题编号不一致")
            heading = re.fullmatch(r"(.+?)；得分\s*([0-9.]+)/([0-9.]+)", labels.get("评分项", ""))
            if not heading or heading[1] != row[3].value or float(heading[2]) != value or float(heading[3]) != maximum:
                raise ValueError(f"{key} 单元格与批注首行的评分项或分数不一致")
            item = {"AI": ai, "轮次": rd, "量表": "v2.2", "条款": code,
                    "评分项": row[3].value.split(" ", 1)[1], "满分": maximum,
                    "档位": value / maximum, "得分": value, "状态": "已评"}
            item.update({target: labels[label] for label, target in fields.items()})
            scores[key] = {"现行评分版本": "S v2.2", "评分原文": item,
                           "关联问题编号": issue, "来源": "交付成果/项目2_AI评测分析.xlsx：评分与用量",
                           "Excel得分批注": comment}
    finally:
        book.close()
    pairs = {(ai, rd) for ai in ("GPT", "Grok", "DeepSeek", "GLM", "Qwen", "Gemini") for rd in (1, 2)} | {("Claude", 1)}
    expected = {f"{ai}-{rd}-{code}" for ai, rd in pairs for code in criteria}
    if set(scores) != expected:
        raise ValueError("S分项应覆盖六系统两轮及Claude第一轮，共325条；Kimi和Claude第二轮留空")
    return scores


def load_evidence(path=None):
    path = path or ROOT / "交付成果/查证数据.js"
    payload = path.read_text(encoding="utf-8-sig")
    return json.loads(payload.split("=", 1)[1].rstrip(";\n\r "))


def merge_current_scores(evidence, scores):
    """更新评分编号及共享问题编号下的评分，保留原始问题证据。"""
    from copy import deepcopy
    from collections import defaultdict
    merged = deepcopy(evidence)
    for value in merged.values():
        value.pop("现行S关联评分", None)
    grouped = defaultdict(list)
    for key, value in scores.items():
        merged[key] = deepcopy(value)
        grouped[value["关联问题编号"]].append(value["评分原文"])
    for issue, items in grouped.items():
        if issue not in scores:
            merged.setdefault(issue, {})["现行S关联评分"] = items
    return merged


def verify_current_scores(evidence=None, scores=None):
    """按325个评分编号逐字段核对分值与依据，不重新评定分数。"""
    scores = scores if scores is not None else read_current_scores()
    evidence = evidence if evidence is not None else load_evidence()
    actual_keys = {k for k, v in evidence.items() if v.get("现行评分版本") == "S v2.2"}
    if actual_keys != set(scores):
        raise ValueError("查证页的现行评分编号集合与Excel不一致")
    for key, expected in scores.items():
        actual = evidence[key]
        for field, value in expected["评分原文"].items():
            if actual.get("评分原文", {}).get(field) != value:
                raise ValueError(f"{key} 查证数据不一致：{field}")
        if actual.get("关联问题编号") != expected["关联问题编号"]:
            raise ValueError(f"{key} 关联问题编号不一致")
        issue = expected["关联问题编号"]
        if issue not in scores:
            linked = evidence.get(issue, {}).get("现行S关联评分", [])
            if expected["评分原文"] not in linked:
                raise ValueError(f"{key} 未同步到共享问题编号 {issue}")
    return {"评分条数": len(scores), "分值与依据": "逐字段一致", "评分机制": "S v2.2"}


def sync_current_scores():
    """只刷新查证数据，保留查证HTML、原始记录和Excel格式。"""
    scores = read_current_scores()
    evidence = merge_current_scores(load_evidence(), scores)
    result = verify_current_scores(evidence, scores)
    path = ROOT / "交付成果/查证数据.js"
    temporary = path.with_suffix(".js.tmp")
    temporary.write_text("window.PROJECT2_EVIDENCE=" + json.dumps(evidence, ensure_ascii=False, separators=(",", ":")) + ";\n", encoding="utf-8")
    temporary.replace(path)
    verify_current_scores(load_evidence(), scores)
    return result


def retain_current_assessment(evidence):
    """保留人工核定的现行S评分及运行说明。"""
    path = ROOT / "交付成果" / "查证数据.js"
    if not path.exists():
        return
    payload = path.read_text(encoding="utf-8-sig")
    current = json.loads(payload.split("=", 1)[1].rstrip(";\n\r "))
    for key, value in current.items():
        if value.get("现行评分版本") == "S v2.2":
            evidence[key] = value
        elif "现行S关联评分" in value:
            evidence.setdefault(key, {})["现行S关联评分"] = value["现行S关联评分"]
        elif key in evidence and "现行S说明" in value:
            evidence[key]["现行S说明"] = value["现行S说明"]
    for key, value in evidence.items():
        if key.startswith("RUN:"):
            if "本次Gate裁定" in value:
                value["历史Gate裁定（旧口径，不用于现行S配对）"] = value.pop("本次Gate裁定")
            if "评分版本" in value:
                value["历史评分版本"] = value.pop("评分版本")


def run(database=DEFAULT_DB):
    """整理评测数据并刷新查证入口。"""
    core = load_core()
    before_hash = hashlib.sha256(database.read_bytes()).hexdigest()

    tables, evidence, _, sql_queries = core.extract(database)
    retain_current_assessment(evidence)
    evidence = merge_current_scores(evidence, read_current_scores())
    field_rows = core.build_field_dictionary_rows(tables)
    evidence_page, evidence_payload = core.write_evidence_page(evidence)

    manifest = {
        "tables": {name: len(rows) for name, rows in tables},
        "field_dictionary_rows": len(field_rows),
    }
    (ROOT / "_实验系统" / "当前导出.json").write_text(
        json.dumps(manifest, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    (ROOT / "_实验系统" / "sql" / "分析取数.sql").write_text(
        "\n".join(sql_queries),
        encoding="utf-8",
    )

    after_hash = hashlib.sha256(database.read_bytes()).hexdigest()
    if before_hash != after_hash:
        raise RuntimeError("SQLite 哈希变化")

    result = {
        **manifest,
        "evidence_shell_bytes": evidence_page.stat().st_size,
        "evidence_payload_bytes": evidence_payload.stat().st_size,
        "sql_database_sha256": before_hash,
        "sql_unchanged": True,
        "no_model_calls": True,
    }
    print(json.dumps(result, ensure_ascii=False))
    return result


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--database", type=Path, default=DEFAULT_DB)
    parser.add_argument("--evidence", help="查询证据编号")
    parser.add_argument("--sync-scores", action="store_true", help="从正式Excel同步325条S得分和批注到查证数据")
    parser.add_argument("--check-scores", action="store_true", help="只读核对Excel与查证数据的325条S评分")
    args = parser.parse_args()
    if args.sync_scores or args.check_scores:
        result = sync_current_scores() if args.sync_scores else verify_current_scores()
        print(json.dumps(result, ensure_ascii=False))
        return

    core = load_core()
    if args.evidence:
        _, evidence, _, _ = core.extract(args.database)
        retain_current_assessment(evidence)
        evidence = merge_current_scores(evidence, read_current_scores())
        if args.evidence not in evidence:
            raise ValueError("不存在此证据编号")
        print(json.dumps(evidence[args.evidence], ensure_ascii=False, indent=2))
        return

    run(args.database)


if __name__ == "__main__":
    main()
