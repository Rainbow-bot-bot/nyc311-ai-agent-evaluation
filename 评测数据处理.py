"""评测数据整理入口。

输入：_实验系统/telemetry/benchmark.sqlite
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


def run(database=DEFAULT_DB):
    """整理评测数据并刷新查证入口。"""
    core = load_core()
    before_hash = hashlib.sha256(database.read_bytes()).hexdigest()

    tables, evidence, _, sql_queries = core.extract(database)
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
    args = parser.parse_args()

    core = load_core()
    if args.evidence:
        _, evidence, _, _ = core.extract(args.database)
        if args.evidence not in evidence:
            raise ValueError("不存在此证据编号")
        print(json.dumps(evidence[args.evidence], ensure_ascii=False, indent=2))
        return

    run(args.database)


if __name__ == "__main__":
    main()
