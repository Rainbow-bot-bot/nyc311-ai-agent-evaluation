"""评测数据整理核心：SQLite → 分析表 → 字段说明与查证。"""

from pathlib import Path
from copy import deepcopy
import argparse
import hashlib
import importlib.util
import json
import math
import re
import sqlite3

ROOT = Path(__file__).resolve().parents[1]

CRITERIA = [
    "A1", "A2", "A3", "A4",
    "B1", "B2", "B3", "B4",
    "C1", "C2", "C3", "C4", "C5",
    "D1", "D2", "D3", "D4",
    "E1", "E2", "E3", "E4",
    "R1", "R2", "R3", "R4", "R5",
]

STATUS = {
    "completed": 1,
    "failed": 2,
    "mixed_excluded": 3,
    "withdrawn_excluded": 4,
    "not_started_clean_run": 5,
    "infrastructure_failure_before_model": 6,
    "telemetry_invalid": 7,
    "manual_partial_unscored": 8,
    "status_unverified": 9,
}

SAMPLE = {
    "frozen_baseline": 1,
    "paired_enhanced": 2,
    "supplemental_only": 3,
    "not_in_main": 0,
    "excluded": 0,
    "excluded_or_pending": 0,
}

OPS = {
    "Read": 1,
    "view_file": 1,
    "Glob": 2,
    "Grep": 2,
    "find_by_name": 2,
    "list_dir": 2,
    "Write": 3,
    "write_to_file": 3,
    "Edit": 4,
    "StrReplace": 4,
    "Bash": 5,
    "Shell": 5,
    "PowerShell": 5,
    "run_command": 5,
    "exec": 6,
    "TaskCreate": 7,
    "TaskUpdate": 7,
    "Skill": 8,
    "GetDynamicTools": 8,
    "CallDynamicTool": 8,
}

LABEL = {
    "contributory": 1,
    "necessary_exploration": 2,
    "administrative": 3,
    "waste": 4,
}

VERDICT = {
    "PASS": 1,
    "numeric_match": 1,
    "SUPPORTED": 1,
    "数值支持": 1,
    "子集数值支持": 1,
    "披露支持": 1,
    "边界明确": 1,
    "SUPPORTED_WITH_LIMITS": 2,
    "比较仍有限制": 2,
    "PARTIAL": 3,
    "PARTIAL_WITH_ERROR": 4,
    "FAIL": 4,
    "numeric_mismatch": 4,
    "数值不一致": 4,
    "口径须修正": 4,
    "UNSUPPORTED_HEADLINE": 5,
    "解释越界": 5,
    "解释过强": 5,
    "说明未披露": 6,

}

TABLE_KIND = {
    "evaluation_numeric_check": 1,
    "evaluation_headline": 2,
    "autonomy_delivery_checks": 3,
    "autonomy_repairs": 4,
}


def verdict_code(value):
    if value.startswith("待修正"):
        return 7
    if "HOLD" in value:
        return 8
    return VERDICT[value]


def h(data):
    """计算 SHA256。"""
    return hashlib.sha256(data).hexdigest()


def number(value):
    """把可识别的数字文本转成数值；未知值保持原样或空。"""
    if value is None or isinstance(value, bool):
        return value
    if isinstance(value, (float, int)):
        return value if math.isfinite(value) else None

    text = str(value).strip().replace(",", "")
    if re.fullmatch(r"[-+]?\d+(?:\.\d+)?", text):
        return float(text)
    return None


def _open_database(db):
    connection = sqlite3.connect(
        Path(db).resolve().as_uri() + "?mode=ro",
        uri=True,
    )
    connection.row_factory = sqlite3.Row
    return connection


def _load_behavior_datasets(connection):
    path = ROOT / "_实验系统" / "telemetry" / "behavior_from_sql.py"
    spec = importlib.util.spec_from_file_location("behavior", path)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)

    return {
        table_name: rows
        for _, table_name, rows, _ in module.datasets(connection)
    }


def _decoded_records(connection, table, version):
    """读取 record_json，并按 run_id 建索引。"""
    sql = f"SELECT run_id, record_json FROM {table} WHERE version=?"
    rows = connection.execute(sql, (version,))
    return {
        row["run_id"]: json.loads(row["record_json"])
        for row in rows
    }


def _score_version(participation):
    """根据参与记录选择正式评分来源。"""
    if participation["inclusion"] == "frozen_baseline":
        return "1.0"
    if participation["inclusion"] == "paired_enhanced":
        return "enhanced_review_1.0"
    if (
        participation["ai"] == "Claude"
        and participation["record_status"] == "mixed_excluded"
    ):
        return "claude_assisted_quality_20260918_v1"
    return None


def _score_status_code(version):
    if version == "1.0":
        return 1
    if version == "enhanced_review_1.0":
        return 2
    if version == "claude_assisted_quality_20260918_v1":
        return 3
    return 0


def _uncached_input(run, metric, participation):
    """统一未缓存输入 Token 的定义。"""
    raw_input = run.get("input_tokens")
    cache_read = run.get("cache_read_tokens")

    if (
        run.get("host_tool") == "codex"
        and raw_input is not None
        and cache_read is not None
    ):
        uncached = raw_input - cache_read
    else:
        uncached = raw_input

    if participation["round_no"] == 1:
        uncached = metric.get("uncached_input_tokens", uncached)

    return uncached


def _build_raw_runs(connection, datasets, evidence):
    """把参与记录、运行记录、用量和评分合成运行基础表。"""
    all_runs = {
        row["run_id"]: dict(row)
        for row in connection.execute("SELECT * FROM benchmark_run")
    }
    participations = [
        dict(row)
        for row in connection.execute(
            "SELECT * FROM evaluation_participation "
            "ORDER BY participation_id"
        )
    ]

    supervision = _decoded_records(
        connection,
        "autonomy_supervision",
        "autonomy_review_20260917_v1",
    )
    usage = _decoded_records(
        connection,
        "autonomy_usage",
        "autonomy_review_20260917_v1",
    )
    conditions = _decoded_records(
        connection,
        "autonomy_conditions",
        "autonomy_review_20260917_v1",
    )

    metrics = {
        row["run_id"]: dict(row)
        for row in connection.execute(
            "SELECT * FROM evaluation_run_metrics "
            "WHERE version IN ('1.0','enhanced_review_1.0')"
        )
    }
    scores = {
        (row["version"], row["ai"], row["criterion"]): row["score"]
        for row in connection.execute("SELECT * FROM evaluation_score_item")
    }
    coverage = {
        row["run_id"]: row
        for row in datasets["tblBehaviorCoverage"]
        if row["run_id"]
    }

    result = []
    for participation in participations:
        run_id = participation["run_id"]
        run = all_runs.get(run_id, {}) if run_id else {}
        metric = metrics.get(run_id, {})
        run_usage = usage.get(run_id, {})
        supervision_row = supervision.get(run_id, {})
        coverage_row = coverage.get(run_id, {})
        condition_row = conditions.get(run_id, {})

        version = _score_version(participation)
        cache_read = run.get("cache_read_tokens")
        uncached_input = _uncached_input(
            run,
            metric,
            participation,
        )

        row = {
            "参与ID": participation["participation_id"],
            "AI": participation["ai"],
            "轮次": participation["round_no"],
            "run_id": run_id,
            "宿主": run.get("host_tool"),
            "模型": run.get("model_name"),
            "状态码": STATUS.get(
                run.get("status", participation["record_status"]),
                9,
            ),
            "参与记录状态码": STATUS.get(
                participation["record_status"],
                9,
            ),
            "样本码": SAMPLE[participation["inclusion"]],
            "评分状态码": _score_status_code(version),
            "原生秒": supervision_row.get(
                "原生观察秒",
                metric.get("native_seconds"),
            ),
            "外部秒": supervision_row.get(
                "外部秒",
                metric.get("external_wall_seconds"),
            ),
            "未缓存输入": run_usage.get(
                "核后未缓存输入",
                uncached_input,
            ),
            "缓存读": run_usage.get(
                "核后缓存读",
                metric.get("cache_read_tokens", cache_read),
            ),
            "缓存写": run_usage.get(
                "核后缓存写",
                metric.get(
                    "cache_write_tokens",
                    run.get("cache_write_tokens"),
                ),
            ),
            "输出Token": run_usage.get(
                "核后输出",
                run.get("output_tokens"),
            ),
            "实付元": None,
            "费用声明元": metric.get("actual_incremental_cny"),
            "追加文本次数": supervision_row.get("追加用户文本指令"),
            "权限点击次数": supervision_row.get("权限点击数"),
            "人工分钟": supervision_row.get("人工分钟"),
            "额外启动清单": condition_row.get("额外启动方法清单"),
            "可查动作数": coverage_row.get("可查询动作数"),
            "有返回文本动作数": coverage_row.get("有结果文本动作数"),
        }

        for criterion in CRITERIA:
            if version:
                row[criterion] = scores.get(
                    (version, participation["ai"], criterion)
                )
            else:
                row[criterion] = None

        evidence_id = "RUN:" + participation["participation_id"]
        row["证据ID"] = evidence_id
        result.append(row)

        evidence[evidence_id] = {
            "原库表": "evaluation_participation",
            "主键": participation["participation_id"],
            "参与原文": participation,
            "运行原文": run,
            "监督": supervision_row,
            "用量": run_usage,
            "指标版本": metric,
            "评分版本": version,
        }

    return result


def _action_theme_codes(event, themes):
    candidates = event["检查类型候选"].split("；")
    return [
        themes[name]
        for name in candidates
        if name in themes
    ]


def _build_raw_actions(connection, datasets, evidence):
    """整理工具动作，并接入已冻结的第二轮动作二审。"""
    reviews = {
        (row["ai"], row["sequence_no"]): dict(row)
        for row in connection.execute(
            "SELECT * FROM evaluation_action "
            "WHERE version='enhanced_action_review_20260918_v1'"
        )
    }

    themes = {
        "数据摸底": 1,
        "比较与拆解": 2,
        "异常下钻": 3,
        "异常敏感性": 4,
        "固定构成": 5,
        "窗口与删失": 6,
        "细分类与口径": 7,
        "结果验收": 8,
        "任务管理/不计实际分析": 9,
    }

    result = []
    last_read = {}

    for event in datasets["tblBehaviorEvents"]:
        evidence_id = event["行为ID"]
        run_id = event["run_id"]
        arguments = event["实际参数原文"]
        result_text = event["检查结果原文"]
        codes = _action_theme_codes(event, themes)

        repeated_read = None
        if OPS.get(event["工具"]) == 1:
            previous = last_read.get(run_id)
            if previous is not None:
                repeated_read = int(
                    previous == (event["工具"], arguments)
                )
        last_read[run_id] = (event["工具"], arguments)

        if len(codes) == 1:
            theme_code = codes[0]
        elif len(codes) > 1:
            theme_code = 98
        else:
            theme_code = 0

        if result_text:
            has_error_text = int(
                bool(
                    re.search(
                        r"Traceback|Error:|exited with code [1-9]|"
                        r'exit_code[\\"\s:]+[1-9]',
                        result_text,
                    )
                )
            )
        else:
            has_error_text = None

        row = {
            "行为ID": evidence_id,
            "AI": event["AI"],
            "轮次": 1 if event["轮次"] == "第一轮" else 2,
            "run_id": run_id,
            "动作序号": event["序号"],
            "操作码": OPS.get(event["工具"], 0),
            "工具": event["工具"],
            "类型候选码": theme_code,
            "类型判定码": 1 if codes else 0,
            "原采集成功": event["原采集成功标记"],
            "返回可见": event["结果文本可见"],
            "返回含错误标记": has_error_text,
            "参数截断": int(
                event["参数原文字符数"] > len(arguments)
            ),
            "结果截断": event.get("结果摘录是否截断"),
            "观测秒": event["观测耗时秒"],
            "修改请求": int(
                event["修改前原文"] is not None
                and event["修改后原文"] is not None
            ),
            "修改确认": event["已确认内容修改"],
            "相邻同参重读": repeated_read,
            "历史动作标签码": 0,
            "标签版本码": 0,
            "证据ID": "ACT:" + evidence_id,
        }

        review = None
        if event["轮次"] == "第二轮":
            review = reviews.get(
                (event["AI"], event["序号"])
            )

        if review:
            row["历史动作标签码"] = LABEL[review["label"]]
            row["标签版本码"] = 3
        else:
            row["历史动作标签码"] = LABEL.get(
                event["已有动作标签"],
                0,
            )
            if event["已有评审版本"] == "1.0":
                row["标签版本码"] = 1

        for name, code in themes.items():
            row["目的提及_" + name] = int(code in codes)

        result.append(row)

        source = deepcopy(event)
        if review:
            source["动作评审依据"] = {
                "原库表": "evaluation_action",
                "主键": [
                    "enhanced_action_review_20260918_v1",
                    event["AI"],
                    event["序号"],
                ],
                "记录": review,
            }
        evidence[row["证据ID"]] = source

    return result


def _append_raw_check(
    checks,
    evidence,
    key,
    ai,
    round_no,
    run_id,
    kind,
    item,
    status,
    owner,
    value_1=None,
    value_2=None,
    value_3=None,
    unit=None,
    change=None,
    attempt=None,
    closed=None,
    source=None,
):
    row = {
        "核验ID": key,
        "AI": ai,
        "轮次": round_no,
        "run_id": run_id,
        "记录类型码": kind,
        "核验项目": str(item)[:60],
        "结论状态码": status,
        "问题归属码": owner,
        "源观测值": number(value_1),
        "参照值": number(value_2),
        "替代基准值": number(value_3),
        "单位码": unit,
        "判断变化": change,
        "修正尝试": attempt,
        "问题闭合": closed,
        "证据ID": "CHK:" + key,
    }
    checks.append(row)
    evidence[row["证据ID"]] = source


def _case_owner(source):
    owner = source["归属"]
    if owner in ("评测方勘误", "实验方", "评测方"):
        return 2
    if "身份" in owner:
        return 3
    if owner == "模型修正":
        return 1
    return 0


def _build_raw_checks(connection, datasets, evidence):
    """把数值核验、闭环检查和敏感性复算整理成统一记录。"""
    checks = []

    for source in datasets["tblBehaviorCases"]:
        kind = TABLE_KIND[source["原库表"]]
        key_parts = json.loads(source["原库主键"])
        if kind == 1:
            item = key_parts[-1]
            value_1 = source["检查前判断"]
            value_2 = source["检查结果"]
        else:
            item = source["分析问题"]
            value_1 = None
            value_2 = None

        _append_raw_check(
            checks,
            evidence,
            source["实例ID"],
            source["AI"],
            1 if source["轮次"] == "第一轮" else 2,
            source["run_id"],
            kind,
            item,
            verdict_code(source["数值或结论状态"]),
            _case_owner(source),
            value_1=value_1,
            value_2=value_2,
            closed=source["是否最终修正"],
            source=source,
        )

    query = (
        "SELECT * FROM autonomy_delivery_checks "
        "WHERE version='astra_pair_evidence_20260917_v1' "
        "ORDER BY record_key"
    )
    unit_codes = {
        "条": 1,
        "份额": 2,
        "秒": 4,
        "月": 6,
        "张": 7,
        "个": 7,
        "次": 7,
        "已执行标记": 8,
    }
    for row in connection.execute(query):
        data = json.loads(row["record_json"])
        source = {
            "原库表": "autonomy_delivery_checks",
            "主键": [row["version"], row["record_key"]],
            "原记录": data,
        }
        _append_raw_check(
            checks,
            evidence,
            data["证据编号"],
            data["AI"],
            1 if data["轮次"] == "第一轮" else 2,
            data["run_id"],
            5,
            data["观测项"],
            0,
            0,
            value_1=data["数值"],
            unit=unit_codes.get(data["单位"]),
            source=source,
        )

    for data in datasets["tblBehaviorLoops"]:
        _append_raw_check(
            checks,
            evidence,
            data["闭环ID"],
            data["AI"],
            1 if data["轮次"] == "第一轮" else 2,
            data["run_id"],
            6,
            data["分析问题"],
            0,
            1,
            change=data["是否改变所列叙述或路线"],
            attempt=data["是否自主修正所列问题"],
            closed=data["所列问题是否闭合"],
            source=data,
        )

    query = (
        "SELECT * FROM autonomy_delivery_checks "
        "WHERE version='matched_window_20260917_v1' "
        "ORDER BY record_key"
    )
    for row in connection.execute(query):
        data = json.loads(row["record_json"])
        source = {
            "原库表": "autonomy_delivery_checks",
            "主键": [row["version"], row["record_key"]],
            "原记录": data,
        }
        _append_raw_check(
            checks,
            evidence,
            "WIN:" + row["record_key"],
            "Qwen",
            2,
            row["run_id"],
            7,
            data["类别"],
            0,
            2,
            value_1=data["原基期11月中位小时"],
            value_2=data["报告期9月中位小时"],
            value_3=data["同月基期9月中位小时"],
            unit=3,
            source=source,
        )

    query = (
        "SELECT * FROM autonomy_delivery_checks "
        "WHERE version='behavior_common_review_20260918_v1' "
        "ORDER BY record_key"
    )
    for row in connection.execute(query):
        data = json.loads(row["record_json"])
        source = {
            "原库表": "autonomy_delivery_checks",
            "主键": [row["version"], row["record_key"]],
            "原记录": data,
        }
        _append_raw_check(
            checks,
            evidence,
            "BEH:" + row["record_key"],
            data["AI"],
            2,
            row["run_id"],
            8,
            data["行为类别"],
            0,
            0,
            source=source,
        )

    return checks


def _audit_queries():
    return [
        "SELECT * FROM evaluation_participation ORDER BY participation_id;",
        "SELECT * FROM benchmark_run;",
        "SELECT * FROM benchmark_tool_call ORDER BY run_id,sequence_no;",
        "SELECT s.*,b.content FROM evaluation_source_dataset s "
        "JOIN evaluation_raw_blob b USING(sha256);",
        "SELECT * FROM evaluation_score_item;",
        "SELECT * FROM evaluation_numeric_check;",
        "SELECT * FROM evaluation_headline;",
        "SELECT * FROM autonomy_delivery_checks;",
        "SELECT * FROM autonomy_repairs;",
        "SELECT * FROM autonomy_supervision;",
        "SELECT * FROM autonomy_usage;",
        "SELECT * FROM autonomy_conditions;",
    ]


def _internal_dictionary():
    """内部码值说明。"""
    return {
        "问题归属码": {
            0: "未归属",
            1: "模型",
            2: "评测方",
            3: "文件身份",
        },
        "类型候选码": {
            0: "未匹配",
            1: "数据摸底",
            2: "比较拆解",
            3: "异常下钻",
            4: "敏感性",
            5: "固定构成",
            6: "窗口删失",
            7: "细类口径",
            8: "验收",
            9: "任务管理",
            98: "多个候选",
        },
        "运行状态": STATUS,
        "样本类别": SAMPLE,
        "操作码": OPS,
        "动作标签": LABEL,
        "结论状态": VERDICT,
        "记录类型": {
            "1": "数值核验",
            "2": "结论评审",
            "3": "共同主题核查",
            "4": "首交修正任务",
            "5": "实例观测值",
            "6": "定向问题前后核查",
            "7": "同月敏感性观测",
        },
        "单位码": {
            "1": "记录数",
            "2": "份额",
            "3": "小时",
            "4": "秒",
            "6": "月数",
            "7": "个/次/张",
            "8": "二值标记",
        },
        "约定": "1=是；0=否；空白=未知/无值。",
        "评分": "第一轮1.0；第二轮enhanced_review_1.0；Claude仅成果质量。",
        "证据": "证据ID对应原库主键、来源哈希和评审记录。",
        "闭环": "attempt=修正尝试；closed=问题闭合。",
        "数值": "源观测值/参照值为同一核验记录两端。",
    }


# 1. 从冻结 SQLite 提取运行、动作和核验的原始结构。
def final_adjudication(tables, evidence):
    path = ROOT / '_实验系统/报告素材/最终评分裁定.json'
    ruling = json.loads(path.read_text(encoding='utf8'))
    data = dict(tables)
    for row in data['评分与用量']:
        code = row['指标'].split()[0]
        if row['AI']=='GPT' and row['轮次']=='第二轮' and code in ('C2','C4'):
            assert row['数值']==ruling['GPT'][code]['原分'], '裁定对应的原评分已变更'
            old = deepcopy(evidence[row['证据编号']])
            row['数值'] = ruling['GPT'][code]['新分']
            row['数据版本'] = ruling['版本']
            evidence[row['证据编号']] = {'原评分':old,'本次裁定':ruling['GPT'][code],'裁定来源':str(path),'裁定SHA256':h(path.read_bytes())}
    parent = next(r for r in data['运行记录'] if r['AI']=='Qwen' and r['轮次']=='第二轮')
    template = next(r for r in data['评分与用量'] if r['参与编号']==parent['参与编号'])
    original=sum(r['数值'] for r in data['评分与用量'] if r['参与编号']==parent['参与编号'] and r['指标类别']=='质量分项')
    assert original==ruling['Qwen']['原分项合计'] and original+ruling['Qwen']['调整']==min(original,ruling['Qwen']['封顶'])
    cap = {k:None for k in template}
    cap.update({'AI':'Qwen','轮次':'第二轮','指标类别':'质量封顶','指标':'Q_CAP 错误时间窗主线封顶调整','数值':ruling['Qwen']['调整'],'单位':'分','数据版本':ruling['版本'],'参与编号':parent['参与编号'],'运行编号':parent['运行编号'],'指标编号':'final-Qwen-cap','证据编号':'REVIEW:Qwen:CAP'})
    data['评分与用量'].append(cap)
    evidence[cap['证据编号']] = ruling['Qwen']
    for row in data['运行记录']:
        if row['AI']=='Gemini' and row['轮次']=='第二轮':
            row['参与类别']='补充观察'
            row['评分状态']='观察评分'
            evidence[row['证据编号']]['本次Gate裁定']=ruling['Gemini']
    evidence['REVIEW:final_20260921'] = ruling
    return [(name,data[name]) for name,_ in tables]


def extract(db):
    """读取原库，依次完成基础整理、样本筛选和指标派生。"""
    connection = _open_database(db)
    try:
        datasets = _load_behavior_datasets(connection)
        evidence = {}

        runs = _build_raw_runs(
            connection,
            datasets,
            evidence,
        )
        actions = _build_raw_actions(
            connection,
            datasets,
            evidence,
        )
        checks = _build_raw_checks(
            connection,
            datasets,
            evidence,
        )

        tables = readable_tables(
            [
                ("运行记录", runs),
                ("行为记录", actions),
                ("核验记录", checks),
            ],
            evidence,
            connection,
        )
        tables = analysis_scope(
            tables,
            evidence,
            connection,
        )
        tables = clarify_analysis_fields(
            tables,
            evidence,
            connection,
        )
        tables = review_extensions(
            tables,
            evidence,
            connection,
        )
        tables = derived_analysis_metrics(
            tables,
            evidence,
            connection,
        )
        tables = split_api_pricing(tables)
        tables = final_adjudication(tables, evidence)

        return (
            tables,
            evidence,
            _internal_dictionary(),
            _audit_queries(),
        )
    finally:
        connection.close()


ROUND_NAMES = {
    0: "校准",
    1: "第一轮",
    2: "第二轮",
}

RUN_STATES = {
    1: "已完成",
    2: "失败",
    3: "混合交付",
    4: "已退出",
    5: "未启动",
    6: "启动失败",
    7: "采集无效",
    8: "人工参与未评分",
    9: "待核实",
}

SAMPLE_NAMES = {
    0: "排除",
    1: "第一轮主样本",
    2: "第二轮主样本",
    3: "补充失败样本",
}

GRADE_NAMES = {
    0: "未评分",
    1: "第一轮",
    2: "第二轮",
    3: "成果质量",
}

HOST_NAMES = {
    "codex": "Codex",
    "cursor": "Cursor",
    "antigravity": "Antigravity",
    "claude_code": "Claude Code",
}

OP_NAMES = {
    0: "其他",
    1: "读取",
    2: "检索",
    3: "写入",
    4: "修改",
    5: "运行命令",
    6: "批量执行",
    7: "任务管理",
    8: "工具或方法加载",
}

ACTION_LABEL_NAMES = {
    0: "尚未评审",
    1: "直接贡献交付",
    2: "必要探索",
    3: "管理动作",
    4: "确认无收益",
}

CHECK_KIND_NAMES = {
    1: "数值核验",
    2: "结论评审",
    3: "共同主题核查",
    4: "交付后修正任务",
    5: "实例观测",
    6: "定向前后核查",
    7: "窗口敏感性观测",
    8: "行为片段复核",
}

CHECK_VERDICT_NAMES = {
    "PASS": "数值一致",
    "numeric_match": "数值一致",
    "FAIL": "数值不一致",
    "numeric_mismatch": "数值不一致",
    "SUPPORTED": "证据支持",
    "SUPPORTED_WITH_LIMITS": "支持但有限制",
    "PARTIAL": "部分支持",
    "PARTIAL_WITH_ERROR": "部分支持且有错误",
    "UNSUPPORTED_HEADLINE": "结论不获支持",
}

SCORE_NAMES = [
    "数据理解",
    "清洗判断",
    "问题选择",
    "范围收缩",
    "数字与口径",
    "证据可查",
    "比较可比",
    "解释边界",
    "增量拆解",
    "深入调查",
    "替代解释",
    "修正判断",
    "适时停止",
    "成品可读",
    "结论可核",
    "说明与复现",
    "最终文件检查",
    "复杂度选择",
    "有效复用",
    "范围管理",
    "资源判断",
    "独立承担",
    "自我纠错",
    "路线控制",
    "完成真实性",
    "协议纪律",
]


def _yes_no(value, missing=None):
    if value == 1:
        return "是"
    if value == 0:
        return "否"
    return missing


def _append_metric(
    metrics,
    evidence,
    connection,
    run,
    category,
    name,
    value,
    unit,
    version,
    detail=None,
):
    """追加一个长表指标，并在需要时回查正式评分来源。"""
    if value is None:
        return

    if not isinstance(value, (int, float)) or not math.isfinite(value):
        raise ValueError(f"指标数值无效：{name}={value!r}")

    metric_id = h(
        (
            run["参与ID"]
            + "|"
            + category
            + "|"
            + name
        ).encode()
    )[:20]
    evidence_id = "MET:" + metric_id

    metrics.append({
        "AI": run["AI"],
        "轮次": ROUND_NAMES[run["轮次"]],
        "指标类别": category,
        "指标": name,
        "数值": value,
        "单位": unit,
        "数据版本": version,
        "参与编号": run["参与ID"],
        "运行编号": run["run_id"],
        "指标编号": metric_id,
        "证据编号": evidence_id,
    })

    if category in ("质量分项", "可靠性分项"):
        score_version = evidence[run["证据ID"]]["评分版本"]
        criterion = name.split(" ", 1)[0]
        scored = connection.execute(
            "SELECT * FROM evaluation_score_item "
            "WHERE version=? AND ai=? AND criterion=?",
            (
                score_version,
                run["AI"],
                criterion,
            ),
        ).fetchone()
        if scored is None or scored["score"] != value:
            raise ValueError("评分来源与数值不一致")

        detail = {
            "原库表": "evaluation_score_item",
            "原库主键": [
                score_version,
                run["AI"],
                criterion,
            ],
            "参与编号": run["参与ID"],
            "运行编号": run["run_id"],
            "评分记录": dict(scored),
            "性质": "已保存评审",
        }

    if detail is not None:
        evidence[evidence_id] = deepcopy(detail)
    else:
        evidence[evidence_id] = {
            "原库证据编号": run["证据ID"],
            "指标": name,
            "来源": evidence[run["证据ID"]],
        }


def _pricing_context(connection):
    path = ROOT / "_实验系统" / "telemetry" / "model_pricing.json"
    catalog = json.loads(path.read_text(encoding="utf-8"))
    fx_row = connection.execute(
        "SELECT fx_rate_to_cny,fx_rate_date "
        "FROM benchmark_run "
        "WHERE host_tool='codex' AND condition_name='baseline'"
    ).fetchone()
    prices = {
        item["ai"]: item
        for item in catalog["models"].values()
    }
    return catalog, h(path.read_bytes()), fx_row, prices


def _gpt_max_input(connection, session_id):
    sources = connection.execute(
        "SELECT s.relative_path,b.content "
        "FROM evaluation_source_dataset s "
        "JOIN evaluation_raw_blob b USING(sha256) "
        "WHERE s.relative_path LIKE ?",
        ("%" + session_id + "%.jsonl",),
    )

    maxima = []
    for source in sources:
        text = bytes(source["content"]).decode("utf-8-sig")
        for line in text.splitlines():
            record = json.loads(line)
            info = (record.get("payload") or {}).get("info") or {}
            last_usage = info.get("last_token_usage")
            if last_usage:
                maxima.append(last_usage.get("input_tokens", 0))

    return max(maxima) if maxima else None


def _price_rates(
    connection,
    run,
    price,
    fx_row,
):
    """根据模型和已观测 Token 选择价格档。"""
    ai = run["AI"]
    extra = {}

    if ai == "GPT":
        maximum_input = _gpt_max_input(
            connection,
            run["运行原文"]["session_id"],
        )
        threshold = price["threshold_input_tokens"]
        if maximum_input is None or maximum_input > threshold:
            return (
                None,
                None,
                "逐请求档位缺失",
                extra,
            )

        rates = {
            key: price["short"][key]
            for key in (
                "input",
                "cache_read",
                "cache_write",
                "output",
            )
        }
        extra = {
            "单次最大输入": maximum_input,
            "固定比较汇率": fx_row[0],
            "汇率日期": fx_row[1],
        }
        return (
            rates,
            rates,
            "标准API价（固定汇率）",
            extra,
        )

    if ai == "DeepSeek":
        return (
            price["idle"],
            price["busy"],
            "闲时–忙时价",
            extra,
        )

    if ai == "Qwen":
        rates = {
            "input": price["input"],
            "cache_read": price["cache_read_explicit"],
            "cache_write": price["cache_create_explicit"],
            "output": price["output"],
        }
        return (
            rates,
            rates,
            "标准API价（显式缓存）",
            extra,
        )

    rates = {
        "input": price["input"],
        "cache_read": price["cache_read"],
        "cache_write": 0,
        "output": price["output"],
    }
    if run["缓存写"] != 0:
        raise ValueError("缓存写单价缺失")

    return (
        rates,
        rates,
        "标准API价",
        extra,
    )


def _add_api_pricing(
    connection,
    run,
    evidence,
    metrics,
    catalog,
    catalog_hash,
    fx_row,
    prices,
):
    """添加 API 目录价参数。"""
    if not run["run_id"]:
        return "Token缺失"

    if run["轮次"] == 0:
        return "校准"

    token_values = [
        run["未缓存输入"],
        run["缓存读"],
        run["缓存写"],
        run["输出Token"],
    ]
    if any(value is None for value in token_values):
        return "Token缺失"

    price = prices[run["AI"]]
    low, high, explanation, extra = _price_rates(
        connection,
        {
            **run,
            "运行原文": evidence[run["证据ID"]]["运行原文"],
        },
        price,
        fx_row,
    )

    if low is None:
        return explanation

    detail = {
        "来源": "冻结API价目参数，非模型运行观测",
        "价目日期": catalog["checked_at"],
        "价目复核记录日期": "2026-09-18",
        "来源网址": price["source"],
        "价目SHA256": catalog_hash,
        "价格": price,
        "比较口径": explanation,
        "补充": extra,
        "零缓存写说明": (
            "未公布写价但本次写入Token为0时，此项费用系数记0；"
            "不表示免费写缓存"
        ),
    }

    token_names = [
        ("input", "未缓存输入"),
        ("cache_read", "缓存读取"),
        ("cache_write", "缓存写入"),
        ("output", "输出"),
    ]

    for key, title in token_names:
        for label, rates in (
            ("低", low),
            ("高", high),
        ):
            if (
                key not in rates
                and not (
                    key == "cache_write"
                    and run["缓存写"] == 0
                )
            ):
                raise ValueError("API 价格字段缺失")

            _append_metric(
                metrics,
                evidence,
                connection,
                run,
                "API单价",
                title + "单价" + label,
                rates.get(key, 0),
                price["currency"] + "/百万Token",
                "价目2026-09-16",
                detail,
            )

    fx = fx_row[0] if price["currency"] == "USD" else 1
    _append_metric(
        metrics,
        evidence,
        connection,
        run,
        "计价参数",
        "折人民币系数",
        fx,
        "元/" + price["currency"],
        "汇率2026-09-16",
        detail,
    )
    _append_metric(
        metrics,
        evidence,
        connection,
        run,
        "计价参数",
        "每百万换算",
        1_000_000,
        "Token/百万Token",
        "单位换算",
        {"定义": "百万Token等于1000000 Token"},
    )
    return explanation


def _add_run_metrics(
    run,
    source,
    evidence,
    connection,
    metrics,
):
    if not run["run_id"] or run["轮次"] == 0:
        return

    usage_version = (
        "usage复核"
        if source["用量"]
        else "原生usage"
    )

    token_fields = [
        ("未缓存输入", "未缓存输入Token"),
        ("缓存读", "缓存读取Token"),
        ("缓存写", "缓存写入Token"),
        ("输出Token", "输出Token"),
    ]
    for key, name in token_fields:
        _append_metric(
            metrics,
            evidence,
            connection,
            run,
            "Token用量",
            name,
            run[key],
            "Token",
            usage_version,
        )

    _append_metric(
        metrics,
        evidence,
        connection,
        run,
        "Token用量",
        "推理Token（输出子集）",
        source["运行原文"].get("reasoning_tokens"),
        "Token",
        "原生usage",
    )

    observed_fields = [
        ("外部秒", "外部计时", "秒"),
        ("追加文本次数", "追加指令次数", "次"),
        ("权限点击次数", "权限点击次数", "次"),
        ("人工分钟", "人工用时", "分钟"),
        ("可查动作数", "可查动作数", "次"),
        ("有返回文本动作数", "有返回文本动作数", "次"),
        ("额外启动清单", "额外启动清单", "是=1/否=0"),
    ]
    for key, name, unit in observed_fields:
        _append_metric(
            metrics,
            evidence,
            connection,
            run,
            "运行观测",
            name,
            run[key],
            unit,
            "SQL观测",
        )

    for criterion, score_name in zip(CRITERIA, SCORE_NAMES):
        category = (
            "可靠性分项"
            if criterion.startswith("R")
            else "质量分项"
        )
        _append_metric(
            metrics,
            evidence,
            connection,
            run,
            category,
            criterion + " " + score_name,
            run[criterion],
            "分",
            GRADE_NAMES[run["评分状态码"]],
        )


def _add_assisted_quality(
    run,
    evidence,
    connection,
    metrics,
):
    if run["run_id"] or run["评分状态码"] != 3:
        return

    for criterion, score_name in zip(CRITERIA, SCORE_NAMES):
        if criterion.startswith("R"):
            continue
        _append_metric(
            metrics,
            evidence,
            connection,
            run,
            "质量分项",
            criterion + " " + score_name,
            run[criterion],
            "分",
            GRADE_NAMES[run["评分状态码"]],
        )


def _format_runs(
    raw_runs,
    evidence,
    connection,
    metrics,
):
    catalog, catalog_hash, fx_row, prices = _pricing_context(connection)
    result = []

    for run in raw_runs:
        source = evidence[run["证据ID"]]

        explanation = _add_api_pricing(
            connection,
            run,
            evidence,
            metrics,
            catalog,
            catalog_hash,
            fx_row,
            prices,
        )
        source["本次API计价说明"] = explanation

        result.append({
            "AI": run["AI"],
            "轮次": ROUND_NAMES[run["轮次"]],
            "运行状态": RUN_STATES[run["状态码"]],
            "参与类别": SAMPLE_NAMES[run["样本码"]],
            "运行工具": HOST_NAMES.get(
                run["宿主"],
                run["宿主"],
            ),
            "评分状态": GRADE_NAMES[run["评分状态码"]],
            "API计价说明": explanation,
            "原生用时（秒）": run["原生秒"],
            "模型": run["模型"],
            "参与编号": run["参与ID"],
            "运行编号": run["run_id"],
            "证据编号": run["证据ID"],
        })

        _add_run_metrics(
            run,
            source,
            evidence,
            connection,
            metrics,
        )
        _add_assisted_quality(
            run,
            evidence,
            connection,
            metrics,
        )

    return result


def _action_purpose(action, source):
    purpose = source.get("分析问题或动作目的") or ""
    if action["工具"] == "exec":
        return "批量执行代码"
    if "项目1" in purpose:
        if "Parquet" in purpose:
            return "查看输入材料和原始数据结构"
        return "检索原始数据目录"
    if len(purpose) > 55:
        return purpose[:54] + "…"
    return purpose


def _format_actions(raw_actions, evidence, by_run):
    result = []

    for action in raw_actions:
        source = evidence[action["证据ID"]]
        requested = action["修改请求"] == 1
        has_return = action["返回可见"] == 1

        if requested and has_return:
            confirmation = _yes_no(action["修改确认"])
        else:
            confirmation = None

        if action["操作码"] == 1:
            repeat = _yes_no(action["相邻同参重读"])
        else:
            repeat = None

        parent = by_run[action["run_id"]]

        result.append({
            "AI": action["AI"],
            "轮次": ROUND_NAMES[action["轮次"]],
            "动作序号": action["动作序号"],
            "操作类型": OP_NAMES[action["操作码"]],
            "动作目的": _action_purpose(action, source),
            "返回文本可见": _yes_no(action["返回可见"]),
            "执行成功标记": _yes_no(action["原采集成功"]),
            "返回摘录含错误字样": _yes_no(action["返回含错误标记"]),
            "结构化替换请求": _yes_no(action["修改请求"]),
            "结构化替换确认": confirmation,
            "相邻同参数重读": repeat,
            "动作评审": ACTION_LABEL_NAMES[
                action["历史动作标签码"]
            ],
            "评审版本": {
                0: None,
                1: "第一轮",
                3: "第二轮",
            }[action["标签版本码"]],
            "观测耗时（秒）": action["观测秒"],
            "工具名称": action["工具"],
            "参与编号": parent["参与ID"],
            "运行编号": action["run_id"],
            "行为编号": action["行为ID"],
            "证据编号": action["证据ID"],
        })

    return result


def _check_unit(check, source, item):
    unit_names = {
        1: "条",
        2: "份额（0—1）",
        3: "小时",
        4: "秒",
        6: "月",
        7: source.get("原记录", {}).get(
            "单位",
            "个／次／张",
        ),
        8: "二值标记",
    }
    unit = unit_names.get(check["单位码"])

    if unit or check["记录类型码"] != 1:
        return unit

    if any(word in item for word in ("记录数", "总量", "条数")):
        return "条"
    if "小时" in item:
        return "小时"
    if "%" in item or "百分比" in item:
        return "%"
    return None


def _check_state(check, source):
    kind = check["记录类型码"]

    if kind == 8:
        return source.get("原记录", {}).get(
            "状态",
            "行为片段复核",
        )

    state = source.get("数值或结论状态")
    if state:
        if state.startswith("待修正"):
            return "待修正"
        if "HOLD" in state:
            return "勘误"
        return CHECK_VERDICT_NAMES.get(state, state)

    if kind in (5, 7):
        return "观测数据"
    return "见修正结果"


def _check_parent(raw_runs, by_run, check):
    if check["run_id"]:
        return by_run[check["run_id"]]

    for parent in raw_runs:
        if (
            parent["参与ID"] == "decision:baseline:Claude"
            and check["AI"] == "Claude"
            and check["轮次"] == 1
        ):
            return parent

    raise ValueError("核验记录父键缺失")


def _format_checks(raw_checks, raw_runs, evidence, by_run):
    result = []

    for check in raw_checks:
        source = evidence[check["证据ID"]]
        kind = check["记录类型码"]

        item = (
            source.get("分析问题")
            or source.get("原记录", {}).get("观测项")
            or check["核验项目"]
        )
        if item.startswith("缺失地址"):
            item = "缺失地址聚合"
        elif item.startswith("未来关闭日期"):
            item = "未来关闭日期"
        elif item.startswith("旧独立复算底稿11项"):
            item = "旧复算底稿11项数值"
        elif item.startswith("同址+同类+同日"):
            item = "同址+同类+同日规则簇"

        if kind == 6:
            change = _yes_no(check["判断变化"])
            attempt = _yes_no(check["修正尝试"])
            closed = _yes_no(check["问题闭合"])
        else:
            change = None
            attempt = None
            closed = None

        if kind == 4:
            closed = _yes_no(check["问题闭合"], "待处理")

        parent = _check_parent(
            raw_runs,
            by_run,
            check,
        )

        result.append({
            "AI": check["AI"],
            "轮次": ROUND_NAMES[check["轮次"]],
            "核验类型": CHECK_KIND_NAMES[kind],
            "核验项目": item,
            "核验结论": _check_state(check, source),
            "原数值": check["源观测值"],
            "对照数值": check["参照值"],
            "另一对照值": check["替代基准值"],
            "单位": _check_unit(check, source, item),
            "所列叙述有改动": change,
            "尝试修正所列问题": attempt,
            "所列问题已修好": closed,
            "参与编号": parent["参与ID"],
            "运行编号": check["run_id"],
            "核验编号": check["核验ID"],
            "证据编号": check["证据ID"],
            "问题归属": {
                0: None,
                1: "模型",
                2: "评测方",
                3: "文件身份",
            }[check["问题归属码"]],
        })

    return result


def _sort_analysis_rows(rows):
    rows.sort(
        key=lambda row: (
            str(row.get("运行编号") or "z"),
            row.get("动作序号", 0),
            row["证据编号"],
        )
    )


def _validate_table_links(runs, metrics, actions, checks, by_run):
    participant_ids = {
        row["参与编号"]
        for row in runs
    }

    def link_ok(row):
        if row["参与编号"] not in participant_ids:
            return False
        if row["运行编号"] is None:
            return row["参与编号"] == "decision:baseline:Claude"
        return (
            by_run[row["运行编号"]]["参与ID"]
            == row["参与编号"]
        )

    for rows in (metrics, actions, checks):
        if not all(link_ok(row) for row in rows):
            raise ValueError("分析表参与编号/运行编号连接失败")

    metric_ids = {
        row["指标编号"]
        for row in metrics
    }
    if len(metric_ids) != len(metrics):
        raise ValueError("评分与用量存在重复指标编号")


# 2. 翻译成中文分析表，并补入评分、Token 和价格指标。
def readable_tables(tables, evidence, connection):
    """把内部结构转成分析表。"""
    raw = dict(tables)
    by_run = {
        row["run_id"]: row
        for row in raw["运行记录"]
        if row["run_id"]
    }
    if len(by_run) != sum(
        bool(row["run_id"])
        for row in raw["运行记录"]
    ):
        raise ValueError("运行编号重复")

    metrics = []
    runs = _format_runs(
        raw["运行记录"],
        evidence,
        connection,
        metrics,
    )
    actions = _format_actions(
        raw["行为记录"],
        evidence,
        by_run,
    )
    checks = _format_checks(
        raw["核验记录"],
        raw["运行记录"],
        evidence,
        by_run,
    )

    for rows in (runs, metrics, actions, checks):
        _sort_analysis_rows(rows)

    _validate_table_links(
        runs,
        metrics,
        actions,
        checks,
        by_run,
    )

    return [
        ("运行记录", runs),
        ("评分与用量", metrics),
        ("行为记录", actions),
        ("核验记录", checks),
    ]


# 3. 应用分析范围
def analysis_scope(tables, evidence, connection):
    """应用分析范围。"""
    data = dict(tables)
    kept = []
    excluded = []
    for row in data['运行记录']:
        p = evidence[row['证据编号']]['参与原文']
        # 保留实际 run，排除 calibration
        has_real_run = p['run_id'] is not None and p['condition_name'] != 'calibration'
        mixed_completion = p['record_status'] == 'mixed_excluded'
        if has_real_run or mixed_completion or p["participation_id"] == "decision:baseline:Kimi":
            kept.append(row)
        else:
            excluded.append({'参与编号': p['participation_id'], '原记录状态': p['record_status'], '原因': p['reason']})
    data['运行记录'] = kept
    parents = {r['参与编号']: r for r in kept}
    for name in ('评分与用量', '行为记录', '核验记录'):
        data[name] = [r for r in data[name] if r['参与编号'] in parents]
    evidence['CLEAN:analysis_scope'] = {'清洗规则': '实际run＋混合完成＋Kimi第一轮退出登记；排除calibration', '剔除记录': excluded, '原记录数': len(kept) + len(excluded), '清洗后记录数': len(kept)}
    for action in data['行为记录']:
        action['执行阶段'] = '原单模型运行'
    sources = connection.execute("SELECT s.*,b.content,m.record_json FROM evaluation_source_dataset s JOIN evaluation_raw_blob b USING(sha256) JOIN evaluation_source_record m ON m.source_id=s.source_id AND m.record_no=0 WHERE s.role='mixed_session_native'").fetchall()
    for source in sources:
        meta = json.loads(source['record_json'])
        pid = meta['participation_id']
        parent = parents[pid]
        body = bytes(source['content'])
        assert h(body) == source['sha256']
        records = [json.loads(s) for s in body.decode('utf-8-sig').splitlines() if s.strip()]
        parent['参与类别'] = '混合成果'
        parent['运行工具'] = 'Cursor'
        parent['运行状态'] = '混合交付'
        parent['评分状态'] = '成果质量'
        parent['模型'] = 'Claude主导 / Grok收尾'
        parent['API计价说明'] = 'Token缺失'
        evidence[parent['证据编号']]['混合会话补采'] = {
            'source_id': source['source_id'],
            'SHA256': source['sha256'],
            '原文行数': len(records),
            '范围': '全会话；Token/逐动作时间/工具返回缺失',
        }
        switch_lines = [i for i, r in enumerate(records, 1) if r.get('type') == 'turn_ended' and 'Switched to grok-4.6' in str(r.get('error', ''))]
        assert len(switch_lines) == 1, '切换边界不唯一'
        switch_line = switch_lines[0]
        phase_counts = {'Claude主体': 0, 'Grok收尾': 0}
        count = 0
        switches = 0
        errors = 0
        users = 0
        for line, r in enumerate(records, 1):
            users += int(r.get('role') == 'user')
            errors += int(r.get('type') == 'turn_ended' and r.get('status') == 'error')
            switches += int(r.get('type') == 'turn_ended' and 'Switched to ' in str(r.get('error', '')))
            for block, b in enumerate((r.get('message') or {}).get('content', [])):
                if not isinstance(b, dict) or b.get('type') != 'tool_use':
                    continue
                count += 1
                args = b.get('input')
                rawargs = args
                if isinstance(args, str):
                    try:
                        args = json.loads(args)
                    except (ValueError, TypeError):
                        args = {}
                if not isinstance(args, dict):
                    args = {}
                tool = b.get('name')
                requested = 'old_string' in args and 'new_string' in args
                purpose = str(args.get('description') or args.get('path') or args.get('command') or '参数见证据')
                purpose = purpose if len(purpose) <= 55 else purpose[:54] + '…'
                aid = 'MIX:' + source['source_id'][:12] + ':' + str(line) + ':' + str(block)
                eid = 'ACT:' + aid
                mix_review = connection.execute("SELECT * FROM evaluation_action WHERE version='claude_mixed_action_review_20260918_v1' AND ai='Claude' AND sequence_no=?", (count,)).fetchone()
                assert mix_review is not None
                review_text = {'contributory': '直接贡献交付', 'necessary_exploration': '必要探索', 'administrative': '管理动作', 'waste': '确认无收益'}[mix_review['label']]
                row = {
                    'AI': parent['AI'],
                    '轮次': parent['轮次'],
                    '动作序号': count,
                    '操作类型': {'Read': '读取', 'Shell': '运行命令', 'Write': '写入', 'StrReplace': '修改', 'Glob': '检索', 'Delete': '删除'}.get(tool, '其他'),
                    '动作目的': purpose[:120],
                    '返回文本可见': '否',
                    '执行成功标记': None,
                    '返回摘录含错误字样': None,
                    '结构化替换请求': '是' if requested else '否',
                    '结构化替换确认': None,
                    '相邻同参数重读': None,
                    '动作评审': review_text,
                    '评审版本': '混合会话',
                    '观测耗时（秒）': None,
                    '工具名称': tool,
                    '参与编号': pid,
                    '运行编号': None,
                    '行为编号': aid,
                    '证据编号': eid,
                }
                phase = 'Claude主体' if line < switch_line else 'Grok收尾'
                row['执行阶段'] = phase
                phase_counts[phase] += 1
                data['行为记录'].append(row)
                evidence[eid] = {
                    '原库表': 'evaluation_raw_blob / evaluation_source_record',
                    '原库主键': [source['source_id'], line],
                    'SHA256': source['sha256'],
                    '参与编号': pid,
                    '原文定位': 'L' + str(line) + ':block' + str(block),
                    '原生动作': b,
                    '执行阶段': phase,
                    '切换提示行': switch_line,
                    '动作评审依据': {
                        '原库表': 'evaluation_action',
                        '主键': ['claude_mixed_action_review_20260918_v1', 'Claude', count],
                        '记录': dict(mix_review),
                    },
                }
        evidence[parent['证据编号']]['工作归属复核'] = {
            '切换提示行': switch_line,
            '切换提示原文': records[switch_line - 1],
            '阶段调用数': phase_counts,
            '成果质量分': sum(
                v['数值']
                for v in data['评分与用量']
                if v['参与编号'] == pid and v['指标类别'] == '质量分项'
            ),
            '可靠性分': None,
        }
        for name, value in [('工具调用请求数（混合会话）', count), ('用户任务消息数', users), ('追加任务消息数', max(users - 1, 0)), ('中断记录数', errors), ('模型切换提示数', switches)]:
            mid = h((pid + '|混合会话观测|' + name).encode())[:20]
            eid = 'MET:' + mid
            data['评分与用量'].append({'AI': parent['AI'], '轮次': parent['轮次'], '指标类别': '混合会话观测', '指标': name, '数值': value, '单位': '次', '数据版本': '混合会话2026-09-18', '参与编号': pid, '运行编号': None, '指标编号': mid, '证据编号': eid})
            evidence[eid] = {'来源': '混合会话原文', 'source_id': source['source_id'], 'SHA256': source['sha256'], '参与编号': pid, '指标': name, '数值': value, '计数口径': '全会话 tool_use 计数'}
    assert len(parents) == len(kept)
    for name, rows in data.items():
        assert all((row['参与编号'] in parents for row in rows))
        assert all((row['运行编号'] == parents[row['参与编号']]['运行编号'] for row in rows))
        assert len({row['证据编号'] for row in rows}) == len(rows)
    return [(name, data[name]) for name, _ in tables]


# 4. 统一单位、文字字段和比较口径
def clarify_analysis_fields(tables, evidence, c):
    """补充计费字段与核验展示字段。"""
    data = dict(tables)
    metrics = data['评分与用量']
    lookup = {(r['参与编号'], r['指标']): r['数值'] for r in metrics}

    for row in metrics:
        pid = row['参与编号']
        name = row['指标']
        rate = row['指标类别'] == 'API单价'
        token_name = name[:-3] + 'Token' if rate else None
        row['计费Token'] = lookup.get((pid, token_name)) if rate else None
        row['人民币系数'] = lookup.get((pid, '折人民币系数')) if rate else None
        row['单价基数'] = lookup.get((pid, '每百万换算')) if rate else None
        row['计价档'] = name[-1] if rate else None

        if rate:
            assert all(row[k] is not None for k in ('计费Token', '人民币系数', '单价基数'))
            evidence[row['证据编号']]['同运行计费字段'] = {
                k: row[k] for k in ('计费Token', '人民币系数', '单价基数')
            }

    enhanced = {
        'janaug26': ('2026年1—8月创建记录', '条', '2026年1—8月'),
        'max_created': ('最大创建时间', '时间戳', '全量 created_date'),
        'negative_closed_subset': ('Closed子集负时长', '条', 'status=Closed'),
        'physical_fields': ('物理字段数', '列', '原始 Parquet'),
        'snow_top5': ('雪冰最高5天记录数', '条', '2026年1—8月 Snow or Ice'),
        'standardized_online26': ('固定构成后ONLINE份额', '%（0—100）', '月份×类别×机构固定构成'),
        'arrest_pct': ('明确拘捕占违停比例', '%（0—100）', 'made an arrest / 违停记录'),
        'heat_dup_B': ('供暖duplicate标记记录数', '条', '报告期；duplicate结案标记'),
        'heat_exclusion_A': ('剔除固定Top10地块后供暖记录：基期', '条', '固定Top10地块排除'),
        'heat_exclusion_B': ('剔除固定Top10地块后供暖记录：报告期', '条', '固定Top10地块排除'),
        'hpd_median': ('HPD关闭时长中位数', '小时', '原稿整数小时桶'),
        'hpd_p90': ('HPD关闭时长P90', '小时', '原稿整数小时桶'),
        'oos_n': ('指定OOS类别记录数', '条', 'oos_timeline'),
        'oos_start': ('指定OOS类别首次出现月', '年月', 'oos_timeline'),
        'snow_high_days': ('雪冰高峰日期记录数', '条', '原稿峰日集合'),
        'street_march': ('2026年3月道路状况记录', '条', '2026年3月 Street Condition'),
        'mature_n': ('成熟队列有效时长记录', '条', '原稿成熟队列'),
    }

    def clip(value, limit=55):
        if value in (None, ''):
            return None
        value = str(value)
        return value if len(value) <= limit else value[:limit - 1] + '…'

    for row in data['核验记录']:
        source = evidence[row['证据编号']]
        kind = row['核验类型']
        row['原始文字值'] = None
        row['对照文字值'] = None
        row['数值口径'] = None
        row['判断来源'] = 'SQL核验'

        if kind == '数值核验':
            key = json.loads(source['原库主键'])
            r = dict(c.execute(
                'SELECT * FROM evaluation_numeric_check WHERE version=? AND ai=? AND check_id=?',
                key,
            ).fetchone())
            cid = r['check_id']
            unit = None
            scope = '原核验对象'

            if r['version'] == '1.0':
                if cid.startswith('MEDIAN-'):
                    unit = '小时'
                elif cid.startswith('STANDARD-'):
                    unit = '倍'
                elif cid == 'MOBILE-PCT':
                    unit = '%（0—100）'
                elif cid == 'RATE-10K':
                    unit = '条/万人'
                    scope = '人口分母算术'
                elif cid == 'ADDRESS-GE1000':
                    unit = '个地址字符串'
                elif cid == 'DRUG-MIX':
                    unit = '文字判断'
                    scope = '2025年1—6月 vs 2025年11月—2026年8月'
                elif cid in (
                    'CORE-N', 'HEAT-PEAK', 'NEGATIVE', 'VALID-DURATION',
                    'Y1', 'Y2', 'J1', 'J2', 'PAIR-N', 'PEAK-N',
                    'PEAK-SNOW', 'ZIP-REMAIN',
                ) or cid.startswith(('DELTA-', 'ADDRESS-', 'POTHOLE-', 'STREET-', 'SNOW-W')):
                    unit = '条'

                if cid in ('Y1', 'Y2') or (cid.startswith('DELTA-') and r['ai'] == 'Grok'):
                    scope = '10月至次年8月'
                if cid in ('J1', 'J2') or (cid.startswith('DELTA-') and r['ai'] == 'GPT'):
                    scope = '两年各1—8月'
            elif cid in enhanced:
                row['核验项目'], unit, scope = enhanced[cid]

            row['单位'] = unit
            row['数值口径'] = scope

            if number(r['reported']) is None:
                row['原始文字值'] = clip(r['reported'])
            if number(r['recomputed']) is None:
                row['对照文字值'] = clip(r['recomputed'])

            if (
                r['status'] in ('PASS', 'numeric_match')
                and row['原数值'] is not None
                and row['对照数值'] != row['原数值']
            ):
                row['核验结论'] = '披露精度一致'

            source['字段释义'] = {
                '单位': unit,
                '数值口径': scope,
                '原始核验记录': r,
            }

        elif kind == '定向前后核查':
            row['判断来源'] = '历史回溯'
            assert source.get('编码原库主键') and source.get('编码者')

        elif kind == '窗口敏感性观测':
            row['判断来源'] = '同月敏感性复算'
            row['数值口径'] = '基期11月 / 报告期9月 / 基期同月9月'

        elif kind == '实例观测':
            row['判断来源'] = '作品摘录'

        elif kind == '行为片段复核':
            a = source['原记录']
            row['原始文字值'] = clip('动作' + a['动作范围'])
            row['对照文字值'] = clip(a['结果'])
            row['判断来源'] = '行为片段复核'

    return [(name, data[name]) for name, _ in tables]


# 5. 接入补充复核
def review_extensions(tables, evidence, c):
    """加入辅助成果、文件身份和修正任务核验。"""
    data = dict(tables)
    checks = data['核验记录']

    category_map = {
        '不适用': None,
        '无证据业务解释': '解释缺证据',
        '评测方口径误判': '评测口径',
        '最早文件身份': '文件身份',
    }
    owner_map = {
        '不适用': None,
        '文件身份核对': '文件身份',
    }
    conclusion_map = {
        '原因尚未验证': '原因未验证',
        '行政事件未经核实': '事件未核实',
    }

    codes = {
        r['record_key']: json.loads(r['record_json'])
        for r in c.execute(
            "SELECT * FROM autonomy_delivery_checks "
            "WHERE version='repair_categories_20260918_v1'"
        )
    }

    for row in checks:
        row['问题类别'] = None
        source = evidence[row['证据编号']]
        if row['核验类型'] == '交付后修正任务':
            key = json.loads(source['原库主键'])[-1]
            if key in codes:
                category = codes[key]['问题类别']
                owner = codes[key]['问题归属']
                row['问题类别'] = category_map.get(category, category)
                row['问题归属'] = owner_map.get(owner, owner)
                source['问题分类依据'] = {
                    '原库表': 'autonomy_delivery_checks',
                    '主键': ['repair_categories_20260918_v1', key],
                    '记录': codes[key],
                }

    parent = next(
        r for r in data['运行记录']
        if r['参与编号'] == 'decision:baseline:Claude'
    )
    version = 'claude_content_review_20260918_v1'

    query = (
        "SELECT r.* FROM evaluation_source_record r "
        "JOIN evaluation_source_dataset s USING(source_id) "
        "WHERE s.role='claude_assisted_content_observations' "
        "ORDER BY r.record_no"
    )
    for raw in c.execute(query):
        a = json.loads(raw['record_json'])
        row = {k: None for k in checks[0]}
        key = a['核验编号']
        eid = 'CHK:' + key
        problem = a['核验结论'] in (
            '标题对象错配',
            '范围不一致',
            '证据不足',
            '原因尚未验证',
            '行政事件未经核实',
        )

        conclusion = conclusion_map.get(a['核验结论'], a['核验结论'])
        category = category_map.get(a['问题类别'], a['问题类别']) if problem else None

        row.update({
            'AI': 'Claude',
            '轮次': '第一轮',
            '核验类型': '辅助成品复核',
            '核验项目': a['核验项目'],
            '核验结论': conclusion,
            '原数值': a['原数值'],
            '对照数值': a['对照数值'],
            '单位': a['单位'],
            '所列叙述有改动': None,
            '尝试修正所列问题': None,
            '所列问题已修好': '否' if problem else None,
            '参与编号': parent['参与编号'],
            '运行编号': parent['运行编号'],
            '核验编号': key,
            '证据编号': eid,
            '问题归属': '辅助成品' if problem else None,
            '原始文字值': None,
            '对照文字值': None,
            '数值口径': None,
            '判断来源': '独立复算',
            '问题类别': category,
        })
        checks.append(row)
        evidence[eid] = {
            '原库表': 'evaluation_source_record',
            '原库主键': [raw['source_id'], raw['record_no']],
            '复核版本': version,
            '复核记录': a,
        }

    query = (
        "SELECT * FROM autonomy_delivery_checks "
        "WHERE version='artifact_identity_20260918_v1' "
        "ORDER BY record_key"
    )
    for raw in c.execute(query):
        a = json.loads(raw['record_json'])
        identity_parent = next(
            r for r in data['运行记录']
            if r['运行编号'] == raw['run_id']
        )
        key = 'IDENT:' + raw['record_key']
        eid = 'CHK:' + key
        row = {k: None for k in checks[0]}
        row.update({
            'AI': a['AI'],
            '轮次': a['轮次'],
            '核验类型': '文件身份核验',
            '核验项目': a['核验项目'],
            '核验结论': a['数值或结论状态'],
            '所列问题已修好': '是',
            '参与编号': identity_parent['参与编号'],
            '运行编号': raw['run_id'],
            '核验编号': key,
            '证据编号': eid,
            '问题归属': '文件身份',
            '数值口径': '文件身份',
            '判断来源': '文件身份核验',
            '问题类别': '文件身份',
        })
        checks.append(row)
        evidence[eid] = {
            '原库表': 'autonomy_delivery_checks',
            '原库主键': ['artifact_identity_20260918_v1', raw['record_key']],
            '原记录': a,
        }

    related = {
        'C01': 'CLAUDE:ONEADDR',
        'C02': 'CLAUDE:DELETE',
        'C03': 'CLAUDE:BOUND',
    }
    lookup = {
        r['核验编号']: r
        for r in checks
        if r['核验类型'] == '辅助成品复核'
    }

    for row in checks:
        if row['核验类型'] != '交付后修正任务':
            continue
        source = evidence[row['证据编号']]
        key = json.loads(source['原库主键'])[-1]
        if key in related:
            target = lookup[related[key]]
            row['问题类别'] = target['问题类别']
            row['问题归属'] = '辅助成品'
            source['问题分类依据'] = {
                '对应核验编号': target['核验编号'],
                '原库判定': evidence[target['证据编号']],
            }

    parents = {
        r['参与编号']: r['运行编号']
        for r in data['运行记录']
    }
    for name, rows in data.items():
        assert all(
            r['参与编号'] in parents
            and r['运行编号'] == parents[r['参与编号']]
            for r in rows
        ), name
        assert len({r['证据编号'] for r in rows}) == len(rows), name

    claude_quality = sum(
        r['数值']
        for r in data['评分与用量']
        if r['参与编号'] == parent['参与编号']
        and r['指标类别'] == '质量分项'
    )
    sql_claude_quality = c.execute(
        "SELECT SUM(score) FROM evaluation_score_item "
        "WHERE version='claude_assisted_quality_20260918_v1' "
        "AND ai='Claude'"
    ).fetchone()[0]
    assert claude_quality == sql_claude_quality == 81

    evidence['REVIEW:project_comparison'] = {
        '研究范围': '前序项目校准',
        'Claude辅助质量': 81,
        '复核': '统一规则',
        '新增模型运行': 0,
    }
    return [(name, data[name]) for name, _ in tables]


# 6. 从逐动作二审和人工负担证据派生 AWR 等可比较指标。
def derived_analysis_metrics(tables, evidence, c):
    """从版本化评审与分析表派生指标。"""
    data = dict(tables)
    metrics = data['评分与用量']
    actions = data['行为记录']
    checks = data['核验记录']
    runs = data['运行记录']
    cond = {r['run_id']: json.loads(r['record_json']) for r in c.execute("SELECT * FROM autonomy_conditions WHERE version='autonomy_review_20260917_v1'")}

    def add(parent, category, name, value, unit, version, detail):
        if value is None:
            return
        mid = h((parent['参与编号'] + '|' + category + '|' + name).encode())[:20]
        if any((r['指标编号'] == mid for r in metrics)):
            return
        eid = 'MET:' + mid
        row = {'AI': parent['AI'], '轮次': parent['轮次'], '指标类别': category, '指标': name, '数值': value, '单位': unit, '数据版本': version, '参与编号': parent['参与编号'], '运行编号': parent['运行编号'], '指标编号': mid, '证据编号': eid, '计费Token': None, '人民币系数': None, '单价基数': None, '计价档': None}
        metrics.append(row)
        evidence[eid] = deepcopy(detail)
    bypid = {r['参与编号']: r for r in runs}
    for pid, parent in bypid.items():
        if pid == 'decision:baseline:Kimi':
            # 只有退出登记，没有首交成果或完整行为记录；不能派生零负担。
            continue
        aa = sorted([r for r in actions if r['参与编号'] == pid], key=lambda r: r['动作序号'])
        if aa:
            reviewed = [r for r in aa if r['动作评审'] != '尚未评审']
            sub = [r for r in reviewed if r['动作评审'] != '管理动作']
            waste = [r for r in sub if r['动作评审'] == '确认无收益']
            detail = {'来源': '行为记录逐动作已版本化二审', '参与编号': pid, '标签定义': '管理动作不进入AWR分母；必要探索不因未写入最终报告而算浪费', '动作证据ID': [r['证据编号'] for r in aa]}
            add(parent, '行为二审', '动作二审覆盖率', len(reviewed) / len(aa), '比例（0—1）', '动作二审2026-09-18', detail)
            add(parent, '行为二审', '实质动作数', len(sub), '次', '动作二审2026-09-18', detail)
            add(parent, '行为二审', '确认无收益动作数', len(waste), '次', '动作二审2026-09-18', detail)
            add(parent, '行为二审', 'AWR_count', len(waste) / len(sub) if sub else None, '比例（0—1）', '动作二审2026-09-18', detail)
            longest = cur = 0
            for r in aa:
                cur = cur + 1 if r['动作评审'] == '确认无收益' else 0
                longest = max(longest, cur)
            add(parent, '行为二审', '最长连续无收益动作数', longest, '次', '动作二审2026-09-18', detail)
            timed = sum((r['观测耗时（秒）'] is not None for r in sub))
            add(parent, '行为二审', '实质动作时长覆盖率', timed / len(sub) if sub else None, '比例（0—1）', '动作二审2026-09-18', detail)
            if sub and timed == len(sub):
                total = sum((r['观测耗时（秒）'] or 0 for r in sub))
                bad = sum((r['观测耗时（秒）'] or 0 for r in waste))
                add(parent, '行为二审', 'AWR_time', bad / total if total else 0, '比例（0—1）', '动作二审2026-09-18', detail)
        rr = [r for r in checks if r['参与编号'] == pid and r['核验类型'] == '交付后修正任务']
        burden = [r for r in rr if r['问题归属'] in ('模型', '辅助成品')]
        evaluator = [r for r in rr if r['问题归属'] not in ('模型', '辅助成品')]
        detail = {'来源': '核验记录中的版本化首交修正任务', '说明': '任务级计数；一项可关联多条核验', '任务证据ID': [r['证据编号'] for r in rr]}
        add(parent, '人工负担证据', '首交后需处理任务数', len(burden), '项', '修正任务', detail)
        add(parent, '人工负担证据', '评测方或身份任务数', len(evaluator), '项', '修正任务', detail)
        cv = cond.get(parent['运行编号'])
        if cv:
            detail = {'来源': 'autonomy_conditions/autonomy_review_20260917_v1', '原记录': cv, '公共输入': 'data_manifest'}
            add(parent, '实验条件', '启动指令字符数', cv.get('启动指令字符数'), '字符', '条件复核2026-09-17', detail)
            add(parent, '实验条件', '额外启动方法清单', cv.get('额外启动方法清单'), '是=1/否=0', '条件复核2026-09-17', detail)
            add(parent, '实验条件', '方法Skill哈希可核', 1 if cv.get('冻结方法SkillSHA256') else 0, '是=1/否=0', '条件复核2026-09-17', detail)
            add(parent, '实验条件', '数据清单哈希可核', 1 if cv.get('数据清单SHA256') else 0, '是=1/否=0', '条件复核2026-09-17', detail)
            add(parent, '实验条件', '独立业务字段字典公共输入证据', 0, '是=1/否=0', '条件复核2026-09-18', detail)
    metrics.sort(key=lambda x: (str(x.get('运行编号') or 'z'), x['指标类别'], x['指标'], x['证据编号']))
    assert len({r['指标编号'] for r in metrics}) == len(metrics)
    return [(name, data[name]) for name, _ in tables]


def split_api_pricing(tables):
    """拆分 API 计价。"""
    data = dict(tables)
    metrics = data['评分与用量']
    kept = []
    pricing = []

    token_order = {
        '未缓存输入': 1,
        '缓存读取': 2,
        '缓存写入': 3,
        '输出': 4,
    }

    for row in metrics:
        category = row['指标类别']

        if category == 'API单价':
            tier = row['计价档']
            suffix = '单价' + tier
            token_type = row['指标']
            if token_type.endswith(suffix):
                token_type = token_type[:-len(suffix)]

            cost = (
                row['数值']
                * row['计费Token']
                * row['人民币系数']
                / row['单价基数']
            )

            pricing.append({
                'AI': row['AI'],
                '轮次': row['轮次'],
                'Token类型': token_type,
                '计价档': tier,
                '单价': row['数值'],
                '单价单位': row['单位'],
                '计费Token': row['计费Token'],
                '人民币系数': row['人民币系数'],
                '单价基数': row['单价基数'],
                '分项费用（元）': cost,
                '数据版本': row['数据版本'],
                '参与编号': row['参与编号'],
                '运行编号': row['运行编号'],
                '证据编号': row['证据编号'],
            })
            continue

        if category == '计价参数':
            continue

        clean = {
            key: value
            for key, value in row.items()
            if key not in ('计费Token', '人民币系数', '单价基数', '计价档')
        }
        kept.append(clean)

    pricing.sort(
        key=lambda row: (
            str(row.get('运行编号') or 'z'),
            token_order.get(row['Token类型'], 99),
            0 if row['计价档'] == '低' else 1,
        )
    )

    data['评分与用量'] = kept
    data['API计价'] = pricing

    return [
        ('运行记录', data['运行记录']),
        ('评分与用量', data['评分与用量']),
        ('API计价', data['API计价']),
        ('行为记录', data['行为记录']),
        ('核验记录', data['核验记录']),
    ]


FIELD_MEANING = {'AI': 'AI 系统',
 '轮次': '实验轮次',
 '运行状态': '运行结果',
 '参与类别': '样本角色',
 '运行工具': '宿主 / Agent 工具',
 '评分状态': '评分状态',
 'API计价说明': 'API 目录价口径',
 '原生用时（秒）': '原生运行时长（秒）',
 '模型': '模型名称/版本',
 '参与编号': '分析表连接键',
 '运行编号': 'benchmark_run.run_id',
 '证据编号': '证据查询键',
 '指标类别': '指标类别',
 '指标': '指标名称',
 '数值': '指标值',
 '单位': '数值单位',
 '数据版本': '评分/审计/价格版本',
 '指标编号': '指标唯一键',
 'Token类型': '计费 Token 类型',
 '单价': 'API 单价',
 '单价单位': 'API 单价单位',
 '计费Token': '计费 Token',
 '人民币系数': '人民币换算系数',
 '单价基数': 'API 单价基数',
 '计价档': '价格档位',
 '分项费用（元）': 'API 目录价分项成本',
 '动作序号': 'run 内动作序号',
 '操作类型': '工具动作类别',
 '动作目的': '动作对象/目的',
 '返回文本可见': '工具返回文本可见性',
 '执行成功标记': '工具执行状态标记',
 '返回摘录含错误字样': '返回摘录错误标记',
 '结构化替换请求': 'old/new 替换请求',
 '结构化替换确认': 'old/new 替换结果确认',
 '相邻同参数重读': '相邻同参数重复读取',
 '动作评审': '动作评审标签',
 '评审版本': '动作评审版本',
 '观测耗时（秒）': '动作观测时长（秒）',
 '工具名称': '原生工具名称',
 '行为编号': '动作唯一键',
 '执行阶段': '混合会话执行阶段',
 '核验类型': '核验类型',
 '核验项目': '核验对象',
 '核验结论': '核验结果',
 '原数值': '原值',
 '对照数值': '对照值',
 '另一对照值': '第三参照值',
 '所列叙述有改动': '叙述/路线变化',
 '尝试修正所列问题': '修正尝试',
 '所列问题已修好': '问题闭合状态',
 '核验编号': '核验唯一键',
 '问题归属': '问题归属',
 '原始文字值': '原始文本',
 '对照文字值': '对照文本',
 '数值口径': '数值范围/口径',
 '判断来源': '判断证据来源',
 '问题类别': '问题类型',
 '原生分钟': '原生用时÷60',
 '成果质量分': 'A-E质量分项合计＋质量封顶调整',
 '独立可靠性分': 'R1-R5 可靠性分项合计',
 '总Token': '输入+缓存读+缓存写+输出',
 'API费用低（元）': '低档 API 目录价等价成本',
 'API费用高（元）': '高档 API 目录价等价成本',

 '数值差（对照减原值）': '对照值−原值'}

FIELD_NULL = {'运行编号': '可空：无独立 run',
 '模型': '可空：无独立模型运行',
 '原生用时（秒）': '可空：无起止时间',
 '计费Token': '必填',
 '人民币系数': '必填',
 '单价基数': '必填',
 '计价档': '必填',
 '观测耗时（秒）': '可空：无动作时长',
 '执行阶段': '可空：非混合会话',
 '原数值': '可空：无原值',
 '对照数值': '可空：无对照值',
 '另一对照值': '可空：非三端/敏感性比较',
 '单位': '可空：纯文字核验',
 '原始文字值': '可空：纯数值核验',
 '对照文字值': '可空：无文字对照',
 '结构化替换确认': '可空：不适用/未核实',
 '相邻同参数重读': '可空：无可比前项',
 '原生分钟': '可空：无原生用时',
 '成果质量分': '可空：质量分项不完整',
 '独立可靠性分': '可空：可靠性分项不完整',
 '总Token': '可空：Token 项不完整',
 'API费用低（元）': '可空：Token/价格不完整',
 'API费用高（元）': '可空：Token/价格不完整',

 '数值差（对照减原值）': '可空：原值/对照值缺失'}

CATEGORICAL_FIELDS={'AI','轮次','运行状态','参与类别','运行工具','评分状态','指标类别','单位','数据版本','Token类型','单价单位','计价档','操作类型','返回文本可见','执行成功标记','返回摘录含错误字样','结构化替换请求','结构化替换确认','相邻同参数重读','动作评审','评审版本','执行阶段','核验类型','核验结论','所列叙述有改动','尝试修正所列问题','所列问题已修好','问题归属','判断来源','问题类别'}


# 7. 生成字段字典，集中说明含义、允许值和空值规则。
def build_field_dictionary_rows(tables):
    extra = {
        '运行记录': [
            '原生分钟', '成果质量分', '独立可靠性分', '总Token',
            'API费用低（元）', 'API费用高（元）',
        ],
        '核验记录': ['数值差（对照减原值）'],
    }

    out = []
    for table, rows in tables:
        fields = list(rows[0]) + extra.get(table, [])
        for field in fields:
            if field in extra.get(table, []):
                vals = []
            else:
                vals = [
                    row.get(field)
                    for row in rows
                    if row.get(field) not in (None, '')
                ]

            if field.endswith('编号') or field in ('参与编号', '运行编号', '证据编号'):
                allowed = '唯一标识/关联键'
            elif field in CATEGORICAL_FIELDS:
                uniq = []
                for value in vals:
                    value = str(value)
                    if value not in uniq:
                        uniq.append(value)
                joined = '；'.join(uniq)
                allowed = (
                    joined
                    if uniq and (field in ('参与类别','评分状态','指标类别') or (len(uniq) <= 8 and len(joined) <= 80))
                    else f'{len(uniq)} 类'
                )
            elif field in (
                '数值', '原数值', '对照数值', '另一对照值',
                '原生用时（秒）', '观测耗时（秒）', '计费Token',
                '人民币系数', '单价基数', '原生分钟', '成果质量分',
                '独立可靠性分', '总Token', 'API费用低（元）',
                'API费用高（元）', '单价', '分项费用（元）',
                '数值差（对照减原值）',
            ):
                allowed = '数值'
            else:
                allowed = '自由文本'

            out.append({
                '表名': table,
                '字段': field,
                '含义': FIELD_MEANING.get(field, field),
                '允许值': allowed,
                '空值规则': FIELD_NULL.get(field, '必填'),
            })

    return out




# 8. 生成查证页和外置证据数据
def write_evidence_page(evidence):
    target = ROOT / '交付成果/查证.html'
    target.parent.mkdir(exist_ok=True)
    payload = ROOT / '交付成果/查证数据.js'
    data = json.dumps(evidence, ensure_ascii=False, separators=(',', ':')).replace('</script>', '<\\/script>')
    payload.write_text('window.PROJECT2_EVIDENCE=' + data + ';', encoding='utf-8')
    page = '<!doctype html>\n<html lang="zh-CN">\n<meta charset="UTF-8">\n<meta name="viewport" content="width=device-width,initial-scale=1">\n<title>项目2 · 证据查证</title>\n<style>\nbody{max-width:1100px;margin:36px auto;font:16px/1.7 system-ui;padding:0 20px;color:#233044}\na{color:#1267af}input{width:min(760px,72%);padding:10px}button{padding:10px 16px}\ntable{width:100%;border-collapse:collapse}td{padding:8px;border-bottom:1px solid #dde3ea;vertical-align:top}\ntd:first-child{width:190px;font-weight:600}.panel{background:#f7f9fb;border:1px solid #dde5ec;padding:16px;border-radius:8px}\n</style>\n<p><a href="开始阅读.html">← 返回首页</a></p>\n<h1>证据查证</h1>\n<p>输入 Excel 中的证据编号。</p>\n<div class="panel"><input id="key" placeholder="例如 CHK:CLAUDE:ONEADDR"><button id="go">查证</button><p id="status"></p><div id="result"></div></div>\n<script src="查证数据.js"></script>\n<script>\nconst data=window.PROJECT2_EVIDENCE||{};\nfunction render(value,node){\n if(value&&typeof value===\'object\'){\n  let table=document.createElement(\'table\');\n  for(let [k,v] of Object.entries(value)){\n   let row=table.insertRow(),a=row.insertCell(),b=row.insertCell();\n   a.textContent=k;\n   if(v&&typeof v===\'object\'){\n    let d=document.createElement(\'details\'),s=document.createElement(\'summary\');\n    s.textContent=\'展开\';d.append(s);render(v,d);b.append(d);\n   }else b.textContent=v===null?\'—\':String(v);\n  }\n  node.append(table);\n }else node.textContent=value===null?\'—\':String(value);\n}\nfunction show(){\n let k=document.getElementById(\'key\').value.trim(),value=data[k];\n document.getElementById(\'status\').textContent=value?k:\'未找到\';\n let node=document.getElementById(\'result\');node.replaceChildren();\n if(value)render(value,node);\n}\ndocument.getElementById(\'go\').onclick=show;\nfunction fromhash(){\n if(location.hash){\n  document.getElementById(\'key\').value=decodeURIComponent(location.hash.slice(1));\n  show();\n }\n}\nwindow.onhashchange=fromhash;fromhash();\n</script>\n</html>'
    target.write_text(page, encoding='utf-8')
    return (target, payload)


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        '--database',
        type=Path,
        default=ROOT / '_实验系统/telemetry/benchmark.sqlite',
    )
    parser.add_argument('--evidence')
    args = parser.parse_args()

    before = h(args.database.read_bytes())
    tables, evidence, _, queries = extract(args.database)
    assert before == h(args.database.read_bytes()), '原数据库发生变化'

    if args.evidence:
        if args.evidence not in evidence:
            raise ValueError('不存在此证据ID')
        print(json.dumps(evidence[args.evidence], ensure_ascii=False, indent=2))
        return

    write_evidence_page(evidence)
    field_rows = build_field_dictionary_rows(tables)
    manifest = {
        'tables': {name: len(rows) for name, rows in tables},
        'field_dictionary_rows': len(field_rows),
    }

    (ROOT / '_实验系统/当前导出.json').write_text(
        json.dumps(manifest, ensure_ascii=False, indent=2),
        encoding='utf-8',
    )
    (ROOT / '_实验系统/sql/分析取数.sql').write_text(
        '\n'.join(queries),
        encoding='utf-8',
    )

    print(json.dumps({
        **manifest,
        'evidence_shell_bytes': (ROOT / '交付成果/查证.html').stat().st_size,
        'evidence_payload_bytes': (ROOT / '交付成果/查证数据.js').stat().st_size,
        'sql_database_sha256': before,
        'sql_unchanged': True,
        'no_model_calls': True,
    }, ensure_ascii=False))


if __name__ == '__main__':
    main()

