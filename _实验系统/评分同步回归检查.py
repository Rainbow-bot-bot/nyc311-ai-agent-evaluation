"""用正式325条评分的内存/临时副本检查同步防错，不写正式交付物。"""
from pathlib import Path
from copy import deepcopy
import hashlib
import importlib.util
import json
import shutil
import tempfile


def main(root=None):
    root = Path(root) if root else Path(__file__).resolve().parents[1]
    spec = importlib.util.spec_from_file_location('score_reference_check', root / '评测数据处理.py')
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    scores, evidence = module.read_current_scores(), module.load_evidence()
    module.verify_current_scores(evidence, scores)
    external = {v['关联问题编号'] for v in scores.values()} - scores.keys()
    key = next(k for k, v in scores.items() if v['关联问题编号'] in external)
    issue = scores[key]['关联问题编号']
    checks = []

    def rejected(name, action):
        try:
            action()
        except ValueError:
            checks.append(name)
        else:
            raise AssertionError(name + ' 未拦截')

    changed = deepcopy(scores)
    missing = '__不存在的测试问题__'
    assert missing not in evidence
    changed[key]['关联问题编号'] = missing
    before = deepcopy(evidence)
    rejected('不存在的编号', lambda: module.merge_current_scores(evidence, changed))
    assert evidence == before
    for record in ({}, {'现行S关联评分': [changed[key]['评分原文']]},
                   {'现行S关联评分': [], '来源': '只有来源标签'},
                   {'问题记录': {'核验发现': '缺少出处'}},
                   {'问题记录': {'来源': '缺少事实'}},
                   {'问题记录': {'核验发现': '  ', '来源': None}}):
        bad = deepcopy(evidence)
        bad[missing] = record
        bad[key] = changed[key]
        rejected('空证据合并', lambda: module.merge_current_scores(bad, changed))
        rejected('空证据核对', lambda: module.verify_current_scores(bad, changed))
    for field, value in [('得分', -1), ('依据', '测试改写依据')]:
        bad = deepcopy(evidence)
        bad[key]['评分原文'][field] = value
        rejected(field + ' 不一致', lambda: module.verify_current_scores(bad, scores))
    merged = module.merge_current_scores(evidence, scores)
    module.verify_current_scores(merged, scores)
    assert merged == evidence, '正常同步应幂等'
    refreshed = {}
    module.retain_current_assessment(refreshed)
    module.verify_current_scores(module.merge_current_scores(refreshed, scores), scores)
    assert all(refreshed[i] == evidence[i] for i in external)
    registered = deepcopy(evidence)
    registered[missing] = {'问题记录': {'核验发现': '仅用于验证新问题登记流程', '来源': '隔离测试记录'}}
    module.verify_current_scores(module.merge_current_scores(registered, changed), changed)

    # 实际走Excel读取和文件写出入口；错误编号同时改K列与批注。
    with tempfile.TemporaryDirectory(prefix='project2_score_refs_') as tmp:
        test_root = Path(tmp)
        for rel in ['资料/评判标准.md', '交付成果/项目2_AI评测分析.xlsx', '交付成果/查证数据.js']:
            dst = test_root / rel
            dst.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(root / rel, dst)
        module.ROOT = test_root
        module.sync_current_scores()
        from openpyxl import load_workbook
        xlsx = test_root / '交付成果/项目2_AI评测分析.xlsx'
        book = load_workbook(xlsx)
        for row in book['评分与用量']:
            if row[2].value == 'S分项' and row[10].value == issue:
                row[10].value = missing
                row[4].comment.text = row[4].comment.text.replace('问题编号：' + issue, '问题编号：' + missing)
                break
        else:
            raise AssertionError('未找到测试评分')
        book.save(xlsx)
        book.close()
        payload = test_root / '交付成果/查证数据.js'
        digest = hashlib.sha256(payload.read_bytes()).hexdigest()
        rejected('Excel错误编号同步', module.sync_current_scores)
        assert hashlib.sha256(payload.read_bytes()).hexdigest() == digest
        assert not payload.with_suffix('.js.tmp').exists()
    result = {'实际评分条数': len(scores), '外部问题数': len(external),
              '拒绝测试次数': len(checks), '错误同步未写文件': True,
              '正常同步与证据刷新': '通过', '重新评定分数': False}
    print(json.dumps(result, ensure_ascii=False))
    return result


if __name__ == '__main__':
    main()
