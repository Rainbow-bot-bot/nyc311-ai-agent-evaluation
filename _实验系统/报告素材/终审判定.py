"""Fail closed on every recorded critical invariant."""
def failures(result):
    failed = []
    if result.get('SQLite完整性') != 'ok': failed.append('SQLite完整性')
    if result.get('数据库未改') is not True: failed.append('数据库发生变化')
    if result.get('原生图表数') != 8: failed.append('原生图表数量')
    for name in ('图表来源.json', '简洁版图表来源.json'):
        item = result.get(name, {})
        for key in ('图片一致', '素材未变'):
            if item.get(key) is not True: failed.append(name + ':' + key)
        if not item.get('图片数') or item.get('图片数') != item.get('来源数'):
            failed.append(name + ':图片与来源数量')
    for name, item in result.get('PDF', {}).items():
        if not item.get('页数') or item.get('乱码替代符', 0): failed.append(name + ':PDF内容')
    if len(result.get('PDF', {})) != 2: failed.append('PDF交付数量')
    if result.get('断链'): failed.append('交付入口断链')
    if result.get('Excel错误单元格'): failed.append('Excel非预期错误')
    return failed
