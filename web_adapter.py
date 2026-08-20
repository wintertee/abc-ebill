#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Pyodide 浏览器端入口：接收 .eml 字节流，返回核对报告 + xlsx 文件(base64)。

在浏览器里由 index.html 加载，复用 main.py 的全部核心逻辑；
处理完全在浏览器本地完成，账单数据不上传任何服务器。
"""
import base64
import io
import zipfile

import main as core


def _b64(b):
    return base64.b64encode(b).decode('ascii')


def process_eml(raw, prefix=''):
    """处理一份 .eml 字节流。返回 dict：
      ok=True  -> {'report': str, 'rows': [..按卡汇总行..],
                   'files': {文件名: base64}, 'zip': 打包好的 zip(base64)}
      ok=False -> {'error': str}
    """
    raw = bytes(raw)  # Pyodide 里 JS 的 Uint8Array 传进来是 memoryview，需转成 bytes
    try:
        info, body = core.parse_eml_bytes(raw)
        account, summary, txns = core.extract_summary_and_txns(
            core.TableExtractor.from_html(body).tables)
        if not summary:
            raise ValueError('未能解析到账务说明部分，无法核对。')
        sums, counts = core.verify(summary, txns)
        cards = core.group_by_card(txns)
        due_date = account.get('到期还款日', '')
        card_rows, card_summary, totals = core.build_card_summary(cards, due_date)
        core.verify_card_totals(sums, totals)
    except ValueError as e:
        return {'ok': False, 'error': str(e)}

    report = core.build_report_text(None, {}, txns, counts, sums, card_rows, totals, summary)

    raw_files = {}
    raw_files['summary.xlsx'] = core.xlsx_bytes(
        '账务说明', ['项目', '内容'], core.build_summary_rows(info, summary)[1:])
    txn_header, txn_rows = core.transaction_columns(txns)
    raw_files['transactions.xlsx'] = core.xlsx_bytes('交易明细', txn_header, txn_rows)
    raw_files['cards.xlsx'] = core.xlsx_bytes('按卡汇总', core.CARD_HEADER, card_rows)
    for c in sorted(cards):
        h, r = core.transaction_columns(cards[c])
        raw_files[f'card_{c}.xlsx'] = core.xlsx_bytes(
            '交易明细', h, r, head_block=(core.CARD_HEADER, [card_summary[c]]))

    # 打包成单个 zip，前端只留一个下载按钮
    zbuf = io.BytesIO()
    with zipfile.ZipFile(zbuf, 'w', zipfile.ZIP_DEFLATED) as z:
        for name, data in raw_files.items():
            z.writestr(f'{prefix}_{name}' if prefix else name, data)

    return {'ok': True, 'report': report, 'rows': card_rows,
            'files': {n: _b64(b) for n, b in raw_files.items()},
            'zip': _b64(zbuf.getvalue())}


if __name__ == '__main__':
    # 本地测试：uv run python web_adapter.py bills/2608.eml
    import sys
    result = process_eml(open(sys.argv[1], 'rb').read())
    if result['ok']:
        print(result['report'])
        print('生成文件:', ', '.join(result['files']))
    else:
        print('错误:', result['error'])
        sys.exit(1)
