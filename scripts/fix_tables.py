#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
修復 HTML 文件中表格 thead 連續重複 <th> 的問題（docx 轉換合併單元格遺留）。
偵測 thead 中連續重複的 th（相同文字），記錄需刪除的列索引，
並在對應 tbody 的每行 tr 中按相同索引刪除 td。
保留原始縮進與格式；Example 類型表格（單一 th）自動跳過。
"""
import re
import os

BASE = "/Users/susi/Documents/Project/ftp.openapi.trae"

FILES = [
    "docs/payments/qr-online/alipay-wap.html",
    "docs/payments/qr-online/wechat-web.html",
    "docs/payments/qr-online/wechat-wap.html",
    "docs/payments/qr-online/alipay-web.html",
    "docs/payments/qr-online/wechat-app.html",
    "docs/payments/qr-online/qr-code.html",
    "docs/payments/qr-online/alipay-plus.html",
    "docs/payments/qr-online/wechat-miniprogram.html",
    "docs/payments/qr-online/alipay-app.html",
    "docs/payments/card-offline/global-payments-offline.html",
    "docs/payments/other/ftp-foreign.html",
    "docs/payments/qr-offline/barcode.html",
    "docs/payments/qr-offline/cash-api.html",
    "docs/payments/card-online/global-payments-online.html",
    "docs/payments/card-online/fintech-payment-online.html",
]

TH_RE = r'<th[^>]*>.*?</th>'
TD_RE = r'<td[^>]*>.*?</td>'


def split_elements(inner, tag_re):
    """將區段內容依指定標籤切分，回傳交替 list [text, elem, text, elem, ...]。"""
    return re.split(r'(' + tag_re + ')', inner, flags=re.DOTALL)


def rebuild(parts, remove_indices):
    """依 remove_indices 移除元素（奇數索引位置），並移除其後方分隔文字的「前導空白」
    （僅空白/換行），藉此保留原始縮進格式與結構標籤（如 </tr>）。
    注意：不可整段刪除後方分隔文字，否則會誤刪 </tr> 等結構標籤。"""
    new_parts = []
    elem_idx = -1
    strip_next_leading_ws = False
    for i, part in enumerate(parts):
        if i % 2 == 1:  # 元素 (th / td)
            elem_idx += 1
            if elem_idx in remove_indices:
                strip_next_leading_ws = True
                continue
            new_parts.append(part)
            strip_next_leading_ws = False
        else:  # 分隔文字
            if strip_next_leading_ws:
                strip_next_leading_ws = False
                part = re.sub(r'^\s+', '', part)
            new_parts.append(part)
    return ''.join(new_parts)


def text_of(elem):
    return re.sub(r'<[^>]+>', '', elem).strip()


def process_table(table_html):
    """處理單一表格，回傳 (新表格 html, 資訊 dict)。"""
    info = {'changed': False, 'remove_indices': set(), 'th_count': 0,
            'rows_total': 0, 'rows_mismatch': 0, 'headers': []}
    thead_m = re.search(r'(<thead[^>]*>)(.*?)(</thead>)', table_html, flags=re.DOTALL)
    tbody_m = re.search(r'(<tbody[^>]*>)(.*?)(</tbody>)', table_html, flags=re.DOTALL)
    if not thead_m or not tbody_m:
        return table_html, info

    thead_open, thead_inner, thead_close = thead_m.group(1), thead_m.group(2), thead_m.group(3)
    th_parts = split_elements(thead_inner, TH_RE)
    th_elems = [th_parts[i] for i in range(1, len(th_parts), 2)]
    info['th_count'] = len(th_elems)
    if len(th_elems) <= 1:
        return table_html, info  # Example 表格或空表，不處理

    th_texts = [text_of(e) for e in th_elems]
    info['headers'] = th_texts

    # 偵測「連續重複」的 th（相同文字、非空），記錄需刪除的索引（保留第一個）
    remove_indices = set()
    for i in range(1, len(th_texts)):
        if th_texts[i] and th_texts[i] == th_texts[i - 1]:
            remove_indices.add(i)
    info['remove_indices'] = remove_indices
    if not remove_indices:
        return table_html, info

    # 重建 thead
    new_thead_inner = rebuild(th_parts, remove_indices)

    # 處理 tbody 中每個 tr，按相同索引刪除 td
    tbody_open, tbody_inner, tbody_close = tbody_m.group(1), tbody_m.group(2), tbody_m.group(3)

    def process_tr(m):
        info['rows_total'] += 1
        tr_open, tr_inner, tr_close = m.group(1), m.group(2), m.group(3)
        td_parts = split_elements(tr_inner, TD_RE)
        td_elems = [td_parts[i] for i in range(1, len(td_parts), 2)]
        if len(td_elems) != len(th_elems):
            info['rows_mismatch'] += 1
        new_tr_inner = rebuild(td_parts, remove_indices)
        return tr_open + new_tr_inner + tr_close

    new_tbody_inner = re.sub(r'(<tr[^>]*>)(.*?)(</tr>)', process_tr, tbody_inner, flags=re.DOTALL)

    info['changed'] = True
    new_table = (table_html[:thead_m.start()]
                 + thead_open + new_thead_inner + thead_close
                 + table_html[thead_m.end():tbody_m.start()]
                 + tbody_open + new_tbody_inner + tbody_close
                 + table_html[tbody_m.end():])
    return new_table, info


def process_file(rel_path):
    abs_path = os.path.join(BASE, rel_path)
    with open(abs_path, 'r', encoding='utf-8') as f:
        content = f.read()
    original = content

    table_count = 0
    fixed_count = 0
    warnings = []

    def repl_table(m):
        nonlocal table_count, fixed_count
        table_count += 1
        new_html, info = process_table(m.group(0))
        if info['changed']:
            fixed_count += 1
            print("    表格#{}: th數={} 移除索引={} 列數={} 不符列數={}".format(
                table_count, info['th_count'], sorted(info['remove_indices']),
                info['rows_total'], info['rows_mismatch']))
            print("      原表頭: {}".format(info['headers']))
            if info['rows_mismatch'] > 0:
                warnings.append("{} 表格#{} 有 {} 列 td 數與 th 數不符（已按索引處理，請複查）".format(
                    rel_path, table_count, info['rows_mismatch']))
        return new_html

    new_content = re.sub(r'<table[^>]*>.*?</table>', repl_table, content, flags=re.DOTALL)

    if new_content != original:
        with open(abs_path, 'w', encoding='utf-8') as f:
            f.write(new_content)
        print("  -> [已修改] 共 {} 個表格, 修復 {} 個".format(table_count, fixed_count))
    else:
        print("  -> [未變更] 共 {} 個表格 (無連續重複 th)".format(table_count))

    return warnings


def main():
    print("開始修復表格連續重複 th 問題\n")
    all_warnings = []
    total_fixed = 0
    for rel in FILES:
        abs_path = os.path.join(BASE, rel)
        if not os.path.exists(abs_path):
            print("[找不到] {}".format(rel))
            continue
        print("處理: {}".format(rel))
        all_warnings.extend(process_file(rel))

    print("\n" + "=" * 60)
    if all_warnings:
        print("警告總覽:")
        for w in all_warnings:
            print("  ! {}".format(w))
    else:
        print("完成，無警告。")


if __name__ == '__main__':
    main()
