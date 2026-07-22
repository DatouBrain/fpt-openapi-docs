#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""檢查 FTP OpenAPI 文檔站點的連結完整性。

掃描 docs/ 下所有 HTML 文件，檢查：
  1. 內部 <a href> 連結指向的文件是否存在
  2. 帶錨點的連結（page.html#section）的錨點 id 是否存在
  3. 每個頁面是否有語言切換器並指向對應的相反語言版本
  4. <link rel=stylesheet> / <script src> / <img src> 引用的資源是否存在
  5. Sidebar 連結是否指向存在的頁面，active class 是否正確
  6. 右側「本頁目錄」的錨點是否對應正文中的實際 id

輸出按嚴重程度分類：BROKEN / MISSING_ANCHOR / WARNING。
"""

import os
import sys
from html.parser import HTMLParser
from collections import defaultdict


DOCS_ROOT = "/Users/susi/Documents/Project/ftp.openapi.trae/docs"
EN_ROOT = os.path.join(DOCS_ROOT, "en")

# 容器上下文標記
CTX_NONE = "none"
CTX_SIDEBAR = "sidebar"
CTX_TOC = "toc-nav"
CTX_LANG_SWITCHER = "lang-switcher"
CTX_HEADER_NAV = "header-nav"


class HtmlAnalyzer(HTMLParser):
    """解析單個 HTML 文件，收集連結、id、結構資訊。"""

    def __init__(self, file_path):
        super().__init__(convert_charrefs=True)
        self.file_path = file_path
        self.ids = set()
        self.links = []  # [{line, href, class, context}]
        self.has_lang_switcher = False
        self.lang_switcher_href = None
        self.lang_switcher_line = None
        self.sidebar_links = []  # [(line, href, class)]
        self.toc_anchors = []  # [(line, anchor)]
        self._context_stack = []

    def _current_context(self):
        return self._context_stack[-1] if self._context_stack else CTX_NONE

    def handle_starttag(self, tag, attrs):
        attrs_d = dict(attrs)
        line = self.getpos()[0]
        cls = attrs_d.get("class", "")
        cls_tokens = cls.split() if cls else []

        # 容器上下文追蹤
        if tag == "aside" and "sidebar" in cls_tokens:
            self._context_stack.append(CTX_SIDEBAR)
        elif tag == "nav" and "toc-nav" in cls_tokens:
            self._context_stack.append(CTX_TOC)
        elif tag == "div" and "lang-switcher" in cls_tokens:
            self._context_stack.append(CTX_LANG_SWITCHER)
            self.has_lang_switcher = True
        elif tag == "nav" and "header-nav" in cls_tokens:
            self._context_stack.append(CTX_HEADER_NAV)

        # 收集 id 屬性
        if "id" in attrs_d and attrs_d["id"]:
            self.ids.add(attrs_d["id"])

        # 收集 <a href>
        if tag == "a" and "href" in attrs_d:
            href = attrs_d["href"]
            ctx = self._current_context()
            entry = {
                "line": line,
                "href": href,
                "class": cls,
                "context": ctx,
            }
            self.links.append(entry)
            if ctx == CTX_SIDEBAR and "sub-link" in cls_tokens:
                self.sidebar_links.append((line, href, cls))
            if ctx == CTX_TOC and href.startswith("#") and len(href) > 1:
                self.toc_anchors.append((line, href[1:]))

        # 語言切換器中的第一個 <a>
        if tag == "a" and CTX_LANG_SWITCHER in self._context_stack:
            if self.lang_switcher_href is None:
                self.lang_switcher_href = attrs_d.get("href")
                self.lang_switcher_line = line

        # 資源引用
        if tag == "link" and attrs_d.get("rel") == "stylesheet" and "href" in attrs_d:
            self.links.append({
                "line": line,
                "href": attrs_d["href"],
                "class": "",
                "context": "stylesheet",
            })
        if tag == "script" and "src" in attrs_d:
            self.links.append({
                "line": line,
                "href": attrs_d["src"],
                "class": "",
                "context": "script",
            })
        if tag == "img" and "src" in attrs_d:
            self.links.append({
                "line": line,
                "href": attrs_d["src"],
                "class": "",
                "context": "img",
            })

    def handle_endtag(self, tag):
        if not self._context_stack:
            return
        top = self._context_stack[-1]
        if (top == CTX_SIDEBAR and tag == "aside") or \
           (top == CTX_TOC and tag == "nav") or \
           (top == CTX_LANG_SWITCHER and tag == "div") or \
           (top == CTX_HEADER_NAV and tag == "nav"):
            self._context_stack.pop()


def is_external(href):
    """判斷是否為外部連結（不需檢查本地文件）。"""
    if not href:
        return True
    lower = href.lower().lstrip()
    if lower.startswith(("http://", "https://", "mailto:", "tel:", "ftp://",
                         "data:", "javascript:", "//")):
        return True
    return False


def resolve_target(current_file, href):
    """將相對 href 解析為絕對路徑。

    Returns:
        (target_path, anchor) - target_path 為 None 表示外部連結
                                anchor 為 "" 表示無錨點
    """
    if is_external(href):
        return None, None
    if "#" in href:
        path_part, anchor = href.split("#", 1)
    else:
        path_part, anchor = href, ""
    if path_part == "":
        # 純錨點（#xxx）— 目標為當前文件
        return current_file, anchor
    target = os.path.normpath(
        os.path.join(os.path.dirname(current_file), path_part)
    )
    return target, anchor


_id_cache = {}


def get_ids_in_file(file_path):
    """取得指定 HTML 文件中所有 id 屬性（帶快取）。"""
    if file_path in _id_cache:
        return _id_cache[file_path]
    ids = set()
    if os.path.isfile(file_path):
        try:
            with open(file_path, "r", encoding="utf-8") as f:
                content = f.read()
            analyzer = HtmlAnalyzer(file_path)
            analyzer.feed(content)
            analyzer.close()
            ids = analyzer.ids
        except Exception:
            ids = set()
    _id_cache[file_path] = ids
    return ids


def is_under_en(file_path):
    """判斷文件是否位於 docs/en/ 目錄下。"""
    return file_path.startswith(EN_ROOT + os.sep)


def check_lang_switcher_target(current_file, href):
    """驗證語言切換器指向相反語言版本。"""
    target, _ = resolve_target(current_file, href)
    if target is None:
        return False, "外部連結或無法解析"
    if not os.path.isfile(target):
        return False, "目標文件不存在: " + os.path.relpath(target, DOCS_ROOT)
    cur_en = is_under_en(current_file)
    tgt_en = is_under_en(target)
    if cur_en == tgt_en:
        return False, "語言切換器指向相同語言目錄: " + os.path.relpath(target, DOCS_ROOT)
    if os.path.basename(target) != os.path.basename(current_file):
        return False, "目標文件名不匹配: " + os.path.basename(target)
    return True, ""


def main():
    # 走訪所有 HTML 文件
    html_files = []
    for dirpath, _, filenames in os.walk(DOCS_ROOT):
        for fn in filenames:
            if fn.endswith(".html"):
                html_files.append(os.path.join(dirpath, fn))
    html_files.sort()

    issues = []  # [(severity, rel_path, line, href, message)]
    total_links = 0
    passed_links = 0
    files_checked = 0

    for fp in html_files:
        files_checked += 1
        try:
            with open(fp, "r", encoding="utf-8") as f:
                content = f.read()
        except Exception as e:
            issues.append(("WARNING", os.path.relpath(fp, DOCS_ROOT),
                           0, "", f"無法讀取文件: {e}"))
            continue

        analyzer = HtmlAnalyzer(fp)
        try:
            analyzer.feed(content)
            analyzer.close()
        except Exception as e:
            issues.append(("WARNING", os.path.relpath(fp, DOCS_ROOT),
                           0, "", f"HTML 解析錯誤: {e}"))
            continue

        # 預先填入 id 快取，避免後續重複解析同一文件
        _id_cache[fp] = analyzer.ids

        rel = os.path.relpath(fp, DOCS_ROOT)
        is_index = os.path.basename(fp) == "index.html"

        # --- 1. 語言切換器檢查 ---
        if not analyzer.has_lang_switcher:
            issues.append(("WARNING", rel, 0, "",
                           "缺少語言切換器 (<div class='lang-switcher'>)"))
        elif analyzer.lang_switcher_href:
            ok, msg = check_lang_switcher_target(
                fp, analyzer.lang_switcher_href)
            if not ok:
                issues.append(("WARNING", rel, analyzer.lang_switcher_line,
                               analyzer.lang_switcher_href,
                               "語言切換器問題: " + msg))

        # --- 2. Sidebar active class 檢查 ---
        current_norm = os.path.normpath(fp)
        active_match_found = False
        for (line, href, cls) in analyzer.sidebar_links:
            target, _ = resolve_target(fp, href)
            if target is None:
                continue
            if os.path.normpath(target) == current_norm:
                active_match_found = True
                if "active" not in cls.split():
                    issues.append(("WARNING", rel, line, href,
                                   "Sidebar 連結指向當前頁面但缺少 'active' class"))
                break
        if not active_match_found and not is_index:
            issues.append(("WARNING", rel, 0, "",
                           "當前頁面未出現在 sidebar 連結中（無 active class 可設置）"))

        # --- 3. TOC 錨點檢查 ---
        for (line, anchor) in analyzer.toc_anchors:
            total_links += 1
            if anchor not in analyzer.ids:
                issues.append(("MISSING_ANCHOR", rel, line, "#" + anchor,
                               f"TOC 錨點 '#{anchor}' 在本頁中找不到對應的 id"))
            else:
                passed_links += 1

        # --- 4. 所有連結檢查（a / link / script / img） ---
        for entry in analyzer.links:
            href = entry["href"]
            ctx = entry["context"]
            line = entry["line"]
            if is_external(href):
                continue
            if href in ("", "#"):
                continue
            # TOC 內的純錨點已在上面檢查過，避免重複
            if ctx == CTX_TOC and href.startswith("#"):
                continue

            total_links += 1
            target, anchor = resolve_target(fp, href)
            if target is None:
                continue
            if not os.path.isfile(target):
                rel_target = (os.path.relpath(target, DOCS_ROOT)
                              if target.startswith(DOCS_ROOT) else target)
                issues.append(("BROKEN", rel, line, href,
                               "目標文件不存在: " + rel_target))
                continue
            if anchor:
                ids = get_ids_in_file(target)
                if anchor not in ids:
                    issues.append(("MISSING_ANCHOR", rel, line, href,
                                   f"錨點 '#{anchor}' 在目標文件中找不到對應的 id"))
                    continue
            passed_links += 1

    # --- 輸出報告 ---
    sev_order = ["BROKEN", "MISSING_ANCHOR", "WARNING"]
    sev_titles = {
        "BROKEN": "BROKEN（文件不存在 / 404）",
        "MISSING_ANCHOR": "MISSING_ANCHOR（文件存在但錨點目標不存在）",
        "WARNING": "WARNING（其他問題）",
    }

    print("=" * 80)
    print("FTP OpenAPI 文檔站點連結完整性檢查報告")
    print("=" * 80)
    print(f"掃描根目錄       : {DOCS_ROOT}")
    print(f"檢查 HTML 文件數 : {files_checked}")
    print(f"檢查連結總數     : {total_links}")
    print(f"通過連結數       : {passed_links}")
    print(f"問題總數         : {len(issues)}")
    print("=" * 80)

    for sev in sev_order:
        sev_issues = [i for i in issues if i[0] == sev]
        print()
        print(f"## {sev_titles[sev]}  ({len(sev_issues)} 個)")
        print("-" * 80)
        if not sev_issues:
            print("  (無)")
            continue
        by_file = defaultdict(list)
        for it in sev_issues:
            by_file[it[1]].append(it)
        for fname in sorted(by_file.keys()):
            print(f"\n  📄 {fname}")
            for it in sorted(by_file[fname], key=lambda x: x[2] or 0):
                _, _, line, href, msg = it
                line_str = f"L{line}" if line else "  -"
                href_str = f"「{href}」" if href else ""
                print(f"    {line_str:>6}  {href_str}  →  {msg}")

    print()
    print("=" * 80)
    print("統計摘要")
    print("=" * 80)
    print(f"  HTML 文件總數       : {files_checked}")
    print(f"  內部連結檢查總數   : {total_links}")
    print(f"  通過 (passed)       : {passed_links}")
    print(f"  失敗 (failed)       : {total_links - passed_links}")
    for sev in sev_order:
        n = sum(1 for i in issues if i[0] == sev)
        print(f"  {sev:<16}: {n}")
    print("=" * 80)

    # 若存在 BROKEN 則以非零退出
    if any(i[0] == "BROKEN" for i in issues):
        sys.exit(1)
    sys.exit(0)


if __name__ == "__main__":
    main()
