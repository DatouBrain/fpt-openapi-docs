#!/usr/bin/env python3
"""
Inject language switcher into all Chinese HTML pages, update main.js to be
language-aware, and append language switcher CSS.

This script is idempotent: it removes any previously injected lang-switcher
block before re-inserting, so it can be run repeatedly.
"""
import os
import re

DOCS = os.path.join(os.path.dirname(os.path.abspath(__file__)), "docs")

ZH_SWITCHER = (
    '<div class="lang-switcher">'
    '<span class="lang-current">繁中</span>'
    '<span class="lang-sep">/</span>'
    '<a href="{en_url}" class="lang-en">EN</a>'
    '</div>'
)

# Regex to find the menu-toggle button to inject before it
MENU_TOGGLE_RE = re.compile(r'(\s*)<button class="menu-toggle"[^>]*>☰</button>')
# Regex to remove a previously injected lang-switcher block
OLD_SWITCHER_RE = re.compile(r'\s*<div class="lang-switcher">.*?</div>\s*\n?', re.DOTALL)


def collect_html_files(root):
    """Yield (abs_path, relpath) for every .html under root, excluding the en/ subtree."""
    for dirpath, dirnames, filenames in os.walk(root):
        # skip the en/ subtree entirely
        rel_dir = os.path.relpath(dirpath, root)
        if rel_dir == "en" or rel_dir.startswith("en" + os.sep):
            continue
        for fn in filenames:
            if fn.endswith(".html"):
                abs_p = os.path.join(dirpath, fn)
                rel_p = os.path.relpath(abs_p, root)
                yield abs_p, rel_p.replace(os.sep, "/")


def en_counterpart_url(relpath):
    """Compute relative URL from a Chinese page to its English counterpart."""
    depth = relpath.count("/")
    prefix = "../" * depth
    return prefix + "en/" + relpath


def inject_switcher(abs_path, relpath):
    with open(abs_path, "r", encoding="utf-8") as f:
        content = f.read()

    # Remove any previously injected switcher
    content = OLD_SWITCHER_RE.sub("\n", content)

    en_url = en_counterpart_url(relpath)
    switcher_html = ZH_SWITCHER.format(en_url=en_url)

    m = MENU_TOGGLE_RE.search(content)
    if not m:
        print("WARN: menu-toggle not found in", relpath)
        return False
    indent = m.group(1)
    insertion = indent + switcher_html + "\n" + indent
    content = content[:m.start()] + insertion + content[m.start():]

    with open(abs_path, "w", encoding="utf-8") as f:
        f.write(content)
    return True


def update_main_js():
    js_path = os.path.join(DOCS, "assets", "js", "main.js")
    with open(js_path, "r", encoding="utf-8") as f:
        content = f.read()

    # Replace the hardcoded Chinese copy button text with language-aware logic.
    # Detect via document.documentElement.lang.
    if "var isEn =" in content:
        # already updated
        return False

    content = content.replace(
        "btn.textContent = '複製';",
        "btn.textContent = (document.documentElement.lang === 'en') ? 'Copy' : '複製';",
    )
    content = content.replace(
        "btn.textContent = '已複製';",
        "btn.textContent = (document.documentElement.lang === 'en') ? 'Copied' : '已複製';",
    )
    content = content.replace(
        "setTimeout(function() { btn.textContent = '複製'; }, 2000);",
        "setTimeout(function() { btn.textContent = (document.documentElement.lang === 'en') ? 'Copy' : '複製'; }, 2000);",
    )

    with open(js_path, "w", encoding="utf-8") as f:
        f.write(content)
    return True


CSS_BLOCK = """
/* ===== Language switcher ===== */
.lang-switcher {
  display: inline-flex;
  align-items: center;
  gap: 6px;
  font-size: 13px;
  margin-left: 16px;
  padding: 4px 10px;
  border: 1px solid var(--border, #e3e8ee);
  border-radius: 6px;
  background: #fff;
  white-space: nowrap;
}
.lang-switcher .lang-current {
  font-weight: 600;
  color: var(--primary, #635bff);
}
.lang-switcher .lang-sep {
  color: #c4cdda;
}
.lang-switcher a {
  color: #6b7280;
  text-decoration: none;
  transition: color .15s;
}
.lang-switcher a:hover {
  color: var(--primary, #635bff);
}
@media (max-width: 900px) {
  .lang-switcher {
    margin-left: 0;
    order: 3;
  }
}
"""


def append_css():
    css_path = os.path.join(DOCS, "assets", "css", "style.css")
    with open(css_path, "r", encoding="utf-8") as f:
        content = f.read()
    if "Language switcher" in content:
        return False
    if not content.endswith("\n"):
        content += "\n"
    content += CSS_BLOCK
    with open(css_path, "w", encoding="utf-8") as f:
        f.write(content)
    return True


def main():
    count = 0
    for abs_p, rel_p in collect_html_files(DOCS):
        if inject_switcher(abs_p, rel_p):
            count += 1
    print("Injected lang-switcher into %d Chinese pages" % count)

    if update_main_js():
        print("Updated main.js (language-aware copy button)")
    else:
        print("main.js already up to date")

    if append_css():
        print("Appended lang-switcher CSS")
    else:
        print("CSS already present")


if __name__ == "__main__":
    main()
