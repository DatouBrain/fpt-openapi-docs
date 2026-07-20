# CLAUDE.md — Fintech Payment OpenAPI 文檔

## 項目定位
Fintech Payment 開放平台 API 文檔靜態網站（繁中/英文雙語），純 HTML/CSS/JS，無構建依賴。

## 本地預覽
```bash
python3 -m http.server 8000 --directory docs
```

## 目錄約定
- `docs/` — 繁中版（根目錄）
- `docs/en/` — 英文版（鏡像結構，路徑深度多一層）
- `docs/common/` — 通用規則頁（7 頁，所有產品共用）
- `docs/payments/` — 產品文檔頁（按 card-online/card-offline/qr-online/qr-offline/other 分類）
- `docs/assets/` — CSS/JS/圖片
- `docs-config.json` — 隱藏頁面配置

## 編輯規則
- 修改任何頁面時，**中英文雙語同步更新**
- 產品頁的簽名章節用連結引用 `common/signature.html`，不要內嵌
- sidebar 結構全站統一，新增頁面需在所有頁面的 sidebar 中添加連結
- 隱藏頁面（ftp-foreign）不在其他頁面的 sidebar 中出現，僅自身 sidebar 包含 Other Pay 分組
- 代碼塊使用 `<pre><code>` 格式，不要用散亂的 `<p>` 段落
- 響應式斷點：≤900px 移動端、≤1400px 普通、≥1920px 2K、≥2560px 4K

## 部署
- push 到 `main` → GitHub Actions 自動部署到 GitHub Pages
- 打包：`zip -r ftp-openapi-docs.zip docs/ -x "*.DS_Store"`

## 當前狀態
- 29 個產品頁 + 7 個通用頁 + 2 個首頁 = 38 個 HTML 頁面（中英各 19）
- Fintech Payment Online service 值已從 sinopay 改為 ftp
- 簽名規則 V1.1（支持 items 巢狀欄位 canonical JSON）
