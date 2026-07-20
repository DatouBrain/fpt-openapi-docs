# Fintech Payment OpenAPI 文檔

Fintech Payment 開放平台的 API 接口文檔靜態網站，提供繁中/英文雙語版本。

## 本地預覽

```bash
python3 -m http.server 8000 --directory docs
```

打開 http://localhost:8000

## 目錄結構

```
docs/
├── index.html              # 繁中首頁
├── assets/                 # 靜態資源（CSS/JS/圖片）
├── common/                 # 通用規則頁（7 頁）
│   ├── quick-start.html    # 快速入門
│   ├── endpoints.html      # API 調用地址
│   ├── data-format.html    # 數據格式
│   ├── signature.html      # 數字簽名
│   ├── request-headers.html# 請求頭說明
│   ├── error-codes.html    # 錯誤碼說明
│   └── notification.html   # 支付通知說明
├── payments/               # 產品文檔頁（15 個）
│   ├── card-online/        # 線上卡支付
│   ├── card-offline/       # 線下卡支付
│   ├── qr-online/          # 線上掃碼支付
│   ├── qr-offline/         # 線下掃碼支付
│   └── other/              # 其他支付
└── en/                     # 英文版（鏡像結構）
```

## 產品文檔

| 分類 | 產品 |
|---|---|
| Card Pay (Online) | Global Payments Online、Fintech Payment Online |
| Card Pay (Offline) | Global Payments Offline（暫未開放） |
| QR Pay (Online) | 動態銀碼掃碼、Alipay WAP/WEB/APP/Plus、WeChat WEB/WAP/APP/小程序 |
| QR Pay (Offline) | 付款碼支付 |
| Other Pay | FTP 外幣支付（暫未開放） |

## 部署

推送到 `main` 分支後，GitHub Actions 自動部署到 GitHub Pages。

**打包給運維部署到 AWS**：
```bash
zip -r ftp-openapi-docs.zip docs/ -x "*.DS_Store"
```

## 構建腳本

- `scripts/convert_docs.py` — docx 轉 HTML（一次性，源文件已歸檔）
- `scripts/fix_tables.py` — 表格修復（一次性，已完成）
- `add_lang_switcher.py` — 語言切換器注入（可用於新增頁面）
- `docs-config.json` — 頁面隱藏配置

## 技術棧

- 純靜態 HTML/CSS/JS，無構建依賴
- 響應式設計（移動端 / 普通屏 / 2K / 4K）
- Stripe API Docs 風格 UI
