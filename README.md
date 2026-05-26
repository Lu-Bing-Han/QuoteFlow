# QuoteFlow

立善科技內部單據生成工具，支援出貨單、報價單、驗機單、維修單、標籤列印、出貨排程、Trello 卡片同步等功能。

## 專案結構

```
src/
├── app.py                    ← 入口（PyInstaller 打包起點）
├── _paths.py                 ← 路徑常數（開發 / 打包雙模式）
├── core/
│   ├── __init__.py
│   ├── parser.py             ← 報價單解析
│   ├── generator.py          ← 出貨單生成
│   ├── generator_fix.py      ← 維修單生成
│   ├── generator_tag.py      ← 維修掛件生成
│   ├── generator_label.py    ← 標籤 PDF 生成
│   ├── generator_inspection.py ← 驗機單生成
│   ├── generator_schedule.py ← 出貨排程生成
│   └── generator_quote.py    ← 報價單生成
├── sync/
│   ├── __init__.py
│   ├── syncer_trello.py      ← Trello 卡片抓取
│   ├── syncer_sheets.py      ← Google Sheets 同步
│   ├── syncer_production.py  ← 生產群組紀錄同步
│   ├── creator_trello.py     ← Trello 卡片建立
│   └── downloader_trello.py  ← Trello 卡片下載
└── ui/
    ├── __init__.py
    ├── app_core.py           ← 主視窗 App 類別 + 共用方法
    ├── mixin_documents.py    ← 出貨單、驗機單、維修單、維修掛件頁籤
    ├── mixin_quote.py        ← 報價單頁籤
    ├── mixin_label.py        ← 標籤生成頁籤
    ├── mixin_schedule.py     ← 出貨排程頁籤
    └── mixin_trello.py       ← Trello 相關頁籤

template/                     ← Excel / PDF 範本、憑證檔案
QuoteFlow.spec                ← PyInstaller 打包設定
requirements.txt              ← Python 套件需求
```

## 環境需求

- Python 3.10+
- 安裝套件：`pip install -r requirements.txt`

## 開發執行

從專案根目錄執行：

```bash
python src/app.py
```

## 打包

```bash
pyinstaller QuoteFlow.spec
```

輸出位置：`dist\QuoteFlow.exe`

## 輸出路徑設定

點擊右上角 ⚙ 圖示，可設定各功能的輸出資料夾或目標 Excel 路徑，設定儲存於執行檔同目錄的 `config.json`。

## template/ 說明

| 檔案 | 說明 |
|------|------|
| template_AP.xlsx | 出貨單範本 |
| template_quote.xlsx | 報價單範本 |
| template_schedule.xlsx | 出貨排程範本（含地址工作表）|
| products.json | 報價產品目錄 |
| credentials.json | Google Sheets OAuth 憑證（自行放置）|
| icon.png | 應用程式圖示 |

## 功能說明

| 頁籤 | 說明 |
|------|------|
| 出貨單 | 從報價單生成出貨單 Excel |
| 報價單 | 三步驟流程（Trello 客戶 → 選取品項 → 生成報價單 Excel）|
| 驗機單 | 從報價單生成驗機單 Excel + Word |
| 維修單 | 從報價單生成維修單 |
| 維修掛件 | 生成維修掛件 Word 文件 |
| 標籤生成 | 批次生成標籤 PDF（多種樣式）|
| 出貨排程 | 抓取 Timetree 行事曆，計算行車時間，寫入排程表 |
| 出貨一覽表 | 同步 Trello 本周下單 → Google Sheets |
| 生產群組紀錄 | 同步 Trello → 生產群組紀錄 Excel |
| 建立卡片 | 從 Excel 批次建立 Trello 卡片 |
| 下載卡片 | 下載 Trello 卡片描述與附件 |
