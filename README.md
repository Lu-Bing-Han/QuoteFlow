# 報價單 → 出貨單 轉換工具

## 專案結構

```
quote_to_shipping/
├── src/
│   ├── app.py        # GUI 主程式 (入口)
│   ├── parser.py     # 報價單解析器
│   └── generator.py  # 出貨單產生器
├── template/
│   └── template.xlsx # 出貨單模板 (請保留，不要刪除)
├── output/           # 生成的出貨單會存在這裡
└── README.md
```

## 環境安裝

```bash
pip install openpyxl pandas
```

## 啟動程式

```bash
cd quote_to_shipping/src
python app.py
```

## 使用流程

1. 點擊「選擇報價單 .xlsx」匯入報價單
2. 左側會自動填入客戶資料（唯讀）
3. 右側補填出貨日期、銷貨單號、製表人員
4. 品項列表可雙擊編輯、新增或刪除
5. 點「生成出貨單」→ 自動存到 output/ 資料夾

## 打包成執行檔 (不需安裝 Python 也能跑)

```bash
pip install pyinstaller
pyinstaller --onefile --windowed --name 出貨單轉換工具 src/app.py
```
生成的執行檔在 `dist/` 資料夾。

## 欄位對應邏輯

| 報價單儲存格 | 意義       | → 出貨單儲存格 |
|------------|-----------|--------------|
| B9         | 報價單號   | (參考用)      |
| B10        | 客戶全名   | C4            |
| B11        | 電話       | C5            |
| B12        | 聯絡人     | F5            |
| H9         | 幣別       | I5            |
| 第15行起   | 品項列表   | 第8行起       |

## 注意事項

- `template/template.xlsx` 是出貨單的底板，格式、字型、框線都從這裡繼承
- 若報價單品項規格跨多行（例如「前小輪」補充說明），程式會自動合併成一行
- 出貨單輸出檔名格式：`出貨單-{客戶名稱}-{出貨日期}.xlsx`
