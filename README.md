# 台灣熱門股票投資組合自動報表系統

這是我學習 Python 資料分析時完成的個人專案，目標是建立一個完整的資料管線：**抓取資料 → 存入資料庫 → 分析計算 → 自動產出 Excel 報表**。

### 專案功能
- 使用 yfinance 抓取台灣上市股票歷史資料（台積電、聯發科等）
- 透過 SQLite 建立資料管線
- 使用 Pandas 計算每日報酬率、累計報酬率、20日移動平均線
- 自動產生含多分頁與圖表的 Excel 報表
- 包含投資組合總覽與風險指標

### 使用技術
- **Python**：函式、類別、迴圈、檔案處理
- **Pandas**：資料清洗、群組分析、報酬率計算
- **SQLite**：資料儲存與查詢
- **Matplotlib**：資料視覺化
- **openpyxl**：Excel 報表自動化與圖片插入
- **yfinance**：股票資料 API

### 專案成果展示
（這裡放上傳後的圖片連結，或直接截圖 Excel）

### 如何執行
1. 安裝套件：`pip install yfinance pandas matplotlib openpyxl`
2. 執行主程式：`python final_report.py`
3. 執行後會自動產生 `stock_portfolio_report_final.xlsx`

### 學習心得
這是我從零開始學習 Python 後，第一個完成的完整資料分析專案。過程中我學會了如何把不同工具（API、資料庫、Excel）串接起來，更加理解資料分析在 FinTech 的實際應用。

---

**日期**：2026年5月  
**狀態**：已完成
