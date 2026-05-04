
import sqlite3
import pandas as pd
import matplotlib.pyplot as plt
from openpyxl import load_workbook   
from openpyxl.drawing.image import Image   

db_name = "stocks.db"
excel_file_name = "stock_portfolio_report_final.xlsx"

# 函式 1：從資料庫讀取資料 
def get_all_data():
    conn = sqlite3.connect(db_name)
    sql = """
    SELECT * FROM stock_prices 
    ORDER BY date, ticker
    """
    df = pd.read_sql(sql, conn)
    conn.close()
    return df

# 函式 2：資料處理與計算 
def process_data(df):
    df_new = df.copy()
    
    # 轉成日期格式 + 排序
    df_new["date"] = pd.to_datetime(df_new["date"])
    df_new = df_new.sort_values(by=["ticker", "date"])
    
    # 計算各種指標
    df_new["daily_return"] = df_new.groupby("ticker")["close"].pct_change()
    df_new["cum_return"] = (1 + df_new["daily_return"]).groupby(df_new["ticker"]).cumprod() - 1
    df_new["ma20"] = df_new.groupby("ticker")["close"].transform(lambda x: x.rolling(window=20).mean())
    
    # 填補缺失值
    df_new["daily_return"] = df_new["daily_return"].fillna(0)
    df_new["ma20"] = df_new["ma20"].ffill()
    
    return df_new

# 函式 3：產生投資組合總覽
def create_portfolio_summary(df):
    summary = df.groupby("ticker").agg(
        latest_close=("close", "last"),
        avg_daily_return=("daily_return", "mean"),
        cum_return=("cum_return", "last"),
        volatility=("daily_return", "std")
    )
    summary = summary.reset_index()
    
    # 轉成百分比
    summary["volatility"] = summary["volatility"] * 100
    summary["cum_return"] = summary["cum_return"] * 100
    summary["avg_daily_return"] = summary["avg_daily_return"] * 100
    
    return summary

#  函式 4：產生所有圖表 
def create_charts(df):
    
    # 畫台積電累計報酬率
    tsmc = df[df["ticker"] == "2330.TW"]
    
    plt.figure(figsize=(10, 6))
    plt.plot(tsmc["date"], tsmc["cum_return"] * 100, color="blue", linewidth=2)
    plt.title("台積電 (2330.TW) 累計報酬率走勢")
    plt.xlabel("日期")
    plt.ylabel("累計報酬率 (%)")
    plt.grid(True)
    plt.savefig("tsmc_cum_return.png")
    print("已儲存 tsmc_cum_return.png")
    
    # 所有股票累計報酬率比較圖
    plt.figure(figsize=(10, 6))
    for ticker in df["ticker"].unique():
        stock = df[df["ticker"] == ticker]
        plt.plot(stock["date"], stock["cum_return"] * 100, label=ticker)
    
    plt.title("所有股票累計報酬率比較")
    plt.xlabel("日期")
    plt.ylabel("累計報酬率 (%)")
    plt.legend()
    plt.grid(True)
    plt.savefig("all_stocks_cum_return.png")
    print("已儲存 all_stocks_cum_return.png")

#  函式 5：產生 Excel 並插入圖片 
def create_final_excel(df, summary):
    
    with pd.ExcelWriter(excel_file_name, engine="openpyxl") as writer:
        summary.to_excel(writer, sheet_name="投資組合總覽", index=False)
        df.to_excel(writer, sheet_name="個股詳細資料", index=False)
        
        # 最近30天資料
        recent = df.groupby("ticker").tail(30)
        recent.to_excel(writer, sheet_name="最近30天資料", index=False)
    
    # 插入圖片到 Excel
    wb = load_workbook(excel_file_name)
    
    # 在「投資組合總覽」分頁插入第一張圖
    ws = wb["投資組合總覽"]
    img1 = Image("tsmc_cum_return.png")
    img1.width = 600
    img1.height = 360
    ws.add_image(img1, "F5")   # 放在 F5 位置
    
    # 在「個股詳細資料」分頁插入第二張圖
    ws2 = wb["個股詳細資料"]
    img2 = Image("all_stocks_cum_return.png")
    img2.width = 600
    img2.height = 360
    ws2.add_image(img2, "H5")
    
    wb.save(excel_file_name)
    print("圖表已自動插入 Excel")

# 主執行部分

# Step 1：讀取資料
df_raw = get_all_data()

# Step 2：資料處理
df = process_data(df_raw)

# Step 3：產生總覽
summary = create_portfolio_summary(df)

# Step 4：產生圖表
create_charts(df)

# Step 5：產生 Excel + 插入圖片
create_final_excel(df, summary)


print("請打開以下檔案查看最終成果：")
print("1. " + excel_file_name + " （最重要！含圖表的專業 Excel 報表）")
print("2. tsmc_cum_return.png")
print("3. all_stocks_cum_return.png")
