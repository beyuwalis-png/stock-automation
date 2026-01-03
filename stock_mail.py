import os
import requests
import pandas as pd
import datetime
import io
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

# --- 郵件設定 (從 GitHub Secrets 環境變數讀取) ---
SENDER_EMAIL = os.environ.get('MY_EMAIL')
RECEIVER_EMAIL = os.environ.get('MY_EMAIL')
APP_PASSWORD = os.environ.get('MY_PASSWORD')

def get_stock_data():
    """使用證交所 Open Data CSV"""
    url = "https://www.twse.com.tw/exchangeReport/STOCK_DAY_ALL?response=open_data"
    headers = {"User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"}
    try:
        response = requests.get(url, headers=headers, timeout=15)
        # 清洗可能存在的特殊空白字元
        clean_text = response.text.replace('\xa0', ' ')
        df = pd.read_csv(io.StringIO(clean_text))
        return df, "OK"
    except Exception as e:
        return None, str(e)

def send_email_report(html_content, date_str):
    # 建立郵件物件
    msg = MIMEMultipart()
    msg['From'] = SENDER_EMAIL
    msg['To'] = RECEIVER_EMAIL
    msg['Subject'] = f"📊 台股強勢股日報 - {date_str}"

    # 指定 'html' 格式與 'utf-8' 編碼
    part = MIMEText(html_content, 'html', 'utf-8')
    msg.attach(part)

    try:
        # 使用 Gmail SMTP 伺服器
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(SENDER_EMAIL, APP_PASSWORD)
        
        # 使用 send_message 自動處理編碼轉換
        server.send_message(msg)
        server.quit()
        print(f"✅ 郵件發送成功！已寄至 {RECEIVER_EMAIL}")
    except Exception as e:
        print(f"❌ 郵件發送失敗: {str(e)}")

def process_and_mail():
    df, status = get_stock_data()
    if df is None or df.empty:
        print(f"❌ 無法取得資料: {status}")
        return

    # 資料清洗與計算
    try:
        cols = ['成交金額', '收盤價', '漲跌價差']
        for col in cols:
            df[col] = pd.to_numeric(df[col].astype(str).str.replace(',', ''), errors='coerce')
        
        df['昨收'] = df['收盤價'] - df['漲跌價差']
        df['漲幅'] = (df['漲跌價差'] / df['昨收']) * 100
        df['成交額(億)'] = (df['成交金額'] / 100000000).round(1)
        df['漲幅'] = df['漲幅'].round(2)

        # 篩選前 20 檔強勢股
        top_20 = df[df['漲幅'] > 2.5].sort_values(by='成交金額', ascending=False).head(20).copy()

        # 為證券名稱建立超連結 (連結至 Yahoo 股市)
        def create_link(row):
            code = str(row['證券代號']).strip()
            name = row['證券名稱']
            url = f"https://tw.stock.yahoo.com/quote/{code}"
            return f'<a href="{url}" style="text-decoration:none; color:#0066cc; font-weight:bold;">{name}</a>'

        top_20['證券名稱'] = top_20.apply(create_link, axis=1)
        top_20 = top_20[['證券代號', '證券名稱', '收盤價', '漲幅', '成交額(億)']]

        # HTML 表格樣式
        html_style = """
        <style>
            table { border-collapse: collapse; width: 100%; font-family: "Microsoft JhengHei", sans-serif; }
            th { background-color: #4CAF50; color: white; padding: 12px; text-align: left; }
            td { padding: 10px; border-bottom: 1px solid #ddd; }
            tr:nth-child(even) { background-color: #f9f9f9; }
            tr:hover { background-color: #f1f1f1; }
            .up { color: #d20000; font-weight: bold; }
        </style>
        """
        
        table_html = top_20.to_html(index=False, classes='stock-table', escape=False)

        full_html = f"""
        <html>
        <head>{html_style}</head>
        <body>
            <h2 style="color: #2c3e50;">📈 台股盤後強勢股篩選報告</h2>
            <p>報告日期：{datetime.datetime.now().strftime('%Y-%m-%d')}</p>
            <p style="font-size: 14px; color: #666;">💡 提示：點擊「證券名稱」可直接跳轉至 Yahoo 股市查看線圖。</p>
            <hr>
            {table_html}
            <br>
            <p style="color: gray; font-size: 12px;">資料來源：臺灣證券交易所 Open Data</p>
        </body>
        </html>
        """

        send_email_report(full_html, datetime.datetime.now().strftime('%Y-%m-%d'))
    except Exception as e:
        print(f"❌ 資料處理發生錯誤: {e}")

if __name__ == "__main__":
    process_and_mail()
