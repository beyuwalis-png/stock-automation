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

    try:
        # 1. 資料清洗
        cols = ['成交金額', '收盤價', '漲跌價差']
        for col in cols:
            df[col] = pd.to_numeric(df[col].astype(str).str.replace(',', ''), errors='coerce')
        
        df['昨收'] = df['收盤價'] - df['漲跌價差']
        df['漲幅'] = (df['漲跌價差'] / df['昨收']) * 100
        df['成交額(億)'] = (df['成交金額'] / 100000000).round(1)
        df['漲幅'] = df['漲幅'].round(2)

        # 2. 核心篩選（漲幅需 > 2.5%）
        filtered_df = df[df['漲幅'] > 2.5].copy()

        # 3. 定義格式化函數 (超連結與顏色)
        def create_link(row):
            code = str(row['證券代號']).strip()
            name = row['證券名稱']
            return f'<a href="https://tw.stock.yahoo.com/quote/{code}" style="text-decoration:none; color:#0066cc; font-weight:bold;">{name}</a>'

        def format_change_color(val):
            color = "#FF0000" if val > 5.0 else "#D20000"
            weight = "bold" if val > 5.0 else "normal"
            return f'<span style="color: {color}; font-weight: {weight};">{val:.2f}%</span>'

        # 4. 準備三種排序的 HTML 表格
        def generate_styled_table(data_df, sort_by):
            temp_df = data_df.sort_values(by=sort_by, ascending=False).head(10).copy()
            # 轉換顯示格式
            temp_df['證券名稱'] = temp_df.apply(create_link, axis=1)
            temp_df['漲幅'] = temp_df['漲幅'].apply(format_change_color)
            return temp_df[['證券代號', '證券名稱', '收盤價', '漲幅', '成交額(億)']].to_html(index=False, escape=False)

        # 產生三個表格
        table_volume = generate_styled_table(filtered_df, '成交金額')
        table_gain = generate_styled_table(filtered_df, '漲幅')
        table_price = generate_styled_table(filtered_df, '收盤價')

        # 5. HTML 樣式與組合
        html_style = """
        <style>
            table { border-collapse: collapse; width: 100%; font-family: "Microsoft JhengHei", sans-serif; margin-bottom: 20px; }
            th { background-color: #4CAF50; color: white; padding: 10px; text-align: left; }
            td { padding: 8px; border-bottom: 1px solid #ddd; }
            h3 { color: #2c3e50; border-left: 5px solid #4CAF50; padding-left: 10px; margin-top: 30px; }
        </style>
        """
        
        full_html = f"""
        <html>
        <head>{html_style}</head>
        <body>
            <h2 style="color: #2c3e50;">📈 台股盤後多維度強勢股報告</h2>
            <p>報告日期：{datetime.datetime.now().strftime('%Y-%m-%d')}</p>
            <hr>
            
            <h3>🔥 資金焦點：成交額 Top 10 (強勢股)</h3>
            {table_volume}
            
            <h3>🚀 漲幅先鋒：漲幅 Top 10 (強勢股)</h3>
            {table_gain}
            
            <h3>💎 高價指標：股價 Top 10 (強勢股)</h3>
            {table_price}
            
            <br>
            <p style="color: gray; font-size: 12px;">註：以上列表皆已先篩選漲幅 > 2.5% 之個股。點擊名稱看線圖。</p>
        </body>
        </html>
        """

        send_email_report(full_html, datetime.datetime.now().strftime('%Y-%m-%d'))
    except Exception as e:
        print(f"❌ 資料處理發生錯誤: {e}")
if __name__ == "__main__":
    process_and_mail()
