import gspread
from oauth2client.service_account import ServiceAccountCredentials
import pandas as pd
from datetime import datetime
import pytz

# Google Sheets setup
SCOPE = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
SPREADSHEET_URL = 'https://docs.google.com/spreadsheets/d/1P48quxwMv9XsYQhXjLOvTRRq8tt3ahJnkbXo4VCxjLc/edit?gid=1615412834'

def convert_timestamp(timestamp_str):
    try:
        # 处理 "上午/下午" 的时间格式
        if '上午' in timestamp_str:
            timestamp_str = timestamp_str.replace('上午', 'AM')
        elif '下午' in timestamp_str:
            timestamp_str = timestamp_str.replace('下午', 'PM')
        
        # 将时间字符串解析为datetime对象
        dt = pd.to_datetime(timestamp_str, format='%Y-%m-%d %p%I:%M:%S')
        return dt.strftime('%Y-%m-%d %H:%M:%S')
    except:
        return timestamp_str

def fetch_and_convert():
    # Authenticate with Google
    credentials = ServiceAccountCredentials.from_json_keyfile_name('service_account.json', SCOPE)
    client = gspread.authorize(credentials)
    
    # Open the spreadsheet
    spreadsheet = client.open_by_url(SPREADSHEET_URL)
    worksheet = spreadsheet.get_worksheet(0)  # Get first worksheet
    
    # Get all values
    data = worksheet.get_all_values()
    
    # Convert to pandas DataFrame
    df = pd.DataFrame(data[1:], columns=data[0])
    
    # 转换时间戳记列
    if '时间戳记' in df.columns:
        df['时间戳记'] = df['时间戳记'].apply(convert_timestamp)
    
    # Get current time in UTC+8
    china_tz = pytz.timezone('Asia/Shanghai')
    current_time = datetime.now(china_tz).strftime('%Y-%m-%d %H:%M:%S')
    
    # Convert to HTML with styling
    html_content = f"""
    <!DOCTYPE html>
    <html lang="zh-CN">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>611Study.ICU - 全国超时学习学校耻辱名单</title>
        <style>
            body {{
                font-family: Arial, sans-serif;
                margin: 20px;
                line-height: 1.6;
            }}
            table {{
                border-collapse: collapse;
                width: 100%;
                margin: 20px 0;
            }}
            th, td {{
                border: 1px solid #ddd;
                padding: 12px;
                text-align: left;
            }}
            th {{
                background-color: #f4f4f4;
            }}
            tr:nth-child(even) {{
                background-color: #f9f9f9;
            }}
            tr:hover {{
                background-color: #f5f5f5;
            }}
            .last-updated {{
                color: #666;
                font-size: 0.9em;
                margin-bottom: 10px;
            }}
            .timezone-notice {{
                color: #ff4444;
                font-size: 0.9em;
                margin-bottom: 20px;
            }}
        </style>
    </head>
    <body>
        <div class="last-updated">最后更新时间：{current_time} (UTC+8)</div>
        <div class="timezone-notice">注意！本站时间与原表格不同，为便于理解，已转换为北京时间！</div>
        {df.to_html(index=False, classes='table', escape=False)}
    </body>
    </html>
    """
    
    # Save to index.html
    with open('index.html', 'w', encoding='utf-8') as f:
        f.write(html_content)

if __name__ == '__main__':
    fetch_and_convert() 