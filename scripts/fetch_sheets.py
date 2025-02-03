import gspread
from oauth2client.service_account import ServiceAccountCredentials
import pandas as pd

# Google Sheets setup
SCOPE = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
SPREADSHEET_URL = 'https://docs.google.com/spreadsheets/d/1P48quxwMv9XsYQhXjLOvTRRq8tt3ahJnkbXo4VCxjLc/edit?gid=1615412834'

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
    
    # Convert to HTML with styling
    html_content = f"""
    <!DOCTYPE html>
    <html lang="en">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>Google Sheets Data</title>
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
                margin-bottom: 20px;
            }}
        </style>
    </head>
    <body>
        <div class="last-updated">Last updated: {pd.Timestamp.now().strftime('%Y-%m-%d %H:%M:%S')} UTC</div>
        {df.to_html(index=False, classes='table', escape=False)}
    </body>
    </html>
    """
    
    # Save to index.html
    with open('index.html', 'w', encoding='utf-8') as f:
        f.write(html_content)

if __name__ == '__main__':
    fetch_and_convert() 