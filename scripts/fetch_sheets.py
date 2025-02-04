import gspread
from oauth2client.service_account import ServiceAccountCredentials
import pandas as pd
from datetime import datetime
import pytz

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
            .controls {{
                margin: 20px 0;
                display: flex;
                gap: 20px;
                flex-wrap: wrap;
            }}
            .search-box {{
                padding: 8px;
                border: 1px solid #ddd;
                border-radius: 4px;
                width: 300px;
                max-width: 100%;
            }}
            .filter-container {{
                display: flex;
                gap: 10px;
                flex-wrap: wrap;
            }}
            .filter-select {{
                padding: 8px;
                border: 1px solid #ddd;
                border-radius: 4px;
                min-width: 150px;
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
                position: relative;
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
            .highlight {{
                background-color: #ffeb3b;
                padding: 2px;
                border-radius: 2px;
            }}
            .hidden {{
                display: none;
            }}
        </style>
    </head>
    <body>
        <div class="last-updated">最后更新时间：{current_time} (UTC+8)</div>
        <div class="timezone-notice">注意！本站时间与原表格一致，<u>仅最后更新时间</u>为北京时间。</div>
        
        <div class="controls">
            <input type="text" class="search-box" placeholder="输入关键词进行搜索..." id="searchInput">
            <div class="filter-container" id="filterContainer">
                <!-- Filter dropdowns will be added here dynamically -->
            </div>
        </div>

        <div id="tableContainer">
            {df.to_html(index=False, classes='table', escape=False)}
        </div>

        <script>
            document.addEventListener('DOMContentLoaded', function() {{
                const table = document.querySelector('table');
                const searchInput = document.getElementById('searchInput');
                const filterContainer = document.getElementById('filterContainer');
                
                // Create filter dropdowns for each column
                const headers = Array.from(table.querySelectorAll('th'));
                headers.forEach((header, index) => {{
                    const values = new Set();
                    Array.from(table.querySelectorAll(`td:nth-child(${{index + 1}})`))
                        .forEach(cell => values.add(cell.textContent.trim()));

                    const select = document.createElement('select');
                    select.className = 'filter-select';
                    select.innerHTML = `
                        <option value="">筛选 ${{header.textContent}}</option>
                        ${{Array.from(values).sort().map(value => 
                            `<option value="${{value}}">${{value}}</option>`
                        ).join('')}}
                    `;
                    
                    select.addEventListener('change', applyFilters);
                    filterContainer.appendChild(select);
                }});

                // Search and highlight functionality
                searchInput.addEventListener('input', applyFilters);

                function applyFilters() {{
                    const searchTerm = searchInput.value.toLowerCase();
                    const filterValues = Array.from(document.querySelectorAll('.filter-select'))
                        .map(select => ({{
                            index: Array.from(select.parentNode.children).indexOf(select),
                            value: select.value.toLowerCase()
                        }}));

                    const rows = Array.from(table.querySelectorAll('tr'));
                    
                    // Skip header row
                    rows.slice(1).forEach(row => {{
                        const cells = Array.from(row.querySelectorAll('td'));
                        const rowText = cells.map(cell => cell.textContent.toLowerCase()).join(' ');
                        
                        // Check if row matches search term and all active filters
                        const matchesSearch = searchTerm === '' || rowText.includes(searchTerm);
                        const matchesFilters = filterValues.every(filter => 
                            filter.value === '' || 
                            cells[filter.index].textContent.toLowerCase() === filter.value
                        );

                        if (matchesSearch && matchesFilters) {{
                            row.classList.remove('hidden');
                            // Highlight search term
                            if (searchTerm) {{
                                cells.forEach(cell => {{
                                    const text = cell.textContent;
                                    const regex = new RegExp(searchTerm, 'gi');
                                    cell.innerHTML = text.replace(regex, match => 
                                        `<span class="highlight">${{match}}</span>`
                                    );
                                }});
                            }} else {{
                                cells.forEach(cell => {{
                                    cell.innerHTML = cell.textContent;
                                }});
                            }}
                        }} else {{
                            row.classList.add('hidden');
                        }}
                    }});
                }}
            }});
        </script>
    </body>
    </html>
    """
    
    # Save to index.html
    with open('index.html', 'w', encoding='utf-8') as f:
        f.write(html_content)

if __name__ == '__main__':
    fetch_and_convert() 