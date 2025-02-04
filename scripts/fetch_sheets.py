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
            :root {{
                --primary-color: #1a73e8;
                --danger-color: #dc3545;
                --warning-color: #ffc107;
            }}
            body {{
                font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif;
                margin: 0;
                padding: 20px;
                line-height: 1.6;
                color: #333;
            }}
            .container {{
                max-width: 1200px;
                margin: 0 auto;
                padding: 0 15px;
            }}
            .controls {{
                margin: 20px 0;
                display: flex;
                gap: 20px;
                flex-wrap: wrap;
            }}
            .search-box {{
                padding: 12px;
                border: 1px solid #ddd;
                border-radius: 8px;
                width: 100%;
                max-width: 400px;
                font-size: 16px;
                transition: all 0.3s ease;
            }}
            .search-box:focus {{
                outline: none;
                border-color: var(--primary-color);
                box-shadow: 0 0 0 2px rgba(26, 115, 232, 0.2);
            }}
            .filter-container {{
                display: flex;
                gap: 15px;
                flex-wrap: wrap;
                width: 100%;
            }}
            .filter-group {{
                flex: 1;
                min-width: 200px;
            }}
            .filter-select {{
                width: 100%;
                padding: 12px;
                border: 1px solid #ddd;
                border-radius: 8px;
                font-size: 16px;
                background-color: white;
                cursor: pointer;
                transition: all 0.3s ease;
            }}
            .filter-select:focus {{
                outline: none;
                border-color: var(--primary-color);
                box-shadow: 0 0 0 2px rgba(26, 115, 232, 0.2);
            }}
            .table-responsive {{
                overflow-x: auto;
                margin: 20px 0;
                border-radius: 8px;
                box-shadow: 0 2px 4px rgba(0,0,0,0.1);
            }}
            table {{
                border-collapse: collapse;
                width: 100%;
                margin: 0;
                background: white;
            }}
            th, td {{
                border: 1px solid #eee;
                padding: 15px;
                text-align: left;
                font-size: 14px;
                white-space: normal;
                word-break: break-word;
            }}
            th {{
                background-color: #f8f9fa;
                font-weight: 600;
                position: sticky;
                top: 0;
                z-index: 10;
            }}
            td {{
                vertical-align: top;
            }}
            tr:nth-child(even) {{
                background-color: #f8f9fa;
            }}
            tr:hover {{
                background-color: #f0f7ff;
            }}
            .last-updated {{
                color: #666;
                font-size: 0.9em;
                margin-bottom: 15px;
            }}
            .timezone-notice {{
                color: var(--danger-color);
                font-size: 0.9em;
                margin-bottom: 20px;
                padding: 10px;
                background-color: rgba(220, 53, 69, 0.1);
                border-radius: 8px;
            }}
            .highlight {{
                background-color: var(--warning-color);
                padding: 2px 4px;
                border-radius: 4px;
            }}
            .hidden {{
                display: none;
            }}
            .stats-container {{
                display: grid;
                grid-template-columns: repeat(auto-fit, minmax(250px, 1fr));
                gap: 20px;
                margin: 20px 0;
            }}
            .stat-card {{
                background: white;
                padding: 20px;
                border-radius: 8px;
                box-shadow: 0 2px 4px rgba(0,0,0,0.1);
            }}
            .stat-value {{
                font-size: 24px;
                font-weight: bold;
                color: var(--primary-color);
            }}
            .stat-label {{
                color: #666;
                font-size: 14px;
            }}
            @media (max-width: 768px) {{
                .controls {{
                    flex-direction: column;
                    gap: 15px;
                }}
                .search-box {{
                    max-width: 100%;
                }}
                .filter-group {{
                    min-width: 100%;
                }}
                /* 移动端表格样式优化 */
                .table-responsive {{
                    margin: 10px -15px;
                    border-radius: 0;
                    box-shadow: none;
                }}
                table {{
                    display: block;
                }}
                table thead {{
                    display: none;
                }}
                table tbody {{
                    display: block;
                }}
                table tr {{
                    display: block;
                    margin: 0 15px 15px;
                    background: white;
                    box-shadow: 0 1px 3px rgba(0,0,0,0.1);
                    border-radius: 8px;
                    position: relative;
                }}
                table td {{
                    display: none;
                    padding: 8px 15px;
                    font-size: 14px;
                    border: none;
                    border-bottom: 1px solid #eee;
                }}
                /* 默认显示的重要信息 */
                table td[data-label="时间戳记"],
                table td[data-label="省份"],
                table td[data-label="城市"],
                table td[data-label="学校名称"],
                table td[data-label="每周在校学习小时数"],
                table td[data-label="24年学生自杀数"] {{
                    display: block;
                }}
                /* 展开/收起按钮 */
                .expand-btn {{
                    position: absolute;
                    right: 15px;
                    top: 50%;
                    transform: translateY(-50%);
                    background: var(--primary-color);
                    color: white;
                    border: none;
                    border-radius: 4px;
                    padding: 6px 12px;
                    font-size: 12px;
                    cursor: pointer;
                    z-index: 1;
                    opacity: 0.9;
                    transition: opacity 0.3s ease;
                }}
                .expand-btn:hover {{
                    opacity: 1;
                }}
                .expanded td {{
                    display: block;
                }}
                /* 优化关键信息显示 */
                table td[data-label="时间戳记"] {{
                    font-weight: 500;
                    background: #f8f9fa;
                    color: var(--primary-color);
                    padding-right: 80px;
                    font-size: 13px;
                }}
                table td[data-label="学校名称"] {{
                    font-size: 16px;
                    font-weight: bold;
                    padding-right: 80px;
                }}
                table td[data-label="每周在校学习小时数"],
                table td[data-label="24年学生自杀数"] {{
                    color: var(--danger-color);
                    font-weight: bold;
                }}
                /* 为每个单元格添加标签 */
                table td:before {{
                    content: attr(data-label);
                    display: block;
                    color: #666;
                    font-size: 12px;
                    font-weight: normal;
                    margin-bottom: 4px;
                }}
                /* 移动端卡片式布局优化 */
                .mobile-card-header {{
                    display: flex;
                    align-items: center;
                    justify-content: space-between;
                    padding-right: 80px;
                }}
                .mobile-card-content {{
                    display: grid;
                    grid-template-columns: repeat(2, 1fr);
                    gap: 8px;
                }}
                @media (max-width: 480px) {{
                    .mobile-card-content {{
                        grid-template-columns: 1fr;
                    }}
                }}
            }}
            /* 添加表格排序样式 */
            .sortable {{
                cursor: pointer;
            }}
            .sortable:after {{
                content: '⇅';
                margin-left: 5px;
                opacity: 0.5;
            }}
            .sortable.asc:after {{
                content: '↑';
                opacity: 1;
            }}
            .sortable.desc:after {{
                content: '↓';
                opacity: 1;
            }}
            /* 学生评论列样式 */
            td[data-label="学生的评论"] {{
                min-width: 200px;
                max-width: 400px;
            }}
        </style>
    </head>
    <body>
        <div class="container">
            <div class="last-updated">最后更新时间：{current_time} (UTC+8)</div>
            <div class="timezone-notice">注意！本站时间与原表格一致，<u>仅最后更新时间</u>为北京时间。</div>
            
            <div class="stats-container" id="totalStatsContainer">
                <!-- 总体统计数据将在这里显示 -->
            </div>
            
            <div class="controls">
                <input type="text" class="search-box" placeholder="输入关键词进行搜索..." id="searchInput">
                <div class="filter-container" id="filterContainer">
                    <!-- 筛选下拉框将通过JavaScript动态添加 -->
                </div>
            </div>

            <div class="stats-container" id="filteredStatsContainer">
                <!-- 筛选后的统计数据将在这里显示 -->
            </div>

            <div class="table-responsive">
                {df.to_html(index=False, classes='table', escape=False)}
            </div>
        </div>

        <script>
            document.addEventListener('DOMContentLoaded', function() {{
                const table = document.querySelector('table');
                const searchInput = document.getElementById('searchInput');
                const filterContainer = document.getElementById('filterContainer');
                const totalStatsContainer = document.getElementById('totalStatsContainer');
                const filteredStatsContainer = document.getElementById('filteredStatsContainer');
                
                // 创建筛选器组
                const headers = Array.from(table.querySelectorAll('th'));
                const originalRows = Array.from(table.querySelectorAll('tr')).slice(1);
                let currentSearchTerm = '';
                
                // 存储所有选项的原始数据
                const originalOptions = {{}};
                
                headers.forEach((header, index) => {{
                    if (['省份', '城市', '区县', '年级'].includes(header.textContent)) {{
                        const values = new Set();
                        Array.from(table.querySelectorAll(`td:nth-child(${{index + 1}})`))
                            .forEach(cell => values.add(cell.textContent.trim()));
                        
                        originalOptions[header.textContent] = Array.from(values).sort();

                        const filterGroup = document.createElement('div');
                        filterGroup.className = 'filter-group';
                        
                        const select = document.createElement('select');
                        select.className = 'filter-select';
                        select.innerHTML = `
                            <option value="">筛选${{header.textContent}}</option>
                            ${{originalOptions[header.textContent].map(value => 
                                `<option value="${{value}}">${{value}}</option>`
                            ).join('')}}
                        `;
                        
                        select.addEventListener('change', () => {{
                            applyFiltersAndSearch();
                        }});
                        
                        filterGroup.appendChild(select);
                        filterContainer.appendChild(filterGroup);
                    }}
                }});

                // 搜索函数
                function searchRows(rows, term) {{
                    if (!term) return rows;
                    
                    return rows.filter(row => {{
                        const cells = Array.from(row.querySelectorAll('td'));
                        const rowText = cells.map(cell => cell.textContent.toLowerCase()).join(' ');
                        return rowText.includes(term.toLowerCase());
                    }});
                }}

                // 筛选函数
                function filterRows(rows) {{
                    const filterValues = Array.from(document.querySelectorAll('.filter-select'))
                        .map(select => ({{
                            name: select.options[0].textContent.replace('筛选', ''),
                            value: select.value
                        }}))
                        .filter(filter => filter.value !== '');

                    if (filterValues.length === 0) return rows;

                    return rows.filter(row => {{
                        const cells = Array.from(row.querySelectorAll('td'));
                        return filterValues.every(filter => {{
                            const cellIndex = headers.findIndex(h => h.textContent === filter.name);
                            return cells[cellIndex].textContent === filter.value;
                        }});
                    }});
                }}

                // 高亮搜索结果
                function highlightSearchTerm(rows, term) {{
                    if (!term) {{
                        rows.forEach(row => {{
                            const cells = Array.from(row.querySelectorAll('td'));
                            cells.forEach(cell => {{
                                cell.innerHTML = cell.textContent;
                            }});
                        }});
                        return;
                    }}

                    rows.forEach(row => {{
                        const cells = Array.from(row.querySelectorAll('td'));
                        cells.forEach(cell => {{
                            const text = cell.textContent;
                            const regex = new RegExp(term, 'gi');
                            cell.innerHTML = text.replace(regex, match => 
                                `<span class="highlight">${{match}}</span>`
                            );
                        }});
                    }});
                }}

                // 更新表格显示
                function updateTableDisplay(visibleRows) {{
                    // 隐藏所有行
                    originalRows.forEach(row => row.classList.add('hidden'));
                    
                    // 显示可见行
                    visibleRows.forEach(row => row.classList.remove('hidden'));
                    
                    // 更新统计
                    updateFilteredStats();
                }}

                // 应用筛选和搜索
                function applyFiltersAndSearch() {{
                    // 第一步：应用搜索
                    let visibleRows = searchRows(originalRows, currentSearchTerm);
                    
                    // 第二步：应用筛选
                    visibleRows = filterRows(visibleRows);
                    
                    // 第三步：更新显示
                    updateTableDisplay(visibleRows);
                    
                    // 第四步：高亮搜索词
                    highlightSearchTerm(visibleRows, currentSearchTerm);
                }}

                // 监听搜索输入
                searchInput.addEventListener('input', (e) => {{
                    currentSearchTerm = e.target.value;
                    applyFiltersAndSearch();
                }});

                // 初始化
                initTotalStats();
                addExpandButtons();
                addMobileDataLabels();
                initTableSort();
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