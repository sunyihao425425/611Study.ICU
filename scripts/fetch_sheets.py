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
                    display: none; /* 隐藏原始表头 */
                }}
                table tbody {{
                    display: block;
                }}
                table tr {{
                    display: block;
                    margin-bottom: 15px;
                    background: white;
                    box-shadow: 0 1px 3px rgba(0,0,0,0.1);
                    border-radius: 8px;
                }}
                table td {{
                    display: block;
                    padding: 10px 15px;
                    font-size: 14px;
                    border: none;
                    border-bottom: 1px solid #eee;
                }}
                table td:last-child {{
                    border-bottom: none;
                }}
                /* 为每个单元格添加标签 */
                table td:before {{
                    content: attr(data-label);
                    display: block;
                    font-weight: bold;
                    color: #666;
                    font-size: 12px;
                    margin-bottom: 4px;
                }}
                /* 优化关键信息显示 */
                table td[data-label="学校名称"] {{
                    font-size: 16px;
                    font-weight: bold;
                    background: #f8f9fa;
                }}
                table td[data-label="每周在校学习小时数"],
                table td[data-label="24年学生自杀数"] {{
                    color: var(--danger-color);
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
        </style>
    </head>
    <body>
        <div class="container">
            <div class="last-updated">最后更新时间：{current_time} (UTC+8)</div>
            <div class="timezone-notice">注意！本站时间与原表格一致，<u>仅最后更新时间</u>为北京时间。</div>
            
            <div class="stats-container" id="statsContainer">
                <!-- 统计卡片将通过JavaScript动态添加 -->
            </div>
            
            <div class="controls">
                <input type="text" class="search-box" placeholder="输入关键词进行搜索..." id="searchInput">
                <div class="filter-container" id="filterContainer">
                    <!-- 筛选下拉框将通过JavaScript动态添加 -->
                </div>
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
                const statsContainer = document.getElementById('statsContainer');
                
                // 创建统计数据
                function updateStats() {{
                    const visibleRows = Array.from(table.querySelectorAll('tr:not(.hidden)')).slice(1);
                    const stats = {{
                        totalSchools: visibleRows.length,
                        avgHours: Math.round(visibleRows.reduce((sum, row) => {{
                            const hours = parseFloat(row.cells[6].textContent || '0');
                            return isNaN(hours) ? sum : sum + hours;
                        }}, 0) / visibleRows.length || 0),
                        totalSuicides: visibleRows.reduce((sum, row) => {{
                            const suicides = row.cells[9].textContent.trim();
                            const num = parseInt(suicides);
                            return isNaN(num) ? sum : sum + num;
                        }}, 0),
                        earlyStart: visibleRows.filter(row => {{
                            const startTime = row.cells[10].textContent;
                            return startTime && (startTime.includes('05:') || startTime.includes('06:'));
                        }}).length
                    }};

                    statsContainer.innerHTML = `
                        <div class="stat-card">
                            <div class="stat-value">${{stats.totalSchools}}</div>
                            <div class="stat-label">符合条件的学校数量</div>
                        </div>
                        <div class="stat-card">
                            <div class="stat-value">${{stats.avgHours}}</div>
                            <div class="stat-label">平均每周在校学习小时数</div>
                        </div>
                        <div class="stat-card">
                            <div class="stat-value">${{stats.totalSuicides || 0}}</div>
                            <div class="stat-label">总自杀人数</div>
                        </div>
                        <div class="stat-card">
                            <div class="stat-value">${{stats.earlyStart}}</div>
                            <div class="stat-label">早于7点上学的学校数</div>
                        </div>
                    `;
                }}

                // 为移动端添加数据标签
                function addMobileDataLabels() {{
                    const rows = Array.from(table.querySelectorAll('tr'));
                    const headers = Array.from(table.querySelectorAll('th')).map(th => th.textContent);
                    
                    rows.slice(1).forEach(row => {{
                        Array.from(row.cells).forEach((cell, index) => {{
                            cell.setAttribute('data-label', headers[index]);
                        }});
                    }});
                }}

                // 添加表格排序功能
                function initTableSort() {{
                    const headers = Array.from(table.querySelectorAll('th'));
                    headers.forEach((header, index) => {{
                        if (['每周在校学习小时数', '每月假期天数', '寒假放假天数', '24年学生自杀数'].includes(header.textContent)) {{
                            header.classList.add('sortable');
                            header.addEventListener('click', () => sortTable(index, header));
                        }}
                    }});
                }}

                function sortTable(columnIndex, header) {{
                    const tbody = table.querySelector('tbody');
                    const rows = Array.from(tbody.querySelectorAll('tr'));
                    const isAsc = !header.classList.contains('asc');
                    
                    // 清除所有排序标记
                    table.querySelectorAll('.sortable').forEach(h => {{
                        h.classList.remove('asc', 'desc');
                    }});
                    
                    // 添加新的排序标记
                    header.classList.add(isAsc ? 'asc' : 'desc');
                    
                    const sortedRows = rows.sort((a, b) => {{
                        const aValue = parseFloat(a.cells[columnIndex].textContent) || 0;
                        const bValue = parseFloat(b.cells[columnIndex].textContent) || 0;
                        return isAsc ? aValue - bValue : bValue - aValue;
                    }});
                    
                    // 重新插入排序后的行
                    tbody.innerHTML = '';
                    sortedRows.forEach(row => tbody.appendChild(row));
                    
                    // 更新统计数据
                    updateStats();
                }}

                // 创建筛选器组
                const headers = Array.from(table.querySelectorAll('th'));
                headers.forEach((header, index) => {{
                    const values = new Set();
                    Array.from(table.querySelectorAll(`td:nth-child(${{index + 1}})`))
                        .forEach(cell => values.add(cell.textContent.trim()));

                    // 只为有意义的列创建筛选器
                    if (['省份', '城市', '区县', '年级'].includes(header.textContent)) {{
                        const filterGroup = document.createElement('div');
                        filterGroup.className = 'filter-group';
                        
                        const select = document.createElement('select');
                        select.className = 'filter-select';
                        select.innerHTML = `
                            <option value="">筛选${{header.textContent}}</option>
                            ${{Array.from(values).sort().map(value => 
                                `<option value="${{value}}">${{value}}</option>`
                            ).join('')}}
                        `;
                        
                        select.addEventListener('change', () => {{
                            applyFilters();
                            // 连锁效应：更新其他筛选器的可选项
                            updateFilterOptions(index);
                        }});
                        
                        filterGroup.appendChild(select);
                        filterContainer.appendChild(filterGroup);
                    }}
                }});

                // 更新筛选器选项（连锁效应）
                function updateFilterOptions(changedIndex) {{
                    const filters = Array.from(document.querySelectorAll('.filter-select'));
                    const activeFilters = filters.map(select => ({{
                        index: headers.findIndex(h => h.textContent === select.options[0].textContent.replace('筛选', '')),
                        value: select.value.toLowerCase()
                    }}));

                    // 获取当前可见的行
                    const visibleRows = Array.from(table.querySelectorAll('tr:not(.hidden)')).slice(1);

                    filters.forEach((select, filterIndex) => {{
                        if (filterIndex === changedIndex) return; // 跳过触发变化的筛选器

                        const columnIndex = headers.findIndex(h => 
                            h.textContent === select.options[0].textContent.replace('筛选', ''));
                        
                        // 收集可用的选项
                        const availableValues = new Set();
                        visibleRows.forEach(row => {{
                            availableValues.add(row.cells[columnIndex].textContent.trim());
                        }});

                        // 保存当前选中的值
                        const currentValue = select.value;
                        
                        // 更新选项
                        select.innerHTML = `
                            <option value="">筛选${{headers[columnIndex].textContent}}</option>
                            ${{Array.from(availableValues).sort().map(value => 
                                `<option value="${{value}}" ${{value === currentValue ? 'selected' : ''}}>${{value}}</option>`
                            ).join('')}}
                        `;
                    }});
                }}

                // 应用筛选和搜索
                function applyFilters() {{
                    const searchTerm = searchInput.value.toLowerCase();
                    const filterValues = Array.from(document.querySelectorAll('.filter-select'))
                        .map(select => ({{
                            index: headers.findIndex(h => h.textContent === select.options[0].textContent.replace('筛选', '')),
                            value: select.value.toLowerCase()
                        }}));

                    const rows = Array.from(table.querySelectorAll('tr'));
                    
                    // 跳过表头行
                    rows.slice(1).forEach(row => {{
                        const cells = Array.from(row.querySelectorAll('td'));
                        const rowText = cells.map(cell => cell.textContent.toLowerCase()).join(' ');
                        
                        // 检查是否匹配搜索词和所有激活的筛选器
                        const matchesSearch = searchTerm === '' || rowText.includes(searchTerm);
                        const matchesFilters = filterValues.every(filter => 
                            filter.value === '' || 
                            cells[filter.index].textContent.toLowerCase() === filter.value
                        );

                        if (matchesSearch && matchesFilters) {{
                            row.classList.remove('hidden');
                            // 高亮搜索词
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

                    // 更新统计数据
                    updateStats();
                }}

                // 监听搜索输入
                searchInput.addEventListener('input', applyFilters);

                // 初始化
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