import pandas as pd
from datetime import datetime
import pytz
import subprocess
import os

# CSV download URL
CSV_URL = 'https://docs.google.com/spreadsheets/d/1P48quxwMv9XsYQhXjLOvTRRq8tt3ahJnkbXo4VCxjLc/export?format=csv'

def fetch_and_convert():
    # Download CSV file using wget
    subprocess.run(['wget', '-q', CSV_URL, '-O', 'data.csv'], check=True)
    
    # Read CSV file
    df = pd.read_csv('data.csv')
    
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
                display: flex;
                flex-direction: column;
                justify-content: center;
            }}
            .stat-value {{
                font-size: 24px;
                font-weight: bold;
                color: var(--primary-color);
                margin-bottom: 8px;
            }}
            .stat-label {{
                color: #666;
                font-size: 14px;
            }}
            .comment-cell {{
                max-width: 300px;
                white-space: pre-wrap;
                word-break: break-word;
            }}
            .time-cell {{
                white-space: nowrap;
            }}
            .fee-cell {{
                text-align: right;
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
                table td[data-label="24年学生自杀数"],
                table td[data-label="放学时间"],
                table td[data-label="寒假补课收费总价格"],
                table td[data-label="学生的评论"] {{
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
            <h1>611Study.ICU - 全国超时学习学校耻辱名单</h1>
            <p class="last-updated">最后更新时间：{current_time} (UTC+8)</p>
            <p class="timezone-notice">⚠️ 注意：所有时间均为24小时制</p>
            
            <div class="controls">
                <input type="text" id="searchInput" class="search-box" placeholder="搜索学校、地区或其他信息...">
                <div class="filter-container" id="filterContainer">
                    <!-- 筛选器将通过JavaScript动态生成 -->
                </div>
            </div>

            <div class="stats-container" id="totalStatsContainer">
                <!-- 总体统计将通过JavaScript动态生成 -->
            </div>

            <div class="stats-container" id="filteredStatsContainer">
                <!-- 筛选后统计将通过JavaScript动态生成 -->
            </div>

            <div class="table-responsive">
                <table>
                    <thead>
                        <tr>
                            <th>时间戳记</th>
                            <th>省份</th>
                            <th>城市</th>
                            <th>区县</th>
                            <th>学校名称</th>
                            <th>年级</th>
                            <th class="sortable">每周在校学习小时数</th>
                            <th class="sortable">每月假期天数</th>
                            <th class="sortable">寒假放假天数</th>
                            <th class="sortable">24年学生自杀数</th>
                            <th>上学时间</th>
                            <th>放学时间</th>
                            <th class="sortable">寒假补课收费总价格</th>
                            <th>学生的评论</th>
                        </tr>
                    </thead>
                    <tbody>
"""

    # Process each row
    for _, row in df.iterrows():
        # Format times
        start_time = row.get('上学时间', '')
        end_time = row.get('放学时间', '')
        fee = row.get('寒假补课收费总价格', 0)
        comment = row.get('学生的评论', '')
        
        html_content += f"""
        <tr>
            <td class="time-cell">{row['时间戳记']}</td>
            <td>{row['省份']}</td>
            <td>{row['城市']}</td>
            <td>{row['区县']}</td>
            <td>{row['学校名称']}</td>
            <td>{row['年级']}</td>
            <td>{row['每周在校学习小时数']}</td>
            <td>{row['每月假期天数']}</td>
            <td>{row['寒假放假天数']}</td>
            <td>{row['24年学生自杀数']}</td>
            <td class="time-cell">{start_time}</td>
            <td class="time-cell">{end_time}</td>
            <td class="fee-cell">{fee if pd.notna(fee) else '0'}</td>
            <td class="comment-cell">{comment if pd.notna(comment) else ''}</td>
        </tr>
"""

    html_content += """
                    </tbody>
                </table>
            </div>
        </div>

        <script>
            document.addEventListener('DOMContentLoaded', function() {
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
                const originalOptions = {};
                
                headers.forEach((header, index) => {
                    if (['省份', '城市', '区县', '年级'].includes(header.textContent)) {
                        const values = new Set();
                        Array.from(table.querySelectorAll(`td:nth-child(${index + 1})`))
                            .forEach(cell => values.add(cell.textContent.trim()));
                        
                        originalOptions[header.textContent] = Array.from(values).sort();

                        const filterGroup = document.createElement('div');
                        filterGroup.className = 'filter-group';
                        
                        const select = document.createElement('select');
                        select.className = 'filter-select';
                        select.innerHTML = `
                            <option value="">筛选${header.textContent}</option>
                            ${{originalOptions[header.textContent].map(value => 
                                `<option value="${value}">${value}</option>`
                            ).join('')}}
                        `;
                        
                        select.addEventListener('change', () => {
                            applyFiltersAndSearch();
                        });
                        
                        filterGroup.appendChild(select);
                        filterContainer.appendChild(filterGroup);
                    }
                });

                // 搜索函数
                function searchRows(rows, term) {
                    if (!term) return rows;
                    
                    return rows.filter(row => {
                        const cells = Array.from(row.querySelectorAll('td'));
                        const rowText = cells.map(cell => cell.textContent.toLowerCase()).join(' ');
                        return rowText.includes(term.toLowerCase());
                    });
                }

                // 筛选函数
                function filterRows(rows) {
                    const filterValues = Array.from(document.querySelectorAll('.filter-select'))
                        .map(select => ({{
                            name: select.options[0].textContent.replace('筛选', ''),
                            value: select.value
                        }}))
                        .filter(filter => filter.value !== '');

                    if (filterValues.length === 0) return rows;

                    return rows.filter(row => {
                        const cells = Array.from(row.querySelectorAll('td'));
                        return filterValues.every(filter => {
                            const cellIndex = headers.findIndex(h => h.textContent === filter.name);
                            return cells[cellIndex].textContent === filter.value;
                        });
                    });
                }

                // 高亮搜索结果
                function highlightSearchTerm(rows, term) {
                    if (!term) {
                        rows.forEach(row => {
                            const cells = Array.from(row.querySelectorAll('td'));
                            cells.forEach(cell => {
                                cell.innerHTML = cell.textContent;
                            });
                        });
                        return;
                    }

                    rows.forEach(row => {
                        const cells = Array.from(row.querySelectorAll('td'));
                        cells.forEach(cell => {
                            const text = cell.textContent;
                            const regex = new RegExp(term, 'gi');
                            cell.innerHTML = text.replace(regex, match => 
                                `<span class="highlight">${match}</span>`
                            );
                        });
                    });
                }

                // 更新表格显示
                function updateTableDisplay(visibleRows) {
                    // 隐藏所有行
                    originalRows.forEach(row => row.classList.add('hidden'));
                    
                    // 显示可见行
                    visibleRows.forEach(row => row.classList.remove('hidden'));
                    
                    // 更新统计
                    updateFilteredStats();
                }

                // 应用筛选和搜索
                function applyFiltersAndSearch() {
                    // 第一步：应用搜索
                    let visibleRows = searchRows(originalRows, currentSearchTerm);
                    
                    // 第二步：应用筛选
                    visibleRows = filterRows(visibleRows);
                    
                    // 第三步：更新显示
                    updateTableDisplay(visibleRows);
                    
                    // 第四步：高亮搜索词
                    highlightSearchTerm(visibleRows, currentSearchTerm);
                }

                // 监听搜索输入
                searchInput.addEventListener('input', (e) => {
                    currentSearchTerm = e.target.value;
                    applyFiltersAndSearch();
                });

                // 计算统计数据
                function calculateStats(rows) {
                    return {{
                        totalSchools: rows.length,
                        avgHours: Math.round(rows.reduce((sum, row) => {{
                            const hours = parseFloat(row.cells[6].textContent || '0');
                            return isNaN(hours) ? sum : sum + hours;
                        }}, 0) / rows.length || 0),
                        totalSuicides: rows.reduce((sum, row) => {{
                            const suicides = row.cells[9].textContent.trim();
                            const num = parseInt(suicides);
                            return isNaN(num) ? sum : sum + num;
                        }}, 0),
                        avgFee: Math.round(rows.reduce((sum, row) => {{
                            const fee = parseFloat(row.cells[12].textContent || '0');
                            return isNaN(fee) ? sum : sum + fee;
                        }}, 0) / rows.length || 0),
                        earlyStart: rows.filter(row => {{
                            const startTime = row.cells[10].textContent;
                            return startTime && (startTime.includes('05:') || startTime.includes('06:'));
                        }}).length,
                        lateEnd: rows.filter(row => {{
                            const endTime = row.cells[11].textContent;
                            return endTime && (endTime.includes('21:') || endTime.includes('22:') || endTime.includes('23:'));
                        }}).length
                    }};
                }

                function createStatsHtml(stats, isFiltered) {
                    return `
                        <div class="stat-card">
                            <div class="stat-value">${stats.totalSchools}</div>
                            <div class="stat-label">${isFiltered ? '筛选后学校数' : '总学校数'}</div>
                        </div>
                        <div class="stat-card">
                            <div class="stat-value">${stats.avgHours}</div>
                            <div class="stat-label">平均每周学习小时数</div>
                        </div>
                        <div class="stat-card">
                            <div class="stat-value">${stats.totalSuicides}</div>
                            <div class="stat-label">2024年学生自杀总数</div>
                        </div>
                        <div class="stat-card">
                            <div class="stat-value">${stats.avgFee}</div>
                            <div class="stat-label">平均寒假补课费用</div>
                        </div>
                        <div class="stat-card">
                            <div class="stat-value">${stats.earlyStart}</div>
                            <div class="stat-label">早于7点上学的学校数</div>
                        </div>
                        <div class="stat-card">
                            <div class="stat-value">${stats.lateEnd}</div>
                            <div class="stat-label">晚于9点放学的学校数</div>
                        </div>
                    `;
                }

                // 初始化总体统计
                function initTotalStats() {
                    const totalStats = calculateStats(originalRows);
                    totalStatsContainer.innerHTML = createStatsHtml(totalStats, false);
                }

                // 更新筛选后的统计
                function updateFilteredStats() {
                    const visibleRows = Array.from(table.querySelectorAll('tr:not(.hidden)')).slice(1);
                    const filteredStats = calculateStats(visibleRows);
                    filteredStatsContainer.innerHTML = createStatsHtml(filteredStats, true);
                }

                // 为移动端添加数据标签
                function addMobileDataLabels() {
                    const rows = Array.from(table.querySelectorAll('tr'));
                    const headers = Array.from(table.querySelectorAll('th')).map(th => th.textContent);
                    
                    rows.slice(1).forEach(row => {
                        Array.from(row.cells).forEach((cell, index) => {
                            cell.setAttribute('data-label', headers[index]);
                        });
                    });
                }

                // 为移动端添加展开/收起功能
                function addExpandButtons() {
                    const rows = Array.from(table.querySelectorAll('tr')).slice(1);
                    rows.forEach(row => {
                        // 创建展开按钮
                        const expandBtn = document.createElement('button');
                        expandBtn.className = 'expand-btn';
                        expandBtn.textContent = '展开详情';
                        expandBtn.onclick = function(e) {
                            e.stopPropagation();
                            const isExpanded = row.classList.toggle('expanded');
                            this.textContent = isExpanded ? '收起' : '展开详情';
                        };
                        
                        // 添加按钮到行
                        if (!row.querySelector('.expand-btn')) {
                            row.appendChild(expandBtn);
                        }
                        
                        // 添加行点击事件
                        row.style.cursor = 'pointer';
                        row.onclick = function() {
                            const isExpanded = this.classList.toggle('expanded');
                            const btn = this.querySelector('.expand-btn');
                            if (btn) {
                                btn.textContent = isExpanded ? '收起' : '展开详情';
                            }
                        };
                    });
                }

                // 添加表格排序功能
                function initTableSort() {
                    const headers = Array.from(table.querySelectorAll('th'));
                    headers.forEach((header, index) => {
                        if (['每周在校学习小时数', '每月假期天数', '寒假放假天数', '24年学生自杀数'].includes(header.textContent)) {
                            header.classList.add('sortable');
                            header.addEventListener('click', () => sortTable(index, header));
                        }
                    });
                }

                function sortTable(columnIndex, header) {
                    const tbody = table.querySelector('tbody');
                    const rows = Array.from(tbody.querySelectorAll('tr'));
                    const isAsc = !header.classList.contains('asc');
                    
                    // 清除所有排序标记
                    table.querySelectorAll('.sortable').forEach(h => {
                        h.classList.remove('asc', 'desc');
                    });
                    
                    // 添加新的排序标记
                    header.classList.add(isAsc ? 'asc' : 'desc');
                    
                    const sortedRows = rows.sort((a, b) => {
                        const aValue = parseFloat(a.cells[columnIndex].textContent) || 0;
                        const bValue = parseFloat(b.cells[columnIndex].textContent) || 0;
                        return isAsc ? aValue - bValue : bValue - aValue;
                    });
                    
                    // 重新插入排序后的行
                    tbody.innerHTML = '';
                    sortedRows.forEach(row => tbody.appendChild(row));
                    
                    // 更新统计数据
                    updateFilteredStats();
                }

                // 初始化
                initTotalStats();
                addExpandButtons();
                addMobileDataLabels();
                initTableSort();
            });
        </script>
    </body>
    </html>
"""

    # Write the HTML file
    with open('index.html', 'w', encoding='utf-8') as f:
        f.write(html_content)

if __name__ == '__main__':
    fetch_and_convert()
    # Clean up CSV file after conversion
    if os.path.exists('data.csv'):
        os.remove('data.csv') 