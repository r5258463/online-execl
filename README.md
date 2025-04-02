html
<!DOCTYPE html>
<html lang="zh-TW">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>簡易線上 Excel</title>
    <style>
        body {
            font-family: Arial, sans-serif;
            margin: 20px;
        }
        #excel-container {
            overflow: auto;
            margin-bottom: 20px;
        }
        table {
            border-collapse: collapse;
        }
        th, td {
            border: 1px solid #ddd;
            padding: 8px;
            min-width: 100px;
            height: 25px;
        }
        th {
            background-color: #f2f2f2;
            text-align: center;
            position: sticky;
            top: 0;
        }
        td:first-child, th:first-child {
            position: sticky;
            left: 0;
            background-color: #f2f2f2;
            z-index: 2;
        }
        th:first-child {
            z-index: 3;
        }
        td input {
            width: 100%;
            height: 100%;
            border: none;
            outline: none;
            box-sizing: border-box;
        }
        .selected {
            background-color: #d4e6f7;
        }
        .controls {
            margin-bottom: 15px;
        }
        button {
            padding: 5px 10px;
            margin-right: 5px;
        }
    </style>
</head>
<body>
    <h1>簡易線上 Excel</h1>
    <div class="controls">
        <button id="add-row">新增行</button>
        <button id="add-col">新增列</button>
        <button id="save">儲存為 CSV</button>
        <button id="load">載入 CSV</button>
        <input type="file" id="file-input" accept=".csv" style="display: none;">
    </div>
    <div id="excel-container">
        <table id="excel-table">
            <thead>
                <tr id="header-row">
                    <th></th> <!-- 左上角空白格 -->
                </tr>
            </thead>
            <tbody id="table-body">
            </tbody>
        </table>
    </div>

    <script>
        // 初始化表格 (10行 x 10列)
        const ROWS = 10;
        const COLS = 10;
        let selectedCell = null;

        // 生成表格
        function initTable() {
            const tableBody = document.getElementById('table-body');
            const headerRow = document.getElementById('header-row');

            // 生成表頭 (A, B, C...)
            for (let col = 0; col < COLS; col++) {
                const th = document.createElement('th');
                th.textContent = String.fromCharCode(65 + col);
                headerRow.appendChild(th);
            }

            // 生成表格內容
            for (let row = 0; row < ROWS; row++) {
                const tr = document.createElement('tr');
                
                // 行號單元格
                const th = document.createElement('th');
                th.textContent = row + 1;
                tr.appendChild(th);

                // 數據單元格
                for (let col = 0; col < COLS; col++) {
                    const td = document.createElement('td');
                    const input = document.createElement('input');
                    input.type = 'text';
                    input.dataset.row = row;
                    input.dataset.col = col;
                    
                    input.addEventListener('focus', (e) => {
                        if (selectedCell) {
                            selectedCell.classList.remove('selected');
                        }
                        selectedCell = e.target.parentElement;
                        selectedCell.classList.add('selected');
                    });
                    
                    td.appendChild(input);
                    tr.appendChild(td);
                }
                tableBody.appendChild(tr);
            }
        }

        // 新增行
        document.getElementById('add-row').addEventListener('click', () => {
            const tableBody = document.getElementById('table-body');
            const rowCount = tableBody.rows.length;
            const colCount = tableBody.rows[0]?.cells.length || 1;

            const tr = document.createElement('tr');
            const th = document.createElement('th');
            th.textContent = rowCount + 1;
            tr.appendChild(th);

            for (let col = 1; col < colCount; col++) {
                const td = document.createElement('td');
                const input = document.createElement('input');
                input.type = 'text';
                input.dataset.row = rowCount;
                input.dataset.col = col - 1;
                
                input.addEventListener('focus', (e) => {
                    if (selectedCell) {
                        selectedCell.classList.remove('selected');
                    }
                    selectedCell = e.target.parentElement;
                    selectedCell.classList.add('selected');
                });
                
                td.appendChild(input);
                tr.appendChild(td);
            }
            tableBody.appendChild(tr);
        });

        // 新增列
        document.getElementById('add-col').addEventListener('click', () => {
            const headerRow = document.getElementById('header-row');
            const tableBody = document.getElementById('table-body');
            const colCount = headerRow.cells.length;
            const colChar = String.fromCharCode(65 + colCount - 1);

            // 更新表頭
            const th = document.createElement('th');
            th.textContent = String.fromCharCode(65 + colCount - 1);
            headerRow.appendChild(th);

            // 更新每行
            Array.from(tableBody.rows).forEach((row, rowIndex) => {
                const td = document.createElement('td');
                const input = document.createElement('input');
                input.type = 'text';
                input.dataset.row = rowIndex;
                input.dataset.col = colCount - 1;
                
                input.addEventListener('focus', (e) => {
                    if (selectedCell) {
                        selectedCell.classList.remove('selected');
                    }
                    selectedCell = e.target.parentElement;
                    selectedCell.classList.add('selected');
                });
                
                td.appendChild(input);
                row.appendChild(td);
            });
        });

        // 儲存為 CSV
        document.getElementById('save').addEventListener('click', () => {
            const tableBody = document.getElementById('table-body');
            const rows = Array.from(tableBody.rows);
            let csvContent = "";

            rows.forEach(row => {
                const cells = Array.from(row.cells).slice(1); // 跳過行號
                const rowData = cells.map(cell => {
                    const input = cell.querySelector('input');
                    return input ? `"${input.value.replace(/"/g, '""')}"` : '""';
                }).join(',');
                csvContent += rowData + '\n';
            });

            const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
            const url = URL.createObjectURL(blob);
            const link = document.createElement('a');
            link.href = url;
            link.download = 'spreadsheet.csv';
            link.click();
        });

        // 載入 CSV
        document.getElementById('load').addEventListener('click', () => {
            document.getElementById('file-input').click();
        });

        document.getElementById('file-input').addEventListener('change', (e) => {
            const file = e.target.files[0];
            if (!file) return;

            const reader = new FileReader();
            reader.onload = (event) => {
                const content = event.target.result;
                const rows = content.split('\n').filter(row => row.trim() !== '');
                const tableBody = document.getElementById('table-body');
                
                // 清空現有表格 (保留標頭)
                tableBody.innerHTML = '';

                // 重新填充數據
                rows.forEach((row, rowIndex) => {
                    const cells = row.split(/,(?=(?:(?:[^"]*"){2})*[^"]*$)/);
                    const tr = document.createElement('tr');
                    const th = document.createElement('th');
                    th.textContent = rowIndex + 1;
                    tr.appendChild(th);

                    cells.forEach((cell, colIndex) => {
                        const td = document.createElement('td');
                        const input = document.createElement('input');
                        input.type = 'text';
                        input.value = cell.replace(/^"|"$/g, '').replace(/""/g, '"');
                        input.dataset.row = rowIndex;
                        input.dataset.col = colIndex;
                        
                        input.addEventListener('focus', (e) => {
                            if (selectedCell) {
                                selectedCell.classList.remove('selected');
                            }
                            selectedCell = e.target.parentElement;
                            selectedCell.classList.add('selected');
                        });
                        
                        td.appendChild(input);
                        tr.appendChild(td);
                    });
                    tableBody.appendChild(tr);
                });
            };
            reader.readAsText(file);
        });

        // 初始化
        initTable();
    </script>
</body>
</html># online-execl
