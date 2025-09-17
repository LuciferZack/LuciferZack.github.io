// ======================
// 全局变量
// ======================
let rawData = [];
let cleanedData = [];
let filteredData = [];
let personList = [];
let applicantList = [];

// ======================
// 页面导航功能
// ======================
document.addEventListener('DOMContentLoaded', function() {
    // Update this section to work with topbar instead of sidebar
    const pageLinks = document.querySelectorAll('.topbar-menu a');
    const pages = document.querySelectorAll('.page');
    
    pageLinks.forEach(link => {
        link.addEventListener('click', function(e) {
            e.preventDefault();
            
            // Update active link
            pageLinks.forEach(l => l.classList.remove('active'));
            this.classList.add('active');
            
            // Show corresponding page
            const pageId = this.getAttribute('data-page');
            
            // If no data-page attribute, determine page by text content
            let targetPageId = pageId;
            if (!targetPageId) {
                const linkText = this.textContent.trim();
                switch(linkText) {
                    case 'Dashboard':
                        targetPageId = 'homepage';
                        break;
                    case 'Forecast':
                        targetPageId = 'forecast';
                        break;
                    case 'Cashflow':
                        targetPageId = 'cashflow';
                        break;
                    case 'Reports':
                        targetPageId = 'dataIntegration';
                        break;
                    default:
                        targetPageId = 'homepage';
                }
            }
            
            pages.forEach(page => page.classList.remove('active'));
            document.getElementById(targetPageId).classList.add('active');
            
            // If navigating to forecast page, initialize it
            if (targetPageId === 'forecast') {
                // Use setTimeout to ensure the page is fully rendered
                setTimeout(function() {
                    if (typeof initForecastPage === 'function') {
                        initForecastPage();
                    }
                    
                    // If we have forecast data, display it
                    if (typeof forecastData !== 'undefined' && forecastData.length > 0) {
                        if (typeof displayForecastTable === 'function') {
                            displayForecastTable(forecastData);
                        }
                    }
                }, 100);
            }
        });
    });

    // Rest of your existing code...

    // ======================
    // 首页功能
    // ======================
    // 文件上传处理
    document.getElementById('fileInput').addEventListener('change', function(e) {
        const file = e.target.files[0];
        if (!file) return;

        const reader = new FileReader();
        reader.onload = function(e) {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
            rawData = XLSX.utils.sheet_to_json(firstSheet, { header: 1 });

            document.getElementById('previewBtn').disabled = false;
            updateStatus(`Files Uploaded: ${file.name}`, 'success');

            // 初始化人员筛选列表
            initFilterLists();
        };
        reader.readAsArrayBuffer(file);
    });

    // 初始化筛选列表
    function initFilterLists() {
        if (rawData.length < 2) return;

        const headers = rawData[0];
        const personIndex = headers.indexOf('Person Respons.');
        const applicantIndex = headers.indexOf('Applicant');

        if (personIndex === -1 || applicantIndex === -1) return;

        // 收集所有负责人和申请人
        const personSet = new Set();
        const applicantSet = new Set();

        for (let i = 1; i < rawData.length; i++) {
            const row = rawData[i];
            if (row[personIndex]) personSet.add(row[personIndex]);
            if (row[applicantIndex]) applicantSet.add(row[applicantIndex]);
        }

        personList = Array.from(personSet).sort();
        applicantList = Array.from(applicantSet).sort();

        // 填充下拉列表
        const personFilter = document.getElementById('personFilter');
        const applicantFilter = document.getElementById('applicantFilter');
        
        personFilter.innerHTML = '<option value="">-- PM Filter --</option>';
        applicantFilter.innerHTML = '<option value="">-- CPM Filter --</option>';

        personList.forEach(person => {
            personFilter.innerHTML += `<option value="${person}">${person}</option>`;
        });

        applicantList.forEach(applicant => {
            applicantFilter.innerHTML += `<option value="${applicant}">${applicant}</option>`;
        });
    }

    // 预览原始数据
    document.getElementById('previewBtn').addEventListener('click', function() {
        if (rawData.length === 0) return;

        displayTable(document.getElementById('rawTable'), rawData);
        document.getElementById('cleanBtn').disabled = false;
        document.getElementById('filterSection').style.display = 'block';

        updateStatus('Original Data loaded, Ready to Cleaning', 'success');
    });

    // 执行数据清洗
    document.getElementById('cleanBtn').addEventListener('click', function() {
        if (rawData.length === 0) return;

        try {
            cleanedData = cleanData(rawData);
            displayTable(document.getElementById('cleanedTable'), cleanedData);

            // 默认筛选后数据与清洗后数据相同
            filteredData = [...cleanedData];
            displayTable(document.getElementById('filteredTable'), filteredData);

            document.getElementById('downloadBtn').disabled = false;
            document.getElementById('startForecastBtn').disabled = false;
            updateStatus('Data cleaning completed! Filtering or downloading data', 'success');
        } catch (error) {
            updateStatus(`error during the cleaning process: ${error.message}`, 'error');
            console.error(error);
        }
    });

    // 应用筛选
    document.getElementById('applyFilterBtn').addEventListener('click', function() {
        if (cleanedData.length === 0) return;

        const selectedPerson = document.getElementById('personFilter').value;
        const selectedApplicant = document.getElementById('applicantFilter').value;

        // 如果没有选择任何筛选条件，显示所有数据
        if (!selectedPerson && !selectedApplicant) {
            filteredData = [...cleanedData];
        } else {
            const headers = cleanedData[0];
            const personIndex = headers.indexOf('Person Respons.');
            const applicantIndex = headers.indexOf('Applicant');

            filteredData = [cleanedData[0]]; // 保留标题行

            for (let i = 1; i < cleanedData.length; i++) {
                const row = cleanedData[i];
                const personMatch = !selectedPerson || row[personIndex] === selectedPerson;
                const applicantMatch = !selectedApplicant || row[applicantIndex] === selectedApplicant;

                if (personMatch && applicantMatch) {
                    filteredData.push(row);
                }
            }
        }

        displayTable(document.getElementById('filteredTable'), filteredData);

        updateStatus(`Filter Completed， ${filteredData.length - 1} pieces of Data in total`, 'success');
    });

    // 重置筛选
    document.getElementById('resetFilterBtn').addEventListener('click', function() {
        document.getElementById('personFilter').value = '';
        document.getElementById('applicantFilter').value = '';
        filteredData = [...cleanedData];
        displayTable(document.getElementById('filteredTable'), filteredData);

        updateStatus('The filtering conditions have been reset', 'success');
    });

// 开始预测
document.getElementById('startForecastBtn').addEventListener('click', function() {
    if (filteredData.length === 0) return;

    try {
        forecastData = prepareForecastData(filteredData);
        
        // 切换到预测页面 (updated for topbar)
        const forecastLink = document.querySelector('.topbar-menu a[data-page="forecast"]');
        if (forecastLink) {
            forecastLink.click();
        } else {
            // Fallback: find by text content
            const allLinks = document.querySelectorAll('.topbar-menu a');
            allLinks.forEach(link => {
                if (link.textContent.trim() === 'Forecast') {
                    link.click();
                }
            });
        }
        
        // After switching to forecast page, display the forecast table
        // Use setTimeout to ensure the page switch is complete
        setTimeout(function() {
            if (forecastData.length > 0) {
                displayForecastTable(forecastData);
            }
        }, 100);
        
        updateStatus('The prediction table is ready and can be edited', 'success');
    } catch (error) {
        updateStatus(`An error occurred when preparing to predict the data: ${error.message}`, 'error');
        console.error(error);
    }
});

    // 下载最终数据
    document.getElementById('downloadBtn').addEventListener('click', function() {
        const dataToDownload = filteredData.length > 0 ? filteredData : cleanedData;
        if (dataToDownload.length === 0) return;

        const worksheet = XLSX.utils.aoa_to_sheet(dataToDownload);
        const workbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(workbook, worksheet, "Processed Data");

        // Get selected CPM and PM values
        const selectedCPM = document.getElementById('applicantFilter').value || 'All';
        const selectedPM = document.getElementById('personFilter').value || 'All';

        XLSX.writeFile(workbook, `Filtered_Data_CPM_${selectedCPM}_PM_${selectedPM}.xlsx`);

        updateStatus('Download started', 'success');
    });

    // 初始化标签页
    const tabButtons = document.querySelectorAll('.tab-button');
    const tabContents = document.querySelectorAll('.tab-content');
    
    tabButtons.forEach(button => {
        button.addEventListener('click', function() {
            const tabId = this.getAttribute('data-tab');

            tabButtons.forEach(btn => btn.classList.remove('active'));
            tabContents.forEach(content => content.classList.remove('active'));

            this.classList.add('active');
            document.getElementById(tabId).classList.add('active');
        });
    });
});

// ======================
// 通用功能函数
// ======================
function updateStatus(message, type) {
    const statusElement = document.getElementById('statusMessage');
    statusElement.textContent = message;
    statusElement.style.color = type === 'error' ? 'var(--danger-color)' : 'var(--success-color)';
}

// 显示表格数据
function displayTable(tableElement, data) {
    tableElement.innerHTML = '';

    if (data.length === 0) return;

    // 创建表头
    const thead = document.createElement('thead');
    const headerRow = document.createElement('tr');

    data[0].forEach(headerText => {
        const th = document.createElement('th');
        th.textContent = headerText;
        headerRow.appendChild(th);
    });

    thead.appendChild(headerRow);
    tableElement.appendChild(thead);

    // 创建表格内容
    const tbody = document.createElement('tbody');
    for (let i = 1; i < data.length; i++) {
        const row = document.createElement('tr');

        data[i].forEach(cellData => {
            const td = document.createElement('td');
            td.textContent = cellData !== undefined ? cellData : '';
            row.appendChild(td);
        });

        tbody.appendChild(row);
    }

    tableElement.appendChild(tbody);
}

// Data cleaning logic
function cleanData(data) {
    if (data.length === 0) return [];

    const headers = data[0];
    const wbsLevelIndex = headers.indexOf('WBS Level');
    const statusIndex = headers.indexOf('Status');
    const segmentIndex = headers.indexOf('Segment');
    const wbsElementIndex = headers.indexOf('WBS Element');
    const pocIndex = headers.indexOf('POC%');
    const planMarginIndex = headers.indexOf('Plan Margin%');
    const actualMarginIndex = headers.indexOf('Actual Margin%');

    // First step: basic filtering
    let filteredRows = [];

    // Add new header row (insert Project Number column)
    const newHeaders = ['Project Number', ...headers];
    filteredRows.push(newHeaders);

    // Collect all .POC project POC% values
    const pocValues = {};
    for (let i = 1; i < data.length; i++) {
        const row = data[i];
        const wbsElement = row[wbsElementIndex] || '';

        if (wbsElement.endsWith('.POC')) {
            const baseProjNumber = wbsElement.split('.')[0];
            pocValues[baseProjNumber] = row[pocIndex];
        }
    }

    // Second step: row filtering
    for (let i = 1; i < data.length; i++) {
        const row = data[i];

        // Skip rows with Status TECO
        if (row[statusIndex] === 'TECO') continue;

        // Only keep rows with Segment SO or MVD
        if (row[segmentIndex] !== 'SO' && row[segmentIndex] !== 'MVD') continue;

        // Add Project Number column
        let projectNumber = row[wbsElementIndex] || '';
        if (projectNumber.includes('.')) {
            projectNumber = projectNumber.split('.')[0];
        }

        // Create new row - preserve all decimal places for margin columns
        const newRow = [projectNumber];
        for (let j = 0; j < row.length; j++) {
            const value = row[j];
            // Check if current column is one of the margin columns
            if (j === planMarginIndex || j === actualMarginIndex) {
                // Keep original value without rounding
                newRow.push(value);
            } else if (typeof value === 'number') {
                // Round other numeric values to 2 decimal places
                newRow.push(parseFloat(value.toFixed(2)));
            } else {
                newRow.push(value);
            }
        }

        if (!row[wbsElementIndex].endsWith('.POC') && pocValues[projectNumber]) {
            newRow[pocIndex + 1] = pocValues[projectNumber]; // +1 because we added Project Number column
        }

        filteredRows.push(newRow);
    }

    // Third step: filter by WBS Level
    const finalRows = [filteredRows[0]]; // Keep header row

    for (let i = 1; i < filteredRows.length; i++) {
        const currentRow = filteredRows[i];
        const currentLevel = parseInt(currentRow[wbsLevelIndex + 1]); // +1 because we added Project Number column

        // Always keep level 1 rows
        if (currentLevel === 1) {
            finalRows.push(currentRow);
            continue;
        }

        // Check if next row exists
        if (i < filteredRows.length - 1) {
            const nextRow = filteredRows[i + 1];
            const nextLevel = parseInt(nextRow[wbsLevelIndex + 1]);

            // If current level is greater than or equal to next level, keep current row
            if (currentLevel >= nextLevel) {
                finalRows.push(currentRow);
            }
        } else {
            // Always keep the last row
            finalRows.push(currentRow);
        }
    }

    return finalRows;
}