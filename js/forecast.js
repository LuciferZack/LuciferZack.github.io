// ======================
// 预测页面功能
// ======================

// 全局变量
let forecastData = [];
let tecoProjects = {};
let hasTecoProjects = false;

// Add this function to forecast.js
function loadAndDisplayForecastData() {
    if (forecastData.length > 0) {
        displayForecastTable(forecastData);
    }
}

// Add this function to forecast.js
function updateForecastPeriod() {
    if (forecastData.length === 0) return;

    try {
        // Prepare new forecast data with updated fiscal year and period
        const newForecastData = prepareForecastData(filteredData);
        
        // Update the global forecastData
        forecastData = newForecastData;
        
        // Redisplay the forecast table
        displayForecastTable(forecastData);
        
        updateStatus('Forecast period updated successfully', 'success');
    } catch (error) {
        updateStatus(`Error updating forecast period: ${error.message}`, 'error');
        console.error(error);
    }
}

// Update the initForecastPage function to add the event listener for the new button
function initForecastPage() {
    // Check if we already have forecast data to display
    if (forecastData.length > 0) {
        displayForecastTable(forecastData);
    }
    
    // Add event listener for the update period button
    const updatePeriodBtn = document.getElementById('updatePeriodBtn');
    if (updatePeriodBtn && !updatePeriodBtn.hasAttribute('data-initialized')) {
        updatePeriodBtn.addEventListener('click', updateForecastPeriod);
        updatePeriodBtn.setAttribute('data-initialized', 'true');
    }
    
    // Rest of your existing event listeners...
    // Make sure these event listeners are only added once
    const saveBtn = document.getElementById('saveForecastBtn');
    if (saveBtn && !saveBtn.hasAttribute('data-initialized')) {
        saveBtn.addEventListener('click', function() {
            updateStatus('The prediction data has been saved (locally).', 'success');
        });
        saveBtn.setAttribute('data-initialized', 'true');
    }

    const exportBtn = document.getElementById('exportForecastBtn');
    if (exportBtn && !exportBtn.hasAttribute('data-initialized')) {
        exportBtn.addEventListener('click', exportForecastData);
        exportBtn.setAttribute('data-initialized', 'true');
    }

    const hasTecoSelect = document.getElementById('hasTeco');
    if (hasTecoSelect && !hasTecoSelect.hasAttribute('data-initialized')) {
        hasTecoSelect.addEventListener('change', function() {
            const tecoControls = document.getElementById('tecoControls');
            tecoControls.style.display = this.value === 'yes' ? 'block' : 'none';
            hasTecoProjects = this.value === 'yes';
        });
        hasTecoSelect.setAttribute('data-initialized', 'true');
    }

    const addTecoBtn = document.getElementById('addTecoBtn');
    if (addTecoBtn && !addTecoBtn.hasAttribute('data-initialized')) {
        addTecoBtn.addEventListener('click', showTecoModal);
        addTecoBtn.setAttribute('data-initialized', 'true');
    }

    const loadPreviousBtn = document.getElementById('loadPreviousBtn');
    if (loadPreviousBtn && !loadPreviousBtn.hasAttribute('data-initialized')) {
        loadPreviousBtn.addEventListener('click', loadPreviousForecast);
        loadPreviousBtn.setAttribute('data-initialized', 'true');
    }

    const confirmTecoBtn = document.getElementById('confirmTecoBtn');
    if (confirmTecoBtn && !confirmTecoBtn.hasAttribute('data-initialized')) {
        confirmTecoBtn.addEventListener('click', confirmTecoSelection);
        confirmTecoBtn.setAttribute('data-initialized', 'true');
    }

    const cancelTecoBtn = document.getElementById('cancelTecoBtn');
    if (cancelTecoBtn && !cancelTecoBtn.hasAttribute('data-initialized')) {
        cancelTecoBtn.addEventListener('click', function() {
            document.getElementById('tecoModal').style.display = 'none';
        });
        cancelTecoBtn.setAttribute('data-initialized', 'true');
    }
}
// Modified prepareForecastData function to include Plan Margin% and Actual Margin%
function prepareForecastData(data) {
    if (data.length < 2) return [];

    const headers = data[0];
    const wbsLevelIndex = headers.indexOf('WBS Level');
    const personIndex = headers.indexOf('Person Respons.');
    const wbsElementIndex = headers.indexOf('WBS Element');
    const wbsDescIndex = headers.indexOf('WBS Description');
    const newOrderIndex = headers.indexOf('New Order');
    const pocIndex = headers.indexOf('POC%');
    const planCostIndex = headers.indexOf('Plan Cost All');
    const committedCostIndex = headers.indexOf('Committed Cost');
    const actualCostIndex = headers.indexOf('Actual Cost');
    const planMarginIndex = headers.indexOf('Plan Margin%');  // Added
    const actualMarginIndex = headers.indexOf('Actual Margin%'); // Added

    // Create forecast data structure
    const forecastData = [];

    // Add header row
    const forecastHeaders = [
        'Level', 'PM', 'Proj.number', 'Proj.description', 'New order',
        'POC%', 'Plan Cost', 'Committed Cost', 'Actual Cost', 'Cost Remain'
    ];

    // Add current fiscal year remaining months COST columns
    const fiscalYear = parseInt(document.getElementById('fiscalYear').value);
    const currentPeriod = parseInt(document.getElementById('currentPeriod').value);

    // Modify to start from current month to P12
    for (let p = currentPeriod; p <= 12; p++) {
        forecastHeaders.push(`COST-F${fiscalYear}P${p}`);
    }

    // Also modify TO columns
    for (let p = currentPeriod; p <= 12; p++) {
        forecastHeaders.push(`TO-F${fiscalYear}P${p}`);
    }

    // Add next year four quarters COST columns
    for (let q = 1; q <= 4; q++) {
        forecastHeaders.push(`COST-F${fiscalYear + 1}Q${q}`);
    }

    // Add next year four quarters TO columns
    for (let q = 1; q <= 4; q++) {
        forecastHeaders.push(`TO-F${fiscalYear + 1}Q${q}`);
    }

    // Add the year after next COST column
    forecastHeaders.push(`COST-F${fiscalYear + 2}`);

    // Add the year after next TO column
    forecastHeaders.push(`TO-F${fiscalYear + 2}`);

    // Add Plan Margin% and Actual Margin% at the end
    forecastHeaders.push('Plan Margin%');
    forecastHeaders.push('Actual Margin%');

    forecastData.push(forecastHeaders);

    // Add data rows
    for (let i = 1; i < data.length; i++) {
        const row = data[i];
        const planCost = parseFloat(row[planCostIndex]) || 0;
        const actualCost = parseFloat(row[actualCostIndex]) || 0;

        const forecastRow = [
            row[wbsLevelIndex], // Level
            row[personIndex], // PM
            row[wbsElementIndex], // Proj.number
            row[wbsDescIndex], // Proj.description
            row[newOrderIndex], // New order
            row[pocIndex], // POC%
            planCost, // Plan Cost
            row[committedCostIndex], // Committed Cost
            actualCost, // Actual Cost
            (planCost - actualCost).toFixed(2) // Cost Remain (fixed to 2 decimal places)
        ];

        // Add empty forecast columns (COST and TO)
        const forecastColsCount = forecastHeaders.length - 12; // 12 = 10 fixed columns + 2 margin columns
        for (let j = 0; j < forecastColsCount; j++) {
            forecastRow.push('');
        }

        // Add Plan Margin% and Actual Margin% values
        forecastRow.push(row[planMarginIndex] || '');  // Plan Margin%
        forecastRow.push(row[actualMarginIndex] || ''); // Actual Margin%

        forecastData.push(forecastRow);
    }

    return forecastData;
}

// Modified displayForecastTable function to handle the new margin columns with 3 decimal places display
function displayForecastTable(data) {
    const forecastTable = document.getElementById('forecastTable');
    forecastTable.innerHTML = '';

    // Set table styles directly
    forecastTable.style.fontSize = '10px';
    forecastTable.style.width = '100%';

    // Add this to ensure the container has the right class
    const container = forecastTable.parentElement;
    container.classList.add('forecast-table-container');
    container.style.height = '500px';
    container.style.fontSize = '10px';
    if (!container.classList.contains('forecast-table-container')) {
        container.classList.add('forecast-table-container');
    }

    if (data.length === 0) return;

    // Create table header
    const thead = document.createElement('thead');
    const headerRow = document.createElement('tr');

    // First 10 columns are fixed titles
    for (let i = 0; i < 10; i++) {
        const th = document.createElement('th');
        th.textContent = data[0][i];
        th.classList.add('fixed-column');
        headerRow.appendChild(th);
    }

    // Add all COST column headers
    for (let i = 10; i < data[0].length - 2; i++) { // -2 to exclude margin columns
        if (data[0][i].startsWith('COST-')) {
            const th = document.createElement('th');
            th.textContent = data[0][i];
            th.classList.add('forecast-header');
            headerRow.appendChild(th);
        }
    }

    // Add all TO column headers
    for (let i = 10; i < data[0].length - 2; i++) { // -2 to exclude margin columns
        if (data[0][i].startsWith('TO-')) {
            const th = document.createElement('th');
            th.textContent = data[0][i];
            th.classList.add('to-header');
            headerRow.appendChild(th);
        }
    }

    // Add margin column headers
    const planMarginHeader = document.createElement('th');
    planMarginHeader.textContent = data[0][data[0].length - 2]; // Plan Margin%
    planMarginHeader.classList.add('margin-column');
    headerRow.appendChild(planMarginHeader);

    const actualMarginHeader = document.createElement('th');
    actualMarginHeader.textContent = data[0][data[0].length - 1]; // Actual Margin%
    actualMarginHeader.classList.add('margin-column');
    headerRow.appendChild(actualMarginHeader);

    thead.appendChild(headerRow);
    forecastTable.appendChild(thead);

    // Create table content
    const tbody = document.createElement('tbody');

    // First identify parent-child relationships and group them
    const projectGroups = {};
    for (let i = 1; i < data.length; i++) {
        const projNumber = data[i][2]; // Proj.number
        const level = parseInt(data[i][0]); // Level

        if (level === 1) {
            // Parent project
            if (!projectGroups[projNumber]) {
                projectGroups[projNumber] = {
                    parentIndex: i,
                    children: []
                };
            }
        } else {
            // Child project - find parent
            const parentProjNumber = projNumber.split('.')[0];
            if (projectGroups[parentProjNumber]) {
                projectGroups[parentProjNumber].children.push(i);
            }
        }
    }

    // Add rows in group order
    for (const projNumber in projectGroups) {
        const group = projectGroups[projNumber];
        const parentRowData = data[group.parentIndex];
        const isTecoProject = tecoProjects[projNumber];
        let tecoPeriodIndex = -1;
        let tecoPeriod = '';

        if (isTecoProject) {
            tecoPeriod = tecoProjects[projNumber];
            tecoPeriodIndex = data[0].indexOf(`COST-${tecoPeriod}`);
        }

        // Add parent row
        const parentRow = document.createElement('tr');
        parentRow.classList.add('parent-row');
        parentRow.dataset.index = group.parentIndex;

        // First 10 columns are fixed data
        for (let j = 0; j < 10; j++) {
            const td = document.createElement('td');
            td.textContent = parentRowData[j];
            td.classList.add('readonly');
            td.classList.add('fixed-column');
            parentRow.appendChild(td);
        }

        // COST columns - parent project automatically sums child projects
        const costColumns = [];
        for (let j = 10; j < data[0].length - 2; j++) { // -2 to exclude margin columns
            if (data[0][j].startsWith('COST-')) {
                costColumns.push(j);

                const td = document.createElement('td');
                td.textContent = ''; // Initially empty, will be calculated later
                td.classList.add('readonly');

                // If this is the TECO period column, mark it specially
                if (isTecoProject && j === tecoPeriodIndex) {
                    td.textContent = 'TECO';
                    td.style.color = 'red';
                    td.style.fontWeight = 'bold';
                }

                parentRow.appendChild(td);
            }
        }

        // TO columns - automatically calculated for parent
        for (let j = 10; j < data[0].length - 2; j++) { // -2 to exclude margin columns
            if (data[0][j].startsWith('TO-')) {
                const td = document.createElement('td');

                // Special calculation for TECO month in parent
                if (isTecoProject && j === tecoPeriodIndex + costColumns.length) {
                    const newOrder = parseFloat(parentRowData[4]) || 0;
                    const pocPercent = parseFloat(parentRowData[5]) || 0;
                    let costBeforeTeco = 0;

                    // Calculate sum of costs before TECO month from all children
                    group.children.forEach(childIdx => {
                        const childData = data[childIdx];
                        for (let k = 0; k < costColumns.length; k++) {
                            const period = data[0][costColumns[k]].replace('COST-', '');
                            if (isPeriodBefore(period, tecoPeriod)) {
                                costBeforeTeco += parseFloat(childData[costColumns[k]]) || 0;
                            }
                        }
                    });

                    const toValue = newOrder * (1 - pocPercent / 100) - costBeforeTeco;
                    td.textContent = toValue.toFixed(2);

                    // Update forecastData
                    forecastData[group.parentIndex][j] = toValue.toFixed(2);
                } else {
                    td.textContent = calculateTOValue(parentRowData, j, costColumns);
                }

                td.classList.add('readonly');
                parentRow.appendChild(td);
            }
        }

        // Add margin columns (readonly) with 3 decimal places display
        const planMarginCell = document.createElement('td');
        const planMarginValue = parentRowData[parentRowData.length - 2]; // Plan Margin%
        // Display with 3 decimal places but keep actual value
        if (planMarginValue !== '' && !isNaN(planMarginValue)) {
            const formattedValue = parseFloat(planMarginValue);
            planMarginCell.textContent = isNaN(formattedValue) ? planMarginValue : formattedValue.toFixed(3);
        } else {
            planMarginCell.textContent = planMarginValue;
        }
        planMarginCell.classList.add('readonly', 'margin-column');
        parentRow.appendChild(planMarginCell);

        const actualMarginCell = document.createElement('td');
        const actualMarginValue = parentRowData[parentRowData.length - 1]; // Actual Margin%
        // Display with 3 decimal places but keep actual value
        if (actualMarginValue !== '' && !isNaN(actualMarginValue)) {
            const formattedValue = parseFloat(actualMarginValue);
            actualMarginCell.textContent = isNaN(formattedValue) ? actualMarginValue : formattedValue.toFixed(3);
        } else {
            actualMarginCell.textContent = actualMarginValue;
        }
        actualMarginCell.classList.add('readonly', 'margin-column');
        parentRow.appendChild(actualMarginCell);

        tbody.appendChild(parentRow);

        // Add child rows and collect COST values
        const childCostValues = Array(costColumns.length).fill(0);
        group.children.forEach(childIndex => {
            const childRowData = data[childIndex];
            const childRow = document.createElement('tr');
            childRow.classList.add('child-row');
            childRow.dataset.index = childIndex;

            // First 10 columns are fixed data
            for (let j = 0; j < 10; j++) {
                const td = document.createElement('td');
                td.textContent = childRowData[j];
                td.classList.add('readonly');
                td.classList.add('fixed-column');
                childRow.appendChild(td);
            }

            // COST columns - editable for child projects
            costColumns.forEach((colIndex, costIdx) => {
                const td = document.createElement('td');

                if (isTecoProject) {
                    // Check if this column is before, at, or after TECO period
                    const currentPeriod = data[0][colIndex].replace('COST-', '');

                    if (currentPeriod === tecoPeriod) {
                        // TECO month - show TECO mark and lock
                        td.textContent = 'TECO';
                        td.style.color = 'red';
                        td.style.fontWeight = 'bold';
                        td.classList.add('readonly');

                        // Mark in forecastData
                        forecastData[childIndex][colIndex] = 'TECO';
                    } else if (isPeriodAfter(currentPeriod, tecoPeriod)) {
                        // After TECO - locked
                        td.textContent = '';
                        td.classList.add('readonly');

                        // Clear in forecastData
                        forecastData[childIndex][colIndex] = '';
                    } else {
                        // Before TECO - editable
                        const input = document.createElement('input');
                        input.type = 'number';
                        input.min = '0';
                        input.step = '0.01';
                        input.value = childRowData[colIndex] || '';
                        input.classList.add('editable');

                        input.addEventListener('input', function() {
                            // Update TO values
                            updateTOValues(childRow, childRowData, costColumns);
                            // Update parent sums
                            updateParentCostSum(parentRow, group.children, costColumns);
                            // Special update for TECO month TO
                            updateTecoMonthTO(childRow, childRowData, costColumns, tecoPeriod);
                        });

                        td.appendChild(input);
                    }
                } else {
                    // Not a TECO project - always editable
                    const input = document.createElement('input');
                    input.type = 'number';
                    input.min = '0';
                    input.step = '0.01';
                    input.value = childRowData[colIndex] || '';
                    input.classList.add('editable');

                    input.addEventListener('input', function() {
                        updateTOValues(childRow, childRowData, costColumns);
                        updateParentCostSum(parentRow, group.children, costColumns);
                    });

                    td.appendChild(input);
                }

                childRow.appendChild(td);

                // Record initial COST value
                if (td.querySelector('input')) {
                    childCostValues[costIdx] += parseFloat(td.querySelector('input').value) || 0;
                } else if (td.textContent && td.textContent !== 'TECO') {
                    childCostValues[costIdx] += parseFloat(td.textContent) || 0;
                }
            });

            // TO columns - automatically calculated for child
            for (let j = 10; j < data[0].length - 2; j++) { // -2 to exclude margin columns
                if (data[0][j].startsWith('TO-')) {
                    const td = document.createElement('td');

                    // Special calculation for TECO month in child
                    if (isTecoProject && j === tecoPeriodIndex + costColumns.length) {
                        // Use parent's New order and POC%
                        const newOrder = parseFloat(parentRowData[4]) || 0;
                        const pocPercent = parseFloat(parentRowData[5]) || 0;
                        let costBeforeTeco = 0;

                        // Calculate sum of costs before TECO month from all children
                        group.children.forEach(childIdx => {
                            const childData = data[childIdx];
                            for (let k = 0; k < costColumns.length; k++) {
                                const period = data[0][costColumns[k]].replace('COST-', '');
                                if (isPeriodBefore(period, tecoPeriod)) {
                                    costBeforeTeco += parseFloat(childData[costColumns[k]]) || 0;
                                }
                            }
                        });

                        const toValue = newOrder * (1 - pocPercent / 100) - costBeforeTeco;
                        td.textContent = toValue.toFixed(2);

                        // Update forecastData
                        forecastData[childIndex][j] = toValue.toFixed(2);
                    } else {
                        td.textContent = calculateTOValue(childRowData, j, costColumns);
                    }

                    td.classList.add('readonly');
                    childRow.appendChild(td);
                }
            }

            // Add margin columns for child rows (readonly) with 3 decimal places display
            const childPlanMarginCell = document.createElement('td');
            const childPlanMarginValue = childRowData[childRowData.length - 2]; // Plan Margin%
            // Display with 3 decimal places but keep actual value
            if (childPlanMarginValue !== '' && !isNaN(childPlanMarginValue)) {
                const formattedValue = parseFloat(childPlanMarginValue);
                childPlanMarginCell.textContent = isNaN(formattedValue) ? childPlanMarginValue : formattedValue.toFixed(3);
            } else {
                childPlanMarginCell.textContent = childPlanMarginValue;
            }
            childPlanMarginCell.classList.add('readonly', 'margin-column');
            childRow.appendChild(childPlanMarginCell);

            const childActualMarginCell = document.createElement('td');
            const childActualMarginValue = childRowData[childRowData.length - 1]; // Actual Margin%
            // Display with 3 decimal places but keep actual value
            if (childActualMarginValue !== '' && !isNaN(childActualMarginValue)) {
                const formattedValue = parseFloat(childActualMarginValue);
                childActualMarginCell.textContent = isNaN(formattedValue) ? childActualMarginValue : formattedValue.toFixed(3);
            } else {
                childActualMarginCell.textContent = childActualMarginValue;
            }
            childActualMarginCell.classList.add('readonly', 'margin-column');
            childRow.appendChild(childActualMarginCell);

            tbody.appendChild(childRow);
        });

        // Update parent row's COST sum
        costColumns.forEach((colIndex, costIdx) => {
            const costTds = parentRow.querySelectorAll('td');
            const costTdIndex = 10 + costIdx;

            // For TECO projects, don't sum after TECO period
            if (isTecoProject) {
                const currentPeriod = data[0][colIndex].replace('COST-', '');
                if (isPeriodAfter(currentPeriod, tecoPeriod)) {
                    costTds[costTdIndex].textContent = '0.00';
                    forecastData[group.parentIndex][colIndex] = '0.00';
                    return;
                }
            }

            costTds[costTdIndex].textContent = childCostValues[costIdx].toFixed(2);
            forecastData[group.parentIndex][colIndex] = childCostValues[costIdx].toFixed(2);
        });

        // Update parent row's TO values (non-TECO months)
        updateTOValues(parentRow, parentRowData, costColumns);
    }

    forecastTable.appendChild(tbody);
}

// 显示TECO模态框
function showTecoModal() {
    const modal = document.getElementById('tecoModal');
    const projectSelect = document.getElementById('tecoProjectSelect');
    
    // Clear previous options
    projectSelect.innerHTML = '';
    
    // Get all parent projects from forecastData
    const parentProjects = [];
    for (let i = 1; i < forecastData.length; i++) {
        if (parseInt(forecastData[i][0]) === 1) { // Level 1 projects
            const projNumber = forecastData[i][2];
            if (!tecoProjects[projNumber]) { // Only show projects not already TECO'd
                parentProjects.push({
                    number: projNumber,
                    description: forecastData[i][3]
                });
            }
        }
    }
    
    // Add projects to select
    parentProjects.forEach(proj => {
        const option = document.createElement('option');
        option.value = proj.number;
        option.textContent = `${proj.number} - ${proj.description}`;
        projectSelect.appendChild(option);
    });
    
    // Populate period select
    const periodSelect = document.getElementById('tecoPeriodSelect');
    periodSelect.innerHTML = '';
    
    const fiscalYear = parseInt(document.getElementById('fiscalYear').value);
    const currentPeriod = parseInt(document.getElementById('currentPeriod').value);
    
    // Add remaining periods in current fiscal year
    for (let p = currentPeriod + 1; p <= 12; p++) {
        const option = document.createElement('option');
        option.value = `F${fiscalYear}P${p}`;
        option.textContent = `F${fiscalYear}P${p}`;
        periodSelect.appendChild(option);
    }
    
    // Add next year's quarters
    for (let q = 1; q <= 4; q++) {
        const option = document.createElement('option');
        option.value = `F${fiscalYear + 1}Q${q}`;
        option.textContent = `F${fiscalYear + 1}Q${q}`;
        periodSelect.appendChild(option);
    }
    
    modal.style.display = 'flex';
}

// 确认TECO选择
function confirmTecoSelection() {
    const projectSelect = document.getElementById('tecoProjectSelect');
    const periodSelect = document.getElementById('tecoPeriodSelect');
    const selectedProjects = Array.from(projectSelect.selectedOptions).map(opt => opt.value);
    const tecoPeriod = periodSelect.value;
    
    if (selectedProjects.length === 0 || !tecoPeriod) {
        alert('Please select the project and TECO point');
        return;
    }
    
    // Add to tecoProjects
    selectedProjects.forEach(proj => {
        tecoProjects[proj] = tecoPeriod;
    });
    
    // Update UI
    updateTecoProjectsDisplay();
    
    // Close modal
    document.getElementById('tecoModal').style.display = 'none';
    
    // Refresh forecast table
    displayForecastTable(forecastData);
}

// 更新TECO项目显示
function updateTecoProjectsDisplay() {
    const container = document.getElementById('tecoProjectsContainer');
    container.innerHTML = '';
    
    Object.keys(tecoProjects).forEach(proj => {
        const div = document.createElement('div');
        div.style.marginBottom = '5px';
        div.innerHTML = `
            <span>${proj} - TECO时间: ${tecoProjects[proj]}</span>
            <button data-project="${proj}" class="removeTecoBtn" style="margin-left: 10px; padding: 2px 5px;">删除</button>
        `;
        container.appendChild(div);
    });
    
    // Add event listeners to remove buttons
    document.querySelectorAll('.removeTecoBtn').forEach(btn => {
        btn.addEventListener('click', function() {
            delete tecoProjects[this.getAttribute('data-project')];
            updateTecoProjectsDisplay();
            displayForecastTable(forecastData);
        });
    });
}

// 比较时间段
function isPeriodAfter(period1, period2) {
    // Compare two period strings like "F2025P9" or "F2026Q1"
    const year1 = parseInt(period1.match(/F(\d+)/)[1]);
    const year2 = parseInt(period2.match(/F(\d+)/)[1]);
    
    if (year1 !== year2) return year1 > year2;
    
    // Same year, compare period
    const type1 = period1.includes('P') ? 'month' : 'quarter';
    const type2 = period2.includes('P') ? 'month' : 'quarter';
    
    if (type1 !== type2) {
        // Months come before quarters in the same year
        return type1 === 'quarter';
    }
    
    // Same type, compare numbers
    const num1 = parseInt(period1.match(/(P|Q)(\d+)/)[2]);
    const num2 = parseInt(period2.match(/(P|Q)(\d+)/)[2]);
    
    return num1 > num2;
}

function isPeriodBefore(period1, period2) {
    return !isPeriodAfter(period1, period2) && period1 !== period2;
}

// 计算TO值
function calculateTOValue(rowData, toColIndex, costColumns) {
    // 找到对应的COST列
    const toHeader = forecastData[0][toColIndex];
    const costHeader = toHeader.replace('TO-', 'COST-');
    const costColIndex = forecastData[0].indexOf(costHeader);
    
    if (costColIndex === -1) return '0.00';
    
    const costValue = parseFloat(rowData[costColIndex]) || 0;
    const planCost = parseFloat(rowData[6]) || 0; // 第7列Plan Cost
    const newOrder = parseFloat(rowData[4]) || 0; // 第5列New order
    
    if (planCost === 0) return '0.00';
    
    const toValue = (costValue / planCost) * newOrder;
    return toValue.toFixed(2);
}

// 更新TO值
function updateTOValues(row, rowData, costColumns) {
    const toCells = row.querySelectorAll('td');
    let toCellIndex = 10 + costColumns.length; // TO列开始位置
    
    for (let j = 10; j < forecastData[0].length; j++) {
        if (forecastData[0][j].startsWith('TO-')) {
            // 检查是否是TECO项目的TECO月
            const projNumber = rowData[2].split('.')[0]; // 获取父项目号
            const tecoPeriod = tecoProjects[projNumber];
            const toPeriod = forecastData[0][j].replace('TO-', '');
            
            if (tecoPeriod && toPeriod === tecoPeriod) {
                // 跳过TECO月的TO计算，因为它有特殊公式
                toCellIndex++;
                continue;
            }
            
            // 正常TO计算
            const toValue = calculateTOValue(rowData, j, costColumns);
            toCells[toCellIndex].textContent = toValue;
            
            // 更新forecastData中的数据
            const rowIndex = parseInt(row.dataset.index);
            if (rowIndex && forecastData[rowIndex]) {
                forecastData[rowIndex][j] = toValue;
            }
            
            toCellIndex++;
        }
    }
}

// 更新TECO月的TO值
function updateTecoMonthTO(row, rowData, costColumns, tecoPeriod) {
    const toCells = row.querySelectorAll('td');
    let toCellIndex = 10 + costColumns.length; // TO列开始位置
    
    // 获取父项目数据（如果是子项目）
    let parentRowData = rowData;
    let isChildRow = parseInt(rowData[0]) > 1; // Level > 1 表示是子项目
    
    if (isChildRow) {
        const parentProjNumber = rowData[2].split('.')[0];
        for (let i = 1; i < forecastData.length; i++) {
            if (parseInt(forecastData[i][0]) === 1 && forecastData[i][2] === parentProjNumber) {
                parentRowData = forecastData[i];
                break;
            }
        }
    }
    
    for (let j = 10; j < forecastData[0].length; j++) {
        if (forecastData[0][j].startsWith('TO-')) {
            const toPeriod = forecastData[0][j].replace('TO-', '');
            
            if (toPeriod === tecoPeriod) {
                // 使用父项目的New order和POC%（直接使用POC%的值，不除以100）
                const newOrder = parseFloat(parentRowData[4]) || 0;
                const pocValue = parseFloat(parentRowData[5]) || 0; // 这里POC%已经是百分比形式
                let costBeforeTeco = 0;
                
                // 计算所有子项目在TECO月前的COST总和
                const parentProjNumber = parentRowData[2];
                const children = [];
                
                // 找到所有子项目
                for (let k = 1; k < forecastData.length; k++) {
                    if (parseInt(forecastData[k][0]) > 1 && 
                        forecastData[k][2].startsWith(parentProjNumber)) {
                        children.push(k);
                    }
                }
                
                // 计算这些子项目在TECO月前的COST总和
                children.forEach(childIndex => {
                    for (let k = 0; k < costColumns.length; k++) {
                        const costPeriod = forecastData[0][costColumns[k]].replace('COST-', '');
                        if (isPeriodBefore(costPeriod, tecoPeriod)) {
                            const childRow = document.getElementById('forecastTable').querySelector(`tr[data-index="${childIndex}"]`);
                            if (childRow) {
                                const costCell = childRow.querySelectorAll('td')[10 + k];
                                if (costCell) {
                                    let costValue = 0;
                                    if (costCell.querySelector('input')) {
                                        costValue = parseFloat(costCell.querySelector('input').value) || 0;
                                    } else if (costCell.textContent && costCell.textContent !== 'TECO') {
                                        costValue = parseFloat(costCell.textContent) || 0;
                                    }
                                    costBeforeTeco += costValue;
                                }
                            }
                        }
                    }
                });
                
                // 修正计算公式：直接使用pocValue，不除以100
                const toValue = newOrder * (1 - pocValue) - costBeforeTeco;
                toCells[toCellIndex].textContent = toValue.toFixed(2);
                
                // 更新forecastData中的数据
                const rowIndex = parseInt(row.dataset.index);
                if (rowIndex && forecastData[rowIndex]) {
                    forecastData[rowIndex][j] = toValue.toFixed(2);
                }
            }
            
            toCellIndex++;
        }
    }
}

// 更新父项目的COST总和
function updateParentCostSum(parentRow, childrenIndices, costColumns) {
    const parentCells = parentRow.querySelectorAll('td');
    const parentRowIndex = parseInt(parentRow.dataset.index);
    const projNumber = forecastData[parentRowIndex][2]; // 父项目号
    const isTecoProject = tecoProjects[projNumber];
    
    // 重置所有COST列
    costColumns.forEach((colIndex, costIdx) => {
        const costTdIndex = 10 + costIdx;
        parentCells[costTdIndex].textContent = '0.00';
        forecastData[parentRowIndex][colIndex] = '0.00';
    });
    
    // 计算所有子项目的COST总和
    childrenIndices.forEach(childIndex => {
        const childRow = document.getElementById('forecastTable').querySelector(`tr[data-index="${childIndex}"]`);
        if (!childRow) return;
        
        const childInputs = childRow.querySelectorAll('input');
        
        costColumns.forEach((colIndex, costIdx) => {
            const costTdIndex = 10 + costIdx;
            const currentSum = parseFloat(parentCells[costTdIndex].textContent) || 0;
            
            // 对于TECO项目，只计算TECO月之前的COST
            if (isTecoProject) {
                const currentPeriod = forecastData[0][colIndex].replace('COST-', '');
                const tecoPeriod = tecoProjects[projNumber];
                
                if (isPeriodAfter(currentPeriod, tecoPeriod)) {
                    // 跳过TECO月之后的COST
                    return;
                }
            }
            
            let childValue = 0;
            if (childInputs[costIdx]) {
                childValue = parseFloat(childInputs[costIdx].value) || 0;
            } else {
                // 可能是TECO月的单元格
                const childTd = childRow.querySelectorAll('td')[10 + costIdx];
                if (childTd && childTd.textContent !== 'TECO') {
                    childValue = parseFloat(childTd.textContent) || 0;
                }
            }
            
            parentCells[costTdIndex].textContent = (currentSum + childValue).toFixed(2);
            forecastData[parentRowIndex][colIndex] = (currentSum + childValue).toFixed(2);
            
            // 更新forecastData中子项目的COST值
            if (childInputs[costIdx]) {
                forecastData[childIndex][colIndex] = childInputs[costIdx].value;
            }
        });
    });
    
    // 更新父项目的TO值
    updateTOValues(parentRow, forecastData[parentRowIndex], costColumns);
    
    // 如果是TECO项目，还需要更新TECO月的TO值
    if (isTecoProject) {
        const tecoPeriod = tecoProjects[projNumber];
        
        // 更新父项目的TECO月TO值
        updateTecoMonthTO(parentRow, forecastData[parentRowIndex], costColumns, tecoPeriod);
        
        // 更新所有子项目的TECO月TO值
        childrenIndices.forEach(childIndex => {
            const childRow = document.getElementById('forecastTable').querySelector(`tr[data-index="${childIndex}"]`);
            if (childRow) {
                updateTecoMonthTO(childRow, forecastData[childIndex], costColumns, tecoPeriod);
            }
        });
    }
}

// 准备导出的预测数据
function forecastDataToExport() {
    const exportData = [];
    exportData.push([...forecastData[0]]);
    
    for (let i = 1; i < forecastData.length; i++) {
        const rowData = [...forecastData[i]];
        const projNumber = rowData[2];
        const isTecoProject = tecoProjects[projNumber];
        
        if (parseInt(rowData[0]) === 1 && isTecoProject) {
            const tecoPeriod = tecoProjects[projNumber];
            const tecoCostCol = forecastData[0].indexOf(`COST-${tecoPeriod}`);
            
            if (tecoCostCol !== -1) {
                rowData[tecoCostCol] = 'TECO';
            }
        }
        
        exportData.push(rowData);
    }
    
    return exportData;
}

// 导出预测数据
function exportForecastData() {
    if (forecastData.length === 0) return;

    const worksheet = XLSX.utils.aoa_to_sheet(forecastDataToExport());
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Turnover Forecast");

    // Get applicant and responsible person from the first row of filtered data
    const applicant = filteredData.length > 1 ? filteredData[1][filteredData[0].indexOf('Applicant')] : '';
    const responsible = filteredData.length > 1 ? filteredData[1][filteredData[0].indexOf('Person Respons.')] : '';
    const fiscalYear = document.getElementById('fiscalYear').value;
    const currentPeriod = document.getElementById('currentPeriod').value;

    XLSX.writeFile(workbook, `Forecast Sheet_${applicant}_${responsible}_F${fiscalYear}P${currentPeriod}.xlsx`);

    updateStatus('The export of predictive data has begun', 'success');
}

// 加载上月预测数据
function loadPreviousForecast() {
    const fileInput = document.getElementById('previousForecast');
    const file = fileInput.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = function(e) {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
        const previousData = XLSX.utils.sheet_to_json(firstSheet, { header: 1 });

        // 同步COST值
        syncPreviousForecast(previousData);
    };
    reader.readAsArrayBuffer(file);
}

// 同步上月预测数据
function syncPreviousForecast(previousData) {
    if (!previousData || previousData.length < 2) return;
    
    const prevHeaders = previousData[0];
    const currentHeaders = forecastData[0];
    
    for (let i = 10; i < prevHeaders.length; i++) {
        if (prevHeaders[i].startsWith('COST-')) {
            const colName = prevHeaders[i];
            const currentColIndex = currentHeaders.indexOf(colName);
            
            if (currentColIndex !== -1) {
                for (let row = 1; row < Math.min(previousData.length, forecastData.length); row++) {
                    // Preserve string values like "TECO"
                    if (typeof previousData[row][i] === 'string' && isNaN(previousData[row][i])) {
                        forecastData[row][currentColIndex] = previousData[row][i];
                    } else {
                        forecastData[row][currentColIndex] = previousData[row][i] || '';
                    }
                }
            }
        }
    }
    
    displayForecastTable(forecastData);
}

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
            
            // If navigating to forecast page and we have forecast data, display it
            if (targetPageId === 'forecast' && typeof forecastData !== 'undefined' && forecastData.length > 0) {
                // Use setTimeout to ensure the page is fully rendered
                setTimeout(function() {
                    if (typeof displayForecastTable === 'function') {
                        displayForecastTable(forecastData);
                    }
                }, 100);
            }
        });
    });
    initForecastPage();
})
