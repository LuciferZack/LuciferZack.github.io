let mergedCashflowData = [];
let currentSortColumn = null;
let sortDirection = 'asc';

// New variables for forecast merging
let mergedForecastData = [];
let forecastFiles = [];

// 在DOMContentLoaded事件监听器中添加新的事件监听器
document.addEventListener('DOMContentLoaded', function() {
    // 现有的代码...
    
    // Add event listeners for forecast table tabs
    const forecastTabButtons = document.querySelectorAll('#dataIntegration .tab-button');
    forecastTabButtons.forEach(button => {
        button.addEventListener('click', function() {
            const tabId = this.getAttribute('data-tab');
            
            // Remove active class from all buttons and content
            document.querySelectorAll('#dataIntegration .tab-button').forEach(btn => {
                btn.classList.remove('active');
            });
            document.querySelectorAll('#dataIntegration .tab-content').forEach(content => {
                content.classList.remove('active');
            });
            
            // Add active class to clicked button and corresponding content
            this.classList.add('active');
            document.getElementById(tabId).classList.add('active');
        });
    });
});

document.addEventListener('DOMContentLoaded', function() {
    document.getElementById('cashflowFiles').addEventListener('change', function(e) {
        const files = e.target.files;
        if (files.length === 0) {
            document.getElementById('mergeCashflowBtn').disabled = true;
            return;
        }
        document.getElementById('mergeCashflowBtn').disabled = false;
    });

    document.getElementById('mergeCashflowBtn').addEventListener('click', mergeCashflowFiles);
    document.getElementById('exportMergedCashflowBtn').addEventListener('click', exportMergedCashflowData);
});
// Add event listeners for forecast merging
document.addEventListener('DOMContentLoaded', function() {
    // Existing code...
    
    // Add new event listeners for forecast merging
    document.getElementById('forecastFiles').addEventListener('change', handleForecastFileSelection);
    document.getElementById('mergeForecastBtn').addEventListener('click', mergeForecastFiles);
    document.getElementById('exportMergedForecastBtn').addEventListener('click', exportMergedForecastData);
});
// Add this to integration.js

// New variables for processed forecast data
let processedForecastData = [];
let previousMonthData = []; // To store previous month's overall data

// Add event listener for previous month data upload
document.addEventListener('DOMContentLoaded', function() {
    document.getElementById('previousMonthFile').addEventListener('change', handlePreviousMonthFile);
    document.getElementById('processForecastBtn').addEventListener('click', processMergedForecastData);
    document.getElementById('exportProcessedForecastBtn').addEventListener('click', exportProcessedForecastData);
});
// Add new variable to store the processed customer & contract information data
let customerContractData = {};

// Add event listener for customer & contract information file upload
document.addEventListener('DOMContentLoaded', function() {
    document.getElementById('customerContractFile').addEventListener('change', handleCustomerContractFile);
});

// Function to handle customer & contract information file
function handleCustomerContractFile(e) {
    const file = e.target.files[0];
    if (!file) return;

    readExcelFile(file).then(data => {
        if (data.length >= 2) {
            // Process the customer & contract information data
            customerContractData = processCustomerContractData(data);
            updateStatus('Customer & Contract Information loaded successfully', 'success');
            
            // Enable process button if we have merged forecast data
            if (mergedForecastData.length > 0 && previousMonthData.length >= 5) {
                document.getElementById('processForecastBtn').disabled = false;
            }
        } else {
            updateStatus('Customer & Contract Information file is invalid or empty', 'error');
        }
    }).catch(error => {
        updateStatus(`Error loading Customer & Contract Information: ${error.message}`, 'error');
        console.error(error);
    });
}

// Function to process customer & contract information data
function processCustomerContractData(data) {
    const result = {};
    
    // Skip header row (assuming first row is headers)
    for (let i = 1; i < data.length; i++) {
        const row = data[i];
        
        // Check if row has enough columns
        if (row.length >= 39) {
            const wbs = row[15]; // 16th column (0-indexed)
            const toValue = parseFloat(row[37]) || 0; // 38th column (0-indexed)
            const gmValue = parseFloat(row[38]) || 0; // 39th column (0-indexed)
            
            if (wbs) {
                // If WBS already exists, add TO and GM values
                if (result[wbs]) {
                    result[wbs].to += toValue;
                    result[wbs].gm += gmValue;
                } else {
                    // Create new entry
                    result[wbs] = {
                        to: toValue,
                        gm: gmValue
                    };
                }
            }
        }
    }
    
    return result;
}
// Replace the existing handlePreviousMonthFile function with this updated version

function handlePreviousMonthFile(e) {
    const file = e.target.files[0];
    if (!file) return;

    readExcelFile(file).then(data => {
        // Process the previous month data correctly
        // Skip first 3 rows, 4th row contains headers, data starts from 5th row
        if (data.length >= 4) {
            previousMonthData = data;
            updateStatus('Previous month data loaded successfully', 'success');
            
            // Enable process button if we have merged forecast data
            if (mergedForecastData.length > 0) {
                document.getElementById('processForecastBtn').disabled = false;
            }
        } else {
            updateStatus('Previous month data file is invalid or empty', 'error');
        }
    }).catch(error => {
        updateStatus(`Error loading previous month data: ${error.message}`, 'error');
        console.error(error);
    });
}

function getPreviousMonthValue(projNumber, columnName) {
    if (previousMonthData.length < 5) return ''; // Need at least 5 rows (headers + 1 data row)

    // Headers are in the 4th row (index 3)
    const headers = previousMonthData[3];
    let wbsElementIndex = headers.indexOf('WBS Element');
    
    if (wbsElementIndex === -1) {
        // Try to find WBS column by other possible names
        const possibleNames = ['WBS', 'Project Number', 'Proj.number', 'WBS Element', 'WBS Element '];
        for (const name of possibleNames) {
            const index = headers.indexOf(name);
            if (index !== -1) {
                wbsElementIndex = index;
                break;
            }
        }
        if (wbsElementIndex === -1) return '';
    }

    // Find column index in processed forecast table headers
    const processedHeaders = processedForecastData[0];
    const targetColumnIndex = processedHeaders.indexOf(columnName);
    
    if (targetColumnIndex === -1) return '';

    // Search for matching project
    // Data starts from the 5th row (index 4)
    for (let i = 4; i < previousMonthData.length; i++) {
        const row = previousMonthData[i];
        if (row && row[wbsElementIndex] && row[wbsElementIndex].toString().trim() === projNumber.toString().trim()) {
            // Direct mapping: use the same column index in the previous month data
            // Since both tables have the same structure (55 columns)
            return row[targetColumnIndex] || '';
        }
    }

    return ''; // Not found
}

function getPreviousMonthValueByIndex(projNumber, columnIndex) {
    if (previousMonthData.length < 5) return '';
    
    // Headers are in the 4th row (index 3)
    const headers = previousMonthData[3];
    let wbsElementIndex = headers.indexOf('WBS Element');
    
    if (wbsElementIndex === -1) {
        // Try to find WBS column by other possible names
        const possibleNames = ['WBS', 'Project Number', 'Proj.number', 'WBS Element', 'WBS Element '];
        for (const name of possibleNames) {
            const index = headers.indexOf(name);
            if (index !== -1) {
                wbsElementIndex = index;
                break;
            }
        }
        if (wbsElementIndex === -1) return '';
    }

    // Search for matching project
    // Data starts from the 5th row (index 4)
    for (let i = 4; i < previousMonthData.length; i++) {
        const row = previousMonthData[i];
        if (row && row[wbsElementIndex] && row[wbsElementIndex].toString().trim() === projNumber.toString().trim()) {
            // Direct mapping by column index
            return row[columnIndex] || '';
        }
    }

    return ''; // Not found
}

// Replace the processMergedForecastData function with this corrected version
function processMergedForecastData() {
    if (mergedForecastData.length === 0) {
        updateStatus('No merged forecast data to process', 'error');
        return;
    }

    if (previousMonthData.length < 5) {
        updateStatus('Please load previous month data first', 'error');
        return;
    }

    // Get current fiscal year and period
    const fiscalYear = parseInt(document.getElementById('fiscalYear').value);
    const currentPeriod = parseInt(document.getElementById('currentPeriod').value);

    // Create headers for processed data
    const headers = [
        'WBS Element',
        'CPM',
        'PM',
        'Proj. description',
        'Type',
        'Beginning BFCklog',
        'Remaining BFCklog',
        'Plan GM%',
        'Plan GM'
    ];

    // Add actual turnover columns for completed periods
    for (let p = 1; p < currentPeriod; p++) {
        headers.push(`FY${fiscalYear.toString().slice(-2)} P${p} AC`);
    }

    // Add forecast columns for current and future periods
    for (let p = currentPeriod; p <= 12; p++) {
        headers.push(`FY${fiscalYear.toString().slice(-2)} P${p} FC`);
    }

    // Add quarterly forecasts for next fiscal year
    for (let q = 1; q <= 4; q++) {
        headers.push(`FY${(fiscalYear + 1).toString().slice(-2)} Q${q}`);
    }

    // Add quarterly TO calculations
    // Past quarters (actual) and current/future quarters (forecast)
    for (let q = 1; q <= 4; q++) {
        if (q <= Math.ceil(currentPeriod / 3)) {
            // Past and current quarters - labeled as "Act"
            headers.push(`FY${fiscalYear.toString().slice(-2)} Q${q} Act`);
        } else {
            // Future quarters - labeled as "FC"
            headers.push(`FY${fiscalYear.toString().slice(-2)} Q${q} FC`);
        }
    }

    // Add yearly totals
    headers.push(`FY${fiscalYear.toString().slice(-2)}`);
    headers.push(`FY${(fiscalYear + 1).toString().slice(-2)}`);
    headers.push(`FY${(fiscalYear + 2).toString().slice(-2)} and Following`);

    // Add GM values for all TO values (these come after all TO values)
    // Actual GM values for completed periods
    for (let p = 1; p < currentPeriod; p++) {
        headers.push(`FY${fiscalYear.toString().slice(-2)} P${p} GM`); // Keep GM suffix for internal use
    }

    // Forecast GM values for current and future periods
    for (let p = currentPeriod; p <= 12; p++) {
        headers.push(`FY${fiscalYear.toString().slice(-2)} P${p} GM`); // Keep GM suffix for internal use
    }

    // Quarterly GM values - follow the same pattern as TO quarters
    // Next fiscal year quarterly GM values
    for (let q = 1; q <= 4; q++) {
        headers.push(`FY${(fiscalYear + 1).toString().slice(-2)} Q${q} GM`); // Keep GM suffix for internal use
    }

    // Current fiscal year quarterly GM values
    // Past quarters (actual) and current/future quarters (forecast)
    for (let q = 1; q <= 4; q++) {
        if (q <= Math.ceil(currentPeriod / 3)) {
            // Past and current quarters - labeled as "Act"
            headers.push(`FY${fiscalYear.toString().slice(-2)} Q${q} GM Act`);
        } else {
            // Future quarters - labeled as "FC"
            headers.push(`FY${fiscalYear.toString().slice(-2)} Q${q} GM FC`);
        }
    }

    // Yearly GM values
    headers.push(`FY${fiscalYear.toString().slice(-2)} GM`); // Keep GM suffix for internal use
    headers.push(`FY${(fiscalYear + 1).toString().slice(-2)} GM`); // Keep GM suffix for internal use
    headers.push(`FY${(fiscalYear + 2).toString().slice(-2)} and Following GM`); // Keep GM suffix for internal use

    processedForecastData = [headers];

    // Process each row of merged forecast data
    for (let i = 1; i < mergedForecastData.length; i++) {
        const row = mergedForecastData[i];
        const processedRow = processForecastRow(row, fiscalYear, currentPeriod);
        processedForecastData.push(processedRow);
    }

    // Display processed forecast table
    displayProcessedForecastTable();
    document.getElementById('exportProcessedForecastBtn').disabled = false;
    updateStatus('Forecast data processed successfully', 'success');
}
// Update the processForecastRow function to use customerContractData
function processForecastRow(row, fiscalYear, currentPeriod) {
    // Find column indices in merged forecast data
    const mergedHeaders = mergedForecastData[0];
    const projNumberIndex = mergedHeaders.indexOf('Proj.number');
    const cpmIndex = mergedHeaders.indexOf('CPM');
    const pmIndex = mergedHeaders.indexOf('PM');
    const projDescIndex = mergedHeaders.indexOf('Proj.description');
    const planMarginIndex = mergedHeaders.indexOf('Plan Margin%');
    const beginningBfCklog = getPreviousMonthValue(row[projNumberIndex], 'Beginning BFCklog');

    // Create processed row with initial values
    const processedRow = [
        row[projNumberIndex], // WBS Element
        row[cpmIndex],        // CPM
        row[pmIndex],         // PM
        row[projDescIndex],   // Proj. description
        '',                   // Type (empty)
        beginningBfCklog,     // Beginning BFCklog
        '',                   // Remaining BFCklog (to be calculated)
        row[planMarginIndex], // Plan GM%
        ''                    // Plan GM (to be calculated)
    ];

    // Get previous period data from customer & contract information
    const wbsNumber = row[projNumberIndex];
    const prevPeriodData = customerContractData[wbsNumber] || { to: 0, gm: 0 };

    // Add actual turnover for completed periods
    for (let p = 1; p < currentPeriod; p++) {
        let acValue = '';
        // For the previous month, use data from customer & contract information
        if (p === currentPeriod - 1) {
            acValue = prevPeriodData.to.toFixed(2);
        } else if (previousMonthData.length >= 5) {
            // For older periods, still use previous month data
            const previousMonthRow = findPreviousMonthRow(row[projNumberIndex]);
            if (previousMonthRow && previousMonthRow.length > 9 + (p - 1)) {
                acValue = previousMonthRow[9 + (p - 1)] || '';
            }
        }
        processedRow.push(acValue);
    }

    // Add forecast turnover for current and future periods
    for (let p = currentPeriod; p <= 12; p++) {
        const periodHeader = `TO-F${fiscalYear}P${p}`;
        const periodIndex = mergedHeaders.indexOf(periodHeader);
        const fcValue = (periodIndex >= 0) ? row[periodIndex] : '';
        processedRow.push(fcValue);
    }

    // Add quarterly forecasts for next fiscal year
    for (let q = 1; q <= 4; q++) {
        const quarterHeader = `TO-F${fiscalYear + 1}Q${q}`;
        const quarterIndex = mergedHeaders.indexOf(quarterHeader);
        const quarterValue = (quarterIndex >= 0) ? row[quarterIndex] : '';
        processedRow.push(quarterValue);
    }

    // Calculate Remaining BFCklog (Beginning BFCklog - current fiscal year TO values)
    let currentYearTO = 0;
    
    // Add actual turnover for completed periods (P1 to currentPeriod-1)
    for (let p = 1; p < currentPeriod; p++) {
        let acValue = '';
        // For the previous month, use data from customer & contract information
        if (p === currentPeriod - 1) {
            acValue = prevPeriodData.to.toFixed(2);
        } else if (previousMonthData.length >= 5) {
            // For older periods, still use previous month data
            const previousMonthRow = findPreviousMonthRow(row[projNumberIndex]);
            if (previousMonthRow && previousMonthRow.length > 9 + (p - 1)) {
                acValue = previousMonthRow[9 + (p - 1)] || '';
            }
        }
        currentYearTO += parseFloat(acValue) || 0;
    }
    
    // Add forecast turnover for current and future periods (currentPeriod to P12)
    for (let p = currentPeriod; p <= 12; p++) {
        const periodHeader = `TO-F${fiscalYear}P${p}`;
        const periodIndex = mergedHeaders.indexOf(periodHeader);
        if (periodIndex >= 0) {
            currentYearTO += parseFloat(row[periodIndex]) || 0;
        }
    }
    
    const remainingBfCklog = (parseFloat(beginningBfCklog) || 0) - currentYearTO;
    processedRow[6] = remainingBfCklog.toFixed(2); // Remaining BFCklog

    // Calculate Plan GM (Beginning BFCklog * Plan GM%)
    const planGM = (parseFloat(beginningBfCklog) || 0) * (parseFloat(row[planMarginIndex]) || 0);
    processedRow[8] = planGM.toFixed(2); // Plan GM

    // Add quarterly TO calculations
    // All quarters calculated the same way - sum of months in quarter
    for (let q = 1; q <= 4; q++) {
        let quarterSum = 0;
        const startMonth = (q - 1) * 3 + 1;
        const endMonth = q * 3;
        
        for (let p = startMonth; p <= Math.min(endMonth, 12); p++) {
            if (p < currentPeriod) {
                // Past months - use actual values
                let acValue = '';
                if (p === currentPeriod - 1) {
                    // For the previous month, use data from customer & contract information
                    acValue = prevPeriodData.to.toFixed(2);
                } else if (previousMonthData.length >= 5) {
                    // For older periods, still use previous month data
                    const previousMonthRow = findPreviousMonthRow(row[projNumberIndex]);
                    if (previousMonthRow && previousMonthRow.length > 9 + (p - 1)) {
                        acValue = previousMonthRow[9 + (p - 1)] || '';
                    }
                }
                quarterSum += parseFloat(acValue) || 0;
            } else {
                // Current and future months - use forecast values
                const periodHeader = `TO-F${fiscalYear}P${p}`;
                const periodIndex = mergedHeaders.indexOf(periodHeader);
                quarterSum += parseFloat(row[periodIndex]) || 0;
            }
        }
        processedRow.push(quarterSum.toFixed(2));
    }

    // Add yearly totals
    // Current fiscal year
    let fyTotal = 0;
    for (let p = 1; p < currentPeriod; p++) {
        // Past months - actual values
        let acValue = '';
        if (p === currentPeriod - 1) {
            // For the previous month, use data from customer & contract information
            acValue = prevPeriodData.to.toFixed(2);
        } else if (previousMonthData.length >= 5) {
            // For older periods, still use previous month data
            const previousMonthRow = findPreviousMonthRow(row[projNumberIndex]);
            if (previousMonthRow && previousMonthRow.length > 9 + (p - 1)) {
                acValue = previousMonthRow[9 + (p - 1)] || '';
            }
        }
        fyTotal += parseFloat(acValue) || 0;
    }
    
    for (let p = currentPeriod; p <= 12; p++) {
        // Future months - forecast values
        const periodHeader = `TO-F${fiscalYear}P${p}`;
        const periodIndex = mergedHeaders.indexOf(periodHeader);
        fyTotal += parseFloat(row[periodIndex]) || 0;
    }
    processedRow.push(fyTotal.toFixed(2));

    // Next fiscal year (sum of quarters)
    let fy1Total = 0;
    for (let q = 1; q <= 4; q++) {
        const quarterHeader = `TO-F${fiscalYear + 1}Q${q}`;
        const quarterIndex = mergedHeaders.indexOf(quarterHeader);
        fy1Total += parseFloat(row[quarterIndex]) || 0;
    }
    processedRow.push(fy1Total.toFixed(2));

    // Following fiscal year (from last column in merged data)
    const lastTOIndex = mergedHeaders.lastIndexOf('TO-F' + (fiscalYear + 2));
    const fy2Value = (lastTOIndex >= 0) ? row[lastTOIndex] : '';
    processedRow.push(fy2Value);

    // Add GM values for all TO values
    // Actual GM values for completed periods
    for (let p = 1; p < currentPeriod; p++) {
        let gmValue = '';
        // For the previous month, use data from customer & contract information
        if (p === currentPeriod - 1) {
            gmValue = prevPeriodData.gm.toFixed(2);
        } else if (previousMonthData.length >= 5) {
            // For older periods, still use previous month data
            const previousMonthRow = findPreviousMonthRow(row[projNumberIndex]);
            if (previousMonthRow && previousMonthRow.length > 9 + 23 + (p - 1)) {
                gmValue = previousMonthRow[9 + 23 + (p - 1)] || '';
            }
        }
        processedRow.push(gmValue);
    }

    // Forecast GM values for current and future periods (TO * Plan GM%)
    const planGMRate = parseFloat(row[planMarginIndex]) || 0;
    
    // Forecast periods (current and future)
    for (let p = currentPeriod; p <= 12; p++) {
        const periodHeader = `TO-F${fiscalYear}P${p}`;
        const periodIndex = mergedHeaders.indexOf(periodHeader);
        const toValue = (periodIndex >= 0) ? parseFloat(row[periodIndex]) || 0 : 0;
        const gmValue = toValue * planGMRate;
        processedRow.push(gmValue.toFixed(2));
    }

    // Quarterly GM values for next fiscal year
    for (let q = 1; q <= 4; q++) {
        const quarterHeader = `TO-F${fiscalYear + 1}Q${q}`;
        const quarterIndex = mergedHeaders.indexOf(quarterHeader);
        const toValue = (quarterIndex >= 0) ? parseFloat(row[quarterIndex]) || 0 : 0;
        const gmValue = toValue * planGMRate;
        processedRow.push(gmValue.toFixed(2));
    }

    // Quarterly GM values - all quarters
    let fyGMTotal = 0; // Keep track of the FY total for consistency
    for (let q = 1; q <= 4; q++) {
        let quarterSum = 0;
        const startMonth = (q - 1) * 3 + 1;
        const endMonth = q * 3;
        
        for (let p = startMonth; p <= Math.min(endMonth, 12); p++) {
            let value = 0;
            if (p < currentPeriod) {
                // Past months - use actual GM values
                if (p === currentPeriod - 1) {
                    // For the previous month, use data from customer & contract information
                    value = prevPeriodData.gm;
                } else if (previousMonthData.length >= 5) {
                    // For older periods, still use previous month data
                    const previousMonthRow = findPreviousMonthRow(row[projNumberIndex]);
                    if (previousMonthRow && previousMonthRow.length > 9 + 23 + (p - 1)) {
                        value = parseFloat(previousMonthRow[9 + 23 + (p - 1)]) || 0;
                    }
                }
            } else {
                // Current and future months - calculate GM based on TO values
                const periodHeader = `TO-F${fiscalYear}P${p}`;
                const periodIndex = mergedHeaders.indexOf(periodHeader);
                const toValue = (periodIndex >= 0) ? parseFloat(row[periodIndex]) || 0 : 0;
                value = toValue * planGMRate;
            }
            quarterSum += value;
        }
        
        const gmValue = quarterSum;
        processedRow.push(gmValue.toFixed(2));
        fyGMTotal += gmValue; // Accumulate for consistent yearly total
    }

    // Yearly GM values
    // Current fiscal year GM - use accumulated quarterly values for consistency
    processedRow.push(fyGMTotal.toFixed(2));

    // Next fiscal year GM
    const nextFYGM = fy1Total * planGMRate;
    processedRow.push(nextFYGM.toFixed(2));

    // Following fiscal year GM
    const followingFYGM = (parseFloat(fy2Value) || 0) * planGMRate;
    processedRow.push(followingFYGM.toFixed(2));

    return processedRow;
}

// Helper function to find a row in previous month data for a given project number
function findPreviousMonthRow(projNumber) {
    if (previousMonthData.length < 5) return null;
    
    const headers = previousMonthData[3];
    let wbsElementIndex = headers.indexOf('WBS Element');
    
    if (wbsElementIndex === -1) {
        const possibleNames = ['WBS', 'Project Number', 'Proj.number', 'WBS Element', 'WBS Element '];
        for (const name of possibleNames) {
            const index = headers.indexOf(name);
            if (index !== -1) {
                wbsElementIndex = index;
                break;
            }
        }
    }
    
    if (wbsElementIndex !== -1) {
        for (let i = 4; i < previousMonthData.length; i++) {
            const prevRow = previousMonthData[i];
            if (prevRow && prevRow[wbsElementIndex] && 
                prevRow[wbsElementIndex].toString().trim() === projNumber.toString().trim()) {
                return prevRow;
            }
        }
    }
    
    return null;
}

// You can add this temporary debugging code to see what columns are available
function debugPreviousMonthColumns() {
    if (previousMonthData.length < 4) return;
    
    const headers = previousMonthData[3];
    console.log("Previous month columns:");
    headers.forEach((header, index) => {
        console.log(`${index}: "${header}"`);
    });
    
    // Also log some sample data
    if (previousMonthData.length > 4) {
        console.log("\nSample data row (index 4):");
        previousMonthData[4].forEach((value, index) => {
            console.log(`${index}: "${value}"`);
        });
    }
}
function displayProcessedForecastTable() {
    const table = document.getElementById('processedForecastTable');
    table.innerHTML = '';

    if (processedForecastData.length === 0) return;

    // Create table header
    const thead = document.createElement('thead');
    const headerRow = document.createElement('tr');

    processedForecastData[0].forEach(headerText => {
        const th = document.createElement('th');
        th.textContent = headerText;
        headerRow.appendChild(th);
    });

    thead.appendChild(headerRow);
    table.appendChild(thead);

    // Create table body
    const tbody = document.createElement('tbody');

    for (let i = 1; i < processedForecastData.length; i++) {
        const row = document.createElement('tr');

        processedForecastData[i].forEach(cellData => {
            const td = document.createElement('td');
            td.textContent = cellData !== undefined ? cellData : '';
            row.appendChild(td);
        });

        tbody.appendChild(row);
    }

    table.appendChild(tbody);
}

function exportProcessedForecastData() {
    if (processedForecastData.length === 0) return;

    const worksheet = XLSX.utils.aoa_to_sheet(processedForecastData);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Processed Forecast");

    // Get fiscal year and period
    const fiscalYear = document.getElementById('fiscalYear').value;
    const currentPeriod = document.getElementById('currentPeriod').value;

    XLSX.writeFile(workbook, `Processed_Forecast_F${fiscalYear}P${currentPeriod}.xlsx`);

    updateStatus('Processed forecast data exported successfully', 'success');
}
function handleForecastFileSelection(e) {
    forecastFiles = Array.from(e.target.files);
    document.getElementById('mergeForecastBtn').disabled = forecastFiles.length === 0;
}

// Update the end of mergeForecastFiles function

// Replace the mergeForecastFiles function with this corrected version

async function mergeForecastFiles() {
    if (forecastFiles.length === 0) return;

    mergedForecastData = [];
    const headers = [
        'CPM', 'Level', 'PM', 'Proj.number', 'Proj.description', 'New order', 'POC%', 
        'Plan Cost', 'Committed Cost', 'Actual Cost', 'Cost Remain'
        // Dynamic headers will be added based on the first file
    ];
    
    try {
        // Process all files
        for (let i = 0; i < forecastFiles.length; i++) {
            const file = forecastFiles[i];
            
            // Extract CPM name and fiscal info from filename
            const fileInfo = extractFileInfo(file.name);
            if (!fileInfo) {
                console.warn(`Skipping file with invalid name format: ${file.name}`);
                continue;
            }
            
            const { cpm, fiscalYear, period } = fileInfo;
            
            // Read file data
            const data = await readExcelFile(file);
            
            if (data.length > 0) {
                // Get headers from first file
                if (mergedForecastData.length === 0) {
                    // Add fixed headers
                    const fixedHeaders = [...headers];
                    
                    // Find dynamic headers (COST-* and TO-*)
                    const firstRow = data[0];
                    const dynamicHeaders = firstRow.filter(header => 
                        typeof header === 'string' && 
                        (header.startsWith('COST-') || header.startsWith('TO-'))
                    );
                    
                    // Always add Plan Margin% and Actual Margin% at the end
                    const marginHeaders = firstRow.filter(header =>
                        header === 'Plan Margin%' || header === 'Actual Margin%'
                    );
                    
                    // Combine all headers
                    mergedForecastData.push([...fixedHeaders, ...dynamicHeaders, ...marginHeaders]);
                }
                
                // Process data rows
                const startRow = isHeaderRow(data[0]) ? 1 : 0;
                
                for (let j = startRow; j < data.length; j++) {
                    const row = data[j];
                    if (row.length > 0) {
                        // Only include parent rows (Level = 1)
                        const levelIndex = data[0].indexOf('Level');
                        if (levelIndex >= 0 && row[levelIndex] == 1) { // Use == instead of === to handle string numbers
                            // Add CPM as first column
                            const processedRow = [cpm, ...row];
                            mergedForecastData.push(processedRow);
                        }
                    }
                }
            }
        }
        
        // Display merged forecast table
        displayMergedForecastTable();
        document.getElementById('exportMergedForecastBtn').disabled = false;
        updateStatus(`Successfully merged ${forecastFiles.length} forecast files`, 'success');
        
        // Enable process button if we also have previous month data
        if (previousMonthData.length >= 5) {
            document.getElementById('processForecastBtn').disabled = false;
        }
    } catch (error) {
        updateStatus(`Error merging forecast files: ${error.message}`, 'error');
        console.error(error);
    }
}

function extractFileInfo(filename) {
    // Expected format: Forecast Sheet_LiMei Jing_ChenLi_F2025P8.xlsx
    const regex = /Forecast Sheet_([^_]+)_[^_]+_F(\d+)P(\d+)/;
    const match = filename.match(regex);
    
    if (match) {
        return {
            cpm: match[1],
            fiscalYear: match[2],
            period: match[3]
        };
    }
    
    return null;
}

function displayMergedForecastTable() {
    const table = document.getElementById('mergedForecastTable');
    table.innerHTML = '';
    
    if (mergedForecastData.length === 0) return;
    
    // Create table header
    const thead = document.createElement('thead');
    const headerRow = document.createElement('tr');
    
    mergedForecastData[0].forEach(headerText => {
        const th = document.createElement('th');
        th.textContent = headerText;
        headerRow.appendChild(th);
    });
    
    thead.appendChild(headerRow);
    table.appendChild(thead);
    
    // Create table body
    const tbody = document.createElement('tbody');
    
    for (let i = 1; i < mergedForecastData.length; i++) {
        const row = document.createElement('tr');
        
        mergedForecastData[i].forEach(cellData => {
            const td = document.createElement('td');
            td.textContent = cellData !== undefined ? cellData : '';
            row.appendChild(td);
        });
        
        tbody.appendChild(row);
    }
    
    table.appendChild(tbody);
}

function exportMergedForecastData() {
    if (mergedForecastData.length === 0) return;
    
    const worksheet = XLSX.utils.aoa_to_sheet(mergedForecastData);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Merged Forecast");
    
    // Get fiscal year and period from turnover forecast
    const fiscalYear = document.getElementById('fiscalYear').value;
    const currentPeriod = document.getElementById('currentPeriod').value;
    
    XLSX.writeFile(workbook, `Merged_Forecast_F${fiscalYear}P${currentPeriod}.xlsx`);
    
    updateStatus('Merged Forecast exported successfully', 'success');
}
// 合并现金流表格函数
async function mergeCashflowFiles() {
    const files = document.getElementById('cashflowFiles').files;
    if (files.length === 0) return;

    mergedCashflowData = [];
    const headers = [
        'WBS', 'Project Name', 'CPM', 'PM', 'Customer Name', 
        'Contract Value', 'Payment Terms', 'Percentage Payment Type', 'Amount', 
        'Currency', 'Via', 'Collection Date in Contract', 
        'FC Period', 'Amount in RMB', 'Actual Received Period', 'comments'
    ];
    
    mergedCashflowData.push(headers);
    
    try {
        // 处理所有文件
        for (let i = 0; i < files.length; i++) {
            const file = files[i];
            const data = await readExcelFile(file);
            
            // 确保数据格式正确
            if (data.length > 0) {
                // 如果第一行是标题行，则从第二行开始处理
                const startRow = isHeaderRow(data[0]) ? 1 : 0;
                
                for (let j = startRow; j < data.length; j++) {
                    // 确保每行数据长度与标题匹配
                    const row = data[j];
                    if (row.length >= headers.length) {
                        mergedCashflowData.push(row.slice(0, headers.length));
                    } else if (row.length > 0) {
                        // 如果数据不足，填充空值
                        const newRow = row.concat(new Array(headers.length - row.length).fill(''));
                        mergedCashflowData.push(newRow);
                    }
                }
            }
        }
        
        // 显示合并后的表格
        displayMergedCashflowTable();
        document.getElementById('exportMergedCashflowBtn').disabled = false;
        updateStatus(`Successful Merging ${files.length} Files`, 'success');
    } catch (error) {
        updateStatus(`Errors in Merging: ${error.message}`, 'error');
        console.error(error);
    }
}

// 读取Excel文件函数
function readExcelFile(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = function(e) {
            try {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { type: 'array' });
                const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
                const jsonData = XLSX.utils.sheet_to_json(firstSheet, { header: 1 });
                resolve(jsonData);
            } catch (error) {
                reject(error);
            }
        };
        reader.onerror = reject;
        reader.readAsArrayBuffer(file);
    });
}

// 检查是否是标题行
function isHeaderRow(row) {
    if (!row || row.length === 0) return false;
    // 简单的检查：如果第一列是"WBS"，则认为是标题行
    return row[0] === 'WBS';
}

// 修改displayMergedCashflowTable函数中的排序绑定
function displayMergedCashflowTable() {
    const table = document.getElementById('mergedCashflowTable');
    table.innerHTML = '';
    
    if (mergedCashflowData.length === 0) return;
    
    // 创建表头
    const thead = document.createElement('thead');
    const headerRow = document.createElement('tr');
    
    mergedCashflowData[0].forEach((headerText, index) => {
        const th = document.createElement('th');
        th.textContent = headerText;
        th.style.cursor = 'pointer';
        
        // 添加排序指示箭头
        if (index === currentSortColumn) {
            th.classList.add(sortDirection === 'asc' ? 'sorted-asc' : 'sorted-desc');
        }
        
        th.addEventListener('click', () => {
            // 如果点击的是已排序的列，则切换排序方向
            if (currentSortColumn === index) {
                sortDirection = sortDirection === 'asc' ? 'desc' : 'asc';
            } else {
                currentSortColumn = index;
                sortDirection = 'asc'; // 默认升序
            }
            sortTable(index, sortDirection);
        });
        
        headerRow.appendChild(th);
    });
    
    thead.appendChild(headerRow);
    table.appendChild(thead);
    
    // 创建表格内容
    const tbody = document.createElement('tbody');
    for (let i = 1; i < mergedCashflowData.length; i++) {
        const row = document.createElement('tr');
        
        mergedCashflowData[i].forEach(cellData => {
            const td = document.createElement('td');
            td.textContent = cellData !== undefined ? cellData : '';
            row.appendChild(td);
        });
        
        tbody.appendChild(row);
    }
    
    table.appendChild(tbody);
}
// 完整的排序函数
function sortTable(columnIndex, direction) {
    if (mergedCashflowData.length < 2) return;
    
    const header = mergedCashflowData[0][columnIndex];
    const dataRows = mergedCashflowData.slice(1);
    
    dataRows.sort((a, b) => {
        const aValue = a[columnIndex] || '';
        const bValue = b[columnIndex] || '';
        
        if (header === 'WBS') {
            return sortWBS(aValue, bValue, direction);
        } else {
            return sortAlphabetical(aValue, bValue, direction);
        }
    });
    
    // 更新数据
    mergedCashflowData = [mergedCashflowData[0], ...dataRows];
    
    // 重新显示表格
    displayMergedCashflowTable();
}


// WBS排序逻辑（保持不变）
function sortWBS(a, b, direction) {
    // 处理空值
    if (!a && !b) return 0;
    if (!a) return direction === 'asc' ? -1 : 1;
    if (!b) return direction === 'asc' ? 1 : -1;
    
    // 分离数字开头的WBS和非标准WBS
    const aIsNumeric = /^\d/.test(a);
    const bIsNumeric = /^\d/.test(b);
    
    if (aIsNumeric && !bIsNumeric) return direction === 'asc' ? -1 : 1;
    if (!aIsNumeric && bIsNumeric) return direction === 'asc' ? 1 : -1;
    if (aIsNumeric && bIsNumeric) {
        // 数字开头的WBS按数字大小排序
        const aNum = parseInt(a.match(/^\d+/)[0]);
        const bNum = parseInt(b.match(/^\d+/)[0]);
        return direction === 'asc' ? aNum - bNum : bNum - aNum;
    }
    
    // 处理标准样例数据 E5PP-V-0082
    const aParts = a.split('-');
    const bParts = b.split('-');
    
    if (aParts.length < 3 || bParts.length < 3) {
        // 非标准格式，按字母顺序排序
        return sortAlphabetical(a, b, direction);
    }
    
    // 比较前缀
    const prefixOrder = ['E5OC', 'E5OP', 'E5PP'];
    const aPrefixIndex = prefixOrder.indexOf(aParts[0]);
    const bPrefixIndex = prefixOrder.indexOf(bParts[0]);
    
    if (aPrefixIndex !== bPrefixIndex) {
        return direction === 'asc' ? aPrefixIndex - bPrefixIndex : bPrefixIndex - aPrefixIndex;
    }
    
    // 前缀相同，比较后缀数字
    const aSuffix = parseInt(aParts[2]);
    const bSuffix = parseInt(bParts[2]);
    
    return direction === 'asc' ? aSuffix - bSuffix : bSuffix - aSuffix;
}

// 字母顺序排序（保持不变）
function sortAlphabetical(a, b, direction) {
    const aStr = String(a).toLowerCase();
    const bStr = String(b).toLowerCase();
    
    if (aStr < bStr) return direction === 'asc' ? -1 : 1;
    if (aStr > bStr) return direction === 'asc' ? 1 : -1;
    return 0;
}


function exportMergedCashflowData() {
    if (mergedCashflowData.length === 0) return;
    
    const worksheet = XLSX.utils.aoa_to_sheet(mergedCashflowData);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Merged Cashflow");
    
    // Get fiscal year and period from turnover forecast
    const fiscalYear = document.getElementById('fiscalYear').value;
    const currentPeriod = document.getElementById('currentPeriod').value;
    
    XLSX.writeFile(workbook, `Overall_Cashin_Forecast_F${fiscalYear}P${currentPeriod}.xlsx`);
    
    updateStatus('Merged Cashin Table Exported', 'success');
}
