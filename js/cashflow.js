let cleanedOverallData = [];
let cashflowData = [];
let previousCashinData = [];
let originalCashflowData = [];
let originalCashflowOrder = [];

let filteredWbsNumbers = []; // To store currently filtered WBS numbers
let isFiltered = false; // To track if we're in filtered mode
// Add this with the other global variables at the top of the file
let addedWbsProjects = new Set(); // To track which WBS projects have been added

// Add these variables at the top with other global variables in cashflow.js
let standardizedInputMode = false;
// Replace your existing DOMContentLoaded handler with this one:
document.addEventListener('DOMContentLoaded', function() {
    document.getElementById('processCashflowBtn').addEventListener('click', processCashflowData);
    document.getElementById('exportCashflowBtn').addEventListener('click', exportCashflowData);
    // Fix the function names to match the actual functions you've defined
    document.getElementById('applyPmFilterBtn').addEventListener('click', applyPmFilterCashflow);
    document.getElementById('resetPmFilterBtn').addEventListener('click', resetPmFilterCashflow);
});
// Apply PM filter - move selected PM's projects to the top
function applyPmFilterCashflow() {
    const pmFilter = document.getElementById('pmFilterCashflow');
    const selectedPm = pmFilter.value;
    
    if (!selectedPm) {
        updateStatus('Please select a PM', 'error');
        return;
    }
    
    // Save original order if not already saved
    if (originalCashflowOrder.length === 0) {
        originalCashflowOrder = [...cashflowData];
    }
    
    // Get headers and PM index
    const headers = cashflowData[0];
    const pmIndex = headers.indexOf('PM');
    
    if (pmIndex === -1) {
        updateStatus('PM column not found', 'error');
        return;
    }
    
    // Separate rows by PM
    const pmRows = [];
    const otherRows = [];
    
    for (let i = 1; i < cashflowData.length; i++) {
        const row = cashflowData[i];
        if (row[pmIndex] === selectedPm) {
            pmRows.push(row);
        } else {
            otherRows.push(row);
        }
    }
    
    // Reconstruct cashflowData with filtered PM rows at the top
    cashflowData = [headers, ...pmRows, ...otherRows];
    
    // Redisplay the table
    displayCashflowTable();
    
    updateStatus(`Moved ${pmRows.length} projects by ${selectedPm} to the top`, 'success');
}

// Reset PM filter and restore original order
function resetPmFilterCashflow() {
    const pmFilter = document.getElementById('pmFilterCashflow');
    pmFilter.value = '';
    
    // Restore original order if it was saved
    if (originalCashflowOrder.length > 0) {
        cashflowData = [...originalCashflowOrder];
        originalCashflowOrder = [];
    }
    
    // Redisplay the table
    displayCashflowTable();
    
    updateStatus('PM filter cleared and original order restored', 'success');
}
// Clean Overall Information data
function cleanOverallData(data) {
    if (data.length === 0) return [];
    
    const headers = data[0];
    const soNumberIndex = 15; // 第16列 (0-based)
    const customerNameIndex = 19; // 第20列
    const orValueIndex = 35; // 第36列
    
    // 只保留需要的列
    const cleanedData = [['WBS Number', 'Customer Name', 'OR']];
    
    // 用于跟踪 WBS Number 第一次出现的位置
    const wbsMap = {};
    
    for (let i = 1; i < data.length; i++) {
        const row = data[i];
        const wbsNumber = row[soNumberIndex];
        const customerName = row[customerNameIndex];
        const orValue = parseFloat(row[orValueIndex]) || 0;
        
        if (!wbsNumber) continue;
        
        if (wbsMap.hasOwnProperty(wbsNumber)) {
            // 非首次出现，将 OR 值加到第一次出现的行
            const firstIndex = wbsMap[wbsNumber];
            cleanedData[firstIndex][2] = (parseFloat(cleanedData[firstIndex][2]) + orValue).toFixed(2);
        } else {
            // 首次出现，添加新行
            wbsMap[wbsNumber] = cleanedData.length;
            cleanedData.push([wbsNumber, customerName, orValue.toFixed(2)]);
        }
    }
    
    return cleanedData;
}

// Modify the processCashflowData function to ensure export button visibility
function processCashflowData() {
    // 获取首页上传的 WBS 数据
    const wbsData = filteredData.length > 0 ? filteredData : cleanedData;
    if (wbsData.length === 0) {
        updateStatus('请先在首页上传并处理WBS数据', 'error');
        return;
    }
    
    // 处理 Previous Cashin Forecast
    const previousCashinFile = document.getElementById('previousCashinFile').files[0];
    if (!previousCashinFile) {
        updateStatus('Please Upload Previous Cashin Forecast File', 'error');
        return;
    }
    
    // 处理 Overall Information
    const overallInfoFile = document.getElementById('overallInfoFile').files[0];
    if (!overallInfoFile) {
        updateStatus('Please Upload Overall Information File', 'error');
        return;
    }
    
    // 读取并处理 Previous Cashin Forecast
    const reader1 = new FileReader();
    reader1.onload = function(e) {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
        previousCashinData = XLSX.utils.sheet_to_json(firstSheet, { header: 1 });
        
        // 读取并处理 Overall Information
        const reader2 = new FileReader();
        reader2.onload = function(e) {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
            const overallData = XLSX.utils.sheet_to_json(firstSheet, { header: 1 });
            
            // 清洗 Overall Information 数据
            cleanedOverallData = cleanOverallData(overallData);
            displayTable(document.getElementById('cleanedOverallTable'), cleanedOverallData);
            
            // 合并数据生成现金流分析
            generateCashflowAnalysis(wbsData, cleanedOverallData, previousCashinData);
            
            // Ensure export button is visible after processing
            document.getElementById('exportCashflowBtn').disabled = false;
            document.getElementById('exportCashflowBtn').style.display = 'inline-block';
            
            updateStatus('Cashin Analysis Created', 'success');
        };
        reader2.readAsArrayBuffer(overallInfoFile);
    };
    reader1.readAsArrayBuffer(previousCashinFile);
}

// Modify the showAdditionalButtons function to properly manage button states
function showAdditionalButtons() {
    // Create and show Add Project button
    let addProjectBtn = document.getElementById('addProjectBtn');
    if (!addProjectBtn) {
        addProjectBtn = document.createElement('button');
        addProjectBtn.id = 'addProjectBtn';
        addProjectBtn.textContent = 'Add Project';
        addProjectBtn.className = 'btn';
        addProjectBtn.style.marginLeft = '10px';
        addProjectBtn.addEventListener('click', addMoreProjects);
        document.getElementById('processCashflowBtn').parentElement.appendChild(addProjectBtn);
    }
    
    // Create and show Clear Filter button
    let clearFilterBtn = document.getElementById('clearFilterBtn');
    if (!clearFilterBtn) {
        clearFilterBtn = document.createElement('button');
        clearFilterBtn.id = 'clearFilterBtn';
        clearFilterBtn.textContent = 'Clear Filter';
        clearFilterBtn.className = 'btn';
        clearFilterBtn.style.marginLeft = '10px';
        clearFilterBtn.addEventListener('click', clearFilter);
        document.getElementById('processCashflowBtn').parentElement.appendChild(clearFilterBtn);
    }
    
    // Hide original process button but keep export button visible
    document.getElementById('processCashflowBtn').style.display = 'none';
    // Do not hide export button - keep it visible
    document.getElementById('exportCashflowBtn').style.display = 'inline-block';
    
    // Show new buttons
    addProjectBtn.style.display = 'inline-block';
    clearFilterBtn.style.display = 'inline-block';
}

// Update clearFilter function to properly restore original buttons
function clearFilter() {
    isFiltered = false;
    filteredWbsNumbers = [];
    addedWbsProjects.clear(); // Clear the set of added projects
    
    // Hide the add project and clear filter buttons
    const addProjectBtn = document.getElementById('addProjectBtn');
    const clearFilterBtn = document.getElementById('clearFilterBtn');
    if (addProjectBtn) addProjectBtn.style.display = 'none';
    if (clearFilterBtn) clearFilterBtn.style.display = 'none';
    
    // Show all original buttons
    document.getElementById('processCashflowBtn').style.display = 'inline-block';
    document.getElementById('exportCashflowBtn').style.display = 'inline-block';
    
    updateStatus('Filter cleared', 'success');
}



function generateCashflowAnalysis(wbsData, overallData, cashinData) {
    // Reset the added projects set when starting a new analysis
    addedWbsProjects.clear();
    
    // 1. Create new cashin forecast headers
    const cashflowHeaders = [
        'WBS', 'Project Name', 'CPM', 'PM', 'Customer Name', 
        'Contract Value', 'Payment Terms', 'Percentage', 'Payment Type', 'Amount', 
        'Currency', 'Via', 'Collection Date in Contract', 
        'FC Period', 'Amount in RMB', 'Actual Received Period', 'comments'
    ];
    
    // 2. Create new cashin forecast data
    cashflowData = [cashflowHeaders];
    
    // 3. Get indexes from raw WBS data
    const rawWbsHeaders = rawData[0];
    const rawWbsNumberIndex = rawWbsHeaders.indexOf('WBS Element');
    const rawProjectNameIndex = rawWbsHeaders.indexOf('WBS Description');
    const rawPersonIndex = rawWbsHeaders.indexOf('Person Respons.');
    const rawApplicantIndex = rawWbsHeaders.indexOf('Applicant');
    
    // 4. Find WBS numbers in Overall Information but not in Previous Cashin Forecast
    const previousWbsNumbers = new Set();
    if (cashinData.length > 0) {
        const cashinHeaders = cashinData[0];
        const cashinWbsIndex = cashinHeaders.indexOf('WBS');
        const cashinCpmIndex = cashinHeaders.indexOf('CPM');
        
        if (cashinWbsIndex !== -1) {
            for (let i = 1; i < cashinData.length; i++) {
                const wbs = cashinData[i][cashinWbsIndex];
                if (wbs) previousWbsNumbers.add(wbs);
                // Also add to added projects since they're already in the table
                addedWbsProjects.add(wbs);
            }
        }
    }
    
    const newWbsNumbers = [];
    for (let i = 1; i < overallData.length; i++) {
        const wbs = overallData[i][0];
        if (wbs && !previousWbsNumbers.has(wbs)) {
            // Find applicant in rawData
            let applicant = '';
            for (let j = 1; j < rawData.length; j++) {
                if (rawData[j][rawWbsNumberIndex] === wbs) {
                    applicant = rawData[j][rawApplicantIndex];
                    break;
                }
            }
            
            newWbsNumbers.push({
                wbs: wbs,
                customer: overallData[i][1],
                or: overallData[i][2],
                applicant: applicant
            });
        }
    }
    
    // 5. Show WBS selection modal with applicant filtering
    if (newWbsNumbers.length > 0) {
        showWbsSelectionModal(newWbsNumbers, function(selectedWbs, selectedApplicant) {
            // 6. Add selected WBS projects
            if (selectedWbs.length > 0) {
                addSelectedProjects(selectedWbs);
            }
            
            // 7. Add filtered Previous Cashin Forecast data
            addPreviousCashinData(cashinData, cashflowHeaders, selectedApplicant);
            
            // 8. Display the table
            displayCashflowTable();
            
            // Store filtered data if a filter was applied
            if (selectedApplicant) {
                isFiltered = true;
                filteredWbsNumbers = newWbsNumbers.filter(item => item.applicant === selectedApplicant);
                
                // Show add project and clear filter buttons
                showAdditionalButtons();
            }
        });
    } else {
        // No new WBS numbers, just add previous data
        addPreviousCashinData(cashinData, cashflowHeaders);
        displayCashflowTable();
    }
        // At the end of the function, after displayCashflowTable(), add:
        setTimeout(() => {
            populatePmFilterCashflow();
        }, 0);

}


function addPreviousCashinData(cashinData, cashflowHeaders, filterApplicant = '') {
    if (cashinData.length > 0) {
        const cashinHeaders = cashinData[0];
        
        // Create column index mapping
        const columnMap = {};
        cashflowHeaders.forEach((targetHeader, targetIndex) => {
            // Try exact match first
            let sourceIndex = cashinHeaders.findIndex(h => h === targetHeader);
            
            // If not found, try flexible matching
            if (sourceIndex === -1) {
                const lowerTarget = targetHeader.toLowerCase();
                sourceIndex = cashinHeaders.findIndex(h => 
                    h && h.toLowerCase().includes(lowerTarget));
            }
            
            if (sourceIndex !== -1) {
                columnMap[targetIndex] = sourceIndex;
            }
        });
        
        // Special handling for specific columns
        const specialMappings = {
            'Collection Date in Contract': ['Collection Date', 'Contract Date'],
            'FC Period': ['Forecast Period', 'Period'],
            'Amount in RMB': ['Amount RMB', 'RMB Amount'],
            'Actual Received Period': ['Actual Period', 'Received Period'],
            'comments': ['comment', 'notes']
        };
        
        // Handle special mappings
        Object.keys(specialMappings).forEach(targetCol => {
            const targetIndex = cashflowHeaders.indexOf(targetCol);
            if (targetIndex !== -1 && !columnMap[targetIndex]) {
                for (const alias of specialMappings[targetCol]) {
                    const sourceIndex = cashinHeaders.findIndex(h => 
                        h && h.toLowerCase() === alias.toLowerCase());
                    if (sourceIndex !== -1) {
                        columnMap[targetIndex] = sourceIndex;
                        break;
                    }
                }
            }
        });
        
        // Find CPM index in previous data
        const cpmIndex = cashinHeaders.findIndex(h => h === 'CPM');
        
        // Add data rows with filtering
        for (let i = 1; i < cashinData.length; i++) {
            const row = cashinData[i];
            
            // Skip if filtering by CPM and doesn't match
            if (filterApplicant && cpmIndex !== -1 && row[cpmIndex] !== filterApplicant) {
                continue;
            }
            
            const newRow = new Array(cashflowHeaders.length).fill('');
            
            // Map all found columns
            Object.keys(columnMap).forEach(targetIndex => {
                const sourceIndex = columnMap[targetIndex];
                if (row[sourceIndex] !== undefined) {
                    newRow[targetIndex] = row[sourceIndex];
                }
            });
            
            cashflowData.push(newRow);
        }
    }
}



// 显示WBS选择弹窗（改进版，支持初始选择状态）
function showWbsSelectionModal(wbsNumbers, callback, mode = 'initial') {
    // 创建模态框
    const modal = document.createElement('div');
    modal.style.position = 'fixed';
    modal.style.top = '0';
    modal.style.left = '0';
    modal.style.width = '100%';
    modal.style.height = '100%';
    modal.style.backgroundColor = 'rgba(0,0,0,0.5)';
    modal.style.display = 'flex';
    modal.style.justifyContent = 'center';
    modal.style.alignItems = 'center';
    modal.style.zIndex = '1000';
    
    // 创建模态框内容
    const modalContent = document.createElement('div');
    modalContent.style.backgroundColor = 'white';
    modalContent.style.padding = '20px';
    modalContent.style.borderRadius = '5px';
    modalContent.style.width = '600px';
    modalContent.style.maxHeight = '80vh';
    modalContent.style.overflow = 'auto';
    
    // 添加标题
    const title = document.createElement('h3');
    title.textContent = mode === 'add' ? 'Add More Projects' : 'Add WBS Project';
    modalContent.appendChild(title);
    
    // 添加选择列表
    const selectContainer = document.createElement('div');
    selectContainer.style.margin = '15px 0';
    selectContainer.style.maxHeight = '300px';
    selectContainer.style.overflow = 'auto';
    
    // Add row count input for each project
    wbsNumbers.forEach(wbsInfo => {
        const itemDiv = document.createElement('div');
        itemDiv.style.margin = '5px 0';
        itemDiv.style.display = 'flex';
        itemDiv.style.alignItems = 'center';
        itemDiv.style.flexWrap = 'wrap';
        
        const checkbox = document.createElement('input');
        checkbox.type = 'checkbox';
        checkbox.id = `wbs-${wbsInfo.wbs}`;
        checkbox.value = wbsInfo.wbs;
        checkbox.checked = false;
        checkbox.style.marginRight = '10px';
        
        const label = document.createElement('label');
        label.htmlFor = `wbs-${wbsInfo.wbs}`;
        label.textContent = `${wbsInfo.wbs} - ${wbsInfo.customer} (OR: ${wbsInfo.or})`;
        label.style.marginRight = '10px';
        label.style.minWidth = '250px';
        
        const rowCountLabel = document.createElement('span');
        rowCountLabel.textContent = 'Rows:';
        rowCountLabel.style.marginRight = '5px';
        
        const rowCountInput = document.createElement('input');
        rowCountInput.type = 'number';
        rowCountInput.min = '1';
        rowCountInput.value = '1';
        rowCountInput.style.width = '50px';
        rowCountInput.style.marginRight = '10px';
        
        itemDiv.appendChild(checkbox);
        itemDiv.appendChild(label);
        itemDiv.appendChild(rowCountLabel);
        itemDiv.appendChild(rowCountInput);
        selectContainer.appendChild(itemDiv);
    });
    
    modalContent.appendChild(selectContainer);
    
    // 添加全选/全不选按钮
    const selectAllDiv = document.createElement('div');
    selectAllDiv.style.margin = '10px 0';
    
    const selectAllBtn = document.createElement('button');
    selectAllBtn.textContent = 'Select All';
    selectAllBtn.className = 'btn';
    selectAllBtn.style.marginRight = '10px';
    // Replace the existing selectAllBtn.onclick function with this improved version
selectAllBtn.onclick = function() {
    selectContainer.querySelectorAll('input[type="checkbox"]').forEach(checkbox => {
        // Only select checkboxes that are currently visible
        const itemDiv = checkbox.parentElement;
        if (itemDiv.style.display !== 'none') {
            checkbox.checked = true;
        }
    });
};
    
    const deselectAllBtn = document.createElement('button');
    deselectAllBtn.textContent = 'Chose None';
    deselectAllBtn.className = 'btn';
    deselectAllBtn.onclick = function() {
        selectContainer.querySelectorAll('input[type="checkbox"]').forEach(checkbox => {
            checkbox.checked = false;
        });
    };
    
    selectAllDiv.appendChild(selectAllBtn);
    selectAllDiv.appendChild(deselectAllBtn);
    modalContent.appendChild(selectAllDiv);
    
    // 添加按钮
    const buttonDiv = document.createElement('div');
    buttonDiv.style.textAlign = 'right';
    buttonDiv.style.marginTop = '15px';
    
    const confirmBtn = document.createElement('button');
    confirmBtn.textContent = 'confirm';
    confirmBtn.className = 'btn';
    confirmBtn.style.marginRight = '10px';
    
    // Add applicantSelect variable to store reference
    let applicantSelect = null;
    
    // Modify the confirm button handler:
    // Modify confirm button handler to handle row counts
    confirmBtn.onclick = function() {
        const selected = [];
        const checkboxes = selectContainer.querySelectorAll('input[type="checkbox"]:checked');
        const selectedApplicant = applicantSelect ? applicantSelect.value : '';
        
        checkboxes.forEach(checkbox => {
            const rowCountInput = checkbox.parentElement.querySelector('input[type="number"]');
            const rowCount = parseInt(rowCountInput.value) || 1;
            const wbs = wbsNumbers.find(item => item.wbs === checkbox.value);
            
            if (wbs) {
                // Add the project multiple times based on row count
                for (let i = 0; i < rowCount; i++) {
                    selected.push(wbs);
                }
            }
        });
        
        document.body.removeChild(modal);
        callback(selected, selectedApplicant);
    };
    
    const cancelBtn = document.createElement('button');
    cancelBtn.textContent = 'cancel';
    cancelBtn.className = 'btn btn-secondary';
    cancelBtn.onclick = function() {
        document.body.removeChild(modal);
        callback([]);
    };
    
    buttonDiv.appendChild(confirmBtn);
    buttonDiv.appendChild(cancelBtn);
    modalContent.appendChild(buttonDiv);
    
    modal.appendChild(modalContent);
    document.body.appendChild(modal);
    
    // Add applicant filter dropdown only for initial mode
    if (mode === 'initial') {
        const applicantFilterDiv = document.createElement('div');
        applicantFilterDiv.style.margin = '10px 0';
        
        const applicantLabel = document.createElement('label');
        applicantLabel.textContent = 'CPM Filter: ';
        applicantLabel.style.marginRight = '10px';
        
        applicantSelect = document.createElement('select'); // Assign to the variable
        applicantSelect.id = 'modalApplicantFilter';
        applicantSelect.style.marginRight = '10px';
        
        // Get unique applicants from rawData (assuming rawData is available globally)
        const applicantSet = new Set();
        if (rawData.length > 0) {
            const headers = rawData[0];
            const applicantIndex = headers.indexOf('Applicant');
            
            if (applicantIndex !== -1) {
                for (let i = 1; i < rawData.length; i++) {
                    const applicant = rawData[i][applicantIndex];
                    if (applicant) applicantSet.add(applicant);
                }
            }
        }
        
        // Add "All" option
        const allOption = document.createElement('option');
        allOption.value = '';
        allOption.textContent = '-- CPM --';
        applicantSelect.appendChild(allOption);
        
        // Add applicant options
        Array.from(applicantSet).sort().forEach(applicant => {
            const option = document.createElement('option');
            option.value = applicant;
            option.textContent = applicant;
            applicantSelect.appendChild(option);
        });
        
        // Add filter button
        const filterBtn = document.createElement('button');
        filterBtn.textContent = 'Filter';
        filterBtn.className = 'btn';
        filterBtn.onclick = function() {
            const selectedApplicant = applicantSelect.value;
            const checkboxes = selectContainer.querySelectorAll('input[type="checkbox"]');
            
            checkboxes.forEach(checkbox => {
                const wbsNumber = checkbox.value;
                const wbsInfo = wbsNumbers.find(item => item.wbs === wbsNumber);
                
                if (!wbsInfo) return;
                
                // Show/hide based on filter
                const itemDiv = checkbox.parentElement;
                if (!selectedApplicant || wbsInfo.applicant === selectedApplicant) {
                    itemDiv.style.display = 'flex';
                } else {
                    itemDiv.style.display = 'none';
                }
            });
        };
        
        applicantFilterDiv.appendChild(applicantLabel);
        applicantFilterDiv.appendChild(applicantSelect);
        applicantFilterDiv.appendChild(filterBtn);
        modalContent.insertBefore(applicantFilterDiv, selectContainer);
    }
}

// Add project button handler
function addMoreProjects() {
    if (!isFiltered || filteredWbsNumbers.length === 0) return;
    
    // Filter out already added projects
    const availableWbsNumbers = filteredWbsNumbers.filter(wbsInfo => !addedWbsProjects.has(wbsInfo.wbs));
    
    if (availableWbsNumbers.length === 0) {
        updateStatus('No more projects available to add', 'info');
        return;
    }
    
    showWbsSelectionModal(availableWbsNumbers, function(selectedWbs) {
        if (selectedWbs.length > 0) {
            addSelectedProjects(selectedWbs);
            displayCashflowTable();
        }
    }, 'add');
}

// Helper function to add selected projects to cashflowData
function addSelectedProjects(selectedWbs) {
    // Get indexes from raw WBS data
    const rawWbsHeaders = rawData[0];
    const rawWbsNumberIndex = rawWbsHeaders.indexOf('WBS Element');
    const rawProjectNameIndex = rawWbsHeaders.indexOf('WBS Description');
    const rawPersonIndex = rawWbsHeaders.indexOf('Person Respons.');
    const rawApplicantIndex = rawWbsHeaders.indexOf('Applicant');
    
    // Keep track of new rows to be added at the top
    const newRows = [];
    
    selectedWbs.forEach(wbsInfo => {
        // Add to the set of added projects
        addedWbsProjects.add(wbsInfo.wbs);
        
        // Find project info in rawData
        let projectName = '';
        let pm = '';
        let cpm = '';
        
        for (let i = 1; i < rawData.length; i++) {
            const row = rawData[i];
            if (row[rawWbsNumberIndex] === wbsInfo.wbs) {
                projectName = row[rawProjectNameIndex];
                pm = row[rawPersonIndex];
                cpm = row[rawApplicantIndex];
                break;
            }
        }
        
        // Calculate Contract Value (OR * 1.13 rounded)
        const orValue = parseFloat(wbsInfo.or) || 0;
        const contractValue = Math.round(orValue * 1.13);
        
        // Add new row with only specified columns populated
        newRows.push([
            wbsInfo.wbs,        // WBS
            projectName,        // Project Name
            cpm,                // CPM
            pm,                 // PM
            wbsInfo.customer,   // Customer Name
            '',                 // Contract Value (empty)
            '',                 // Payment Terms (empty)
            '',                 // Percentage (empty)
            '',                 // Payment Type (empty)
            contractValue,      // Amount (only this has value)
            'RMB',              // Currency
            '',                 // Via (empty)
            '',                 // Collection Date in Contract (empty)
            '',                 // FC Period (empty)
            '',                 // Amount in RMB (empty)
            '',                 // Actual Received Period (empty)
            ''                  // comments (empty)
        ]);
    });
    
    // Add new rows at the beginning (after headers)
    cashflowData.splice(1, 0, ...newRows);
}

function displayCashflowTable() {
    const container = document.querySelector('.cashflow-table-container');
    const table = document.getElementById('cashflowTable');
    
    // Set container width to match forecast table
    container.style.width = '100%';
    container.style.overflowX = 'auto';
    
    // Clear existing table
    table.innerHTML = '';
    
    // Set table to not exceed container width
    table.style.width = 'auto';
    table.style.minWidth = '100%';
    
    if (cashflowData.length === 0) return;
    
    // Get current fiscal year and period from the forecast page
    const fiscalYear = parseInt(document.getElementById('fiscalYear').value) || 2025;
    const currentPeriod = parseInt(document.getElementById('currentPeriod').value) || 8;
    
    // Generate period options for the next 2 years (24 periods)
    const periodOptions = [];
    let year = fiscalYear;
    let period = currentPeriod;
    
    // Add current and future periods for 2 years
    for (let i = 0; i < 24; i++) {
        periodOptions.push(`FC${year.toString().slice(-2)}P${period.toString().padStart(2, '0')}`);
        
        period++;
        if (period > 12) {
            period = 1;
            year++;
        }
    }
    
    // Define dropdown options for specific columns
    const columnOptions = {
        'Payment Terms': ['DP', 'before delivery', 'after delivery', 'after FAC'],
        'Payment Type': ['Milestone', 'DP', 'Not due AR', 'Overdue AR'],
        'Currency': ['CNY', 'EUR', 'USD'],
        'Via': ['T/T', 'BD', 'T/T+BD'],
        'Collection Date in Contract': periodOptions,
        'FC Period': periodOptions,
        'Actual Received Period': periodOptions
    };
    
    // Create table header
    const thead = document.createElement('thead');
    const headerRow = document.createElement('tr');
    
    cashflowData[0].forEach(headerText => {
        const th = document.createElement('th');
        th.textContent = headerText;
        headerRow.appendChild(th);
    });
    
    thead.appendChild(headerRow);
    table.appendChild(thead);
    
    // Create table body
    const tbody = document.createElement('tbody');
    for (let i = 1; i < cashflowData.length; i++) {
        const row = document.createElement('tr');
        
        cashflowData[i].forEach((cellData, colIndex) => {
            const td = document.createElement('td');
            
            // WBS column is read-only
            if (colIndex === 0) {
                td.textContent = cellData;
                td.classList.add('readonly');
            } else {
                const headerName = cashflowData[0][colIndex];
                
                // Check if this column should have dropdown options
                if (columnOptions[headerName]) {
                    // Create container for input and dropdown button
                    const container = document.createElement('div');
                    container.style.position = 'relative';
                    container.style.width = '100%';
                    container.style.zIndex = 'auto'; // 添加这一行
                    
                    // Create input field
                    const input = document.createElement('input');
                    input.type = 'text';
                    input.value = cellData || '';
                    input.style.width = '100%';
                    input.style.padding = '2px 20px 2px 4px';
                    input.style.fontSize = '12px';
                    input.style.border = '1px solid #ccc';
                    input.style.borderRadius = '3px';
                    
                    // Create dropdown button
                    const dropdownBtn = document.createElement('button');
                    dropdownBtn.innerHTML = '▼';
                    dropdownBtn.style.position = 'absolute';
                    dropdownBtn.style.right = '2px';
                    dropdownBtn.style.top = '50%';
                    dropdownBtn.style.transform = 'translateY(-50%)';
                    dropdownBtn.style.background = 'none';
                    dropdownBtn.style.border = 'none';
                    dropdownBtn.style.cursor = 'pointer';
                    dropdownBtn.style.fontSize = '10px';
                    dropdownBtn.style.padding = '0';
                    dropdownBtn.style.width = '16px';
                    dropdownBtn.style.height = '16px';
                    
                    // Create dropdown list
                    const dropdownList = document.createElement('div');
                    dropdownList.style.position = 'absolute';
                    dropdownList.style.top = '100%';
                    dropdownList.style.left = '0';
                    dropdownList.style.right = '0';
                    dropdownList.style.backgroundColor = 'white';
                    dropdownList.style.border = '1px solid #ccc';
                    dropdownList.style.borderRadius = '3px';
                    dropdownList.style.boxShadow = '0 2px 5px rgba(0,0,0,0.2)';
                    dropdownList.style.zIndex = '9999';
                    dropdownList.style.maxHeight = '150px';
                    dropdownList.style.overflowY = 'auto';
                    dropdownList.style.display = 'none';
                    
                    // Add options to dropdown list
                    columnOptions[headerName].forEach(option => {
                        const optionElement = document.createElement('div');
                        optionElement.textContent = option;
                        optionElement.style.padding = '4px 8px';
                        optionElement.style.cursor = 'pointer';
                        optionElement.style.fontSize = '12px';
                        
                        optionElement.addEventListener('mouseenter', function() {
                            this.style.backgroundColor = '#f0f0f0';
                        });
                        
                        optionElement.addEventListener('mouseleave', function() {
                            this.style.backgroundColor = 'white';
                        });
                        
                        optionElement.addEventListener('click', function() {
                            input.value = option;
                            cashflowData[i][colIndex] = option;
                            dropdownList.style.display = 'none';
                        });
                        
                        dropdownList.appendChild(optionElement);
                    });
                    
                    // Handle input changes
                    input.addEventListener('change', function() {
                        cashflowData[i][colIndex] = this.value;
                    });
                    
                    // Toggle dropdown visibility
                    dropdownBtn.addEventListener('click', function(e) {
                        e.stopPropagation();
                        if (dropdownList.style.display === 'none') {
                            dropdownList.style.display = 'block';
                        } else {
                            dropdownList.style.display = 'none';
                        }
                    });
                    
                    // Close dropdown when clicking elsewhere
                    document.addEventListener('click', function(e) {
                        if (e.target !== input && e.target !== dropdownBtn) {
                            dropdownList.style.display = 'none';
                        }
                    });
                    
                    // Add elements to container
                    container.appendChild(input);
                    container.appendChild(dropdownBtn);
                    container.appendChild(dropdownList);
                    
                    td.appendChild(container);
                } else {
                    // Regular editable input for other columns
                    const input = document.createElement('input');
                    input.type = 'text';
                    input.value = cellData || '';
                    input.style.width = '100%';
                    input.style.padding = '2px';
                    input.style.fontSize = '12px';
                    input.addEventListener('change', function() {
                        cashflowData[i][colIndex] = this.value;
                    });
                    td.appendChild(input);
                }
            }
            
            row.appendChild(td);
        });
        
        tbody.appendChild(row);
    }
    
    table.appendChild(tbody);
    
    // Populate PM filter dropdown after table is displayed
    setTimeout(() => {
        populatePmFilterCashflow();
    }, 0);
}
// 导出现金流分析数据
function exportCashflowData() {
    if (cashflowData.length === 0) return;
    
    const worksheet = XLSX.utils.aoa_to_sheet(cashflowData);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Cashin Forecast");
    
    // Get selected name from the first row (assuming it's in column 2 - "Project Name")
    let selectedName = 'Unknown';
    if (cashflowData.length > 1 && cashflowData[1][2]) {
        selectedName = cashflowData[1][2];
    }
    
    // Get fiscal year and period from turnover forecast
    const fiscalYear = document.getElementById('fiscalYear').value;
    const currentPeriod = document.getElementById('currentPeriod').value;
    
    XLSX.writeFile(workbook, `Cashin_Forecast_${selectedName}_F${fiscalYear}P${currentPeriod}.xlsx`);
    
    updateStatus('Cashin Data Exported', 'success');
}

function initPmFilter() {
    const pmFilter = document.getElementById('pmFilter');
    if (!pmFilter) return;
    
    // Clear existing options
    pmFilter.innerHTML = '<option value="">All PMs</option>';
    
    // Create a set to store unique PMs
    const pmSet = new Set();
    
    // Extract PMs from cashflowData
    if (cashflowData && cashflowData.length > 1) {
        const headers = cashflowData[0];
        const pmIndex = headers.indexOf('PM');
        
        if (pmIndex !== -1) {
            for (let i = 1; i < cashflowData.length; i++) {
                const row = cashflowData[i];
                if (row && row.length > pmIndex) {
                    const pm = row[pmIndex];
                    // Better handling of empty values
                    if (pm != null && pm.toString().trim() !== '') {
                        pmSet.add(pm.toString().trim());
                    }
                }
            }
            
            console.log('Found PMs:', Array.from(pmSet)); // For debugging
        } else {
            console.log('PM column not found in headers'); // For debugging
        }
    } else {
        console.log('No cashflow data available'); // For debugging
    }
    
    // Add unique PMs to the dropdown
    Array.from(pmSet).sort().forEach(pm => {
        const option = document.createElement('option');
        option.value = pm;
        option.textContent = pm;
        pmFilter.appendChild(option);
    });
    
    // Restore the current filter if it exists
    if (currentPmFilter) {
        pmFilter.value = currentPmFilter;
    }
}

// Populate PM filter dropdown with unique PMs
function populatePmFilterCashflow() {
    const pmFilter = document.getElementById('pmFilterCashflow');
    if (!pmFilter) return;
    
    // Clear existing options except the first one
    pmFilter.innerHTML = '<option value="">-- All PMs --</option>';
    
    // Create a set to store unique PMs
    const pmSet = new Set();
    
    // Extract PMs from cashflowData
    if (cashflowData && cashflowData.length > 1) {
        const headers = cashflowData[0];
        const pmIndex = headers.indexOf('PM');
        
        if (pmIndex !== -1) {
            for (let i = 1; i < cashflowData.length; i++) {
                const row = cashflowData[i];
                if (row && row.length > pmIndex) {
                    const pm = row[pmIndex];
                    if (pm != null && pm.toString().trim() !== '') {
                        pmSet.add(pm.toString().trim());
                    }
                }
            }
        }
    }
    
    // Add unique PMs to the dropdown
    Array.from(pmSet).sort().forEach(pm => {
        const option = document.createElement('option');
        option.value = pm;
        option.textContent = pm;
        pmFilter.appendChild(option);
    });
}
