// Global variables
let selectedFiles = [];
let processedData = {
    sheets: [],
    combinedWB: null,
    processingStats: {
        startTime: 0,
        endTime: 0
    }
};

// DOM elements
const uploadArea = document.getElementById('uploadArea');
const browseButton = document.getElementById('browseButton');
const fileInput = document.getElementById('fileInput');
const fileCountEl = document.getElementById('fileCount');
const summaryFileCount = document.getElementById('summaryFileCount');
const summarySheetCount = document.getElementById('summarySheetCount');
const summaryRowCount = document.getElementById('summaryRowCount');
const summaryColumnCount = document.getElementById('summaryColumnCount');
const processBtn = document.getElementById('processBtn');
const resetBtn = document.getElementById('resetBtn');
const downloadBtn = document.getElementById('downloadBtn');
const progressBar = document.getElementById('progressBar');
const progressFill = document.getElementById('progressFill');
const sheetSelect = document.getElementById('sheetSelect');
const previewTable = document.getElementById('previewTable');
const tableHeader = document.getElementById('tableHeader');
const previewBody = document.getElementById('previewBody');
const previewInfo = document.getElementById('previewInfo');
const previewCount = document.getElementById('previewCount');
const summaryBody = document.getElementById('summaryBody');
const summaryTableFooter = document.querySelector('#summaryTable tfoot');
const totalSheetsEl = document.getElementById('totalSheets');
const totalRowsEl = document.getElementById('totalRows');
const totalColumnsEl = document.getElementById('totalColumns');
const reportModal = document.getElementById('reportModal');
const closeModal = document.getElementById('closeModal');
const closeModalBtn = document.getElementById('closeModalBtn');
const confirmDownload = document.getElementById('confirmDownload');
const modalFileCount = document.getElementById('modalFileCount');
const modalSheetCount = document.getElementById('modalSheetCount');
const modalRowCount = document.getElementById('modalRowCount');
const modalColumnCount = document.getElementById('modalColumnCount');
const modalTime = document.getElementById('modalTime');
const modalFileSize = document.getElementById('modalFileSize');
const infoTabs = document.querySelectorAll('.info-tab');
const tabPanes = document.querySelectorAll('.tab-pane');

// Initialize
function init() {
    setupEventListeners();
    resetTool();
}

// Setup event listeners
function setupEventListeners() {
    browseButton.addEventListener('click', (e) => {
        e.preventDefault();
        e.stopPropagation();
        fileInput.click();
    });

    uploadArea.addEventListener('click', (e) => {
        if (e.target.closest('#browseButton')) return;
        fileInput.click();
    });

    uploadArea.addEventListener('dragover', (e) => {
        e.preventDefault();
        uploadArea.classList.add('drag-over');
    });

    uploadArea.addEventListener('dragleave', () => {
        uploadArea.classList.remove('drag-over');
    });

    uploadArea.addEventListener('drop', (e) => {
        e.preventDefault();
        uploadArea.classList.remove('drag-over');
        if (e.dataTransfer.files.length) {
            handleFiles(e.dataTransfer.files);
        }
    });

    fileInput.addEventListener('change', (e) => {
        if (e.target.files && e.target.files.length > 0) {
            handleFiles(e.target.files);
        }
    });

    processBtn.addEventListener('click', processWorkbooks);
    resetBtn.addEventListener('click', resetTool);
    downloadBtn.addEventListener('click', showDownloadModal);
    closeModal.addEventListener('click', () => reportModal.style.display = 'none');
    closeModalBtn.addEventListener('click', () => reportModal.style.display = 'none');
    confirmDownload.addEventListener('click', downloadCombinedWorkbook);
    sheetSelect.addEventListener('change', updatePreview);

    // Tab functionality
    infoTabs.forEach(tab => {
        tab.addEventListener('click', () => {
            const tabId = tab.getAttribute('data-tab');
            infoTabs.forEach(t => t.classList.remove('active'));
            tab.classList.add('active');
            tabPanes.forEach(pane => {
                pane.classList.remove('active');
                if (pane.id === `${tabId}-tab`) {
                    pane.classList.add('active');
                }
            });
        });
    });
}

// Handle file selection
function handleFiles(files) {
    selectedFiles = Array.from(files);
    updateFileCount();

    processedData = {
        sheets: [],
        combinedWB: null,
        processingStats: { startTime: 0, endTime: 0 }
    };

    processBtn.disabled = false;
    sheetSelect.innerHTML = '<option value="">No sheets available</option>';
    sheetSelect.disabled = true;
    previewInfo.textContent = `${selectedFiles.length} workbook(s) selected. Click "Process Workbooks" to combine.`;

    // Show file list in preview
    let fileListHTML = '';
    selectedFiles.forEach((file, index) => {
        const fileSize = (file.size / 1024).toFixed(1);
        fileListHTML += `
            <div style="display: flex; align-items: center; gap: 0.8rem; padding: 0.5rem; background: white; border-radius: 6px; margin-bottom: 0.5rem;">
                <i class="fas fa-file-excel" style="color: var(--accent-blue);"></i>
                <span style="flex: 1;">${file.name}</span>
                <span style="color: var(--text-light); font-size: 0.9rem;">${fileSize} KB</span>
            </div>
        `;
    });

    previewBody.innerHTML = `
        <tr>
            <td colspan="100" style="padding: 2rem;">
                <h3 style="color: var(--primary-blue); margin-bottom: 1rem; display: flex; align-items: center; gap: 0.8rem;">
                    <i class="fas fa-folder-open"></i> Selected Workbooks (${selectedFiles.length})
                </h3>
                <div style="background: var(--lighter-blue); padding: 1.5rem; border-radius: 8px; max-height: 300px; overflow-y: auto;">
                    ${fileListHTML}
                </div>
                <div style="margin-top: 1.5rem; padding: 1rem; background: #e7f4ff; border-left: 4px solid var(--accent-blue); border-radius: 0 8px 8px 0;">
                    <div style="display: flex; align-items: center; gap: 0.8rem; color: var(--primary-blue); font-weight: 600;">
                        <i class="fas fa-info-circle"></i>
                        <span>Click "Process Workbooks" to combine the selected Excel workbooks</span>
                    </div>
                </div>
            </td>
        </tr>
    `;

    previewCount.textContent = 'Showing 0 rows';

    // Reset summary table
    summaryBody.innerHTML = `
        <tr>
            <td colspan="6" style="text-align: center; padding: 3rem; color: var(--text-light);">
                <i class="fas fa-file-excel" style="font-size: 3rem; margin-bottom: 1rem; display: block;"></i>
                <h3 style="margin-bottom: 0.5rem; color: var(--primary-blue);">No Data Processed</h3>
                <p>Upload Excel workbooks and click "Process Workbooks" to see the summary</p>
            </td>
        </tr>
    `;
    summaryTableFooter.style.display = 'none';

    fileInput.value = '';
}

function updateFileCount() {
    fileCountEl.textContent = selectedFiles.length;
    summaryFileCount.textContent = selectedFiles.length;
}

async function processWorkbooks() {
    if (selectedFiles.length === 0) {
        showNotification('Please select at least one Excel workbook to combine.', 'error');
        return;
    }

    processFilesWithProgress();
}

async function processFilesWithProgress() {
    processedData.processingStats.startTime = Date.now();

    processBtn.innerHTML = '<div class="loading"></div> Processing Workbooks...';
    processBtn.disabled = true;
    downloadBtn.disabled = true;
    progressBar.style.display = 'block';
    progressFill.style.width = '0%';

    processedData = {
        sheets: [],
        combinedWB: XLSX.utils.book_new(),
        processingStats: processedData.processingStats
    };

    let processedCount = 0;
    const totalFiles = selectedFiles.length;
    let totalRows = 0;
    let maxColumns = 0;
    let sheetCounter = 0;

    summaryBody.innerHTML = '';

    for (const file of selectedFiles) {
        try {
            const reader = new FileReader();
            await new Promise((resolve, reject) => {
                reader.onload = function(e) {
                    try {
                        const data = new Uint8Array(e.target.result);
                        const workbook = XLSX.read(data, { type: 'array' });

                        workbook.SheetNames.forEach((originalSheetName) => {
                            const worksheet = workbook.Sheets[originalSheetName];
                            const sheetData = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '' });

                            const rows = sheetData.length;
                            const cols = rows > 0 ? Math.max(...sheetData.map(row => row.length)) : 0;

                            totalRows += rows;
                            maxColumns = Math.max(maxColumns, cols);

                            const sheetKey = `S_${sheetCounter++}`;

                            processedData.sheets.push({
                                key: sheetKey,
                                fileName: file.name,
                                originalSheetName: originalSheetName,
                                finalSheetName: originalSheetName,
                                rows: rows,
                                columns: cols,
                                data: sheetData
                            });

                            XLSX.utils.book_append_sheet(processedData.combinedWB, worksheet, sheetKey);

                            const tr = document.createElement('tr');
                            tr.innerHTML = `
                                <td>${file.name}</td>
                                <td>${originalSheetName}</td>
                                <td><input type="text" class="sheet-name-edit" value="${originalSheetName}" data-key="${sheetKey}"></td>
                                <td>${rows}</td>
                                <td>${cols}</td>
                                <td><span style="color: var(--success-green);"><i class="fas fa-check-circle"></i> Processed</span></td>
                            `;
                            summaryBody.appendChild(tr);
                        });

                        processedCount++;
                        progressFill.style.width = `${(processedCount / totalFiles) * 100}%`;

                        document.querySelectorAll('.sheet-name-edit').forEach(input => {
                            input.addEventListener('input', function() {
                                const key = this.getAttribute('data-key');
                                const newName = this.value || this.getAttribute('value');
                                updateSheetName(key, newName);
                            });
                        });

                        resolve();
                    } catch (error) {
                        reject(error);
                    }
                };
                reader.onerror = reject;
                reader.readAsArrayBuffer(file);
            });
        } catch (error) {
            console.error('Error processing file:', error);
            showNotification(`Error processing workbook "${file.name}": ${error.message}`, 'error');
            processedCount++;
            progressFill.style.width = `${(processedCount / totalFiles) * 100}%`;
        }
    }

    finishProcessing(totalRows, maxColumns);
}

function finishProcessing(totalRows, maxColumns) {
    processedData.processingStats.endTime = Date.now();
    const processingTime = ((processedData.processingStats.endTime - processedData.processingStats.startTime) / 1000).toFixed(2);

    summarySheetCount.textContent = processedData.sheets.length;
    summaryRowCount.textContent = totalRows;
    summaryColumnCount.textContent = maxColumns;

    totalSheetsEl.textContent = processedData.sheets.length;
    totalRowsEl.textContent = totalRows;
    totalColumnsEl.textContent = maxColumns;
    summaryTableFooter.style.display = '';

    sheetSelect.innerHTML = '';
    sheetSelect.disabled = false;

    processedData.sheets.forEach((sheet, index) => {
        const option = document.createElement('option');
        option.value = sheet.key;
        option.textContent = `${sheet.fileName} - ${sheet.finalSheetName} (${sheet.rows} rows)`;
        sheetSelect.appendChild(option);
    });

    if (processedData.sheets.length > 0) {
        updatePreview();
    }

    downloadBtn.disabled = false;
    processBtn.innerHTML = '<i class="fas fa-cogs"></i> Process Workbooks';
    processBtn.disabled = false;
    progressBar.style.display = 'none';

    showNotification(`Successfully processed ${selectedFiles.length} workbook(s) with ${processedData.sheets.length} sheet(s) in ${processingTime}s`, 'success');
}

function updateSheetName(key, newName) {
    const sheet = processedData.sheets.find(s => s.key === key);
    if (sheet) {
        sheet.finalSheetName = newName;
        const option = sheetSelect.querySelector(`option[value="${key}"]`);
        if (option) {
            option.textContent = `${sheet.fileName} - ${newName} (${sheet.rows} rows)`;
        }
        if (sheetSelect.value === key) {
            updatePreview();
        }
    }
}

function updatePreview() {
    const selectedSheetKey = sheetSelect.value;
    if (!selectedSheetKey || processedData.sheets.length === 0) {
        previewBody.innerHTML = `
            <tr>
                <td colspan="100" style="text-align: center; padding: 4rem; color: var(--text-light);">
                    <i class="fas fa-file-excel" style="font-size: 3rem; margin-bottom: 1rem; display: block;"></i>
                    <h3 style="margin-bottom: 0.5rem; color: var(--primary-blue);">No Data to Display</h3>
                    <p>Upload Excel workbooks and click "Process Workbooks" to see the preview here</p>
                </td>
            </tr>
        `;
        previewInfo.textContent = 'No data loaded';
        previewCount.textContent = 'Showing 0 rows';
        return;
    }

    const sheet = processedData.sheets.find(s => s.key === selectedSheetKey);
    if (!sheet) return;

    const sheetData = sheet.data;
    const maxRowsToShow = Math.min(sheetData.length, 50);

    if (sheetData.length > 0) {
        const headers = sheetData[0].map((_, index) => `Column ${index + 1}`);
        tableHeader.innerHTML = '<tr><th>Row</th>' + headers.map(h => `<th>${h}</th>`).join('') + '</tr>';
    }

    previewBody.innerHTML = '';
    for (let r = 0; r < maxRowsToShow; r++) {
        const rowData = sheetData[r] || [];
        let rowHTML = '<tr>';
        rowHTML += `<td style="font-weight: 600;">${r + 1}</td>`;
        for (let c = 0; c < sheet.columns; c++) {
            const value = rowData[c] !== undefined ? rowData[c] : '';
            const cellClass = value === '' || value === 0 ? 'zero-value' : '';
            rowHTML += `<td class="${cellClass}">${value}</td>`;
        }
        rowHTML += '</tr>';
        previewBody.innerHTML += rowHTML;
    }

    if (sheetData.length > maxRowsToShow) {
        previewBody.innerHTML += `
            <tr>
                <td colspan="${sheet.columns + 1}" style="text-align: center; padding: 1rem; background: var(--light-blue); color: var(--primary-blue); font-weight: 600;">
                    <i class="fas fa-info-circle"></i> Showing first ${maxRowsToShow} rows only. Full data will be included in the downloaded workbook.
                </td>
            </tr>
        `;
    }

    previewInfo.textContent = `Sheet: ${sheet.fileName} - ${sheet.finalSheetName}`;
    previewCount.textContent = `Showing ${Math.min(maxRowsToShow, sheetData.length)} of ${sheetData.length} rows`;
}

function showDownloadModal() {
    if (processedData.sheets.length === 0) {
        showNotification('Please process workbooks first before downloading.', 'error');
        return;
    }

    // Populate modal
    modalFileCount.textContent = selectedFiles.length;
    modalSheetCount.textContent = processedData.sheets.length;
    const totalRows = processedData.sheets.reduce((sum, sheet) => sum + sheet.rows, 0);
    const maxColumns = Math.max(...processedData.sheets.map(sheet => sheet.columns));
    modalRowCount.textContent = totalRows;
    modalColumnCount.textContent = maxColumns;
    const processingTime = ((processedData.processingStats.endTime - processedData.processingStats.startTime) / 1000).toFixed(2);
    modalTime.textContent = `${processingTime}s`;
    const estimatedSizeKB = Math.round((totalRows * maxColumns * 15) / 1024);
    modalFileSize.textContent = `${estimatedSizeKB} KB`;

    reportModal.style.display = 'flex';
}

function downloadCombinedWorkbook() {
    reportModal.style.display = 'none';
    performDownload();
}

function performDownload() {
    try {
        // Collect all final sheet names and ensure uniqueness
        const existingNames = [];
        processedData.sheets.forEach(sheet => {
            let uniqueName = sheet.finalSheetName.substring(0, 31).replace(/[\\/*?:[\]]/g, '_');
            let counter = 1;
            while (existingNames.includes(uniqueName)) {
                uniqueName = `${sheet.finalSheetName.substring(0, 28)}_${counter}`.substring(0, 31).replace(/[\\/*?:[\]]/g, '_');
                counter++;
            }
            sheet.uniqueName = uniqueName;
            existingNames.push(uniqueName);
        });

        // Create summary sheet data
        const summaryData = [
            ['Workbook Combiner - Summary Report'],
            ['Generated on:', new Date().toLocaleString()],
            [''],
            ['Workbooks Selected:', selectedFiles.length],
            ['Sheets Combined:', processedData.sheets.length],
            ['Total Rows:', processedData.sheets.reduce((sum, sheet) => sum + sheet.rows, 0)],
            ['Max Columns:', Math.max(...processedData.sheets.map(sheet => sheet.columns))],
            [''],
            ['File Name', 'Original Sheet', 'Final Sheet', 'Rows', 'Columns', 'Status']
        ];

        processedData.sheets.forEach(sheet => {
            summaryData.push([
                sheet.fileName,
                sheet.originalSheetName,
                sheet.finalSheetName,
                sheet.rows,
                sheet.columns,
                'Processed'
            ]);
        });

        // Generate unique name for summary sheet
        let summarySheetName = 'Summary_Report';
        let counter = 1;
        while (existingNames.includes(summarySheetName)) {
            summarySheetName = `Summary_Report_${counter}`;
            counter++;
        }

        // Create workbook
        const downloadWB = XLSX.utils.book_new();

        // Create summary worksheet and set column widths
        const summaryWS = XLSX.utils.aoa_to_sheet(summaryData);
        // Set column widths (approx. based on max content length)
        const colWidths = [];
        for (let i = 0; i < 6; i++) {
            let maxLen = 10;
            summaryData.forEach(row => {
                if (row[i]) {
                    const len = row[i].toString().length;
                    if (len > maxLen) maxLen = len;
                }
            });
            colWidths.push({ wch: maxLen + 2 });
        }
        summaryWS['!cols'] = colWidths;

        // Add summary sheet first
        XLSX.utils.book_append_sheet(downloadWB, summaryWS, summarySheetName);

        // Add other sheets with unique names
        processedData.sheets.forEach(sheet => {
            const worksheet = processedData.combinedWB.Sheets[sheet.key];
            XLSX.utils.book_append_sheet(downloadWB, worksheet, sheet.uniqueName);
        });

        // Generate filename and download
        const date = new Date();
        const dateStr = date.toISOString().slice(0, 10).replace(/-/g, '_');
        const randomNum = Math.floor(1000 + Math.random() * 9000);
        const filename = `combined_workbook_${dateStr}_${randomNum}.xlsx`;

        const wbout = XLSX.write(downloadWB, { bookType: 'xlsx', type: 'array' });
        const blob = new Blob([wbout], { type: 'application/octet-stream' });
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = filename;
        document.body.appendChild(a);
        a.click();
        setTimeout(() => {
            document.body.removeChild(a);
            URL.revokeObjectURL(url);
        }, 100);

        showNotification('Combined workbook downloaded successfully!', 'success');
    } catch (error) {
        console.error('Error downloading file:', error);
        showNotification('Error generating workbook: ' + error.message, 'error');
    }
}

function resetTool() {
    selectedFiles = [];
    processedData = {
        sheets: [],
        combinedWB: null,
        processingStats: { startTime: 0, endTime: 0 }
    };

    updateFileCount();
    summarySheetCount.textContent = '0';
    summaryRowCount.textContent = '0';
    summaryColumnCount.textContent = '0';

    processBtn.disabled = true;
    downloadBtn.disabled = true;
    progressBar.style.display = 'none';

    sheetSelect.innerHTML = '<option value="">No sheets available</option>';
    sheetSelect.disabled = true;

    previewBody.innerHTML = `
        <tr>
            <td colspan="100" style="text-align: center; padding: 4rem; color: var(--text-light);">
                <i class="fas fa-file-excel" style="font-size: 3rem; margin-bottom: 1rem; display: block;"></i>
                <h3 style="margin-bottom: 0.5rem; color: var(--primary-blue);">No Data to Display</h3>
                <p>Upload Excel workbooks and click "Process Workbooks" to see the preview here</p>
            </td>
        </tr>
    `;
    previewInfo.textContent = 'No data loaded. Please upload Excel workbooks to begin.';
    previewCount.textContent = 'Showing 0 rows';

    summaryBody.innerHTML = `
        <tr>
            <td colspan="6" style="text-align: center; padding: 3rem; color: var(--text-light);">
                <i class="fas fa-file-excel" style="font-size: 3rem; margin-bottom: 1rem; display: block;"></i>
                <h3 style="margin-bottom: 0.5rem; color: var(--primary-blue);">No Data Processed</h3>
                <p>Upload Excel workbooks and click "Process Workbooks" to see the summary</p>
            </td>
        </tr>
    `;
    summaryTableFooter.style.display = 'none';

    fileInput.value = '';
    showNotification('Tool has been reset successfully.', 'success');
}

function showNotification(message, type = 'info') {
    const notification = document.createElement('div');
    notification.style.cssText = `
        position: fixed;
        top: 20px;
        right: 20px;
        padding: 1rem 1.5rem;
        border-radius: 8px;
        color: white;
        font-weight: 600;
        z-index: 9999;
        display: flex;
        align-items: center;
        gap: 0.8rem;
        box-shadow: 0 4px 12px rgba(0, 0, 0, 0.15);
        animation: slideIn 0.3s ease;
        max-width: 400px;
    `;

    if (type === 'success') notification.style.background = 'var(--success-green)';
    else if (type === 'error') notification.style.background = 'var(--error-red)';
    else if (type === 'warning') notification.style.background = 'var(--warning-orange)';
    else notification.style.background = 'var(--accent-blue)';

    let icon = 'info-circle';
    if (type === 'success') icon = 'check-circle';
    if (type === 'error') icon = 'exclamation-circle';
    if (type === 'warning') icon = 'exclamation-triangle';

    notification.innerHTML = `<i class="fas fa-${icon}"></i><span>${message}</span>`;
    document.body.appendChild(notification);

    setTimeout(() => {
        notification.style.animation = 'slideOut 0.3s ease';
        setTimeout(() => {
            if (notification.parentNode) notification.parentNode.removeChild(notification);
        }, 300);
    }, 5000);

    if (!document.querySelector('#notification-styles')) {
        const style = document.createElement('style');
        style.id = 'notification-styles';
        style.textContent = `
            @keyframes slideIn { from { transform: translateX(100%); opacity: 0; } to { transform: translateX(0); opacity: 1; } }
            @keyframes slideOut { from { transform: translateX(0); opacity: 1; } to { transform: translateX(100%); opacity: 0; } }
        `;
        document.head.appendChild(style);
    }
}

// Start the application
init();