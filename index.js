// Constants and Configuration
const CONFIG = {
    EXPORT_SETTINGS: {
        bom: true,
        columnDelimiter: ',',
        columnHeaders: true,
        exportHiddenColumns: true,
        exportHiddenRows: true,
        filename: 'Spreadsheet_[YYYY]-[MM]-[DD]',
        mimeType: 'text/csv;charset=utf-8;',
        rowDelimiter: '\r\n',
        rowHeaders: true
    },
    INITIAL_DATA: Array(20).fill().map(() => Array(12).fill(''))
};

// Initialize HyperFormula
const hyperformulaInstance = HyperFormula.buildEmpty({
    licenseKey: 'gpl-v3'
});

// Help modal functionality
const showHelp = () => {
    const helpModal = new bootstrap.Modal(document.getElementById('helpModal'));
    helpModal.show();
};

// Initialize Handsontable with improved error handling
function initializeHandsontable() {
    const container = document.querySelector('#table');
    if (!container) {
        console.error('Handsontable container not found');
        return null;
    }

    return new Handsontable(container, {
        data: CONFIG.INITIAL_DATA,
        contextMenu: true,
        customBorders: true,
        rowHeaders: true,
        colHeaders: true,
        comments: true,
        mergeCells: true,
        manualRowResize: true,
        manualColumnResize: true,
        colWidths: 100,
        formulas: {
            engine: hyperformulaInstance,
            enabled: true,
            sheetName: 'Sheet1'
        },
        licenseKey: 'non-commercial-and-evaluation',
        afterChange: function(changes) {
            if (!changes) return;
            
            changes.forEach(([row, prop, oldValue, newValue]) => {
                if (newValue && typeof newValue === 'string' && newValue.startsWith('=')) {
                    // Store formula in cell metadata
                    this.setCellMeta(row, prop, 'formula', newValue.substring(1));
                }
            });
        },
        afterLoadData: function() {
            this.render(); // Re-render to show calculated values
        }
    });
}

// Add this helper function to get cell data with formulas
function getCellDataWithFormulas(hot) {
    const data = [];
    const rows = hot.countRows();
    const cols = hot.countCols();

    for (let row = 0; row < rows; row++) {
        const rowData = [];
        for (let col = 0; col < cols; col++) {
            const cellProperties = hot.getCellMeta(row, col);
            // Get formula if it exists, otherwise get cell value
            const formula = cellProperties.formula ? '=' + cellProperties.formula : hot.getDataAtCell(row, col);
            rowData.push(formula);
        }
        data.push(rowData);
    }
    return data;
}

// Replace the handleCSVUpload function with this more comprehensive file handler
function handleFileUpload(file) {
    return new Promise((resolve, reject) => {
        const fileExtension = file.name.split('.').pop().toLowerCase();
        
        if (fileExtension === 'csv') {
            const reader = new FileReader();
            reader.onerror = () => reject(new Error('File reading failed'));
            reader.onload = (e) => {
                try {
                    const text = e.target.result;
                    const data = text.split('\n')
                        .map(row => row.trim())
                        .filter(row => row.length > 0)
                        .map(row => {
                            // Parse CSV while preserving formulas
                            let inQuote = false;
                            let cells = [];
                            let currentCell = '';
                            
                            for (let char of row) {
                                if (char === '"') {
                                    inQuote = !inQuote;
                                } else if (char === ',' && !inQuote) {
                                    cells.push(currentCell.trim());
                                    currentCell = '';
                                } else {
                                    currentCell += char;
                                }
                            }
                            cells.push(currentCell.trim());
                            return cells.map(cell => cell.replace(/(^"|"$)/g, ''));
                        });
                    resolve(data);
                } catch (error) {
                    reject(new Error(`Failed to parse CSV file: ${error.message}`));
                }
            };
            reader.readAsText(file);
        } else if (['xlsx', 'xls'].includes(fileExtension)) {
            const reader = new FileReader();
            reader.onerror = () => reject(new Error('File reading failed'));
            reader.onload = (e) => {
                try {
                    const data = e.target.result;
                    const workbook = XLSX.read(data, { type: 'array' });
                    const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
                    
                    // Get both values and formulas
                    const jsonData = XLSX.utils.sheet_to_json(firstSheet, { 
                        header: 1,
                        raw: false,
                        defval: ''
                    });
                    
                    // Preserve formulas
                    for (let R = 0; R < jsonData.length; ++R) {
                        for (let C = 0; C < jsonData[R].length; ++C) {
                            const cell = firstSheet[XLSX.utils.encode_cell({r: R, c: C})];
                            if (cell && cell.f) {
                                jsonData[R][C] = '=' + cell.f;
                            }
                        }
                    }
                    
                    resolve(jsonData);
                } catch (error) {
                    reject(new Error(`Failed to parse Excel file: ${error.message}`));
                }
            };
            reader.readAsArrayBuffer(file);
        } else {
            reject(new Error('Unsupported file format'));
        }
    });
}

// Initialize table
const hot = initializeHandsontable();
if (!hot) {
    throw new Error('Failed to initialize Handsontable');
}

// Event Listeners
document.querySelector('#export-file')?.addEventListener('click', () => {
    try {
        // Get data with formulas
        const data = getCellDataWithFormulas(hot);
        
        // Export to CSV (formulas will be preserved as strings)
        const csvContent = data.map(row => 
            row.map(cell => {
                // Properly escape and quote cells containing formulas or commas
                if (cell && (cell.toString().includes(',') || cell.toString().startsWith('='))) {
                    return `"${cell.replace(/"/g, '""')}"`;
                }
                return cell || ''; // Convert null/undefined to empty string
            }).join(',')
        ).join('\n');
        
        const blob = new Blob([csvContent], { type: CONFIG.EXPORT_SETTINGS.mimeType });
        
        // Create download link and trigger download
        const fileName = CONFIG.EXPORT_SETTINGS.filename
            .replace('[YYYY]', new Date().getFullYear())
            .replace('[MM]', String(new Date().getMonth() + 1).padStart(2, '0'))
            .replace('[DD]', String(new Date().getDate()).padStart(2, '0'));
            
        const link = document.createElement('a');
        link.href = URL.createObjectURL(blob);
        link.download = `${fileName}.csv`;
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
        
    } catch (error) {
        console.error('Export failed:', error);
        alert('Failed to export file. Please try again.');
    }
});

document.querySelector('#addRowButton')?.addEventListener('click', () => {
    hot.alter('insert_row_below', hot.countRows() - 1);
});

document.querySelector('#addColumnButton')?.addEventListener('click', () => {
    hot.alter('insert_col_end', hot.countCols() - 1);
});

document.querySelector('#helpButton')?.addEventListener('click', showHelp);

// Update the file input event listener
document.querySelector('#csvFileInput')?.addEventListener('change', async (event) => {
    try {
        const file = event.target.files?.[0];
        if (!file) return;
        
        // Show loading state if needed
        const importButton = event.target.parentElement.querySelector('button');
        const originalText = importButton.innerHTML;
        importButton.innerHTML = '<i class="bi bi-hourglass-split me-1"></i>Loading...';
        importButton.disabled = true;
        
        const data = await handleFileUpload(file);
        hot.loadData(data);
        
        // Reset button state
        importButton.innerHTML = originalText;
        importButton.disabled = false;
        
    } catch (error) {
        console.error('File upload failed:', error);
        alert('Failed to load file. Please check the file format.');
        
        // Reset file input
        event.target.value = '';
    }
});

