// Constants and Configuration
const CONFIG = {
    EXPORT_SETTINGS: {
        bom: true,
        columnDelimiter: ',',
        columnHeaders: true,
        exportHiddenColumns: true,
        exportHiddenRows: true,
        filename: 'Table_[YYYY]-[MM]-[DD]',
        mimeType: 'text/csv;charset=utf-8;',
        rowDelimiter: '\r\n',
        rowHeaders: true
    },
    INITIAL_DATA: (() => {
        // Calculate rows and columns based on screen size
        const screenHeight = window.innerHeight;
        const screenWidth = window.innerWidth;
        
        const rows = Math.min(36, Math.max(20, Math.floor((screenHeight - 100) / 20)));
        const cols = Math.min(18, Math.max(10, Math.floor((screenWidth - 100) / 75)));
        
        return Array(rows).fill().map(() => Array(cols).fill(''));
    })()
};

// Add this after the CONFIG object
const FORMULA_SUGGESTIONS = [
    { 
        name: 'SUM', 
        description: 'Add up values', 
        example: '=SUM(A1:B9)',
        syntax: 'Select a range of cells to sum'
    },
    { 
        name: 'AVERAGE', 
        description: 'Calculate average', 
        example: '=AVERAGE(A1:B9)',
        syntax: 'Select a range of cells to average'
    },
    { 
        name: 'MAX', 
        description: 'Find maximum value', 
        example: '=MAX(A1:B9)',
        syntax: 'Select a range of cells to find maximum'
    },
    { 
        name: 'MIN', 
        description: 'Find minimum value', 
        example: '=MIN(A1:B9)',
        syntax: 'Select a range of cells to find minimum'
    },
    { 
        name: 'IF', 
        description: 'Conditional logic', 
        example: '=IF(A1>10, "High", "Low")',
        syntax: 'Compare a cell with a value and return result'
    },
    { 
        name: 'POWER', 
        description: 'Raise number to power', 
        example: '=POWER(A1, 2)',
        syntax: 'First value is base, second is exponent'
    },
    { 
        name: 'SQRT', 
        description: 'Calculate square root', 
        example: '=SQRT(A1)',
        syntax: 'Select a cell to calculate its square root'
    },
    { 
        name: 'ROUND', 
        description: 'Round to nearest integer', 
        example: '=ROUND(A1)',
        syntax: 'Select a cell to round its value'
    }
];

// Initialize HyperFormula
const hyperformulaInstance = HyperFormula.buildEmpty({
    licenseKey: 'gpl-v3'
});

// Help modal functionality
const showHelp = () => {
    const helpModal = new bootstrap.Modal(document.getElementById('helpModal'));
    helpModal.show();
};

// Add these functions after the FORMULA_SUGGESTIONS constant
function getCellLabel(row, col) {
    const colLabel = String.fromCharCode(65 + col); // A, B, C, etc.
    return `${colLabel}${row + 1}`;
}

function updateFormulaBar(hot, row, col) {
    const formulaInput = document.getElementById('formulaInput');
    
    if (row === undefined || col === undefined) {
        formulaInput.value = '';
        formulaInput.classList.remove('editing-formula');
        return;
    }

    const cellMeta = hot.getCellMeta(row, col);
    const cellValue = hot.getDataAtCell(row, col);
    
    if (cellMeta.formula) {
        formulaInput.value = '=' + cellMeta.formula;
        formulaInput.classList.add('editing-formula');
    } else {
        formulaInput.value = cellValue || '';
        formulaInput.classList.remove('editing-formula');
    }
}

// Modify the createFormulaSuggestions function
function createFormulaSuggestions() {
    const dropdown = document.createElement('div');
    dropdown.className = 'formula-suggestions';
    document.body.appendChild(dropdown);

    return {
        element: dropdown,
        show(rect, filter = '') {
            const filteredSuggestions = FORMULA_SUGGESTIONS.filter(s => 
                s.name.toLowerCase().includes(filter.toLowerCase())
            );

            dropdown.innerHTML = filteredSuggestions.map(suggestion => `
                <div class="suggestion-item">
                    <div class="suggestion-name">${suggestion.name}</div>
                    <div class="suggestion-description">${suggestion.description}</div>
                    <div class="suggestion-syntax">${suggestion.syntax}</div>
                    <div class="suggestion-example">${suggestion.example}</div>
                </div>
            `).join('');

            dropdown.style.display = 'block';
            dropdown.style.left = `${rect.left}px`;
            dropdown.style.top = `${rect.bottom}px`;
        },
        hide() {
            dropdown.style.display = 'none';
        },
        isVisible() {
            return dropdown.style.display === 'block';
        }
    };
}

// Modify the handleFormulaInput function to better handle formula updates
function handleFormulaInput(instance, formulaInput, suggestions) {
    const selected = instance.getSelected();
    if (!selected) return;

    const [row, col] = selected[0];
    const value = formulaInput.value;

    instance.setDataAtCell(row, col, value, 'temp');

    if (value === '=') {
        const rect = formulaInput.getBoundingClientRect();
        suggestions.show(rect);
    } else if (value.startsWith('=')) {
        const filter = value.substring(1);
        if (filter) {
            const rect = formulaInput.getBoundingClientRect();
            suggestions.show(rect, filter);
        }
    } else {
        suggestions.hide();
    }
}

// Modify the initializeHandsontable function
function initializeHandsontable() {
    const container = document.querySelector('#table');
    if (!container) return null;

    const suggestions = createFormulaSuggestions();
    const formulaInput = document.getElementById('formulaInput');
    let currentCell = null;

    document.addEventListener('click', (event) => {
        if (!suggestions.element.contains(event.target) && 
            event.target !== formulaInput) {
            suggestions.hide();
        }
    });

    formulaInput.addEventListener('keydown', (e) => {
        if (e.key === 'Enter' || e.key === 'Escape') {
            e.preventDefault();
            suggestions.hide();
            
            if (e.key === 'Enter' && currentCell) {
                const value = e.target.value;
                instance.setDataAtCell(currentCell.row, currentCell.col, value);
                instance.render();
                
                const nextRow = currentCell.row + 1;
                instance.selectCell(nextRow, currentCell.col);
                currentCell = { row: nextRow, col: currentCell.col };
            }
        }
    });

    const instance = new Handsontable(container, {
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
        afterChange: function(changes, source) {
            if (!changes) return;
            
            changes.forEach(([row, prop, oldValue, newValue]) => {
                if (currentCell && currentCell.row === row && currentCell.col === prop) {
                    const cellMeta = this.getCellMeta(row, prop);
                    if (cellMeta.formula) {
                        formulaInput.value = '=' + cellMeta.formula;
                    } else {
                        formulaInput.value = newValue || '';
                    }
                }
                
                if (source !== 'temp' && newValue && typeof newValue === 'string' && newValue.startsWith('=')) {
                    try {
                        const formula = newValue.substring(1);
                        this.setCellMeta(row, prop, 'formula', formula);
                    } catch (error) {
                        this.setDataAtCell(row, prop, oldValue, 'silent');
                    }
                } else if (source !== 'temp') {
                    this.setCellMeta(row, prop, 'formula', undefined);
                }
            });
            
            if (source !== 'temp') {
                this.render();
            }
        },
        afterLoadData: function() {
            this.render(); // Re-render to show calculated values
        },
        beforeKeyDown: function(event) {
            if (event.key === 'Enter' || event.key === 'Escape') {
                suggestions.hide();
            }
        },
        afterBeginEditing: function(row, col) {
            if (currentCell && currentCell.row === row && currentCell.col === col) {
                const cellMeta = this.getCellMeta(row, col);
                if (cellMeta.formula) {
                    formulaInput.value = '=' + cellMeta.formula;
                } else {
                    const value = this.getDataAtCell(row, col);
                    formulaInput.value = value || '';
                }
            }
        },
        afterSelection: function(row, col) {
            currentCell = { row, col };
            updateFormulaBar(this, row, col);
            suggestions.hide();
        },
        afterDeselect: function() {
            suggestions.hide();
        },
        afterEndEditing: function(row, col, prop, value) {
            if (currentCell && currentCell.row === row && currentCell.col === col) {
                formulaInput.value = value || '';
            }
        }
    });

    // Modify formula input handler
    formulaInput.addEventListener('input', (e) => {
        if (!currentCell) {
            return;
        }

        const value = e.target.value;
        
        // Add editing-formula class when editing a formula
        if (value.startsWith('=')) {
            formulaInput.classList.add('editing-formula');
            // Store the formula in cell meta immediately
            instance.setCellMeta(currentCell.row, currentCell.col, 'formula', value.substring(1));

            // Show suggestions based on the formula being typed
            const rect = formulaInput.getBoundingClientRect();
            if (value === '=') {
                // Show all suggestions when only = is typed
                suggestions.show(rect);
            } else {
                // Extract function name for filtering (everything between = and first parenthesis or end)
                const match = value.match(/^=([A-Za-z]*)/);
                if (match) {
                    const functionName = match[1];
                    // Only show suggestions if we're still typing the function name
                    if (!value.includes(')')) {
                        suggestions.show(rect, functionName);
                    } else {
                        suggestions.hide(); // Hide when formula is complete
                    }
                } else {
                    suggestions.hide();
                }
            }
        } else {
            formulaInput.classList.remove('editing-formula');
            instance.setCellMeta(currentCell.row, currentCell.col, 'formula', undefined);
            suggestions.hide();
        }

        // Update cell with the current value
        instance.setDataAtCell(currentCell.row, currentCell.col, value, 'temp');
    });

    // Modify blur handler
    formulaInput.addEventListener('blur', (e) => {
        if (!currentCell) {
            return;
        }

        const value = e.target.value;
        
        // Keep editing-formula class if it's still a formula
        if (!value.startsWith('=')) {
            formulaInput.classList.remove('editing-formula');
        }
        
        suggestions.hide(); // Hide suggestions when focus is lost
        
        // Update cell with final value
        instance.setDataAtCell(currentCell.row, currentCell.col, value);
        instance.render();
    });

    // Modify suggestions click handler
    suggestions.element.addEventListener('click', (event) => {
        const suggestionItem = event.target.closest('.suggestion-item');
        if (suggestionItem && currentCell) {
            const formula = suggestionItem.dataset.formula;
            const formulaTemplate = `=${formula}(...)`; // Changed to use ... instead of open parenthesis
            
            instance.setDataAtCell(currentCell.row, currentCell.col, formulaTemplate, 'temp');
            formulaInput.value = formulaTemplate;
            formulaInput.focus();
            
            // Place cursor between parentheses
            const cursorPos = formulaTemplate.length - 2;
            formulaInput.setSelectionRange(cursorPos, cursorPos);
            
            suggestions.hide();
        }
    });

    // Add click handler for the table to update formula bar
    container.addEventListener('mousedown', () => {
        // Don't clear currentCell here, it will be updated in afterSelection
    });

    return instance;
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
        alert('Failed to load file. Please check the file format.');
        
        // Reset file input
        event.target.value = '';
    }
});

