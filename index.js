// Constants and Configuration
const CONFIG = {
    CSV_EXPORT: {
        bom: false,
        columnDelimiter: ',',
        columnHeaders: false,
        exportHiddenColumns: true,
        exportHiddenRows: true,
        fileExtension: 'csv',
        filename: 'CSV-file_[YYYY]-[MM]-[DD]',
        mimeType: 'text/csv',
        rowDelimiter: '\r\n',
        rowHeaders: true
    },
    INITIAL_DATA: Array(16).fill().map(() => Array(10).fill(''))
};

// Initialize HyperFormula
const hyperformulaInstance = HyperFormula.buildEmpty({
    licenseKey: 'gpl-v3'
});

// Help modal content
const showHelp = () => {
    alert(`Handsontable Functions Help:

1. Basic Operations:
   - Click any cell to edit
   - Use arrow keys to navigate
   - Press Enter to confirm, Esc to cancel
   
2. Formulas (start with =):
   Basic Math:
   - Addition: =A1+B1
   - Subtraction: =A1-B1
   - Multiplication: =A1*B1
   - Division: =A1/B1
   - Power: =POWER(A1,2) or =A1^2
   - Square Root: =SQRT(A1)
   
   Statistical:
   - Sum: =SUM(A1:A5)
   - Average: =AVERAGE(A1:A5)
   - Count: =COUNT(A1:A5)
   - Maximum: =MAX(A1:A5)
   - Minimum: =MIN(A1:A5)
   - Median: =MEDIAN(A1:A5)
   
   Logical:
   - IF: =IF(A1>10,"High","Low")
   - AND: =AND(A1>10,B1<20)
   - OR: =OR(A1>10,B1<20)
   - NOT: =NOT(A1>10)
   
   Text:
   - Concatenate: =CONCATENATE(A1," ",B1)
   - Left: =LEFT(A1,3)
   - Right: =RIGHT(A1,3)
   - Length: =LEN(A1)
   
   Date:
   - Today: =TODAY()
   - Now: =NOW()
   - Year: =YEAR(A1)
   - Month: =MONTH(A1)
   - Day: =DAY(A1)
   
   Lookup:
   - VLOOKUP: =VLOOKUP(lookup_value,range,col_index)
   - HLOOKUP: =HLOOKUP(lookup_value,range,row_index)
   
3. Features:
   - Right-click for context menu
   - Drag column/row borders to resize
   - Double-click borders for auto-resize
   
4. Keyboard Shortcuts:
   - Ctrl+C: Copy
   - Ctrl+V: Paste
   - Ctrl+Z: Undo
   - Ctrl+Y: Redo`);
};

// Initialize Handsontable with improved error handling
function initializeHandsontable() {
    const container = document.querySelector('#example');
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
            engine: hyperformulaInstance
        },
        licenseKey: 'non-commercial-and-evaluation',
        afterChange: (changes) => {
            if (changes) {
                console.log('Data changed:', changes);
            }
        },
        afterError: (error) => {
            console.error('Handsontable error:', error);
        }
    });
}

// CSV file handling
function handleCSVUpload(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        
        reader.onerror = () => reject(new Error('File reading failed'));
        
        reader.onload = (e) => {
            try {
                const text = e.target.result;
                const data = text.split('\n').map(row => row.split(','));
                resolve(data);
            } catch (error) {
                reject(new Error(`Failed to parse CSV file: ${error.message}`));
            }
        };

        reader.readAsText(file);
    });
}

// Initialize table
const hot = initializeHandsontable();
if (!hot) {
    throw new Error('Failed to initialize Handsontable');
}

// Event Listeners
document.querySelector('#export-file')?.addEventListener('click', () => {
    const exportPlugin = hot.getPlugin('exportFile');
    exportPlugin.downloadFile('csv', CONFIG.CSV_EXPORT);
});

document.querySelector('#addRowButton')?.addEventListener('click', () => {
    hot.alter('insert_row_below', hot.countRows() - 1);
});

document.querySelector('#addColumnButton')?.addEventListener('click', () => {
    hot.alter('insert_col_end', hot.countCols() - 1);
});

document.querySelector('#helpButton')?.addEventListener('click', showHelp);

// CSV file input handler
document.querySelector('#csvFileInput')?.addEventListener('change', async (event) => {
    try {
        const file = event.target.files?.[0];
        if (!file) return;
        
        const data = await handleCSVUpload(file);
        hot.loadData(data);
    } catch (error) {
        console.error('CSV upload failed:', error);
        alert('Failed to load CSV file. Please check the file format.');
    }
});

