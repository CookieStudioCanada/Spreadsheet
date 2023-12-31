// Select the Handsontable container and the export button
const container = document.querySelector('#example');
const button = document.querySelector('#export-file');

// Hyperformula instance creation
const hyperformulaInstance = HyperFormula.buildEmpty({
    licenseKey: 'gpl-v3'
});

// Initialize Handsontable with HyperFormula
const hot = new Handsontable(container, {
    data: [
        ['', '', '', '','', '', '','', '', ''],
        ['', '', '', '','', '', '','', '', ''],
        ['', '', '', '','', '', '','', '', ''],
        ['', '', '', '','', '', '','', '', ''],
        ['', '', '', '','', '', '','', '', ''],
        ['', '', '', '','', '', '','', '', ''],
        ['', '', '', '','', '', '','', '', ''],
        ['', '', '', '','', '', '','', '', ''],
        ['', '', '', '','', '', '','', '', ''],
        ['', '', '', '','', '', '','', '', ''],
      ],
    contextMenu: true,
    customBorders: true,
    rowHeaders: true,
    colHeaders: true,
    comments: true,
    mergeCells: true,
    manualRowResize: true,
    manualColumnResize: true,
    formulas: {
        engine: hyperformulaInstance
    },
    licenseKey: 'non-commercial-and-evaluation',
    // other configuration options...
});

// Get the export plugin
const exportPlugin = hot.getPlugin('exportFile');

// Event listener for the CSV export button
button.addEventListener('click', () => {
  exportPlugin.downloadFile('csv', {
    bom: false,
    columnDelimiter: ',',
    columnHeaders: false,
    exportHiddenColumns: true,
    exportHiddenRows: true,
    fileExtension: 'csv',
    filename: 'Handsontable-CSV-file_[YYYY]-[MM]-[DD]',
    mimeType: 'text/csv',
    rowDelimiter: '\r\n',
    rowHeaders: true
  });
});

// Convert Handsontable data to text for email
function convertDataToText(hotInstance) {
    const data = hotInstance.getData();
    return data.map(row => row.join(', ')).join('\n');
}

// Create a mailto link with Handsontable data
function createMailToLink(hotInstance) {
    const emailBody = convertDataToText(hotInstance);
    const subject = encodeURIComponent("Handsontable Data");
    const body = encodeURIComponent(emailBody);

    return `mailto:?subject=${subject}&body=${body}`;
}

// Open the default email client with the Handsontable data
function openEmailClient(hotInstance) {
    const mailtoLink = createMailToLink(hotInstance);
    window.open(mailtoLink, '_blank');
}

// Event listener for the email data button
document.getElementById('emailDataButton').addEventListener('click', function() {
    openEmailClient(hot);
});

function addRow(hotInstance) {
    hotInstance.alter('insert_row_below', hotInstance.countRows() - 1);
}

function addColumn(hotInstance) {
    hotInstance.alter('insert_col_end', hotInstance.countCols() - 1);
}

// Event listeners for adding rows and columns
document.getElementById('addRowButton').addEventListener('click', function() {
    addRow(hot);
});

document.getElementById('addColumnButton').addEventListener('click', function() {
    addColumn(hot);
});