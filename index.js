// Select the Handsontable container and the export button
const container = document.querySelector('#example');
const button = document.querySelector('#export-file');

// Hyperformula instance creation
const hyperformulaInstance = HyperFormula.buildEmpty({
    licenseKey: 'gpl-v3'
});

const data = [
    // Profits Section
    ['Revenue', 'Expenses', 'Profit', 'Taxes (%)', 'Taxes dues', ''],
    [0, 0, '=A2-B2', 0.135, '=C2*D2', ''],

    // Space
    ['','','','','','',''],

    // Capital Gains Section
    ['Cost', 'Proceeds', 'Capital Gain', 'Tax on Gain (50.17%)', 'CDA', 'RDTOH'],
    [0, 0, '=(B5-A5)/2', '=(C5*50.17)/100', '=(B5-A5)/2', 0],

    // Space
    ['','','','','','',''],

  ];
  

// Initialize Handsontable with HyperFormula
const hot = new Handsontable(container, {
    data: data,
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
    filename: 'CSV-file_[YYYY]-[MM]-[DD]',
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

// Load data CSV

document.getElementById('csvFileInput').addEventListener('change', function(event) {
    const file = event.target.files[0];
    const reader = new FileReader();
    
    reader.onload = function(e) {
        const text = e.target.result;
        const data = text.split('\n').map(row => row.split(','));
        hot.loadData(data); // Assuming 'hot' is your Handsontable instance
    };

    reader.readAsText(file);
});

// Load data - Excel

document.getElementById('excelFileInput').addEventListener('change', function(event) {
    const file = event.target.files[0];
    const reader = new FileReader();

    reader.onload = function(e) {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        const json = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
        hot.loadData(json); // Load data into Handsontable
    };

    reader.readAsArrayBuffer(file);
});

// Function to open the modal
function openModal() {
    var modal = document.getElementById("editModal");
    var span = document.getElementsByClassName("close")[0];

    // When the user clicks on <span> (x), close the modal
    span.onclick = function() {
        modal.style.display = "none";
    }

    // When the user clicks anywhere outside of the modal, close it
    window.onclick = function(event) {
        if (event.target == modal) {
            modal.style.display = "none";
        }
    }

    // Get the data from Handsontable and display it in the textarea
    var data = hot.getData(); // Assume hot is your Handsontable instance
    document.getElementById("dataTextArea").value = JSON.stringify(data, null, 2);

    // Display the modal
    modal.style.display = "block";
}

// Function to save changes from the modal
document.getElementById("saveButton").addEventListener('click', function() {
    var updatedData = JSON.parse(document.getElementById("dataTextArea").value);
    hot.loadData(updatedData); // Update the Handsontable with new data
    document.getElementById("editModal").style.display = "none"; // Close the modal
});

// Add a button or a way to open the modal
document.getElementById("yourButtonId").addEventListener('click', function() {
    openModal();
});