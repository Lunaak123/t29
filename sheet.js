let data = [];
let filteredData = [];
let workbook; // Global variable to store workbook

// Function to load the Excel file and display the first sheet initially
async function loadExcelSheet(fileUrl) {
    try {
        const response = await fetch(fileUrl);
        const arrayBuffer = await response.arrayBuffer();
        workbook = XLSX.read(new Uint8Array(arrayBuffer), { type: 'array' });

        populateSheetList(); // Populate sheet names into dropdown
        loadSheetByName(workbook.SheetNames[0]); // Load first sheet by default

    } catch (error) {
        console.error("Error loading Excel sheet:", error);
    }
}

// Populate the sub-sheet list in the dropdown
function populateSheetList() {
    const sheetSelection = document.getElementById('sheet-selection');
    sheetSelection.innerHTML = ''; // Clear existing options

    workbook.SheetNames.forEach((sheetName, index) => {
        const option = document.createElement('option');
        option.value = sheetName;
        option.textContent = sheetName;
        sheetSelection.appendChild(option);
    });

    // Event listener to load selected sheet
    sheetSelection.addEventListener('change', (event) => {
        loadSheetByName(event.target.value);
    });
}

// Load the data of a specific sheet by name
function loadSheetByName(sheetName) {
    const sheet = workbook.Sheets[sheetName];
    data = XLSX.utils.sheet_to_json(sheet, { defval: null });
    filteredData = [...data];
    displaySheet(filteredData);
}

// Function to display the Excel sheet as an HTML table
function displaySheet(sheetData) {
    const sheetContentDiv = document.getElementById('sheet-content');
    sheetContentDiv.innerHTML = '';

    if (sheetData.length === 0) {
        sheetContentDiv.innerHTML = '<p>No data available</p>';
        return;
    }

    const table = document.createElement('table');
    const headerRow = document.createElement('tr');
    Object.keys(sheetData[0]).forEach(header => {
        const th = document.createElement('th');
        th.textContent = header;
        headerRow.appendChild(th);
    });
    table.appendChild(headerRow);

    sheetData.forEach(row => {
        const tr = document.createElement('tr');
        Object.values(row).forEach(cell => {
            const td = document.createElement('td');
            td.textContent = cell === null || cell === "" ? 'NULL' : cell;
            tr.appendChild(td);
        });
        table.appendChild(tr);
    });

    sheetContentDiv.appendChild(table);
}

// Function to apply the selected operations
function applyOperation() {
    const primaryColumn = document.getElementById('primary-column').value.trim();
    const operationColumnsInput = document.getElementById('operation-columns').value.trim();
    const operationType = document.getElementById('operation-type').value;
    const operation = document.getElementById('operation').value;

    if (!primaryColumn || !operationColumnsInput) {
        alert('Please enter the primary column and columns to operate on.');
        return;
    }

    const operationColumns = operationColumnsInput.split(',').map(col => col.trim());

    filteredData = data.filter(row => {
        const isPrimaryNull = row[primaryColumn] === null || row[primaryColumn] === "";
        const columnChecks = operationColumns.map(col => {
            if (operation === 'null') return row[col] === null || row[col] === "";
            else return row[col] !== null && row[col] !== "";
        });

        if (operationType === 'and') return !isPrimaryNull && columnChecks.every(check => check);
        else return !isPrimaryNull && columnChecks.some(check => check);
    });

    filteredData = filteredData.map(row => {
        const filteredRow = {};
        filteredRow[primaryColumn] = row[primaryColumn];
        operationColumns.forEach(col => {
            filteredRow[col] = row[col] === null || row[col] === "" ? 'NULL' : row[col];
        });
        return filteredRow;
    });

    displaySheet(filteredData);
}

// Function to open the download modal
function openDownloadModal() {
    document.getElementById('download-modal').style.display = 'flex';
}

function closeDownloadModal() {
    document.getElementById('download-modal').style.display = 'none';
}

function downloadExcel() {
    const filename = document.getElementById('filename').value.trim() || 'download';
    const format = document.getElementById('file-format').value;
    const exportData = filteredData.map(row => {
        return Object.keys(row).reduce((acc, key) => {
            acc[key] = row[key] === null || row[key] === "" ? 'NULL' : row[key];
            return acc;
        }, {});
    });

    let worksheet = XLSX.utils.json_to_sheet(exportData);
    let workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Filtered Data');

    if (format === 'xlsx') XLSX.writeFile(workbook, `${filename}.xlsx`);
    else if (format === 'csv') XLSX.writeFile(workbook, `${filename}.csv`);

    closeDownloadModal();
}

document.getElementById('apply-operation').addEventListener('click', applyOperation);
document.getElementById('download-button').addEventListener('click', openDownloadModal);
document.getElementById('confirm-download').addEventListener('click', downloadExcel);
document.getElementById('close-modal').addEventListener('click', closeDownloadModal);

window.addEventListener('load', () => {
    const fileUrl = getQueryParam('fileUrl');
    loadExcelSheet(fileUrl);
});

function getQueryParam(param) {
    const urlParams = new URLSearchParams(window.location.search);
    return urlParams.get(param);
}
