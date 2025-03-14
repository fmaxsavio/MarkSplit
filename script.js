const marks = [4, 4, 4, 4, 4, 26, 26, 28];
let processedWorkbook;

function processExcel() {
    const fileInput = document.getElementById("uploadFile").files[0];
    if (!fileInput) {
        alert("Please upload an Excel file.");
        return;
    }

    const reader = new FileReader();
    reader.readAsArrayBuffer(fileInput);
    reader.onload = function (e) {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: "array" });
        const sheet = workbook.Sheets[workbook.SheetNames[0]];

        // Identify last row with data in column D
        let lastRow = findLastRow(sheet, "D");

        // Process each value in column D from D3 onwards
        for (let row = 3; row <= lastRow; row++) {
            let cellRef = "D" + row;
            let cell = sheet[cellRef];

            if (!cell || isNaN(cell.v)) continue; // Skip invalid or empty cells

            const inputMarks = parseInt(cell.v);
            const splitMarks = splitMarksFunction(inputMarks);
            const columns = ["E", "F", "G", "H", "I", "J", "K", "L"];

            splitMarks.forEach((val, index) => {
                const newCellRef = columns[index] + row;
                sheet[newCellRef] = { t: "n", v: val }; // Ensure numeric data type
            });
        }

        // Explicitly update the range to reflect new cells
        sheet["!ref"] = `A1:K${lastRow}`;

        // Save processed workbook
        processedWorkbook = workbook;

        // Generate and save the file
        saveWorkbookToFile(workbook, "Output.xlsx");

        document.getElementById("downloadBtn").style.display = "inline";
        document.getElementById("status").innerText = "Processing complete! Click Download.";
    };
}

function downloadExcel() {
    if (!processedWorkbook) return;
    saveWorkbookToFile(processedWorkbook, "Output.xlsx");
}

// Function to split marks correctly
function splitMarksFunction(input) {
    let remaining = input;
    let splitValues = [0, 0, 0, 0, 0, 0, 0, 0]; // Corresponds to E-L

    // Assign marks to columns E-I (Max 4 each)
    for (let i = 0; i < 5; i++) {
        if (remaining >= 4) {
            splitValues[i] = 4;
            remaining -= 4;
        } else {
            splitValues[i] = remaining;
            remaining = 0;
        }
    }

    // Assign marks to columns J-K (Max 26 each)
    for (let i = 5; i < 7; i++) {
        if (remaining >= 26) {
            splitValues[i] = 26;
            remaining -= 26;
        } else {
            splitValues[i] = remaining;
            remaining = 0;
        }
    }

    // Assign remaining marks to column L (Max 28)
    splitValues[7] = remaining; // Whatever is left goes to K

    return splitValues;
}

// Function to determine the last row in column D
function findLastRow(sheet, column) {
    let lastRow = 2; // Start at D3 (index 2)
    while (true) {
        let cellRef = column + (lastRow + 1);
        if (!sheet[cellRef] || isNaN(sheet[cellRef].v)) break;
        lastRow++;
    }
    return lastRow;
}

// Function to save the workbook to a file
function saveWorkbookToFile(workbook, filename) {
    const wbout = XLSX.write(workbook, { bookType: "xlsx", type: "array" });
    const blob = new Blob([wbout], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
    
    const link = document.createElement("a");
    link.href = URL.createObjectURL(blob);
    link.download = filename;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
}
