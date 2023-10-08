

document.addEventListener("DOMContentLoaded", function () {
    const generateBtn = document.getElementById("generate-btn");
    const exportBtn = document.getElementById("export-btn");
    const tableContainer = document.getElementById("table-container");

    generateBtn.addEventListener("click", generateExcel);
    exportBtn.addEventListener("click", exportExcel);

    let data = [];

    function generateExcel() {
        const columns = parseInt(document.getElementById("columns").value);
        const rows = parseInt(document.getElementById("rows").value);

      

        let tableHTML = "<table>";
        for (let i = 1; i <= rows; i++) {
            tableHTML += "<tr>";
            const rowData = [];
            for (let j = 1; j <= columns; j++) {
                const cellValue = "";
                rowData.push(cellValue);
                tableHTML += `<td contenteditable="true" class="cell">${cellValue}</td>`;
            }
            tableHTML += "</tr>";
            data.push(rowData);
        }
        tableHTML += "</table>";

        tableContainer.innerHTML = tableHTML;
    }

    function exportExcel() {
        
        const table = tableContainer.querySelector("table");
        const rows = table.querySelectorAll("tr");

        rows.forEach((row, rowIndex) => {
            const cells = row.querySelectorAll("td");
            cells.forEach((cell, columnIndex) => {
                const cellValue = cell.textContent;
                data[rowIndex][columnIndex] = cellValue;
            });
        });

        const ws = XLSX.utils.aoa_to_sheet(data);
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, "Sheet1");
        XLSX.writeFile(wb, "excel_sheet.xlsx");
    }
});
