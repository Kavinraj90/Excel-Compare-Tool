let resultData = [];

function readExcel(file) {
    return new Promise((resolve) => {
        const reader = new FileReader();

        reader.onload = function(e) {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, {type: 'array'});
            const sheet = workbook.Sheets[workbook.SheetNames[0]];
            const json = XLSX.utils.sheet_to_json(sheet, {header:1});
            resolve(json);
        };

        reader.readAsArrayBuffer(file);
    });
}

async function compareFiles() {

    const f1 = document.getElementById('file1').files[0];
    const f2 = document.getElementById('file2').files[0];

    if (!f1 || !f2) {
        alert("Upload both files!");
        return;
    }

    document.getElementById("fileNames").innerText =
        `File1: ${f1.name} | File2: ${f2.name}`;

    const data1 = await readExcel(f1);
    const data2 = await readExcel(f2);

    let maxRows = Math.max(data1.length, data2.length);
    let diffCount = 0;

    resultData = [];

    let html = "<table>";

    for (let i = 0; i < maxRows; i++) {
        html += "<tr>";
        let row = [];

        let row1 = data1[i] || [];
        let row2 = data2[i] || [];

        let maxCols = Math.max(row1.length, row2.length);

        for (let j = 0; j < maxCols; j++) {

            let v1 = row1[j] || "";
            let v2 = row2[j] || "";

            if (v1 !== v2) {
                diffCount++;
                html += `<td class="diff">${v1} | ${v2}</td>`;
                row.push(`${v1} | ${v2}`);
            } else {
                html += `<td>${v1}</td>`;
                row.push(v1);
            }
        }

        resultData.push(row);
        html += "</tr>";
    }

    html += "</table>";

    document.getElementById("output").innerHTML = html;
    document.getElementById("resultText").innerText =
        "Total Differences: " + diffCount;
}

// Filter only differences
function filterDiff() {
    let cells = document.querySelectorAll("td");

    cells.forEach(cell => {
        if (!cell.classList.contains("diff")) {
            cell.style.display = "none";
        }
    });
}

// Download result Excel
function downloadExcel() {
    let ws = XLSX.utils.aoa_to_sheet(resultData);
    let wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Result");

    XLSX.writeFile(wb, "comparison_result.xlsx");
}

// Drag & Drop
const dropArea = document.getElementById("dropArea");

dropArea.addEventListener("dragover", e => {
    e.preventDefault();
});

dropArea.addEventListener("drop", e => {
    e.preventDefault();
    const files = e.dataTransfer.files;

    document.getElementById("file1").files = files;
});