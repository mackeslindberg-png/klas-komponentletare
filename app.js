const excelFile = document.getElementById("excelFile");

excelFile.addEventListener("change", function () {
    const file = excelFile.files[0];

    if (!file) {
        alert("Ingen fil vald.");
        return;
    }

    const reader = new FileReader();

    reader.onload = function (event) {
        const data = new Uint8Array(event.target.result);
        const workbook = XLSX.read(data, { type: "array" });

        const firstSheetName = workbook.SheetNames[0];
        const sheet = workbook.Sheets[firstSheetName];

        const rows = XLSX.utils.sheet_to_json(sheet, { header: 1 });

        console.log(rows);
        alert("Excel inläst. Antal rader: " + rows.length);
    };

    reader.readAsArrayBuffer(file);
});
