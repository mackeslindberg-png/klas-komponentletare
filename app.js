const excelFile = document.getElementById("excelFile");

excelFile.addEventListener("change", function () {
    const file = excelFile.files[0];

    if (!file) {
        alert("Ingen fil vald.");
        return;
    }

    alert("Excel-fil vald: " + file.name);
});
