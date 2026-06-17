const excelFile = document.getElementById("excelFile");
const resultat = document.getElementById("resultat");

excelFile.addEventListener("change", function () {
    const file = excelFile.files[0];

    if (!file) {
        resultat.innerHTML = "Ingen fil vald.";
        return;
    }

    const reader = new FileReader();

    reader.onload = function (event) {
        const data = new Uint8Array(event.target.result);
        const workbook = XLSX.read(data, { type: "array" });

        const firstSheetName = workbook.SheetNames[0];
        const sheet = workbook.Sheets[firstSheetName];

        const rows = XLSX.utils.sheet_to_json(sheet, { header: 1 });

        let html = "<h3>Excel inläst</h3>";
        html += "<p>Antal rader: " + rows.length + "</p>";
        html += "<ul>";

        for (let i = 1; i < rows.length; i++) {
            const rad = rows[i];

            const komponentTyp = rad[2];
            const komponentNummer = rad[3];
            const artikelNummer = rad[5];

            if (komponentTyp && komponentNummer && artikelNummer) {
                html += "<li>";
                html += "<strong>" + komponentTyp + "</strong><br>";
                html += "Komp.nr: " + komponentNummer + "<br>";
                html += "Art.nr: " + artikelNummer;
                html += "</li>";
            }
        }

        html += "</ul>";

        resultat.innerHTML = html;
    };

    reader.readAsArrayBuffer(file);
});
