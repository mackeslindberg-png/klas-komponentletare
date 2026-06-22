const excelFile = document.getElementById("excelFile");
const resultat = document.getElementById("resultat");

let komponenter = [];

function normalisera(varde) {
    return String(varde || "")
        .toUpperCase()
        .replace(/\s/g, "")
        .replace(/-/g, "")
        .trim();
}

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

        komponenter = [];

        for (let i = 1; i < rows.length; i++) {
            const rad = rows[i];

            const komponentTyp = rad[2];
            const komponentNummer = rad[3];
            const artikelNummer = rad[5];

            if (komponentTyp && komponentNummer && artikelNummer) {
                komponenter.push({
                    typ: komponentTyp,
                    komp: String(komponentNummer),
                    art: String(artikelNummer)
                });
            }
        }

        visaLista();
    };

    reader.readAsArrayBuffer(file);
});

function visaLista() {
    let html = "<h3>Excel inläst</h3>";
    html += "<p>Antal komponenter: " + komponenter.length + "</p>";

    html += `
        <label>Sök komponentnummer:</label><br>
        <input type="text" id="sokRuta" placeholder="Ex: 22761">
        <button onclick="sokKomponent()">Sök</button>
        <div id="sokResultat"></div>
        <hr>
    `;

    html += "<ul>";

    komponenter.forEach(k => {
        html += "<li>";
        html += "<strong>" + k.typ + "</strong><br>";
        html += "Komp.nr: " + k.komp + "<br>";
        html += "Art.nr: " + k.art;
        html += "</li>";
    });

    html += "</ul>";

    resultat.innerHTML = html;
}

function sokKomponent() {
    const sokRuta = document.getElementById("sokRuta");
    const sokResultat = document.getElementById("sokResultat");

    const sok = normalisera(sokRuta.value);

    if (!sok) {
        sokResultat.innerHTML = "<p>Skriv ett komponentnummer.</p>";
        return;
    }

    const hittad = komponenter.find(k => normalisera(k.komp) === sok);

    if (!hittad) {
        sokResultat.innerHTML = "<p><strong>Ingen träff.</strong></p>";
        return;
    }

    sokResultat.innerHTML = `
        <div class="match">
            <h3>Träff</h3>
            <p><strong>${hittad.typ}</strong></p>
            <p>Komp.nr: ${hittad.komp}</p>
            <p>Art.nr: ${hittad.art}</p>
        </div>
    `;
}
