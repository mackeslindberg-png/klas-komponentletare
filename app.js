const excelFile = document.getElementById("excelFile");
const resultat = document.getElementById("resultat");
const startCamera = document.getElementById("startCamera");
const readNumber = document.getElementById("readNumber");
const camera = document.getElementById("camera");
const snapshot = document.getElementById("snapshot");
const ocrStatus = document.getElementById("ocrStatus");

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

    const reader = new FileReader();

    reader.onload = function (event) {
        const data = new Uint8Array(event.target.result);
        const workbook = XLSX.read(data, { type: "array" });
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json(sheet, { header: 1 });

        komponenter = [];

        for (let i = 1; i < rows.length; i++) {
            const rad = rows[i];

            if (rad[2] && rad[3] && rad[5]) {
                komponenter.push({
                    typ: String(rad[2]),
                    komp: String(rad[3]),
                    art: String(rad[5])
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
        html += `
            <li>
                <strong>${k.typ}</strong><br>
                Komp.nr: ${k.komp}<br>
                Art.nr: ${k.art}
            </li>
        `;
    });

    html += "</ul>";
    resultat.innerHTML = html;
}

function sokKomponent() {
    const sok = normalisera(document.getElementById("sokRuta").value);
    const sokResultat = document.getElementById("sokResultat");

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

startCamera.addEventListener("click", async function () {
    try {
        ocrStatus.innerHTML = "Startar kamera...";

        const stream = await navigator.mediaDevices.getUserMedia({
            video: { facingMode: "environment" },
            audio: false
        });

        camera.srcObject = stream;
        ocrStatus.innerHTML = "Kamera startad.";
    } catch (error) {
        ocrStatus.innerHTML = "Kamerafel: " + error.message;
    }
});

readNumber.addEventListener("click", async function () {
    if (!camera.srcObject) {
        ocrStatus.innerHTML = "Starta kameran först.";
        return;
    }

    try {
        ocrStatus.innerHTML = "Tar bild från kameran...";

        snapshot.width = camera.videoWidth;
        snapshot.height = camera.videoHeight;

        const ctx = snapshot.getContext("2d");
        ctx.drawImage(camera, 0, 0, snapshot.width, snapshot.height);

        ocrStatus.innerHTML = "Bild tagen. Startar textläsning...";

        const result = await Tesseract.recognize(snapshot, "eng", {
            logger: function (m) {
                if (m.status) {
                    ocrStatus.innerHTML =
                        "OCR: " + m.status +
                        (m.progress ? " " + Math.round(m.progress * 100) + "%" : "");
                }
            }
        });

        const text = result.data.text || "";

        ocrStatus.innerHTML = `
            <h3>OCR-resultat</h3>
            <pre>${text}</pre>
        `;
    } catch (error) {
        ocrStatus.innerHTML = "OCR-fel: " + error.message;
    }
});
