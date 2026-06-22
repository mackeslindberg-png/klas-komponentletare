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
        .replace(/#/g, "")
        .trim();
}

function innehallerNormaliserat(text, sokvarde) {
    return normalisera(text).includes(normalisera(sokvarde));
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
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json(sheet, { header: 1 });

        komponenter = [];

        for (let i = 1; i < rows.length; i++) {
            const rad = rows[i];

            if (rad[2] && rad[3] && rad[5]) {
                komponenter.push({
                    typ: String(rad[2]),
                    komp: String(rad[3]),
                    art: String(rad[5]),
                    kontrollerad: false,
                    avvikelse: false
                });
            }
        }

        visaLista();
    };

    reader.readAsArrayBuffer(file);
});

function visaLista() {
    let kontrollerade = komponenter.filter(k => k.kontrollerad).length;
    let avvikelser = komponenter.filter(k => k.avvikelse).length;

    let html = "<h3>Excel inläst</h3>";
    html += "<p>Antal komponenter: " + komponenter.length + "</p>";
    html += "<p>Kontrollerade: " + kontrollerade + " / " + komponenter.length + "</p>";
    html += "<p>Avvikelser: " + avvikelser + "</p>";

    html += `
        <label>Sök komponentnummer:</label><br>
        <input type="text" id="sokRuta" placeholder="Ex: 41714">
        <button onclick="sokKomponent()">Sök</button>
        <div id="sokResultat"></div>
        <hr>
    `;

    html += "<ul>";

    komponenter.forEach(k => {
        let klass = "";

        if (k.avvikelse) {
            klass = "avvikelse";
        } else if (k.kontrollerad) {
            klass = "ok";
        }

        html += `
            <li class="${klass}">
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

    if (komponenter.length === 0) {
        ocrStatus.innerHTML = "Läs in Excel-filen först.";
        return;
    }

    try {
        ocrStatus.innerHTML = "Tar bild från kameran...";

        snapshot.width = camera.videoWidth;
        snapshot.height = camera.videoHeight;

        const ctx = snapshot.getContext("2d");
        ctx.drawImage(camera, 0, 0, snapshot.width, snapshot.height);

        ocrStatus.innerHTML = "Bild tagen. Läser text...";

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

        tolkaOcrText(text);

    } catch (error) {
        ocrStatus.innerHTML = "OCR-fel: " + error.message;
    }
});

function tolkaOcrText(text) {
    const hittadKomponent = komponenter.find(k =>
        innehallerNormaliserat(text, k.komp)
    );

    if (!hittadKomponent) {
        ocrStatus.innerHTML = `
            <h3>Ingen komponent hittad</h3>
            <p>OCR läste:</p>
            <pre>${text}</pre>
        `;
        return;
    }

    const artikelStammer = innehallerNormaliserat(text, hittadKomponent.art);

    if (artikelStammer) {
        hittadKomponent.kontrollerad = true;
        hittadKomponent.avvikelse = false;

        ocrStatus.innerHTML = `
            <div class="match">
                <h3>✅ OK</h3>
                <p><strong>${hittadKomponent.typ}</strong></p>
                <p>Komp.nr: ${hittadKomponent.komp}</p>
                <p>Art.nr: ${hittadKomponent.art}</p>
            </div>
        `;
    } else {
        hittadKomponent.kontrollerad = true;
        hittadKomponent.avvikelse = true;

        ocrStatus.innerHTML = `
            <div class="fel">
                <h3>❌ Mismatch</h3>
                <p><strong>${hittadKomponent.typ}</strong></p>
                <p>Komp.nr hittat: ${hittadKomponent.komp}</p>
                <p>Förväntat art.nr: ${hittadKomponent.art}</p>
                <p>OCR läste:</p>
                <pre>${text}</pre>
            </div>
        `;
    }

    visaLista();
}
