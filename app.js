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
        .replace(/[ÅÄ]/g, "A")
        .replace(/Ö/g, "O")
        .replace(/[^A-Z0-9]/g, "");
}

function normaliseraNummer(varde) {
    return String(varde || "")
        .toUpperCase()
        .replace(/[OQ]/g, "0")
        .replace(/[IL]/g, "1")
        .replace(/S/g, "5")
        .replace(/B/g, "8")
        .replace(/[^A-Z0-9]/g, "");
}

function poangForMatch(ocrText, komponent) {
    const ocr = normaliseraNummer(ocrText);
    const komp = normaliseraNummer(komponent.komp);
    const art = normaliseraNummer(komponent.art);

    let poang = 0;
    let detaljer = [];

    if (komp && ocr.includes(komp)) {
        poang += 100;
        detaljer.push("Komponentnummer hittat");
    }

    if (art && ocr.includes(art)) {
        poang += 100;
        detaljer.push("Artikelnummer hittat");
    }

    // Extra stöd för Volvo/Parker-format:
    // Ex: P55194809#T2616 0705 PG#
    if (komp && ocr.includes("P" + komp)) {
        poang += 30;
        detaljer.push("P-prefix + komponentnummer hittat");
    }

    if (art && ocr.includes("T" + art)) {
        poang += 30;
        detaljer.push("T-prefix + artikelnummer hittat");
    }

    return {
        poang,
        detaljer
    };
}

function hittaBastaMatch(ocrText) {
    let basta = null;

    for (const komponent of komponenter) {
        const match = poangForMatch(ocrText, komponent);

        if (!basta || match.poang > basta.poang) {
            basta = {
                komponent,
                poang: match.poang,
                detaljer: match.detaljer
            };
        }
    }

    return basta;
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
    const kontrollerade = komponenter.filter(k => k.kontrollerad).length;
    const avvikelser = komponenter.filter(k => k.avvikelse).length;

    let html = "<h3>Excel inläst</h3>";
    html += "<p>Antal komponenter: " + komponenter.length + "</p>";
    html += "<p>Kontrollerade: " + kontrollerade + " / " + komponenter.length + "</p>";
    html += "<p>Avvikelser: " + avvikelser + "</p>";

    html += `
        <label>Sök komponentnummer eller artikelnummer:</label><br>
        <input type="text" id="sokRuta" placeholder="Ex: 55194809 eller 2616 0705 PG">
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
    const sok = document.getElementById("sokRuta").value;
    const sokResultat = document.getElementById("sokResultat");

    if (!sok) {
        sokResultat.innerHTML = "<p>Skriv ett nummer.</p>";
        return;
    }

    const hittad = komponenter.find(k =>
        normaliseraNummer(k.komp).includes(normaliseraNummer(sok)) ||
        normaliseraNummer(k.art).includes(normaliseraNummer(sok)) ||
        normaliseraNummer(sok).includes(normaliseraNummer(k.komp)) ||
        normaliseraNummer(sok).includes(normaliseraNummer(k.art))
    );

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
    const match = hittaBastaMatch(text);

    if (!match || match.poang < 80) {
        ocrStatus.innerHTML = `
            <div class="fel">
                <h3>Ingen säker träff</h3>
                <p>OCR läste:</p>
                <pre>${text}</pre>
            </div>
        `;
        return;
    }

    const k = match.komponent;

    const kompHittad = normaliseraNummer(text).includes(normaliseraNummer(k.komp));
    const artHittad = normaliseraNummer(text).includes(normaliseraNummer(k.art));

    k.kontrollerad = true;
    k.avvikelse = !artHittad;

    if (artHittad) {
        ocrStatus.innerHTML = `
            <div class="match">
                <h3>✅ OK</h3>
                <p><strong>${k.typ}</strong></p>
                <p>Komp.nr: ${k.komp}</p>
                <p>Art.nr: ${k.art}</p>
                <p>${match.detaljer.join(", ")}</p>
            </div>
        `;
    } else {
        ocrStatus.innerHTML = `
            <div class="fel">
                <h3>⚠️ Komponent hittad, artikelnummer ej bekräftat</h3>
                <p><strong>${k.typ}</strong></p>
                <p>Komp.nr: ${k.komp}</p>
                <p>Förväntat art.nr: ${k.art}</p>
                <p>OCR läste:</p>
                <pre>${text}</pre>
            </div>
        `;
    }

    visaLista();
}
