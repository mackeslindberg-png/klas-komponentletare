const excelFile = document.getElementById("excelFile");
const resultat = document.getElementById("resultat");
const startCamera = document.getElementById("startCamera");
const readNumber = document.getElementById("readNumber");
const camera = document.getElementById("camera");
const snapshot = document.getElementById("snapshot");
const ocrStatus = document.getElementById("ocrStatus");
const rotationSlider = document.getElementById("rotationSlider");
const rotationValue = document.getElementById("rotationValue");

let komponenter = [];
let rotation = 0;

rotationSlider.addEventListener("input", function () {
    rotation = Number(rotationSlider.value);
    rotationValue.textContent = rotation;
});

function setRotation(value) {
    rotation = value;
    rotationSlider.value = value;
    rotationValue.textContent = value;
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

function skapaKandidater(text) {
    const ren = normaliseraNummer(text);
    const delar = ren.match(/[A-Z0-9]{4,}/g) || [];
    const kandidater = new Set();

    delar.forEach(d => {
        kandidater.add(d);

        const baraSiffror = d.replace(/[^0-9]/g, "");
        if (baraSiffror.length >= 4) kandidater.add(baraSiffror);

        if (d.length > 6) {
            kandidater.add(d.slice(-6));
            kandidater.add(d.slice(-7));
            kandidater.add(d.slice(-8));
        }

        if (baraSiffror.length > 6) {
            kandidater.add(baraSiffror.slice(-6));
            kandidater.add(baraSiffror.slice(-7));
            kandidater.add(baraSiffror.slice(-8));
        }
    });

    return Array.from(kandidater).filter(k => k.length >= 4);
}

function likhet(a, b) {
    a = normaliseraNummer(a);
    b = normaliseraNummer(b);

    if (!a || !b) return 0;
    if (a === b) return 100;
    if (a.includes(b) || b.includes(a)) return 85;

    let kort = a.length <= b.length ? a : b;
    let lang = a.length > b.length ? a : b;

    let basta = 0;

    for (let i = 0; i <= lang.length - kort.length; i++) {
        let del = lang.slice(i, i + kort.length);
        let samma = 0;

        for (let j = 0; j < kort.length; j++) {
            if (kort[j] === del[j]) samma++;
        }

        let procent = Math.round((samma / kort.length) * 100);
        if (procent > basta) basta = procent;
    }

    return basta;
}

function poangForKomponent(ocrText, komponent) {
    const kandidater = skapaKandidater(ocrText);
    const komp = normaliseraNummer(komponent.komp);
    const art = normaliseraNummer(komponent.art);

    let poang = 0;
    let detaljer = [];

    kandidater.forEach(k => {
        const kompLikhet = likhet(k, komp);
        const artLikhet = likhet(k, art);

        if (kompLikhet > poang) {
            poang = kompLikhet;
            detaljer = [`Kandidat ${k} liknar komp.nr ${komponent.komp} (${kompLikhet}%)`];
        }

        if (artLikhet > poang) {
            poang = artLikhet;
            detaljer = [`Kandidat ${k} liknar art.nr ${komponent.art} (${artLikhet}%)`];
        }
    });

    return { poang, detaljer, kandidater };
}

function hittaToppMatcher(ocrText) {
    return komponenter
        .map(k => {
            const match = poangForKomponent(ocrText, k);
            return {
                komponent: k,
                poang: match.poang,
                detaljer: match.detaljer,
                kandidater: match.kandidater
            };
        })
        .filter(m => m.poang >= 60)
        .sort((a, b) => b.poang - a.poang)
        .slice(0, 3);
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

        if (k.avvikelse) klass = "avvikelse";
        else if (k.kontrollerad) klass = "ok";

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

    const matcher = hittaToppMatcher(sok);

    if (matcher.length === 0) {
        sokResultat.innerHTML = "<p><strong>Ingen träff.</strong></p>";
        return;
    }

    sokResultat.innerHTML = skapaMatchHtml(matcher, sok, null);
}

function markeraOk(index) {
    const k = komponenter[index];
    k.kontrollerad = true;
    k.avvikelse = false;
    visaLista();
}

function markeraAvvikelse(index) {
    const k = komponenter[index];
    k.kontrollerad = true;
    k.avvikelse = true;
    visaLista();
}

function skapaMatchHtml(matcher, ocrText, usedCanvas) {
    let html = "";

    if (usedCanvas) {
        html += `
            <p><strong>OCR-beskärning:</strong></p>
            <img src="${usedCanvas.toDataURL()}" style="max-width:100%;">
        `;
    }

    html += "<h3>Möjliga träffar</h3>";

    matcher.forEach((m, i) => {
        const k = m.komponent;
        const index = komponenter.indexOf(k);

        html += `
            <div class="match">
                <h3>${i + 1}. ${k.typ}</h3>
                <p>Säkerhet: ${m.poang}%</p>
                <p>Komp.nr: ${k.komp}</p>
                <p>Art.nr: ${k.art}</p>
                <p>${m.detaljer.join(", ")}</p>
                <button onclick="markeraOk(${index})">Markera OK</button>
                <button onclick="markeraAvvikelse(${index})">Markera avvikelse</button>
            </div>
        `;
    });

    html += `
        <p><strong>OCR/text läste:</strong></p>
        <pre>${ocrText}</pre>
    `;

    return html;
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

function roteraCanvas(sourceCanvas, grader) {
    const radians = grader * Math.PI / 180;
    const sin = Math.abs(Math.sin(radians));
    const cos = Math.abs(Math.cos(radians));

    const width = sourceCanvas.width;
    const height = sourceCanvas.height;

    const rotatedCanvas = document.createElement("canvas");
    rotatedCanvas.width = Math.floor(width * cos + height * sin);
    rotatedCanvas.height = Math.floor(width * sin + height * cos);

    const ctx = rotatedCanvas.getContext("2d");
    ctx.translate(rotatedCanvas.width / 2, rotatedCanvas.height / 2);
    ctx.rotate(radians);
    ctx.drawImage(sourceCanvas, -width / 2, -height / 2);

    return rotatedCanvas;
}

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
        ocrStatus.innerHTML = "Tar bild från scan-rutan...";

        const videoWidth = camera.videoWidth;
        const videoHeight = camera.videoHeight;

        snapshot.width = videoWidth;
        snapshot.height = videoHeight;

        const ctx = snapshot.getContext("2d");
        ctx.drawImage(camera, 0, 0, videoWidth, videoHeight);

        const cropX = videoWidth * 0.15;
        const cropY = videoHeight * 0.38;
        const cropWidth = videoWidth * 0.70;
        const cropHeight = videoHeight * 0.24;

        const cropCanvas = document.createElement("canvas");
        cropCanvas.width = cropWidth;
        cropCanvas.height = cropHeight;

        const cropCtx = cropCanvas.getContext("2d");
        cropCtx.drawImage(
            snapshot,
            cropX,
            cropY,
            cropWidth,
            cropHeight,
            0,
            0,
            cropWidth,
            cropHeight
        );

        const rotatedCanvas = roteraCanvas(cropCanvas, rotation);

        ocrStatus.innerHTML = `
            <h3>OCR-beskärning</h3>
            <p>Rotation: ${rotation}°</p>
            <img src="${rotatedCanvas.toDataURL()}" style="max-width:100%;">
            <p>Läser text...</p>
        `;

        const result = await Tesseract.recognize(rotatedCanvas, "eng", {
            logger: function (m) {
                if (m.status) {
                    ocrStatus.innerHTML = `
                        <h3>OCR-beskärning</h3>
                        <p>Rotation: ${rotation}°</p>
                        <img src="${rotatedCanvas.toDataURL()}" style="max-width:100%;">
                        <p>OCR: ${m.status} ${m.progress ? Math.round(m.progress * 100) + "%" : ""}</p>
                    `;
                }
            }
        });

        const text = result.data.text || "";
        const matcher = hittaToppMatcher(text);

        if (matcher.length === 0) {
            ocrStatus.innerHTML = `
                <div class="fel">
                    <h3>Ingen träff</h3>
                    <p><strong>OCR-beskärning:</strong></p>
                    <img src="${rotatedCanvas.toDataURL()}" style="max-width:100%;">
                    <p>OCR läste:</p>
                    <pre>${text}</pre>
                </div>
            `;
            return;
        }

        ocrStatus.innerHTML = skapaMatchHtml(matcher, text, rotatedCanvas);

    } catch (error) {
        ocrStatus.innerHTML = "OCR-fel: " + error.message;
    }
});
