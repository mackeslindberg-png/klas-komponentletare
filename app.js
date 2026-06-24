const excelFile = document.getElementById("excelFile");
const listImage = document.getElementById("listImage");
const resultat = document.getElementById("resultat");
const startCamera = document.getElementById("startCamera");
const startScan = document.getElementById("startScan");
const stopScan = document.getElementById("stopScan");
const scanInfo = document.getElementById("scanInfo");
const camera = document.getElementById("camera");
const snapshot = document.getElementById("snapshot");
const ocrStatus = document.getElementById("ocrStatus");

let komponenter = [];
let scanning = false;
let scanningBusy = false;
let scanTimer = null;

function normalisera(varde) {
    return String(varde || "")
        .toUpperCase()
        .replace(/[OQ]/g, "0")
        .replace(/[IL]/g, "1")
        .replace(/S/g, "5")
        .replace(/B/g, "8")
        .replace(/[^A-Z0-9]/g, "");
}

function likhet(a, b) {
    a = normalisera(a);
    b = normalisera(b);

    if (!a || !b) return 0;
    if (a === b) return 100;
    if (a.includes(b) || b.includes(a)) return 92;

    let kort = a.length <= b.length ? a : b;
    let lang = a.length > b.length ? a : b;
    let basta = 0;

    for (let i = 0; i <= lang.length - kort.length; i++) {
        let del = lang.slice(i, i + kort.length);
        let lika = 0;

        for (let j = 0; j < kort.length; j++) {
            if (kort[j] === del[j]) lika++;
        }

        let score = Math.round((lika / kort.length) * 100);
        if (score > basta) basta = score;
    }

    return basta;
}

function skapaKandidater(text) {
    const ren = normalisera(text);
    const delar = ren.match(/[A-Z0-9]{4,}/g) || [];
    const kandidater = new Set();

    delar.forEach(d => {
        kandidater.add(d);

        const siffror = d.replace(/[^0-9]/g, "");
        if (siffror.length >= 4) kandidater.add(siffror);

        if (d.length >= 6) {
            kandidater.add(d.slice(-6));
            kandidater.add(d.slice(-7));
            kandidater.add(d.slice(-8));
            kandidater.add(d.slice(-9));
            kandidater.add(d.slice(-10));
        }

        if (siffror.length >= 6) {
            kandidater.add(siffror.slice(-6));
            kandidater.add(siffror.slice(-7));
            kandidater.add(siffror.slice(-8));
            kandidater.add(siffror.slice(-9));
            kandidater.add(siffror.slice(-10));
        }
    });

    return Array.from(kandidater).filter(k => k.length >= 4);
}

function matchaMotLista(text) {
    const kandidater = skapaKandidater(text);
    let matcher = [];

    komponenter.forEach(k => {
        let bästaPoäng = 0;
        let bästaText = "";

        kandidater.forEach(kandidat => {
            const kompPoäng = likhet(kandidat, k.komp);
            const artPoäng = likhet(kandidat, k.art);

            if (kompPoäng > bästaPoäng) {
                bästaPoäng = kompPoäng;
                bästaText = `Kandidat ${kandidat} liknar komp.nr ${k.komp}`;
            }

            if (artPoäng > bästaPoäng) {
                bästaPoäng = artPoäng;
                bästaText = `Kandidat ${kandidat} liknar art.nr ${k.art}`;
            }
        });

        if (bästaPoäng >= 60) {
            matcher.push({
                komponent: k,
                poäng: bästaPoäng,
                text: bästaText
            });
        }
    });

    return matcher
        .sort((a, b) => b.poäng - a.poäng)
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
    html += `<p>Antal komponenter: ${komponenter.length}</p>`;
    html += `<p>Kontrollerade: ${kontrollerade} / ${komponenter.length}</p>`;
    html += `<p>Avvikelser: ${avvikelser}</p>`;

    html += `
        <label>Sök manuellt:</label><br>
        <input type="text" id="sokRuta" placeholder="Ex: 03H350105">
        <button onclick="sokManuellt()">Sök</button>
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

function sokManuellt() {
    const text = document.getElementById("sokRuta").value;
    const matcher = matchaMotLista(text);
    document.getElementById("sokResultat").innerHTML = skapaMatchHtml(matcher, text, null);
}

function skapaMatchHtml(matcher, text, bild) {
    if (!matcher.length) {
        return `
            <div class="fel">
                <h3>Ingen träff</h3>
                <p>Läst text:</p>
                <pre>${text}</pre>
            </div>
        `;
    }

    let html = "";

    if (bild) {
        html += `<img src="${bild.toDataURL()}" style="max-width:100%;">`;
    }

    html += "<h3>Möjliga träffar</h3>";

    matcher.forEach((m, i) => {
        const k = m.komponent;
        const index = komponenter.indexOf(k);

        html += `
            <div class="match">
                <h3>${i + 1}. ${k.typ}</h3>
                <p>Säkerhet: ${m.poäng}%</p>
                <p>Komp.nr: ${k.komp}</p>
                <p>Art.nr: ${k.art}</p>
                <p>${m.text}</p>
                <button onclick="markeraOk(${index})">Markera OK</button>
                <button onclick="markeraAvvikelse(${index})">Markera avvikelse</button>
            </div>
        `;
    });

    html += `<pre>${text}</pre>`;
    return html;
}

function markeraOk(index) {
    komponenter[index].kontrollerad = true;
    komponenter[index].avvikelse = false;
    visaLista();
}

function markeraAvvikelse(index) {
    komponenter[index].kontrollerad = true;
    komponenter[index].avvikelse = true;
    visaLista();
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

startScan.addEventListener("click", function () {
    if (!camera.srcObject) {
        alert("Starta kameran först.");
        return;
    }

    if (!komponenter.length) {
        alert("Läs in Excel-filen först.");
        return;
    }

    scanning = true;
    scanInfo.innerHTML = "Scan: aktiv";
    ocrStatus.innerHTML = "Smart scan startad.";

    scanTimer = setInterval(körSmartScan, 2500);
});

stopScan.addEventListener("click", function () {
    scanning = false;
    scanningBusy = false;
    clearInterval(scanTimer);
    scanInfo.innerHTML = "Scan: stoppad";
});

function beskärScanRuta() {
    const videoWidth = camera.videoWidth;
    const videoHeight = camera.videoHeight;

    snapshot.width = videoWidth;
    snapshot.height = videoHeight;

    const ctx = snapshot.getContext("2d");
    ctx.drawImage(camera, 0, 0, videoWidth, videoHeight);

    const cropX = videoWidth * 0.10;
    const cropY = videoHeight * 0.20;
    const cropWidth = videoWidth * 0.80;
    const cropHeight = videoHeight * 0.60;

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

    return cropCanvas;
}

function roteraCanvas(sourceCanvas, grader) {
    const radians = grader * Math.PI / 180;
    const sin = Math.abs(Math.sin(radians));
    const cos = Math.abs(Math.cos(radians));

    const width = sourceCanvas.width;
    const height = sourceCanvas.height;

    const canvas = document.createElement("canvas");
    canvas.width = Math.floor(width * cos + height * sin);
    canvas.height = Math.floor(width * sin + height * cos);

    const ctx = canvas.getContext("2d");
    ctx.translate(canvas.width / 2, canvas.height / 2);
    ctx.rotate(radians);
    ctx.drawImage(sourceCanvas, -width / 2, -height / 2);

    return canvas;
}

async function körSmartScan() {
    if (!scanning || scanningBusy) return;

    scanningBusy = true;

    try {
        const crop = beskärScanRuta();
        const rotationer = [0, 90, 180, 270];

        let bästaMatcher = [];
        let bästaText = "";
        let bästaBild = crop;

        for (const grad of rotationer) {
            const bild = roteraCanvas(crop, grad);

            ocrStatus.innerHTML = `
                <p>Smart scan körs...</p>
                <p>Testar rotation: ${grad}°</p>
                <img src="${bild.toDataURL()}" style="max-width:100%;">
            `;

            const result = await Tesseract.recognize(bild, "eng");
            const text = result.data.text || "";
            const matcher = matchaMotLista(text);

            if (matcher.length && (!bästaMatcher.length || matcher[0].poäng > bästaMatcher[0].poäng)) {
                bästaMatcher = matcher;
                bästaText = text;
                bästaBild = bild;
            }

            if (bästaMatcher.length && bästaMatcher[0].poäng >= 90) {
                break;
            }
        }

        if (!bästaMatcher.length) {
            ocrStatus.innerHTML = `
                <div class="fel">
                    <h3>Ingen träff ännu</h3>
                    <p>Fortsätter scanna...</p>
                    <img src="${bästaBild.toDataURL()}" style="max-width:100%;">
                </div>
            `;
        } else {
            const bästa = bästaMatcher[0];
            const k = bästa.komponent;

            if (bästa.poäng >= 90) {
                k.kontrollerad = true;
                k.avvikelse = false;
                visaLista();

                ocrStatus.innerHTML = `
                    <div class="match">
                        <h3>✅ AUTO OK</h3>
                        <p><strong>${k.typ}</strong></p>
                        <p>Säkerhet: ${bästa.poäng}%</p>
                        <p>Komp.nr: ${k.komp}</p>
                        <p>Art.nr: ${k.art}</p>
                        <p>${bästa.text}</p>
                        <img src="${bästaBild.toDataURL()}" style="max-width:100%;">
                    </div>
                `;
            } else {
                ocrStatus.innerHTML = skapaMatchHtml(bästaMatcher, bästaText, bästaBild);
            }
        }

    } catch (error) {
        ocrStatus.innerHTML = "Scan-fel: " + error.message;
    }

    scanningBusy = false;
}
listImage.addEventListener("change", async function () {
    const file = listImage.files[0];

    if (!file) {
        resultat.innerHTML = "Ingen bild vald.";
        return;
    }

    resultat.innerHTML = "Läser komponentlista från bild...";

    try {
        const result = await Tesseract.recognize(file, "eng", {
            logger: function (m) {
                if (m.status) {
                    resultat.innerHTML =
                        "OCR lista: " + m.status +
                        (m.progress ? " " + Math.round(m.progress * 100) + "%" : "");
                }
            }
        });

        const text = result.data.text || "";

        byggListaFranBildText(text);

    } catch (error) {
        resultat.innerHTML = "Kunde inte läsa listbilden: " + error.message;
    }
});

function byggListaFranBildText(text) {
    const rader = text
        .split("\n")
        .map(rad => rad.trim())
        .filter(rad => rad.length > 0);

    let hittade = [];

    rader.forEach(rad => {
        const nummer = rad.match(/[A-Z0-9][A-Z0-9\s\-]{3,}/gi) || [];

        if (nummer.length >= 2) {
            const renRad = rad.replace(/\s+/g, " ").trim();

            hittade.push({
                rå: renRad,
                typ: "Okänd komponent",
                komp: nummer[0].trim(),
                art: nummer[nummer.length - 1].trim(),
                kontrollerad: false,
                avvikelse: false
            });
        }
    });

    if (hittade.length === 0) {
        resultat.innerHTML = `
            <div class="fel">
                <h3>Ingen lista hittad</h3>
                <p>OCR läste:</p>
                <pre>${text}</pre>
            </div>
        `;
        return;
    }

    komponenter = hittade;

    let html = `
        <div class="match">
            <h3>Lista skapad från bild</h3>
            <p>Hittade ${komponenter.length} möjliga rader.</p>
            <p><strong>Viktigt:</strong> kontrollera att komp.nr och art.nr hamnat rätt.</p>
        </div>
    `;

    html += "<h3>OCR-text</h3>";
    html += `<pre>${text}</pre>`;

    html += "<h3>Tolkad lista</h3><ul>";

    komponenter.forEach(k => {
        html += `
            <li>
                <strong>${k.typ}</strong><br>
                Komp.nr: ${k.komp}<br>
                Art.nr: ${k.art}<br>
                <small>${k.rå}</small>
            </li>
        `;
    });

    html += "</ul>";

    resultat.innerHTML = html;
}
