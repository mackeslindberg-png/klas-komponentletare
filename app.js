const excelFile = document.getElementById("excelFile");
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
let väntarPåGodkännande = false;

function normalisera(varde) {
    return String(varde || "")
        .toUpperCase()
        .replace(/[OQ]/g, "0")
        .replace(/[IL]/g, "1")
        .replace(/S/g, "5")
        .replace(/B/g, "8")
        .replace(/Z/g, "2")
        .replace(/G/g, "6")
        .replace(/[^A-Z0-9]/g, "");
}

function baraSiffror(varde) {
    return normalisera(varde).replace(/[^0-9]/g, "");
}

function levenshtein(a, b) {
    a = String(a || "");
    b = String(b || "");

    const dp = Array.from({ length: a.length + 1 }, () =>
        Array(b.length + 1).fill(0)
    );

    for (let i = 0; i <= a.length; i++) dp[i][0] = i;
    for (let j = 0; j <= b.length; j++) dp[0][j] = j;

    for (let i = 1; i <= a.length; i++) {
        for (let j = 1; j <= b.length; j++) {
            const kostnad = a[i - 1] === b[j - 1] ? 0 : 1;
            dp[i][j] = Math.min(
                dp[i - 1][j] + 1,
                dp[i][j - 1] + 1,
                dp[i - 1][j - 1] + kostnad
            );
        }
    }

    return dp[a.length][b.length];
}

function likhetStrikt(kandidat, target) {
    kandidat = baraSiffror(kandidat);
    target = baraSiffror(target);

    if (!kandidat || !target) return 0;

    // Minst 6 tecken, annars ska den inte kunna bli en träff.
    if (kandidat.length < 6 || target.length < 6) return 0;

    if (kandidat === target) return 100;

    // Om OCR fått med extra tecken före/efter.
    if (kandidat.includes(target) || target.includes(kandidat)) {
        const kortaste = Math.min(kandidat.length, target.length);
        if (kortaste >= 7) return 96;
    }

    // Jämför bara om längderna är rimligt nära.
    const längdSkillnad = Math.abs(kandidat.length - target.length);
    if (längdSkillnad > 2) return 0;

    const dist = levenshtein(kandidat, target);
    const maxLen = Math.max(kandidat.length, target.length);
    const score = Math.round((1 - dist / maxLen) * 100);

    return score;
}

function skapaKandidater(text) {
    const ren = normalisera(text);
    const delar = ren.match(/[A-Z0-9]{6,}/g) || [];
    const kandidater = new Set();

    delar.forEach(d => {
        const siffror = d.replace(/[^0-9]/g, "");

        // Hela OCR-raden
        if (d.length >= 6) kandidater.add(d);
        if (siffror.length >= 6) kandidater.add(siffror);

        // Sista 6-12 tecknen, bra för t.ex. P17507546#T26224368PG
        for (let len = 6; len <= 12; len++) {
            if (d.length >= len) kandidater.add(d.slice(-len));
            if (siffror.length >= len) kandidater.add(siffror.slice(-len));
        }
    });

    return Array.from(kandidater)
        .filter(k => baraSiffror(k).length >= 6)
        .sort((a, b) => baraSiffror(b).length - baraSiffror(a).length);
}

function matchaMotLista(text) {
    const kandidater = skapaKandidater(text);
    let artMatcher = [];

    // 1. ART.NR FÖRST
    komponenter.forEach(k => {
        const art = baraSiffror(k.art);
        if (!art || art.length < 6) return;

        kandidater.forEach(kandidat => {
            const kandidatSiffror = baraSiffror(kandidat);
            if (kandidatSiffror.length < 6) return;

            const poäng = likhetStrikt(kandidatSiffror, art);

            if (poäng >= 80) {
                artMatcher.push({
                    komponent: k,
                    poäng,
                    text: `Kandidat ${kandidatSiffror} matchar ART.NR ${k.art}`,
                    matchTyp: "art"
                });
            }
        });
    });

    artMatcher = artMatcher
        .sort((a, b) => b.poäng - a.poäng)
        .slice(0, 3);

    if (artMatcher.length > 0) {
        return artMatcher;
    }

    // 2. KOMP.NR ENDAST OM INGET ART.NR HITTAS
    let kompMatcher = [];

    komponenter.forEach(k => {
        const komp = baraSiffror(k.komp);
        if (!komp || komp.length < 6) return;

        kandidater.forEach(kandidat => {
            const kandidatSiffror = baraSiffror(kandidat);
            if (kandidatSiffror.length < 6) return;

            const poäng = likhetStrikt(kandidatSiffror, komp);

            if (poäng >= 85) {
                kompMatcher.push({
                    komponent: k,
                    poäng,
                    text: `Kandidat ${kandidatSiffror} matchar komp.nr ${k.komp}`,
                    matchTyp: "komp"
                });
            }
        });
    });

    return kompMatcher
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
        väntarPåGodkännande = false;

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
        <input type="text" id="sokRuta" placeholder="Ex: 0326062014">
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

    const bästa = matcher[0];

    if (bästa.poäng >= 90) {
        const k = bästa.komponent;
        const index = komponenter.indexOf(k);

        html += `
            <div class="match">
                <h3>Vill du godkänna?</h3>
                <p><strong>${k.typ}</strong></p>
                <p>Säkerhet: ${bästa.poäng}%</p>
                <p>Träff på: ${bästa.matchTyp === "art" ? "Art.nr" : "Komp.nr"}</p>
                <p>Komp.nr: ${k.komp}</p>
                <p>Art.nr: ${k.art}</p>
                <p>${bästa.text}</p>
                <button onclick="markeraOk(${index})">Godkänn</button>
                <button onclick="markeraAvvikelse(${index})">Avvikelse</button>
                <button onclick="ignoreraTraff()">Ignorera</button>
            </div>
        `;
    }

    html += "<h3>Möjliga träffar</h3>";

    matcher.forEach((m, i) => {
        const k = m.komponent;
        const index = komponenter.indexOf(k);

        html += `
            <div class="match">
                <h3>${i + 1}. ${k.typ}</h3>
                <p>Säkerhet: ${m.poäng}%</p>
                <p>Träff på: ${m.matchTyp === "art" ? "Art.nr" : "Komp.nr"}</p>
                <p>Komp.nr: ${k.komp}</p>
                <p>Art.nr: ${k.art}</p>
                <p>${m.text}</p>
                <button onclick="markeraOk(${index})">Godkänn</button>
                <button onclick="markeraAvvikelse(${index})">Avvikelse</button>
            </div>
        `;
    });

    html += `<pre>${text}</pre>`;
    return html;
}

function ignoreraTraff() {
    väntarPåGodkännande = false;
    ocrStatus.innerHTML = "Träff ignorerad. Tryck Starta smart scan för att fortsätta.";
    scanInfo.innerHTML = "Scan: pausad";
}

function markeraOk(index) {
    väntarPåGodkännande = false;
    komponenter[index].kontrollerad = true;
    komponenter[index].avvikelse = false;
    ocrStatus.innerHTML = `Godkänd: ${komponenter[index].typ}`;
    scanInfo.innerHTML = "Scan: pausad";
    visaLista();
}

function markeraAvvikelse(index) {
    väntarPåGodkännande = false;
    komponenter[index].kontrollerad = true;
    komponenter[index].avvikelse = true;
    ocrStatus.innerHTML = `Avvikelse markerad: ${komponenter[index].typ}`;
    scanInfo.innerHTML = "Scan: pausad";
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

    väntarPåGodkännande = false;
    scanning = true;
    scanningBusy = false;
    scanInfo.innerHTML = "Scan: aktiv";
    ocrStatus.innerHTML = "Smart scan startad.";

    if (scanTimer) clearInterval(scanTimer);

    scanTimer = setInterval(körSmartScan, 2500);
});

stopScan.addEventListener("click", function () {
    scanning = false;
    scanningBusy = false;
    väntarPåGodkännande = false;

    if (scanTimer) {
        clearInterval(scanTimer);
        scanTimer = null;
    }

    scanInfo.innerHTML = "Scan: stoppad";
});

function stoppaScanVidTraff() {
    väntarPåGodkännande = true;
    scanning = false;
    scanningBusy = false;

    if (scanTimer) {
        clearInterval(scanTimer);
        scanTimer = null;
    }

    scanInfo.innerHTML = "Scan: pausad - träff hittad";
}

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
    cropCtx.drawImage(snapshot, cropX, cropY, cropWidth, cropHeight, 0, 0, cropWidth, cropHeight);

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
    if (!scanning || scanningBusy || väntarPåGodkännande) return;

    scanningBusy = true;

    try {
        const crop = beskärScanRuta();
        const rotationer = [0, 90, 180, 270];

        let bästaMatcher = [];
        let bästaText = "";
        let bästaBild = crop;

        for (const grad of rotationer) {
            if (!scanning || väntarPåGodkännande) break;

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

            if (bästaMatcher.length && bästaMatcher[0].poäng >= 90) break;
        }

        if (!bästaMatcher.length) {
            if (!väntarPåGodkännande) {
                ocrStatus.innerHTML = `
                    <div class="fel">
                        <h3>Ingen träff ännu</h3>
                        <p>Fortsätter scanna...</p>
                        <img src="${bästaBild.toDataURL()}" style="max-width:100%;">
                    </div>
                `;
            }
        } else {
            stoppaScanVidTraff();
            ocrStatus.innerHTML = skapaMatchHtml(bästaMatcher, bästaText, bästaBild);
            return;
        }

    } catch (error) {
        ocrStatus.innerHTML = "Scan-fel: " + error.message;
    }

    scanningBusy = false;
}
