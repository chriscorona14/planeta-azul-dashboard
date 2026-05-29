const fs = require('fs');

let mainJs = fs.readFileSync('main.js', 'utf8');

const startStr = "window.processCxpWorkbook = async function(workbook) {";
const startIdx = mainJs.indexOf(startStr);

if (startIdx === -1) {
    console.log("Could not find start");
    process.exit(1);
}

const endStr = "window.processCxpFile = async function(file) {";
const endIdx = mainJs.indexOf(endStr);

if (endIdx === -1) {
    console.log("Could not find end");
    process.exit(1);
}

const newLogic = `window.processCxpWorkbook = async function(workbook) {
    // 1. Identify sheets
    const names = workbook.SheetNames;
    const historicoName = names.find(n => n.includes("Historico") && n.includes("CXP"));
    const balanzaName = names.find(n => n.includes("Balanza"));
    const analisisName = names.find(n => n.includes("Analisis") || n.includes("Análisis"));
    
    if (!historicoName || !balanzaName || !analisisName) {
        console.warn("Could not find one of the required sheets (Historico CXP, Balanza, Analisis).");
        return;
    }

    const { sheet_to_json } = XLSX.utils;
    const cleanVal = (v) => {
        if (v === undefined || v === null || v === '') return 0;
        if (typeof v === 'number') return v;
        let str = String(v).trim();
        const isNegative = str.startsWith('\'(') || str.startsWith('(');
        str = str.replace(/[^0-9.-]/g, '');
        let n = parseFloat(str);
        if (isNaN(n)) return 0;
        return isNegative ? -Math.abs(n) : n;
    };
    
    // --- PASO 1 - IDENTIFICAR PERIODO ---
    const histRows = sheet_to_json(workbook.Sheets[historicoName], { header: 1 });
    if (histRows.length === 0) return;
    
    const histHeaders = histRows[0];
    let periodColIdx = -1;
    let nombreSocioColIdx = -1;
    let totalCxpColIdx = -1;
    let colNoVencido = -1, col0_30 = -1, col31_60 = -1, col61_90 = -1, col91_120 = -1, col121_150 = -1, col151_180 = -1, col180Mas = -1;

    for (let i = 0; i < histHeaders.length; i++) {
        let val = String(histHeaders[i]).toLowerCase().trim();
        if (val === 'period') periodColIdx = i;
        if (val === 'nombresocio') nombreSocioColIdx = i;
        if (val === 'total saldo cxp') totalCxpColIdx = i;
        if (val === 'saldo no vencido') colNoVencido = i;
        if (val === '0 a 30') col0_30 = i;
        if (val === '31 a 60') col31_60 = i;
        if (val === '61 a 90') col61_90 = i;
        if (val === '91 a 120') col91_120 = i;
        if (val === '121 a 150') col121_150 = i;
        if (val === '151 a 180') col151_180 = i;
        if (val === '> 180') col180Mas = i;
    }

    let maxDate = new Date(0);
    let maxPeriod = "";
    
    for (let i = 1; i < histRows.length; i++) {
        let p = histRows[i][periodColIdx];
        if (p) {
            let parts = String(p).split('/');
            if (parts.length === 2) {
                let m = parseInt(parts[0], 10);
                let y = parseInt(parts[1], 10);
                let d = new Date(y, m - 1, 1);
                if (d > maxDate) {
                    maxDate = d;
                    maxPeriod = String(p);
                }
            }
        }
    }
    
    if (maxPeriod === "") {
        console.warn("No valid periods found in Historico");
        return;
    }
    
    let dates = [];
    let curD = new Date(maxDate.getFullYear(), maxDate.getMonth(), 1);
    for (let i = 0; i < 5; i++) {
        dates.unshift(new Date(curD.getFullYear(), curD.getMonth(), 1));
        curD.setMonth(curD.getMonth() - 1);
    }
    
    const periods = dates.map(d => d.getMonth() + 1 + "/" + d.getFullYear());
    const shortMonthsList = ["ene", "feb", "mar", "abr", "may", "jun", "jul", "ago", "sep", "oct", "nov", "dic"];
    const labels = dates.map(d => shortMonthsList[d.getMonth()] + "-" + d.getFullYear().toString().slice(-2));

    // --- PASO 2 - HOJA BALANZA ---
    const balRows = sheet_to_json(workbook.Sheets[balanzaName], { header: 1 });
    const balYears = balRows[0];
    const balMonthsName = balRows[1];
    
    let balIndices = [];
    for (let i = 0; i < periods.length; i++) {
        let y = dates[i].getFullYear();
        let m = dates[i].getMonth() + 1;
        let foundIdx = -1;
        for (let j = 2; j < balYears.length; j++) {
            if (balYears[j] == y && balMonthsName[j] == m) {
                foundIdx = j;
                break;
            }
        }
        balIndices.push(foundIdx);
    }
    
    const rowCXP = balRows[12];
    const rowOtrasCXP = balRows[21];
    const rowBal = balRows[24];
    
    let arrBalGen = [], arrCXP = [], arrProveedoresProv = [];
    for (let i = 0; i < periods.length; i++) {
        let idx = balIndices[i];
        if (idx !== -1) {
            arrBalGen.push(-cleanVal(rowBal[idx]) / 1000000);
            arrCXP.push(-cleanVal(rowCXP[idx]) / 1000000);
            arrProveedoresProv.push(-cleanVal(rowOtrasCXP[idx]) / 1000000);
        } else {
            arrBalGen.push(0); arrCXP.push(0); arrProveedoresProv.push(0);
        }
    }

    // --- PASO 3 - AGING ---
    let arrCorriente = [0,0,0,0,0];
    let arr0_30 = [0,0,0,0,0];
    let arr31_60 = [0,0,0,0,0];
    let arr61_90 = [0,0,0,0,0];
    let arr91_120 = [0,0,0,0,0];
    let arr121_150 = [0,0,0,0,0];
    let arr151_180 = [0,0,0,0,0];
    let arr180mas = [0,0,0,0,0];

    for (let i = 1; i < histRows.length; i++) {
        let p = String(histRows[i][periodColIdx]);
        let pIdx = periods.indexOf(p);
        if (pIdx !== -1) {
            arrCorriente[pIdx] += -cleanVal(histRows[i][colNoVencido]);
            arr0_30[pIdx] += -cleanVal(histRows[i][col0_30]);
            arr31_60[pIdx] += -cleanVal(histRows[i][col31_60]);
            arr61_90[pIdx] += -cleanVal(histRows[i][col61_90]);
            arr91_120[pIdx] += -cleanVal(histRows[i][col91_120]);
            arr121_150[pIdx] += -cleanVal(histRows[i][col121_150]);
            arr151_180[pIdx] += -cleanVal(histRows[i][col151_180]);
            arr180mas[pIdx] += -cleanVal(histRows[i][col180Mas]);
        }
    }
    
    const toMM = (arr) => arr.map(v => v / 1000000);
    arrCorriente = toMM(arrCorriente);
    arr0_30 = toMM(arr0_30);
    arr31_60 = toMM(arr31_60);
    arr61_90 = toMM(arr61_90);
    arr91_120 = toMM(arr91_120);
    arr121_150 = toMM(arr121_150);
    arr151_180 = toMM(arr151_180);
    arr180mas = toMM(arr180mas);
    
    // --- PASO 4 - TOP 14 PROVEEDORES ---
    let lastMonthProvTotals = {};
    for (let i = 1; i < histRows.length; i++) {
        let p = String(histRows[i][periodColIdx]);
        if (p === maxPeriod) {
            let prov = String(histRows[i][nombreSocioColIdx] || '');
            let v = -cleanVal(histRows[i][totalCxpColIdx]) / 1000000;
            if (!lastMonthProvTotals[prov]) lastMonthProvTotals[prov] = 0;
            lastMonthProvTotals[prov] += v;
        }
    }
    
    let provList = Object.keys(lastMonthProvTotals).map(k => ({name: k, val: lastMonthProvTotals[k]}));
    provList.sort((a,b) => Math.abs(b.val) - Math.abs(a.val));
    let top14Names = provList.slice(0, 14).map(p => p.name);
    
    let top14Saldos = {};
    for (let n of top14Names) {
        top14Saldos[n] = [0,0,0,0,0];
    }
    
    for (let i = 1; i < histRows.length; i++) {
        let p = String(histRows[i][periodColIdx]);
        let pIdx = periods.indexOf(p);
        if (pIdx !== -1) {
            let prov = String(histRows[i][nombreSocioColIdx] || '');
            if (top14Saldos[prov]) {
                top14Saldos[prov][pIdx] += -cleanVal(histRows[i][totalCxpColIdx]) / 1000000;
            }
        }
    }
    
    let arrOtros = [0,0,0,0,0];
    let arrTotal = [0,0,0,0,0];
    for (let i = 0; i < 5; i++) {
        let sumTop14 = 0;
        for (let n of top14Names) {
            sumTop14 += top14Saldos[n][i];
        }
        arrOtros[i] = arrBalGen[i] - sumTop14;
        arrTotal[i] = sumTop14 + arrOtros[i];
    }

    // --- PASO 5 - COSTOS YTD ---
    const anaRows = sheet_to_json(workbook.Sheets[analisisName], { header: 1 });
    const anaDates = anaRows[2]; // Fila 3, índice 2
    const anaCostos = anaRows[38]; // Fila 39, índice 38
    
    let arrCostosYTD = [0,0,0,0,0];
    
    const EXCEL_EPOCH = new Date(1899, 11, 30); // in excel 1= 1900-01-01 but there's leaps
    const getMonthAndYearFromExcel = (cell) => {
        if (!cell) return null;
        if (typeof cell === 'number') {
            let d = new Date(EXCEL_EPOCH.getTime() + cell * 86400000);
            return { m: d.getMonth() + 1, y: d.getFullYear() };
        }
        if (cell instanceof Date) {
            return { m: cell.getMonth() + 1, y: cell.getFullYear() };
        }
        let tk = String(cell).substring(0,10);
        let d = new Date(tk + 'T12:00:00Z');
        if (!isNaN(d.getTime())) return { m: d.getMonth() + 1, y: d.getFullYear() };
        
        let d2 = new Date(cell);
        if (!isNaN(d2.getTime())) return { m: d2.getMonth() + 1, y: d2.getFullYear() };
        return null;
    };
    
    let anaIndices = [];
    for (let i = 0; i < periods.length; i++) {
        let y = dates[i].getFullYear();
        let m = dates[i].getMonth() + 1;
        let foundIdx = -1;
        
        if (anaDates) {
            for (let j = 1; j < anaDates.length; j++) {
                let dt = getMonthAndYearFromExcel(anaDates[j]);
                if (dt && dt.m === m && dt.y === y) {
                    foundIdx = j; break;
                }
            }
        }
        anaIndices.push(foundIdx);
    }
    
    for (let i = 0; i < periods.length; i++) {
        let idx = anaIndices[i];
        if (idx !== -1 && anaCostos) {
            arrCostosYTD[i] = cleanVal(anaCostos[idx]); // Ya en MM
        }
    }
    
    // --- PASO 6 - DPO ---
    let arrDPO = [0,0,0,0,0];
    for (let i = 0; i < 5; i++) {
        let cytd = arrCostosYTD[i];
        if (cytd > 0) {
            arrDPO[i] = Math.round(arrBalGen[i] / (cytd / 30));
        }
    }

    window.cxpStandaloneData = {
        labels,
        periods,
        BalanceGeneral: arrBalGen,
        CXP: arrCXP,
        Provisionales: arrProveedoresProv,
        Corriente: arrCorriente,
        Aging: {
            "0_30": arr0_30,
            "31_60": arr31_60,
            "61_90": arr61_90,
            "91_120": arr91_120,
            "121_150": arr121_150,
            "151_180": arr151_180,
            "180Mas": arr180mas
        },
        Top14Names: top14Names,
        Top14Saldos: top14Saldos,
        OtrosProveedores: arrOtros,
        Total: arrTotal,
        CostosYTD: arrCostosYTD,
        DPO: arrDPO
    };
    
    try {
        const db = await getFinanceDB();
        const tx = db.transaction('finance_cache', 'readwrite');
        tx.objectStore('finance_cache').put({ data: window.cxpStandaloneData, timestamp: Date.now() }, 'MASTER_STANDALONE_CXP_DATA');
    } catch(e) {
        console.error("Error saving to standalone indexeddb cxp", e);
    }
    
    if (window.currentActiveView === 'view-cxp') { 
        if (typeof window.renderCxpView === 'function') {
            window.renderCxpView(window.cxpStandaloneData);
        }
    }
}
`;

const res = mainJs.slice(0, startIdx) + newLogic + '\n' + mainJs.slice(endIdx);
fs.writeFileSync('main.js', res);

console.log("Replaced processCxpWorkbook");
