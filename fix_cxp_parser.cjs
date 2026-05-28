const fs = require('fs');

const RAW_CXP_PARSER = `
    window.processCxpWorkbook = async function(workbook) {
        let cxpSheetName = workbook.SheetNames.find(n => /cxp|cuentas por pagar|aging|antiguedad|proveedores/i.test(n)) || workbook.SheetNames[0];
        let sheet = workbook.Sheets[cxpSheetName];
        let rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: null });

        const normalizeText = (t) => { if (!t) return ""; return t.toString().toLowerCase().normalize("NFD").replace(/[\\u0300-\\u036f]/g, "").trim(); };
        const cleanVal = (v) => {
            if (v === undefined || v === null || v === '') return 0;
            if (typeof v === 'number') return v;
            let str = v.toString().replace(/,/g, '');
            let n = parseFloat(str);
            return isNaN(n) ? 0 : n;
        };

        // Detect if RAW format (has NombreSocio or CodSocio)
        let isRawFormat = false;
        let headerRowIdx = -1;
        for (let i = 0; i < Math.min(rows.length, 10); i++) {
            if (!rows[i]) continue;
            let rowStr = rows[i].map(c => normalizeText(c)).join("|");
            if (rowStr.includes('nombresocio') || rowStr.includes('codsocio') || rowStr.includes('saldo no vencido')) {
                isRawFormat = true;
                headerRowIdx = i;
                break;
            }
        }

        if (isRawFormat) {
            console.log("Detected RAW CxP structure, parsing aggregation...");
            const headers = rows[headerRowIdx].map(c => normalizeText(c));
            
            const getIdx = (k) => headers.findIndex(h => h.includes(k));
            
            const idxNombre = getIdx("nombresocio") !== -1 ? getIdx("nombresocio") : getIdx("socio");
            const idxPeriod = getIdx("period");
            const idxFecha = getIdx("fechagenerado");
            
            // Amount indices
            const idxTotal = getIdx("total saldo");
            const idxNoVencido = getIdx("saldo no vencido");
            const idxVencido = getIdx("saldo vencido");
            const idx0_30 = headers.findIndex(h => h.match(/^0( a |-)30/));
            const idx31_60 = headers.findIndex(h => h.match(/^31( a |-)60/));
            const idx61_90 = headers.findIndex(h => h.match(/^61( a |-)90/));
            const idx91_120 = headers.findIndex(h => h.match(/^91( a |-)120/));
            const idx121_150 = headers.findIndex(h => h.match(/^121( a |-)150/));
            const idx151_180 = headers.findIndex(h => h.match(/^151( a |-)180/));
            const idx180Mas = headers.findIndex(h => h.match(/> ?180/));

            const aggregatedByPeriod = {};

            const monthNames = ["ene", "feb", "mar", "abr", "may", "jun", "jul", "ago", "sep", "oct", "nov", "dic"];

            for (let i = headerRowIdx + 1; i < rows.length; i++) {
                const row = rows[i];
                if (!row || !row.length) continue;
                
                let periodKey = "unknown";
                if (idxPeriod !== -1 && row[idxPeriod]) {
                    // Custom parsing for 'ene-26' -> '0-2026'
                    const p = normalizeText(row[idxPeriod]);
                    const matchedM = monthNames.findIndex(m => p.includes(m));
                    const yMatch = p.match(/\\d{2,4}$/);
                    if (matchedM !== -1 && yMatch) {
                        let y = parseInt(yMatch[0]);
                        if (y < 100) y += 2000;
                        periodKey = \`\${matchedM}-\${y}\`;
                    }
                } else if (idxFecha !== -1 && row[idxFecha]) {
                    // Parse '31/1/2026' or excel date
                    let dateObj;
                    let c = row[idxFecha];
                    if (c instanceof Date) dateObj = c;
                    else if (typeof c === 'number') dateObj = new Date((c - 25569) * 86400 * 1000);
                    else {
                        const parts = c.split(' ')[0].split('/'); // assumed dd/mm/yyyy
                        if (parts.length >= 3) {
                            dateObj = new Date(parseInt(parts[2]), parseInt(parts[1])-1, parseInt(parts[0]));
                        }
                    }
                    if (dateObj && !isNaN(dateObj.getTime())) {
                        periodKey = \`\${dateObj.getMonth()}-\${dateObj.getFullYear()}\`;
                    }
                }
                
                if (periodKey === "unknown") continue;

                if (!aggregatedByPeriod[periodKey]) {
                    aggregatedByPeriod[periodKey] = {
                        providerRows: [],
                        corriente: 0,
                        dias0_30: 0, dias31_60: 0, dias61_90: 0, dias91_120: 0,
                        dias121_150: 0, dias151_180: 0, dias180Mas: 0,
                        cxpTotal: 0, vencido: 0,
                        
                        alplaHispaniola: 0, polyplas: 0, grupoRojas: 0, raviCaribe: 0, valcopack: 0,
                        termopack: 0, cartoneraApolo: 0, multiplast: 0, flexopack: 0, etiofset: 0,
                        smurfit: 0, plasticosCaribe: 0, industriasNacionales: 0, distribuidoraCorripo: 0,
                        otrosProveedores: 0,
                        isProjectedDetail: false
                    };
                }

                const agg = aggregatedByPeriod[periodKey];
                const provider = normalizeText(row[idxNombre]);
                const total = cleanVal(row[idxTotal]);
                
                agg.corriente += cleanVal(row[idxNoVencido]);
                agg.dias0_30 += idx0_30 !== -1 ? cleanVal(row[idx0_30]) : 0;
                agg.dias31_60 += idx31_60 !== -1 ? cleanVal(row[idx31_60]) : 0;
                agg.dias61_90 += idx61_90 !== -1 ? cleanVal(row[idx61_90]) : 0;
                agg.dias91_120 += idx91_120 !== -1 ? cleanVal(row[idx91_120]) : 0;
                agg.dias121_150 += idx121_150 !== -1 ? cleanVal(row[idx121_150]) : 0;
                agg.dias151_180 += idx151_180 !== -1 ? cleanVal(row[idx151_180]) : 0;
                agg.dias180Mas += idx180Mas !== -1 ? cleanVal(row[idx180Mas]) : 0;
                agg.cxpTotal += total;
                agg.vencido += idxVencido !== -1 ? cleanVal(row[idxVencido]) : 0;

                // Match suppliers
                if (provider.includes('alpla')) agg.alplaHispaniola += total;
                else if (provider.includes('polyplas')) agg.polyplas += total;
                else if (provider.includes('rojas')) agg.grupoRojas += total;
                else if (provider.includes('ravi')) agg.raviCaribe += total;
                else if (provider.includes('valcopack')) agg.valcopack += total;
                else if (provider.includes('termopack')) agg.termopack += total;
                else if (provider.includes('apolo')) agg.cartoneraApolo += total;
                else if (provider.includes('multiplast')) agg.multiplast += total;
                else if (provider.includes('flexopack')) agg.flexopack += total;
                else if (provider.includes('etrefset') || provider.includes('etiquetas') || provider.includes('etiofset')) agg.etiofset += total;
                else if (provider.includes('smurfit')) agg.smurfit += total;
                else if (provider.includes('plasticos del caribe') || provider.includes('plisticos del caribe')) agg.plasticosCaribe += total;
                else if (provider.includes('industrias nacionales')) agg.industriasNacionales += total;
                else if (provider.includes('corripo')) agg.distribuidoraCorripo += total;
                else agg.otrosProveedores += total;
            }

            // Sync with globalFinancialData
            window.globalFinancialData.forEach(point => {
                const pKey = \`\${point.sortDate.getMonth()}-\${point.sortDate.getFullYear()}\`;
                if (aggregatedByPeriod[pKey]) {
                    const agg = aggregatedByPeriod[pKey];
                    agg.balanceGeneral = agg.cxpTotal;
                    agg.cxpBase = agg.cxpTotal;
                    agg.provisionSinFactura = 0; // Not detailed in raw file
                    agg.conciliacion = 0; // exact match because we sum it
                    point.cxpDetail = agg;
                }
            });

            // Fallback costs calculations
            let yearlyCostosGastoYTD = {};
            window.globalFinancialData.forEach(point => {
                const year = point.sortDate.getFullYear();
                if (yearlyCostosGastoYTD[year] === undefined) {
                    yearlyCostosGastoYTD[year] = 0;
                }
                
                if (point.cxpDetail) {
                    const costos = point.pnl ? (point.pnl.costos || 0) : 0;
                    const opex = point.pnl ? (point.pnl.opex || 0) : 0;
                    const capex = point.cashflowDetail ? (point.cashflowDetail.capex || 0) : 0;
                    const costosGastoMensual = Math.abs(costos) + Math.abs(opex) + Math.abs(capex);
                    yearlyCostosGastoYTD[year] += costosGastoMensual;
                    point.cxpDetail.costosGastoYtd = yearlyCostosGastoYTD[year];
                    
                    const elapsed_months = (point.sortDate.getMonth() === 11 && year === 2025) ? 12 : (point.sortDate.getMonth() + 1);
                    const days = elapsed_months * 30.4;
                    if (point.cxpDetail.costosGastoYtd > 0) {
                        point.cxpDetail.dpo = Math.round((point.cxpDetail.cxpTotal / point.cxpDetail.costosGastoYtd) * days);
                    } else {
                        point.cxpDetail.dpo = 0;
                    }
                }
            });

        } else {
            // ORIGINAL PIVOT LOGIC HERE
            // ... I will preserve the original logic dynamically using regex replace
`;

let mainJs = fs.readFileSync('main.js', 'utf8');

// Find processCxpWorkbook function boundaries
const startIdx = mainJs.indexOf('window.processCxpWorkbook = async function(workbook) {');
const endIdx = mainJs.indexOf('window.processCxpFile = async function(file) {');

if (startIdx !== -1 && endIdx !== -1) {
    const originalLogic = mainJs.slice(startIdx + 'window.processCxpWorkbook = async function(workbook) {'.length, endIdx);
    
    // We want to skip the first 3 lines of original logic, since we redefine them
    const lines = originalLogic.split('\n');
    const pivotLogic = lines.slice(4).join('\n');
    
    const newFunction = RAW_CXP_PARSER + pivotLogic + "\n    }\n";
    
    fs.writeFileSync('main.js', mainJs.slice(0, startIdx) + newFunction + mainJs.slice(endIdx));
    console.log('Successfully injected raw CSV parser supporting logic for CxP.');
} else {
    console.log('Could not locate boundaries for processCxpWorkbook in main.js.');
}
