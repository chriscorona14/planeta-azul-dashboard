import fs from 'fs';

let content = fs.readFileSync('main.js', 'utf8');

// Replace the put for MASTER_FINANCE_KEY in processCxpWorkbook
content = content.replace(/tx\.objectStore\('finance_cache'\)\.put\(window\.globalFinancialData,\s*'MASTER_FINANCE_KEY'\);/g, "tx.objectStore('finance_cache').put({ data: window.globalFinancialData, timestamp: Date.now() }, 'MASTER_FINANCE_KEY');");

// Put the CXP_STANDALONE_KEY logic when globalFinancialData is empty
content = content.replace(/window\.cachedStandaloneCxp = \{ indices, cxpRows, rows \};\n\s*\}\n\s*\};\n\n\s*window\.processCxpFile = async function/g, `window.cachedStandaloneCxp = { indices, cxpRows, rows };
            try {
                const db = await new Promise(r => { const req = indexedDB.open('PlanetaAzulDB', 3); req.onsuccess = () => r(req.result); });
                const tx = db.transaction('finance_cache', 'readwrite');
                tx.objectStore('finance_cache').put({ data: window.cachedStandaloneCxp, timestamp: Date.now() }, 'CXP_STANDALONE_KEY');
            } catch(e) {
                console.error("Error saving standalone CXP to indexedDB:", e);
            }
        }
    };

    window.applyCachedStandaloneCxp = async function() {
        if (!window.cachedStandaloneCxp || !window.globalFinancialData || window.globalFinancialData.length === 0) return;
        const { indices, cxpRows, rows } = window.cachedStandaloneCxp;
        
        window.globalFinancialData.forEach(point => {
            const key = \`\${point.sortDate.getMonth()}-\${getSortYear(point)}\`;
            const idx = indices[key];
            if (idx !== undefined && idx !== -1) {
                const cleanVal = (val) => {
                    if (val === undefined || val === null || val === '') return 0;
                    if (typeof val === 'number') return val;
                    return parseFloat(val.toString().trim().replace(/,/g, '')) || 0;
                };
                const getCxpVal = (row) => row ? cleanVal(row[idx]) : 0;
                
                const cxpObj = {
                    provisionSinFactura: getCxpVal(cxpRows.provisionSinFactura),
                    corriente: getCxpVal(cxpRows.corriente),
                    dias0_30: getCxpVal(cxpRows.dias0_30),
                    dias31_60: getCxpVal(cxpRows.dias31_60),
                    dias61_90: getCxpVal(cxpRows.dias61_90),
                    dias91_120: getCxpVal(cxpRows.dias91_120),
                    dias121_150: getCxpVal(cxpRows.dias121_150),
                    dias151_180: getCxpVal(cxpRows.dias151_180),
                    dias180Mas: getCxpVal(cxpRows.dias180Mas),
                    alplaHispaniola: getCxpVal(cxpRows.alplaHispaniola),
                    polyplas: getCxpVal(cxpRows.polyplas),
                    grupoRojas: getCxpVal(cxpRows.grupoRojas),
                    raviCaribe: getCxpVal(cxpRows.raviCaribe),
                    valcopack: getCxpVal(cxpRows.valcopack),
                    termopack: getCxpVal(cxpRows.termopack),
                    cartoneraApolo: getCxpVal(cxpRows.cartoneraApolo),
                    multiplast: getCxpVal(cxpRows.multiplast),
                    flexopack: getCxpVal(cxpRows.flexopack),
                    etiofset: getCxpVal(cxpRows.etiofset),
                    smurfit: getCxpVal(cxpRows.smurfit),
                    plasticosCaribe: getCxpVal(cxpRows.plasticosCaribe),
                    industriasNacionales: getCxpVal(cxpRows.industriasNacionales),
                    distribuidoraCorripo: getCxpVal(cxpRows.distribuidoraCorripo),
                    otrosProveedores: getCxpVal(cxpRows.otrosProveedores),
                    costosGastoYtd: getCxpVal(cxpRows.costosGastoYtd),
                    dpo: getCxpVal(cxpRows.dpo)
                };

                const totalAntiguedad = (cxpObj.provisionSinFactura || 0) + (cxpObj.corriente || 0) + 
                                        (cxpObj.dias0_30 || 0) + (cxpObj.dias31_60 || 0) + 
                                        (cxpObj.dias61_90 || 0) + (cxpObj.dias91_120 || 0) + 
                                        (cxpObj.dias121_150 || 0) + (cxpObj.dias151_180 || 0) + 
                                        (cxpObj.dias180Mas || 0);

                cxpObj.cxpTotal = totalAntiguedad;
                point.cxpDetail = cxpObj;
            }
        });

        let yearlyCostosGastoYTD = {};
        window.globalFinancialData.forEach(point => {
            const year = getSortYear(point);
            if (yearlyCostosGastoYTD[year] === undefined) {
                yearlyCostosGastoYTD[year] = 0;
            }
            
            if (point.cxpDetail && point.cxpDetail.costosGastoYtd > 0) {
            } else if (point.cxpDetail) {
                const costos = point.pnl ? (point.pnl.costos || 0) : 0;
                const opex = point.pnl ? (point.pnl.opex || 0) : 0;
                const capex = point.cashflowDetail ? (point.cashflowDetail.capex || 0) : 0;
                const costosGastoMensual = Math.abs(costos) + Math.abs(opex) + Math.abs(capex);
                yearlyCostosGastoYTD[year] += costosGastoMensual;
                point.cxpDetail.costosGastoYtd = yearlyCostosGastoYTD[year];
            }
            
            if (point.cxpDetail) {
                const elapsed_months = (point.sortDate.getMonth() === 11 && getSortYear(point) === 2025) ? 12 : (point.sortDate.getMonth() + 1);
                const days = elapsed_months * 30.4;
                
                if (point.cxpDetail.costosGastoYtd > 0) {
                    point.cxpDetail.dpo = Math.round((point.cxpDetail.cxpTotal / point.cxpDetail.costosGastoYtd) * days);
                } else {
                    point.cxpDetail.dpo = 0;
                }
            }
        });

        try {
            const db = await new Promise(r => { const req = indexedDB.open('PlanetaAzulDB', 3); req.onsuccess = () => r(req.result); });
            const tx = db.transaction('finance_cache', 'readwrite');
            tx.objectStore('finance_cache').put({ data: window.globalFinancialData, timestamp: Date.now() }, 'MASTER_FINANCE_KEY');
        } catch(e) {
            console.error("Error saving merged CXP to indexedDB:", e);
        }
    };

    window.processCxpFile = async function`);

fs.writeFileSync('main.js', content);
console.log('Fixed cxp!');
