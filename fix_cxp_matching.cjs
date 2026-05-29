const fs = require('fs');
let mainJs = fs.readFileSync('main.js', 'utf8');

const sIdx = mainJs.indexOf('if (window.globalFinancialData && data.periods && data.periods.length > 0) {');
const eIdx = mainJs.indexOf('if (periodLabel) {', sIdx);

const newLogic = `if (window.globalFinancialData && data.periods && data.periods.length > 0) {
        data.CostosYTD = data.periods.map((p, i) => {
            const parts = String(p).split('/');
            const m = parseInt(parts[0], 10);
            const y = parseInt(parts[1], 10);
            
            const gItem = window.globalFinancialData.find(d => {
                if (d.sortDate) {
                    const dt = new Date(d.sortDate);
                    return dt.getMonth() + 1 === m && dt.getFullYear() === y;
                }
                return false;
            });
            
            if (gItem && gItem.wcFullRows) {
                const r = gItem.wcFullRows.find(rw => {
                    const c = String(rw.concept).toLowerCase().trim();
                    return (c.includes("capex") && c.includes("opex") && c.includes("costo"));
                });
                if (r && r.values) {
                    // Try exact match
                    if (r.values[gItem.date] !== undefined && r.values[gItem.date] !== null && r.values[gItem.date] !== "") return r.values[gItem.date];
                    
                    // Fallback to fuzzy match keys in r.values
                    for (let key in r.values) {
                        if (String(key).toLowerCase().trim() === String(gItem.date).toLowerCase().trim()) {
                            if (Math.abs(Number(r.values[key])) > 0) return r.values[key];
                        }
                    }
                    
                    // If all fails, fall back to matching month and year string in key
                    const shortMonths = ['ene','feb','mar','abr','may','jun','jul','ago','sep','oct','nov','dic'];
                    const tgtStr = shortMonths[m-1] + "-" + String(y).slice(-2);
                    for (let key in r.values) {
                        const kl = String(key).toLowerCase();
                        if (kl.includes(shortMonths[m-1]) && kl.includes(String(y).slice(-2))) {
                            if (Math.abs(Number(r.values[key])) > 0) return r.values[key];
                        }
                    }
                }
            }
            return data.CostosYTD[i]; // fallback to existing
        });
        
        data.DPO = data.periods.map((p, i) => {
            const parts = String(p).split('/');
            const m = parseInt(parts[0], 10);
            const y = parseInt(parts[1], 10);
            
            const gItem = window.globalFinancialData.find(d => {
                if (d.sortDate) {
                    const dt = new Date(d.sortDate);
                    return dt.getMonth() + 1 === m && dt.getFullYear() === y;
                }
                return false;
            });
            
            if (gItem && gItem.wcFullRows) {
                const r = gItem.wcFullRows.find(rw => {
                    const c = String(rw.concept).toLowerCase().trim();
                    return c === "dpo" || c.includes("dpo");
                });
                if (r && r.values) {
                    if (r.values[gItem.date] !== undefined && r.values[gItem.date] !== null && r.values[gItem.date] !== "") return r.values[gItem.date];
                    
                    for (let key in r.values) {
                         if (String(key).toLowerCase().trim() === String(gItem.date).toLowerCase().trim()) {
                            if (Math.abs(Number(r.values[key])) > 0) return r.values[key];
                        }
                    }
                    
                    const shortMonths = ['ene','feb','mar','abr','may','jun','jul','ago','sep','oct','nov','dic'];
                    for (let key in r.values) {
                        const kl = String(key).toLowerCase();
                        if (kl.includes(shortMonths[m-1]) && kl.includes(String(y).slice(-2))) {
                            if (Math.abs(Number(r.values[key])) > 0) return r.values[key];
                        }
                    }
                }
            }
            return data.DPO[i]; // fallback to existing
        });
    }

    `;
    
mainJs = mainJs.slice(0, sIdx) + newLogic + mainJs.slice(eIdx);
fs.writeFileSync('main.js', mainJs);
console.log("Updated fuzzy matching for Capex and DPO");
