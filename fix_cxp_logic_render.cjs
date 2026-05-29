const fs = require('fs');
let mainJs = fs.readFileSync('main.js', 'utf8');

const startStr = "function renderCxpView(";
const startIdx = mainJs.indexOf(startStr);

if (startIdx === -1) {
    console.log("Could not find start");
    process.exit(1);
}

const endStr = "function renderDeudaView(";
const endIdx = mainJs.indexOf(endStr);

if (endIdx === -1) {
    console.log("Could not find end");
    process.exit(1);
}

const newLogic = `window.renderCxpView = function(overrideData) {
    const data = overrideData || window.cxpStandaloneData;
    const headerEl = document.getElementById('cxpHeader');
    const bodyEl = document.getElementById('cxpBody');
    const periodLabel = document.getElementById('cxpPeriodLabel');
    if (!headerEl || !bodyEl || !data || !data.labels) return;

    if (periodLabel) {
        let yr = "20XX";
        if (data.periods && data.periods.length > 0) {
            let lastP = data.periods[data.periods.length - 1];
            yr = String(lastP).split('/')[1] || "20XX";
        }
        periodLabel.innerText = \`Análisis de Cuentas Por Pagar \${yr}\`;
    }

    let headerHTML = '<tr><th>Concepto</th>';
    data.labels.forEach((lbl) => {
        headerHTML += \`<th style="text-align:right;">\${lbl}</th>\`;
    });
    headerHTML += '</tr>';
    headerEl.innerHTML = headerHTML;
    
    const formatCurrencyStr = (v, minDec = 2, maxDec = 2) => {
        if (v === undefined || v === null) return '-';
        if (typeof v !== 'number') return v;
        return v.toLocaleString('en-US', { minimumFractionDigits: minDec, maximumFractionDigits: maxDec });
    };

    const formatInt = (v) => v !== undefined && v !== null ? Math.round(v).toLocaleString('en-US') : '-';

    let bodyHTML = '';

    const addRow = (label, values, isTotal = false) => {
        const rowStyle = isTotal ? 'font-weight:700; background:rgba(0,0,0,0.02);' : '';
        let h = \`<tr style="\${rowStyle}">\` 
            + \`<td>\${label}</td>\`;
        
        values.forEach((v) => {
            const valNum = parseFloat(v);
            const isNegative = !isNaN(valNum) && valNum < -0.009;
            const textCls = isNegative ? 'negative-val' : '';
            h += \`<td style="text-align:right;" class="\${textCls}">\${formatCurrencyStr(v)}</td>\`;
        });
        h += '</tr>';
        return h;
    };

    // Resumen General
    bodyHTML += addRow('Balance General', data.BalanceGeneral, true);
    
    // Aging
    bodyHTML += \`<tr><td colspan="\${data.labels.length + 1}" style="font-weight:700; background:rgba(0,0,0,0.04); font-size: 0.85rem; text-transform: uppercase;">Aging</td></tr>\`;
    bodyHTML += addRow('CXP', data.CXP);
    bodyHTML += addRow('Proveedores provisión sin fact', data.Provisionales);
    bodyHTML += addRow('Corriente (Saldo No Vencido)', data.Corriente);
    bodyHTML += addRow('0 a 30', data.Aging['0_30']);
    bodyHTML += addRow('31 a 60', data.Aging['31_60']);
    bodyHTML += addRow('61 a 90', data.Aging['61_90']);
    bodyHTML += addRow('91 a 120', data.Aging['91_120']);
    bodyHTML += addRow('121 a 150', data.Aging['121_150']);
    bodyHTML += addRow('151 a 180', data.Aging['151_180']);
    bodyHTML += addRow('> 180', data.Aging['180Mas']);

    // Proveedores Top 14
    bodyHTML += \`<tr style="height:20px"><td colspan="\${data.labels.length + 1}"></td></tr>\`;
    
    let provLine = '<tr><td style="font-weight:700; background:rgba(0,0,0,0.04); font-size: 0.85rem; text-transform: uppercase;">Saldos de Top Proveedores</td>';
    data.labels.forEach((lbl) => provLine += \`<td style="text-align:right; font-weight:700; background:rgba(0,0,0,0.04);">\${lbl}</td>\`);
    provLine += '</tr>';
    bodyHTML += provLine;

    data.Top14Names.forEach((name) => {
        bodyHTML += addRow(name, data.Top14Saldos[name] || [0,0,0,0,0]);
    });

    bodyHTML += addRow('Otros Proveedores', data.OtrosProveedores);
    bodyHTML += addRow('Total', data.Total, true);
    
    bodyHTML += \`<tr style="height:20px"><td colspan="\${data.labels.length + 1}"></td></tr>\`;
    
    // Costos YTD
    let costosRow = '<tr><td>Costos + Gasto (Opex+Capex) YTD</td>';
    data.CostosYTD.forEach((v) => {
        // formatCostos can be up to thousands, we can use 2 decimals
        costosRow += \`<td style="text-align:right;">\${formatCurrencyStr(v)}</td>\`
    });
    costosRow += '</tr>';
    bodyHTML += costosRow;
    
    // DPO
    let dpoRow = '<tr><td>DPO</td>';
    data.DPO.forEach((v) => {
        dpoRow += \`<td style="text-align:right;">\${formatInt(v)}</td>\`
    });
    dpoRow += '</tr>';
    bodyHTML += dpoRow;

    bodyEl.innerHTML = bodyHTML;
}

// Ensure the old function name works due to previous references if any
function renderCxpView(overrideData) {
    if (window.renderCxpView) {
        window.renderCxpView(overrideData);
    }
}

`;

const res = mainJs.slice(0, startIdx) + newLogic + '\n\n' + mainJs.slice(endIdx);
fs.writeFileSync('main.js', res);
console.log("Replaced renderCxpView");
