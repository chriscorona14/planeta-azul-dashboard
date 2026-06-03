export let costoUnitarioData = null;
export let lastParsedWorkbook = null;

export function hasCostoUnitarioData() {
    return costoUnitarioData !== null;
}

const MONTH_COLS_REAL = ['P','Q','R','S','T','U','V','W','X','Y','Z','AA'];
const MONTH_COLS_PPTO = ['AC','AD','AE','AF','AG','AH','AI','AJ','AK','AL','AM','AN'];

function parseCostoUnitario(sheet) {
    if (!sheet) return null;
    
    const data = {
        botella: [],
        botellon: []
    };

    function parseBlock(startRow, endRow) {
        let blockRows = [];
        for (let r = startRow; r <= endRow; r++) {
            // Concept name is in column C
            let conceptCell = sheet['C' + r];
            if (!conceptCell || !conceptCell.v) continue;
            let concept = String(conceptCell.v).trim();
            
            // Just take all rows that have a concept. We might want to format them appropriately based on content.
            // Also rows usually alternate between DOP, Costo Unitario, %, etc. We extract the exact format dynamically based on B column (Metric type) and A column?
            // Actually, for the target view: we want Concept, Valor (DOP) and Costo Unit.
            // In the excel file, the concepts are repeated. e.g Row 10 is Agua Tratada (DOP), Row 11 is Costo Unitario (Agua Tratada). 
            // Looking at the csv:
            // "DOP",,"Agua Tratada",...
            // "Costo Unitario",,"Agua Tratada",...
            // Or the row has 'DOP' in A, 'Costo Unitario' in A. 
            // We just store all rows and map them later.
            
            let rowDict = {
                concept: concept,
                colA: sheet['A' + r] ? String(sheet['A' + r].v).trim() : '',
                colB: sheet['B' + r] ? String(sheet['B' + r].v).trim() : '',
                real: [],
                ppto: []
            };

            for (let i = 0; i < 12; i++) {
                let cellR = sheet[MONTH_COLS_REAL[i] + r];
                let cellP = sheet[MONTH_COLS_PPTO[i] + r];
                rowDict.real.push(cellR && cellR.t === 'n' ? cellR.v : (cellR ? parseFloat(cellR.v) || 0 : 0));
                rowDict.ppto.push(cellP && cellP.t === 'n' ? cellP.v : (cellP ? parseFloat(cellP.v) || 0 : 0));
            }
            blockRows.push(rowDict);
        }
        return blockRows;
    }

    // Botella 0.5 LTS is rows 8 to 43
    data.botella = parseBlock(8, 43);
    // BOTELLON is rows 47 to 84
    data.botellon = parseBlock(47, 84);

    return data;
}

export function processManualFile(arrayBuffer) {
    return new Promise((resolve, reject) => {
        try {
            const data = new Uint8Array(arrayBuffer);
            const workbook = window.XLSX.read(data, { type: 'array' });
            processCostoUnitarioWorkbook(workbook);
            resolve(true);
        } catch (e) {
            console.error("Costo Unitario parse error", e);
            resolve(false);
        }
    });
}

export function processCostoUnitarioWorkbook(workbook) {
    if (!workbook) return;
    const sheetName = "Costos Unit V2";
    if (!workbook.Sheets[sheetName]) {
        console.warn("Hoja 'Costos Unit V2' no encontrada");
        return;
    }
    costoUnitarioData = parseCostoUnitario(workbook.Sheets[sheetName]);
    lastParsedWorkbook = workbook;
}

export function renderCostoUnitario(monthIndex, prodType) {
    if (!costoUnitarioData) return;

    // Determine target Month string
    const monthsStr = ["ENE", "FEB", "MAR", "ABR", "MAY", "JUN", "JUL", "AGO", "SEP", "OCT", "NOV", "DIC"];
    
    // Regla: Real closed until April (index 3). 
    // We check the first few rows of real data.
    
    let block = costoUnitarioData[prodType];
    
    let tcRow = block.find(r => r.concept.includes('Total Costo') && r.colA === 'DOP');
    let monthIsReal = [];
    for (let m = 0; m <= monthIndex; m++) {
        let isMReal = m <= 3;
        if (m > 3 && tcRow && tcRow.real[m] > 0) {
            isMReal = true;
        }
        monthIsReal.push(isMReal);
    }
    
    // Update the tag label (using the latest selected month as "current")
    let currentIsReal = monthIsReal[monthIndex];
    // No actualizamos costoUnitarioDateType porque fue removido de la vista

    const thead = document.getElementById("costo-unitario-thead");
    if (thead) {
        thead.innerHTML = "";
        let thr = document.createElement("tr");
        let thConcept = document.createElement("th");
        thConcept.style = "background:#0f172a; color:white; border:none; padding: 16px; min-width: 250px; text-align: left; font-weight: 700;";
        thConcept.innerText = "Concepto";
        thr.appendChild(thConcept);
        
        for (let m = 0; m <= monthIndex; m++) {
            let th = document.createElement("th");
            let tStr = monthIsReal[m] ? "REAL" : "PPTO";
            let bgCol = monthIsReal[m] ? "#1e293b" : "#f97316";
            th.style = `background:${bgCol}; color:white; border-bottom: 2px solid #38bdf8; padding: 16px; text-align: right; text-transform: uppercase;`;
            th.innerText = `${monthsStr[m]}-26\n(${tStr})`;
            thr.appendChild(th);
        }
        thead.appendChild(thr);
    }

    const tbody = document.getElementById("costo-unitario-tbody");
    if (!tbody) return;
    
    tbody.innerHTML = "";

    let renderedConcepts = new Set();
    
    let displayRows = [];

    for (let i = 0; i < block.length; i++) {
        let r_dop = block[i];
        if (renderedConcepts.has(i)) continue;

        let concept = r_dop.concept;
        
        const normConcept = concept.normalize("NFD").replace(/[\u0300-\u036f]/g, "").toUpperCase().trim();
        if (normConcept.includes('TOTAL COSTO CON DEPRECIACI') ||
            normConcept.includes('COSTO DE VENTAS (DOP) CON DEP') ||
            normConcept === 'COSTO DE VENTAS (DOP)') {
            continue;
        }
        
        if (prodType === 'botellon' && concept.toUpperCase().includes('APA BOTELLON 18.9 LTS (X1)')) continue;
        if (prodType === 'botella' && concept.toUpperCase().includes('AGUA PLANETA AZUL 16.9 OZ CLEAR (20/1)')) continue;
        
        let unitariosByMonth = [];
        
        // try to find Costo Unitario row
        let unitRowIndex = -1;
        for (let j = i+1; j < Math.min(i+5, block.length); j++) {
            if (block[j].concept === concept && block[j].colA === 'Costo Unitario') {
                unitRowIndex = j;
                break;
            }
        }

        if (unitRowIndex !== -1) {
            renderedConcepts.add(unitRowIndex);
        }
        renderedConcepts.add(i);

        let rowType = 'normal';
        let isTotal = concept.toLowerCase().includes('total');
        if (isTotal) rowType = 'total';

        for (let m = 0; m <= monthIndex; m++) {
            let val = "-";
            if (unitRowIndex !== -1) {
                val = monthIsReal[m] ? block[unitRowIndex].real[m] : block[unitRowIndex].ppto[m];
            } else {
                val = monthIsReal[m] ? r_dop.real[m] : r_dop.ppto[m];
            }
            unitariosByMonth.push(val);
        }
        
        let isPct = false;
        if (unitRowIndex !== -1) {
            isPct = block[unitRowIndex].colA === '%' || block[unitRowIndex].concept === 'Margen del Costo Bruto';
        } else {
            isPct = r_dop.colA === '%' || r_dop.concept === 'Margen del Costo Bruto';
        }

        displayRows.push({
            concept: concept,
            valores: unitariosByMonth,
            type: rowType,
            isPct: isPct
        });
    }

    const fmtNum = (n, dec=2) => {
        if (n === "-" || Number.isNaN(n) || n === null || n === undefined) return "-";
        return Number(n).toLocaleString('en-US', { minimumFractionDigits: dec, maximumFractionDigits: dec });
    };

    displayRows.forEach(dr => {
        let tr = document.createElement("tr");

        let styleLabel = "padding: 14px 16px; border-bottom: 1px solid #f1f5f9; color: var(--text-primary); font-size: 0.9rem;";
        
        if (dr.type === 'total') {
            styleLabel += " font-weight: 800; background: #f8fafc;";
        }
        if (dr.concept === 'Costo Unitario') {
            styleLabel += " font-weight: 800; color: #0284c7;";
        }

        tr.innerHTML = `<td style="${styleLabel}">${dr.concept}</td>`;
        
        for (let m = 0; m <= monthIndex; m++) {
            let td = document.createElement("td");
            let styleUnit = "padding: 14px 16px; border-bottom: 1px solid #f1f5f9; color: var(--sidebar); font-size: 0.95rem; text-align: right; font-weight: 600; font-variant-numeric: tabular-nums;";
            
            if (dr.type === 'total') {
                styleUnit += " font-weight: 800; background: #f8fafc;";
            }
            if (dr.concept === 'Costo Unitario') {
                styleUnit += " color: #0284c7;";
            }
            
            let v = dr.valores[m];
            let unitStr = dr.isPct ? fmtNum(v, 2) + "%" : fmtNum(v, typeof v === 'number' && v < 1 && v > 0 ? 4 : 2);
            td.style = styleUnit;
            td.innerText = unitStr;
            tr.appendChild(td);
        }
        
        tbody.appendChild(tr);
    });
}
