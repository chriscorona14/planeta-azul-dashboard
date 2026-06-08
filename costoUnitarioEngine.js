export let costoUnitarioData = null;
export let lastParsedWorkbook = null;

export function hasCostoUnitarioData() {
    return costoUnitarioData !== null;
}

const MONTH_COLS_REAL25 = ['D','E','F','G','H','I','J','K','L','M','N','O'];
const MONTH_COLS_REAL = ['P','Q','R','S','T','U','V','W','X','Y','Z','AA'];
const MONTH_COLS_PPTO = ['AC','AD','AE','AF','AG','AH','AI','AJ','AK','AL','AM','AN'];

function parseCostoUnitario(sheet) {
    if (!sheet) return null;
    
    // Dynamically find columns in Row 7 (or 6)
    let colsReal25 = [];
    let colsReal = [];
    let colsPpto = [];
    
    function getColName(n) {
        let name = '';
        while (n > 0) {
            let m = (n - 1) % 26;
            name = String.fromCharCode(65 + m) + name;
            n = Math.floor((n - m) / 26);
        }
        return name;
    }

    const monthPrefixes = ['ene', 'feb', 'mar', 'abr', 'may', 'jun', 'jul', 'ago', 'sep', 'oct', 'nov', 'dic'];
    
    // We expect 3 distinct blocks of 12 months.
    // Real 2025: usually early columns
    // Real 2026: middle columns
    // Ppto 2026: late columns
    // We will scan row 7 from A to AZ
    let foundMonths = [];
    for (let c = 1; c <= 52; c++) {
        let col = getColName(c);
        let cell = sheet[col + '7'] || sheet[col + '6'];
        if (cell && cell.v) {
            let val = String(cell.v).toLowerCase().trim();
            // check if it contains a month
            let mIdx = monthPrefixes.findIndex(mp => val.includes(mp));
            // Excel dates might be represented as numbers (e.g., 45000)
            if (mIdx === -1 && cell.w) {
                let wVal = String(cell.w).toLowerCase().trim();
                mIdx = monthPrefixes.findIndex(mp => wVal.includes(mp));
            }
            if (mIdx !== -1) {
                // Determine which group it belongs to based on text or column position
                let is25 = val.includes('25') || val.includes('2025') || (cell.w && cell.w.includes('25'));
                let isPpto = val.includes('ppto') || val.includes('presupuesto');
                foundMonths.push({ col: col, mIdx: mIdx, is25: is25, isPpto: isPpto, val: val });
            }
        }
    }
    
    // If we didn't find dynamic headers, fallback to constants
    const defaultReal25 = ['D','E','F','G','H','I','J','K','L','M','N','O'];
    const defaultReal = ['P','Q','R','S','T','U','V','W','X','Y','Z','AA'];
    const defaultPpto = ['AC','AD','AE','AF','AG','AH','AI','AJ','AK','AL','AM','AN'];
    
    for (let i = 0; i < 12; i++) {
        colsReal25.push(defaultReal25[i]);
        colsReal.push(defaultReal[i]);
        colsPpto.push(defaultPpto[i]);
    }
    
    // If we found a good sequence, override
    if (foundMonths.length >= 12) {
        // Group by sequence
        let real25Group = foundMonths.filter(m => m.is25);
        let pptoGroup = foundMonths.filter(m => m.isPpto);
        let realGroup = foundMonths.filter(m => !m.is25 && !m.isPpto);
        
        // If the grouping is robust, apply it
        if (realGroup.length >= 12) {
            for (let i=0; i<12; i++) {
                let match = realGroup.find(m => m.mIdx === i);
                if (match) colsReal[i] = match.col;
            }
        }
        if (pptoGroup.length >= 12) {
            for (let i=0; i<12; i++) {
                let match = pptoGroup.find(m => m.mIdx === i);
                if (match) colsPpto[i] = match.col;
            }
        }
        if (real25Group.length >= 12) {
            for (let i=0; i<12; i++) {
                let match = real25Group.find(m => m.mIdx === i);
                if (match) colsReal25[i] = match.col;
            }
        }
    }

    const data = {
        botella: [],
        botellon: []
    };

    function parseBlock(startRow, endRow) {
        let blockRows = [];
        for (let r = startRow; r <= endRow; r++) {
            let conceptCell = sheet['C' + r];
            if (!conceptCell || !conceptCell.v) continue;
            let concept = String(conceptCell.v).trim();
            
            let rowDict = {
                concept: concept,
                colA: sheet['A' + r] ? String(sheet['A' + r].v).trim() : '',
                colB: sheet['B' + r] ? String(sheet['B' + r].v).trim() : '',
                real25: [],
                real: [],
                ppto: []
            };

            for (let i = 0; i < 12; i++) {
                let cellR25 = sheet[colsReal25[i] + r];
                let cellR = sheet[colsReal[i] + r];
                let cellP = sheet[colsPpto[i] + r];
                rowDict.real25.push(cellR25 && cellR25.t === 'n' ? cellR25.v : (cellR25 ? parseFloat(cellR25.v) || 0 : 0));
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
    return new Promise(async (resolve, reject) => {
        try {
            const data = new Uint8Array(arrayBuffer);
            const workbook = window.XLSX.read(data, { type: 'array' });
            await processCostoUnitarioWorkbook(workbook);
            resolve(true);
        } catch (e) {
            console.error("Costo Unitario parse error", e);
            resolve(false);
        }
    });
}

// IndexedDB Helpers
function openDB() {
  return new Promise((resolve, reject) => {
    const req = indexedDB.open('PlanetaAzulDB', 7);
    req.onupgradeneeded = (e) => {
      if (!e.target.result.objectStoreNames.contains('finance_cache')) {
        e.target.result.createObjectStore('finance_cache');
      }
    };
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error);
  });
}

function dbPut(db, key, value) {
  return new Promise((resolve, reject) => {
    const tx = db.transaction('finance_cache', 'readwrite');
    tx.objectStore('finance_cache').put(value, key);
    tx.oncomplete = resolve;
    tx.onerror = reject;
  });
}

function dbGet(db, key) {
  return new Promise((resolve) => {
    const req = db.transaction('finance_cache', 'readonly').objectStore('finance_cache').get(key);
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => resolve(null);
  });
}

export async function processCostoUnitarioWorkbook(workbook) {
    if (!workbook) return;
    const sheetName = "Costos Unit V2";
    if (!workbook.Sheets[sheetName]) {
        console.warn("Hoja 'Costos Unit V2' no encontrada");
        return;
    }
    costoUnitarioData = parseCostoUnitario(workbook.Sheets[sheetName]);
    lastParsedWorkbook = workbook;
    
    try {
      const db = await openDB();
      await dbPut(db, 'COSTO_UNITARIO_KEY', { data: costoUnitarioData, timestamp: Date.now() });
      console.log("💾 [costoUnitario] Datos persistidos con éxito en IndexedDB.");
    } catch (e) {
      console.warn('[costoUnitario] Fail to cache:', e);
    }
}

export async function loadCostoUnitarioCache() {
  try {
    const db = await openDB();
    const record = await dbGet(db, 'COSTO_UNITARIO_KEY');
    if (record && record.data) {
      costoUnitarioData = record.data;
      console.log("📂 [costoUnitario] Caché cargada de IndexedDB.");
      return true;
    }
  } catch (e) {
    console.warn('[costoUnitario] Fallo al cargar caché:', e);
  }
  return false;
}

export function renderCostoUnitario(monthIndex, prodType, vista = 'tendencia') {
    if (!costoUnitarioData) return;

    if (vista === 'resumen') {
        renderCostoUnitarioResumen(monthIndex, prodType);
    } else {
        renderCostoUnitarioTendencia(monthIndex, prodType);
    }
}

function renderCostoUnitarioTendencia(monthIndex, prodType) {
    const monthsStr = ["Ene", "Feb", "Mar", "Abr", "May", "Jun", "Jul", "Ago", "Sep", "Oct", "Nov", "Dic"];
    let block = costoUnitarioData[prodType];
    
    let tcRow = block.find(r => (r.concept.toLowerCase().includes('total costo') || r.concept.toUpperCase().includes('TOTAL COSTO')) && !r.concept.toLowerCase().includes('depreciacion') && !r.concept.toLowerCase().includes('con dep'));
    let volRow = block.find(r => r.concept.toLowerCase().includes('cantidad prod'));
    let monthIsReal = [];
    for (let m = 0; m <= monthIndex; m++) {
        let isMReal = true;
        
        let realVal = tcRow ? tcRow.real[m] : (volRow ? volRow.real[m] : 0);
        let pptoVal = tcRow ? tcRow.ppto[m] : (volRow ? volRow.ppto[m] : 0);
        
        // If there is no real data (value is 0 or missing) but there is PPTO data, fallback to PPTO.
        // We check the "Total Costo" or "Cantidad Producción" row as the source of truth for the month.
        if ((!realVal || realVal === 0) && (pptoVal && pptoVal !== 0)) {
           isMReal = false; 
        }
        monthIsReal.push(isMReal);
    }
    
    const table = document.getElementById("costo-unitario-table");
    if (table) table.style.width = '100%'; // Full width for Tendencia

    const thead = document.getElementById("costo-unitario-thead");
    if (thead) {
        thead.innerHTML = "";
        let thr = document.createElement("tr");
        let thConcept = document.createElement("th");
        thConcept.style = "background:#174c86; color:white; border: 1px solid #f8fafc; border-right: 1px solid white; padding: 14px 16px; min-width: 250px; text-align: left; font-weight: 700; font-size: 0.85rem; text-transform: uppercase;";
        thConcept.innerText = "Concepto";
        thr.appendChild(thConcept);
        
        for (let m = 0; m <= monthIndex; m++) {
            let th = document.createElement("th");
            let bgCol = monthIsReal[m] ? "#174c86" : "#f97316";
            th.style = `background:${bgCol}; color:white; border: 1px solid white; padding: 14px 8px; text-align: right; font-weight: 700; font-size: 0.85rem; text-transform: uppercase;`;
            th.innerText = monthIsReal[m] ? `${monthsStr[m].toUpperCase()} 2026` : `PPTO ${monthsStr[m].toUpperCase()} 2026`;
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
        const checkIsPct = (row) => {
            if (!row) return false;
            const colAVal = String(row.colA || '').trim();
            const colBVal = String(row.colB || '').trim();
            const cVal = String(row.concept || '').trim();
            return colAVal === '%' || colAVal.includes('%') || 
                   colBVal === '%' || colBVal.includes('%') || 
                   cVal === '%' || cVal.includes('%') || 
                   cVal.toLowerCase().includes('margen');
        };
        
        if (unitRowIndex !== -1) {
            isPct = checkIsPct(block[unitRowIndex]) || checkIsPct(r_dop);
        } else {
            isPct = checkIsPct(r_dop);
        }

        let isVol = concept.toLowerCase().includes('volumen') || concept.toLowerCase().includes('cantidad');

        displayRows.push({
            concept: concept,
            valores: unitariosByMonth,
            type: rowType,
            isPct: isPct,
            isVol: isVol
        });
    }

    const fmtNum = (n, isVol, isPct) => {
        if (n === "-" || Number.isNaN(n) || n === null || n === undefined) return "-";
        let val = Number(n);
        let str = '';
        if (isPct) {
            str = Math.abs(val).toLocaleString('en-US', { minimumFractionDigits: 1, maximumFractionDigits: 1 }) + "%";
        } else if (isVol) {
            str = Math.abs(val).toLocaleString('en-US', { minimumFractionDigits: 0, maximumFractionDigits: 0 });
        } else {
            str = Math.abs(val).toLocaleString('en-US', { minimumFractionDigits: 1, maximumFractionDigits: 1 });
        }
        return val < 0 ? `(${str})` : str;
    };

    displayRows.forEach(dr => {
        let tr = document.createElement("tr");

        let styleLabel = "padding: 14px 16px; border-bottom: 1px solid #f1f5f9; color: var(--text-primary); font-size: 0.95rem;";
        
        if (dr.type === 'total') {
            styleLabel += " font-weight: 800; background: #f8fafc;";
        }
        if (dr.concept === 'Costo Unitario') {
            styleLabel += " font-weight: 800; color: #0284c7;";
        }

        tr.innerHTML = `<td style="${styleLabel}">${dr.concept}</td>`;
        
        for (let m = 0; m <= monthIndex; m++) {
            let td = document.createElement("td");
            let styleUnit = "padding: 14px 8px; border-bottom: 1px solid #f1f5f9; color: var(--sidebar); font-size: 0.95rem; text-align: right; font-weight: 600; font-variant-numeric: tabular-nums;";
            
            if (dr.type === 'total') {
                styleUnit += " font-weight: 800; background: #f8fafc;";
            }
            if (dr.concept === 'Costo Unitario') {
                styleUnit += " color: #0284c7;";
            }
            
            let v = dr.valores[m];
            let vCorrected = v;
            if (dr.isPct && typeof v === 'number' && v < 1.1) {
                // Adjust raw fraction to percentage (since it seems raw is 0.xx)
                // Need to be careful here if not all % are fractions. 
                // In earlier version, we didn't multiply by 100 in Tendencia, we just did fmtNum(v, 2) + "%". 
                // Let's multiply if we are using the generic fmtNum.
                vCorrected = v * 100; 
            }
            let unitStr = fmtNum(vCorrected, dr.isVol, dr.isPct);
            td.style = styleUnit;
            td.innerText = unitStr;
            tr.appendChild(td);
        }
        
        tbody.appendChild(tr);
    });
}

function renderCostoUnitarioResumen(monthIndex, prodType) {
    const monthsStr = ["Ene", "Feb", "Mar", "Abr", "May", "Jun", "Jul", "Ago", "Sep", "Oct", "Nov", "Dic"];
    const mStr = monthsStr[monthIndex];
    let block = costoUnitarioData[prodType];

    const table = document.getElementById("costo-unitario-table");
    if (table) table.style.width = 'auto'; // Auto width for compact Resumen
    
    const thead = document.getElementById("costo-unitario-thead");
    if (thead) {
        thead.innerHTML = `
            <tr>
                <th rowspan="2" style="background:#174c86; color:white; border: 1px solid #f8fafc; border-right: 1px solid white; padding: 14px 16px; min-width: 250px; width: auto; text-align: left; font-weight: 700; font-size: 0.85rem; vertical-align: middle; text-transform: uppercase;">Concepto</th>
                <th rowspan="2" style="background:#174c86; color:white; border: 1px solid white; padding: 8px 10px; text-align: center; vertical-align: middle; width: 90px; font-weight: 700; font-size: 0.85rem; text-transform: uppercase;">${mStr.toUpperCase()} 2025</th>
                <th rowspan="2" style="background:#174c86; color:white; border: 1px solid white; padding: 8px 10px; text-align: center; vertical-align: middle; width: 90px; font-weight: 700; font-size: 0.85rem; text-transform: uppercase;">${mStr.toUpperCase()} 2026</th>
                <th rowspan="2" style="background:white; border:none; width: 4px; padding: 0;"></th>
                <th rowspan="2" style="background:#f97316; color:white; border: 1px solid white; padding: 8px 10px; text-align: center; vertical-align: middle; width: 100px; font-weight: 700; font-size: 0.85rem; text-transform: uppercase;">PPTO ${mStr.toUpperCase()} 2026</th>
                <th rowspan="2" style="background:white; border:none; width: 4px; padding: 0;"></th>
                <th colspan="2" style="background:black; color:white; border: 1px solid white; padding: 8px 10px; text-align: center; font-weight: 700; font-size: 0.85rem; text-transform: uppercase;">Var</th>
            </tr>
            <tr>
                <th style="background:#174c86; color:white; border: 1px solid white; padding: 6px 10px; text-align: center; font-size: 0.80rem; font-weight: 700; text-transform: uppercase;">vs 2025</th>
                <th style="background:#174c86; color:white; border: 1px solid white; padding: 6px 10px; text-align: center; font-size: 0.80rem; font-weight: 700; text-transform: uppercase;">vs PPTO</th>
            </tr>
        `;
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
        
        let unitRowIndex = -1;
        for (let j = i+1; j < Math.min(i+5, block.length); j++) {
            if (block[j].concept === concept && block[j].colA === 'Costo Unitario') {
                unitRowIndex = j;
                break;
            }
        }

        if (unitRowIndex !== -1) renderedConcepts.add(unitRowIndex);
        renderedConcepts.add(i);

        let rowType = concept.toLowerCase().includes('total') ? 'total' : 'normal';

        let val25, val26, valPpto;
        if (unitRowIndex !== -1 && concept !== 'Cantidad Producción por presentación' && !concept.toLowerCase().includes('volumen')) {
            val25 = block[unitRowIndex].real25[monthIndex];
            val26 = block[unitRowIndex].real[monthIndex];
            valPpto = block[unitRowIndex].ppto[monthIndex];
        } else {
            val25 = r_dop.real25[monthIndex];
            val26 = r_dop.real[monthIndex];
            valPpto = r_dop.ppto[monthIndex];
        }
        
        let isPct = false;
        const checkIsPct = (row) => {
            if (!row) return false;
            const colAVal = String(row.colA || '').trim();
            const colBVal = String(row.colB || '').trim();
            const cVal = String(row.concept || '').trim();
            return colAVal === '%' || colAVal.includes('%') || 
                   colBVal === '%' || colBVal.includes('%') || 
                   cVal === '%' || cVal.includes('%') || 
                   cVal.toLowerCase().includes('margen');
        };
        
        if (unitRowIndex !== -1) {
            isPct = checkIsPct(block[unitRowIndex]) || checkIsPct(r_dop);
        } else {
            isPct = checkIsPct(r_dop);
        }

        // Special override for volume or quantities that use large commas
        let isVol = concept.toLowerCase().includes('volumen') || concept.toLowerCase().includes('cantidad');

        displayRows.push({ concept, val25, val26, valPpto, type: rowType, isPct, isVol });
    }

    const fmtNum = (n, isVol, isPct) => {
        if (n === "-" || Number.isNaN(n) || n === null || n === undefined) return "-";
        let val = Number(n);
        let str = '';
        if (isPct) {
            str = Math.abs(val).toLocaleString('en-US', { minimumFractionDigits: 1, maximumFractionDigits: 1 }) + "%";
        } else if (isVol) {
            str = Math.abs(val).toLocaleString('en-US', { minimumFractionDigits: 0, maximumFractionDigits: 0 });
        } else {
            str = Math.abs(val).toLocaleString('en-US', { minimumFractionDigits: 1, maximumFractionDigits: 1 });
        }
        return val < 0 ? `(${str})` : str;
    };

    displayRows.forEach(dr => {
        let tr = document.createElement("tr");

        let styleLabel = "padding: 10px 16px; border-bottom: 1px solid #f1f5f9; color: var(--text-primary); font-size: 0.95rem;";
        if (dr.type === 'total') styleLabel += " font-weight: 800; background: #f8fafc;";
        else if (dr.concept.includes('Costo') || dr.concept === '%') styleLabel += " font-style: italic;";
        
        let val25Str = fmtNum(dr.val25, dr.isVol, dr.isPct);
        let val26Str = fmtNum(dr.val26, dr.isVol, dr.isPct);
        let pptoStr = fmtNum(dr.valPpto, dr.isVol, dr.isPct);
        
        let var25 = dr.val26 - dr.val25;
        let varPpto = dr.val26 - dr.valPpto;
        
        // For %, the variance is usually simple arithmetic difference. e.g. 55% - 50% = 5%
        // We calculate usually 55 - 50 = +5.0%
        if (dr.isPct && typeof dr.val26 === 'number' && typeof dr.valPpto === 'number' && dr.val26 < 1.1) {
            var25 = (dr.val26 - dr.val25) * 100;
            varPpto = (dr.val26 - dr.valPpto) * 100;
            // Also need to convert val25/val26/valPpto to *100 if they are raw ratios like 0.65
            dr.val25 = dr.val25 * 100;
            dr.val26 = dr.val26 * 100;
            dr.valPpto = dr.valPpto * 100;
            val25Str = fmtNum(dr.val25, dr.isVol, dr.isPct);
            val26Str = fmtNum(dr.val26, dr.isVol, dr.isPct);
            pptoStr = fmtNum(dr.valPpto, dr.isVol, dr.isPct);
        }

        let var25Str = fmtNum(var25, dr.isVol, dr.isPct);
        let varPptoStr = fmtNum(varPpto, dr.isVol, dr.isPct);

        let trStyle = dr.type === 'total' ? "background: #f8fafc; font-weight: bold;" : "";

        tr.innerHTML = `
            <td style="${styleLabel}">${dr.concept}</td>
            <td style="padding: 10px 16px; border-bottom: 1px solid #f1f5f9; text-align: right; font-size: 0.95rem; font-variant-numeric: tabular-nums; ${trStyle}">${val25Str}</td>
            <td style="padding: 10px 16px; border-bottom: 1px solid #f1f5f9; text-align: right; font-weight: bold; font-size: 0.95rem; font-variant-numeric: tabular-nums; ${trStyle}">${val26Str}</td>
            <td style="background:white; border-bottom: 1px solid #f1f5f9;"></td>
            <td style="padding: 10px 16px; border-bottom: 1px solid #f1f5f9; text-align: right; font-size: 0.95rem; font-variant-numeric: tabular-nums; ${trStyle}">${pptoStr}</td>
            <td style="background:white; border-bottom: 1px solid #f1f5f9;"></td>
            <td style="padding: 10px 16px; border-bottom: 1px solid #f1f5f9; text-align: right; font-size: 0.95rem; font-variant-numeric: tabular-nums; ${trStyle}">${var25Str}</td>
            <td style="padding: 10px 16px; border-bottom: 1px solid #f1f5f9; text-align: right; font-size: 0.95rem; font-variant-numeric: tabular-nums; ${trStyle}">${varPptoStr}</td>
        `;
        tbody.appendChild(tr);
    });
}
