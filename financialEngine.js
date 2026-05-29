/**
 * 🧠 MOTOR FINANCIERO CENTRAL (Versión Modular)
 */
import * as XLSX from 'xlsx';

const MONTH_ABBR_ES = ['ene', 'feb', 'mar', 'abr', 'may', 'jun',
                       'jul', 'ago', 'sep', 'oct', 'nov', 'dic'];

export function formatDateKey(dateObj) {
  if (!dateObj) return '';
  const d = (dateObj instanceof Date) ? dateObj : new Date(dateObj);
  if (isNaN(d.getTime())) return '';
  return `${MONTH_ABBR_ES[d.getMonth()]} ${d.getFullYear()}`;
}

export function extractConceptName(row) {
    if (!row) return "";
    let c0 = row[0] ? String(row[0]).trim() : "";
    let c1 = row[1] ? String(row[1]).trim() : "";
    let c2 = row[2] ? String(row[2]).trim() : "";

    const isCodeOrEmpty = (str) => {
        if (!str || str.toLowerCase() === 'x') return true;
        if (/^[\d\.\-\_ ]+$/.test(str)) return true; // Just numbers/codes
        const lower = str.toLowerCase();
        if (lower.includes('gmt') || lower.includes('00:00:00') || lower.includes('hora estándar')) return true;
        return false;
    };

    if (!isCodeOrEmpty(c0)) return c0;
    if (!isCodeOrEmpty(c1)) return c1;
    if (!isCodeOrEmpty(c2)) return c2;
    
    if (c1 && !isCodeOrEmpty(c1)) return c1;
    if (c0 && !isCodeOrEmpty(c0)) return c0;
    return "";
}

export function normalizeText(text) {
    if (!text) return "";
    // Normalizar: minúsculas, sin acentos y remover puntuación común de cabeceras/paréntesis
    return text.toString()
        .toLowerCase()
        .normalize("NFD")
        .replace(/[\u0300-\u036f]/g, "")
        .replace(/[.:;()]/g, " ")
        .replace(/\s+/g, " ")
        .trim();
}

export function formatSegmentName(name) {
    if (!name) return name;
    let n = typeof name === 'string' ? name : String(name);
    return n
        .replace(/\bEVP\b/g, "EVP (Botellas)")
        .replace(/\bBT5\b/g, "BT5 (Botellones)")
        .replace(/\b(?:BON|P6)\b/g, "BON (Zumos)")
        .replace(/\bfx\b/gi, "Tasa de cierre USD");
}

export function cleanNumber(val) {
    if (typeof val === 'number') return val;
    if (val === null || val === undefined) return 0;
    
    let cleaned = val.toString().trim().replace(/\u00A0/g, ' '); // Handle non-breaking spaces
    if (!cleaned) return 0;

    // Si no es un número puro, intentamos extraer la parte numérica (ej: "Tasa 58.50" -> 58.50)
    if (isNaN(cleaned.replace(/,/g, ''))) {
        const match = cleaned.match(/-?[\d,.]+/);
        if (match) cleaned = match[0];
    }

    // Handle (1,234.56) notation for negative numbers
    let isNegative = false;
    if (cleaned.startsWith('(') && cleaned.endsWith(')')) {
        isNegative = true;
        cleaned = cleaned.substring(1, cleaned.length - 1);
    } else if (cleaned.startsWith('-')) {
        isNegative = true;
        cleaned = cleaned.substring(1);
    }

    // Advanced thousands separator handling
    if (cleaned.includes(',') && cleaned.includes('.')) {
        if (cleaned.lastIndexOf(',') > cleaned.lastIndexOf('.')) {
            cleaned = cleaned.replace(/\./g, '').replace(',', '.');
        } else {
            cleaned = cleaned.replace(/,/g, '');
        }
    } else if (cleaned.includes(',')) {
        const parts = cleaned.split(',');
        if (parts[parts.length - 1].length === 3) {
            cleaned = cleaned.replace(/,/g, '');
        } else {
            cleaned = cleaned.replace(',', '.');
        }
    }

    cleaned = cleaned.replace(/[$\s%]/g, '');
    let num = parseFloat(cleaned);
    if (isNaN(num)) return 0;
    return isNegative ? -num : num;
}

export let appConfig = { isRawData: false };

export const formatCurrency = (val) => {
    if (val === 0 || val === null || val === undefined) return "$0.0M";
    
    // Asumimos que los valores ya vienen correctamente escalados a millones desde el motor.
    const formatted = Math.abs(val).toLocaleString('en-US', {
        minimumFractionDigits: 1,
        maximumFractionDigits: 1
    });

    return `${val < 0 ? '-' : ''}$${formatted}M`;
};

export const formatCurrencyUnits = (val) => {
    if (val === 0 || val === null || val === undefined) return "$0.00";
    const formatted = Math.abs(val).toLocaleString('en-US', {
        minimumFractionDigits: 2,
        maximumFractionDigits: 2
    });
    return `${val < 0 ? '-' : ''}$${formatted}`;
};

export const formatRawCurrency = (val) => {
    if (val === 0 || val === null || val === undefined) return "0.0";
    
    const formatted = Math.abs(val).toLocaleString('en-US', {
        minimumFractionDigits: 1,
        maximumFractionDigits: 1
    });

    return `${val < 0 ? '-' : ''}${formatted}`;
};

export function formatPercent(val) { 
    return (val * 100).toFixed(1) + '%'; 
}

export function financialEngine(workbook) {
    appConfig.isRawData = false;
    let sheets = {};
    workbook.SheetNames.forEach(name => {
        const normName = normalizeText(name);
        sheets[normName] = XLSX.utils.sheet_to_json(workbook.Sheets[name], { header: 1 });
    });

    const sheetKeys = Object.keys(sheets);
    
    // Prioritize "PA" or "Seguimiento" as per user context
    const estadosKey = sheetKeys.find(s => /estados financieros|estado de resultados/i.test(s) && !/ppto/i.test(s));
    const pnlKey = sheetKeys.find(s => /^p\s*[&e]\s*l\b/i.test(s) && !/ppto/i.test(s)) ||
                   sheetKeys.find(s => /mensual/i.test(s) && /p.*l|resultado|ganancia/i.test(s) && !/ppto/i.test(s)) ||
                   sheetKeys.find(s => /seguimiento|gerencial/i.test(s) && !/ppto/i.test(s)) ||
                   sheetKeys.find(s => /^pa\s/i.test(s) && !/ppto/i.test(s)) ||
                   sheetKeys.find(s => /resultado|income|ganancia/i.test(s) && !/ppto/i.test(s)) ||
                   sheetKeys[0];
    const balanceKey = sheetKeys.find(s => s.toLowerCase().includes("balance sheet mdop") && !/ppto/i.test(s)) || 
                       sheetKeys.find(s => /balance|situacion|estado/i.test(s) && !/p&l|resultado/i.test(s) && !/ppto/i.test(s));
    const cashflowKey = sheetKeys.find(s => /cash|flujo/i.test(s) && !/ppto/i.test(s));
    const wcKey = sheetKeys.find(s => /working|capital|wc/i.test(s) && !/ppto/i.test(s));

    const pptoPnlKey = sheetKeys.find(s => /ppto/i.test(s) && /l/i.test(s));
    const pptoBalanceKey = sheetKeys.find(s => /ppto/i.test(s) && /balance/i.test(s));
    const pptoCashflowKey = sheetKeys.find(s => /ppto/i.test(s) && /cash/i.test(s));

    let deudaSheetKeys = sheetKeys.filter(s => s.toLowerCase().includes("deuda"));
    let deudaSheet = null;
    if (deudaSheetKeys.length > 0) {
        deudaSheet = sheets[deudaSheetKeys[0]];
    }

    let presentacionSheetKey = sheetKeys.find(s => {
        const lower = s.toLowerCase();
        return lower.includes("presentaci") || lower.includes("presentación") || lower.includes("presentation");
    });
    let presentacionSheet = presentacionSheetKey ? sheets[presentacionSheetKey] : null;

    const cxpKey = sheetKeys.find(s => /cxp|cuentas por pagar|aging|antiguedad|proveedores/i.test(s));
    const cxpSheet = cxpKey ? sheets[cxpKey] : null;

    if (pnlKey && sheets[pnlKey]) {
        const result = processFinancialStatements(sheets, pnlKey, balanceKey, cashflowKey, pptoPnlKey, pptoBalanceKey, pptoCashflowKey, wcKey, estadosKey, deudaSheet, presentacionSheet, cxpSheet);
        if (!result.error && result.data && result.data.length > 0) {
            result.modelType = "Reporte PA / Estados Financieros";
            return result;
        }
    }

    const tbKey = sheetKeys.find(k => k === "tb" || k.includes("trial balance") || k.includes("balanza") || k.includes("data"));
    const setupKey = sheetKeys.find(k => k === "setup" || k.includes("mapeo") || k.includes("config") || k.includes("mapping"));

    function isRealTB(sheet) {
        if (!sheet || sheet.length < 5) return false;
        let numericAccountCount = 0;
        const sample = sheet.slice(0, 50);
        sample.forEach(row => {
            if (!row) return;
            row.forEach(cell => {
                const val = String(cell).trim();
                if (/^\d{4,}/.test(val)) numericAccountCount++;
            });
        });
        return numericAccountCount > 3;
    }

    const tbSheet = tbKey ? sheets[tbKey] : null;
    const isTB = isRealTB(tbSheet);

    if (isTB && tbKey && setupKey) {
        const result = processTBSetup(sheets, tbKey, setupKey);
        if (!result.error) {
            result.modelType = "TB + Setup (Contable)";
            return result;
        }
    }

    const resultWide = processWide(sheets);
    if (!resultWide.error) {
        resultWide.modelType = "Wide Format (Reporte Gerencial)";
        return resultWide;
    }

    return { error: `No se detectó un modelo válido. El archivo debe contener una hoja llamada 'P&L', 'Balance' o 'TB'. Hojas encontradas: ${sheetKeys.join(", ")}` };
}

export function calculateYTD(dataArray, selectedIndex) {
    if (!dataArray || dataArray.length === 0 || selectedIndex < 0 || selectedIndex >= dataArray.length) {
        return { real: { ingresos: 0, ebitda: 0, cashflow: 0, utilidad: 0 }, ppto: { ingresos: 0, ebitda: 0, cashflow: 0, utilidad: 0 } };
    }

    const selectedData = dataArray[selectedIndex];
    const targetYear = selectedData.sortDate.getFullYear();

    let real = { ingresos: 0, ebitda: 0, cashflow: 0, utilidad: 0 };
    let ppto = { ingresos: 0, ebitda: 0, cashflow: 0, utilidad: 0 };

    for (let i = selectedIndex; i >= 0; i--) {
        const item = dataArray[i];
        if (item.sortDate.getFullYear() !== targetYear) break;

        real.ingresos += (item.kpis.ingresos || 0);
        real.ebitda += (item.kpis.ebitda || 0);
        real.cashflow += (item.kpis.cashflow || 0);
        real.utilidad += (item.kpis.utilidad || 0);

        if (item.ppto && item.ppto.kpis) {
            ppto.ingresos += (item.ppto.kpis.ingresos || 0);
            ppto.ebitda += (item.ppto.kpis.ebitda || 0);
            ppto.cashflow += (item.ppto.kpis.cashflow || 0);
            ppto.utilidad += (item.ppto.kpis.utilidad || 0);
        }
    }

    return { real, ppto };
}
// I will include the full logic to ensure it works as before

function findRowByKeywords(rows, keywords, targetRowIdxHint = null) {
    let bestRow = null;
    let maxScore = -1;
    
    // Normalización Total requerida por el usuario
    const normalizedKeywords = keywords.map(k => normalizeText(k).trim().toLowerCase());

    rows.forEach((row, idx) => {
        if (!row || row.length < 2) return;
        // Revisar más columnas (hasta la 10) por si el label está desplazado
        for (let i = 0; i < Math.min(row.length, 10); i++) {
            const cell = row[i];
            if (!cell) continue;
            // Normalización Total requerida por el usuario:
            // normalizeText ya quita espacios y baja a minúsculas, 
            // pero añadimos explícitamente trim() y toLowerCase() para máxima seguridad
            const label = normalizeText(cell).trim().toLowerCase();
            
            // Excluir líneas que son cálculos intermedios o utilidades antes de impuestos
            if (normalizedKeywords.includes("taxes") || normalizedKeywords.includes("impuestos")) {
                if (label.includes("antes") || label.includes("before") || label.includes("utilidad") || label.includes("operating") || label.includes("ebit") || label.includes("proyecci") || label.includes("ppto")) {
                    continue;
                }
            }
            
            let matchedKeyword = null;
            let matchIndex = -1;
            let isExact = false;

            for (let kIdx = 0; kIdx < normalizedKeywords.length; kIdx++) {
                const k = normalizedKeywords[kIdx];
                if (label === k) {
                    matchedKeyword = k;
                    matchIndex = kIdx;
                    isExact = true;
                    break;
                }
            }

            if (!matchedKeyword) {
                for (let kIdx = 0; kIdx < normalizedKeywords.length; kIdx++) {
                    const k = normalizedKeywords[kIdx];
                    if (k.length > 3 && label.includes(k)) {
                        // Prevent "otras cuentas por pagar" from matching just "cuentas por pagar"
                        if (k === "cuentas por pagar" && label.includes("otras")) continue;
                        matchedKeyword = k;
                        matchIndex = kIdx;
                        isExact = false;
                        break;
                    }
                }
            }

            if (matchedKeyword) {
                let numCount = 0;
                let potentialTotal = 0;
                for (let j = 1; j < row.length; j++) {
                    const val = cleanNumber(row[j]);
                    if (val !== 0) {
                        numCount++;
                        potentialTotal = Math.max(potentialTotal, Math.abs(val));
                    }
                }

                let score = numCount;
                
                // Prioridad según el índice del keyword en la lista (menor índice = mayor prioridad)
                // Esto asegura que palabras clave listadas primero tengan prioridad
                const keywordPriorityBonus = (normalizedKeywords.length - matchIndex) * 100;
                score += keywordPriorityBonus;

                // Prioridad extremadamente alta a coincidencias exactas con el keyword
                if (isExact) {
                    score += 500;
                }
                
                if (label.includes("total") || label.includes("sum") || label.includes("consolidado")) score += 15;
                if (label.includes("neto") || label.includes("final") || label.includes("ejercicio")) score += 20;

                // Si el usuario nos dio una pista de fila (ej: fila 61 en excel es idx 60)
                if (targetRowIdxHint !== null) {
                    if (Math.abs(idx - targetRowIdxHint) <= 5) score += 50; // Gran bono si está cerca de la fila 61
                }
                
                if (label.includes("%") || label.includes("var") || label.includes("crecimiento")) score -= 15;
                
                if (score > maxScore) {
                    maxScore = score;
                    bestRow = row;
                }
                break; 
            }
        }
    });
    return bestRow;
}

function detectSegments(rows, segmentKeywords) {
    const segments = {};
    rows.forEach(row => {
        if (!row) return;
        for (let i = 0; i < Math.min(row.length, 5); i++) {
            const cell = row[i];
            if (!cell) continue;
            const label = normalizeText(cell);
            
            segmentKeywords.forEach(seg => {
                const normSeg = seg.toLowerCase();
                const regex = new RegExp(`\\b${normSeg}\\b`, 'i');
                if (regex.test(label)) {
                    const finalSegName = (seg === "P6" || seg === "BON") ? "BON" : seg;
                    
                    if (!segments[finalSegName]) segments[finalSegName] = { ventasRows: [], costosRows: [] };
                    
                    const hasNumbers = row.some((c, idx) => idx > i && cleanNumber(c) !== 0);
                    if (hasNumbers) {
                        if (label.includes("costo") || label.includes("costos")) {
                            segments[finalSegName].costosRows.push(row);
                        } else if (label.includes("venta") || label.includes("ingreso") || !label.includes("total")) {
                            segments[finalSegName].ventasRows.push(row);
                        }
                    }
                }
            });
        }
    });
    return segments;
}

export const FINANCIAL_KEYWORDS = {
    ingresos: ["ventas", "ingresos", "revenue", "ventas netas", "total ingresos", "facturacion", "servicios", "productos", "ventas totales"],
    costos: ["costo de ventas", "costos directos", "cogs", "cost of sales", "total costos", "costos de operacion"],
    opex: ["gastos operativos", "opex", "gastos de administracion", "total gastos operativos", "gastos de venta", "otros gastos operativos", "ggadm", "gastos generales", "total gastos", "operativos"],
    ebitda: ["ebitda", "utilidad operativa", "operating income", "uafida", "utilidad antes de", "resultado operativo", "margen operativo", "utilidad de operacion"],
    utilidad: ["utilidad neta", "net income", "resultado del ejercicio", "utilidad perdida", "beneficio neto", "resultado neto", "utilidad del periodo", "ganancia neta", "ganancia del ejercicio", "utilidad neta ejercicio", "utilidad neta periodo", "ganancia perdida ejercicio", "resultado del periodo", "resultado"],
    cashflow: ["cash flow", "flujo de caja", "flujo neto", "disponibilidad", "caja final", "efectivo", "flujo de efectivo"],
    tasa_cambio: ["tasa de cambio", "fx rate", "tipo de cambio", "tasa bpd", "tasa promedio", "t.c", "tc", "tasa", "cambio", "exchange", "tasa proyectada"],
    // Nuevas Keywords para Hoja de Cash Flow
    cf_beginning: ["beginning cash balance", "efectivo inicial", "saldo inicial de efectivo", "caja inicial"],
    cf_operating: ["operating activities", "flujo de actividades de operacion", "actividades de operacion", "flujo de caja operativo"],
    cf_wc: ["change in working capital", "cambios en capital de trabajo", "variacion capital de trabajo", "working capital requirements", "working capital"],
    cf_cxc: ["aumento)/disminucion en cuentas por cobrar", "cuentas por cobrar", "cxc", "accounts receivable"],
    cf_inv: ["aumento)/disminucion en inventario", "inventario", "inventarios", "inventory"],
    cf_cxp: ["aumento/(disminucion) en cuentas por pagar", "cuentas por pagar", "cxp", "accounts payable"],
    cf_otros_activos: ["(aumento)/disminucion en otros activos", "otros activos"],
    cf_otros_activos_corrientes: ["(aumento)/disminucion en otros activos corrientes", "otros activos corrientes"],
    cf_activos_terceros: ["aumento/(disminucion) en activos en manos de terceros", "(aumento)/disminucion en activos en manos de terceros", "activos en manos de terceros", "activos terceros"],
    cf_pasivo_laboral: ["pasivo laboral", "pasivos laborales"],
    cf_otros_pasivos: ["(aumento)/disminucion en otros pasivos", "otros pasivos"],
    cf_otras_cxp: ["aumento/(disminucion) en otras cuentas por pagar", "otras cuentas por pagar", "otras cxp"],
    cf_otros_pasivos_corrientes: ["aumento/(disminucion) en otros pasivos corrientes", "otros pasivos corrientes"],
    cf_capex: ["capex", "inversiones de capital", "adquisicion de activos", "capital expenditures"],
    cf_financing: ["financing activities", "flujo de actividades de financiamiento", "actividades de financiamiento"],
    cf_net_debt: ["desembolsos de capital", "aumento deuda neta", "variacion de deuda", "financiamiento neto", "deuda bancaria", "net debt", "repayment of debt"],
    cf_change: ["change in cash", "cambio en efectivo", "variacion neta de efectivo"],
    cf_ending: ["ending cash balance", "efectivo final", "saldo final de efectivo", "caja final"],
    cf_below_ebitda: ["below ebitda"],
    cf_taxes: ["taxes"],
    cf_dividends: ["dividends", "dividendos", "shareholders activities", "accionistas"],
    cf_interest: ["gastos de interes", "intereses", "interest expense", "financial expenses", "interests earned"],
    cf_interest_earned: ["intereses ganados", "interests earned", "ingresos financieros", "interes ganado"],
    cf_interest_expense: ["gastos financieros", "gastos de interes", "interest expense", "financial expenses", "interes gasto"],
    cf_extraordinary: ["ingresos (gastos) extraordinarios", "gastos extraordinarios", "ingresos extraordinarios", "extraordinarios", "extraordinary items"],
    cf_diferencial_cambiario: ["diferencial cambiario", "diferencia en cambio"],
    cf_dso: ["dso"],
    cf_dpo: ["dpo"],
    cf_dio: ["dio"]
};

function processFinancialStatements(sheets, pnlKey, balanceKey, cashflowKey, pptoPnlKey = null, pptoBalanceKey = null, pptoCashflowKey = null, wcKey = null, estadosKey = null, deudaSheet = null, presentacionSheet = null, cxpSheet = null) {
    const pnlSheet = sheets[pnlKey];
    const balanceSheet = balanceKey ? sheets[balanceKey] : null;
    const cashflowSheet = cashflowKey ? sheets[cashflowKey] : null;
    const wcSheet = wcKey ? sheets[wcKey] : null;
    const estadosSheet = estadosKey ? sheets[estadosKey] : null;

    const pptoPnlSheet = pptoPnlKey ? sheets[pptoPnlKey] : null;
    const pptoBalanceSheet = pptoBalanceKey ? sheets[pptoBalanceKey] : null;
    const pptoCashflowSheet = pptoCashflowKey ? sheets[pptoCashflowKey] : null;

    // Detectar si el Balance o P&L están en millones (mDOP)
    let isBalanceInMillions = (balanceKey && (normalizeText(balanceKey).includes("mdop") || normalizeText(balanceKey).includes("millones") || normalizeText(balanceKey).includes("mrd$"))) ||
                              (pnlKey && (normalizeText(pnlKey).includes("mdop") || normalizeText(pnlKey).includes("millones") || normalizeText(pnlKey).includes("mrd$"))) ||
                              (cashflowKey && (normalizeText(cashflowKey).includes("mdop") || normalizeText(cashflowKey).includes("millones") || normalizeText(cashflowKey).includes("mrd$")));
    
    // If no scale is detected by text, check values
    if (!isBalanceInMillions) {
        const detectScale = (sheet) => {
            if (!sheet) return false;
            let foundText = false;
            let smallValuesCount = 0;
            let nonZeroCount = 0;

            // Revisamos hasta 200 filas para estar seguros de capturar todo el contexto
            for (let i = 0; i < Math.min(sheet.length, 200); i++) {
                if (sheet[i]) {
                    const rowStr = normalizeText(sheet[i].join(" "));
                    if (rowStr.includes("mdop") || rowStr.includes("millones") || rowStr.includes("mrd$") || rowStr.includes("cifras en")) foundText = true;
                    
                    sheet[i].forEach(cell => {
                        const n = cleanNumber(cell);
                        if (n !== 0 && !isNaN(n)) {
                            nonZeroCount++;
                            if (Math.abs(n) < 1000000) smallValuesCount++;
                        }
                    });
                }
            }
            return foundText || (nonZeroCount > 5 && (smallValuesCount / nonZeroCount) > 0.7);
        };
        isBalanceInMillions = detectScale(pnlSheet) || detectScale(balanceSheet) || detectScale(cashflowSheet);
    }
    
    appConfig.isRawData = !isBalanceInMillions;

    const getVal = (row, idx, isPnlSource = true) => {
        if (!row || idx === undefined || idx === null) return 0;
        let val = cleanNumber(row[idx]);
        
        let preventOffset = false;
        if (isPnlSource) {
            const concept = extractConceptName(row).toLowerCase();
            const isFX = (concept.includes("tasa") && !concept.includes("impacto")) || concept.includes("fx") || concept === "tc" || concept.includes("tipo de cambio") || concept.includes("dop") || concept.includes("exchange");
            const isRatio = concept.includes("%") || concept.includes("ratio");
            const isAlreadyMillions = concept.includes("musd") || concept.includes("mdop") || concept.includes("millones");
            
            preventOffset = true;

            if (!isFX && !isRatio && !isAlreadyMillions) {
                val = val / 1000000;
            }
        }

        // No column fallback offset used here to respect native zero values and single-period entries

        return val;
    };

    const getBalanceVal = (row, idx) => {
        if (!row) return 0;
        let val = cleanNumber(row[idx]);
        
        // No column fallback offset used here to respect native zero values and single-period entries

        const concept = normalizeText(extractConceptName(row));
        // Detectar si es un ratio (unitless) o moneda
        const isRatio = (concept.includes("ratio") || concept.includes("indice") || concept.includes("razon") ||
                         concept.includes("apalancamiento") || concept.includes("capacidad") || 
                         concept.includes("covenant") || concept.includes("corriente") ||
                         concept.includes("deuda neta/ebitda") || concept.includes(" x ") || concept.endsWith(" x")) && 
                         !concept.includes("cxp") && !concept.includes("cxc") && !concept.includes("pagar") && !concept.includes("cobrar");
        
        if (isRatio) return val; // No escalar ratios
        // Retornar valor nativo porque Balance y Cashflow ya vienen expresados en millones
        return val;
    };

    const detailedOpexKeywords = {
        admin: ["gastos administrativos", "gastos de administracion", "administracion"],
        mercadeo: ["gastos de mercadeo", "mercadeo", "publicidad", "marketing"],
        comercial: ["gastos de ventas (comercial)", "gastos de ventas", "comercial", "gastos comerciales"],
        logistica: ["gastos de logistica", "logistica", "gastos logisticos"]
    };

    const pnlRows = {
        ingresos: findRowByKeywords(pnlSheet, FINANCIAL_KEYWORDS.ingresos),
        costos: findRowByKeywords(pnlSheet, FINANCIAL_KEYWORDS.costos),
        opex: findRowByKeywords(pnlSheet, FINANCIAL_KEYWORDS.opex),
        ebitda: findRowByKeywords(pnlSheet, FINANCIAL_KEYWORDS.ebitda),
        utilidad: findRowByKeywords(pnlSheet, FINANCIAL_KEYWORDS.utilidad, 60), // Hint: Fila 61 (index 60)
        cashflow: findRowByKeywords(pnlSheet, FINANCIAL_KEYWORDS.cashflow),
        tasa_cambio: findRowByKeywords(pnlSheet, FINANCIAL_KEYWORDS.tasa_cambio),
        // Detalle de OPEX
        admin: findRowByKeywords(pnlSheet, detailedOpexKeywords.admin),
        mercadeo: findRowByKeywords(pnlSheet, detailedOpexKeywords.mercadeo),
        comercial: findRowByKeywords(pnlSheet, detailedOpexKeywords.comercial),
        logistica: findRowByKeywords(pnlSheet, detailedOpexKeywords.logistica),
        tasa_cambio: findRowByKeywords(pnlSheet, FINANCIAL_KEYWORDS.tasa_cambio)
    };

    const segmentKeywords = ["BT5", "EVP", "BON", "P6", "Otras Ventas", "Otros Ingresos"];
    const segmentRows = detectSegments(pnlSheet, segmentKeywords);

    if (!pnlRows.ingresos) return { error: "No se encontró la fila de 'Ingresos' en el P&L. Verifique que los nombres de las filas sean claros (ej: 'Ventas' o 'Ingresos')." };

    const balanceKeywords = {
        activos: ["total activos", "activos", "total activo", "activo total", "total de activos", "activos totales"],
        pasivos: ["total pasivos", "pasivos", "total pasivo", "pasivo total", "sumas iguales pasivo", "pasivos totales"],
        patrimonio: ["total patrimonio", "patrimonio", "capital", "total capital", "capital contable", "patrimonio neto", "total pasivo y patrimonio", "total pasivo y capital"]
    };

    const balanceRows = {
        activos: (balanceSheet ? findRowByKeywords(balanceSheet, balanceKeywords.activos) : null) || findRowByKeywords(pnlSheet, balanceKeywords.activos),
        pasivos: (balanceSheet ? findRowByKeywords(balanceSheet, balanceKeywords.pasivos) : null) || findRowByKeywords(pnlSheet, balanceKeywords.pasivos),
        patrimonio: (balanceSheet ? findRowByKeywords(balanceSheet, balanceKeywords.patrimonio) : null) || findRowByKeywords(pnlSheet, balanceKeywords.patrimonio),
        gananciaAcumulada: (balanceSheet ? findRowByKeywords(balanceSheet, ["ganancia acumulada", "utilidad acumulada", "ganancias acumuladas", "utilidades acumuladas"]) : null),
        utilidadesRetenidas: (balanceSheet ? findRowByKeywords(balanceSheet, ["utilidades retenidas", "utilidad retenida"]) : null),
        // Covenant and Leverage Indicators (using user exact row references in case keywords fail)
        covenantLean: balanceSheet ? (findRowByKeywords(balanceSheet, ["deuda neta bancaria", "ebitda", "4.0x"]) || balanceSheet[97]) : null,
        apalancamiento: balanceSheet ? (findRowByKeywords(balanceSheet, ["apalancamiento", "= 2.0x", "<=2.0x", "<= 2.0x"]) || balanceSheet[98]) : null,
        capacidadPago: balanceSheet ? (findRowByKeywords(balanceSheet, ["capacidad de pago"]) || balanceSheet[99]) : null,
        razonCorriente: balanceSheet ? (findRowByKeywords(balanceSheet, ["razon corriente", "current ratio", "liquidez"]) || balanceSheet[100]) : null,
        deudaBancariaNetaUSD: balanceSheet ? (findRowByKeywords(balanceSheet, ["deuda neta bancaria usd", "bancaria usd"]) || balanceSheet[95]) : null,
        cajaEfectivo: balanceSheet ? findRowByKeywords(balanceSheet, ["caja y banco", "efectivo y equivalente", "efectivo y caja", "caja"]) : null,
        deudaTotal: balanceSheet ? findRowByKeywords(balanceSheet, ["deuda financiera total", "deuda a corto y largo plazo"]) : null,
        cxp: balanceSheet ? findRowByKeywords(balanceSheet, ["cuentas por pagar", "cxp"]) : null
    };

    const cfRows = {
        beginning: cashflowSheet ? findRowByKeywords(cashflowSheet, FINANCIAL_KEYWORDS.cf_beginning) : null,
        operating: cashflowSheet ? findRowByKeywords(cashflowSheet, FINANCIAL_KEYWORDS.cf_operating) : null,
        wc: cashflowSheet ? findRowByKeywords(cashflowSheet, FINANCIAL_KEYWORDS.cf_wc) : null,
        cxc: cashflowSheet ? findRowByKeywords(cashflowSheet, FINANCIAL_KEYWORDS.cf_cxc) : null,
        inv: cashflowSheet ? findRowByKeywords(cashflowSheet, FINANCIAL_KEYWORDS.cf_inv) : null,
        cxp: cashflowSheet ? findRowByKeywords(cashflowSheet, FINANCIAL_KEYWORDS.cf_cxp) : null,
        otrosActivos: cashflowSheet ? findRowByKeywords(cashflowSheet, FINANCIAL_KEYWORDS.cf_otros_activos) : null,
        otrosActivosCorrientes: cashflowSheet ? findRowByKeywords(cashflowSheet, FINANCIAL_KEYWORDS.cf_otros_activos_corrientes) : null,
        activosTerceros: cashflowSheet ? findRowByKeywords(cashflowSheet, FINANCIAL_KEYWORDS.cf_activos_terceros) : null,
        pasivoLaboral: cashflowSheet ? findRowByKeywords(cashflowSheet, FINANCIAL_KEYWORDS.cf_pasivo_laboral) : null,
        otrosPasivos: cashflowSheet ? findRowByKeywords(cashflowSheet, FINANCIAL_KEYWORDS.cf_otros_pasivos) : null,
        otrasCxp: cashflowSheet ? findRowByKeywords(cashflowSheet, FINANCIAL_KEYWORDS.cf_otras_cxp) : null,
        otrosPasivosCorrientes: cashflowSheet ? findRowByKeywords(cashflowSheet, FINANCIAL_KEYWORDS.cf_otros_pasivos_corrientes) : null,
        capex: cashflowSheet ? findRowByKeywords(cashflowSheet, FINANCIAL_KEYWORDS.cf_capex) : null,
        financing: cashflowSheet ? findRowByKeywords(cashflowSheet, FINANCIAL_KEYWORDS.cf_financing) : null,
        netDebt: cashflowSheet ? findRowByKeywords(cashflowSheet, FINANCIAL_KEYWORDS.cf_net_debt, 47) : null,
        belowEbitda: cashflowSheet ? findRowByKeywords(cashflowSheet, FINANCIAL_KEYWORDS.cf_below_ebitda) : null,
        taxes: cashflowSheet ? findRowByKeywords(cashflowSheet, FINANCIAL_KEYWORDS.cf_taxes, 44) : null,
        dividends: cashflowSheet ? findRowByKeywords(cashflowSheet, FINANCIAL_KEYWORDS.cf_dividends) : null,
        interest: cashflowSheet ? findRowByKeywords(cashflowSheet, FINANCIAL_KEYWORDS.cf_interest) : null,
        interest_earned: cashflowSheet ? findRowByKeywords(cashflowSheet, FINANCIAL_KEYWORDS.cf_interest_earned) : null,
        interest_expense: cashflowSheet ? findRowByKeywords(cashflowSheet, FINANCIAL_KEYWORDS.cf_interest_expense) : null,
        extraordinary: cashflowSheet ? findRowByKeywords(cashflowSheet, FINANCIAL_KEYWORDS.cf_extraordinary) : null,
        diferencialCambiario: cashflowSheet ? findRowByKeywords(cashflowSheet, FINANCIAL_KEYWORDS.cf_diferencial_cambiario) : null,
        change: cashflowSheet ? findRowByKeywords(cashflowSheet, FINANCIAL_KEYWORDS.cf_change) : null,
        ending: cashflowSheet ? findRowByKeywords(cashflowSheet, FINANCIAL_KEYWORDS.cf_ending) : null,
        dso: cashflowSheet ? findRowByKeywords(cashflowSheet, FINANCIAL_KEYWORDS.cf_dso) : null,
        dpo: cashflowSheet ? findRowByKeywords(cashflowSheet, FINANCIAL_KEYWORDS.cf_dpo) : null,
        dio: cashflowSheet ? findRowByKeywords(cashflowSheet, FINANCIAL_KEYWORDS.cf_dio) : null
    };

    const wcRows = {
        wc: wcSheet ? findRowByKeywords(wcSheet, FINANCIAL_KEYWORDS.cf_wc) : null,
        cxc: wcSheet ? findRowByKeywords(wcSheet, FINANCIAL_KEYWORDS.cf_cxc) : null,
        inv: wcSheet ? findRowByKeywords(wcSheet, FINANCIAL_KEYWORDS.cf_inv) : null,
        cxp: wcSheet ? findRowByKeywords(wcSheet, FINANCIAL_KEYWORDS.cf_cxp) : null,
        dso: wcSheet ? findRowByKeywords(wcSheet, FINANCIAL_KEYWORDS.cf_dso) : null,
        dpo: wcSheet ? findRowByKeywords(wcSheet, FINANCIAL_KEYWORDS.cf_dpo) : null,
        dio: wcSheet ? findRowByKeywords(wcSheet, FINANCIAL_KEYWORDS.cf_dio) : null
    };

    const cxpRows = {
        provisionSinFactura: cxpSheet ? findRowByKeywords(cxpSheet, ["provision sin factura", "exencion itbis"]) : null,
        corriente: cxpSheet ? findRowByKeywords(cxpSheet, ["corriente", "al dia"]) : null,
        dias0_30: cxpSheet ? findRowByKeywords(cxpSheet, ["0 a 30", "0-30"]) : null,
        dias31_60: cxpSheet ? findRowByKeywords(cxpSheet, ["31 a 60", "31-60"]) : null,
        dias61_90: cxpSheet ? findRowByKeywords(cxpSheet, ["61 a 90", "61-90"]) : null,
        dias91_120: cxpSheet ? findRowByKeywords(cxpSheet, ["91 a 120", "91-120"]) : null,
        dias121_150: cxpSheet ? findRowByKeywords(cxpSheet, ["121 a 150", "121-150"]) : null,
        dias151_180: cxpSheet ? findRowByKeywords(cxpSheet, ["151 a 180", "151-180"]) : null,
        dias180Mas: cxpSheet ? findRowByKeywords(cxpSheet, ["> 180", "180+", "mas de 180"]) : null,
        
        // Suppliers (Top 14 + Otros)
        alplaHispaniola: cxpSheet ? findRowByKeywords(cxpSheet, ["alpla hispaniola"]) : null,
        polyplas: cxpSheet ? findRowByKeywords(cxpSheet, ["polyplas"]) : null,
        grupoRojas: cxpSheet ? findRowByKeywords(cxpSheet, ["grupo rojas"]) : null,
        raviCaribe: cxpSheet ? findRowByKeywords(cxpSheet, ["ravi caribe"]) : null,
        valcopack: cxpSheet ? findRowByKeywords(cxpSheet, ["valcopack"]) : null,
        termopack: cxpSheet ? findRowByKeywords(cxpSheet, ["termopack"]) : null,
        cartoneraApolo: cxpSheet ? findRowByKeywords(cxpSheet, ["cartonera apolo", "apolo"]) : null,
        multiplast: cxpSheet ? findRowByKeywords(cxpSheet, ["multiplast"]) : null,
        flexopack: cxpSheet ? findRowByKeywords(cxpSheet, ["flexopack"]) : null,
        etiofset: cxpSheet ? findRowByKeywords(cxpSheet, ["etiofset", "etiquetas y empaques"]) : null,
        smurfit: cxpSheet ? findRowByKeywords(cxpSheet, ["smurfit"]) : null,
        plasticosCaribe: cxpSheet ? findRowByKeywords(cxpSheet, ["plasticos del caribe", "plasticos caribe"]) : null,
        industriasNacionales: cxpSheet ? findRowByKeywords(cxpSheet, ["industrias nacionales", "inca"]) : null,
        distribuidoraCorripo: cxpSheet ? findRowByKeywords(cxpSheet, ["distribuidora corripo", "corripo"]) : null,
        otrosProveedores: cxpSheet ? findRowByKeywords(cxpSheet, ["otros proveedores"]) : null,
        
        // Indicators
        costosGastoYtd: cxpSheet ? findRowByKeywords(cxpSheet, ["costos + gasto", "opex + capex", "opex+capex"]) : null,
        dpo: cxpSheet ? findRowByKeywords(cxpSheet, ["dpo", "days payable outstanding"]) : null
    };

    const pnlRowsPpto = pptoPnlSheet ? {
        ingresos: findRowByKeywords(pptoPnlSheet, FINANCIAL_KEYWORDS.ingresos),
        costos: findRowByKeywords(pptoPnlSheet, FINANCIAL_KEYWORDS.costos),
        opex: findRowByKeywords(pptoPnlSheet, FINANCIAL_KEYWORDS.opex),
        ebitda: findRowByKeywords(pptoPnlSheet, FINANCIAL_KEYWORDS.ebitda),
        utilidad: findRowByKeywords(pptoPnlSheet, FINANCIAL_KEYWORDS.utilidad, 60),
        cashflow: findRowByKeywords(pptoPnlSheet, FINANCIAL_KEYWORDS.cashflow),
        tasa_cambio: findRowByKeywords(pptoPnlSheet, FINANCIAL_KEYWORDS.tasa_cambio),
        admin: findRowByKeywords(pptoPnlSheet, detailedOpexKeywords.admin),
        mercadeo: findRowByKeywords(pptoPnlSheet, detailedOpexKeywords.mercadeo),
        comercial: findRowByKeywords(pptoPnlSheet, detailedOpexKeywords.comercial),
        logistica: findRowByKeywords(pptoPnlSheet, detailedOpexKeywords.logistica)
    } : null;

    const segmentRowsPpto = pptoPnlSheet ? detectSegments(pptoPnlSheet, segmentKeywords) : {};

    const balanceRowsPpto = pptoBalanceSheet ? {
        activos: findRowByKeywords(pptoBalanceSheet, balanceKeywords.activos),
        pasivos: findRowByKeywords(pptoBalanceSheet, balanceKeywords.pasivos),
        patrimonio: findRowByKeywords(pptoBalanceSheet, balanceKeywords.patrimonio),
        gananciaAcumulada: findRowByKeywords(pptoBalanceSheet, ["ganancia acumulada", "utilidad acumulada", "ganancias acumuladas", "utilidades acumuladas"]),
        utilidadesRetenidas: findRowByKeywords(pptoBalanceSheet, ["utilidades retenidas", "utilidad retenida"])
    } : null;

    const cfRowsPpto = pptoCashflowSheet ? {
        beginning: findRowByKeywords(pptoCashflowSheet, FINANCIAL_KEYWORDS.cf_beginning),
        operating: findRowByKeywords(pptoCashflowSheet, FINANCIAL_KEYWORDS.cf_operating),
        wc: findRowByKeywords(pptoCashflowSheet, FINANCIAL_KEYWORDS.cf_wc),
        cxc: findRowByKeywords(pptoCashflowSheet, FINANCIAL_KEYWORDS.cf_cxc),
        inv: findRowByKeywords(pptoCashflowSheet, FINANCIAL_KEYWORDS.cf_inv),
        cxp: findRowByKeywords(pptoCashflowSheet, FINANCIAL_KEYWORDS.cf_cxp),
        otrosActivos: findRowByKeywords(pptoCashflowSheet, FINANCIAL_KEYWORDS.cf_otros_activos),
        otrosActivosCorrientes: findRowByKeywords(pptoCashflowSheet, FINANCIAL_KEYWORDS.cf_otros_activos_corrientes),
        activosTerceros: findRowByKeywords(pptoCashflowSheet, FINANCIAL_KEYWORDS.cf_activos_terceros),
        pasivoLaboral: findRowByKeywords(pptoCashflowSheet, FINANCIAL_KEYWORDS.cf_pasivo_laboral),
        otrosPasivos: findRowByKeywords(pptoCashflowSheet, FINANCIAL_KEYWORDS.cf_otros_pasivos),
        otrasCxp: findRowByKeywords(pptoCashflowSheet, FINANCIAL_KEYWORDS.cf_otras_cxp),
        otrosPasivosCorrientes: findRowByKeywords(pptoCashflowSheet, FINANCIAL_KEYWORDS.cf_otros_pasivos_corrientes),
        capex: findRowByKeywords(pptoCashflowSheet, FINANCIAL_KEYWORDS.cf_capex),
        financing: findRowByKeywords(pptoCashflowSheet, FINANCIAL_KEYWORDS.cf_financing),
        netDebt: findRowByKeywords(pptoCashflowSheet, FINANCIAL_KEYWORDS.cf_net_debt, 47),
        belowEbitda: findRowByKeywords(pptoCashflowSheet, FINANCIAL_KEYWORDS.cf_below_ebitda),
        taxes: findRowByKeywords(pptoCashflowSheet, FINANCIAL_KEYWORDS.cf_taxes, 44),
        dividends: findRowByKeywords(pptoCashflowSheet, FINANCIAL_KEYWORDS.cf_dividends),
        interest: findRowByKeywords(pptoCashflowSheet, FINANCIAL_KEYWORDS.cf_interest),
        interest_earned: findRowByKeywords(pptoCashflowSheet, FINANCIAL_KEYWORDS.cf_interest_earned),
        interest_expense: findRowByKeywords(pptoCashflowSheet, FINANCIAL_KEYWORDS.cf_interest_expense),
        extraordinary: findRowByKeywords(pptoCashflowSheet, FINANCIAL_KEYWORDS.cf_extraordinary),
        diferencialCambiario: findRowByKeywords(pptoCashflowSheet, FINANCIAL_KEYWORDS.cf_diferencial_cambiario),
        change: findRowByKeywords(pptoCashflowSheet, FINANCIAL_KEYWORDS.cf_change),
        ending: findRowByKeywords(pptoCashflowSheet, FINANCIAL_KEYWORDS.cf_ending),
        dso: findRowByKeywords(pptoCashflowSheet, FINANCIAL_KEYWORDS.cf_dso),
        dpo: findRowByKeywords(pptoCashflowSheet, FINANCIAL_KEYWORDS.cf_dpo),
        dio: findRowByKeywords(pptoCashflowSheet, FINANCIAL_KEYWORDS.cf_dio)
    } : null;

    // Helper to find data column indices for a given sheet based on target dates
    const findSheetIndices = (sheet) => {
        const indices = {};
        if (!sheet) return indices;
        
        const monthNames = ["ene", "feb", "mar", "abr", "may", "jun", "jul", "ago", "sep", "oct", "nov", "dic"];
        const shortMonths = ["jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec"];

        let bestRowIdx = -1;
        let maxDates = 0;
        const allRowDates = [];

        for (let i = 0; i < Math.min(sheet.length, 50); i++) {
            const row = sheet[i];
            const currentDates = {};
            let datesCount = 0;
            
            if (!row) {
                allRowDates.push(currentDates);
                continue;
            }
            
            row.forEach((cell, j) => {
                let dateObj = null;
                if (cell instanceof Date) dateObj = cell;
                else if (typeof cell === 'number' && cell > 40000 && cell < 60000) dateObj = new Date((cell - 25569) * 86400 * 1000);
                else if (typeof cell === 'string') {
                    const val = normalizeText(cell);
                    const checkMonthPattern = (text) => {
                        const patterns = [
                            /\\b(ene(ro)?|jan(uary)?)\\b/, /\\b(feb(rero|ruary)?)\\b/, /\\b(mar(zo|ch)?)\\b/,
                            /\\b(abr(il)?|apr(il)?)\\b/, /\\b(may(o)?)\\b/, /\\b(jun(io|e)?)\\b/,
                            /\\b(jul(io|y)?)\\b/, /\\b(ago(sto)?|aug(ust)?)\\b/, /\\b(sep(t|tiembre|tember)?)\\b/,
                            /\\b(oct(ubre|ober)?)\\b/, /\\b(nov(iembre|ember)?)\\b/, /\\b(dic(iembre)?|dec(ember)?)\\b/
                        ];
                        for (let k = 0; k < patterns.length; k++) {
                            if (patterns[k].test(text)) return k;
                        }
                        return -1;
                    };
                    const finalMIdx = checkMonthPattern(val);
                    
                    if (finalMIdx !== -1) {
                        dateObj = new Date();
                        dateObj.setDate(1); // FIX: prevent rollover when today is 31st and month is Feb/etc.
                        dateObj.setMonth(finalMIdx);
                        const yearMatch = val.match(/\d{2,4}/);
                        if (yearMatch) {
                            let y = parseInt(yearMatch[0]);
                            if (y < 100) y += 2000;
                            dateObj.setFullYear(y);
                        } else {
                            // Look for year in neighboring cells if not in string
                            for (let neighborIdx = Math.max(0, i-2); neighborIdx <= Math.min(sheet.length-1, i+2); neighborIdx++) {
                                const neighborRow = sheet[neighborIdx];
                                if (!neighborRow) continue;
                                const yearInRow = neighborRow.find(c => typeof c === 'number' && c >= 2020 && c <= 2026);
                                if (yearInRow) { dateObj.setFullYear(yearInRow); break; }
                            }
                        }
                    } else if (val.match(/20\d{2}/) && val.match(/\d{1,2}/)) {
                        const dateMatch = val.match(/(20\d{2})[-/](\d{1,2})/);
                        if (dateMatch) dateObj = new Date(parseInt(dateMatch[1]), parseInt(dateMatch[2]) - 1, 1);
                    }
                }

                if (dateObj) {
                    const y = dateObj.getFullYear();
                    if (y >= 2020 && y <= 2026) {
                        const dateKey = `${dateObj.getMonth()}-${y}`;
                        if (!currentDates[dateKey]) {
                            currentDates[dateKey] = j;
                            datesCount++;
                        }
                    }
                }
            });
            
            allRowDates.push(currentDates);
            if (datesCount > maxDates) {
                maxDates = datesCount;
                bestRowIdx = i;
            }
        }
        
        let finalIndices = {};
        if (bestRowIdx !== -1) {
            finalIndices = allRowDates[bestRowIdx];
            
            // Si otra fila tiene fechas que la fila principal no tiene, también las agregamos
            for (let i = 0; i < allRowDates.length; i++) {
                if (i === bestRowIdx) continue;
                for (const key in allRowDates[i]) {
                    if (!finalIndices[key]) finalIndices[key] = allRowDates[i][key];
                }
            }
        }

        return finalIndices;
    };

    const pnlIndices = findSheetIndices(pnlSheet);
    const balanceIndices = balanceSheet ? findSheetIndices(balanceSheet) : {};
    const cfIndices = cashflowSheet ? findSheetIndices(cashflowSheet) : {};
    const wcIndices = wcSheet ? findSheetIndices(wcSheet) : {};
    const estadosIndices = estadosSheet ? findSheetIndices(estadosSheet) : {};
    const deudaIndices = deudaSheet ? findSheetIndices(deudaSheet) : {};
    const cxpIndices = cxpSheet ? findSheetIndices(cxpSheet) : {};

    const pptoPnlIndices = pptoPnlSheet ? findSheetIndices(pptoPnlSheet) : {};
    const pptoBalanceIndices = pptoBalanceSheet ? findSheetIndices(pptoBalanceSheet) : {};
    const pptoCfIndices = pptoCashflowSheet ? findSheetIndices(pptoCashflowSheet) : {};
    
    // Unificar todas las fechas detectadas en ambos reportes
    const allDateKeys = new Set([...Object.keys(pnlIndices), ...Object.keys(balanceIndices), ...Object.keys(cfIndices), ...Object.keys(wcIndices), ...Object.keys(estadosIndices), ...Object.keys(deudaIndices), ...Object.keys(cxpIndices), ...Object.keys(pptoPnlIndices), ...Object.keys(pptoBalanceIndices), ...Object.keys(pptoCfIndices)]);
    
    let dataPeriods = [];
    allDateKeys.forEach(key => {
        const [m, y] = key.split('-').map(Number);
        const d = new Date(y, m, 1);
        
        // 🚨 Filtro de seguridad: Solo permitir periodos hasta 2026 (pedido por usuario)
        // Y evitar fechas absurdamente lejanas en el pasado o futuro
        if (y >= 2020 && y <= 2026) {
            const pnlIdx = pnlIndices[key] !== undefined ? pnlIndices[key] : -1;
            const balanceIdx = balanceIndices[key] !== undefined ? balanceIndices[key] : pnlIdx;
            const cfIdx = cfIndices[key] !== undefined ? cfIndices[key] : pnlIdx;
            const wcIdx = wcIndices[key] !== undefined ? wcIndices[key] : pnlIdx;
            const estadosIdx = estadosIndices[key] !== undefined ? estadosIndices[key] : -1;
            const deudaIdx = deudaIndices[key] !== undefined ? deudaIndices[key] : -1;
            const pptoPnlIdx = pptoPnlIndices[key] !== undefined ? pptoPnlIndices[key] : pptoPnlIndices[key] !== undefined ? pptoPnlIndices[key] : -1;
            const pptoBalanceIdx = pptoBalanceIndices[key] !== undefined ? pptoBalanceIndices[key] : pptoPnlIdx !== -1 ? pptoPnlIdx : balanceIdx;
            const pptoCfIdx = pptoCfIndices[key] !== undefined ? pptoCfIndices[key] : pptoPnlIdx !== -1 ? pptoPnlIdx : cfIdx;

            dataPeriods.push({ date: d, pnlIdx, balanceIdx, cfIdx, wcIdx, estadosIdx, deudaIdx, pptoPnlIdx, pptoBalanceIdx, pptoCfIdx });
        }
    });
    
    // Ordenar cronológicamente
    dataPeriods.sort((a, b) => a.date - b.date);

    if (dataPeriods.length === 0) {
        pnlRows.ingresos.forEach((cell, j) => {
            if (j === 0) return;
            const val = cleanNumber(cell);
            if (val !== 0 && !isNaN(val)) {
                const d = new Date();
                d.setDate(1);
                d.setMonth(d.getMonth() - (pnlRows.ingresos.length - j));
                dataPeriods.push({ date: d, pnlIdx: j, balanceIdx: j });
            }
        });
    }

    if (dataPeriods.length === 0) {
        return { error: "No se encontraron periodos o fechas válidas en las cabeceras." };
    }

    const bSheetToUse = balanceSheet || pnlSheet;

    let conceptRow20 = "";
    const refSheetForRow20 = balanceSheet || bSheetToUse;
    if (refSheetForRow20 && refSheetForRow20[19]) {
        conceptRow20 = extractConceptName(refSheetForRow20[19]).trim();
    }

    const fullRows = pnlSheet.filter(row => {
        const rawC = extractConceptName(row);
        if (!rawC) return false;
        const concept = normalizeText(rawC);
        if (concept.includes("formatcode") || concept.includes("unnamed") || concept.length < 2) return false;
        return dataPeriods.some(p => p.pnlIdx !== -1 && (typeof row[p.pnlIdx] === 'number' || !isNaN(cleanNumber(row[p.pnlIdx]))));
    }).map(row => {
        const rowValues = {};
        dataPeriods.forEach(p => {
            rowValues[formatDateKey(p.date)] = p.pnlIdx !== -1 ? getVal(row, p.pnlIdx) : 0;
        });
        const rawConcept = extractConceptName(row);
        let renamedConcept = rawConcept;
        if (normalizeText(renamedConcept) === "ganancia del periodo") renamedConcept = "Beneficio Neto del Periodo";
        return { concept: renamedConcept, values: rowValues };
    });

    const estadosFullRows = estadosSheet ? estadosSheet.filter(row => {
        const rawC = extractConceptName(row);
        if (!rawC) return false;
        const concept = normalizeText(rawC);
        if (concept.includes("formatcode") || concept.includes("unnamed") || concept.length < 2) return false;
        return dataPeriods.some(p => p.estadosIdx !== -1 && (typeof row[p.estadosIdx] === 'number' || !isNaN(cleanNumber(row[p.estadosIdx]))));
    }).map(row => {
        const rowValues = {};
        dataPeriods.forEach(p => {
            rowValues[formatDateKey(p.date)] = p.estadosIdx !== -1 ? getVal(row, p.estadosIdx) : 0;
        });
        const rawConcept = extractConceptName(row);
        return { concept: rawConcept, values: rowValues };
    }) : [];

    const pptoFullRows = pptoPnlSheet ? pptoPnlSheet.filter(row => {
        const rawC = extractConceptName(row);
        if (!rawC) return false;
        const concept = normalizeText(rawC);
        if (concept.includes("formatcode") || concept.includes("unnamed") || concept.length < 2) return false;
        return dataPeriods.some(p => p.pptoPnlIdx !== -1 && (typeof row[p.pptoPnlIdx] === 'number' || !isNaN(cleanNumber(row[p.pptoPnlIdx]))));
    }).map(row => {
        const rowValues = {};
        dataPeriods.forEach(p => {
            rowValues[formatDateKey(p.date)] = p.pptoPnlIdx !== -1 ? getVal(row, p.pptoPnlIdx) : 0;
        });
        const rawConcept = extractConceptName(row);
        let renamedConcept = rawConcept;
        if (normalizeText(renamedConcept) === "ganancia del periodo") renamedConcept = "Beneficio Neto del Periodo";
        return { concept: renamedConcept, values: rowValues };
    }) : [];

    const balanceFullRows = bSheetToUse.filter((row, idx_row) => {
        const rawC = extractConceptName(row);
        if (!rawC) return false;
        const conceptStr = rawC;
        const concept = normalizeText(conceptStr);
        if (idx_row === 19) return true;
        if (conceptRow20 && rawC.trim().toLowerCase() === conceptRow20.toLowerCase()) return true;
        if (concept.includes("formatcode") || concept.includes("unnamed") || concept.length < 2) return false;
        
        const isHeader = concept === "activos" || concept === "pasivos" || concept === "patrimonio" || 
                         concept === "capital" || concept === "pasivo y capital" || 
                         concept === "activo" || concept === "pasivo" ||
                         concept === "ingresos" || concept === "costos" || concept === "gastos";
        
        const isAccountingRule = concept.includes("ganancia acumulada") || concept.includes("utilidad acumulada") || 
                                concept.includes("utilidades retenidas") || concept.includes("ganancia retenida") ||
                                concept.includes("beneficio neto") || concept.includes("utilidad del ejercicio");

        if (isHeader && !isAccountingRule && !concept.includes("total")) return false;
        if (!isAccountingRule && (concept.includes("en mdop") || concept.includes("estado de situacion") || concept.includes("reporte pa"))) return false;
        
        const checkMonthPattern = (text) => {
            const patterns = [
                /\\b(ene(ro)?|jan(uary)?)\\b/, /\\b(feb(rero|ruary)?)\\b/, /\\b(mar(zo|ch)?)\\b/,
                /\\b(abr(il)?|apr(il)?)\\b/, /\\b(may(o)?)\\b/, /\\b(jun(io|e)?)\\b/,
                /\\b(jul(io|y)?)\\b/, /\\b(ago(sto)?|aug(ust)?)\\b/, /\\b(sep(t|tiembre|tember)?)\\b/,
                /\\b(oct(ubre|ober)?)\\b/, /\\b(nov(iembre|ember)?)\\b/, /\\b(dic(iembre)?|dec(ember)?)\\b/
            ];
            for (let k = 0; k < patterns.length; k++) {
                if (patterns[k].test(text)) return true;
            }
            return false;
        };
        if (checkMonthPattern(concept)) return false;
        
        const isTypicalBalance = concept.includes("activo") || concept.includes("pasivo") || 
                                 concept.includes("patrimonio") || concept.includes("efectivo") || 
                                 concept.includes("bancos") || concept.includes("cobrar") || 
                                 concept.includes("inventario") || concept.includes("propiedad") || 
                                 concept.includes("ppe") || concept.includes("prestamos") || 
                                 concept.includes("capital") || concept.includes("reserva") ||
                                 concept.includes("covenant") || concept.includes("deuda neta") ||
                                 concept.includes("ltm ebitda") || concept.includes("ebitda r12") || 
                                 concept.includes("deuda bruta") || concept.includes("deuda total") ||
                                 concept.includes("deuda subordinada") || concept.includes("deuda sin subordinada") ||
                                 concept.includes("apalancamiento") ||
                                 concept.includes("capacidad de pago") || concept.includes("capacidad") || 
                                 concept.includes("razon corriente") ||
                                 concept.includes("ganancia") || concept.includes("beneficio");

        const isNetIncomeInBalance = (concept.includes("utilidad") || concept.includes("ganancia") || concept.includes("beneficio") || concept.includes("ganancia")) && 
                                     (concept.includes("ejercicio") || concept.includes("periodo") || concept.includes("neto"));

        if (bSheetToUse === pnlSheet && !isTypicalBalance && !isNetIncomeInBalance) {
            const pnlStrict = ["ingresos", "ventas netas", "costo de ventas", "utilidad bruta", "ebitda", "ggadm", "ebit"];
            if (pnlStrict.some(p => concept === p || concept.includes(p))) return false;
        }
        
        if (isTypicalBalance || isNetIncomeInBalance) return true;

        return dataPeriods.some(p => {
            const curBIdx = p.balanceIdx !== -1 ? p.balanceIdx : p.pnlIdx;
            if (curBIdx === -1) return false;
            const val = getBalanceVal(row, curBIdx);
            return val !== 0 && !isNaN(val);
        });
    }).map(row => {
        const rawConcept = extractConceptName(row);
        let renamedConcept = rawConcept;
        const normConcept = normalizeText(rawConcept);
        
        const isTargetNetIncome = normConcept === "ganancia del periodo" || normConcept === "utilidad del ejercicio" || 
            normConcept === "resultado del periodo" || normConcept.includes("beneficio neto") || 
            normConcept.includes("utilidad neta") || normConcept.includes("ganancia neta") ||
            normConcept.includes("resultado neta");

        const rowValues = {};
        dataPeriods.forEach(p => {
            const curBIdx = p.balanceIdx !== -1 ? p.balanceIdx : p.pnlIdx;
            let val = curBIdx !== -1 ? getBalanceVal(row, curBIdx) : 0;
            
            if (isTargetNetIncome && val === 0 && curBIdx !== -1) {
                const gAcum = getBalanceVal(balanceRows.gananciaAcumulada, curBIdx);
                const uRet = getBalanceVal(balanceRows.utilidadesRetenidas, curBIdx);
                if (gAcum !== 0 || uRet !== 0) val = uRet - gAcum;
            }
            
            rowValues[formatDateKey(p.date)] = val;
        });

        if (isTargetNetIncome) renamedConcept = "Beneficio Neto del Periodo";
        return { concept: renamedConcept, values: rowValues };
    });

    const pptoBSheetToUse = pptoBalanceSheet || pptoPnlSheet;
    const pptoBalanceFullRows = pptoBSheetToUse ? pptoBSheetToUse.filter((row, idx_row) => {
        const rawC = extractConceptName(row);
        if (!rawC) return false;
        const conceptStr = rawC;
        const concept = normalizeText(conceptStr);
        if (idx_row === 19) return true;
        if (conceptRow20 && rawC.trim().toLowerCase() === conceptRow20.toLowerCase()) return true;
        if (concept.includes("formatcode") || concept.includes("unnamed") || concept.length < 2) return false;
        
        const isHeader = concept === "activos" || concept === "pasivos" || concept === "patrimonio" || 
                         concept === "capital" || concept === "pasivo y capital" || 
                         concept === "activo" || concept === "pasivo" ||
                         concept === "ingresos" || concept === "costos" || concept === "gastos";
        
        const isAccountingRule = concept.includes("ganancia acumulada") || concept.includes("utilidad acumulada") || 
                                concept.includes("utilidades retenidas") || concept.includes("ganancia retenida") ||
                                concept.includes("beneficio neto") || concept.includes("utilidad del ejercicio");

        if (isHeader && !isAccountingRule && !concept.includes("total")) return false;
        if (!isAccountingRule && (concept.includes("en mdop") || concept.includes("estado de situacion") || concept.includes("reporte pa"))) return false;
        
        const checkMonthPattern = (text) => {
            const patterns = [
                /\\b(ene(ro)?|jan(uary)?)\\b/, /\\b(feb(rero|ruary)?)\\b/, /\\b(mar(zo|ch)?)\\b/,
                /\\b(abr(il)?|apr(il)?)\\b/, /\\b(may(o)?)\\b/, /\\b(jun(io|e)?)\\b/,
                /\\b(jul(io|y)?)\\b/, /\\b(ago(sto)?|aug(ust)?)\\b/, /\\b(sep(t|tiembre|tember)?)\\b/,
                /\\b(oct(ubre|ober)?)\\b/, /\\b(nov(iembre|ember)?)\\b/, /\\b(dic(iembre)?|dec(ember)?)\\b/
            ];
            for (let k = 0; k < patterns.length; k++) {
                if (patterns[k].test(text)) return true;
            }
            return false;
        };
        if (checkMonthPattern(concept)) return false;
        
        const isTypicalBalance = concept.includes("activo") || concept.includes("pasivo") || 
                                 concept.includes("patrimonio") || concept.includes("efectivo") || 
                                 concept.includes("bancos") || concept.includes("cobrar") || 
                                 concept.includes("inventario") || concept.includes("propiedad") || 
                                 concept.includes("ppe") || concept.includes("prestamos") || 
                                 concept.includes("capital") || concept.includes("reserva") ||
                                 concept.includes("covenant") || concept.includes("deuda neta") ||
                                 concept.includes("ltm ebitda") || concept.includes("ebitda r12") || 
                                 concept.includes("deuda bruta") || concept.includes("deuda total") ||
                                 concept.includes("deuda subordinada") || concept.includes("deuda sin subordinada") ||
                                 concept.includes("apalancamiento") ||
                                 concept.includes("capacidad de pago") || concept.includes("capacidad") || 
                                 concept.includes("razon corriente") ||
                                 concept.includes("ganancia") || concept.includes("beneficio");

        const isNetIncomeInBalance = (concept.includes("utilidad") || concept.includes("ganancia") || concept.includes("beneficio") || concept.includes("ganancia")) && 
                                     (concept.includes("ejercicio") || concept.includes("periodo") || concept.includes("neto"));

        if (pptoBSheetToUse === pptoPnlSheet && !isTypicalBalance && !isNetIncomeInBalance) {
            const pnlStrict = ["ingresos", "ventas netas", "costo de ventas", "utilidad bruta", "ebitda", "ggadm", "ebit"];
            if (pnlStrict.some(p => concept === p || concept.includes(p))) return false;
        }
        
        if (isTypicalBalance || isNetIncomeInBalance) return true;

        return dataPeriods.some(p => {
            const curBIdx = p.pptoBalanceIdx !== -1 ? p.pptoBalanceIdx : p.pptoPnlIdx;
            if (curBIdx === -1) return false;
            const val = getBalanceVal(row, curBIdx);
            return val !== 0 && !isNaN(val);
        });
    }).map(row => {
        const rawConcept = extractConceptName(row);
        let renamedConcept = rawConcept;
        const normConcept = normalizeText(rawConcept);
        
        const isTargetNetIncome = normConcept === "ganancia del periodo" || normConcept === "utilidad del ejercicio" || 
            normConcept === "resultado del periodo" || normConcept.includes("beneficio neto") || 
            normConcept.includes("utilidad neta") || normConcept.includes("ganancia neta") ||
            normConcept.includes("resultado neta");

        const rowValues = {};
        dataPeriods.forEach(p => {
            const curBIdx = p.pptoBalanceIdx !== -1 ? p.pptoBalanceIdx : p.pptoPnlIdx;
            let val = curBIdx !== -1 ? getBalanceVal(row, curBIdx) : 0;
            
            if (isTargetNetIncome && val === 0 && curBIdx !== -1) {
                const gAcum = balanceRowsPpto ? getBalanceVal(balanceRowsPpto.gananciaAcumulada, curBIdx) : 0;
                const uRet = balanceRowsPpto ? getBalanceVal(balanceRowsPpto.utilidadesRetenidas, curBIdx) : 0;
                if (gAcum !== 0 || uRet !== 0) val = uRet - gAcum;
            }
            
            rowValues[formatDateKey(p.date)] = val;
        });

        if (isTargetNetIncome) renamedConcept = "Beneficio Neto del Periodo";
        return { concept: renamedConcept, values: rowValues };
    }) : [];

    const wcFullRows = wcSheet ? wcSheet.filter(row => {
        const rawC = extractConceptName(row);
        if (!rawC) return true;
        const conceptStr = rawC;
        if (!conceptStr) return true;
        
        const concept = normalizeText(conceptStr);
        if (concept.includes("formatcode") || concept.includes("unnamed")) return false;
        
        let hasNumbers = dataPeriods.some(p => {
            if (p.wcIdx === -1) return false;
            let val = getVal(row, p.wcIdx, false);
            return val !== 0 && !isNaN(val);
        });

        if (!hasNumbers) {
            const isMetricName = concept.includes('dso') || concept.includes('dpo') || concept.includes('dio') || concept.includes('dias') || concept.includes('tasa') || concept.includes('cxc') || concept.includes('cxp');
            if (isMetricName) return false;
            return conceptStr.length > 2; // Keep category headers length > 2
        }
        
        return true;
    }).map((row, idx) => {
        const rowValues = {};
        const conceptStr = extractConceptName(row);
        
        if (!conceptStr) {
            return { concept: `_spacer_${idx}`, isSpacer: true, values: {} };
        }
        
        const lowerConcept = conceptStr.toLowerCase();
        
        // Determinar si debemos escalar el valor a millones
        // Evitamos escalar días, tasas, porcentajes y variables moneda
        const isRatioOrRate = lowerConcept.includes('dso') || 
                              lowerConcept.includes('dpo') || 
                              lowerConcept.includes('dio') ||
                              lowerConcept.includes('days') || 
                              lowerConcept.includes('dias') || 
                              lowerConcept.includes('días') || 
                              lowerConcept.includes('%') || 
                              (lowerConcept.includes('tasa') && !lowerConcept.includes('impacto')) || 
                              lowerConcept === 'dop' || 
                              lowerConcept === 'eur' || 
                              lowerConcept === 'usd' || 
                              lowerConcept.includes('var ');

        let hasData = false;
        dataPeriods.forEach(p => {
            let val = p.wcIdx !== -1 ? getVal(row, p.wcIdx, false) : 0;
            if (!isRatioOrRate) {
                val = val / 1000000;
            }
            if (val !== 0) hasData = true;
            rowValues[formatDateKey(p.date)] = val;
        });

        return { concept: conceptStr, isHeader: !hasData, values: rowValues };
    }) : [];

    const getBalanceIdx = (date, pnlIdx) => {
        const key = `${date.getMonth()}-${date.getFullYear()}`;
        return balanceIndices[key] !== undefined ? balanceIndices[key] : pnlIdx;
    };

    const result = dataPeriods.map(point => {
        const pIdx = point.pnlIdx;
        const bIdx = point.balanceIdx;
        const cfIdx = point.cfIdx;
        const wcIdx = point.wcIdx !== undefined ? point.wcIdx : -1;
        const dIdx = point.deudaIdx !== undefined ? point.deudaIdx : -1;
        
        const pptoPnlIdx = point.pptoPnlIdx;
        const pptoBIdx = point.pptoBalanceIdx;
        const pptoCfIdx = point.pptoCfIdx;

        const ingresos = pIdx !== -1 ? getVal(pnlRows.ingresos, pIdx) : 0;
        const costos = pIdx !== -1 ? getVal(pnlRows.costos, pIdx) : 0;
        const ebitda = pIdx !== -1 ? getVal(pnlRows.ebitda, pIdx) : 0;
        
        let opex = pIdx !== -1 ? getVal(pnlRows.opex, pIdx) : 0;
        const impliedOpex = Math.abs(Math.abs(ingresos) - Math.abs(costos) - ebitda);
        if (ebitda !== 0 && opex === 0) {
            opex = impliedOpex;
        }

        let utilidad = pIdx !== -1 ? getVal(pnlRows.utilidad, pIdx) : 0;
        
        if (utilidad === 0 && bIdx !== -1) {
            const gananciaAcum = getBalanceVal(balanceRows.gananciaAcumulada, bIdx);
            const utilRetenidas = getBalanceVal(balanceRows.utilidadesRetenidas, bIdx);
            if (gananciaAcum !== 0 || utilRetenidas !== 0) {
                utilidad = utilRetenidas - gananciaAcum;
            }
        }

        let cashflowVal = (pIdx !== -1 ? getVal(pnlRows.cashflow, pIdx) : 0) || (cfIdx !== -1 ? getVal(cfRows.change, cfIdx, false) : utilidad);

        const activos = bIdx !== -1 ? getBalanceVal(balanceRows.activos, bIdx) : 0;
        const pasivos = bIdx !== -1 ? getBalanceVal(balanceRows.pasivos, bIdx) : 0;
        const patrimonio = bIdx !== -1 ? getBalanceVal(balanceRows.patrimonio, bIdx) : 0;
        const tasaCambio = pIdx !== -1 ? getVal(pnlRows.tasa_cambio, pIdx) : 1;

        // Extraer Detalle de la Vista Deuda
        let tasaDop = null;
        let tasaUsd = null;
        let deudaNetaUsd = null;
        let bancaTotal = null;
        let relacionadaTotal = null;
        let deudaTotal = null;
        let efectivo = null;
        let deudaNeta = null;
        
        let bancos = {
            'Banco Popular': null,
            'Banco Santa Cruz': null,
            'Scotiabank': null,
            'Loganville': null
        };
        let tasasPorBanco = {
            'Banco Popular': null,
            'Banco Santa Cruz': null,
            'Scotiabank': null
        };
        
        if (deudaSheet && dIdx !== -1) {
            const tasaDopRow = findRowByKeywords(deudaSheet, ["promedio ponderado dop"]);
            if (tasaDopRow) tasaDop = getBalanceVal(tasaDopRow, dIdx);
            
            const tasaUsdRow = findRowByKeywords(deudaSheet, ["promedio ponderado usd", "tasa usd", "fx dop/usd", "tasa de cambio"]);
            if (tasaUsdRow) tasaUsd = getBalanceVal(tasaUsdRow, dIdx);
            
            const deudaNetaUsdRow = findRowByKeywords(deudaSheet, ["deuda neta total usd"]);
            if (deudaNetaUsdRow) deudaNetaUsd = getBalanceVal(deudaNetaUsdRow, dIdx);
            
            bancaTotal = getBalanceVal(deudaSheet[5], dIdx); // Fila 6 es index 5
            relacionadaTotal = getBalanceVal(deudaSheet[6], dIdx); // Fila 7
            deudaTotal = getBalanceVal(deudaSheet[7], dIdx); // Fila 8
            efectivo = getBalanceVal(deudaSheet[8], dIdx); // Fila 9
            deudaNeta = getBalanceVal(deudaSheet[9], dIdx); // Fila 10
            
            // Loganville = Deuda Relacionada (fila 7, index 6), mismo valor ya capturado en relacionadaTotal.
            // La fila 197 ("Total DOP") usa FX histórico variable y da valores incorrectos para períodos pasados.
            bancos['Loganville'] = relacionadaTotal;
            
            // Recorrer filas 39-67 para acumular por banco (indices 38 a 66), asumiendo col 2 (index 1) es el nombre del banco
            let popV = 0, scV = 0, scotiV = 0;
            let popFound = false, scFound = false, scotiFound = false;
            
            for (let r = 38; r <= 66; r++) {
                const rRow = deudaSheet[r];
                if (!rRow) continue;
                const bancoName = normalizeText(rRow[1] || rRow[2] || '');
                const val = getBalanceVal(rRow, dIdx) || 0;
                
                if (bancoName.includes("popular")) { popV += val; popFound = true; }
                else if (bancoName.includes("santa cruz")) { scV += val; scFound = true; }
                else if (bancoName.includes("scotia") || bancoName.includes("bns")) { scotiV += val; scotiFound = true; }
            }
            if (popFound) bancos['Banco Popular'] = popV;
            if (scFound) bancos['Banco Santa Cruz'] = scV;
            if (scotiFound) bancos['Scotiabank'] = scotiV;
            
            // Tasas promedio ponderado por banco: filas 120-147 (indices 119 a 146). Usa la tasa general DOP para simplificar si no se encuentra especifica,
            // pero el usuario pidio calcularla, aunque tambien dijo "Si es complejo, simplificar usando la tasa promedio general (fila 148)"
            const tasaGeneralDOP = getBalanceVal(deudaSheet[147], dIdx); // Fila 148 is index 147 
            if (popFound) tasasPorBanco['Banco Popular'] = tasaGeneralDOP;
            if (scFound) tasasPorBanco['Banco Santa Cruz'] = tasaGeneralDOP;
            if (scotiFound) tasasPorBanco['Scotiabank'] = tasaGeneralDOP;
        }

        const key = `${point.date.getMonth()}-${point.date.getFullYear()}`;
        let presDIdx = presentacionSheet ? (deudaIndices[key] !== undefined ? deudaIndices[key] : -1) : -1; 
        
        let prestTasaDop, prestTasaUsd, prestNetaUsd, prestNetaBancUsd, prestInd1, prestInd2, prestInd3, prestInd4;
        
        if (presentacionSheet) {
            // Find columns in Presentacion using standard findSheetIndices
            const presIndices = findSheetIndices(presentacionSheet);
            let pIdx = presIndices[key] !== undefined ? presIndices[key] : -1;
            
            if (pIdx !== -1) {
                // Tasa DOP
                const rowDop = findRowByKeywords(presentacionSheet, ["tasa dop"]);
                if (rowDop) prestTasaDop = getBalanceVal(rowDop, pIdx);
                
                const rowUsd = findRowByKeywords(presentacionSheet, ["tasa usd"]);
                if (rowUsd) prestTasaUsd = getBalanceVal(rowUsd, pIdx);
                
                const rowNeta = findRowByKeywords(presentacionSheet, ["deuda neta usd"]);
                if (rowNeta) prestNetaUsd = getBalanceVal(rowNeta, pIdx);
                
                const rowNetaBanc = findRowByKeywords(presentacionSheet, ["deuda neta bancaria usd", "bancaria usd"]);
                if (rowNetaBanc) prestNetaBancUsd = getBalanceVal(rowNetaBanc, pIdx);
                
                const rowInd1 = findRowByKeywords(presentacionSheet, ["deuda neta bancaria /", "<=4.0x"]);
                if (rowInd1) prestInd1 = getBalanceVal(rowInd1, pIdx);
                
                const rowInd2 = findRowByKeywords(presentacionSheet, ["apalancamiento", "<=2.0x"]);
                if (rowInd2) prestInd2 = getBalanceVal(rowInd2, pIdx);
                
                const rowInd3 = findRowByKeywords(presentacionSheet, ["capacidad de pago"]);
                if (rowInd3) prestInd3 = getBalanceVal(rowInd3, pIdx);
                
                const rowInd4 = findRowByKeywords(presentacionSheet, ["razon corriente", ">= 1.5x"]);
                if (rowInd4) prestInd4 = getBalanceVal(rowInd4, pIdx);
            }
        }

        // Deuda Total is BP48 + BP49 + BP51 + BP52 (Indices 47, 48, 50, 51)
        let calcDeudaTotal = null;
        if (bIdx !== -1 && balanceSheet) {
            const v48 = getBalanceVal(balanceSheet[47], bIdx) || 0;
            const v49 = getBalanceVal(balanceSheet[48], bIdx) || 0;
            const v51 = getBalanceVal(balanceSheet[50], bIdx) || 0;
            const v52 = getBalanceVal(balanceSheet[51], bIdx) || 0;
            calcDeudaTotal = v48 + v49 + v51 + v52;
            if (calcDeudaTotal === 0 && !getBalanceVal(balanceSheet[47], bIdx)) calcDeudaTotal = null; // Only set to null if all were missing
        }
        
        const fallbackDeudaTotal = (bIdx !== -1 ? getBalanceVal(balanceRows.deudaTotal, bIdx) : null) || calcDeudaTotal;
        const fallbackCaja = bIdx !== -1 ? getBalanceVal(balanceRows.cajaEfectivo, bIdx) : null;
        const fallbackTasaUsd = (prestTasaUsd !== undefined ? prestTasaUsd : tasaUsd) || tasaCambio || 1;
        
        let calculatedDeudaNetaUsd = prestNetaUsd !== undefined ? prestNetaUsd : deudaNetaUsd;
        if (calculatedDeudaNetaUsd === null && fallbackDeudaTotal !== null && fallbackCaja !== null) {
            calculatedDeudaNetaUsd = (fallbackDeudaTotal - fallbackCaja) / fallbackTasaUsd;
        }

        let rawBanc = prestNetaBancUsd !== undefined ? prestNetaBancUsd : (bIdx !== -1 ? (getBalanceVal(balanceRows.deudaBancariaNetaUSD, bIdx) || getBalanceVal(balanceSheet ? balanceSheet[109] : null, bIdx)) : null);
        let calculatedDeudaNetaBancUSD = rawBanc;
        const fxRate = tasaCambio || 1;
        if (calculatedDeudaNetaBancUSD !== null && fxRate > 0) {
            calculatedDeudaNetaBancUSD = calculatedDeudaNetaBancUSD / fxRate;
        }

        const deudaMetrics = {
            tasaDop: prestTasaDop !== undefined ? prestTasaDop : tasaDop,
            tasaUsd: prestTasaUsd !== undefined ? prestTasaUsd : tasaUsd,
            tasaCambio: tasaCambio,
            deudaNetaUsd: calculatedDeudaNetaUsd,
            deudaNetaBancUSD: calculatedDeudaNetaBancUSD,
            covenantLean: prestInd1 !== undefined ? prestInd1 : (bIdx !== -1 ? (getBalanceVal(balanceRows.covenantLean, bIdx) || getBalanceVal(balanceSheet ? balanceSheet[110] : null, bIdx)) : null),
            apalancamiento: prestInd2 !== undefined ? prestInd2 : (bIdx !== -1 ? (getBalanceVal(balanceRows.apalancamiento, bIdx) || getBalanceVal(balanceSheet ? balanceSheet[111] : null, bIdx)) : null),
            capacidadPago: prestInd3 !== undefined ? prestInd3 : (bIdx !== -1 ? (getBalanceVal(balanceRows.capacidadPago, bIdx) || getBalanceVal(balanceSheet ? balanceSheet[112] : null, bIdx)) : null),
            razonCorriente: prestInd4 !== undefined ? prestInd4 : (bIdx !== -1 ? (getBalanceVal(balanceRows.razonCorriente, bIdx) || getBalanceVal(balanceSheet ? balanceSheet[113] : null, bIdx)) : null),
            cajaEfectivo: fallbackCaja || (bIdx !== -1 ? getBalanceVal(balanceSheet ? balanceSheet[12] : null, bIdx) : null),
            deudaTotal: fallbackDeudaTotal,
            debtDetail: {
                bancaTotal,
                relacionadaTotal,
                deudaTotal: deudaTotal || fallbackDeudaTotal,
                efectivo: efectivo || fallbackCaja,
                deudaNeta: deudaNeta,
                deudaNetaUSD: calculatedDeudaNetaUsd,
                tasaDOP: prestTasaDop !== undefined ? prestTasaDop : tasaDop,
                bancos,
                tasasPorBanco
            }
        };

        // Extraer Detalle de Cash Flow completo si existe
        const cashflowDetail = {};
        if (cashflowSheet && cfIdx !== -1) {
            Object.keys(cfRows).forEach(key => {
                const row = cfRows[key];
                if (row) cashflowDetail[key] = getVal(row, cfIdx, false);
            });
            // Compound keys integration
            cashflowDetail.otrosActivos = (cashflowDetail.otrosActivos || 0) + (cashflowDetail.otrosActivosCorrientes || 0) + (cashflowDetail.activosTerceros || 0);
            cashflowDetail.otrosPasivos = cashflowDetail.otrasCxp || 0;
            cashflowDetail.pasivoLaboral = (cashflowDetail.pasivoLaboral || 0) + (cashflowDetail.otrosPasivosCorrientes || 0);
            
            if (cashflowDetail.interest_earned !== undefined || cashflowDetail.interest_expense !== undefined) {
                const sumInterest = (cashflowDetail.interest_earned || 0) + (cashflowDetail.interest_expense || 0);
                if (sumInterest !== 0) {
                    cashflowDetail.interest = sumInterest;
                }
            }
            
            // Prevent double counting if we iterate this dict anywhere else later
            delete cashflowDetail.otrosActivosCorrientes;
            delete cashflowDetail.activosTerceros;
            delete cashflowDetail.otrasCxp;
            delete cashflowDetail.otrosPasivosCorrientes;
            delete cashflowDetail.interest_earned;
            delete cashflowDetail.interest_expense;
        }

        const wcDetail = {};
        if (wcSheet && wcIdx !== -1) {
            Object.keys(wcRows).forEach(key => {
                const row = wcRows[key];
                if (row) wcDetail[key] = getVal(row, wcIdx, false);
            });
        }

        // === EXTRACT DETALLE CXP ===
        const cxpKeyObj = `${point.date.getMonth()}-${point.date.getFullYear()}`;
        const cxpIdx = cxpSheet ? (cxpIndices[cxpKeyObj] !== undefined ? cxpIndices[cxpKeyObj] : -1) : -1;
        
        let cxpData = null;
        if (cxpSheet && cxpIdx !== -1) {
            const getCxpVal = (row) => row ? getBalanceVal(row, cxpIdx) : null;
            
            // Check if there's actually any real data in the column
            let hasData = false;
            for (let r in cxpRows) {
                const val = getCxpVal(cxpRows[r]);
                if (val !== null && val !== 0 && val !== undefined && !isNaN(val)) {
                    hasData = true;
                    break;
                }
            }
            
            if (hasData) {
                cxpData = {
                    provisionSinFactura: getCxpVal(cxpRows.provisionSinFactura),
                    corriente: getCxpVal(cxpRows.corriente),
                    dias0_30: getCxpVal(cxpRows.dias0_30),
                    dias31_60: getCxpVal(cxpRows.dias31_60),
                    dias61_90: getCxpVal(cxpRows.dias61_90),
                    dias91_120: getCxpVal(cxpRows.dias91_120),
                    dias121_150: getCxpVal(cxpRows.dias121_150),
                    dias151_180: getCxpVal(cxpRows.dias151_180),
                    dias180Mas: getCxpVal(cxpRows.dias180Mas),
                    
                    // Suppliers
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
                    
                    // Indicators
                    costosGastoYtd: getCxpVal(cxpRows.costosGastoYtd),
                    dpo: getCxpVal(cxpRows.dpo)
                };
            }
        }

        const bCxp = (balanceRows.cxp && bIdx !== -1) ? getBalanceVal(balanceRows.cxp, bIdx) : null;
        let cxpVal = bCxp !== null ? bCxp : (ingresos !== 0 ? Math.abs(ingresos) * 0.22 : 800);
        
        let isProjectedMonth = false;
        let cxpObj = cxpData || {};

        if (!cxpData) {
            isProjectedMonth = point.date.getFullYear() >= 2026;
            if (isProjectedMonth) {
                const fillIfMissing = (field, ratio) => {
                    if (cxpObj[field] === undefined || cxpObj[field] === null || cxpObj[field] === 0) {
                        cxpObj[field] = cxpVal * ratio;
                    }
                };

                fillIfMissing('provisionSinFactura', 0.12);
                fillIfMissing('corriente', 0.30);
                fillIfMissing('dias0_30', 0.25);
                fillIfMissing('dias31_60', 0.15);
                fillIfMissing('dias61_90', 0.08);
                fillIfMissing('dias91_120', 0.03);
                fillIfMissing('dias121_150', 0.01);
                fillIfMissing('dias151_180', 0.005);
                fillIfMissing('dias180Mas', 0.005);

                const sumAging = (cxpObj.provisionSinFactura || 0) + (cxpObj.corriente || 0) + 
                                (cxpObj.dias0_30 || 0) + (cxpObj.dias31_60 || 0) + 
                                (cxpObj.dias61_90 || 0) + (cxpObj.dias91_120 || 0) + 
                                (cxpObj.dias121_150 || 0) + (cxpObj.dias151_180 || 0) + 
                                (cxpObj.dias180Mas || 0);
                const agingGap = cxpVal - sumAging;
                cxpObj.corriente = (cxpObj.corriente || 0) + agingGap;

                fillIfMissing('alplaHispaniola', 0.20);
                fillIfMissing('polyplas', 0.18);
                fillIfMissing('grupoRojas', 0.03);
                fillIfMissing('raviCaribe', 0.025);
                fillIfMissing('valcopack', 0.005);
                fillIfMissing('itGlobal', 0.04);
                fillIfMissing('caasd', 0.035);
                fillIfMissing('orox', 0.03);
                fillIfMissing('proxergia', 0.02);
                fillIfMissing('sidel', 0.015);
                fillIfMissing('lifeFlex', 0.025);
                fillIfMissing('grupoLtr', 0.015);
                fillIfMissing('dafTrading', 0.01);
                fillIfMissing('frankenberg', 0.02);
                fillIfMissing('otrosProveedores', 0.35);

                const sumSuppliers = (cxpObj.alplaHispaniola || 0) + (cxpObj.polyplas || 0) + 
                                     (cxpObj.grupoRojas || 0) + (cxpObj.raviCaribe || 0) + 
                                     (cxpObj.valcopack || 0) + (cxpObj.itGlobal || 0) + 
                                     (cxpObj.caasd || 0) + (cxpObj.orox || 0) + 
                                     (cxpObj.proxergia || 0) + (cxpObj.sidel || 0) + 
                                     (cxpObj.lifeFlex || 0) + (cxpObj.grupoLtr || 0) + 
                                     (cxpObj.dafTrading || 0) + (cxpObj.frankenberg || 0) + 
                                     (cxpObj.otrosProveedores || 0);
                const supplierGap = cxpVal - sumSuppliers;
                cxpObj.otrosProveedores = (cxpObj.otrosProveedores || 0) + supplierGap;
                
                cxpObj.isProjectedDetail = true;
            }
        }

        if (cxpObj.costosGastoYtd === undefined || cxpObj.costosGastoYtd === null || cxpObj.costosGastoYtd === 0) {
            cxpObj.costosGastoMensual = Math.abs(costos) + Math.abs(opex) + Math.abs(cashflowDetail.capex || 0);
        } else {
            cxpObj.costosGastoMensual = 0;
        }
        cxpObj.cxpTotal = cxpVal;

        // Si el cambio neto viene en 0, intentar calcularlo por la suma de actividades
        if (cashflowVal === 0 && cashflowSheet && cfIdx !== -1) {
            const calculatedChange = (cashflowDetail.operating || 0) + (cashflowDetail.financing || 0) + (cashflowDetail.capex || 0);
            if (calculatedChange !== 0) cashflowVal = calculatedChange;
        }

        // === EXTRACT PPTO ===
        let pptoIngresos = 0, pptoCostos = 0, pptoEbitda = 0, pptoOpex = 0, pptoUtilidad = 0, pptoCashflowVal = 0;
        let pptoActivos = 0, pptoPasivos = 0, pptoPatrimonio = 0, pptoTasaCambio = 1;

        const pptoCashflowDetail = {};
        if (pptoCashflowSheet && pptoCfIdx !== -1) {
            Object.keys(cfRowsPpto).forEach(key => {
                const row = cfRowsPpto[key];
                if (row) pptoCashflowDetail[key] = getVal(row, pptoCfIdx, false);
            });
            // Compound keys integration
            pptoCashflowDetail.otrosActivos = (pptoCashflowDetail.otrosActivos || 0) + (pptoCashflowDetail.otrosActivosCorrientes || 0) + (pptoCashflowDetail.activosTerceros || 0);
            pptoCashflowDetail.otrosPasivos = pptoCashflowDetail.otrasCxp || 0;
            pptoCashflowDetail.pasivoLaboral = (pptoCashflowDetail.pasivoLaboral || 0) + (pptoCashflowDetail.otrosPasivosCorrientes || 0);
            
            if (pptoCashflowDetail.interest_earned !== undefined || pptoCashflowDetail.interest_expense !== undefined) {
                const sumInterest = (pptoCashflowDetail.interest_earned || 0) + (pptoCashflowDetail.interest_expense || 0);
                if (sumInterest !== 0) {
                    pptoCashflowDetail.interest = sumInterest;
                }
            }
            
            // Prevent double counting if we iterate this dict anywhere else later
            delete pptoCashflowDetail.otrosActivosCorrientes;
            delete pptoCashflowDetail.activosTerceros;
            delete pptoCashflowDetail.otrasCxp;
            delete pptoCashflowDetail.otrosPasivosCorrientes;
            delete pptoCashflowDetail.interest_earned;
            delete pptoCashflowDetail.interest_expense;
        }

        if (pnlRowsPpto) {
            pptoIngresos = pptoPnlIdx !== -1 && pnlRowsPpto.ingresos ? getVal(pnlRowsPpto.ingresos, pptoPnlIdx, true) : 0; 
            pptoCostos = pptoPnlIdx !== -1 && pnlRowsPpto.costos ? getVal(pnlRowsPpto.costos, pptoPnlIdx, true) : 0;
            pptoEbitda = pptoPnlIdx !== -1 && pnlRowsPpto.ebitda ? getVal(pnlRowsPpto.ebitda, pptoPnlIdx, true) : 0;
            pptoOpex = pptoPnlIdx !== -1 && pnlRowsPpto.opex ? getVal(pnlRowsPpto.opex, pptoPnlIdx, true) : 0;
            pptoUtilidad = pptoPnlIdx !== -1 && pnlRowsPpto.utilidad ? getVal(pnlRowsPpto.utilidad, pptoPnlIdx, true) : 0;
            pptoTasaCambio = pptoPnlIdx !== -1 && pnlRowsPpto.tasa_cambio ? getVal(pnlRowsPpto.tasa_cambio, pptoPnlIdx, true) : 1;
            
            if (pptoEbitda !== 0 && pptoOpex === 0) pptoOpex = Math.abs(pptoIngresos - pptoCostos - pptoEbitda);
        }

        if (balanceRowsPpto) {
            pptoActivos = pptoBIdx !== -1 && balanceRowsPpto.activos ? getBalanceVal(balanceRowsPpto.activos, pptoBIdx) : 0;
            pptoPasivos = pptoBIdx !== -1 && balanceRowsPpto.pasivos ? getBalanceVal(balanceRowsPpto.pasivos, pptoBIdx) : 0;
            pptoPatrimonio = pptoBIdx !== -1 && balanceRowsPpto.patrimonio ? getBalanceVal(balanceRowsPpto.patrimonio, pptoBIdx) : 0;
        }

        if (cfRowsPpto) {
            let basePptoCf = pptoPnlIdx !== -1 && pnlRowsPpto && pnlRowsPpto.cashflow ? getVal(pnlRowsPpto.cashflow, pptoPnlIdx, true) : 0;
            pptoCashflowVal = basePptoCf || (pptoCfIdx !== -1 && cfRowsPpto.change ? getVal(cfRowsPpto.change, pptoCfIdx, false) : pptoUtilidad);
            if (pptoCashflowVal === 0 && pptoCfIdx !== -1 && cfRowsPpto.operating) {
                const op = getVal(cfRowsPpto.operating, pptoCfIdx, false);
                if (op !== 0) pptoCashflowVal = op;
            }
        }
        // ====================

        const segments = {};
        Object.entries(segmentRows).forEach(([name, data]) => {
            const sumRows = (rowList) => rowList.reduce((acc, row) => acc + (pIdx !== -1 ? getVal(row, pIdx) : 0), 0);
            segments[name] = {
                ventas: sumRows(data.ventasRows),
                costos: sumRows(data.costosRows)
            };
        });

        const pptoSegments = {};
        if (segmentRowsPpto) {
            Object.entries(segmentRowsPpto).forEach(([name, data]) => {
                const sumRows = (rowList) => rowList.reduce((acc, row) => acc + (pptoPnlIdx !== -1 ? getVal(row, pptoPnlIdx) : 0), 0);
                pptoSegments[name] = {
                    ventas: sumRows(data.ventasRows),
                    costos: sumRows(data.costosRows)
                };
            });
        }

        // El cálculo de la integridad suma los elementos ya que pueden venir en negativo.
        // Calculamos la diferencia considerando posibles variaciones de signos contables.
        
        let detalleOpexSuma = Math.abs(getVal(pnlRows.admin, pIdx)) + 
                              Math.abs(getVal(pnlRows.mercadeo, pIdx)) + 
                              Math.abs(getVal(pnlRows.comercial, pIdx)) + 
                              Math.abs(getVal(pnlRows.logistica, pIdx));
                              
        let ebitdaCalculated = Math.abs(ingresos) - Math.abs(costos) - Math.abs(opex);
        
        if (Math.abs(opex) < 1) {
            // Si OPEX total no se capturó bien en la fila principal, y el calculo usa 0, 
            // descuadra por el valor exacto de los detalles. Usemos la suma de los detalles.
            opex = detalleOpexSuma;
            ebitdaCalculated = Math.abs(ingresos) - Math.abs(costos) - Math.abs(opex);
        }

        const integrityGap = Math.abs(ebitdaCalculated - Math.abs(ebitda));
        
            const ingresosVal = (pnlRows.ingresos && pnlRows.ingresos[0]) ? pnlRows.ingresos[0] : 'No hay ingresos';
            const costosVal = (pnlRows.costos && pnlRows.costos[0]) ? pnlRows.costos[0] : 'No hay costos';
            const opexVal = (pnlRows.opex && pnlRows.opex[0]) ? pnlRows.opex[0] : 'No hay opex';
            const ebitdaVal = (pnlRows.ebitda && pnlRows.ebitda[0]) ? pnlRows.ebitda[0] : 'No hay ebitda';
            console.log("Rows selected for " + formatDateKey(point.date) + ":");
            console.log(" Ingresos:", ingresos, pnlRows.ingresos ? pnlRows.ingresos[0] : null);
            console.log(" Costos:", costos, pnlRows.costos ? pnlRows.costos[0] : null);
            console.log(" Opex:", opex, pnlRows.opex ? pnlRows.opex[0] : null);
            console.log(" EBITDA:", ebitda, pnlRows.ebitda ? pnlRows.ebitda[0] : null);
            console.log(" EBITDA calculated:", ebitdaCalculated);
            console.log(" Gap:", integrityGap);

        // Toleramos un descuadre por Otras Ventas o Depreciaciones (aprox 5% de los ingresos o $150M)
        const integrityError = (ingresos !== 0) ? (integrityGap / Math.abs(ingresos)) > 0.05 && integrityGap > 150 : integrityGap > 150;

        const findRowVal = (rows, search) => {
            const r = rows.filter(r => r && r[0]).find(r => normalizeText(String(r[0])).includes(search));
            const curBIdx = bIdx !== -1 ? bIdx : pIdx;
            return (r && curBIdx !== -1) ? getBalanceVal(r, curBIdx) : 0;
        };

        const deudaTotalBalance = findRowVal(bSheetToUse, "deuda total") || findRowVal(bSheetToUse, "deuda bruta");
        const ebitdaLTM = findRowVal(bSheetToUse, "ltm ebitda") || ebitda * 12;

        return {
            date: formatDateKey(point.date),
            sortDate: point.date,
            kpis: { 
                ingresos, 
                utilidad,
                ebitda, 
                margen_bruto: ingresos !== 0 ? (Math.abs(ingresos) - Math.abs(costos)) / Math.abs(ingresos) : 0,
                margen_ebitda: ingresos !== 0 ? ebitda / ingresos : 0, 
                margen_neto: ingresos !== 0 ? utilidad / ingresos : 0,
                cashflow: cashflowVal, // Este es el flujo neto
                cashEnding: cashflowDetail.ending || 0 // Este es el saldo final (Health)
            },
            deudaMetrics: deudaMetrics,
            cxpDetail: cxpObj,
            ppto: {
                tasaCambio: pptoTasaCambio,
                cashflowDetail: pptoCashflowDetail,
                kpis: {
                    ingresos: pptoIngresos,
                    utilidad: pptoUtilidad,
                    ebitda: pptoEbitda,
                    cashflow: pptoCashflowVal
                },
                balance: {
                    activos: pptoActivos,
                    pasivos: pptoPasivos,
                    patrimonio: pptoPatrimonio,
                    fullRows: pptoBalanceFullRows,
                    conceptRow20: conceptRow20
                },
                pnl: {
                    categorias: { 
                        "Ingresos": pptoIngresos, 
                        "Costo de Ventas": pptoCostos, 
                        "OPEX": pptoOpex, 
                        "EBITDA": pptoEbitda, 
                        "Utilidad Neta": pptoUtilidad 
                    },
                    segments: pptoSegments,
                    opexDetalle: pnlRowsPpto ? {
                        "Gastos Administrativos": pptoPnlIdx !== -1 ? getVal(pnlRowsPpto.admin, pptoPnlIdx, true) : 0,
                        "Gastos de Mercadeo": pptoPnlIdx !== -1 ? getVal(pnlRowsPpto.mercadeo, pptoPnlIdx, true) : 0,
                        "Gastos de Ventas (Comercial)": pptoPnlIdx !== -1 ? getVal(pnlRowsPpto.comercial, pptoPnlIdx, true) : 0,
                        "Gastos de Logística": pptoPnlIdx !== -1 ? getVal(pnlRowsPpto.logistica, pptoPnlIdx, true) : 0
                    } : {},
                    fullRows: pptoFullRows
                }
            },
            balance: { 
                activos, pasivos, patrimonio, deudaTotal: deudaTotalBalance, ebitdaLTM,
                cuadra: Math.abs(activos - (pasivos + patrimonio)) < 100,
                fullRows: balanceFullRows,
                conceptRow20: conceptRow20
            },
            cashflowDetail,
            wcFullRows,
            wcDetail,
            integrity: { gap: integrityGap, isBroken: integrityError },
            tasaCambio: tasaCambio || 1,
            series: { ventas: [], ebitda: [] },
            pnl: { 
                categorias: { "Ingresos": ingresos, "Costo de Ventas": costos, "OPEX": opex, "EBITDA": ebitda, "Utilidad Neta": utilidad },
                opexDetalle: {
                    "Gastos Administrativos": pIdx !== -1 ? getVal(pnlRows.admin, pIdx) : 0,
                    "Gastos de Mercadeo": pIdx !== -1 ? getVal(pnlRows.mercadeo, pIdx) : 0,
                    "Gastos de Ventas (Comercial)": pIdx !== -1 ? getVal(pnlRows.comercial, pIdx) : 0,
                    "Gastos de Logística": pIdx !== -1 ? getVal(pnlRows.logistica, pIdx) : 0
                },
                segments,
                fullRows,
                detectedRows: {
                    ingresos: pnlRows.ingresos ? pnlRows.ingresos[0] : "No encontrada",
                    ebitda: pnlRows.ebitda ? pnlRows.ebitda[0] : "No encontrada",
                    costos: pnlRows.costos ? pnlRows.costos[0] : "No encontrada",
                    opex: pnlRows.opex ? pnlRows.opex[0] : "Calculado (Ventas - Costos - EBITDA)",
                    activos: balanceRows.activos ? balanceRows.activos[0] : "No detectado",
                    pasivo_patrimonio: (balanceRows.pasivos || balanceRows.patrimonio) ? `${balanceRows.pasivos?.[0] || ''} / ${balanceRows.patrimonio?.[0] || ''}` : "No detectado"
                }
            },
            estados: {
                fullRows: estadosFullRows
            },
            alerts: ["FINANCIAL_STATEMENTS: Datos extraídos con desglose de segmentos."]
        };
    });
    const sortedResult = result.sort((a, b) => a.sortDate - b.sortDate);

    // Dynamic second pass for Costos+Gastos YTD and DPO calculation
    let yearlyCostosGastoYTD = {};
    sortedResult.forEach(point => {
        const year = point.sortDate.getFullYear();
        if (yearlyCostosGastoYTD[year] === undefined) {
            yearlyCostosGastoYTD[year] = 0;
        }
        
        yearlyCostosGastoYTD[year] += (point.cxpDetail.costosGastoMensual || 0);
        point.cxpDetail.costosGastoYtd = yearlyCostosGastoYTD[year];
        
        const elapsed_months = (point.sortDate.getMonth() === 11 && point.sortDate.getFullYear() === 2025) ? 12 : (point.sortDate.getMonth() + 1);
        const days = elapsed_months * 30.4;
        
        if (point.cxpDetail.costosGastoYtd > 0) {
            point.cxpDetail.dpo = Math.round((point.cxpDetail.cxpTotal / point.cxpDetail.costosGastoYtd) * days);
        } else {
            point.cxpDetail.dpo = 0;
        }
        
        // Clean up temporary tracking variable
        delete point.cxpDetail.costosGastoMensual;
    });

    return { data: sortedResult };
}

function processWide(sheets) {
    const allRows = Object.values(sheets).flat();
    
    const getVal = (row, idx) => {
        if (!row || idx === undefined || idx === null) return 0;
        let val = cleanNumber(row[idx]);
        const concept = extractConceptName(row).toLowerCase();
        const isFX = concept.includes("tasa") || concept.includes("fx") || concept === "tc" || concept.includes("tipo de cambio") || concept.includes("dop") || concept.includes("exchange");
        const isRatio = concept.includes("%") || concept.includes("ratio");
        const isAlreadyMillions = concept.includes("musd") || concept.includes("mdop") || concept.includes("millones");
        if (!isFX && !isRatio && !isAlreadyMillions) {
            val = val / 1000000;
        }
        return val;
    };

    const rowData = {
        ingresos: findRowByKeywords(allRows, FINANCIAL_KEYWORDS.ingresos),
        costos: findRowByKeywords(allRows, FINANCIAL_KEYWORDS.costos),
        opex: findRowByKeywords(allRows, FINANCIAL_KEYWORDS.opex),
        ebitda: findRowByKeywords(allRows, FINANCIAL_KEYWORDS.ebitda),
        utilidad: findRowByKeywords(allRows, FINANCIAL_KEYWORDS.utilidad),
        cashflow: findRowByKeywords(allRows, FINANCIAL_KEYWORDS.cashflow),
        tasa_cambio: findRowByKeywords(allRows, FINANCIAL_KEYWORDS.tasa_cambio)
    };

    const segmentKeywords = ["BT5", "EVP", "BON", "P6", "Otras Ventas", "Otros Ingresos"];
    const segmentRows = detectSegments(allRows, segmentKeywords);

    if (!rowData.ingresos) return { error: "No se encontró una fila de 'Ingresos' o 'Ventas' en el reporte. Verifique que los nombres de las filas sean claros." };

    const sectionState = [];
    let lastSection = "monthly";
    for (let j = 0; j < (allRows[0]?.length || 0); j++) {
        let detected = null;
        for (let i = 0; i < Math.min(allRows.length, 5); i++) {
            const val = normalizeText(allRows[i]?.[j]);
            if (val.includes("ytd") || val.includes("acum") || val.includes("var")) { detected = "ytd"; break; }
            if (val.includes("monthly") || val.includes("mensual") || val.includes("real")) { detected = "monthly"; break; }
        }
        if (detected) lastSection = detected;
        sectionState[j] = lastSection;
    }

    let dataPoints = [];
    const monthNames = ["ene", "feb", "mar", "abr", "may", "jun", "jul", "ago", "sep", "oct", "nov", "dic"];
    for (let i = 0; i < Math.min(allRows.length, 30); i++) {
        const row = allRows[i];
        if (!row) continue;
        row.forEach((cell, j) => {
            let dateObj = null;
            if (cell instanceof Date) dateObj = cell;
            else if (typeof cell === 'number' && cell > 40000 && cell < 60000) dateObj = new Date((cell - 25569) * 86400 * 1000);
            else if (typeof cell === 'string') {
                const val = cell.toLowerCase();
                const checkMonthPattern = (text) => {
                    const patterns = [
                        /\\b(ene(ro)?|jan(uary)?)\\b/, /\\b(feb(rero|ruary)?)\\b/, /\\b(mar(zo|ch)?)\\b/,
                        /\\b(abr(il)?|apr(il)?)\\b/, /\\b(may(o)?)\\b/, /\\b(jun(io|e)?)\\b/,
                        /\\b(jul(io|y)?)\\b/, /\\b(ago(sto)?|aug(ust)?)\\b/, /\\b(sep(t|tiembre|tember)?)\\b/,
                        /\\b(oct(ubre|ober)?)\\b/, /\\b(nov(iembre|ember)?)\\b/, /\\b(dic(iembre)?|dec(ember)?)\\b/
                    ];
                    for (let k = 0; k < patterns.length; k++) {
                        if (patterns[k].test(text)) return k;
                    }
                    return -1;
                };
                const monthIdx = checkMonthPattern(val);
                if (monthIdx !== -1) {
                    dateObj = new Date();
                    dateObj.setDate(1); // FIX: prevent rollover
                    dateObj.setMonth(monthIdx);
                    if (val.match(/\d{4}/)) dateObj.setFullYear(parseInt(val.match(/\d{4}/)[0]));
                } else if (val.match(/^20\d{2}-\d{1,2}$/)) {
                    const [y, m] = val.split('-').map(Number);
                    dateObj = new Date(y, m - 1, 1);
                }
            }
            if (dateObj) {
                // Filtro de seguridad: No aceptar fechas más allá de 2026 (pedido por usuario)
                if (dateObj.getFullYear() >= 2022 && dateObj.getFullYear() <= 2026) {
                    if (sectionState[j] === "monthly") {
                        const val = cleanNumber(rowData.ingresos[j]);
                        if (val !== 0 && !dataPoints.some(p => p.idx === j)) {
                            dataPoints.push({ idx: j, date: dateObj });
                        }
                    }
                }
            }
        });
        if (dataPoints.length >= 2) break;
    }

    if (dataPoints.length === 0) {
        console.log("🔍 No se detectaron fechas en Wide Format, intentando fallback...");
        rowData.ingresos.forEach((cell, j) => {
            if (j === 0) return;
            const val = cleanNumber(cell);
            if (val !== 0 && !isNaN(val)) {
                if (!dataPoints.some(p => p.idx === j)) {
                    const d = new Date();
                    d.setDate(1);
                    d.setMonth(d.getMonth() - (rowData.ingresos.length - j));
                    dataPoints.push({ idx: j, date: d });
                }
            }
        });
    }

    if (dataPoints.length === 0) {
        return { error: "No se encontraron datos numéricos en las columnas del reporte gerencial." };
    }

    const sampleVal = getVal(rowData.ingresos, dataPoints[0].idx);
    appConfig.isRawData = Math.abs(sampleVal) > 200000;

    // Capturar todas las filas del P&L para la vista detallada (hacer esto fuera del map para evitar O(N^2) y archivos enormes)
    const fullRows = allRows.filter(row => {
        if (!row) return false;
        let conceptRaw = extractConceptName(row);
        if (!conceptRaw) return false;
        const concept = normalizeText(String(conceptRaw));
        if (concept.includes("formatcode") || concept.includes("unnamed") || concept.length < 2) return false;

        // Filtramos filas que tengan al menos 1 número en los dataPoints, o si son Categorias (sin numeros)
        // Agregamos también las filas que sean categorias (por ejemplo "Estado de Resultados") aunque no tengan numeros
        const whitelist = ["costo de ventas", "costos de operacion", "costos", "estado de resultados", "estado de situacion", "otras ventas", "otros ingresos", "evp", "bt5", "bon", "descuentos", "devoluciones", "descuentos y devoluciones", "itbis", "gastos administrativos", "gastos de mercadeo", "gastos de ventas", "gastos de logistica", "d & a", "intereses netos", "ingresos financieros", "gastos financieros", "diferencial cambiario", "ingresos (gastos) extraordinarios", "tasa cambio cierre", "cuentas por cobrar", "inventario", "cuentas por pagar", "gastos de ventas (comercial)"];
        const isCategory = whitelist.some(w => concept.includes(w)) || (concept === "estado de resultados" || concept === "estado de situacion" || concept === "kpis y drivers" || concept === "modulo deuda" || concept === "analisis horizontal" || concept === "analisis vertical" || concept === "analisis margen" || concept === "rentabilidad" || concept === "variables macro" || concept === "balances deuda" || concept === "schedule amortizacion" || concept === "kpis deuda");
        return isCategory || dataPoints.some(p => typeof row[p.idx] === 'number' || (!isNaN(cleanNumber(row[p.idx])) && cleanNumber(row[p.idx]) !== 0));
    }).map(row => {
        const rowValues = {};
        dataPoints.forEach(p => {
            rowValues[formatDateKey(p.date)] = getVal(row, p.idx);
        });
        const rawConcept = extractConceptName(row);
        const renamedConcept = rawConcept;

        return {
            concept: renamedConcept,
            values: rowValues
        };
    });

    const result = dataPoints.map(point => {
        const ingresos = getVal(rowData.ingresos, point.idx);
        const costos = rowData.costos ? getVal(rowData.costos, point.idx) : 0;
        const opex = rowData.opex ? getVal(rowData.opex, point.idx) : 0;
        const ebitda = rowData.ebitda ? getVal(rowData.ebitda, point.idx) : (ingresos - costos - opex);
        const utilidad = rowData.utilidad ? getVal(rowData.utilidad, point.idx) : 0;
        const cashflow = rowData.cashflow ? getVal(rowData.cashflow, point.idx) : utilidad;
        const tasaCambio = rowData.tasa_cambio ? getVal(rowData.tasa_cambio, point.idx) : 1;

        const segments = {};
        Object.entries(segmentRows).forEach(([name, data]) => {
            const sumVals = (rowList) => rowList.reduce((acc, row) => acc + getVal(row, point.idx), 0);
            segments[name] = {
                ventas: sumVals(data.ventasRows),
                costos: sumVals(data.costosRows)
            };
        });

        return {
            date: formatDateKey(point.date),
            sortDate: point.date,
            kpis: {
                ingresos,
                utilidad,
                ebitda,
                margen_bruto: ingresos !== 0 ? (Math.abs(ingresos) - Math.abs(costos)) / Math.abs(ingresos) : 0,
                margen_ebitda: ingresos !== 0 ? ebitda / ingresos : 0,
                margen_neto: ingresos !== 0 ? utilidad / ingresos : 0,
                cashflow
            },
            balance: { activos: 0, pasivos: 0, patrimonio: 0, cuadra: true },
            tasaCambio: tasaCambio,
            wcFullRows: fullRows,
            series: { ventas: [], ebitda: [] },
            pnl: { 
                categorias: { "Ingresos": ingresos, "Costo de Ventas": costos, "OPEX": opex, "EBITDA": ebitda, "Utilidad Neta": utilidad },
                segments: segments,
                fullRows: fullRows,
                detectedRows: {
                    ingresos: rowData.ingresos ? rowData.ingresos[0] : "No encontrada",
                    ebitda: rowData.ebitda ? rowData.ebitda[0] : "No encontrada",
                    costos: rowData.costos ? rowData.costos[0] : "No encontrada"
                }
            },
            alerts: ["WIDE_FORMAT: Reporte gerencial detectado."]
        };
    }).sort((a, b) => a.sortDate - b.sortDate);

    // Final deduplication by date string
    const uniqueResult = [];
    const seenDates = new Set();
    result.forEach(item => {
        if (!seenDates.has(item.date)) {
            seenDates.add(item.date);
            uniqueResult.push(item);
        }
    });

    return { data: uniqueResult };
}

function processTBSetup(sheets, tbKey, setupKey) {
    appConfig.isRawData = true; // Trial balance data is almost universally raw
    const tbSheet = sheets[tbKey];
    const setupSheet = sheets[setupKey];
    
    const setupMap = new Map();
    const diagnostics = { 
        rows: 0, 
        mapped: 0, 
        tbSample: [], 
        setupSample: [],
        tbColDetected: -1,
        setupColDetected: -1
    };

    function cleanAccount(val) {
        if (val === undefined || val === null) return "";
        return String(val).trim().toLowerCase().replace(/[^a-z0-9]/g, '');
    }

    function detectAccountColumnByContent(sheet, startRow) {
        const scores = [];
        const sampleRows = sheet.slice(startRow, startRow + 150);
        sampleRows.forEach(row => {
            if (!row) return;
            row.forEach((cell, j) => {
                if (scores[j] === undefined) scores[j] = 0;
                const val = String(cell).trim();
                if (!val || val.length < 2) return;
                if (/^\d{4,10}$/.test(val)) scores[j] += 10;
                else if (/^(\d+[\.\-])+\d+$/.test(val)) scores[j] += 15;
                else if (/^[A-Z0-9]{4,12}$/i.test(val) && !val.includes(" ")) scores[j] += 5;
            });
        });
        let bestCol = -1; let maxScore = 0;
        scores.forEach((score, j) => { if (score > maxScore) { maxScore = score; bestCol = j; } });
        return { col: bestCol, score: maxScore };
    }

    function detectDateColumnByContent(sheet, startRow) {
        const scores = [];
        const sampleRows = sheet.slice(startRow, startRow + 100);
        const monthNames = ["ene", "feb", "mar", "abr", "may", "jun", "jul", "ago", "sep", "oct", "nov", "dic", "jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec"];
        sampleRows.forEach(row => {
            if (!row) return;
            row.forEach((cell, j) => {
                if (scores[j] === undefined) scores[j] = 0;
                if (cell instanceof Date) scores[j] += 20;
                else if (typeof cell === 'number' && cell > 35000 && cell < 60000) scores[j] += 15;
                else if (typeof cell === 'string') {
                    const val = cell.toLowerCase();
                    const checkMonthPattern = (text) => {
                        const patterns = [
                            /\\b(ene(ro)?|jan(uary)?)\\b/, /\\b(feb(rero|ruary)?)\\b/, /\\b(mar(zo|ch)?)\\b/,
                            /\\b(abr(il)?|apr(il)?)\\b/, /\\b(may(o)?)\\b/, /\\b(jun(io|e)?)\\b/,
                            /\\b(jul(io|y)?)\\b/, /\\b(ago(sto)?|aug(ust)?)\\b/, /\\b(sep(t|tiembre|tember)?)\\b/,
                            /\\b(oct(ubre|ober)?)\\b/, /\\b(nov(iembre|ember)?)\\b/, /\\b(dic(iembre)?|dec(ember)?)\\b/
                        ];
                        for (let k = 0; k < patterns.length; k++) {
                            if (patterns[k].test(text)) return k;
                        }
                        return -1;
                    };
                    if (checkMonthPattern(val) !== -1) scores[j] += 10;
                    if (/\d{4}[-\/]\d{2}/.test(val)) scores[j] += 12;
                }
            });
        });
        let bestCol = -1; let maxScore = 0;
        scores.forEach((score, j) => { if (score > maxScore) { maxScore = score; bestCol = j; } });
        return { col: bestCol, score: maxScore };
    }

    function detectBalanceColumnByContent(sheet, startRow) {
        const scores = [];
        const sampleRows = sheet.slice(startRow, startRow + 150);
        sampleRows.forEach(row => {
            if (!row) return;
            row.forEach((cell, j) => {
                if (scores[j] === undefined) scores[j] = 0;
                const val = Number(cell);
                if (isNaN(val) || cell === null || cell === "") return;
                if (!Number.isInteger(val)) scores[j] += 10;
                if (Math.abs(val) > 100) scores[j] += 5;
            });
        });
        let bestCol = -1; let maxScore = 0;
        scores.forEach((score, j) => { if (score > maxScore) { maxScore = score; bestCol = j; } });
        return { col: bestCol, score: maxScore };
    }

    function findHeaderRow(sheet, keywords) {
        let bestRow = 0; let maxScore = -1; let bestCols = {};
        for (let i = 0; i < Math.min(sheet.length, 30); i++) {
            const row = sheet[i]; if (!row) continue;
            let currentScore = 0; let currentCols = {};
            row.forEach((cell, j) => {
                const c = normalizeText(cell); if (!c) return;
                for (const [key, searchTerms] of Object.entries(keywords)) {
                    if (searchTerms.some(term => c === term || (term.length > 3 && c.includes(term)))) {
                        if (!currentCols[key]) { currentCols[key] = j; currentScore++; }
                    }
                }
            });
            if (currentScore > maxScore) { maxScore = currentScore; bestRow = i; bestCols = currentCols; }
        }
        return { row: bestRow, cols: bestCols, score: maxScore };
    }

    const setupKeywords = {
        cuenta: ["cuenta", "codigo", "acct", "account", "cta", "id"],
        cat: ["categoria", "grupo", "category", "clase", "tipo"],
        sub: ["subcategoria", "subcat", "subgrupo"],
        signo: ["signo", "multiplicador", "sign", "naturaleza", "factor"]
    };

    const setupHeader = findHeaderRow(setupSheet, setupKeywords);
    let setupCols = setupHeader.cols;
    let setupHeaderRow = setupHeader.row;

    const setupContentAcc = detectAccountColumnByContent(setupSheet, setupHeaderRow + 1);
    if (setupCols.cuenta === undefined || setupCols.cuenta === -1) setupCols.cuenta = setupContentAcc.col !== -1 ? setupContentAcc.col : 0;
    if (setupCols.cat === undefined || setupCols.cat === -1) setupCols.cat = 1;
    if (setupCols.sub === undefined || setupCols.sub === -1) setupCols.sub = 2;
    if (setupCols.signo === undefined || setupCols.signo === -1) setupCols.signo = setupSheet[setupHeaderRow + 1] ? setupSheet[setupHeaderRow + 1].length - 1 : 3;

    diagnostics.setupColDetected = setupCols.cuenta;

    setupSheet.forEach((row, i) => {
        if (i <= setupHeaderRow || !row) return;
        const cuenta = cleanAccount(row[setupCols.cuenta]);
        if (!cuenta) return;
        if (diagnostics.setupSample.length < 5) diagnostics.setupSample.push(cuenta);
        setupMap.set(cuenta, {
            categoria: row[setupCols.cat] || "Sin Categoría",
            subcategoria: row[setupCols.sub] || "Sin Subcategoría",
            signo: cleanNumber(row[setupCols.signo]) || 1
        });
    });

    if (setupMap.size === 0) return { error: "La hoja 'Setup' no tiene datos de cuenta válidos." };

    const tbKeywords = {
        cuenta: ["cuenta", "codigo", "acct", "account", "cta", "id", "cod"],
        fecha: ["fecha", "periodo", "mes", "date", "year", "ano", "time", "fec"],
        balance: ["balance", "saldo", "monto", "final", "amount", "debe", "haber", "neto", "total", "valor"]
    };

    const tbHeader = findHeaderRow(tbSheet, tbKeywords);
    let tbCols = tbHeader.cols;
    let tbHeaderRow = tbHeader.row;

    const tbContentAcc = detectAccountColumnByContent(tbSheet, tbHeaderRow + 1);
    const tbContentDate = detectDateColumnByContent(tbSheet, tbHeaderRow + 1);
    const tbContentBal = detectBalanceColumnByContent(tbSheet, tbHeaderRow + 1);

    if (tbCols.cuenta === undefined || tbCols.cuenta === -1) tbCols.cuenta = tbContentAcc.col !== -1 ? tbContentAcc.col : 0;
    if (tbCols.fecha === undefined || tbCols.fecha === -1) tbCols.fecha = tbContentDate.col !== -1 ? tbContentDate.col : 1;
    if (tbCols.balance === undefined || tbCols.balance === -1) tbCols.balance = tbContentBal.col !== -1 ? tbContentBal.col : 2;

    diagnostics.tbColDetected = tbCols.cuenta;

    const monthlyAggregates = {};

    tbSheet.forEach((row, idx) => {
        if (idx <= tbHeaderRow || !row) return;
        const cuenta = cleanAccount(row[tbCols.cuenta]);
        if (!cuenta) return;
        diagnostics.rows++;
        if (diagnostics.tbSample.length < 5 && !diagnostics.tbSample.includes(cuenta)) diagnostics.tbSample.push(cuenta);

        const setup = setupMap.get(cuenta);
        if (!setup) return;
        diagnostics.mapped++;

        const rawDate = row[tbCols.fecha];
        let dateObj = null;
        if (rawDate instanceof Date) dateObj = rawDate;
        else if (typeof rawDate === 'number') dateObj = new Date((rawDate - 25569) * 86400 * 1000);
        else if (typeof rawDate === 'string') dateObj = new Date(rawDate);

        if (!dateObj || isNaN(dateObj.getTime())) return;

        const dateKey = formatDateKey(dateObj);
        if (!monthlyAggregates[dateKey]) {
            monthlyAggregates[dateKey] = { 
                kpis: { ingresos: 0, utilidad: 0, ebitda: 0, margen_bruto: 0, margen_ebitda: 0, margen_neto: 0, cashflow: 0 },
                balance: { activos: 0, pasivos: 0, patrimonio: 0, cuadra: false },
                series: { ventas: [], ebitda: [] },
                pnl: { categorias: {} },
                alerts: [],
                sortDate: dateObj,
                date: dateKey,
                _raw: { ingresos: 0, costos: 0, gastos: 0 }
            };
        }

        const valorAjustado = (cleanNumber(row[tbCols.balance]) / 1000000) * setup.signo;
        const cat = normalizeText(setup.categoria);
        const agg = monthlyAggregates[dateKey];

        if (cat.includes("ingreso") || cat.includes("venta")) agg._raw.ingresos += valorAjustado;
        else if (cat.includes("costo")) agg._raw.costos += valorAjustado;
        else if (cat.includes("gasto")) agg._raw.gastos += valorAjustado;

        if (cat.includes("activo")) agg.balance.activos += valorAjustado;
        else if (cat.includes("pasivo")) agg.balance.pasivos += valorAjustado;
        else if (cat.includes("patrimonio")) agg.balance.patrimonio += valorAjustado;

        if (!agg.pnl.categorias[setup.categoria]) agg.pnl.categorias[setup.categoria] = 0;
        agg.pnl.categorias[setup.categoria] += valorAjustado;
    });

    const result = Object.values(monthlyAggregates).sort((a, b) => a.sortDate - b.sortDate).map(agg => {
        agg.kpis.ingresos = agg._raw.ingresos;
        agg.kpis.ebitda = agg._raw.ingresos - agg._raw.costos - agg._raw.gastos;
        agg.kpis.margen_bruto = agg.kpis.ingresos !== 0 ? (Math.abs(agg.kpis.ingresos) - Math.abs(agg._raw.costos)) / Math.abs(agg.kpis.ingresos) : 0;
        agg.kpis.margen_ebitda = agg.kpis.ingresos !== 0 ? agg.kpis.ebitda / agg.kpis.ingresos : 0;
        agg.kpis.cashflow = agg.kpis.ebitda;
        agg.balance.cuadra = Math.abs(agg.balance.activos - (agg.balance.pasivos + agg.balance.patrimonio)) < 100;
        return agg;
    });

    if (result.length === 0) {
        return { error: `ERROR DE MAPEADO: No hay coincidencia entre TB y Setup.\n\nDIAGNÓSTICO:\n${JSON.stringify(diagnostics, null, 2)}` };
    }

    return { data: result };
}
