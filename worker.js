import * as XLSX from 'xlsx';
import { financialEngine } from './financialEngine.js';

self.onmessage = function(e) {
    try {
        const buffer = e.data.buffer;
        const fileType = e.data.fileType || 'master';
        
        self.postMessage({ type: 'progress', progress: 50, message: "Decodificando archivo Excel en segundo plano..." });
        
        let workbook;

        // ==========================================
        // 1. PROCESAMIENTO DE VENTAS CEO (SIN cellDates)
        // ==========================================
        if (fileType === 'ventas_ceo') {
            workbook = XLSX.read(new Uint8Array(buffer), { type: 'array' });
            
            self.postMessage({ type: 'progress', progress: 75, message: "Procesando datos de Ventas CEO..." });
            
            const sheetNames = workbook.SheetNames;
            const consejoSheetName = sheetNames.find(n => n.toLowerCase().includes('consejo'));
            const dataSheetName = sheetNames.find(n => n.toLowerCase().includes('data por mes') || n.toLowerCase().includes('datos por mes'));
            
            if (!consejoSheetName && !dataSheetName) {
                self.postMessage({ 
                    type: 'error', 
                    error: `Estructura inválida. No se encontraron las hojas "Consejo" o "Data por mes". Hojas detectadas: [${sheetNames.join(', ')}]` 
                });
                return; 
            }

            // BÚSQUEDA INTELIGENTE DE LA FILA DE TÍTULOS PARA 'TABLAS CONSEJO'
            let consejoRows = null;
            if (consejoSheetName) {
                const sheet = workbook.Sheets[consejoSheetName];
                // Extraemos todo como matriz para escanear las filas
                const rawData = XLSX.utils.sheet_to_json(sheet, { header: 1 });
                
                let headerRowIndex = 1; // Por defecto (rango 1), asume que el título está en la segunda fila
                for (let i = 0; i < rawData.length; i++) {
                    if (!rawData[i]) continue;
                    
                    // Buscamos si esta fila contiene la palabra TOTAL
                    const hasTotal = rawData[i].some(c => c && String(c).toUpperCase().trim() === 'TOTAL');
                    if (hasTotal) {
                        // Si encontramos TOTAL, los títulos siempre son la fila de arriba (i - 1)
                        headerRowIndex = i > 0 ? i - 1 : 0; 
                        break;
                    }
                }
                
                // Extraemos la tabla final indicándole el inicio correcto
                consejoRows = XLSX.utils.sheet_to_json(sheet, { range: headerRowIndex, defval: 0 });
            }

            let bestSheetName = sheetNames[0];
            if (!consejoSheetName) {
                let maxScore = -1;
                for (let name of sheetNames) {
                    const sheetTmp = workbook.Sheets[name];
                    const rowsTmp = XLSX.utils.sheet_to_json(sheetTmp, { header: 1 });
                    let score = 0;
                    for (let r of rowsTmp) {
                        if (!r) continue;
                        for (let c of r) {
                            if (c === undefined || c === null) continue;
                            const term = String(c).toLowerCase().trim();
                            if (term === 'producto' || term === 'descripción' || term === 'descripcion') score += 10;
                            if (term === 'tipo') score += 5;
                            if (term.includes('ventas netas dop') || term.includes('ventas netas')) score += 8;
                            if (term.includes('volumen unidades') || term.includes('volumen')) score += 8;
                            if (term === '2026' || term.includes('ppto')) score += 5;
                        }
                    }
                    if (score > maxScore && score > 0) {
                        maxScore = score;
                        bestSheetName = name;
                    }
                }
            }

            const result = {
                consejoSheetName,
                dataSheetName,
                bestSheetName,
                consejoRows: consejoRows,
                dataRows: dataSheetName ? XLSX.utils.sheet_to_json(workbook.Sheets[dataSheetName], { header: 1 }) : null,
                bestRows: XLSX.utils.sheet_to_json(workbook.Sheets[bestSheetName], { header: 1 })
            };
            
            workbook = null;
            self.postMessage({ type: 'done_ventas', result });
            return;
        }

        // ==========================================
        // 2. PROCESAMIENTO MASTER FINANCIERO (CON cellDates)
        // ==========================================
        workbook = XLSX.read(new Uint8Array(buffer), { type: 'array', cellDates: true });
        
        self.postMessage({ type: 'progress', progress: 75, message: "Ejecutando motor de datos financieros..." });
        let engineResult = financialEngine(workbook);
        
        workbook = null;

        if (engineResult.error || !engineResult.data || engineResult.data.length === 0) {
            self.postMessage({ 
                type: 'error', 
                error: engineResult.error || "No se pudieron extraer datos numéricos del archivo." 
            });
            return;
        }

        self.postMessage({ 
            type: 'done', 
            engineResult 
        });
    } catch (err) {
        self.postMessage({ type: 'error', error: err.message || "Ocurrió un error en el worker." });
    }
};
