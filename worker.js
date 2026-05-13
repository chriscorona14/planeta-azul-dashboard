import * as XLSX from 'xlsx';
import { financialEngine } from './financialEngine.js';

self.onmessage = function(e) {
    try {
        const buffer = e.data.buffer;
        const fileType = e.data.fileType || 'master';
        
        self.postMessage({ type: 'progress', progress: 50, message: "Decodificando archivo Excel en segundo plano..." });
        let workbook = XLSX.read(new Uint8Array(buffer), { type: 'array', cellDates: true });
        
        if (fileType === 'ventas_ceo') {
            self.postMessage({ type: 'progress', progress: 75, message: "Procesando datos de Ventas CEO..." });
            
            const consejoSheetName = workbook.SheetNames.find(n => n.toLowerCase().includes('consejo'));
            const dataSheetName = workbook.SheetNames.find(n => n.toLowerCase().includes('data por mes'));
            
            let bestSheetName = workbook.SheetNames[0];
            if (!consejoSheetName) {
                for (let name of workbook.SheetNames) {
                    const sheetTmp = workbook.Sheets[name];
                    const rowsTmp = XLSX.utils.sheet_to_json(sheetTmp, { header: 1 });
                    const hasProducto = rowsTmp.some(r => r && r.some(c => String(c).toLowerCase().trim() === 'producto' || String(c).toLowerCase().trim() === 'descripción'));
                    if (hasProducto) {
                        bestSheetName = name;
                        break;
                    }
                }
            }

            const result = {
                consejoSheetName,
                dataSheetName,
                bestSheetName,
                consejoRows: consejoSheetName ? XLSX.utils.sheet_to_json(workbook.Sheets[consejoSheetName], { range: 2, defval: 0 }) : null,
                dataRows: dataSheetName ? XLSX.utils.sheet_to_json(workbook.Sheets[dataSheetName], { header: 1 }) : null,
                bestRows: XLSX.utils.sheet_to_json(workbook.Sheets[bestSheetName], { header: 1 })
            };
            
            workbook = null;
            self.postMessage({ type: 'done_ventas', result });
            return;
        }

        // Master processing
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
